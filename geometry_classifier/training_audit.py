"""Scan a training folder (mixed BMS + standard jobs), build CORRECT_ME files,
compare against the live rules classifier, and emit actionable suggestions.

Designed for folders like C:\\Users\\lenovo\\Downloads\\TRAINING where each
subfolder may contain steel sheets, BOM PDFs, XT exports, and SolidWorks files.
"""
from __future__ import annotations

import csv
import json
import re
import sys
import threading
import time
from collections import Counter, defaultdict
from pathlib import Path

BASE = Path(__file__).resolve().parent
OUT_DIR = BASE / "outputs" / "training"
DATA_DIR = BASE / "data" / "training"
REPORT_PATH = DATA_DIR / "last_audit_report.json"
SUGGESTIONS_PATH = DATA_DIR / "rule_suggestions.md"
PROGRESS_PATH = DATA_DIR / "training_progress.json"

_progress_lock = threading.Lock()
_progress: dict = {
    "running": False,
    "phase": "idle",
    "current_job": "",
    "job_index": 0,
    "job_total": 0,
    "message": "",
    "use_qwen": False,
    "qwen_model": "",
    "error": "",
}

# Import sibling modules
if str(BASE.parent) not in sys.path:
    sys.path.insert(0, str(BASE.parent))

from geometry_classifier import train_from_quote_sheets  # noqa: E402
from geometry_classifier.qwen_classify_xt_csv import (  # noqa: E402
    ROLES,
    classify_geometry,
    read_rows,
)

C_NUMBER_RE = re.compile(r"\b(C\d{4,6})\b", re.I)
JOB_NUMBER_RE = re.compile(r"\b(\d{6,8})\b")

BMS_MARKERS = (
    "moldbase",
    "mold base",
    "mold-base",
    "pot block",
    "pot-block",
    "potblock",
    "bms-",
    "-bom",
    "_bom",
    "bom pricing",
    "bom.pdf",
)

STANDARD_MARKERS = (
    "a-plate",
    "a_plate",
    "b-plate",
    "b_plate",
    "sc-retainer",
    "sc-backup",
    "ej-ret",
    "ej-backup",
    "ldr-pin",
    "latch-lock",
    "dynacast-",
)

SUGGESTION_TEMPLATES = {
    "a_plate": "Add or strengthen A-PLATE / full-footprint stack rule for plate at stack position {}",
    "b_plate": "Add or strengthen B-PLATE token + parting-line rule; do not let leader pins flip A/B {}",
    "ejector_plate": "Ejector stack: thinner centered plate = ejector_plate (never ejector_retainer_plate) {}",
    "bottom_ejector_plate": "Ejector stack: thicker/lower plate = bottom_ejector_plate {}",
    "sc_retainer_plate": "Latch-lock/SC base: SC-RETAINER-PLATE token or stack position {}",
    "sc_backup_plate": "Latch-lock/SC base: SC-BACKUP-PLATE token or stack position {}",
    "leader_pin": "Leader pin: match long round bar to bushing at same XY plane; rails anchor stack first {}",
    "leader_pin_bushing": "Leader bushing: short cylinder near guide hardware; reversed pins must not flip A/B {}",
    "latch_lock": "PLC/LATCH-LOCK/SAFETY-STRAP token marks plate-sequenced base; secondary parting only {}",
    "rail": "Rail: long narrow full-length side block between bottom and support zone {}",
    "bottom_clamp_plate": "CLAMP-PLATE token or bottom full-footprint plate {}",
    "support_plate": "Support plate in 5-plate standard stack between B and bottom clamp {}",
}


def _safe_float(v) -> float:
    try:
        return float(str(v).strip())
    except Exception:
        return 0.0


def extract_job_id(folder: Path) -> str:
    """Prefer C-number from filenames, else folder name."""
    for path in folder.rglob("*"):
        if not path.is_file():
            continue
        m = C_NUMBER_RE.search(path.name)
        if m:
            return m.group(1).upper()
    name = folder.name
    m = C_NUMBER_RE.search(name)
    if m:
        return m.group(1).upper()
    m = JOB_NUMBER_RE.search(name)
    if m:
        return m.group(1)
    return name


def detect_base_type(folder: Path) -> tuple[str, list[str]]:
    """Return ('bms'|'standard'|'unknown', list of reasons)."""
    signals_bms: list[str] = []
    signals_std: list[str] = []

    for path in folder.rglob("*"):
        if not path.is_file():
            continue
        low = path.name.lower()
        if any(m in low for m in BMS_MARKERS):
            signals_bms.append(path.name)
        if low.endswith((".sldasm", ".sldprt")) and "mold" in low:
            signals_bms.append(path.name)
        if any(m in low.replace("_", "-") for m in STANDARD_MARKERS):
            signals_std.append(path.name)

    # BASE subfolder with only part files often indicates BMS pot-block assembly
    base_dir = folder / "BASE"
    if base_dir.is_dir():
        children = list(base_dir.iterdir())
        if children and not train_from_quote_sheets.find_xt_csv(folder):
            signals_bms.append("BASE/ subfolder without XT export (typical BMS layout)")

    if signals_bms and not signals_std:
        return "bms", signals_bms[:8]
    if signals_std and not signals_bms:
        return "standard", signals_std[:8]
    if signals_bms and signals_std:
        # Steel sheet + BOM is common on finished BMS jobs — BOM wins for macro path
        if any("bom" in s.lower() or "moldbase" in s.lower() for s in signals_bms):
            return "bms", signals_bms[:8] + ["(also has standard-like files)"]
        return "standard", signals_std[:8] + ["(also has BMS-like files)"]
    return "unknown", []


def discover_job_folders(root: Path, max_depth: int = 3) -> list[Path]:
    """Find folders that look like quote jobs (steel/quote/BOM/CAD), recursively."""
    root = root.resolve()
    if not root.exists():
        return []

    def is_job_folder(p: Path) -> bool:
        if not p.is_dir():
            return False
        names = [f.name.lower() for f in p.iterdir() if f.is_file()]
        dirs = [d.name.lower() for d in p.iterdir() if d.is_dir()]
        has_sheet = any(
            "steel" in n or "quote" in n for n in names if n.endswith((".xls", ".xlsx", ".xlsm"))
        )
        has_bom = any("bom" in n or "moldbase" in n for n in names)
        has_cad = any(n.endswith((".sldasm", ".igs", ".xt", ".step", ".stp")) for n in names)
        has_xt = train_from_quote_sheets.find_xt_csv(p) is not None
        has_base = "base" in dirs
        return has_sheet or has_bom or has_xt or (has_cad and has_sheet) or has_base

    found: list[Path] = []

    def walk(p: Path, depth: int) -> None:
        if depth > max_depth:
            return
        if is_job_folder(p):
            found.append(p)
            return  # don't descend into subfolders of a job
        for child in sorted(p.iterdir()):
            if child.is_dir() and child.name.lower() not in ("outputs", "training", ".git", "node_modules"):
                walk(child, depth + 1)

    if is_job_folder(root):
        found.append(root)
    else:
        walk(root, 0)
    return found


def _compact_row(row: dict) -> dict:
    return {
        "i": str(row.get("Index", "")),
        "t": _safe_float(row.get("Thickness")),
        "w": _safe_float(row.get("Width")),
        "l": _safe_float(row.get("Length")),
        "v": _safe_float(row.get("BBoxVolume_cuin")),
        "x": _safe_float(row.get("CenterX")),
        "y": _safe_float(row.get("CenterY")),
        "z": _safe_float(row.get("CenterZ")),
        "name": str(row.get("Component", ""))[:80],
    }


def _write_progress(**fields) -> None:
    with _progress_lock:
        _progress.update(fields)
        _progress["updated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        PROGRESS_PATH.write_text(json.dumps(_progress, indent=2), encoding="utf-8")


def get_progress() -> dict:
    if PROGRESS_PATH.exists():
        try:
            saved = json.loads(PROGRESS_PATH.read_text(encoding="utf-8"))
            with _progress_lock:
                _progress.update(saved)
        except Exception:
            pass
    with _progress_lock:
        return dict(_progress)


def _load_truth_from_correct_me(correct_me_path: Path) -> dict[str, str]:
    truth: dict[str, str] = {}
    if not correct_me_path.exists():
        return truth
    with correct_me_path.open("r", newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            role = (row.get("CorrectRole") or "").strip()
            if role:
                truth[str(row.get("Index", ""))] = role
    return truth


def _audit_predictions_against_truth(
    truth: dict[str, str], pred_map: dict[str, str], label: str
) -> dict:
    mismatches = []
    correct = 0
    for idx, expected in truth.items():
        got = pred_map.get(idx, "")
        if got == expected:
            correct += 1
        else:
            mismatches.append(
                {
                    "index": idx,
                    "expected": expected,
                    "predicted": got or "(none)",
                    "source": label,
                }
            )
    total = len(truth)
    return {
        "compared": total,
        "correct": correct,
        "mismatches": mismatches,
        "accuracy_pct": round(100 * correct / max(total, 1), 1),
    }


def run_qwen_on_xt(xt_path: Path, model: str = "qwen3.5:9b", timeout_minutes: int = 0) -> dict:
    """Run Ollama/Qwen on one XT export (slow — for training only)."""
    from geometry_classifier.qwen_classify_xt_csv import (  # noqa: E402
        build_prompt,
        classify_geometry,
        extract_json,
        read_rows,
        run_ollama,
    )

    rows = read_rows(str(xt_path), include_names=True)
    if len(rows) > 5000:
        rows = rows[:5000]
    prompt = build_prompt(rows, str(xt_path), long_knowledge=True)
    started = time.time()
    try:
        raw = run_ollama(prompt, model, timeout_minutes)
        data = extract_json(raw)
        return {
            "qwen_ran": True,
            "qwen_model": model,
            "elapsed_sec": round(time.time() - started, 1),
            "data": data,
        }
    except Exception as exc:
        fallback = classify_geometry(rows)
        return {
            "qwen_ran": False,
            "qwen_model": model,
            "error": str(exc),
            "elapsed_sec": round(time.time() - started, 1),
            "data": fallback,
        }


def audit_rules_against_correct_me(correct_me_path: Path, xt_path: Path) -> dict:
    """Compare deterministic rules vs CORRECT_ME ground truth."""
    truth = _load_truth_from_correct_me(correct_me_path)
    if not truth:
        return {"compared": 0, "correct": 0, "mismatches": [], "accuracy_pct": 0}

    xt_rows = [_compact_row(r) for r in train_from_quote_sheets.read_xt_rows(xt_path)]
    predicted = classify_geometry(xt_rows)
    pred_map = {str(c["index"]): c["role"] for c in predicted.get("classifications", [])}
    audit = _audit_predictions_against_truth(truth, pred_map, "rules")
    mismatches = []
    for mm in audit["mismatches"]:
        comp = next((r for r in xt_rows if r["i"] == mm["index"]), {})
        mismatches.append(
            {
                **mm,
                "component": comp.get("name", ""),
                "thickness": comp.get("t"),
                "width": comp.get("w"),
                "length": comp.get("l"),
            }
        )
    audit["mismatches"] = mismatches
    return audit


def audit_qwen_against_correct_me(correct_me_path: Path, qwen_data: dict) -> dict:
    truth = _load_truth_from_correct_me(correct_me_path)
    if not truth:
        return {"compared": 0, "correct": 0, "mismatches": [], "accuracy_pct": 0}
    pred_map = {
        str(c.get("index", "")): c.get("role", "")
        for c in qwen_data.get("classifications", [])
    }
    return _audit_predictions_against_truth(truth, pred_map, "qwen")


def build_suggestions(all_mismatches: list[dict], job_results: list[dict]) -> list[dict]:
    """Turn mismatch patterns into macro/classifier suggestions."""
    by_role: Counter = Counter()
    patterns: defaultdict[str, list] = defaultdict(list)

    for job in job_results:
        for mm in job.get("audit", {}).get("mismatches", []):
            key = mm.get("expected", "")
            if key:
                by_role[key] += 1
                patterns[key].append(mm)

    suggestions: list[dict] = []
    for role, count in by_role.most_common():
        template = SUGGESTION_TEMPLATES.get(role, f"Improve rule coverage for role '{role}'")
        examples = patterns[role][:3]
        ex_text = "; ".join(
            f"idx {e['index']} {e['component'][:40]} (got {e['predicted']})" for e in examples
        )
        suggestions.append(
            {
                "priority": "high" if count >= 3 else "medium" if count >= 2 else "low",
                "role": role,
                "occurrences": count,
                "suggestion": template,
                "examples": ex_text,
                "action": _action_for_role(role),
            }
        )

    # Macro-level guidance (always include)
    suggestions.insert(
        0,
        {
            "priority": "critical",
            "role": "macro",
            "occurrences": 0,
            "suggestion": (
                "Module6121: keep RunAiBridgeClassification guard — BMS jobs must call "
                "AiBridgeNotifyBms only; standard jobs call /api/vba/classify (rules, ~seconds). "
                "Never run Qwen during live quotes."
            ),
            "examples": "",
            "action": "macro",
        },
    )
    bms_count = sum(1 for j in job_results if j.get("base_type") == "bms")
    std_count = sum(1 for j in job_results if j.get("base_type") == "standard")
    if bms_count:
        suggestions.append(
            {
                "priority": "critical",
                "role": "bms",
                "occurrences": bms_count,
                "suggestion": (
                    f"{bms_count} BMS job(s) detected: macro must use BOM-driven flow "
                    "(ReadCustomerBom, Match BOM to CAD). AI bridge registers job only — no classify."
                ),
                "examples": ", ".join(
                    j["job_id"] for j in job_results if j.get("base_type") == "bms"
                )[:200],
                "action": "macro_bms",
            }
        )
    if std_count:
        suggestions.append(
            {
                "priority": "high",
                "role": "standard",
                "occurrences": std_count,
                "suggestion": (
                    f"{std_count} standard job(s): macro calls /api/vba/classify after XT export. "
                    "Improve Python rules using CORRECT_ME files below — not bigger Qwen at quote time."
                ),
                "examples": "",
                "action": "classifier_rules",
            }
        )
    return suggestions


def _action_for_role(role: str) -> str:
    if role in ROLES:
        return f"Update classify_geometry() and SHORT_RULES in qwen_classify_xt_csv.py for {role}"
    return "Update mold_geometry_knowledge.md and train_from_quote_sheets token_map"


def process_bms_job(job_id: str, folder: Path) -> dict:
    """Catalog BMS job for training manifest — no AI classification."""
    steel = train_from_quote_sheets.find_workbook(folder, train_from_quote_sheets.STEEL_CANDIDATES)
    quote = train_from_quote_sheets.find_workbook(folder, train_from_quote_sheets.QUOTE_CANDIDATES)
    bom_files = [
        str(p)
        for p in folder.rglob("*")
        if p.is_file() and ("bom" in p.name.lower() or "moldbase" in p.name.lower())
    ]
    plates = []
    if steel:
        plates = train_from_quote_sheets.read_sheet_plates(steel)
    elif quote:
        plates = train_from_quote_sheets.read_sheet_plates(quote)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    manifest_path = OUT_DIR / f"{job_id}_BMS_TRAINING.json"
    payload = {
        "job_id": job_id,
        "base_type": "bms",
        "folder": str(folder),
        "bom_files": bom_files[:10],
        "steel_sheet": str(steel) if steel else "",
        "quote_sheet": str(quote) if quote else "",
        "plates_from_sheet": plates,
        "macro_guidance": (
            "Use BOM-driven pot-block flow in Module6121. "
            "RunAiBridgeClassification must skip classify and call AiBridgeNotifyBms only."
        ),
    }
    manifest_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    return {
        "job_id": job_id,
        "status": "ok_bms",
        "base_type": "bms",
        "bom_files": len(bom_files),
        "plates_found": len(plates),
        "output": str(manifest_path),
        "macro_guidance": payload["macro_guidance"],
    }


def run_full_audit(
    jobs_root: str | None = None,
    manifest_path: str | None = None,
    use_qwen: bool = False,
    qwen_model: str = "qwen3.5:9b",
) -> dict:
    """Scan training folder; optionally run Qwen per job in background (slow)."""
    if use_qwen:
        start_background_audit(jobs_root, manifest_path, qwen_model)
        return {
            "started": True,
            "background": True,
            "use_qwen": True,
            "qwen_model": qwen_model,
            "message": "Training with Qwen started in background. Keep this PC awake; poll status for progress.",
            **get_progress(),
        }
    return _run_audit_worker(jobs_root, manifest_path, use_qwen=False, qwen_model=qwen_model)


def start_background_audit(
    jobs_root: str | None,
    manifest_path: str | None,
    qwen_model: str = "qwen3.5:9b",
) -> None:
    if get_progress().get("running"):
        raise RuntimeError("Training already running")

    _write_progress(
        running=True,
        phase="starting",
        message="Starting training with Qwen...",
        use_qwen=True,
        qwen_model=qwen_model,
        error="",
        job_index=0,
        job_total=0,
        current_job="",
    )

    def _thread():
        try:
            _run_audit_worker(jobs_root, manifest_path, use_qwen=True, qwen_model=qwen_model)
        except Exception as exc:
            _write_progress(running=False, phase="error", error=str(exc), message=str(exc))
        finally:
            with _progress_lock:
                if _progress.get("phase") != "error":
                    _write_progress(running=False, phase="done", message="Training complete")

    threading.Thread(target=_thread, daemon=True).start()


def _run_audit_worker(
    jobs_root: str | None,
    manifest_path: str | None,
    use_qwen: bool,
    qwen_model: str,
) -> dict:
    job_folders: list[tuple[str, Path]] = []

    if manifest_path:
        for entry in train_from_quote_sheets._load_manifest(Path(manifest_path)):
            folder = Path(entry.get("job_folder", ""))
            if folder.exists():
                job_folders.append((entry.get("job_id") or folder.name, folder))
    elif jobs_root:
        root = Path(jobs_root)
        _write_progress(phase="scan", message=f"Discovering jobs in {root}...")
        for folder in discover_job_folders(root):
            job_folders.append((extract_job_id(folder), folder))
    else:
        return {"error": "jobs_root or manifest_path required", "jobs_processed": 0}

    _write_progress(job_total=len(job_folders), job_index=0)
    results: list[dict] = []

    for idx, (job_id, folder) in enumerate(job_folders):
        _write_progress(
            job_index=idx + 1,
            current_job=job_id,
            phase="scan",
            message=f"Scanning {job_id} ({idx + 1}/{len(job_folders)})...",
        )
        base_type, signals = detect_base_type(folder)
        entry: dict = {
            "job_id": job_id,
            "folder": str(folder),
            "base_type": base_type,
            "detection_signals": signals,
        }

        if base_type == "bms":
            entry.update(process_bms_job(job_id, folder))
            results.append(entry)
            continue

        proc = train_from_quote_sheets.process_job(job_id, folder)
        entry.update(proc)
        entry["base_type"] = base_type if base_type != "unknown" else "standard"

        xt_path = proc.get("xt_csv") or ""
        if not xt_path:
            found = train_from_quote_sheets.find_xt_csv(folder)
            xt_path = str(found) if found else ""

        if proc.get("status") == "ok" and proc.get("output") and xt_path:
            audit = audit_rules_against_correct_me(Path(proc["output"]), Path(xt_path))
            entry["audit"] = audit
            entry["rules_accuracy_pct"] = audit.get("accuracy_pct", 0)

        if use_qwen and xt_path:
            _write_progress(
                phase="qwen",
                current_job=job_id,
                message=f"Qwen ({qwen_model}) analyzing {job_id} — 15–40+ min/job on CPU...",
            )
            qwen_out = run_qwen_on_xt(Path(xt_path), model=qwen_model, timeout_minutes=0)
            entry["qwen_ran"] = qwen_out.get("qwen_ran", False)
            entry["qwen_elapsed_sec"] = qwen_out.get("elapsed_sec", 0)
            if qwen_out.get("error"):
                entry["qwen_error"] = qwen_out["error"]

            OUT_DIR.mkdir(parents=True, exist_ok=True)
            qwen_json_path = OUT_DIR / f"{job_id}_qwen_training.json"
            qwen_json_path.write_text(json.dumps(qwen_out.get("data", {}), indent=2), encoding="utf-8")
            entry["qwen_output"] = str(qwen_json_path)

            if proc.get("output") and Path(proc["output"]).exists():
                q_audit = audit_qwen_against_correct_me(Path(proc["output"]), qwen_out.get("data", {}))
                entry["qwen_audit"] = q_audit
                entry["qwen_accuracy_pct"] = q_audit.get("accuracy_pct", 0)

        results.append(entry)

    return _finalize_summary(results, jobs_root or "", use_qwen, qwen_model)


def _finalize_summary(
    results: list[dict],
    jobs_root: str,
    use_qwen: bool,
    qwen_model: str,
) -> dict:
    all_mismatches: list[dict] = []
    qwen_mismatches: list[dict] = []
    for r in results:
        all_mismatches.extend(r.get("audit", {}).get("mismatches", []))
        qwen_mismatches.extend(r.get("qwen_audit", {}).get("mismatches", []))

    suggestions = build_suggestions(all_mismatches, results)
    if qwen_mismatches:
        suggestions.append(
            {
                "priority": "high",
                "role": "qwen_vs_steel",
                "occurrences": len(qwen_mismatches),
                "suggestion": (
                    "Qwen disagreed with steel-sheet ground truth. Turn these into classifier "
                    "rules and knowledge-file examples — do not use raw Qwen at quote time."
                ),
                "examples": "; ".join(
                    f"{m['index']} expected {m['expected']} got {m['predicted']}"
                    for m in qwen_mismatches[:5]
                ),
                "action": "update_rules_from_qwen_gaps",
            }
        )

    summary = {
        "jobs_processed": len(results),
        "jobs_ok": sum(1 for r in results if r.get("status") in ("ok", "ok_bms", "ok_steel_only")),
        "jobs_skipped": sum(1 for r in results if r.get("status") == "skipped"),
        "bms_jobs": sum(1 for r in results if r.get("base_type") == "bms"),
        "standard_jobs": sum(1 for r in results if r.get("base_type") == "standard"),
        "overall_rules_accuracy_pct": _overall_accuracy(results, "audit"),
        "overall_qwen_accuracy_pct": _overall_accuracy(results, "qwen_audit"),
        "use_qwen": use_qwen,
        "qwen_model": qwen_model if use_qwen else "",
        "results": results,
        "suggestions": suggestions,
        "output_dir": str(OUT_DIR),
        "jobs_root": jobs_root,
    }

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    _write_suggestions_md(suggestions, summary)
    (DATA_DIR / "last_training_run.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return summary


def _overall_accuracy(results: list[dict], key: str = "audit") -> float:
    total_c = total_ok = 0
    for r in results:
        audit = r.get(key, {})
        if audit.get("compared"):
            total_c += audit["compared"]
            total_ok += audit.get("correct", 0)
    return round(100 * total_ok / max(total_c, 1), 1) if total_c else 0.0


def _write_suggestions_md(suggestions: list[dict], summary: dict) -> None:
    lines = [
        "# CMS Training Audit — Rule & Macro Suggestions",
        "",
        f"Jobs scanned: **{summary.get('jobs_processed', 0)}** "
        f"(BMS: {summary.get('bms_jobs', 0)}, standard: {summary.get('standard_jobs', 0)})",
        f"Rules accuracy: **{summary.get('overall_rules_accuracy_pct', 0)}%**",
    ]
    if summary.get("use_qwen"):
        lines.append(f"Qwen accuracy vs steel sheets: **{summary.get('overall_qwen_accuracy_pct', 0)}%**")
    lines.extend(
        [
            "",
            "> Target recurring CMS job types. Qwen is for training only — quotes use fast rules.",
            "",
            "## Suggestions (priority order)",
            "",
        ]
    )
    for s in suggestions:
        lines.append(f"### [{s['priority'].upper()}] {s['role']}")
        lines.append(s["suggestion"])
        if s.get("examples"):
            lines.append(f"- Examples: {s['examples']}")
        if s.get("occurrences"):
            lines.append(f"- Seen in **{s['occurrences']}** job(s)")
        lines.append(f"- **Action:** {s.get('action', 'review')}")
        lines.append("")

    lines.append("## Per-job results")
    lines.append("")
    for r in summary.get("results", []):
        st = r.get("status", "?")
        bt = r.get("base_type", "?")
        acc = r.get("rules_accuracy_pct", r.get("accuracy_pct", "n/a"))
        qacc = r.get("qwen_accuracy_pct", "")
        extra = f", Qwen {qacc}%" if qacc != "" else ""
        lines.append(f"- **{r.get('job_id')}** ({bt}) — {st}, rules {acc}%{extra}")
    SUGGESTIONS_PATH.write_text("\n".join(lines), encoding="utf-8")


def status() -> dict:
    out: dict = {"jobs_processed": 0, "output_dir": str(OUT_DIR), "suggestions": []}
    prog = get_progress()
    out.update(prog)
    if REPORT_PATH.exists():
        try:
            out.update(json.loads(REPORT_PATH.read_text(encoding="utf-8")))
        except Exception:
            pass
    elif (DATA_DIR / "last_training_run.json").exists():
        try:
            out.update(json.loads((DATA_DIR / "last_training_run.json").read_text(encoding="utf-8")))
        except Exception:
            pass
    out.update(prog)
    return out


def suggestions_markdown() -> str:
    if SUGGESTIONS_PATH.exists():
        return SUGGESTIONS_PATH.read_text(encoding="utf-8")
    return "# No training audit yet\n\nRun **Training Scan** in Settings with your TRAINING folder path.\n"
