"""Scan a training folder (mixed BMS + standard jobs), build CORRECT_ME files,
compare against the live rules classifier, and emit actionable suggestions.

Designed for folders like C:\\Users\\lenovo\\Downloads\\TRAINING where each
subfolder may contain steel sheets, BOM PDFs, XT exports, and SolidWorks files.
"""
from __future__ import annotations

import csv
import json
import platform
import re
import subprocess
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
_cancel_event = threading.Event()
_active_ollama_proc: subprocess.Popen | None = None
_active_ollama_lock = threading.Lock()

_progress: dict = {
    "running": False,
    "phase": "idle",
    "current_job": "",
    "job_index": 0,
    "job_total": 0,
    "message": "",
    "detail": "",
    "use_qwen": False,
    "qwen_model": "",
    "export_xt": False,
    "error": "",
    "cancelled": False,
    "qwen_thinking": False,
    "qwen_elapsed_sec": 0,
    "started_at": "",
    "elapsed_sec": 0,
}

# Import sibling modules
if str(BASE.parent) not in sys.path:
    sys.path.insert(0, str(BASE.parent))

from geometry_classifier import train_from_quote_sheets  # noqa: E402
from geometry_classifier import training_xt_export  # noqa: E402
from geometry_classifier.qwen_classify_xt_csv import (  # noqa: E402
    ROLES,
    classify_geometry,
    read_rows,
)

C_NUMBER_RE = re.compile(r"\b(C\d{4,6})\b", re.I)
JOB_NUMBER_RE = re.compile(r"\b(\d{6,8})\b")

# Folder-name prefixes that mean STANDARD mold base (customer jobs).
STANDARD_FOLDER_PREFIXES = (
    "dynacast-",
    "industrialmold-",
    "itwmedical-",
    "itw-",
    "pcs-",
    "dme-",
)

# Folder-name markers that mean BMS / pot-block (macro BOM path).
BMS_FOLDER_MARKERS = (
    "bms-",
    "pot-block",
    "potblock",
    "pot_block",
)

# File-level hints only used when folder name is ambiguous.
BMS_FILE_MARKERS = (
    "moldbase",
    "mold base",
    "mold-base",
    "pot block",
    "pot-block",
    "potblock",
)

STANDARD_FILE_MARKERS = (
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
    """Return ('bms'|'standard'|'unknown', list of reasons).

    Folder name wins: BMS-... = BMS; Dynacast-/IndustrialMold-/ITW... = standard.
    Generic BOM/moldbase files alone must NOT force BMS (standard jobs often have them).
    """
    folder_low = folder.name.lower().replace("_", "-")

    # 1) Explicit BMS folder name (e.g. BMS-863700102-C18608)
    if any(m in folder_low for m in BMS_FOLDER_MARKERS) or folder_low.startswith("bms"):
        return "bms", [f"folder name: {folder.name}"]

    # 2) Explicit customer / standard folder name
    if any(folder_low.startswith(p) for p in STANDARD_FOLDER_PREFIXES):
        return "standard", [f"folder name: {folder.name}"]

    signals_bms: list[str] = []
    signals_std: list[str] = []

    for path in folder.rglob("*"):
        if not path.is_file():
            continue
        low = path.name.lower()
        low_dash = low.replace("_", "-")
        if any(m in low for m in BMS_FILE_MARKERS):
            signals_bms.append(path.name)
        if any(m in low_dash for m in STANDARD_FILE_MARKERS):
            signals_std.append(path.name)

    # Steel / quote sheet without BMS folder name → treat as standard for training
    has_steel = any(
        ("steel" in p.name.lower() or "quote" in p.name.lower())
        and p.suffix.lower() in (".xls", ".xlsx", ".xlsm")
        for p in folder.iterdir()
        if p.is_file()
    )
    if has_steel and not signals_bms:
        return "standard", ["steel/quote sheet present"] + signals_std[:6]
    if has_steel and signals_std:
        return "standard", ["steel/quote sheet + standard tokens"] + signals_std[:6]

    if signals_std and not signals_bms:
        return "standard", signals_std[:8]
    if signals_bms and not signals_std:
        return "bms", signals_bms[:8]
    if signals_bms and signals_std:
        return "standard", signals_std[:8] + ["(ambiguous files — default standard)"]
    if has_steel:
        return "standard", ["steel/quote sheet present"]
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
        # After a hard cancel, ignore worker updates that try to keep running=True
        if _cancel_event.is_set() and fields.get("running") is True:
            return
        if _cancel_event.is_set() and _progress.get("phase") == "cancelled":
            # Only allow final cancelled/error writes
            if fields.get("phase") not in (None, "cancelled", "error", "done"):
                fields = {k: v for k, v in fields.items() if k in ("detail", "message", "elapsed_sec", "updated_at")}
                if not fields:
                    return
        _progress.update(fields)
        if _cancel_event.is_set() and _progress.get("phase") == "cancelled":
            _progress["running"] = False
            _progress["qwen_thinking"] = False
        _progress["updated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
        started = _progress.get("started_at") or ""
        if started and _progress.get("running"):
            try:
                t0 = time.mktime(time.strptime(started, "%Y-%m-%dT%H:%M:%S"))
                _progress["elapsed_sec"] = int(time.time() - t0)
            except Exception:
                pass
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
        out = dict(_progress)
    started = out.get("started_at") or ""
    if started and out.get("running"):
        try:
            t0 = time.mktime(time.strptime(started, "%Y-%m-%dT%H:%M:%S"))
            out["elapsed_sec"] = int(time.time() - t0)
        except Exception:
            pass
    return out


def is_cancelled() -> bool:
    return _cancel_event.is_set()


def request_cancel() -> dict:
    """Force-stop training immediately: kill Ollama + XT export wait, clear running flag."""
    global _active_ollama_proc
    _cancel_event.set()
    killed: list[str] = []

    with _active_ollama_lock:
        proc = _active_ollama_proc
        if proc is not None and proc.poll() is None:
            try:
                proc.kill()
                killed.append("ollama")
            except Exception:
                pass
        _active_ollama_proc = None

    try:
        training_xt_export.request_xt_cancel()
        killed.append("xt_export")
    except Exception:
        pass

    # Hard-kill leftover ollama CLI on Windows if still running
    if platform.system() == "Windows":
        for name in ("ollama.exe",):
            try:
                subprocess.run(
                    ["taskkill", "/F", "/IM", name],
                    capture_output=True,
                    timeout=5,
                )
                killed.append(name)
            except Exception:
                pass

    _write_progress(
        running=False,
        phase="cancelled",
        cancelled=True,
        qwen_thinking=False,
        qwen_elapsed_sec=0,
        message="Training cancelled",
        detail="Stopped immediately" + (f" · killed {', '.join(killed)}" if killed else ""),
        error="",
    )
    return get_progress()


def _check_cancelled(results: list[dict] | None = None) -> bool:
    if not is_cancelled():
        return False
    _write_progress(
        running=False,
        phase="cancelled",
        qwen_thinking=False,
        message="Training cancelled by user",
        detail=f"Stopped after {len(results or [])} job(s)",
    )
    return True


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
    """Run Ollama/Qwen on one XT export with live thinking heartbeats (training only)."""
    global _active_ollama_proc
    from geometry_classifier.qwen_classify_xt_csv import (  # noqa: E402
        build_prompt,
        classify_geometry,
        extract_json,
        read_rows,
    )

    rows = read_rows(str(xt_path), include_names=True)
    n_rows = len(rows)
    if n_rows > 5000:
        rows = rows[:5000]
        n_rows = 5000

    _write_progress(
        phase="qwen",
        qwen_thinking=False,
        qwen_elapsed_sec=0,
        detail=f"Building prompt from {n_rows} CAD components…",
        message=f"Preparing Qwen ({model}) for current job…",
    )
    prompt = build_prompt(rows, str(xt_path), long_knowledge=True)
    prompt_chars = len(prompt)

    import shutil

    if shutil.which("ollama") is None:
        return {
            "qwen_ran": False,
            "qwen_model": model,
            "error": "Ollama is not installed or not on PATH",
            "elapsed_sec": 0,
            "data": classify_geometry(rows),
        }

    started = time.time()
    _write_progress(
        phase="qwen",
        qwen_thinking=True,
        qwen_elapsed_sec=0,
        detail=f"Ollama loading {model} · prompt {prompt_chars:,} chars · {n_rows} parts",
        message=f"Qwen is thinking on {Path(xt_path).name}…",
    )

    cmd = ["ollama", "run", model]
    try:
        proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except FileNotFoundError:
        return {
            "qwen_ran": False,
            "qwen_model": model,
            "error": "Ollama is not installed or not on PATH",
            "elapsed_sec": 0,
            "data": classify_geometry(rows),
        }

    with _active_ollama_lock:
        _active_ollama_proc = proc

    stdout_chunks: list[str] = []
    stderr_chunks: list[str] = []
    cancelled = False

    def _reader(stream, bucket: list[str]) -> None:
        try:
            while True:
                chunk = stream.read(4096)
                if not chunk:
                    break
                bucket.append(chunk)
        except Exception:
            pass

    t_out = threading.Thread(target=_reader, args=(proc.stdout, stdout_chunks), daemon=True)
    t_err = threading.Thread(target=_reader, args=(proc.stderr, stderr_chunks), daemon=True)
    t_out.start()
    t_err.start()

    try:
        assert proc.stdin is not None
        proc.stdin.write(prompt)
        proc.stdin.close()
    except Exception as exc:
        proc.kill()
        with _active_ollama_lock:
            _active_ollama_proc = None
        return {
            "qwen_ran": False,
            "qwen_model": model,
            "error": f"Failed to send prompt to Ollama: {exc}",
            "elapsed_sec": round(time.time() - started, 1),
            "data": classify_geometry(rows),
        }

    timeout_sec = timeout_minutes * 60 if timeout_minutes > 0 else None
    last_beat = 0.0
    while True:
        if is_cancelled():
            cancelled = True
            try:
                proc.kill()
            except Exception:
                pass
            break
        rc = proc.poll()
        elapsed = time.time() - started
        if timeout_sec and elapsed > timeout_sec:
            try:
                proc.kill()
            except Exception:
                pass
            with _active_ollama_lock:
                _active_ollama_proc = None
            return {
                "qwen_ran": False,
                "qwen_model": model,
                "error": f"Ollama timed out after {timeout_minutes} minutes",
                "elapsed_sec": round(elapsed, 1),
                "data": classify_geometry(rows),
            }
        if rc is not None:
            break
        if elapsed - last_beat >= 2.0:
            last_beat = elapsed
            mins = int(elapsed // 60)
            secs = int(elapsed % 60)
            out_len = sum(len(c) for c in stdout_chunks)
            _write_progress(
                phase="qwen",
                qwen_thinking=True,
                qwen_elapsed_sec=int(elapsed),
                detail=(
                    f"Qwen thinking… {mins}m {secs:02d}s · model {model} · "
                    f"{n_rows} parts · output {out_len:,} chars so far"
                ),
                message=f"Qwen is thinking on current job ({mins}m {secs:02d}s)…",
            )
        time.sleep(0.5)

    t_out.join(timeout=5)
    t_err.join(timeout=5)
    with _active_ollama_lock:
        _active_ollama_proc = None

    elapsed = round(time.time() - started, 1)
    if cancelled:
        _write_progress(qwen_thinking=False, detail="Qwen cancelled")
        return {
            "qwen_ran": False,
            "qwen_model": model,
            "error": "cancelled",
            "elapsed_sec": elapsed,
            "data": {},
            "cancelled": True,
        }

    raw = "".join(stdout_chunks)
    err = "".join(stderr_chunks).strip()
    if proc.returncode not in (0, None) and not raw.strip():
        _write_progress(qwen_thinking=False, detail=f"Qwen failed: {err[:200]}")
        return {
            "qwen_ran": False,
            "qwen_model": model,
            "error": err or f"Ollama exit code {proc.returncode}",
            "elapsed_sec": elapsed,
            "data": classify_geometry(rows),
        }

    try:
        data = extract_json(raw)
        n_cls = len(data.get("classifications", [])) if isinstance(data, dict) else 0
        _write_progress(
            qwen_thinking=False,
            qwen_elapsed_sec=int(elapsed),
            detail=f"Qwen finished in {elapsed}s · {n_cls} classifications",
            message="Qwen finished — scoring against steel sheet…",
        )
        return {
            "qwen_ran": True,
            "qwen_model": model,
            "elapsed_sec": elapsed,
            "data": data,
        }
    except Exception as exc:
        fallback = classify_geometry(rows)
        _write_progress(qwen_thinking=False, detail=f"Qwen parse failed: {exc}")
        return {
            "qwen_ran": False,
            "qwen_model": model,
            "error": str(exc),
            "elapsed_sec": elapsed,
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
    export_xt: bool = True,
) -> dict:
    """Scan training folder. Always runs in background so the UI can show live status
    (XT export and Qwen both take minutes per job)."""
    start_background_audit(
        jobs_root, manifest_path, qwen_model, export_xt=export_xt, use_qwen=use_qwen
    )
    return {
        "started": True,
        "background": True,
        "use_qwen": use_qwen,
        "qwen_model": qwen_model if use_qwen else "",
        "export_xt": export_xt,
        "message": "Training started. Watch the status bar for current task.",
        **get_progress(),
    }


def start_background_audit(
    jobs_root: str | None,
    manifest_path: str | None,
    qwen_model: str = "qwen3.5:9b",
    export_xt: bool = True,
    use_qwen: bool = True,
) -> None:
    if get_progress().get("running"):
        raise RuntimeError("Training already running")

    _cancel_event.clear()
    started_at = time.strftime("%Y-%m-%dT%H:%M:%S")
    _write_progress(
        running=True,
        phase="starting",
        message="Starting training scan…",
        detail="Discovering job folders",
        use_qwen=use_qwen,
        qwen_model=qwen_model if use_qwen else "",
        export_xt=export_xt,
        error="",
        cancelled=False,
        qwen_thinking=False,
        qwen_elapsed_sec=0,
        job_index=0,
        job_total=0,
        current_job="",
        started_at=started_at,
        elapsed_sec=0,
    )

    def _thread():
        try:
            summary = _run_audit_worker(
                jobs_root, manifest_path, use_qwen=use_qwen, qwen_model=qwen_model, export_xt=export_xt
            )
            if is_cancelled() or summary.get("cancelled"):
                _write_progress(
                    running=False,
                    phase="cancelled",
                    qwen_thinking=False,
                    message="Training cancelled",
                    detail=f"Partial results saved ({summary.get('jobs_processed', 0)} jobs)",
                )
            else:
                _write_progress(
                    running=False,
                    phase="done",
                    qwen_thinking=False,
                    message="Training complete",
                    detail=f"{summary.get('jobs_processed', 0)} jobs · rules {summary.get('overall_rules_accuracy_pct', 0)}%",
                )
        except Exception as exc:
            _write_progress(
                running=False,
                phase="error",
                qwen_thinking=False,
                error=str(exc),
                message=str(exc),
            )

    threading.Thread(target=_thread, daemon=True).start()


def _run_audit_worker(
    jobs_root: str | None,
    manifest_path: str | None,
    use_qwen: bool,
    qwen_model: str,
    export_xt: bool = True,
) -> dict:
    job_folders: list[tuple[str, Path]] = []

    if manifest_path:
        for entry in train_from_quote_sheets._load_manifest(Path(manifest_path)):
            folder = Path(entry.get("job_folder", ""))
            if folder.exists():
                job_folders.append((entry.get("job_id") or folder.name, folder))
    elif jobs_root:
        root = Path(jobs_root)
        _write_progress(
            phase="scan",
            message=f"Discovering jobs in {root}…",
            detail="Walking TRAINING folder tree",
        )
        for folder in discover_job_folders(root):
            if is_cancelled():
                break
            job_folders.append((extract_job_id(folder), folder))
    else:
        return {"error": "jobs_root or manifest_path required", "jobs_processed": 0}

    if is_cancelled():
        return {"cancelled": True, "jobs_processed": 0, "results": []}

    _write_progress(
        job_total=len(job_folders),
        job_index=0,
        detail=f"Found {len(job_folders)} job folder(s)",
        message=f"Found {len(job_folders)} jobs — starting scan…",
    )
    results: list[dict] = []

    for idx, (job_id, folder) in enumerate(job_folders):
        if _check_cancelled(results):
            summary = _finalize_summary(results, jobs_root or "", use_qwen, qwen_model, export_xt)
            summary["cancelled"] = True
            return summary

        _write_progress(
            job_index=idx + 1,
            current_job=job_id,
            phase="scan",
            qwen_thinking=False,
            message=f"Scanning {job_id} ({idx + 1}/{len(job_folders)})…",
            detail=f"Detecting base type · {folder.name}",
        )
        base_type, signals = detect_base_type(folder)
        entry: dict = {
            "job_id": job_id,
            "folder": str(folder),
            "base_type": base_type,
            "detection_signals": signals,
        }

        if export_xt and not train_from_quote_sheets.find_xt_csv(folder):
            if training_xt_export.folder_has_cad(folder):
                _write_progress(
                    job_index=idx + 1,
                    current_job=job_id,
                    phase="xt_export",
                    message=f"SolidWorks exporting CAD dimensions for {job_id}…",
                    detail=f"Opening CAD in {folder.name} ({idx + 1}/{len(job_folders)})",
                )
                if is_cancelled():
                    results.append(entry)
                    summary = _finalize_summary(results, jobs_root or "", use_qwen, qwen_model, export_xt)
                    summary["cancelled"] = True
                    return summary
                xt_result = training_xt_export.ensure_xt_export(folder, job_id)
                entry["xt_export"] = xt_result
                if xt_result.get("status") == "cancelled" or is_cancelled():
                    results.append(entry)
                    summary = _finalize_summary(results, jobs_root or "", use_qwen, qwen_model, export_xt)
                    summary["cancelled"] = True
                    return summary
                _write_progress(
                    detail=f"XT export: {xt_result.get('status')} · {xt_result.get('reason') or xt_result.get('message') or ''}"
                )
            else:
                entry["xt_export"] = {
                    "ok": False,
                    "status": "skipped",
                    "reason": "no CAD files — cannot auto-export XT",
                }

        if base_type == "bms":
            _write_progress(detail=f"{job_id} = BMS — cataloging BOM/steel (no AI classify)")
            entry.update(process_bms_job(job_id, folder))
            results.append(entry)
            continue

        _write_progress(detail=f"{job_id} = {base_type} — matching steel sheet to CAD…")
        proc = train_from_quote_sheets.process_job(job_id, folder)
        entry.update(proc)
        entry["base_type"] = base_type if base_type != "unknown" else "standard"

        xt_path = proc.get("xt_csv") or ""
        if not xt_path:
            found = train_from_quote_sheets.find_xt_csv(folder)
            xt_path = str(found) if found else ""

        if proc.get("status") == "ok" and proc.get("output") and xt_path:
            _write_progress(detail=f"Auditing rules vs steel sheet for {job_id}…")
            audit = audit_rules_against_correct_me(Path(proc["output"]), Path(xt_path))
            entry["audit"] = audit
            entry["rules_accuracy_pct"] = audit.get("accuracy_pct", 0)

        if use_qwen and xt_path:
            if is_cancelled():
                results.append(entry)
                summary = _finalize_summary(results, jobs_root or "", use_qwen, qwen_model, export_xt)
                summary["cancelled"] = True
                return summary
            _write_progress(
                phase="qwen",
                current_job=job_id,
                job_index=idx + 1,
                message=f"Starting Qwen ({qwen_model}) on {job_id}…",
                detail=f"XT: {Path(xt_path).name}",
            )
            qwen_out = run_qwen_on_xt(Path(xt_path), model=qwen_model, timeout_minutes=0)
            entry["qwen_ran"] = qwen_out.get("qwen_ran", False)
            entry["qwen_elapsed_sec"] = qwen_out.get("elapsed_sec", 0)
            if qwen_out.get("error"):
                entry["qwen_error"] = qwen_out["error"]
            if qwen_out.get("cancelled"):
                results.append(entry)
                summary = _finalize_summary(results, jobs_root or "", use_qwen, qwen_model, export_xt)
                summary["cancelled"] = True
                return summary

            OUT_DIR.mkdir(parents=True, exist_ok=True)
            qwen_json_path = OUT_DIR / f"{job_id}_qwen_training.json"
            qwen_json_path.write_text(json.dumps(qwen_out.get("data", {}), indent=2), encoding="utf-8")
            entry["qwen_output"] = str(qwen_json_path)

            if proc.get("output") and Path(proc["output"]).exists():
                q_audit = audit_qwen_against_correct_me(Path(proc["output"]), qwen_out.get("data", {}))
                entry["qwen_audit"] = q_audit
                entry["qwen_accuracy_pct"] = q_audit.get("accuracy_pct", 0)

        results.append(entry)

    return _finalize_summary(results, jobs_root or "", use_qwen, qwen_model, export_xt)


def _finalize_summary(
    results: list[dict],
    jobs_root: str,
    use_qwen: bool,
    qwen_model: str,
    export_xt: bool = True,
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
        "export_xt": export_xt,
        "xt_exported_jobs": sum(
            1 for r in results if r.get("xt_export", {}).get("status") == "exported"
        ),
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
    """Live progress always wins while a scan is running or just cancelled."""
    out: dict = {"jobs_processed": 0, "output_dir": str(OUT_DIR), "suggestions": []}
    prog = get_progress()
    # Don't let a stale report overwrite a live/cancelled progress state
    if prog.get("running") or prog.get("phase") in ("cancelled", "error", "starting", "scan", "xt_export", "qwen"):
        out.update(prog)
        if not prog.get("running") and prog.get("phase") == "done" and REPORT_PATH.exists():
            try:
                report = json.loads(REPORT_PATH.read_text(encoding="utf-8"))
                for k in ("results", "suggestions", "jobs_processed", "jobs_ok", "overall_rules_accuracy_pct", "overall_qwen_accuracy_pct", "bms_jobs", "standard_jobs", "xt_exported_jobs"):
                    if k in report:
                        out[k] = report[k]
            except Exception:
                pass
        return out

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
