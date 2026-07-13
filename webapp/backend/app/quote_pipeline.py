"""One-click quote orchestrator — runs the full CMS pipeline without manual uploads.

Flow (matches the old CMS_Launcher + Module6121 + DME lookup process):
  1. Gather files from email attachments or a selected folder
  2. Write cms_email.txt (+ read back cms_handoff.txt for assigned C-number)
  3. Run cms_price_lookup.py --all (DME prices into CSV)
  4. Launch CMS_Launcher.vbs /usemail → SolidWorks + Module6121
  5. Module6121 calls /api/vba/classify mid-run for non-BMS bases (AI in the middle)
  6. On completion, Module6121 POSTs /api/vba/job-complete → sync artifacts to webapp
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import time
from pathlib import Path

from . import config, jobs

LOCAL_WORKSPACE = Path(os.environ.get("CMS_LOCAL_WORKSPACE", r"C:\CMS_Local_Workspace"))
HANDOFF_FILE = LOCAL_WORKSPACE / "cms_handoff.txt"
EMAIL_OUTPUT_FILE = LOCAL_WORKSPACE / "cms_email.txt"
CANCEL_FILE = LOCAL_WORKSPACE / "cms_quote_cancel.txt"
TRAINING_TRIGGER = LOCAL_WORKSPACE / "cms_training_xt.txt"
MACRO_STATUS_FILE = LOCAL_WORKSPACE / "cms_macro_status.txt"
MACRO_STARTED_FILE = LOCAL_WORKSPACE / "cms_macro_started.txt"
MACRO_DONE_FILE = LOCAL_WORKSPACE / "cms_macro_done.txt"
MACRO_ERROR_FILE = LOCAL_WORKSPACE / "cms_macro_error.txt"
STATUS_DIR = config.DATA_DIR / "quote_status"
REPO_ROOT = Path(__file__).resolve().parent.parent.parent.parent

_active_launcher_procs: dict[str, subprocess.Popen] = {}
_cancelled_quotes: set[str] = set()


def _delete_if_exists(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except Exception:
        pass


def _clear_macro_launch_status_files() -> None:
    for p in (MACRO_STATUS_FILE, MACRO_STARTED_FILE, MACRO_DONE_FILE, MACRO_ERROR_FILE):
        _delete_if_exists(p)


def _clear_quote_cancel(quote_id: str) -> None:
    """Allow re-quote after Cancel — sticky cancel must not block a new launch."""
    qid = (quote_id or "").strip()
    if qid:
        _cancelled_quotes.discard(qid)
        # Also drop common C-number variants the UI may have used as quote_id.
        _cancelled_quotes.discard(qid.upper())
        _cancelled_quotes.discard(qid.replace("-", ""))
        if qid.upper().startswith("C") and "-" not in qid:
            _cancelled_quotes.discard("C-" + qid[1:])
    try:
        if CANCEL_FILE.exists():
            text = CANCEL_FILE.read_text(encoding="utf-8", errors="ignore")
            if not qid or qid in text or qid.upper() in text.upper():
                CANCEL_FILE.unlink(missing_ok=True)
    except Exception:
        pass


_XT_EXTS = {".x_t", ".x_b", ".step", ".stp", ".igs", ".iges"}


def _is_generated_base_path(path: Path | str) -> bool:
    u = str(path).replace("/", "\\").upper()
    return "\\BASE\\" in u or u.endswith("\\BASE")


def _digits_only(s: str) -> str:
    return "".join(ch for ch in (s or "") if ch.isdigit())


def _is_foreign_job_cad(path: Path | str, c_number: str = "", cust_job: str = "") -> bool:
    """Reject CAD that clearly belongs to another C##### or BMS job id."""
    text = str(path).replace("/", "\\").upper()
    want_c = (c_number or "").strip().upper().replace("-", "")
    if want_c and not want_c.startswith("C"):
        want_c = "C" + want_c
    want_job = _digits_only(cust_job)

    c_tokens = re.findall(r"(?<![A-Z0-9])C(\d{4,6})(?!\d)", text)
    if want_c and c_tokens:
        want_digits = want_c[1:] if want_c.startswith("C") else want_c
        if want_digits not in c_tokens and any(t != want_digits for t in c_tokens):
            return True

    job_tokens = re.findall(r"\d{8,}", text)
    if want_job and job_tokens:
        if want_job not in job_tokens and any(t != want_job for t in job_tokens):
            return True
    return False


def _find_best_xt(
    folder: Path,
    c_number: str = "",
    cust_job: str = "",
) -> Path | None:
    """Prefer Parasolid XT (then STEP/IGES) under a staged job folder; never \\base\\."""
    if not folder or not folder.is_dir():
        return None
    scored: list[tuple[int, Path]] = []
    want_c = (c_number or "").strip().upper().replace("-", "")
    if want_c and not want_c.startswith("C"):
        want_c = "C" + want_c
    want_job = _digits_only(cust_job)

    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        if _is_generated_base_path(p):
            continue
        if any(part.upper() == "BASE" for part in p.parts):
            continue
        if _is_foreign_job_cad(p, want_c, want_job):
            continue
        ext = p.suffix.lower()
        if ext not in _XT_EXTS:
            continue
        score = 120 if ext in {".x_t", ".x_b"} else 110 if ext in {".step", ".stp"} else 100
        name_u = p.name.upper()
        path_u = str(p).upper()
        if want_c and want_c in path_u:
            score += 500
        if want_job and want_job in path_u:
            score += 500
        if "RFQ" in name_u and score < 400:
            score -= 40
        if score < 400 and ("MOLD_BASE" in name_u or "OUTSOURCE" in name_u):
            score -= 20
        scored.append((score, p))
    if not scored:
        return None
    scored.sort(key=lambda t: (-t[0], -t[1].stat().st_mtime))
    return scored[0][1]


def stage_job_to_local_workspace(
    c_number: str,
    source_dirs: list[str | Path] | None = None,
    cust_job: str = "",
) -> dict:
    """Copy job/attach files into C:\\CMS_Local_Workspace\\C##### and return local XT path.

    Used so SolidWorks opens the XT from the local workspace, not the network share.
    """
    c = (c_number or "").strip().upper().replace("-", "")
    if c and not c.startswith("C"):
        c = "C" + c
    if not c:
        return {"local_folder": "", "cad_path": "", "copied": 0}

    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    dest = LOCAL_WORKSPACE / c
    copied = 0
    want_job = _digits_only(cust_job)

    # Fresh stage each launch so we don't mix old BASE exports with new files.
    if dest.exists():
        try:
            shutil.rmtree(dest, ignore_errors=True)
        except Exception:
            pass
    dest.mkdir(parents=True, exist_ok=True)

    for raw in source_dirs or []:
        src = Path(str(raw or "").strip())
        if not src.is_dir():
            continue
        if _is_generated_base_path(src):
            continue
        try:
            for item in src.iterdir():
                if item.name.startswith("."):
                    continue
                if item.is_dir() and item.name.upper() == "BASE":
                    continue
                if _is_foreign_job_cad(item, c, want_job):
                    continue
                target = dest / item.name
                if item.is_file():
                    shutil.copy2(item, target)
                    copied += 1
                elif item.is_dir():
                    shutil.copytree(
                        item,
                        target,
                        dirs_exist_ok=True,
                        ignore=lambda _dir, names: [
                            n
                            for n in names
                            if n.upper() == "BASE"
                            or _is_foreign_job_cad(Path(_dir) / n, c, want_job)
                        ],
                    )
                    copied += sum(1 for _ in target.rglob("*") if _.is_file())
        except Exception:
            continue

    xt = _find_best_xt(dest, c_number=c, cust_job=want_job)
    return {
        "local_folder": str(dest),
        "cad_path": str(xt) if xt else "",
        "copied": copied,
    }


def _write_handoff_atomic(text: str) -> None:
    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    tmp = HANDOFF_FILE.with_suffix(HANDOFF_FILE.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    _delete_if_exists(HANDOFF_FILE)
    tmp.replace(HANDOFF_FILE)


def _write_batch_handoff(jobs: list[dict[str, str]]) -> None:
    """Write BatchCount + JobN.* handoff for sequential multi-quote runs."""
    lines = [f"BatchCount={len(jobs)}", ""]
    for i, job in enumerate(jobs, start=1):
        for key in (
            "CNum",
            "QuoteNum",
            "CustJob",
            "SimilarTo",
            "ShipDate",
            "RootPath",
            "JobFolder",
            "CustomerPrefix",
            "CustomerName",
            "AttachDir",
            "CadPath",
        ):
            lines.append(f"Job{i}.{key}={job.get(key, '')}")
        lines.append("")
    _write_handoff_atomic("\n".join(lines) + "\n")


def _ensure_status_dir() -> None:
    STATUS_DIR.mkdir(parents=True, exist_ok=True)


def _status_path(quote_id: str) -> Path:
    safe = re.sub(r"[^\w\-]", "_", quote_id)
    return STATUS_DIR / f"{safe}.json"


def set_status(quote_id: str, **fields) -> dict:
    _ensure_status_dir()
    path = _status_path(quote_id)
    current: dict = {}
    if path.exists():
        try:
            current = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    current.update(fields)
    current["updated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    path.write_text(json.dumps(current, indent=2), encoding="utf-8")
    return current


def get_status(quote_id: str) -> dict | None:
    path = _status_path(quote_id)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _read_handoff() -> dict[str, str]:
    out: dict[str, str] = {}
    if not HANDOFF_FILE.exists():
        return out
    for line in HANDOFF_FILE.read_text(encoding="utf-8", errors="replace").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            out[k.strip()] = v.strip()
    return out


def _find_launcher() -> Path | None:
    for p in (
        LOCAL_WORKSPACE / "CMS_Launcher.vbs",
        REPO_ROOT / "CMS_Launcher.vbs",
    ):
        if p.exists():
            return p
    return None


def _find_python_script(name: str) -> Path | None:
    for p in (
        LOCAL_WORKSPACE / name,
        REPO_ROOT / name,
    ):
        if p.exists():
            return p
    return None


def _deploy_launcher_assets() -> None:
    """Copy launcher scripts from the repo into C:\\CMS_Local_Workspace on Windows."""
    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    for name in ("CMS_Launcher.vbs", "RunSolidWorksMacro.ps1", "RunTrainingXtLauncher.vbs"):
        src = REPO_ROOT / name
        if src.exists():
            try:
                shutil.copy2(src, LOCAL_WORKSPACE / name)
            except Exception:
                pass


def run_dme_price_lookup(wait: bool = False) -> bool:
    """Refresh DME prices in Purchased Components Prices.csv (same as launcher).

    By default starts in the background so SolidWorks can launch immediately.
    Pass wait=True only when prices must be ready before the macro reads them.
    """
    script = _find_python_script("cms_price_lookup.py")
    if not script:
        set_status("_system", last_price_lookup="skipped_no_script")
        return False
    try:
        if wait:
            subprocess.run(
                ["python", str(script), "--all"],
                capture_output=True,
                text=True,
                timeout=120,
            )
        else:
            subprocess.Popen(
                ["python", str(script), "--all"],
                close_fds=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        return True
    except Exception as e:
        set_status("_system", last_price_lookup_error=str(e))
        return False


def launch_full_quote(quote_id: str, attach_dir: str, email_info: dict | None = None) -> dict:
    """Write handoff files, run DME lookup, start CMS_Launcher /usemail."""
    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    _deploy_launcher_assets()
    _delete_if_exists(TRAINING_TRIGGER)
    _clear_macro_launch_status_files()
    _clear_quote_cancel(quote_id)

    info = email_info or {}
    c_number = (info.get("c_number") or "").strip().upper()
    if not c_number:
        # Prefer C##### from subject / cust_job / attach path (BMS-...-C18603)
        blob = " ".join(
            [
                str(info.get("subject", "")),
                str(info.get("cust_job", "")),
                str(attach_dir),
                str(quote_id),
            ]
        )
        m = re.search(r"[-_]C(\d{4,6})\b", blob, re.I) or re.search(r"\bC[- ]?(\d{4,6})\b", blob, re.I)
        if m:
            c_number = "C" + m.group(1)

    # Pull attach/job files into C:\CMS_Local_Workspace\C##### and prefer local XT.
    stage_sources = [
        attach_dir,
        info.get("job_folder") or "",
        info.get("root_path") or "",
    ]
    staged = stage_job_to_local_workspace(
        c_number,
        stage_sources,
        cust_job=str(info.get("cust_job") or ""),
    ) if c_number else {}
    local_cad = str(staged.get("cad_path") or "")
    local_folder = str(staged.get("local_folder") or "")
    # Only advertise local stage when we actually found CAD or copied files.
    if not local_cad and int(staged.get("copied") or 0) <= 0:
        local_folder = ""

    lines = {
        "Found": "1",
        "Subject": info.get("subject", ""),
        "CustJob": info.get("cust_job", ""),
        "CNum": c_number,
        "SimilarTo": info.get("similar_to", ""),
        "ShipDate": info.get("ship_date", ""),
        "Attachments": str(info.get("attachments", 0)),
        "AttachDir": attach_dir,
        "LocalJobFolder": local_folder,
        "CadPath": local_cad,
        "Error": "",
    }
    EMAIL_OUTPUT_FILE.write_text(
        "\n".join(f"{k}={v}" for k, v in lines.items()) + "\n",
        encoding="utf-8",
    )

    set_status(
        quote_id,
        phase="starting",
        message=(
            f"Staged to {local_folder}; opening local XT..."
            if local_cad
            else "Opening CAD in SolidWorks, then running Module6121.swp..."
        ),
        attach_dir=attach_dir,
        c_number=c_number or None,
        local_job_folder=local_folder or None,
        cad_path=local_cad or None,
    )

    run_dme_price_lookup(wait=False)

    set_status(
        quote_id,
        phase="launching",
        message=(
            f"Opening local XT then Module6121.swp..."
            if local_cad
            else "Opening CAD in SolidWorks first, then Module6121.swp..."
        ),
    )

    launcher = _find_launcher()
    launched = False
    if launcher:
        try:
            proc = subprocess.Popen(["wscript", str(launcher), "/usemail"], close_fds=True)
            _active_launcher_procs[quote_id] = proc
            launched = True
        except Exception as e:
            set_status(quote_id, phase="error", message=str(e))
            return {"launched": False, "error": str(e)}

    # Give launcher a moment to write cms_handoff.txt with assigned C-number,
    # and preferably wait until VBA acknowledges STARTED.
    c_num = c_number
    handoff: dict = {}
    for _ in range(60):
        time.sleep(0.5)
        if quote_id in _cancelled_quotes:
            set_status(quote_id, phase="cancelled", message="Quote cancelled before launch")
            return {"launched": False, "cancelled": True, "quote_id": quote_id}
        handoff = _read_handoff()
        c_num = handoff.get("CNum", "") or handoff.get("QuoteNum", "").replace("-", "") or c_num
        if MACRO_STARTED_FILE.exists() or MACRO_ERROR_FILE.exists():
            break
        if c_num and handoff.get("JobFolder"):
            # Handoff ready; keep waiting briefly for STARTED acknowledgment.
            continue

    if c_num:
        jobs.create_job(c_num, display_name=info.get("subject", c_num)[:80], customer=info.get("cust_job", ""))
        started = MACRO_STARTED_FILE.exists()
        set_status(
            quote_id,
            phase="running",
            message=(
                f"Module6121 acknowledged start for {c_num}..."
                if started
                else f"SolidWorks opening CAD — Module6121 quoting {c_num}..."
            ),
            c_number=c_num,
            job_id=c_num,
            handoff=handoff if c_num else {},
            macro_started=started,
        )
    else:
        set_status(
            quote_id,
            phase="running",
            message="SolidWorks opening CAD, then Module6121.swp...",
            job_id=quote_id,
        )

    return {
        "launched": launched,
        "quote_id": quote_id,
        "job_id": c_num or quote_id,
        "c_number": c_num,
        "handoff_file": str(HANDOFF_FILE),
        "macro_started": MACRO_STARTED_FILE.exists(),
    }


def launch_batch_quotes(items: list[dict]) -> dict:
    """Launch multiple quotes as one sequential SolidWorks batch.

    Each item: quote_id, attach_dir, and optional email fields
    (subject, cust_job, c_number, similar_to, ship_date, root_path, job_folder, cad_path).
    """
    if not items:
        return {"launched": False, "error": "No quotes in batch"}

    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    _deploy_launcher_assets()
    _delete_if_exists(TRAINING_TRIGGER)
    _clear_macro_launch_status_files()

    batch_jobs: list[dict[str, str]] = []
    quote_ids: list[str] = []

    for item in items:
        quote_id = str(item.get("quote_id") or item.get("c_number") or "").strip()
        attach_dir = str(item.get("attach_dir") or "").strip()
        info = item.get("email_info") or item
        c_number = (info.get("c_number") or item.get("c_number") or quote_id or "").strip().upper()
        if not c_number:
            blob = " ".join(
                [
                    str(info.get("subject", "")),
                    str(info.get("cust_job", "")),
                    attach_dir,
                    quote_id,
                ]
            )
            m = re.search(r"[-_]C(\d{4,6})\b", blob, re.I) or re.search(r"\bC[- ]?(\d{4,6})\b", blob, re.I)
            if m:
                c_number = "C" + m.group(1)
        if not c_number:
            continue

        # Re-quote after Cancel is allowed — clear sticky cancel for this id.
        _clear_quote_cancel(quote_id)
        _clear_quote_cancel(c_number)

        staged = stage_job_to_local_workspace(
            c_number,
            [
                attach_dir,
                info.get("job_folder") or item.get("job_folder") or "",
                info.get("root_path") or item.get("root_path") or "",
            ],
            cust_job=str(info.get("cust_job") or ""),
        )
        cad_path = str(staged.get("cad_path") or "")
        if not cad_path:
            raw_cad = str(info.get("cad_path") or item.get("cad_path") or "")
            if raw_cad and not _is_foreign_job_cad(raw_cad, c_number, str(info.get("cust_job") or "")):
                cad_path = raw_cad
        local_folder = str(staged.get("local_folder") or "")

        job = {
            "CNum": c_number,
            "QuoteNum": str(info.get("quote_num") or c_number),
            "CustJob": str(info.get("cust_job") or ""),
            "SimilarTo": str(info.get("similar_to") or ""),
            "ShipDate": str(info.get("ship_date") or ""),
            "RootPath": str(info.get("root_path") or item.get("root_path") or ""),
            "JobFolder": str(info.get("job_folder") or item.get("job_folder") or local_folder or ""),
            "CustomerPrefix": str(info.get("customer_prefix") or ""),
            "CustomerName": str(info.get("customer_name") or ""),
            "AttachDir": attach_dir or local_folder,
            "CadPath": cad_path,
        }
        batch_jobs.append(job)
        qid = quote_id or c_number
        quote_ids.append(qid)
        set_status(
            qid,
            phase="queued",
            message=f"Queued in batch ({len(batch_jobs)} jobs)...",
            c_number=c_number,
            attach_dir=attach_dir,
            batch=True,
        )
        jobs.create_job(c_number, display_name=str(info.get("subject", c_number))[:80], customer=str(info.get("cust_job", "")))

    if not batch_jobs:
        return {"launched": False, "error": "No valid C-numbers in batch"}

    _write_batch_handoff(batch_jobs)
    EMAIL_OUTPUT_FILE.write_text(
        "\n".join(
            [
                f"Found={len(batch_jobs)}",
                f"BatchCount={len(batch_jobs)}",
                f"CNum={','.join(j['CNum'] for j in batch_jobs)}",
                f"AttachDir={batch_jobs[0].get('AttachDir', '')}",
                "Error=",
            ]
        )
        + "\n",
        encoding="utf-8",
    )

    run_dme_price_lookup(wait=False)

    for qid in quote_ids:
        set_status(qid, phase="launching", message=f"Launching batch of {len(batch_jobs)} quotes...")

    # Prefer PowerShell runner (retry + STARTED ack). Fall back to VBS /usemail for single.
    ps1 = LOCAL_WORKSPACE / "RunSolidWorksMacro.ps1"
    if not ps1.exists():
        ps1 = REPO_ROOT / "RunSolidWorksMacro.ps1"
    swp = LOCAL_WORKSPACE / "Module6121.swp"
    launched = False
    batch_id = f"BATCH-{quote_ids[0]}" if quote_ids else "BATCH"

    try:
        if ps1.exists() and swp.exists():
            sw_exe = r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (3)\SLDWORKS.EXE"
            progid = "SldWorks.Application.31"
            proc = subprocess.Popen(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(ps1),
                    "-MacroPath",
                    str(swp),
                    "-SwExe",
                    sw_exe,
                    "-ProgId",
                    progid,
                    "-Procedure",
                    "main",
                    "-TimeoutSeconds",
                    "90",
                ],
                close_fds=True,
            )
            _active_launcher_procs[batch_id] = proc
            launched = True
        else:
            launcher = _find_launcher()
            if launcher and len(batch_jobs) == 1:
                proc = subprocess.Popen(["wscript", str(launcher), "/usemail"], close_fds=True)
                _active_launcher_procs[batch_id] = proc
                launched = True
            elif not swp.exists():
                return {"launched": False, "error": f"Missing compiled macro: {swp}"}
            else:
                return {"launched": False, "error": f"Missing macro runner: {ps1}"}
    except Exception as e:
        for qid in quote_ids:
            set_status(qid, phase="error", message=str(e))
        return {"launched": False, "error": str(e)}

    # Wait briefly for STARTED acknowledgment.
    for _ in range(60):
        time.sleep(0.5)
        if MACRO_STARTED_FILE.exists() or MACRO_ERROR_FILE.exists():
            break

    started = MACRO_STARTED_FILE.exists()
    for i, qid in enumerate(quote_ids):
        c_num = batch_jobs[i]["CNum"]
        set_status(
            qid,
            phase="running",
            message=(
                f"Batch {i + 1}/{len(batch_jobs)}: Module6121 started ({c_num})"
                if started
                else f"Batch {i + 1}/{len(batch_jobs)}: waiting for Module6121 ({c_num})"
            ),
            c_number=c_num,
            job_id=c_num,
            batch=True,
            batch_count=len(batch_jobs),
            batch_index=i + 1,
            macro_started=started,
        )

    return {
        "launched": launched,
        "batch": True,
        "batch_count": len(batch_jobs),
        "quote_ids": quote_ids,
        "c_numbers": [j["CNum"] for j in batch_jobs],
        "handoff_file": str(HANDOFF_FILE),
        "macro_started": started,
    }


def _folder_looks_like_bms(folder: Path) -> bool:
    """Detect Tempcraft / BMS pot-block jobs that must never get A/B/rail AI roles."""
    return jobs._folder_looks_like_bms(folder)


def sync_completed_job(job_id: str, folder_path: str, base_type: str = "standard") -> dict:
    """Import finished macro outputs into the webapp registry."""
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(f"Completed job folder not found: {folder_path}")

    # Never let a mis-tagged standard sync run A/B/rail classify on pot-block jobs.
    resolved_type = (base_type or "standard").strip().lower()
    if resolved_type != "bms" and (
        _folder_looks_like_bms(folder) or jobs._xt_looks_like_pot_block(folder)
    ):
        resolved_type = "bms"

    job = jobs.import_from_folder(str(folder))
    job_dir = jobs._job_dir(job_id)
    if resolved_type != "bms" and jobs._xt_looks_like_pot_block(job_dir):
        resolved_type = "bms"
    jobs.update_meta(job_id, base_type=resolved_type, quote_status="completed", source_folder=str(folder))

    # Auto-classify if XT exists but no classification yet (non-BMS).
    if resolved_type != "bms" and job.get("has_raw_csv") and not job.get("has_classification"):
        try:
            job = jobs.classify_job(job_id, mode="rules")
        except Exception:
            pass

    job = jobs.get_job(job_id)

    set_status(
        job_id,
        phase="completed",
        message="Quote finished — all files synced.",
        job_id=job_id,
        folder_path=str(folder),
    )
    return job


def find_local_job_folder(c_number: str) -> Path | None:
    """Locate C:\\CMS_Local_Workspace\\C##### after macro runs."""
    c = c_number.replace("-", "").upper()
    if not c.startswith("C"):
        c = "C" + c
    candidates = [
        LOCAL_WORKSPACE / c,
        LOCAL_WORKSPACE / c_number,
    ]
    for p in candidates:
        if p.exists() and p.is_dir():
            return p
    # Newest folder matching C-number pattern
    if LOCAL_WORKSPACE.exists():
        matches = sorted(
            (d for d in LOCAL_WORKSPACE.iterdir() if d.is_dir() and re.search(rf"\b{re.escape(c)}\b", d.name, re.I)),
            key=lambda d: d.stat().st_mtime,
            reverse=True,
        )
        if matches:
            return matches[0]
    return None


def is_quote_cancelled(quote_id: str) -> bool:
    if quote_id in _cancelled_quotes:
        return True
    status = get_status(quote_id)
    return bool(status and status.get("phase") == "cancelled")


def cancel_quote(quote_id: str) -> dict:
    """Stop a background quote run and mark it cancelled in the status file."""
    _cancelled_quotes.add(quote_id)
    proc = _active_launcher_procs.pop(quote_id, None)
    if proc is not None and proc.poll() is None:
        try:
            proc.kill()
        except Exception:
            pass

    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    try:
        CANCEL_FILE.write_text(f"QuoteId={quote_id}\n", encoding="utf-8")
    except Exception:
        pass

    current = get_status(quote_id) or {"quote_id": quote_id}
    set_status(
        quote_id,
        phase="cancelled",
        message="Quote cancelled by user",
        job_id=current.get("job_id") or quote_id,
        dismissed=True,
    )
    return get_status(quote_id) or {"phase": "cancelled", "quote_id": quote_id}


def _macro_log_says_done(folder: Path) -> bool:
    """True when Module6121 wrote a DONE line (standard or BMS)."""
    for name in ("CMS_Base_Export_Log.txt", "CMS_Training_XT_Log.txt"):
        p = folder / name
        if not p.is_file():
            continue
        try:
            text = p.read_text(encoding="utf-8", errors="ignore")[-8000:].upper()
        except Exception:
            continue
        if "DONE JOB" in text or "DONE ACTIVE CAD QUOTE" in text:
            return True
        if "TOTAL JOB TIME:" in text or "TOTAL ACTIVE RUN TIME:" in text:
            return True
    return False


def poll_completion(quote_id: str) -> dict:
    """Check if macro has finished by looking for output files or status."""
    status = get_status(quote_id) or {"phase": "unknown", "quote_id": quote_id}
    if status.get("phase") == "cancelled":
        return status
    job_id = status.get("job_id") or status.get("c_number") or quote_id

    local = find_local_job_folder(job_id)
    if local:
        has_xt = (local / "XT_Export_CAD_Dimensions.csv").exists()
        has_quote = any(local.glob("*quote*.xls*")) or any(local.glob("*Quote*.xls*"))
        has_purchased = (local / "Purchased Components Quote.csv").exists()
        has_steel = any(local.glob("*steel*.xls*")) or any(local.glob("*J000*.xls*"))
        has_done_log = _macro_log_says_done(local)
        # Standard jobs may write steel/quote under slightly different names;
        # also accept DONE log so the UI does not stay on "running" forever.
        ready = has_xt and (has_quote or has_purchased or has_steel or has_done_log)
        if ready:
            if status.get("phase") != "completed":
                try:
                    sync_completed_job(job_id, str(local))
                    status = get_status(quote_id) or status
                except Exception as e:
                    status["sync_error"] = str(e)
            status["outputs_found"] = True
            status["local_folder"] = str(local)
        else:
            status["outputs_found"] = has_xt
            status["local_folder"] = str(local)

    return status


def list_active_quotes() -> list[dict]:
    _ensure_status_dir()
    active_phases = {"queued", "starting", "launching", "running"}
    out: list[dict] = []
    for path in sorted(STATUS_DIR.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        phase = data.get("phase", "")
        if phase in active_phases or phase == "completed":
            if phase == "completed" and data.get("dismissed"):
                continue
            out.append(poll_completion(data.get("quote_id") or path.stem))
    return out[:20]
