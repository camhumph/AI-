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
STATUS_DIR = config.DATA_DIR / "quote_status"
REPO_ROOT = Path(__file__).resolve().parent.parent.parent.parent

_active_launcher_procs: dict[str, subprocess.Popen] = {}
_cancelled_quotes: set[str] = set()


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
    if TRAINING_TRIGGER.exists():
        try:
            TRAINING_TRIGGER.unlink()
        except Exception:
            pass

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
    lines = {
        "Found": "1",
        "Subject": info.get("subject", ""),
        "CustJob": info.get("cust_job", ""),
        "CNum": c_number,
        "SimilarTo": info.get("similar_to", ""),
        "ShipDate": info.get("ship_date", ""),
        "Attachments": str(info.get("attachments", 0)),
        "AttachDir": attach_dir,
        "Error": "",
    }
    EMAIL_OUTPUT_FILE.write_text(
        "\n".join(f"{k}={v}" for k, v in lines.items()) + "\n",
        encoding="utf-8",
    )

    set_status(
        quote_id,
        phase="starting",
        message="Opening CAD in SolidWorks, then running Module6121.swp...",
        attach_dir=attach_dir,
        c_number=c_number or None,
    )

    run_dme_price_lookup(wait=False)

    set_status(quote_id, phase="launching", message="Opening CAD in SolidWorks first, then Module6121.swp...")

    if quote_id in _cancelled_quotes:
        set_status(quote_id, phase="cancelled", message="Quote cancelled before launch")
        return {"launched": False, "cancelled": True, "quote_id": quote_id}

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

    # Give launcher a moment to write cms_handoff.txt with assigned C-number.
    c_num = c_number
    handoff: dict = {}
    for _ in range(30):
        time.sleep(0.5)
        handoff = _read_handoff()
        c_num = handoff.get("CNum", "") or handoff.get("QuoteNum", "").replace("-", "") or c_num
        if c_num and handoff.get("JobFolder"):
            break

    if c_num:
        jobs.create_job(c_num, display_name=info.get("subject", c_num)[:80], customer=info.get("cust_job", ""))
        set_status(
            quote_id,
            phase="running",
            message=f"SolidWorks opened CAD — Module6121 quoting {c_num}...",
            c_number=c_num,
            job_id=c_num,
            handoff=handoff if c_num else {},
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
    }


def _folder_looks_like_bms(folder: Path) -> bool:
    """Detect Tempcraft / BMS pot-block jobs that must never get A/B/rail AI roles."""
    blob = folder.name.lower()
    if any(m in blob for m in ("bms", "tempcraft", "howmet", "potblock", "pot-block", "pot_block")):
        return True
    try:
        for path in folder.rglob("*"):
            if not path.is_file():
                continue
            low = path.name.lower()
            if any(
                m in low
                for m in (
                    "rfq_mb_asm",
                    "mb_asm",
                    "smed",
                    "holder block",
                    "pot block",
                    "id holder",
                    "od holder",
                )
            ):
                return True
            if path.suffix.lower() in (".csv", ".txt", ".log") and path.stat().st_size < 2_000_000:
                try:
                    text = path.read_text(encoding="utf-8", errors="ignore")[:8000].upper()
                except Exception:
                    continue
                if "ID HOLDER" in text or "OD HOLDER" in text or "SMED" in text or "POT BLOCK" in text:
                    return True
                if "BASE TYPE: POT" in text or "BOM-DRIVEN" in text:
                    return True
    except Exception:
        pass
    return False


def sync_completed_job(job_id: str, folder_path: str, base_type: str = "standard") -> dict:
    """Import finished macro outputs into the webapp registry."""
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(f"Completed job folder not found: {folder_path}")

    # Never let a mis-tagged standard sync run A/B/rail classify on pot-block jobs.
    resolved_type = (base_type or "standard").strip().lower()
    if resolved_type != "bms" and _folder_looks_like_bms(folder):
        resolved_type = "bms"

    job = jobs.import_from_folder(str(folder))
    jobs.update_meta(job_id, base_type=resolved_type, quote_status="completed", source_folder=str(folder))

    # Auto-classify if XT exists but no classification yet (non-BMS).
    if resolved_type != "bms" and job.get("has_raw_csv") and not job.get("has_classification"):
        try:
            job = jobs.classify_job(job_id, mode="rules")
        except Exception:
            pass

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
        if has_xt and (has_quote or has_purchased):
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
