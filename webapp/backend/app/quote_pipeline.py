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
import subprocess
import time
from pathlib import Path

from . import config, jobs

LOCAL_WORKSPACE = Path(os.environ.get("CMS_LOCAL_WORKSPACE", r"C:\CMS_Local_Workspace"))
HANDOFF_FILE = LOCAL_WORKSPACE / "cms_handoff.txt"
EMAIL_OUTPUT_FILE = LOCAL_WORKSPACE / "cms_email.txt"
STATUS_DIR = config.DATA_DIR / "quote_status"
REPO_ROOT = Path(__file__).resolve().parent.parent.parent.parent


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


def run_dme_price_lookup() -> bool:
    """Refresh DME prices in Purchased Components Prices.csv (same as launcher)."""
    script = _find_python_script("cms_price_lookup.py")
    if not script:
        set_status("_system", last_price_lookup="skipped_no_script")
        return False
    try:
        subprocess.run(
            ["python", str(script), "--all"],
            capture_output=True,
            text=True,
            timeout=600,
        )
        return True
    except Exception as e:
        set_status("_system", last_price_lookup_error=str(e))
        return False


def launch_full_quote(quote_id: str, attach_dir: str, email_info: dict | None = None) -> dict:
    """Write handoff files, run DME lookup, start CMS_Launcher /usemail."""
    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)

    info = email_info or {}
    lines = {
        "Found": "1",
        "Subject": info.get("subject", ""),
        "CustJob": info.get("cust_job", ""),
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
        message="Updating DME purchased-component prices...",
        attach_dir=attach_dir,
    )

    run_dme_price_lookup()

    set_status(quote_id, phase="launching", message="Starting SolidWorks + Module6121...")

    launcher = _find_launcher()
    launched = False
    if launcher:
        try:
            subprocess.Popen(["wscript", str(launcher), "/usemail"], close_fds=True)
            launched = True
        except Exception as e:
            set_status(quote_id, phase="error", message=str(e))
            return {"launched": False, "error": str(e)}

    # Give launcher a moment to write cms_handoff.txt with assigned C-number.
    c_num = ""
    for _ in range(20):
        time.sleep(0.5)
        handoff = _read_handoff()
        c_num = handoff.get("CNum", "") or handoff.get("QuoteNum", "").replace("-", "")
        if c_num:
            break

    if c_num:
        jobs.create_job(c_num, display_name=info.get("subject", c_num)[:80], customer=info.get("cust_job", ""))
        set_status(
            quote_id,
            phase="running",
            message="Module6121 is quoting in SolidWorks (AI runs for standard bases)...",
            c_number=c_num,
            job_id=c_num,
            handoff=handoff if c_num else {},
        )
    else:
        set_status(
            quote_id,
            phase="running",
            message="SolidWorks macro started — waiting for C-number assignment...",
            job_id=quote_id,
        )

    return {
        "launched": launched,
        "quote_id": quote_id,
        "job_id": c_num or quote_id,
        "c_number": c_num,
        "handoff_file": str(HANDOFF_FILE),
    }


def sync_completed_job(job_id: str, folder_path: str, base_type: str = "standard") -> dict:
    """Import finished macro outputs into the webapp registry."""
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(f"Completed job folder not found: {folder_path}")

    job = jobs.import_from_folder(str(folder))
    jobs.update_meta(job_id, base_type=base_type, quote_status="completed", source_folder=str(folder))

    # Auto-classify if XT exists but no classification yet (non-BMS).
    if base_type != "bms" and job.get("has_raw_csv") and not job.get("has_classification"):
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


def poll_completion(quote_id: str) -> dict:
    """Check if macro has finished by looking for output files or status."""
    status = get_status(quote_id) or {"phase": "unknown", "quote_id": quote_id}
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
        if phase in active_phases:
            out.append(poll_completion(data.get("quote_id") or path.stem))
    return out[:20]


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
