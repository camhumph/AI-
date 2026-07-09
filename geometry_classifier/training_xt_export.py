"""Launch Module6121 RunTrainingXtExport for training folders missing XT CSV.

Windows + SolidWorks only. The webapp training scan writes cms_training_xt.txt,
starts RunSolidWorksMacro.ps1 with procedure RunTrainingXtExport, and polls
cms_training_xt_done.txt for completion.
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import time
from pathlib import Path

LOCAL_WORKSPACE = Path(os.environ.get("CMS_LOCAL_WORKSPACE", r"C:\CMS_Local_Workspace"))
TRAINING_HANDOFF = LOCAL_WORKSPACE / "cms_training_xt.txt"
TRAINING_DONE = LOCAL_WORKSPACE / "cms_training_xt_done.txt"
REPO_ROOT = Path(__file__).resolve().parent.parent

SW_EXE_DEFAULT = r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe"
SW_PROGID_DEFAULT = "SldWorks.Application.30"

CAD_PRIORITY: dict[str, int] = {
    ".sldasm": 100,
    ".easm": 90,
    ".asm": 85,
    ".step": 80,
    ".stp": 80,
    ".x_t": 70,
    ".x_b": 70,
    ".igs": 60,
    ".iges": 60,
    ".sldprt": 50,
    ".prt": 45,
}


def is_windows() -> bool:
    return platform.system() == "Windows"


def folder_has_cad(folder: Path) -> bool:
    if not folder.exists():
        return False
    for path in folder.rglob("*"):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        if path.suffix.lower() in CAD_PRIORITY:
            return True
    return False


def deploy_runtime_files() -> None:
    """Copy macro + PS runner from repo into CMS_Local_Workspace when newer."""
    if not is_windows():
        return
    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    for name in ("Module6121.swb", "RunSolidWorksMacro.ps1"):
        src = REPO_ROOT / name
        dst = LOCAL_WORKSPACE / name
        if src.exists() and (not dst.exists() or src.stat().st_mtime > dst.stat().st_mtime):
            shutil.copy2(src, dst)


def find_macro_path() -> Path | None:
    deploy_runtime_files()
    for p in (
        LOCAL_WORKSPACE / "Module6121.swb",
        REPO_ROOT / "Module6121.swb",
        LOCAL_WORKSPACE / "Module6121.swp",
        REPO_ROOT / "Module6121.swp",
    ):
        if p.exists():
            return p
    return None


def find_runner_script() -> Path | None:
    for p in (REPO_ROOT / "RunSolidWorksMacro.ps1", LOCAL_WORKSPACE / "RunSolidWorksMacro.ps1"):
        if p.exists():
            return p
    return None


def _parse_done_file(path: Path) -> dict[str, str]:
    out: dict[str, str] = {}
    if not path.exists():
        return out
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            out[k.strip()] = v.strip()
    return out


def export_xt_via_macro(folder: Path, job_id: str, timeout_sec: int = 1200) -> dict:
    """Run SolidWorks macro to write XT_Export_CAD_Dimensions.csv into folder."""
    folder = folder.resolve()
    if not is_windows():
        return {"ok": False, "status": "skipped", "reason": "SolidWorks XT export requires Windows"}

    if not folder_has_cad(folder):
        return {"ok": False, "status": "skipped", "reason": "no CAD files in folder"}

    macro = find_macro_path()
    runner = find_runner_script()
    if not macro:
        return {"ok": False, "status": "error", "reason": "Module6121.swb/.swp not found"}
    if not runner:
        return {"ok": False, "status": "error", "reason": "RunSolidWorksMacro.ps1 not found"}

    output_csv = folder / "XT_Export_CAD_Dimensions.csv"
    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    if TRAINING_DONE.exists():
        TRAINING_DONE.unlink()

    TRAINING_HANDOFF.write_text(
        "\n".join(
            [
                f"JobFolder={folder}",
                f"JobId={job_id}",
                f"OutputCsv={output_csv}",
                f"DoneFile={TRAINING_DONE}",
            ]
        )
        + "\n",
        encoding="utf-8",
    )

    sw_exe = os.environ.get("CMS_SOLIDWORKS_EXE", SW_EXE_DEFAULT)
    sw_progid = os.environ.get("CMS_SOLIDWORKS_PROGID", SW_PROGID_DEFAULT)
    log_file = str(LOCAL_WORKSPACE / "CMS_Training_XT_Log.txt")

    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(runner),
        "-MacroPath",
        str(macro),
        "-SwExe",
        sw_exe,
        "-ProgId",
        sw_progid,
        "-Procedure",
        "RunTrainingXtExport",
        "-LogFile",
        log_file,
    ]

    try:
        subprocess.Popen(cmd, close_fds=True)
    except Exception as exc:
        return {"ok": False, "status": "error", "reason": str(exc)}

    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        done = _parse_done_file(TRAINING_DONE)
        status = done.get("Status", "").upper()
        if status:
            xt_path = Path(done.get("XtCsv") or output_csv)
            if status == "OK" and xt_path.exists():
                return {
                    "ok": True,
                    "status": "exported",
                    "xt_csv": str(xt_path),
                    "part_count": done.get("PartCount", ""),
                    "message": done.get("Message", ""),
                }
            return {
                "ok": False,
                "status": "error",
                "reason": done.get("Message") or f"XT export returned {status}",
            }
        time.sleep(2)

    return {
        "ok": False,
        "status": "timeout",
        "reason": f"Timed out after {timeout_sec}s waiting for SolidWorks macro",
    }


def ensure_xt_export(folder: Path, job_id: str, timeout_sec: int = 1200) -> dict:
    """Return existing XT or launch macro to create one."""
    from geometry_classifier import train_from_quote_sheets

    existing = train_from_quote_sheets.find_xt_csv(folder)
    if existing:
        return {"ok": True, "status": "exists", "xt_csv": str(existing)}
    return export_xt_via_macro(folder, job_id, timeout_sec=timeout_sec)
