"""Launch Module6121 RunTrainingXtExport for training folders missing XT CSV.

Windows + SolidWorks only. The webapp training scan writes cms_training_xt.txt,
then starts RunSolidWorksMacro.ps1 (RunMacro2 + OLE retry) with procedure main.
main() and RunFromLauncher() both route to RunTrainingXtExport when that handoff
file exists. Falls back to RunTrainingXtLauncher.vbs if PowerShell is unavailable.
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import threading
import time
from pathlib import Path

LOCAL_WORKSPACE = Path(os.environ.get("CMS_LOCAL_WORKSPACE", r"C:\CMS_Local_Workspace"))
TRAINING_HANDOFF = LOCAL_WORKSPACE / "cms_training_xt.txt"
TRAINING_DONE = LOCAL_WORKSPACE / "cms_training_xt_done.txt"
REPO_ROOT = Path(__file__).resolve().parent.parent

# SolidWorks 2023 on this shop PC is the "(3)" install.
# Plain "SOLIDWORKS\SLDWORKS.exe" opens 2025 — do not use that.
SW_EXE_DEFAULT = r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (3)\SLDWORKS.EXE"
SW_PROGID_DEFAULT = "SldWorks.Application.31"  # 31 = SW 2023 (32=2024, 33=2025)

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


def folder_has_cad(folder: Path, max_depth: int = 2) -> bool:
    """Shallow CAD check — avoid slow network rglob."""
    if not folder.exists():
        return False

    def walk(p: Path, depth: int) -> bool:
        try:
            for child in p.iterdir():
                if child.name.startswith("~$") or child.name.startswith("."):
                    continue
                if child.is_file() and child.suffix.lower() in CAD_PRIORITY:
                    return True
                if child.is_dir() and depth < max_depth and walk(child, depth + 1):
                    return True
        except (PermissionError, OSError):
            return False
        return False

    return walk(folder, 0)


_active_xt_proc: subprocess.Popen | None = None
_xt_cancel = False
_xt_launch_lock = threading.Lock()


def request_xt_cancel() -> None:
    """Break any in-progress XT export wait loop and kill the macro launcher."""
    global _xt_cancel, _active_xt_proc
    _xt_cancel = True
    try:
        TRAINING_DONE.write_text(
            "Status=CANCELLED\nMessage=Cancelled by user\nXtCsv=\nPartCount=0\n",
            encoding="utf-8",
        )
    except Exception:
        pass
    proc = _active_xt_proc
    if proc is not None and proc.poll() is None:
        try:
            proc.kill()
        except Exception:
            pass
    _active_xt_proc = None


def _sw_paths() -> tuple[str, str]:
    return (
        os.environ.get("CMS_SOLIDWORKS_EXE", SW_EXE_DEFAULT),
        os.environ.get("CMS_SOLIDWORKS_PROGID", SW_PROGID_DEFAULT),
    )


def deploy_runtime_files(*, force_macro: bool = False) -> None:
    """Copy macro + launchers from repo into CMS_Local_Workspace."""
    if not is_windows():
        return
    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    always_copy = {"Module6121.swb", "Module6121.bas"}
    for name in ("Module6121.swb", "Module6121.bas", "RunSolidWorksMacro.ps1", "RunTrainingXtLauncher.vbs"):
        src = REPO_ROOT / name
        dst = LOCAL_WORKSPACE / name
        if not src.exists():
            continue
        if force_macro and name in always_copy:
            shutil.copy2(src, dst)
        elif not dst.exists() or src.stat().st_mtime > dst.stat().st_mtime:
            shutil.copy2(src, dst)


def find_training_launcher() -> Path | None:
    deploy_runtime_files(force_macro=True)
    for p in (
        LOCAL_WORKSPACE / "RunTrainingXtLauncher.vbs",
        REPO_ROOT / "RunTrainingXtLauncher.vbs",
    ):
        if p.exists():
            return p
    return None


def find_macro_path() -> Path | None:
    """Use the compiled SolidWorks macro package (.swp) — that is what RunMacro expects."""
    deploy_runtime_files(force_macro=True)
    swp = LOCAL_WORKSPACE / "Module6121.swp"
    if swp.exists():
        return swp
    # Last resort only if the shop has never saved a .swp to the workspace.
    swb = LOCAL_WORKSPACE / "Module6121.swb"
    if swb.exists():
        return swb
    repo_swb = REPO_ROOT / "Module6121.swb"
    if repo_swb.exists():
        return repo_swb
    return None


def find_runner_script() -> Path | None:
    deploy_runtime_files(force_macro=True)
    for p in (LOCAL_WORKSPACE / "RunSolidWorksMacro.ps1", REPO_ROOT / "RunSolidWorksMacro.ps1"):
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


def _wait_for_idle_launcher(timeout_sec: int = 900) -> bool:
    """Do not stack multiple SolidWorks macro launchers on top of each other."""
    global _active_xt_proc
    if _active_xt_proc is None or _active_xt_proc.poll() is not None:
        _active_xt_proc = None
        return True
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if _xt_cancel:
            return False
        if _active_xt_proc.poll() is not None:
            _active_xt_proc = None
            return True
        time.sleep(0.5)
    try:
        _active_xt_proc.kill()
    except Exception:
        pass
    _active_xt_proc = None
    return True


def _launch_macro_runner(macro: Path, log_file: str) -> tuple[subprocess.Popen | None, bool]:
    """Start macro via PowerShell (RunMacro2). Falls back to VBS launcher.

    Returns (process, started). When PowerShell succeeds, started=True and process is None.
    """
    sw_exe, sw_progid = _sw_paths()
    runner = find_runner_script()
    if runner:
        for procedure in ("main", "RunFromLauncher"):
            try:
                result = subprocess.run(
                    [
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
                        procedure,
                        "-LogFile",
                        log_file,
                    ],
                    cwd=str(LOCAL_WORKSPACE),
                    capture_output=True,
                    text=True,
                    timeout=180,
                )
                if result.returncode == 0:
                    return None, True
            except Exception:
                continue

    launcher = find_training_launcher()
    if not launcher:
        return None, False
    proc = subprocess.Popen(
        ["wscript", str(launcher)],
        close_fds=True,
        cwd=str(LOCAL_WORKSPACE),
    )
    return proc, False


def export_xt_via_macro(folder: Path, job_id: str, timeout_sec: int = 1200) -> dict:
    """Run SolidWorks macro to write XT_Export_CAD_Dimensions.csv into folder."""
    global _active_xt_proc, _xt_cancel

    with _xt_launch_lock:
        _xt_cancel = False
        folder = folder.resolve()
        if not is_windows():
            return {"ok": False, "status": "skipped", "reason": "SolidWorks XT export requires Windows"}

        if not folder_has_cad(folder):
            return {"ok": False, "status": "skipped", "reason": "no CAD files in folder"}

        macro = find_macro_path()
        if not macro:
            return {"ok": False, "status": "error", "reason": "Module6121.swp not found in C:\\CMS_Local_Workspace"}

        if not _wait_for_idle_launcher():
            return {"ok": False, "status": "cancelled", "reason": "Cancelled by user"}

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

        log_file = str(LOCAL_WORKSPACE / "CMS_Training_XT_Launcher_Log.txt")

        try:
            launcher_proc, macro_started = _launch_macro_runner(macro, log_file)
            _active_xt_proc = launcher_proc
        except Exception as exc:
            _active_xt_proc = None
            return {"ok": False, "status": "error", "reason": str(exc)}

        if not macro_started and launcher_proc is None:
            return {
                "ok": False,
                "status": "error",
                "reason": "Could not start SolidWorks macro runner (see CMS_Training_XT_Launcher_Log.txt)",
            }

        deadline = time.time() + timeout_sec
        launcher_deadline = time.time() + 900 if launcher_proc else 0
        while time.time() < deadline:
            if _xt_cancel:
                _active_xt_proc = None
                return {"ok": False, "status": "cancelled", "reason": "Cancelled by user"}
            done = _parse_done_file(TRAINING_DONE)
            status = done.get("Status", "").upper()
            if status:
                _active_xt_proc = None
                xt_path = Path(done.get("XtCsv") or output_csv)
                if status == "OK" and xt_path.exists():
                    return {
                        "ok": True,
                        "status": "exported",
                        "xt_csv": str(xt_path),
                        "part_count": done.get("PartCount", ""),
                        "message": done.get("Message", ""),
                    }
                if status == "CANCELLED":
                    return {"ok": False, "status": "cancelled", "reason": "Cancelled by user"}
                return {
                    "ok": False,
                    "status": "error",
                    "reason": done.get("Message") or f"XT export returned {status}",
                }
            if launcher_proc and time.time() < launcher_deadline:
                if launcher_proc.poll() is not None:
                    if launcher_proc.returncode not in (0, None):
                        log_hint = ""
                        log_path = Path(log_file)
                        if log_path.exists():
                            try:
                                log_hint = log_path.read_text(encoding="utf-8", errors="replace")[-500:]
                            except Exception:
                                pass
                        _active_xt_proc = None
                        return {
                            "ok": False,
                            "status": "error",
                            "reason": (
                                f"SolidWorks macro launcher failed (exit {launcher_proc.returncode}). "
                                f"See {log_file}. {log_hint[-250:]}"
                            ),
                        }
                    launcher_proc = None
                    _active_xt_proc = None
            time.sleep(0.4)

        _active_xt_proc = None
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
