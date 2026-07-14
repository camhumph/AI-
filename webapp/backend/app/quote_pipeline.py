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
LAUNCHER_STATUS_FILE = LOCAL_WORKSPACE / "cms_launcher_status.txt"
CAD_JOB_MISMATCH_FILE = LOCAL_WORKSPACE / "cms_cad_job_mismatch.txt"
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
    for p in (
        MACRO_STATUS_FILE,
        MACRO_STARTED_FILE,
        MACRO_DONE_FILE,
        MACRO_ERROR_FILE,
        LAUNCHER_STATUS_FILE,
        CANCEL_FILE,
        CAD_JOB_MISMATCH_FILE,
    ):
        _delete_if_exists(p)


def _append_quote_log(message: str) -> None:
    """Append a webapp line to CMS_Quote_Log.txt and refresh launcher status."""
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] webapp: {message}"
    try:
        LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
        with (LOCAL_WORKSPACE / "CMS_Quote_Log.txt").open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    try:
        LAUNCHER_STATUS_FILE.write_text(line + "\n", encoding="utf-8")
    except Exception:
        pass


def _popen_wscript(*args: str) -> subprocess.Popen:
    """Start wscript so SolidWorks UI can appear on the shop desktop."""
    cmd = ["wscript", "//nologo", *[str(a) for a in args]]
    kwargs: dict = {}
    if os.name == "nt":
        # Avoid close_fds on Windows — can prevent GUI child processes from starting.
        kwargs["close_fds"] = False
        # New process group; keep the desktop session so SW is visible.
        create = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        if create:
            kwargs["creationflags"] = create
    else:
        kwargs["close_fds"] = True
    return subprocess.Popen(cmd, cwd=str(LOCAL_WORKSPACE), **kwargs)


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


def _is_month_folder_job_token(digits: str) -> bool:
    """True for Ron month-folder ids like 000000007 (July), not real BMS job #s."""
    d = (digits or "").strip()
    if not d.isdigit():
        return False
    # Month folders are zero-padded to 9 digits: 000000001 .. 000000012
    if len(d) == 9 and d.startswith("000000") and 1 <= int(d) <= 12:
        return True
    # Also ignore any long all-zero / leading-zero pad that isn't a shop job id.
    if len(d) >= 8 and d.startswith("00000"):
        return True
    return False


def _cad_mismatch_scan_text(cad_path: str) -> str:
    """Filename + parent folder only — ignore month folders higher in the path."""
    raw = (cad_path or "").replace("/", "\\").strip()
    if not raw:
        return ""
    parts = [p for p in raw.split("\\") if p]
    if not parts:
        return raw
    # e.g. ...\BMS-863700114-C18611\863700102_RFQ....x_t
    return "\\".join(parts[-2:]) if len(parts) >= 2 else parts[-1]


def _cad_folder_job_mismatch_warning(
    cad_path: str,
    cust_job: str = "",
    folder_hint: str = "",
) -> str:
    """Soft warning when XT/CAD name uses a different BMS job # than the quote folder.

    Same physical mold is often named under an older BMS id (e.g. 851100021 XT
    quoted as folder 851100043) — still allow the quote; just note it.
    """
    cad = (cad_path or "").strip()
    if not cad:
        return ""
    want = _digits_only(cust_job)
    if not want:
        m = re.search(r"(?<!\d)(\d{8,})(?!\d)", folder_hint or "")
        if m and not _is_month_folder_job_token(m.group(1)):
            want = m.group(1)
    if not want or _is_month_folder_job_token(want):
        return ""
    scan = _cad_mismatch_scan_text(cad)
    tokens = [
        t
        for t in re.findall(r"\d{8,}", scan)
        if not _is_month_folder_job_token(t)
    ]
    others = [t for t in tokens if t != want]
    if not others:
        return ""
    return (
        f"Quoting different job-number CAD files than folder job {want} "
        f"(CAD uses {others[0]}). Continuing."
    )


def _find_best_xt(
    folder: Path,
    c_number: str = "",
    cust_job: str = "",
    *,
    max_depth: int = 4,
) -> Path | None:
    """Prefer Parasolid XT (then STEP/IGES) under a staged job folder; never \\base\\."""
    if not folder or not folder.is_dir():
        return None
    scored: list[tuple[int, Path]] = []
    want_c = (c_number or "").strip().upper().replace("-", "")
    if want_c and not want_c.startswith("C"):
        want_c = "C" + want_c
    want_job = _digits_only(cust_job)
    root_depth = len(folder.parts)

    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        # Cap depth so network Browse/AttachDir scans stay responsive.
        if len(p.parts) - root_depth > max_depth:
            continue
        if _is_generated_base_path(p):
            continue
        if any(part.upper() == "BASE" for part in p.parts):
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
        scored.append((score, p))
    if not scored:
        return None
    scored.sort(key=lambda t: (-t[0], -t[1].stat().st_mtime))
    return scored[0][1]


def _cad_hint_from_sources(
    c_number: str,
    source_dirs: list[str | Path],
    cust_job: str = "",
) -> dict:
    """Find XT on the network/job folder without copying (launcher stages locally)."""
    want_job = _digits_only(cust_job)
    folder_hint = " ".join(str(s) for s in source_dirs if s)
    best: Path | None = None
    for raw in source_dirs:
        src = Path(str(raw or "").strip())
        if not src.is_dir():
            continue
        # Shallow scan only — deep network rglob made the webapp feel stuck.
        hit = _find_best_xt(src, c_number=c_number, cust_job=want_job, max_depth=3)
        if hit:
            best = hit
            break
    cad_path = str(best) if best else ""
    warning = _cad_folder_job_mismatch_warning(cad_path, cust_job=want_job, folder_hint=folder_hint)
    return {"cad_path": cad_path, "warning": warning}


def stage_job_to_local_workspace(
    c_number: str,
    source_dirs: list[str | Path] | None = None,
    cust_job: str = "",
) -> dict:
    """Copy job/attach files into C:\\CMS_Local_Workspace\\C##### and return local XT path.

    Prefer the launcher's StageJobToLocalWorkspace for live quotes (one copy).
    This helper remains for callers that need an explicit local mirror.
    """
    c = (c_number or "").strip().upper().replace("-", "")
    if c and not c.startswith("C"):
        c = "C" + c
    if not c:
        return {"local_folder": "", "cad_path": "", "copied": 0, "warning": ""}

    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    dest = LOCAL_WORKSPACE / c
    copied = 0
    want_job = _digits_only(cust_job)
    folder_hint = " ".join(str(s) for s in (source_dirs or []) if s)

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
                target = dest / item.name
                if item.is_file():
                    shutil.copy2(item, target)
                    copied += 1
                elif item.is_dir():
                    shutil.copytree(item, target, dirs_exist_ok=True)
                    copied += sum(1 for _ in target.rglob("*") if _.is_file())
        except Exception:
            continue

    xt = _find_best_xt(dest, c_number=c, cust_job=want_job)
    cad_path = str(xt) if xt else ""
    warning = _cad_folder_job_mismatch_warning(cad_path, cust_job=want_job, folder_hint=folder_hint)
    return {
        "local_folder": str(dest),
        "cad_path": cad_path,
        "copied": copied,
        "warning": warning,
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


def _deploy_launcher_assets() -> dict:
    """Copy launcher scripts from the repo into C:\\CMS_Local_Workspace on Windows.

    Returns a small report so the UI/log can prove the new runners were installed.
    """
    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    report: dict = {"deployed": [], "failed": [], "runners": {}}
    for name in (
        "CMS_Launcher.vbs",
        "RunSolidWorksMacro.ps1",
        "RunModule6121.vbs",
        "RunTrainingXtLauncher.vbs",
    ):
        src = REPO_ROOT / name
        dst = LOCAL_WORKSPACE / name
        if not src.exists():
            report["failed"].append(f"{name}: missing in repo")
            continue
        try:
            shutil.copy2(src, dst)
            report["deployed"].append(name)
            try:
                report["runners"][name] = {
                    "path": str(dst),
                    "bytes": dst.stat().st_size,
                    "mtime": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(dst.stat().st_mtime)),
                }
            except Exception:
                pass
        except Exception as e:
            report["failed"].append(f"{name}: {e}")

    # Append deploy proof into the shared quote log when possible.
    try:
        log = LOCAL_WORKSPACE / "CMS_Quote_Log.txt"
        with log.open("a", encoding="utf-8") as f:
            f.write(
                f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] webapp: deployed launchers "
                f"{', '.join(report['deployed']) or '(none)'}\n"
            )
            if report["failed"]:
                f.write(
                    f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] webapp: deploy failures "
                    f"{'; '.join(report['failed'])}\n"
                )
    except Exception:
        pass
    return report


def _launch_module6121_runner() -> tuple[subprocess.Popen | None, str]:
    """Fallback: RunModule6121.vbs (RunMacro only — does not stage/open CAD).

    Prefer _start_cms_launcher for live quotes.
    """
    _deploy_launcher_assets()
    swp = LOCAL_WORKSPACE / "Module6121.swp"
    if not swp.exists():
        return None, f"missing compiled macro: {swp}"

    vbs = LOCAL_WORKSPACE / "RunModule6121.vbs"
    if not vbs.exists():
        vbs = REPO_ROOT / "RunModule6121.vbs"
    if vbs.exists():
        _append_quote_log(f"launching macro-runner-v3 via {vbs.name} (fallback, no CAD stage)")
        try:
            proc = _popen_wscript(str(vbs), str(swp))
        except Exception as e:
            return None, f"failed to start {vbs.name}: {e}"
        return proc, f"wscript {vbs} (macro-runner-v3)"

    ps1 = LOCAL_WORKSPACE / "RunSolidWorksMacro.ps1"
    if not ps1.exists():
        ps1 = REPO_ROOT / "RunSolidWorksMacro.ps1"
    if ps1.exists():
        sw_exe = r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (3)\SLDWORKS.EXE"
        progid = "SldWorks.Application.31"
        _append_quote_log(f"launching macro-runner-v3 via {ps1.name} (fallback)")
        try:
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
                    "120",
                ],
                cwd=str(LOCAL_WORKSPACE),
                close_fds=(os.name != "nt"),
            )
        except Exception as e:
            return None, f"failed to start {ps1.name}: {e}"
        return proc, f"powershell {ps1}"

    return None, "no RunModule6121.vbs or RunSolidWorksMacro.ps1 found"


def _start_cms_launcher() -> tuple[subprocess.Popen | None, str]:
    """Start CMS_Launcher.vbs /usemail — stages CAD, opens SolidWorks, runs Module6121."""
    _deploy_launcher_assets()
    launcher = _find_launcher()
    if not launcher:
        return None, "CMS_Launcher.vbs not found in C:\\CMS_Local_Workspace or repo root"
    swp = LOCAL_WORKSPACE / "Module6121.swp"
    if not swp.exists():
        _append_quote_log(f"WARNING Module6121.swp missing at {swp} — launcher will fail until recompiled")
    _append_quote_log(f"starting CMS_Launcher.vbs /usemail ({launcher})")
    try:
        proc = _popen_wscript(str(launcher), "/usemail")
    except FileNotFoundError:
        return None, "wscript.exe not found — is Windows Script Host available?"
    except Exception as e:
        return None, f"failed to start CMS_Launcher.vbs: {e}"

    # Quick sanity: if wscript dies instantly, surface it instead of fake "opening CAD".
    time.sleep(0.8)
    code = proc.poll()
    if code is not None and code != 0:
        _append_quote_log(f"CMS_Launcher.vbs exited immediately code={code}")
        return None, f"CMS_Launcher.vbs exited immediately (code {code}). See CMS_Quote_Log.txt."
    return proc, f"wscript {launcher} /usemail"


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
    # Do not keep a prior BatchCount handoff — that blocked single quotes.
    _delete_if_exists(HANDOFF_FILE)
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

    # Hint CadPath from AttachDir (no Python copy — launcher stages once to local).
    stage_sources = [
        attach_dir,
        info.get("job_folder") or "",
        info.get("root_path") or "",
    ]
    hinted = _cad_hint_from_sources(
        c_number,
        stage_sources,
        cust_job=str(info.get("cust_job") or ""),
    ) if c_number else {}
    local_cad = str(hinted.get("cad_path") or "")
    cad_warning = str(hinted.get("warning") or "")
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

    start_msg = (
        "Launcher will stage to CMS_Local_Workspace and open XT..."
        if attach_dir
        else "Opening CAD in SolidWorks, then running Module6121.swp..."
    )
    if cad_warning:
        start_msg = f"{cad_warning} {start_msg}"

    set_status(
        quote_id,
        phase="starting",
        message=start_msg,
        attach_dir=attach_dir,
        c_number=c_number or None,
        cad_path=local_cad or None,
        warning=cad_warning or None,
        cad_job_mismatch=True if cad_warning else None,
    )

    run_dme_price_lookup(wait=False)

    launch_msg = "Starting SolidWorks + Module6121..."
    if cad_warning:
        launch_msg = f"{cad_warning} {launch_msg}"

    set_status(
        quote_id,
        phase="launching",
        message=launch_msg,
        warning=cad_warning or None,
        cad_job_mismatch=True if cad_warning else None,
    )

    proc, how = _start_cms_launcher()
    launched = proc is not None
    if proc is not None:
        _active_launcher_procs[quote_id] = proc
        set_status(quote_id, phase="launching", message=f"Started {how}", runner=how)
    else:
        set_status(quote_id, phase="error", message=how, runner=how)
        return {"launched": False, "error": how, "quote_id": quote_id}

    # Brief handoff peek only — do not block the UI for 30s.
    c_num = c_number
    handoff: dict = {}
    for _ in range(8):
        time.sleep(0.25)
        if quote_id in _cancelled_quotes:
            set_status(quote_id, phase="cancelled", message="Quote cancelled before launch")
            return {"launched": False, "cancelled": True, "quote_id": quote_id}
        handoff = _read_handoff()
        c_num = handoff.get("CNum", "") or handoff.get("QuoteNum", "").replace("-", "") or c_num
        if c_num or MACRO_STARTED_FILE.exists() or MACRO_ERROR_FILE.exists():
            break
        # Detect instant launcher death after the initial check.
        if proc.poll() is not None and proc.returncode not in (None, 0) and not MACRO_STARTED_FILE.exists():
            err = f"CMS_Launcher exited code={proc.returncode} before macro start. See CMS_Quote_Log.txt."
            _append_quote_log(err)
            set_status(quote_id, phase="error", message=err)
            return {"launched": False, "error": err, "quote_id": quote_id}

    if c_num:
        jobs.create_job(c_num, display_name=info.get("subject", c_num)[:80], customer=info.get("cust_job", ""))
        started = MACRO_STARTED_FILE.exists()
        run_msg = (
            f"Module6121 acknowledged start for {c_num}..."
            if started
            else f"SolidWorks opening CAD — Module6121 quoting {c_num}..."
        )
        if cad_warning:
            run_msg = f"{cad_warning} {run_msg}"
        set_status(
            quote_id,
            phase="running",
            message=run_msg,
            c_number=c_num,
            job_id=c_num,
            handoff=handoff if c_num else {},
            macro_started=started,
            warning=cad_warning or None,
            cad_job_mismatch=True if cad_warning else None,
            cad_path=local_cad or handoff.get("CadPath") or None,
        )
    else:
        set_status(
            quote_id,
            phase="running",
            message="SolidWorks opening CAD, then Module6121.swp...",
            job_id=quote_id,
            warning=cad_warning or None,
            cad_job_mismatch=True if cad_warning else None,
        )

    return {
        "launched": launched,
        "quote_id": quote_id,
        "job_id": c_num or quote_id,
        "c_number": c_num,
        "handoff_file": str(HANDOFF_FILE),
        "macro_started": MACRO_STARTED_FILE.exists(),
        "warning": cad_warning or None,
        "cad_job_mismatch": True if cad_warning else None,
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

        # Launcher/macro stages locally — do not full-copy here (was making the UI slow).
        hinted = _cad_hint_from_sources(
            c_number,
            [
                attach_dir,
                info.get("job_folder") or item.get("job_folder") or "",
                info.get("root_path") or item.get("root_path") or "",
            ],
            cust_job=str(info.get("cust_job") or ""),
        )
        cad_path = str(hinted.get("cad_path") or info.get("cad_path") or item.get("cad_path") or "")
        cad_warning = str(hinted.get("warning") or "")

        job = {
            "CNum": c_number,
            "QuoteNum": str(info.get("quote_num") or c_number),
            "CustJob": str(info.get("cust_job") or ""),
            "SimilarTo": str(info.get("similar_to") or ""),
            "ShipDate": str(info.get("ship_date") or ""),
            "RootPath": str(info.get("root_path") or item.get("root_path") or ""),
            "JobFolder": str(info.get("job_folder") or item.get("job_folder") or ""),
            "CustomerPrefix": str(info.get("customer_prefix") or ""),
            "CustomerName": str(info.get("customer_name") or ""),
            "AttachDir": attach_dir,
            "CadPath": cad_path,
        }
        batch_jobs.append(job)
        qid = quote_id or c_number
        quote_ids.append(qid)
        queue_msg = f"Queued in batch ({len(batch_jobs)} jobs)..."
        if cad_warning:
            queue_msg = f"{cad_warning} {queue_msg}"
        set_status(
            qid,
            phase="queued",
            message=queue_msg,
            c_number=c_number,
            attach_dir=attach_dir,
            batch=True,
            cad_path=cad_path or None,
            warning=cad_warning or None,
            cad_job_mismatch=True if cad_warning else None,
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

    # Always use CMS_Launcher (stages CAD + opens SolidWorks). RunModule6121 alone
    # was leaving quotes stuck with only "webapp: deployed" and no SW start.
    launched = False
    batch_id = f"BATCH-{quote_ids[0]}" if quote_ids else "BATCH"
    how = ""

    try:
        proc, how = _start_cms_launcher()
        if proc is not None:
            _active_launcher_procs[batch_id] = proc
            launched = True
            for qid in quote_ids:
                set_status(qid, phase="launching", message=f"Started {how}", runner=how)
        else:
            # Last resort: RunMacro-only runner (no CAD open).
            proc2, how2 = _launch_module6121_runner()
            how = how2
            if proc2 is not None:
                _active_launcher_procs[batch_id] = proc2
                launched = True
                for qid in quote_ids:
                    set_status(qid, phase="launching", message=f"Started {how}", runner=how)
            else:
                err = how or how2 or "Could not start CMS_Launcher / Module6121 runner"
                for qid in quote_ids:
                    set_status(qid, phase="error", message=err)
                return {"launched": False, "error": err}
    except Exception as e:
        for qid in quote_ids:
            set_status(qid, phase="error", message=str(e))
        return {"launched": False, "error": str(e)}

    # Brief STARTED peek — do not block the UI for 30s.
    for _ in range(6):
        time.sleep(0.25)
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
                if started and i == 0
                else f"Batch {i + 1}/{len(batch_jobs)}: queued for Module6121 ({c_num})"
            ),
            c_number=c_num,
            job_id=c_num,
            batch=True,
            batch_count=len(batch_jobs),
            batch_index=i + 1,
            macro_started=started and i == 0,
            runner=how,
        )

    return {
        "launched": launched,
        "batch": True,
        "batch_count": len(batch_jobs),
        "quote_ids": quote_ids,
        "c_numbers": [j["CNum"] for j in batch_jobs],
        "handoff_file": str(HANDOFF_FILE),
        "macro_started": started,
        "runner": how,
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


def _read_tail(path: Path, max_chars: int = 2500) -> str:
    try:
        if not path.exists():
            return ""
        text = path.read_text(encoding="utf-8", errors="replace")
        return text[-max_chars:].strip()
    except Exception:
        return ""


def _current_launch_log_lines(log_tail: str) -> list[str]:
    """Use only the current launch slice of CMS_Quote_Log.txt.

    Older failed RunMacro attempts stay in the same file forever; matching those
    made the UI keep saying "Old macro-runner still running" even after v3
    deploy / a live quote.
    """
    lines = [ln for ln in (log_tail or "").splitlines() if ln.strip()]
    if not lines:
        return []
    # Prefer everything after the latest webapp deploy / runner start marker.
    start = None
    for i, ln in enumerate(lines):
        low = ln.lower()
        if (
            "webapp: deployed launchers" in low
            or "webapp: launching macro-runner-v3" in low
            or "webapp: starting cms_launcher" in low
            or "macro-runner-v3:" in low
            or "launcher process alive" in low
            or "launcher: ===== launcher" in low
        ):
            start = i
    if start is None:
        return lines[-25:]
    # Do not expand backward — that reintroduces historical ok=False noise.
    return lines[start:][-40:]


def _recent_runner_is_old(log_lines: list[str]) -> bool:
    """True only when the newest runner lines are pre-v3 (not historical)."""
    runner_lines = [
        ln
        for ln in log_lines
        if "macro-runner" in ln.lower() and "webapp:" not in ln.lower()
    ]
    if not runner_lines:
        return False
    last = runner_lines[-1].lower()
    # v3 VBS / PS1 prefixes "macro-runner-v3:"
    if "macro-runner-v3" in last:
        return False
    # Old PS1 style: "[...] macro-runner: attempt=..." without v3 in the line.
    recent = "\n".join(ln.lower() for ln in runner_lines[-8:])
    return "macro-runner:" in last and "ok=false" in recent


def _read_cad_job_mismatch() -> dict:
    """Parse cms_cad_job_mismatch.txt written by Module6121 when CAD job # ≠ folder job #."""
    out: dict = {}
    if not CAD_JOB_MISMATCH_FILE.exists():
        return out
    try:
        for line in CAD_JOB_MISMATCH_FILE.read_text(encoding="utf-8", errors="replace").splitlines():
            if "=" in line:
                k, v = line.split("=", 1)
                out[k.strip()] = v.strip()
    except Exception:
        return {}
    return out


def _cad_mismatch_warning_text(info: dict | None = None, fallback: str = "") -> str:
    info = info or {}
    expected = (info.get("ExpectedJob") or "").strip()
    cad = (info.get("CadTitle") or info.get("CadPath") or "").strip()
    if expected and cad:
        return (
            f"Quoting different job-number CAD files than folder job {expected}. "
            f"CAD: {cad}. Continuing."
        )
    msg = (info.get("Message") or "").strip()
    if msg:
        return msg
    return fallback or (
        "Quoting different job-number CAD files than the folder job. Continuing."
    )


def _collect_launch_diagnostics(status: dict) -> dict:
    """Read launcher/macro status files so the UI can show why a quote is stuck."""
    diag: dict = {
        "macro_started": MACRO_STARTED_FILE.exists(),
        "macro_done": MACRO_DONE_FILE.exists(),
        "macro_error": MACRO_ERROR_FILE.exists(),
        "handoff_exists": HANDOFF_FILE.exists(),
        "email_handoff_exists": EMAIL_OUTPUT_FILE.exists(),
    }
    status_txt = _read_tail(MACRO_STATUS_FILE, 800)
    error_txt = _read_tail(MACRO_ERROR_FILE, 1200)
    started_txt = _read_tail(MACRO_STARTED_FILE, 400)
    done_txt = _read_tail(MACRO_DONE_FILE, 400)
    launcher_status = _read_tail(LAUNCHER_STATUS_FILE, 800)
    log_tail = _read_tail(LOCAL_WORKSPACE / "CMS_Quote_Log.txt", 12000)
    if not log_tail:
        log_tail = _read_tail(Path(r"C:\Users\lenovo\Downloads\CMS_Quote_Log.txt"), 12000)
    live_log = _read_tail(LOCAL_WORKSPACE / "CMS_Module6121_Live_Log.txt", 4000)
    mismatch_info = _read_cad_job_mismatch()

    log_lines = _current_launch_log_lines(log_tail)
    recent_log = "\n".join(log_lines)

    # Ignore stale cms_launcher_status.txt left by older runners (e.g. July 13
    # "macro-runner:" lines) when the current log slice has no matching step.
    if launcher_status:
        status_core = launcher_status.strip().lstrip("\ufeff")
        low_status = status_core.lower()
        in_current = bool(status_core) and status_core in recent_log
        written_by_webapp = low_status.startswith("[") and "webapp:" in low_status
        written_by_v3 = "macro-runner-v3" in low_status or "launcher:" in low_status or "launcher process alive" in low_status
        if not in_current and not written_by_webapp and not written_by_v3:
            launcher_status = ""
        elif not in_current and "macro-runner:" in low_status and "macro-runner-v3" not in low_status:
            launcher_status = ""

    if status_txt:
        diag["macro_status"] = status_txt
    if error_txt:
        diag["macro_error_text"] = error_txt
    if started_txt:
        diag["macro_started_text"] = started_txt
    if done_txt:
        diag["macro_done_text"] = done_txt
    if launcher_status:
        diag["launcher_last_step"] = launcher_status
    if log_lines:
        diag["launcher_log_tail"] = "\n".join(log_lines[-12:])

    # Soft CAD job-number mismatch (continue quoting, but flag in red in the UI).
    blob = "\n".join(
        [
            recent_log,
            live_log or "",
            status_txt or "",
            str(status.get("warning") or ""),
            str(status.get("message") or ""),
        ]
    ).lower()
    if mismatch_info.get("Mismatch") == "1" or mismatch_info.get("ExpectedJob"):
        diag["cad_job_mismatch"] = True
        diag["cad_job_mismatch_text"] = _cad_mismatch_warning_text(mismatch_info)
    elif (
        "does not match handoff customer job" in blob
        or "cad job # mismatch" in blob
        or "differs from folder job" in blob
        or "quoting different job-number" in blob
    ):
        diag["cad_job_mismatch"] = True
        diag["cad_job_mismatch_text"] = _cad_mismatch_warning_text(
            fallback=str(status.get("warning") or "")
        )
    if diag.get("cad_job_mismatch"):
        diag["warning"] = diag["cad_job_mismatch_text"]

    # Human-readable stuck reason (current launch only — ignore old log noise)
    phase = (status.get("phase") or "").lower()
    last_log = log_lines[-1] if log_lines else ""
    if error_txt:
        err_low = error_txt.lower()
        # Soft mismatch must not mark the quote failed when we allow continuing.
        if "does not match handoff customer job" in err_low and diag.get("cad_job_mismatch"):
            pass
        else:
            diag["stuck_reason"] = f"Macro error: {error_txt.splitlines()[-1][:240]}"
    elif phase in {"launching", "running", "starting", "queued"}:
        if not MACRO_STARTED_FILE.exists():
            last = launcher_status or last_log
            low = recent_log.lower()
            if "module6121.swp not found" in low:
                diag["stuck_reason"] = "Module6121.swp missing in C:\\CMS_Local_Workspace — recompile the macro."
            elif _recent_runner_is_old(log_lines):
                diag["stuck_reason"] = (
                    "Old macro-runner still running. Pull latest code, restart the webapp, "
                    "and confirm the log shows 'macro-runner-v3'. Also recompile Module6121.swp."
                )
            elif "getmacromethods returned no entry" in low or "getmacromethods returned no entry points" in low:
                diag["stuck_reason"] = (
                    "Module6121.swp has no runnable entry points — recompile "
                    "Module6121.bas→.swp (see webapp\\COMPILE_MODULE6121.bat)."
                )
            elif ("runmacro" in low or " attempt " in low) and "ok=false" in low and "success" not in low:
                diag["stuck_reason"] = (
                    "SolidWorks RunMacro failed (ok=False). Recompile Module6121.swp "
                    "with VBA module name Module61211 (COMPILE_MODULE6121.bat), "
                    "enable macros in SW options, then retry. "
                    "Log should show Module61211.main / macro-runner-v3."
                )
            elif "did not start" in low or "could not connect" in low:
                diag["stuck_reason"] = "SolidWorks did not start or connect. Check CMS_SOLIDWORKS_EXE / SW 2023 install."
            elif "opendoc/loadfile failed" in low or "opendoc/loadfile returned nothing" in low:
                # Pre-open is best-effort; Module6121 opens CadPath from handoff.
                if "macro acknowledged started" in low:
                    pass
                elif "running macro with retry" in low:
                    diag["stuck_reason"] = (
                        "Waiting for macro STARTED (launcher OpenDoc was best-effort; "
                        "macro opens CadPath from handoff)."
                    )
                else:
                    diag["stuck_reason"] = (
                        "Waiting for macro after CAD pre-open warning. "
                        "CadPath is in cms_handoff.txt — macro should open it next."
                    )
            elif "no cad" in low or "cad: (none)" in low or "cad=no" in low:
                diag["stuck_reason"] = (
                    "No CAD/XT found before macro run (often still inside a ZIP). "
                    "Launcher now extracts ZIPs; pull latest CMS_Launcher.vbs and retry."
                )
            elif "webapp: starting cms_launcher" in low and "launcher process alive" not in low and "launcher started" not in low:
                diag["stuck_reason"] = (
                    "Webapp started CMS_Launcher.vbs but it has not logged yet. "
                    "If SolidWorks never opens, run manually: "
                    "wscript C:\\CMS_Local_Workspace\\CMS_Launcher.vbs /usemail"
                )
            elif last:
                diag["stuck_reason"] = f"Waiting for macro STARTED. Last launcher step: {last[-220:]}"
            else:
                diag["stuck_reason"] = (
                    "Launcher/macro has not written cms_macro_started.txt yet. "
                    "Check C:\\CMS_Local_Workspace\\CMS_Quote_Log.txt"
                )
        elif not MACRO_DONE_FILE.exists():
            diag["stuck_reason"] = (
                "Macro started but has not finished yet (no cms_macro_done.txt). "
                "SolidWorks may still be processing — see job CMS_Base_Export_Log.txt."
            )
    return diag


def poll_completion(quote_id: str) -> dict:
    """Check if macro has finished by looking for output files or status."""
    status = get_status(quote_id) or {"phase": "unknown", "quote_id": quote_id}
    if status.get("phase") == "cancelled":
        return status
    job_id = status.get("job_id") or status.get("c_number") or quote_id

    # Always attach live launcher/macro diagnostics while active (or on error).
    diag = _collect_launch_diagnostics(status)
    status["diagnostics"] = diag
    if diag.get("stuck_reason"):
        status["stuck_reason"] = diag["stuck_reason"]
    if diag.get("cad_job_mismatch"):
        status["cad_job_mismatch"] = True
        status["warning"] = diag.get("cad_job_mismatch_text") or diag.get("warning") or status.get("warning")
    elif diag.get("warning") and not status.get("warning"):
        status["warning"] = diag["warning"]
    if diag.get("macro_error") and status.get("phase") not in {"completed", "cancelled", "error"}:
        # Soft CAD mismatch notice must not flip the quote to error.
        err_blob = (diag.get("macro_error_text") or "").lower()
        if diag.get("cad_job_mismatch") and "does not match handoff customer job" in err_blob:
            pass
        else:
            status["phase"] = "error"
            status["message"] = diag.get("stuck_reason") or "Macro reported an error"
            try:
                set_status(
                    quote_id,
                    phase="error",
                    message=status["message"],
                    stuck_reason=status.get("stuck_reason"),
                    warning=status.get("warning"),
                    cad_job_mismatch=status.get("cad_job_mismatch"),
                    diagnostics=diag,
                )
            except Exception:
                pass

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
                    status["diagnostics"] = diag
                    if diag.get("stuck_reason"):
                        status["stuck_reason"] = diag["stuck_reason"]
                except Exception as e:
                    status["sync_error"] = str(e)
            status["outputs_found"] = True
            status["local_folder"] = str(local)
        else:
            status["outputs_found"] = has_xt
            status["local_folder"] = str(local)
            # Surface last lines of the job export log when still running.
            job_log = _read_tail(local / "CMS_Base_Export_Log.txt", 1500)
            if job_log:
                status.setdefault("diagnostics", diag)["job_log_tail"] = "\n".join(
                    [ln for ln in job_log.splitlines() if ln.strip()][-8:]
                )

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
        if phase in active_phases or phase == "completed" or phase == "error":
            if phase == "completed" and data.get("dismissed"):
                continue
            out.append(poll_completion(data.get("quote_id") or path.stem))
    return out[:20]
