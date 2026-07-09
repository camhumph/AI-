import mimetypes
import os
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import config, credentials, email_service, jobs, pricing, quote_pipeline, vba_bridge

app = FastAPI(title="CMS AI Quoting")

# Local-only tool: the server binds 127.0.0.1 (see run instructions/README);
# CORS stays open for the localhost Vite dev server on another port.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# --------------------------------------------------------------------------
# Health
# --------------------------------------------------------------------------
@app.get("/api/health")
def health():
    return {
        "ok": True,
        "jobs_root": str(config.JOBS_ROOT),
        "email_configured": config.EMAIL_CONFIGURED,
        "smtp_configured": config.SMTP_CONFIGURED,
    }


# --------------------------------------------------------------------------
# Jobs
# --------------------------------------------------------------------------
class CreateJobBody(BaseModel):
    job_id: str
    display_name: Optional[str] = ""
    customer: Optional[str] = ""


class ClassifyBody(BaseModel):
    mode: str = "rules"


@app.get("/api/jobs")
def api_list_jobs():
    return jobs.list_jobs()


@app.post("/api/jobs")
def api_create_job(body: CreateJobBody):
    return jobs.create_job(body.job_id, body.display_name, body.customer)


@app.get("/api/jobs/{job_id}")
def api_get_job(job_id: str):
    job = jobs.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"Job '{job_id}' not found")
    return job


@app.post("/api/jobs/{job_id}/classify")
def api_classify_job(job_id: str, body: ClassifyBody):
    job = jobs.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"Job '{job_id}' not found")
    if job.get("base_type") == "bms":
        raise HTTPException(
            status_code=409,
            detail="This is a BMS / pot-block base. Its quote is BOM-driven by "
            "Module6121 and the AI classifier is intentionally disabled for it.",
        )
    try:
        return jobs.classify_job(job_id, body.mode)
    except FileNotFoundError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/jobs/{job_id}/upload")
async def api_upload(job_id: str, subfolder: str, file: UploadFile = File(...)):
    allowed = {"raw", "images", "models", "documents"}
    if subfolder not in allowed:
        raise HTTPException(status_code=400, detail=f"subfolder must be one of {allowed}")
    data = await file.read()
    target_name = "XT_Export_CAD_Dimensions.csv" if subfolder == "raw" else file.filename
    return jobs.save_upload(job_id, "" if subfolder == "raw" else subfolder, target_name, data)


@app.get("/api/jobs/{job_id}/file/{subfolder}/{filename}")
def api_get_file(job_id: str, subfolder: str, filename: str):
    path = jobs.get_asset_path(job_id, subfolder, filename)
    if path is None:
        raise HTTPException(status_code=404, detail="File not found")
    media_type, _ = mimetypes.guess_type(str(path))
    return FileResponse(path, media_type=media_type or "application/octet-stream")


@app.get("/api/jobs/{job_id}/quote-sheet")
def api_quote_sheet(job_id: str):
    job = jobs.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"Job '{job_id}' not found")
    sheet = pricing.build_quote_sheet(job)
    vba_bridge.write_bridge_files(job_id, sheet, job.get("job_analysis", {}))
    return sheet


# --------------------------------------------------------------------------
# Pricing config
# --------------------------------------------------------------------------
@app.get("/api/pricing")
def api_get_pricing():
    return pricing.load_rates()


@app.put("/api/pricing")
def api_put_pricing(rates: dict):
    return pricing.save_rates(rates)


# --------------------------------------------------------------------------
# Module6121 / VBA bridge
# --------------------------------------------------------------------------
class VbaClassifyBody(BaseModel):
    job_id: str
    csv_path: str
    base_type: str = "standard"  # "standard" | "bms"


@app.post("/api/vba/classify", response_class=PlainTextResponse)
def api_vba_classify(body: VbaClassifyBody):
    """Called by Module6121 right after it writes XT_Export_CAD_Dimensions.csv.

    Registers/updates the job (so it appears in the dashboard immediately),
    and for STANDARD bases runs the AI classifier and returns the bridge CSV
    (Index,Component,Role,ResolvedName,Confidence,Quote,Price,SecondaryPartingLine)
    the macro parses to name plates.

    BMS / pot-block bases are NEVER classified: the macro's BOM-driven flow
    owns them. They are only registered (marked base_type=bms) for visibility.
    """
    job_id = body.job_id.strip() or "ACTIVE"
    jobs.create_job(job_id, display_name=job_id)
    imported = jobs.import_raw_csv(job_id, body.csv_path)

    base_type = "bms" if body.base_type.strip().lower() in ("bms", "pot", "pot_block") else "standard"
    jobs.update_meta(job_id, base_type=base_type)

    if base_type == "bms":
        return "BMS_REGISTERED: BOM-driven flow retained; AI classification skipped."

    if not imported:
        raise HTTPException(
            status_code=400,
            detail=f"CSV not readable at '{body.csv_path}'. The macro and this app "
            "must run on the same machine (or share the path).",
        )

    try:
        job = jobs.classify_job(job_id, mode="rules")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Classification failed: {e}")

    sheet = pricing.build_quote_sheet(job)
    vba_bridge.write_bridge_files(job_id, sheet, job.get("job_analysis", {}))
    csv_file = config.VBA_BRIDGE_DIR / f"{job_id}_part_names.csv"
    return csv_file.read_text(encoding="utf-8")


@app.get("/api/bridge/{job_id}")
def api_bridge_json(job_id: str):
    payload = vba_bridge.read_bridge_json(job_id)
    if payload is None:
        # Not generated yet -- build it from the current quote sheet.
        job = jobs.get_job(job_id)
        if job is None:
            raise HTTPException(status_code=404, detail=f"Job '{job_id}' not found")
        sheet = pricing.build_quote_sheet(job)
        vba_bridge.write_bridge_files(job_id, sheet, job.get("job_analysis", {}))
        payload = vba_bridge.read_bridge_json(job_id)
    return payload


@app.get("/api/bridge/{job_id}/csv")
def api_bridge_csv(job_id: str):
    path = config.VBA_BRIDGE_DIR / f"{job_id}_part_names.csv"
    if not path.exists():
        api_bridge_json(job_id)  # generate it
    if not path.exists():
        raise HTTPException(status_code=404, detail="Bridge export not available")
    return FileResponse(path, media_type="text/csv", filename=path.name)


# --------------------------------------------------------------------------
# Email
# --------------------------------------------------------------------------
class ReplyBody(BaseModel):
    to: str
    subject: str
    body: str
    in_reply_to: Optional[str] = ""


class ReplyAllBody(BaseModel):
    to_addrs: list[str]
    cc_addrs: list[str] = []
    subject: str
    body: str
    in_reply_to: Optional[str] = ""


class ComposeBody(BaseModel):
    to_addrs: list[str]
    cc_addrs: list[str] = []
    subject: str
    body: str


class ForwardBody(BaseModel):
    to: str
    body: str = ""


class MarkReadBody(BaseModel):
    read: bool = True


class StarBody(BaseModel):
    starred: bool = True


class EmailSettingsBody(BaseModel):
    imap_host: Optional[str] = "imap.gmail.com"
    imap_port: Optional[int] = 993
    imap_user: Optional[str] = ""
    imap_password: Optional[str] = ""  # blank = keep existing
    imap_folder: Optional[str] = "INBOX"
    imap_ssl: Optional[bool] = True
    smtp_host: Optional[str] = "smtp.gmail.com"
    smtp_port: Optional[int] = 587
    smtp_user: Optional[str] = ""
    smtp_password: Optional[str] = ""  # blank = keep existing
    smtp_from: Optional[str] = ""
    gmail_address: Optional[str] = "cms1engineering@gmail.com"


class QuoteEmailBody(BaseModel):
    launch_macro: bool = True


@app.get("/api/settings/email")
def api_get_email_settings():
    return credentials.public_view()


@app.put("/api/settings/email")
def api_put_email_settings(body: EmailSettingsBody):
    saved = credentials.save(body.model_dump(exclude_none=True))
    config.reload_email_settings()
    return saved


@app.post("/api/settings/email/test")
def api_test_email():
    try:
        email_service.list_messages(limit=1)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"IMAP test failed: {e}")
    return {"ok": True, "message": "Inbox connection successful."}


@app.get("/api/workspace/browse")
def api_browse_workspace(path: str = ""):
    return jobs.browse_workspace(path)


class ImportFolderBody(BaseModel):
    folder_path: str
    run_quote: bool = True


@app.post("/api/jobs/import-folder")
def api_import_folder(body: ImportFolderBody):
    try:
        return jobs.import_from_folder(body.folder_path, run_quote=body.run_quote)
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


class JobCompleteBody(BaseModel):
    job_id: str
    folder_path: str
    base_type: str = "standard"
    status: str = "completed"


@app.post("/api/vba/job-complete")
def api_vba_job_complete(body: JobCompleteBody):
    """Called by Module6121 when ProcessOneJob finishes — syncs outputs to webapp."""
    try:
        job = quote_pipeline.sync_completed_job(body.job_id, body.folder_path, body.base_type)
        return {"synced": True, "job_id": job.get("job_id")}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/quote/status/{quote_id}")
def api_quote_status(quote_id: str):
    status = quote_pipeline.poll_completion(quote_id)
    if not status:
        raise HTTPException(status_code=404, detail="Quote run not found")
    return status


@app.get("/api/quote/active")
def api_quote_active():
    return quote_pipeline.list_active_quotes()


@app.get("/api/email/status")
def api_email_status():
    view = credentials.public_view()
    return {
        "configured": view["configured"],
        "smtp_configured": view["smtp_configured"],
        "imap_host": view["imap_host"] or None,
        "imap_user": view["imap_user"] or None,
        "credentials_path": view["credentials_path"],
    }


@app.get("/api/email/messages")
def api_email_messages(limit: int = 50, q: str = ""):
    try:
        messages = email_service.list_messages(limit, query=q)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not fetch inbox: {e}")

    job_ids = {j["job_id"].upper() for j in jobs.list_jobs()}
    for m in messages:
        m["matched_jobs"] = [t for t in m["job_tokens"] if t in job_ids]
    return messages


@app.get("/api/email/messages/{message_id}")
def api_email_message(message_id: str):
    try:
        message = email_service.get_message(message_id)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not fetch message: {e}")
    if message is None:
        raise HTTPException(status_code=404, detail="Message not found")

    job_ids = {j["job_id"].upper() for j in jobs.list_jobs()}
    message["matched_jobs"] = [t for t in message["job_tokens"] if t in job_ids]
    return message


@app.post("/api/email/messages/{message_id}/reply")
def api_email_reply(message_id: str, body: ReplyBody):
    try:
        email_service.send_reply(body.to, body.subject, body.body, body.in_reply_to)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not send reply: {e}")
    return {"sent": True}


@app.post("/api/email/messages/{message_id}/reply-all")
def api_email_reply_all(message_id: str, body: ReplyAllBody):
    try:
        email_service.send_reply_all(
            body.to_addrs, body.cc_addrs, body.subject, body.body, body.in_reply_to
        )
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not send reply: {e}")
    return {"sent": True}


@app.post("/api/email/messages/{message_id}/forward")
def api_email_forward(message_id: str, body: ForwardBody):
    try:
        email_service.forward_message(message_id, body.to, body.body)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not forward: {e}")
    return {"sent": True}


@app.post("/api/email/compose")
def api_email_compose(body: ComposeBody):
    try:
        email_service.send_compose(body.to_addrs, body.subject, body.body, body.cc_addrs)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not send email: {e}")
    return {"sent": True}


@app.patch("/api/email/messages/{message_id}/read")
def api_email_mark_read(message_id: str, body: MarkReadBody):
    try:
        email_service.mark_read(message_id, body.read)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not update message: {e}")
    return {"ok": True, "read": body.read}


@app.patch("/api/email/messages/{message_id}/star")
def api_email_star(message_id: str, body: StarBody):
    try:
        email_service.toggle_star(message_id, body.starred)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not update message: {e}")
    return {"ok": True, "starred": body.starred}


@app.delete("/api/email/messages/{message_id}")
def api_email_delete(message_id: str, permanent: bool = False):
    try:
        email_service.delete_message(message_id, permanent=permanent)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not delete message: {e}")
    return {"deleted": True}


@app.post("/api/email/messages/{message_id}/archive")
def api_email_archive(message_id: str):
    try:
        email_service.archive_message(message_id)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not archive message: {e}")
    return {"archived": True}


@app.post("/api/email/messages/{message_id}/quote")
def api_quote_email(message_id: str, body: QuoteEmailBody = QuoteEmailBody()):
    """One-click Quote: pull attachments, write cms_email.txt, start launcher."""
    try:
        return email_service.quote_from_message(message_id, launch_macro=body.launch_macro)
    except email_service.EmailNotConfigured as e:
        raise HTTPException(status_code=409, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not start quote: {e}")


@app.get("/api/training/status")
def api_training_status():
    _ensure_classifier_path()
    from geometry_classifier import training_audit  # type: ignore

    return training_audit.status()


@app.get("/api/training/suggestions")
def api_training_suggestions():
    _ensure_classifier_path()
    from geometry_classifier import training_audit  # type: ignore

    report = training_audit.status()
    return {
        "markdown": training_audit.suggestions_markdown(),
        "suggestions": report.get("suggestions", []),
        "overall_rules_accuracy_pct": report.get("overall_rules_accuracy_pct", 0),
        "jobs_processed": report.get("jobs_processed", 0),
        "bms_jobs": report.get("bms_jobs", 0),
        "standard_jobs": report.get("standard_jobs", 0),
    }


class TrainingRunBody(BaseModel):
    manifest_path: Optional[str] = None
    jobs_root: Optional[str] = None
    scan: bool = True
    use_qwen: bool = True
    qwen_model: str = "qwen3.5:9b"


def _ensure_classifier_path():
    import sys

    classifier_dir = config.GEOMETRY_CLASSIFIER_DIR
    parent = str(classifier_dir.parent)
    if parent not in sys.path:
        sys.path.insert(0, parent)


@app.post("/api/training/run")
def api_training_run(body: TrainingRunBody):
    """Scan TRAINING folder (BMS + standard), build CORRECT_ME, audit rules, suggest fixes."""
    _ensure_classifier_path()
    try:
        from geometry_classifier import training_audit  # type: ignore

        default_root = os.environ.get(
            "CMS_TRAINING_ROOT",
            r"C:\Users\lenovo\Downloads\TRAINING",
        )
        jobs_root = body.jobs_root or default_root
        result = training_audit.run_full_audit(
            jobs_root=jobs_root if body.scan else None,
            manifest_path=body.manifest_path,
            use_qwen=body.use_qwen,
            qwen_model=body.qwen_model,
        )
        return result
    except RuntimeError as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# --------------------------------------------------------------------------
# Built frontend (single local process: 127.0.0.1:8000 serves UI + API).
# --------------------------------------------------------------------------
_FRONTEND_DIST = Path(__file__).resolve().parent.parent.parent / "frontend" / "dist"
if _FRONTEND_DIST.exists():
    app.mount("/assets", StaticFiles(directory=_FRONTEND_DIST / "assets"), name="assets")

    @app.get("/{full_path:path}", include_in_schema=False)
    def spa(full_path: str):
        candidate = _FRONTEND_DIST / full_path
        if full_path and candidate.is_file():
            return FileResponse(candidate)
        return FileResponse(_FRONTEND_DIST / "index.html")
