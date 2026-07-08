import mimetypes
from typing import Optional

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from . import config, email_service, jobs, pricing, vba_bridge

app = FastAPI(title="CMS AI Quoting")

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


@app.get("/api/email/status")
def api_email_status():
    return {
        "configured": config.EMAIL_CONFIGURED,
        "smtp_configured": config.SMTP_CONFIGURED,
        "imap_host": config.IMAP_HOST or None,
        "imap_user": config.IMAP_USER or None,
    }


@app.get("/api/email/messages")
def api_email_messages(limit: int = 30):
    try:
        messages = email_service.list_messages(limit)
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
