"""IMAP/SMTP email integration.

Credentials are stored in the webapp Settings page (email_credentials.json
under CMS_DATA_DIR). Environment variables CMS_IMAP_* / CMS_SMTP_* override
the file when set.
"""
import email
import imaplib
import os
import re
import shutil
import subprocess
from pathlib import Path
import smtplib
from email.header import decode_header
from email.message import EmailMessage
from email.utils import parsedate_to_datetime

from . import config, jobs

JOB_TOKEN_RE = re.compile(r"\b([A-Z]{1,2}\d{4,6})\b")
MIN_JOB_DIGITS = 7
LOCAL_WORKSPACE = Path(os.environ.get("CMS_LOCAL_WORKSPACE", r"C:\CMS_Local_Workspace"))
EMAIL_OUTPUT_FILE = LOCAL_WORKSPACE / "cms_email.txt"
DOWNLOADS_FOLDER = Path(os.environ.get("CMS_DOWNLOADS_FOLDER", r"C:\Users\lenovo\Downloads"))


class EmailNotConfigured(Exception):
    pass


def _decode(value) -> str:
    if not value:
        return ""
    parts = decode_header(value)
    out = []
    for text, enc in parts:
        if isinstance(text, bytes):
            out.append(text.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def _connect():
    if not config.EMAIL_CONFIGURED:
        raise EmailNotConfigured(
            "Email is not configured. Set CMS_IMAP_HOST / CMS_IMAP_USER / "
            "CMS_IMAP_PASSWORD (an app password works for Gmail/Outlook) as "
            "secrets, then reload."
        )
    imap = imaplib.IMAP4_SSL(config.IMAP_HOST, config.IMAP_PORT) if config.IMAP_USE_SSL \
        else imaplib.IMAP4(config.IMAP_HOST, config.IMAP_PORT)
    imap.login(config.IMAP_USER, config.IMAP_PASSWORD)
    imap.select(config.IMAP_FOLDER)
    return imap


def guess_job_tokens(*texts) -> list:
    tokens = set()
    for text in texts:
        if not text:
            continue
        tokens.update(JOB_TOKEN_RE.findall(text.upper()))
    return sorted(tokens)


def list_messages(limit: int = 30) -> list:
    imap = _connect()
    try:
        status, data = imap.search(None, "ALL")
        if status != "OK":
            return []
        ids = data[0].split()
        ids = ids[-limit:][::-1]
        messages = []
        for msg_id in ids:
            status, msg_data = imap.fetch(
                msg_id, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])"
            )
            if status != "OK" or not msg_data or not msg_data[0]:
                continue
            raw_headers = msg_data[0][1]
            msg = email.message_from_bytes(raw_headers)
            subject = _decode(msg.get("Subject"))
            from_ = _decode(msg.get("From"))
            date_hdr = msg.get("Date")
            try:
                date_iso = parsedate_to_datetime(date_hdr).isoformat() if date_hdr else ""
            except Exception:
                date_iso = date_hdr or ""
            messages.append(
                {
                    "id": msg_id.decode(),
                    "from": from_,
                    "subject": subject,
                    "date": date_iso,
                    "job_tokens": guess_job_tokens(subject),
                }
            )
        return messages
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def get_message(message_id: str) -> dict:
    imap = _connect()
    try:
        status, msg_data = imap.fetch(message_id.encode(), "(RFC822)")
        if status != "OK" or not msg_data or not msg_data[0]:
            return None
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        subject = _decode(msg.get("Subject"))
        from_ = _decode(msg.get("From"))
        to = _decode(msg.get("To"))
        date_hdr = msg.get("Date")
        message_id_hdr = msg.get("Message-ID", "")

        body_text, body_html, attachments = "", "", []
        if msg.is_multipart():
            for part in msg.walk():
                disp = str(part.get("Content-Disposition") or "")
                ctype = part.get_content_type()
                if "attachment" in disp or part.get_filename():
                    attachments.append(
                        {
                            "filename": _decode(part.get_filename()),
                            "content_type": ctype,
                            "size": len(part.get_payload(decode=True) or b""),
                        }
                    )
                elif ctype == "text/plain" and not body_text:
                    body_text = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    )
                elif ctype == "text/html" and not body_html:
                    body_html = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    )
        else:
            payload = msg.get_payload(decode=True) or b""
            text = payload.decode(msg.get_content_charset() or "utf-8", errors="replace")
            if msg.get_content_type() == "text/html":
                body_html = text
            else:
                body_text = text

        job_tokens = guess_job_tokens(
            subject, " ".join(a["filename"] for a in attachments)
        )

        return {
            "id": message_id,
            "from": from_,
            "to": to,
            "subject": subject,
            "date": date_hdr,
            "message_id_header": message_id_hdr,
            "body_text": body_text,
            "body_html": body_html,
            "attachments": attachments,
            "job_tokens": job_tokens,
        }
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def send_reply(to_addr: str, subject: str, body: str, in_reply_to: str = "") -> None:
    if not config.SMTP_CONFIGURED:
        raise EmailNotConfigured(
            "SMTP is not configured. Open Settings in the webapp and save your "
            "Gmail app password there."
        )
    msg = EmailMessage()
    msg["From"] = config.SMTP_FROM
    msg["To"] = to_addr
    msg["Subject"] = subject if subject.lower().startswith("re:") else f"Re: {subject}"
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        msg["References"] = in_reply_to
    msg.set_content(body)

    with smtplib.SMTP(config.SMTP_HOST, config.SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(config.SMTP_USER, config.SMTP_PASSWORD)
        smtp.send_message(msg)


def _clean_job_token(token: str) -> str:
    token = (token or "").strip()
    if not token:
        return ""
    upper = token.upper()
    bad_prefixes = ("STEP-", "STP-", "SLDPRT-", "SLDASM-", "X-T-", "XT-", "PARASOLID-")
    bad_exact = {"STEP", "STP", "SLDPRT", "SLDASM", "X-T", "XT", "PARASOLID"}
    if upper in bad_exact or any(upper.startswith(p) for p in bad_prefixes):
        return ""
    return token[:60]


def _first_job_token(text: str) -> str:
    patterns = (
        r"\b[A-Za-z][A-Za-z0-9]*(?:[-_][A-Za-z0-9]+)+\b",
        r"\b[A-Z]{1,4}\d{3,}\b",
    )
    for pat in patterns:
        for m in re.finditer(pat, text or ""):
            token = _clean_job_token(m.group(0))
            if token and re.search(r"\d", token):
                return token
    return ""


def _extract_number_after(text: str, label: str) -> str:
    m = re.search(rf"{re.escape(label)}\s*[:#]?\s*([A-Za-z0-9][A-Za-z0-9\-_/]*)", text, re.I)
    return _clean_job_token(m.group(1)) if m else ""


def _first_long_number(text: str, min_digits: int = MIN_JOB_DIGITS) -> str:
    for m in re.finditer(r"\d+", text or ""):
        if len(m.group(0)) >= min_digits:
            return m.group(0)
    return ""


def extract_quote_info(msg: email.message.Message) -> dict:
    """Pull customer job #, ship date, similar-to from a quote email."""
    subject = _decode(msg.get("Subject"))
    body_text = ""
    attachment_names: list[str] = []

    if msg.is_multipart():
        for part in msg.walk():
            disp = str(part.get("Content-Disposition") or "")
            fn = part.get_filename()
            if fn:
                attachment_names.append(_decode(fn))
            elif part.get_content_type() == "text/plain" and not body_text and "attachment" not in disp:
                body_text = part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="replace"
                )
    else:
        body_text = (msg.get_payload(decode=True) or b"").decode(
            msg.get_content_charset() or "utf-8", errors="replace"
        )

    names_blob = " ".join(attachment_names)
    text = f"{subject}\n{body_text}"
    job = (
        _extract_number_after(text, "JOB#")
        or _extract_number_after(text, "JOB #")
        or _extract_number_after(text, "JOB NUMBER")
    )
    if not job:
        job = _first_long_number(names_blob)
    if not job:
        job = _first_long_number(subject)
    if not job:
        job = _first_long_number(text + "\n" + names_blob)
    if not job:
        job = _first_job_token(names_blob)
    if not job:
        job = _first_job_token(subject)
    if not job:
        job = _first_job_token(text + "\n" + names_blob)
    job = _clean_job_token(job)

    ship_m = re.search(r"SHIP\s*DATE\s*[:#]?\s*([^\n\r]+)", text, re.I)
    similar_m = re.search(r"SIMILAR\s*TO\s*[:#]?\s*([^\n\r]+)", text, re.I)

    return {
        "subject": subject,
        "cust_job": job,
        "similar_to": (similar_m.group(1).strip() if similar_m else ""),
        "ship_date": (ship_m.group(1).strip() if ship_m else ""),
        "attachment_names": attachment_names,
    }


def _save_attachments(msg: email.message.Message, job_token: str) -> tuple[int, Path]:
    job_token = _clean_job_token(job_token) or "unknown"
    base = DOWNLOADS_FOLDER / "CMS_Incoming" / job_token
    incoming_root = (DOWNLOADS_FOLDER / "CMS_Incoming").resolve()
    base_resolved = base.resolve()
    if str(base_resolved).startswith(str(incoming_root)) and base_resolved.exists():
        shutil.rmtree(base_resolved)
    base_resolved.mkdir(parents=True, exist_ok=True)
    count = 0
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        fn = part.get_filename()
        if not fn:
            continue
        safe = re.sub(r'[\\/:*?"<>|]', "_", _decode(fn)).strip()
        if not safe:
            continue
        data = part.get_payload(decode=True)
        if not data:
            continue
        (base_resolved / safe).write_bytes(data)
        count += 1
    return count, base_resolved


def _write_email_handoff(info: dict, attach_dir: Path, attach_count: int) -> None:
    LOCAL_WORKSPACE.mkdir(parents=True, exist_ok=True)
    lines = {
        "Found": "1",
        "Subject": info.get("subject", ""),
        "CustJob": info.get("cust_job", ""),
        "SimilarTo": info.get("similar_to", ""),
        "ShipDate": info.get("ship_date", ""),
        "Attachments": str(attach_count),
        "AttachDir": str(attach_dir),
        "Error": "",
    }
    EMAIL_OUTPUT_FILE.write_text(
        "\n".join(f"{k}={v}" for k, v in lines.items()) + "\n",
        encoding="utf-8",
    )


def _launch_quote_flow() -> bool:
    """Start CMS_Launcher.vbs /usemail on Windows when available."""
    candidates = [
        LOCAL_WORKSPACE / "CMS_Launcher.vbs",
        Path(__file__).resolve().parent.parent.parent.parent / "CMS_Launcher.vbs",
    ]
    for vbs in candidates:
        if vbs.exists():
            try:
                subprocess.Popen(["wscript", str(vbs), "/usemail"], close_fds=True)
                return True
            except Exception:
                pass
    return False


def quote_from_message(message_id: str, launch_macro: bool = True) -> dict:
    """Download attachments, write cms_email.txt, optionally launch SolidWorks flow."""
    imap = _connect()
    try:
        status, msg_data = imap.fetch(message_id.encode(), "(RFC822)")
        if status != "OK" or not msg_data or not msg_data[0]:
            raise ValueError("Message not found")
        msg = email.message_from_bytes(msg_data[0][1])
    finally:
        try:
            imap.logout()
        except Exception:
            pass

    info = extract_quote_info(msg)
    tokens = guess_job_tokens(info["subject"])
    job_token = info["cust_job"] or (tokens[0] if tokens else "")
    if not job_token:
        job_token = f"EMAIL-{message_id}"

    attach_count, attach_dir = _save_attachments(msg, job_token)
    _write_email_handoff(info, attach_dir, attach_count)

    jobs.create_job(job_token, display_name=info["subject"][:80], customer=info["cust_job"])
    job_dir = config.JOBS_ROOT / job_token.replace("..", "").replace("/", "_")
    docs = job_dir / "documents"
    docs.mkdir(parents=True, exist_ok=True)
    for src in attach_dir.glob("*"):
        if src.is_file():
            dest = docs / src.name
            if not dest.exists():
                shutil.copy2(src, dest)

    launched = False
    quote_id = info["cust_job"] or job_token
    if launch_macro:
        from . import quote_pipeline

        quote_pipeline.set_status(
            quote_id,
            phase="queued",
            message="Preparing quote — downloading attachments done.",
            cust_job=info["cust_job"],
        )
        result = quote_pipeline.launch_full_quote(
            quote_id,
            str(attach_dir),
            {
                "subject": info["subject"],
                "cust_job": info["cust_job"],
                "similar_to": info["similar_to"],
                "ship_date": info["ship_date"],
                "attachments": attach_count,
            },
        )
        launched = result.get("launched", False)
        job_token = result.get("job_id") or job_token

    return {
        "job_id": job_token,
        "quote_id": quote_id,
        "subject": info["subject"],
        "cust_job": info["cust_job"],
        "attachments_saved": attach_count,
        "attach_dir": str(attach_dir),
        "launcher_started": launched,
        "email_handoff": str(EMAIL_OUTPUT_FILE),
        "poll_url": f"/api/quote/status/{quote_id}",
    }
