"""IMAP/SMTP email integration.

Reads credentials only from environment variables (Cursor Cloud Agent
Secrets in this environment, or a local .env on a real CMS machine):

  CMS_IMAP_HOST, CMS_IMAP_PORT, CMS_IMAP_USER, CMS_IMAP_PASSWORD, CMS_IMAP_FOLDER
  CMS_SMTP_HOST, CMS_SMTP_PORT, CMS_SMTP_USER, CMS_SMTP_PASSWORD, CMS_SMTP_FROM

If IMAP is not configured, list_messages()/get_message() raise
EmailNotConfigured so the API can return a friendly "connect your email"
state instead of fake data.
"""
import email
import imaplib
import re
import smtplib
from email.header import decode_header
from email.message import EmailMessage
from email.utils import parsedate_to_datetime

from . import config

JOB_TOKEN_RE = re.compile(r"\b([A-Z]{1,2}\d{4,6})\b")


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
            "SMTP is not configured. Set CMS_SMTP_HOST / CMS_SMTP_USER / "
            "CMS_SMTP_PASSWORD as secrets, then reload."
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
