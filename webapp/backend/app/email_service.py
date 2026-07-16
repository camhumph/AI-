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


def _parse_flags(fetch_line: bytes | tuple) -> dict:
    seen = flagged = False
    raw = b""
    if isinstance(fetch_line, tuple):
        raw = fetch_line[0] if isinstance(fetch_line[0], bytes) else b""
    elif isinstance(fetch_line, bytes):
        raw = fetch_line
    text = raw.decode("utf-8", errors="replace").upper()
    seen = "\\SEEN" in text
    flagged = "\\FLAGGED" in text
    return {"seen": seen, "starred": flagged}


def _snippet_from_msg(msg: email.message.Message, max_len: int = 120) -> str:
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get_filename():
                continue
            if part.get_content_type() == "text/plain":
                body = (part.get_payload(decode=True) or b"").decode(
                    part.get_content_charset() or "utf-8", errors="replace"
                )
                break
    else:
        body = (msg.get_payload(decode=True) or b"").decode(
            msg.get_content_charset() or "utf-8", errors="replace"
        )
    body = re.sub(r"\s+", " ", body).strip()
    return body[:max_len] + ("…" if len(body) > max_len else "")


def _extract_name_addr(from_hdr: str) -> tuple[str, str]:
    m = re.match(r"^(.*?)\s*<([^>]+)>$", (from_hdr or "").strip())
    if m:
        return m.group(1).strip('" '), m.group(2).strip()
    return from_hdr, from_hdr


def list_messages(limit: int = 40, query: str = "", folder: str | None = None) -> list:
    """Fast inbox list — headers + flags only (no body peek per message)."""
    imap = _connect()
    try:
        if folder and folder != config.IMAP_FOLDER:
            imap.select(folder)
        criteria = "ALL"
        q = (query or "").strip()
        if q:
            safe = q.replace('"', "")
            # Subject/From only — BODY search is very slow on large mailboxes
            criteria = f'(OR SUBJECT "{safe}" FROM "{safe}")'
        status, data = imap.search(None, criteria)
        if status != "OK":
            return []
        ids = data[0].split()
        ids = ids[-limit:][::-1]
        if not ids:
            return []

        # Batch fetch headers for all IDs in one round-trip
        id_list = b",".join(ids)
        status, msg_data = imap.fetch(
            id_list,
            "(FLAGS BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])",
        )
        if status != "OK" or not msg_data:
            return []

        # Build id -> (flags, header_bytes) from fetch response
        by_id: dict[bytes, tuple[dict, bytes]] = {}
        current_id: bytes | None = None
        for item in msg_data:
            if isinstance(item, bytes):
                # e.g. b')' separators — ignore
                continue
            if not isinstance(item, tuple) or len(item) < 2:
                continue
            meta = item[0] or b""
            meta_s = meta.decode("utf-8", errors="replace") if isinstance(meta, bytes) else str(meta)
            # Extract message sequence number from fetch meta like b'123 (FLAGS ...'
            m = re.match(r"(\d+)\s*\(", meta_s)
            if m:
                current_id = m.group(1).encode()
            flags = _parse_flags(item) if "FLAGS" in meta_s.upper() else {"seen": False, "starred": False}
            header_bytes = item[1] or b""
            if current_id is not None:
                prev = by_id.get(current_id)
                if prev:
                    # merge flags if we already have header
                    merged_flags = prev[0] if prev[0].get("seen") or prev[0].get("starred") else flags
                    by_id[current_id] = (merged_flags if "FLAGS" in meta_s.upper() else prev[0], header_bytes or prev[1])
                else:
                    by_id[current_id] = (flags, header_bytes)

        messages = []
        for msg_id in ids:
            flags, header_bytes = by_id.get(msg_id, ({"seen": False, "starred": False}, b""))
            if not header_bytes:
                continue
            msg = email.message_from_bytes(header_bytes)
            subject = _decode(msg.get("Subject"))
            from_ = _decode(msg.get("From"))
            name, addr = _extract_name_addr(from_)
            date_hdr = msg.get("Date")
            try:
                date_iso = parsedate_to_datetime(date_hdr).isoformat() if date_hdr else ""
            except Exception:
                date_iso = date_hdr or ""
            messages.append(
                {
                    "id": msg_id.decode(),
                    "from": from_,
                    "from_name": name or addr,
                    "from_addr": addr,
                    "subject": subject,
                    "date": date_iso,
                    "snippet": "",
                    "seen": flags.get("seen", False),
                    "starred": flags.get("starred", False),
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

        flags = {"seen": False, "starred": False}
        try:
            st, fd = imap.fetch(message_id.encode(), "(FLAGS)")
            if st == "OK" and fd:
                flags = _parse_flags(fd[0])
        except Exception:
            pass

        return {
            "id": message_id,
            "from": from_,
            "to": to,
            "cc": _decode(msg.get("Cc")),
            "subject": subject,
            "date": date_hdr,
            "message_id_header": message_id_hdr,
            "body_text": body_text,
            "body_html": body_html,
            "attachments": attachments,
            "job_tokens": job_tokens,
            "seen": flags["seen"],
            "starred": flags["starred"],
        }
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def _msg_id_bytes(message_id: str) -> bytes:
    return message_id.encode() if isinstance(message_id, str) else message_id


def mark_read(message_id: str, read: bool = True) -> None:
    imap = _connect()
    try:
        flag = "\\Seen" if read else "-FLAGS"
        if read:
            imap.store(_msg_id_bytes(message_id), "+FLAGS", "\\Seen")
        else:
            imap.store(_msg_id_bytes(message_id), "-FLAGS", "\\Seen")
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def toggle_star(message_id: str, starred: bool) -> None:
    imap = _connect()
    try:
        if starred:
            imap.store(_msg_id_bytes(message_id), "+FLAGS", "\\Flagged")
        else:
            imap.store(_msg_id_bytes(message_id), "-FLAGS", "\\Flagged")
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def delete_message(message_id: str, permanent: bool = False) -> None:
    imap = _connect()
    try:
        mid = _msg_id_bytes(message_id)
        if not permanent:
            for trash in ("[Gmail]/Trash", "Trash", "[Google Mail]/Trash"):
                try:
                    imap.copy(mid, trash)
                    break
                except Exception:
                    continue
        imap.store(mid, "+FLAGS", "\\Deleted")
        imap.expunge()
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def archive_message(message_id: str) -> None:
    imap = _connect()
    try:
        mid = _msg_id_bytes(message_id)
        try:
            imap.store(mid, "-X-GM-LABELS", r"(\Inbox)")
        except Exception:
            for archive in ("[Gmail]/All Mail", "All Mail"):
                try:
                    imap.copy(mid, archive)
                    break
                except Exception:
                    continue
            imap.store(mid, "+FLAGS", "\\Deleted")
            imap.expunge()
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def _smtp_send(msg: EmailMessage) -> None:
    if not config.SMTP_CONFIGURED:
        raise EmailNotConfigured(
            "SMTP is not configured. Open Settings in the webapp and save your "
            "Gmail app password there."
        )
    with smtplib.SMTP(config.SMTP_HOST, config.SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(config.SMTP_USER, config.SMTP_PASSWORD)
        smtp.send_message(msg)


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
    _smtp_send(msg)


def send_reply_all(
    to_addrs: list[str],
    cc_addrs: list[str],
    subject: str,
    body: str,
    in_reply_to: str = "",
) -> None:
    msg = EmailMessage()
    msg["From"] = config.SMTP_FROM
    msg["To"] = ", ".join(to_addrs)
    if cc_addrs:
        msg["Cc"] = ", ".join(cc_addrs)
    msg["Subject"] = subject if subject.lower().startswith("re:") else f"Re: {subject}"
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        msg["References"] = in_reply_to
    msg.set_content(body)
    _smtp_send(msg)


def send_compose(
    to_addrs: list[str],
    subject: str,
    body: str,
    cc_addrs: list[str] | None = None,
) -> None:
    msg = EmailMessage()
    msg["From"] = config.SMTP_FROM
    msg["To"] = ", ".join(to_addrs)
    if cc_addrs:
        msg["Cc"] = ", ".join(cc_addrs)
    msg["Subject"] = subject
    msg.set_content(body)
    _smtp_send(msg)


def send_forward(to_addr: str, subject: str, body: str, original: email.message.Message) -> None:
    fwd_subject = subject if subject.lower().startswith("fwd:") else f"Fwd: {subject}"
    orig_text, orig_html = "", ""
    if original.is_multipart():
        for part in original.walk():
            ctype = part.get_content_type()
            if ctype == "text/plain" and not orig_text:
                orig_text = part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="replace"
                )
            elif ctype == "text/html" and not orig_html:
                orig_html = part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="replace"
                )
    else:
        text = (original.get_payload(decode=True) or b"").decode(
            original.get_content_charset() or "utf-8", errors="replace"
        )
        if original.get_content_type() == "text/html":
            orig_html = text
        else:
            orig_text = text

    msg = EmailMessage()
    msg["From"] = config.SMTP_FROM
    msg["To"] = to_addr
    msg["Subject"] = fwd_subject
    if body.strip():
        combined = f"{body.strip()}\n\n---------- Forwarded message ----------\n{orig_text}"
    else:
        combined = f"---------- Forwarded message ----------\n{orig_text}"
    if orig_html:
        msg.set_content(combined)
        msg.add_alternative(orig_html, subtype="html")
    else:
        msg.set_content(combined)
    _smtp_send(msg)


def forward_message(message_id: str, to_addr: str, body: str = "") -> None:
    imap = _connect()
    try:
        status, msg_data = imap.fetch(_msg_id_bytes(message_id), "(RFC822)")
        if status != "OK" or not msg_data or not msg_data[0]:
            raise ValueError("Message not found")
        original = email.message_from_bytes(msg_data[0][1])
        subject = _decode(original.get("Subject"))
        send_forward(to_addr, subject, body, original)
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def extract_c_number(*texts: str) -> str:
    """Prefer an existing CMS C-number like C18603 from BMS-...-C18603 names."""
    for text in texts:
        if not text:
            continue
        # Prefer -C##### / _C##### (BMS-851100029-C18603)
        m = re.search(r"[-_]C(\d{4,6})\b", text, re.I)
        if m:
            return "C" + m.group(1)
        m = re.search(r"\bC[- ]?(\d{4,6})\b", text, re.I)
        if m:
            return "C" + m.group(1)
    return ""


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
    c_number = extract_c_number(names_blob, subject, text, job)

    ship_m = re.search(r"SHIP\s*DATE\s*[:#]?\s*([^\n\r]+)", text, re.I)
    similar_m = re.search(r"SIMILAR\s*TO\s*[:#]?\s*([^\n\r]+)", text, re.I)

    return {
        "subject": subject,
        "cust_job": job,
        "c_number": c_number,
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
    c_number = info.get("c_number") or extract_c_number(
        info.get("subject", ""),
        info.get("cust_job", ""),
        str(attach_dir),
        " ".join(info.get("attachment_names") or []),
    )
    lines = {
        "Found": "1",
        "Subject": info.get("subject", ""),
        "CustJob": info.get("cust_job", ""),
        "CNum": c_number,
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
    from . import quote_pipeline

    proc, how = quote_pipeline._start_cms_launcher()
    return proc is not None


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
    # Prefer existing C-number (BMS-...-C18603) for folder/job id when present.
    c_number = info.get("c_number") or extract_c_number(
        info["subject"], info["cust_job"], " ".join(info.get("attachment_names") or [])
    )
    job_token = c_number or info["cust_job"] or (tokens[0] if tokens else "")
    if not job_token:
        job_token = f"EMAIL-{message_id}"

    attach_count, attach_dir = _save_attachments(msg, job_token)
    # Re-scan attach path for C-number (folder names often carry BMS-...-C#####)
    if not c_number:
        c_number = extract_c_number(str(attach_dir), job_token)
        info["c_number"] = c_number
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
    quote_id = c_number or info["cust_job"] or job_token
    if launch_macro:
        from . import quote_pipeline

        quote_pipeline.set_status(
            quote_id,
            phase="queued",
            message="Preparing quote — downloading attachments done.",
            cust_job=info["cust_job"],
            c_number=c_number,
        )
        result = quote_pipeline.launch_full_quote(
            quote_id,
            str(attach_dir),
            {
                "subject": info["subject"],
                "cust_job": info["cust_job"],
                "c_number": c_number,
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
        "c_number": c_number,
        "attachments_saved": attach_count,
        "attach_dir": str(attach_dir),
        "launcher_started": launched,
        "email_handoff": str(EMAIL_OUTPUT_FILE),
        "poll_url": f"/api/quote/status/{quote_id}",
    }


def quote_from_messages(message_ids: list[str], launch_macro: bool = True) -> dict:
    """Prepare multiple email quotes, then launch them as one sequential SolidWorks batch."""
    if not message_ids:
        return {"launched": False, "error": "No message ids", "results": []}

    prepared: list[dict] = []
    results: list[dict] = []

    for mid in message_ids:
        mid = str(mid).strip()
        if not mid:
            continue
        # Prepare attachments/handoff fields without launching each one separately.
        one = quote_from_message(mid, launch_macro=False)
        results.append(one)
        prepared.append(
            {
                "quote_id": one.get("quote_id") or one.get("job_id") or mid,
                "attach_dir": one.get("attach_dir") or "",
                "c_number": one.get("c_number") or "",
                "email_info": {
                    "subject": one.get("subject") or "",
                    "cust_job": one.get("cust_job") or "",
                    "c_number": one.get("c_number") or "",
                    "similar_to": "",
                    "ship_date": "",
                },
            }
        )

    if not prepared:
        return {"launched": False, "error": "No quotes prepared", "results": results}

    if not launch_macro:
        return {
            "launched": False,
            "batch": True,
            "batch_count": len(prepared),
            "results": results,
            "quote_ids": [p["quote_id"] for p in prepared],
        }

    from . import quote_pipeline

    batch = quote_pipeline.launch_batch_quotes(prepared)
    return {
        **batch,
        "results": results,
    }
