#!/usr/bin/env python3
# ============================================================
# cms_gmail_search.py
# ------------------------------------------------------------
# Two modes:
#   * Run by itself (double-click / "python cms_gmail_search.py")
#       -> hands off to CMS_Launcher.vbs, which runs the WHOLE
#          job (Gmail -> proposal -> SolidWorks).
#   * Run by the launcher with  --from-launcher
#       -> logs into Gmail, finds the NEWEST matching email,
#          pulls the customer Job # / ship date / similar-to,
#          and writes them to cms_email.txt for the launcher.
#
# Standard library only. Logs to the Downloads folder.
#
# ONE-TIME GMAIL SETUP:
#   1. Turn on 2-Step Verification on the Gmail account.
#   2. Google Account > Security > App passwords -> make one for
#      "Mail" and paste the 16 characters into APP_PASSWORD below.
# ============================================================

import imaplib
import email
import os
import re
import sys
import datetime
import subprocess
import shutil
from email.header import decode_header
from email.utils import parseaddr

# ===================== SETTINGS =====================
GMAIL_ADDRESS      = "cms1engineering@gmail.com"  # default; overridden by credentials file
# Credentials are saved in the webapp Settings page (never gmail_app_password.txt).
# Same JSON file Module6121.bas reads for SMTP.
def _credentials_paths():
    paths = []
    data_dir = os.environ.get("CMS_DATA_DIR", r"C:\CMS_Local_Workspace\cms_data")
    paths.append(os.path.join(data_dir, "email_credentials.json"))
    paths.append(r"C:\CMS_Local_Workspace\cms_data\email_credentials.json")
    here = os.path.dirname(os.path.abspath(__file__))
    paths.append(os.path.join(here, "webapp", "backend", "data", "email_credentials.json"))
    return paths


def _load_email_credentials():
  cred = {}
  for path in _credentials_paths():
    try:
      with open(path, "r", encoding="utf-8") as f:
        data = __import__("json").load(f)
      if isinstance(data, dict):
        cred = data
        break
    except Exception:
      pass
  global GMAIL_ADDRESS
  if cred.get("gmail_address"):
    GMAIL_ADDRESS = cred["gmail_address"]
  elif cred.get("imap_user"):
    GMAIL_ADDRESS = cred["imap_user"]
  return cred


def _read_app_password():
    cred = _load_email_credentials()
    pw = (cred.get("imap_password") or cred.get("smtp_password") or "").strip()
    if pw:
        return pw
    # Legacy fallback only if the old file still exists (migrate via Settings page).
    legacy = r"C:\CMS_Local_Workspace\gmail_app_password.txt"
    try:
        with open(legacy, "r", encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return ""


APP_PASSWORD       = _read_app_password()         # 16-char Gmail app password
SUBJECT_KEYWORD    = ""                           # "" = list ALL unopened (forwards/RFQs have no fixed word)
SEARCH_UNREAD_ONLY = True                         # prefer unread, fall back to read if none
MARK_AS_READ       = False                        # mark the picked email read after pulling it
DOWNLOADS_FOLDER   = r"C:\Users\lenovo\Downloads" # where attachments are saved
OUTPUT_FILE        = r"C:\CMS_Local_Workspace\cms_email.txt"
LOG_FILE           = r"C:\Users\lenovo\Downloads\CMS_Quote_Log.txt"
MIN_JOB_DIGITS     = 7                            # a customer job # is a run of >= this many digits
# ====================================================


def log(msg):
    try:
        folder = os.path.dirname(LOG_FILE)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write("[%s] gmail: %s\n" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass


def _decode(s):
    if s is None:
        return ""
    out = ""
    for text, enc in decode_header(s):
        if isinstance(text, bytes):
            try:
                out += text.decode(enc or "utf-8", "replace")
            except Exception:
                out += text.decode("utf-8", "replace")
        else:
            out += text
    return out


def get_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and "attachment" not in str(part.get("Content-Disposition")):
                try:
                    return part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "replace")
                except Exception:
                    pass
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                import html as _html
                try:
                    raw = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "replace")
                    return _html.unescape(re.sub(r"<[^>]+>", " ", raw))
                except Exception:
                    pass
        return ""
    try:
        return msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", "replace")
    except Exception:
        return str(msg.get_payload())


def extract_number_after(text, token):
    u = text.upper()
    i = u.find(token.upper())
    if i < 0:
        return ""
    i += len(token)
    res, started = "", False
    while i < len(u):
        ch = u[i]
        if ch.isdigit():
            res += ch
            started = True
        elif started:
            break
        i += 1
    return res


def extract_text_after(text, token):
    u = text.upper()
    i = u.find(token.upper())
    if i < 0:
        return ""
    i += len(token)
    rest = text[i:]
    line = rest.splitlines()[0] if rest.splitlines() else ""
    return line.strip().lstrip(":-").strip()


def clean_company_token(token):
    token = (token or "").strip()
    token = re.sub(r"(?i)(?:\.|-|_)?(?:step|stp|sldasm|sldprt|x_t|xt|igs|iges)(?:[-_]\d+)?(?:\..*)?$", "", token)
    token = re.sub(r"(?i)\.[A-Za-z0-9]{2,6}(?:[-_]\d+)?$", "", token)
    token = re.sub(r"(?i)\b(please|quote|rfq|bom|proposal|request|assembly)\b", " ", token)
    token = re.sub(r"[^A-Za-z0-9]+", "-", token).strip("-")
    token = re.sub(r"-+", "-", token)
    return token[:50]


def attachment_file_token(msg):
    for part in msg.walk():
        try:
            nm = part.get_filename()
        except Exception:
            nm = None
        if not nm:
            continue
        nm = _decode(nm)
        stem = re.sub(r"(?i)\.(step|stp|sldasm|sldprt|x_t|xt|igs|iges|pdf|xlsx?|csv).*$", "", nm)
        tok = clean_company_token(stem)
        if tok and tok.upper() not in {"STEP", "STP", "SLDASM", "SLDPRT", "XT", "X-T"}:
            return tok
    return ""


def sender_company_token(msg):
    name, addr = parseaddr(_decode(msg.get("From")))
    tok = clean_company_token(name)
    if tok and tok.lower() not in {"cms", "gmail", "quotes", "engineering"}:
        return tok
    dom = (addr.split("@", 1)[1] if "@" in addr else "").split(".")[0]
    return clean_company_token(dom)


def customer_identity_from_message(msg, job):
    subject = _decode(msg.get("Subject"))
    names = get_attachment_names(msg)
    from_text = _decode(msg.get("From"))
    all_text = " ".join([subject, names, from_text])
    if re.search(r"(?i)\bBMS\b", all_text):
        return "BMS", "BMS"
    prefix = sender_company_token(msg) or attachment_file_token(msg) or clean_company_token(job)
    if not prefix:
        prefix = "Quote"
    return prefix, prefix

def first_long_number(text, minlen):
    m = re.search(r"\d{%d,}" % minlen, text)
    return m.group(0) if m else ""


def clean_job_token(token):
    token = (token or "").strip()
    token = re.sub(r"(?i)\b(CMS|QUOTE|PROPOSAL|BOM|RFQ|FINAL|REV)\b.*$", "", token).strip()
    token = token.strip(" _-.,;:()[]{}")
    token = re.sub(r"\s+", "", token)
    token = re.sub(r'[\\/:*?"<>|]', "_", token)
    upper = token.upper().replace("_", "-")
    # Attachment names often contain CAD/export suffixes like ".step-3.x_t".
    # Those are not customer job numbers, and using them makes folders like
    # BMS-step-3-C18608.
    bad_prefixes = ("STEP-", "STP-", "SLDPRT-", "SLDASM-", "X-T-", "XT-", "PARASOLID-")
    bad_exact = {"STEP", "STP", "SLDPRT", "SLDASM", "X-T", "XT", "PARASOLID"}
    if upper in bad_exact or any(upper.startswith(p) for p in bad_prefixes):
        return ""
    return token[:60]


def first_job_token(text):
    text = text or ""
    # Compact customer/job tokens only. Do not allow spaces here; otherwise
    # normal subject words like "Please quote Dynacast-2223488" can be swallowed
    # and cleaned down to the wrong value.
    patterns = (
        r"\b[A-Za-z][A-Za-z0-9]*(?:[-_][A-Za-z0-9]+)+\b",
        r"\b[A-Z]{1,4}\d{3,}\b",
    )
    for pat in patterns:
        for m in re.finditer(pat, text):
            token = clean_job_token(m.group(0))
            if token and re.search(r"\d", token):
                return token
    return ""


def get_attachment_names(msg):
    names = []
    for part in msg.walk():
        try:
            fn = part.get_filename()
        except Exception:
            fn = None
        if fn:
            names.append(_decode(fn))
    return " ".join(names)


def save_attachments(msg, job):
    """Save every attachment into Downloads\\CMS_Incoming\\<job>. Returns (count, folder)."""
    job = clean_job_token(job) or "unknown"
    base = os.path.join(DOWNLOADS_FOLDER, "CMS_Incoming", job)
    try:
        # Each picked email gets a fresh attachment folder. This prevents stale
        # files from older "unknown" or same-job emails from being copied into
        # the new Ron quote folder.
        incoming_root = os.path.abspath(os.path.join(DOWNLOADS_FOLDER, "CMS_Incoming"))
        base_abs = os.path.abspath(base)
        if base_abs.startswith(incoming_root + os.sep) and os.path.isdir(base_abs):
            shutil.rmtree(base_abs)
        os.makedirs(base_abs, exist_ok=True)
        base = base_abs
    except Exception:
        return 0, base
    n = 0
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        fn = part.get_filename()
        if not fn:
            continue
        fn = _decode(fn)
        safe = re.sub(r'[\\/:*?"<>|]', "_", fn).strip()
        if not safe:
            continue
        try:
            data = part.get_payload(decode=True)
            if data is None:
                continue
            with open(os.path.join(base, safe), "wb") as f:
                f.write(data)
            n += 1
        except Exception:
            pass
    return n, base


def write_output(d):
    folder = os.path.dirname(OUTPUT_FILE)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        for k, v in d.items():
            f.write("%s=%s\n" % (k, str(v).replace("\r", " ").replace("\n", " ").strip()))


HIDE_SPAM = True   # picker hides newsletters / no-reply / marketing mail

_SPAM_FROM = ("no-reply", "noreply", "no_reply", "donotreply", "do-not-reply",
              "notifications@", "newsletter", "mailer-daemon", "marketing@",
              "ads-account", "@info.", "@email.", "@notifications", "googleads")
_SPAM_SUBJECT = ("[newsletter]", "request received", "password change",
                 "get to know your", "updates to our", "starting soon",
                 "find the right tool", "using claude", "bring your ideas",
                 "ways claude", "aha moment", "precision under",
                 "transforming energy", "calling all", "account password",
                 "unsubscribe", "webinar", "back up your pc")


def is_probably_spam(subject, sender):
    s = (subject or "").lower()
    f = (sender or "").lower()
    for p in _SPAM_FROM:
        if p in f:
            return True
    for p in _SPAM_SUBJECT:
        if p in s:
            return True
    return False


def _extract_message_info(msg):
    """Pull Job#, similar-to, ship date, subject from one email message."""
    subject = _decode(msg.get("Subject"))
    body = get_body(msg)
    names = get_attachment_names(msg)
    text = subject + "\n" + body
    job = (extract_number_after(text, "JOB#")
           or extract_number_after(text, "JOB #")
           or extract_number_after(text, "JOB NUMBER"))
    if not job:
        job = first_long_number(names, MIN_JOB_DIGITS)
    if not job:
        job = first_long_number(subject, MIN_JOB_DIGITS)
    if not job:
        job = first_long_number(text + "\n" + names, MIN_JOB_DIGITS)
    if not job:
        job = first_job_token(names)
    if not job:
        job = first_job_token(subject)
    if not job:
        job = first_job_token(text + "\n" + names)
    job = clean_job_token(job)
    return {
        "Found": "1",
        "Subject": subject,
        "CustJob": job,
        "SimilarTo": extract_number_after(text, "SIMILAR TO"),
        "ShipDate": (extract_text_after(text, "SHIP DATE") or extract_text_after(text, "SHIP")),
    }


def do_search():
    out = {"Found": "0", "Subject": "", "CustJob": "", "SimilarTo": "", "ShipDate": "",
           "Attachments": "0", "AttachDir": "", "Error": ""}
    M = None
    try:
        M = imaplib.IMAP4_SSL("imap.gmail.com")
        M.login(GMAIL_ADDRESS, APP_PASSWORD)
        M.select("INBOX")

        kw = SUBJECT_KEYWORD.strip()
        if kw:
            criteria = '(UNSEEN SUBJECT "%s")' % kw if SEARCH_UNREAD_ONLY else '(SUBJECT "%s")' % kw
        else:
            criteria = "(UNSEEN)" if SEARCH_UNREAD_ONLY else "ALL"

        typ, data = M.uid("search", None, criteria)
        uids = data[0].split() if data and data[0] else []

        # If nothing unread matched, fall back to read emails with the keyword
        if not uids and SEARCH_UNREAD_ONLY:
            fb = '(SUBJECT "%s")' % kw if kw else "ALL"
            typ, data = M.uid("search", None, fb)
            uids = data[0].split() if data and data[0] else []

        if uids:
            newest = max(uids, key=lambda u: int(u))      # highest UID = most recent arrival
            typ, msgdata = M.uid("fetch", newest, "(RFC822)")
            msg = email.message_from_bytes(msgdata[0][1])

            info = _extract_message_info(msg)
            out.update(info)

            # Save all attachments into the Downloads folder
            cnt, adir = save_attachments(msg, info["CustJob"])
            out["Attachments"] = str(cnt)
            out["AttachDir"] = adir
            log("saved %d attachment(s) to %s" % (cnt, adir))

            if MARK_AS_READ:
                M.uid("store", newest, "+FLAGS", "\\Seen")

        M.logout()
    except Exception as e:
        out["Error"] = str(e)
        try:
            if M is not None:
                M.logout()
        except Exception:
            pass

    write_output(out)

    if out["Error"]:
        log("ERROR " + out["Error"])
        print("Gmail search error:", out["Error"])
    elif out["Found"] == "1":
        log("found '%s' -> Job#=%s ship=%s similar=%s"
            % (out["Subject"][:60], out["CustJob"], out["ShipDate"], out["SimilarTo"]))
        print("Found newest quote email:", out["Subject"])
        print("  Customer Job#:", out["CustJob"] or "(none)",
              "| Similar to:", out["SimilarTo"] or "(none)",
              "| Ship:", out["ShipDate"] or "(none)")
    else:
        log("no matching email found")
        print('No email with "%s" in the subject was found.' % SUBJECT_KEYWORD)


def send_proposal():
    """Read the proposal handoff written by the macro and email a BOM-pricing summary
    (total shown in red) back to the shop inbox. Skips if the macro already sent it."""
    path = OUTPUT_FILE.replace("cms_email.txt", "cms_proposal.txt")
    d = {}
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            for line in f:
                if "=" in line:
                    k, v = line.split("=", 1)
                    d[k.strip()] = v.strip()
    except Exception as e:
        print("Proposal file not found:", e)
        return

    if d.get("Sent", "0") == "1":
        log("proposal email already sent by macro (CDO); skipping Python send.")
        return

    to = d.get("To", "cms1engineering@gmail.com")
    custjob = d.get("CustJob", "")
    cnum = d.get("CNum", "").replace("-", "")          # example shows C18472 (no hyphen)
    items = d.get("Items", "")
    total = d.get("Total", "0")
    prefix = d.get("CustomerPrefix", "") or d.get("CustomerName", "") or "QUOTE"
    if custjob.upper().startswith((prefix + "-").upper()):
        custjob = custjob[len(prefix)+1:]
    subject = "%s%s %s / PROPOSAL-BOM PRICING" % (prefix, (" - " + custjob) if custjob else "", cnum)

    text = "%s\n\nBOM: %s   $%s\n" % (subject, items, total)
    html = ('<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px">'
            '<p>%s</p>'
            '<p>BOM: %s &nbsp;&nbsp; <span style="color:#d00000"><b>$%s</b></span></p>'
            '</body></html>') % (subject, items, total)

    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    m = MIMEMultipart("alternative")
    m["Subject"] = subject
    m["From"] = GMAIL_ADDRESS
    m["To"] = to
    m.attach(MIMEText(text, "plain"))
    m.attach(MIMEText(html, "html"))
    try:
        s = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        s.login(GMAIL_ADDRESS, APP_PASSWORD)
        s.sendmail(GMAIL_ADDRESS, [to], m.as_string())
        s.quit()
        log("proposal email sent to %s: %s | $%s" % (to, items, total))
        print("Proposal email sent to", to)
        try:    # mark sent so it isn't sent twice
            with open(path, "a", encoding="utf-8") as f:
                f.write("Sent=1\n")
        except Exception:
            pass
    except Exception as e:
        log("proposal email ERROR " + str(e))
        print("Could not send proposal email:", e)


def list_unopened():
    """Return (rows, error). rows = list of (uid, subject, from, date) for unopened
    matching emails, fetched with PEEK so listing them does not mark them read."""
    rows = []
    M = None
    try:
        M = imaplib.IMAP4_SSL("imap.gmail.com")
        M.login(GMAIL_ADDRESS, APP_PASSWORD)
        M.select("INBOX")
        kw = SUBJECT_KEYWORD.strip()
        crit = ('(UNSEEN SUBJECT "%s")' % kw) if kw else "(UNSEEN)"
        typ, data = M.uid("search", None, crit)
        uids = data[0].split() if data and data[0] else []
        for u in uids:
            typ, hd = M.uid("fetch", u, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])")
            subj = frm = dt = ""
            if hd and hd[0]:
                hmsg = email.message_from_bytes(hd[0][1])
                subj = _decode(hmsg.get("Subject"))
                frm = _decode(hmsg.get("From"))
                dt = _decode(hmsg.get("Date"))
            rows.append((u.decode() if isinstance(u, bytes) else str(u), subj, frm, dt))
        M.logout()
        if HIDE_SPAM:
            rows = [r for r in rows if not is_probably_spam(r[1], r[2])]
        # newest first
        rows.sort(key=lambda r: int(r[0]), reverse=True)
        return rows, ""
    except Exception as e:
        try:
            if M:
                M.logout()
        except Exception:
            pass
        return [], str(e)


def process_selected(uids):
    """For each chosen email: download its attachments, write cms_email.txt, then run
    CMS_Launcher.vbs /usemail (which uses that file instead of re-searching) and wait."""
    msgs = []
    M = None
    try:
        M = imaplib.IMAP4_SSL("imap.gmail.com")
        M.login(GMAIL_ADDRESS, APP_PASSWORD)
        M.select("INBOX")
        for u in uids:
            typ, dat = M.uid("fetch", u.encode() if isinstance(u, str) else u, "(RFC822)")
            if dat and dat[0]:
                msgs.append(email.message_from_bytes(dat[0][1]))
        M.logout()
    except Exception as e:
        print("Fetch error:", e)
        try:
            if M:
                M.logout()
        except Exception:
            pass
        return

    here = os.path.dirname(os.path.abspath(__file__))
    vbs = os.path.join(here, "CMS_Launcher.vbs")
    for idx, msg in enumerate(msgs, 1):
        info = _extract_message_info(msg)
        out = {"Found": "1", "Subject": info["Subject"], "CustJob": info["CustJob"],
               "SimilarTo": info["SimilarTo"], "ShipDate": info["ShipDate"],
               "Attachments": "0", "AttachDir": "", "Error": ""}
        cnt, adir = save_attachments(msg, info["CustJob"])
        out["Attachments"] = str(cnt)
        out["AttachDir"] = adir
        write_output(out)
        log("picked %d/%d Job#=%s '%s' (%d files)" % (idx, len(msgs), info["CustJob"], info["Subject"][:50], cnt))
        print("Quoting %d of %d:  Job# %s" % (idx, len(msgs), info["CustJob"] or "?"))
        if os.path.exists(vbs):
            try:
                subprocess.call(["wscript", vbs, "/usemail"])   # waits for this job to finish
            except Exception as e:
                print("  launcher error:", e)
            # The macro wrote cms_proposal.txt during the job. Send it now (Python SMTP
            # is reliable); send_proposal() skips it if the macro already sent via CDO.
            try:
                send_proposal()
            except Exception as e:
                print("  proposal email error:", e)
        else:
            print("  CMS_Launcher.vbs not found next to this script; wrote cms_email.txt only.")
    print("Done. Processed %d email(s)." % len(msgs))


def pick_unopened():
    """Show a checkbox list (with Select All) of unopened quote emails to choose which to do."""
    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception as e:
        print("Tkinter not available:", e)
        return

    rows, err = list_unopened()
    root = tk.Tk()
    root.title("CMS - Select emails to quote")
    root.geometry("780x540")
    if err:
        tk.Label(root, text="Gmail error: " + err, fg="red", wraplength=740).pack(padx=14, pady=20)
        root.mainloop()
        return
    if not rows:
        tk.Label(root, text="No unopened emails were found in the inbox.").pack(padx=14, pady=20)
        root.mainloop()
        return

    tk.Label(root, text="Unopened emails - check the job(s) you want to quote:",
             font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(12, 4))

    vars_ = []
    selall = tk.IntVar()

    def toggle_all():
        for v in vars_:
            v.set(selall.get())

    tk.Checkbutton(root, text="Select All", variable=selall, command=toggle_all,
                   font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14)

    canv = tk.Canvas(root, highlightthickness=0)
    inner = tk.Frame(canv)
    sb = tk.Scrollbar(root, orient="vertical", command=canv.yview)
    canv.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    canv.pack(fill="both", expand=True, padx=14, pady=6)
    canv.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda e: canv.configure(scrollregion=canv.bbox("all")))

    for (uid, subj, frm, dt) in rows:
        v = tk.IntVar()
        vars_.append(v)
        label = "%s    -    %s    (%s)" % (subj[:74] or "(no subject)", frm[:30], dt[:24])
        tk.Checkbutton(inner, variable=v, text=label, anchor="w", justify="left",
                       wraplength=700).pack(anchor="w")

    def do_quote():
        chosen = [rows[i][0] for i, v in enumerate(vars_) if v.get()]
        if not chosen:
            messagebox.showinfo("CMS", "Check at least one email first.")
            return
        root.destroy()
        process_selected(chosen)

    tk.Button(root, text="Quote Selected", command=do_quote, bg="#1e40af", fg="white",
              font=("Segoe UI", 11, "bold"), padx=16, pady=6).pack(pady=12)
    root.mainloop()


def launch_full_flow():
    """User ran this script directly -> open the webapp inbox (no Tk picker)."""
    log("redirecting user to webapp inbox")
    print("Email quoting is now done in the CMS AI Quoting webapp.")
    print("Open http://127.0.0.1:8000/email and click the blue Quote button.")
    try:
        import webbrowser
        webbrowser.open("http://127.0.0.1:8000/email")
    except Exception:
        pass
    here = os.path.dirname(os.path.abspath(__file__))
    bat = os.path.join(here, "webapp", "START_CMS_QUOTING_APP.bat")
    if os.path.exists(bat):
        print("If the site cannot be reached, double-click:")
        print(" ", bat)


def main():
    if "--from-launcher" in sys.argv:
        do_search()
    elif "--send-proposal" in sys.argv:
        send_proposal()
    elif "--pick" in sys.argv:
        pick_unopened()
    else:
        launch_full_flow()


if __name__ == "__main__":
    main()


