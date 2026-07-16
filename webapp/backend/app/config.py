"""Central configuration for the CMS AI Quoting web app.

Everything here is overridable with environment variables so the same code
runs unchanged on this cloud sandbox and on a real CMS Windows machine
pointed at C:\\CMS_Local_Workspace.
"""
import os
from pathlib import Path

BACKEND_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ.get("CMS_DATA_DIR", BACKEND_DIR / "data"))

# Root folder that contains one sub-folder per quote job. Each job folder can
# contain: XT_Export_CAD_Dimensions.csv, classification CSV/JSON, *.jpg view
# renders, *.stl models, and any PDF/quote-sheet/steel-sheet documents.
# On a real CMS machine this should point at C:\CMS_Local_Workspace.
JOBS_ROOT = Path(os.environ.get("CMS_JOBS_ROOT", DATA_DIR / "jobs"))

# Folders to browse when picking a quote job (C-number folders on the shop PC).
# Default: network Downloads where month folders / job packages land.
WORKSPACE_ROOT = Path(
    os.environ.get("CMS_WORKSPACE_ROOT", r"\\Mycloudex2ultra\mexico\Downloads")
)
# Additional roots scanned for existing quote folders (network drive, month folders).
WORKSPACE_EXTRA_ROOTS = [
    p.strip()
    for p in os.environ.get("CMS_WORKSPACE_EXTRA_ROOTS", "").split(";")
    if p.strip()
]

# Where the AI classifier lives (repo-relative), used to actually (re)run
# classification against a job's raw XT_Export_CAD_Dimensions.csv.
GEOMETRY_CLASSIFIER_DIR = Path(
    os.environ.get(
        "CMS_GEOMETRY_CLASSIFIER_DIR",
        BACKEND_DIR.parent.parent / "geometry_classifier",
    )
)

# Where AI -> Module6121 "bridge" exports are written. Module6121.bas (or any
# VBA macro) reads the CSV/JSON here to pull AI-resolved part names/roles.
VBA_BRIDGE_DIR = Path(os.environ.get("CMS_VBA_BRIDGE_DIR", DATA_DIR / "vba_bridge"))

PRICING_CONFIG_PATH = Path(
    os.environ.get("CMS_PRICING_CONFIG", DATA_DIR / "pricing_config.json")
)

# The shop's real purchased-component price list (same CSV Module6121 reads).
# Hardware roles (leader pins, bushings, latch locks/straps...) price from
# this file when a usable UnitPrice exists.
PURCHASED_PRICES_CSV = Path(
    os.environ.get(
        "CMS_PURCHASED_PRICES_CSV",
        BACKEND_DIR.parent.parent / "Purchased Components Prices.csv",
    )
)

# --- Email (IMAP/SMTP) ---------------------------------------------------
# Loaded from webapp/backend/data/email_credentials.json (Settings page) with
# optional CMS_* environment overrides. Never hardcoded or committed.
def _email_settings():
    from . import credentials

    cred = credentials.load()
    return cred


def reload_email_settings():
    """Re-read credentials after Settings save (called from main.py)."""
    global IMAP_HOST, IMAP_PORT, IMAP_USER, IMAP_PASSWORD, IMAP_FOLDER, IMAP_USE_SSL
    global SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_FROM
    global EMAIL_CONFIGURED, SMTP_CONFIGURED
    cred = _email_settings()
    IMAP_HOST = cred.get("imap_host", "")
    IMAP_PORT = int(cred.get("imap_port") or 993)
    IMAP_USER = cred.get("imap_user", "")
    IMAP_PASSWORD = cred.get("imap_password", "")
    IMAP_FOLDER = cred.get("imap_folder", "INBOX")
    IMAP_USE_SSL = bool(cred.get("imap_ssl", True))
    SMTP_HOST = cred.get("smtp_host", "")
    SMTP_PORT = int(cred.get("smtp_port") or 587)
    SMTP_USER = cred.get("smtp_user", "")
    SMTP_PASSWORD = cred.get("smtp_password", "")
    SMTP_FROM = cred.get("smtp_from") or SMTP_USER
    EMAIL_CONFIGURED = bool(IMAP_HOST and IMAP_USER and IMAP_PASSWORD)
    SMTP_CONFIGURED = bool(SMTP_HOST and SMTP_USER and SMTP_PASSWORD)


reload_email_settings()

JOBS_ROOT.mkdir(parents=True, exist_ok=True)
VBA_BRIDGE_DIR.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)
