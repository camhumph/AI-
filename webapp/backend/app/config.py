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

# --- Email (IMAP/SMTP) ---------------------------------------------------
# These are intentionally read from the environment only (Cursor Cloud
# Agent Secrets, or a real .env on the CMS machine). Nothing email-related
# is ever hardcoded or committed.
IMAP_HOST = os.environ.get("CMS_IMAP_HOST", "")
IMAP_PORT = int(os.environ.get("CMS_IMAP_PORT", "993"))
IMAP_USER = os.environ.get("CMS_IMAP_USER", "")
IMAP_PASSWORD = os.environ.get("CMS_IMAP_PASSWORD", "")
IMAP_FOLDER = os.environ.get("CMS_IMAP_FOLDER", "INBOX")
IMAP_USE_SSL = os.environ.get("CMS_IMAP_SSL", "true").lower() != "false"

SMTP_HOST = os.environ.get("CMS_SMTP_HOST", "")
SMTP_PORT = int(os.environ.get("CMS_SMTP_PORT", "587"))
SMTP_USER = os.environ.get("CMS_SMTP_USER", "")
SMTP_PASSWORD = os.environ.get("CMS_SMTP_PASSWORD", "")
SMTP_FROM = os.environ.get("CMS_SMTP_FROM", SMTP_USER)

EMAIL_CONFIGURED = bool(IMAP_HOST and IMAP_USER and IMAP_PASSWORD)
SMTP_CONFIGURED = bool(SMTP_HOST and SMTP_USER and SMTP_PASSWORD)

JOBS_ROOT.mkdir(parents=True, exist_ok=True)
VBA_BRIDGE_DIR.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)
