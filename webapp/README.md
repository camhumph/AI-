# CMS AI Quoting Web App

A full-stack quoting console for the CMS mold-geometry AI pipeline, wired
**live into Module6121** (the SolidWorks quoting macro at the repo root):

- **Module6121 integration (seamless, local-only):** right after the macro
  writes `XT_Export_CAD_Dimensions.csv`, it POSTs the CSV path to this app at
  `http://127.0.0.1:8000/api/vba/classify`. The app runs the AI classifier
  and returns one resolved role per CAD part
  (`A Plate`, `B Plate`, `SC Retainer Plate`, `Bottom Ejector Plate`, rails,
  leader pins, latch locks...). The macro applies those roles to its standard
  plate list before filling the quote workbook and steel sheet. Every macro
  run also auto-registers the job here, so it appears in the dashboard
  instantly.
- **BMS protection:** BMS / pot-block bases NEVER touch the AI. The macro
  guards them (`RunAiBridgeClassification` exits for non-standard bases) and
  the app refuses to classify any job marked `base_type=bms` -- their proven
  BOM-driven flow is untouched. BMS jobs still register in the dashboard,
  clearly badged.
- Browse/upload CAD job folders and (re)run the AI classifier on standard
  bases from the UI.
- See every part of a quote in one place: rendered JPEG views, an STL 3D
  viewer, a grouped/priced parts table, and any documents (quote sheet,
  steel sheet, PDFs) for that job.
- A pricing engine with a total quote price per job. Hardware roles price
  from the shop's real `Purchased Components Prices.csv` (same file the
  macro reads); plate rates are editable placeholders in Settings.
- An email inbox (IMAP) with a reply composer (SMTP) and a "Quote This"
  button that jumps straight into the matching job.

```
webapp/
  backend/                  FastAPI app (Python)
  frontend/                 Vite + React + TypeScript + Tailwind UI
  START_CMS_QUOTING_APP.bat one-click local-only start for the shop PC
```

## Local-only hosting (off the network)

This app is designed to run **on the same PC as SolidWorks/Module6121 and
bind to 127.0.0.1 only** -- it is never exposed to the network or internet.
`START_CMS_QUOTING_APP.bat` starts it that way, and the macro's
`AI_BRIDGE_URL` constant points at `http://127.0.0.1:8000`.

## Running

### One-click on the shop PC (production, local-only)

First-time setup:

```bat
cd webapp\backend  && pip install -r requirements.txt
cd webapp\frontend && npm install && npm run build
```

Then just double-click `webapp\START_CMS_QUOTING_APP.bat`. One process on
`127.0.0.1:8000` serves both the UI and the API (the backend serves the
built frontend automatically when `frontend/dist` exists).

### Development mode

```bash
cd webapp/backend
pip install -r requirements.txt
python3 seed_demo_data.py         # optional: seeds a real T001015 demo job
python3 -m uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload

cd webapp/frontend
npm install
npm run dev                        # http://localhost:5173, proxies /api to :8000
```

## How the Module6121 integration works

Inside `Module6121.bas` (repo root):

| Piece | What it does |
|---|---|
| `RunAiBridgeClassification` | Called in `ProcessOneJob` and `RunActiveAssembly` right after `DetectBaseTypeIsStandard`. Skips + logs for BMS/pot-block bases. POSTs the CAD CSV path to `/api/vba/classify`; falls back to a bridge file in `C:\CMS_Local_Workspace\AI_Bridge\` when the app isn't running; if neither exists, the macro's own geometry rules run alone (nothing breaks). |
| `BuildStdFromAiBridge` | New first-priority source in `ClassifyStandardBasePlates`: applies the AI's per-part roles (>= 3 plates required, else geometry fallback). Logs every AI-assigned plate with confidence. |
| `AiPlateNameForRole` | Maps AI role keys to the macro's plate names and existing quote rows -- `bottom_ejector_plate` lands on the same PIN row the thick ejector plate always used, `ejector_plate` keeps the EJECTOR row. Nothing about the quote workbook layout changes. |
| Latch locks | `latch_lock` roles are logged as secondary-parting-line markers and never flip A/B or take a leader-pin price. |

The response CSV format (also written to `data/vba_bridge/<job>_part_names.csv`):

```
Index,Component,Role,ResolvedName,Confidence,Quote,Price,SecondaryPartingLine
```

**Deploying the macro changes:** only `Module6121.bas` (repo root) carries the
new integration. `Module6121.swb` / `Module6121.swp` are still the OLD
compiled/source copies -- SolidWorks can't be run in this environment to
rebuild them. On the shop PC: open the SolidWorks VBA editor, remove the old
module, import the updated `Module6121.bas`, and re-save the `.swb`/`.swp`
(the same way previous macro updates were deployed). `CMS_Launcher.vbs` runs
the `.swb` first, so this step is required before the launcher picks up the
AI bridge.

## Configuration (environment variables)

All configuration is environment-driven so the same code runs in this sandbox
and on a real CMS machine pointed at `C:\CMS_Local_Workspace`.

| Variable | Purpose | Default |
|---|---|---|
| `CMS_JOBS_ROOT` | Folder containing one sub-folder per quote job | `backend/data/jobs` |
| `CMS_GEOMETRY_CLASSIFIER_DIR` | Path to `geometry_classifier/` | `../../geometry_classifier` |
| `CMS_VBA_BRIDGE_DIR` | Where Module6121 bridge CSV/JSON exports land | `backend/data/vba_bridge` |
| `CMS_PRICING_CONFIG` | Pricing rates JSON file | `backend/data/pricing_config.json` |
| `CMS_IMAP_HOST` / `CMS_IMAP_PORT` / `CMS_IMAP_USER` / `CMS_IMAP_PASSWORD` / `CMS_IMAP_FOLDER` | IMAP inbox connection | unset (inbox shows "connect your email") |
| `CMS_SMTP_HOST` / `CMS_SMTP_PORT` / `CMS_SMTP_USER` / `CMS_SMTP_PASSWORD` / `CMS_SMTP_FROM` | SMTP for sending replies | unset (reply disabled) |

**Set these as Secrets in the Cursor Dashboard (Cloud Agents > Secrets) for
this repo**, not in code. For Gmail/Outlook/Office365, use an app password,
not your normal login password.

## Job folder layout

Each job is a folder under `CMS_JOBS_ROOT`:

```
<JobID>/
  meta.json                          display name, customer, notes
  XT_Export_CAD_Dimensions.csv       raw SolidWorks CAD export (source of truth)
  classification.csv / .json         AI classification result (written by the app)
  images/*.jpg                       rendered assembly views
  models/*.stl                       3D models (STL viewer)
  documents/*                        quote sheet / steel sheet / PDFs
```

Point `CMS_JOBS_ROOT` at your real `C:\CMS_Local_Workspace` (or a synced
copy) to browse real jobs instead of the seeded demo.

## Pricing

Hardware roles (leader pins, bushings, latch locks/straps, support pillars)
price from the shop's real **`Purchased Components Prices.csv`** at the repo
root -- the same file Module6121 reads (override the path with
`CMS_PURCHASED_PRICES_CSV`). Only rows with a non-zero `UnitPrice` are used.

Plate rates are still **placeholders** (no plate price book exists in this
repo). Edit them from the Settings page (persisted to `CMS_PRICING_CONFIG`).
Pricing modes per role:

- `flat` -- fixed price per unit (hardware, latch locks, pins)
- `per_cuin` -- rate x (Thickness x Width x Length), a rough material-volume
  proxy for plates
- `per_inch` -- rate x Length, for rails

Note: latch-lock/PLC/safety-strap hardware classifies with `quote=False` by
default (dozens of individual fasteners inside one latch assembly would
otherwise wildly inflate the total). Flip that in
`geometry_classifier/qwen_classify_xt_csv.py` if you'd rather quote it, ideally
per-assembly rather than per-fastener.

## SECURITY: Gmail app password (action required)

The previous versions of `Module6121.bas` / `cms_gmail_search.py` uploaded to
GitHub contained a **hardcoded Gmail app password**. Even though it has been
scrubbed from the current files, it is still visible in the repository's git
history, so you must:

1. **Revoke it now** at myaccount.google.com/apppasswords
   (account `cms1engineering@gmail.com`).
2. Create a new app password and save it -- alone on one line -- in
   `C:\CMS_Local_Workspace\gmail_app_password.txt` on the shop PC.
   Both the macro and `cms_gmail_search.py` now read it from that file.

Also, the macro's automatic proposal email is now gated by
`PROPOSAL_EMAIL_MODE` in `Module6121.bas` (`"PROMPT"` by default -- it asks
before sending; set `"OFF"` to only write the preview file, `"AUTO"` for the
old fire-and-forget behavior).

## Known limitations / honesty notes

- **Email** only activates once IMAP/SMTP secrets are set; there is no fake
  inbox data -- it fails closed with a clear "connect your email" state.
- **Plate pricing** is placeholder math (hardware pricing is real, from the
  shop CSV).
- **"Quote This" from an email** uses a simple job-token regex
  (`[A-Z]{1,2}\d{4,6}`, e.g. `J8420`, `T001015`, `C18606`) matched against
  known job IDs in the subject/attachment names. It's a heuristic, not
  guaranteed matching.
- **No STL files ship in this repo** (none exist in the source workspace);
  upload one from the Quote Detail page's "3D Model" tab to see the viewer.
- **The VBA changes are untestable here** (no SolidWorks/Windows/Excel in
  this environment). The macro edits are conservative -- every AI call is
  wrapped in error handlers that fall back to the existing geometry rules --
  but the first run on the shop PC should be watched with the log open.
- **AI for BMS bases is intentionally off.** If you later want the AI
  trained on BMS/pot-block bases, the path is: collect corrected BMS
  examples via the same CORRECT_ME workflow, add pot/holder roles to the
  classifier, and only then relax the guards (macro + backend both).
