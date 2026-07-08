# CMS AI Quoting Web App

A full-stack quoting console for the CMS mold-geometry AI pipeline:

- Browse/upload CAD job folders and (re)run the AI classifier
  (`geometry_classifier/qwen_classify_xt_csv.py`) on them.
- See every part of a quote in one place: rendered JPEG views, an STL 3D
  viewer, a grouped/priced parts table, and any documents (quote sheet,
  steel sheet, PDFs) for that job.
- A configurable pricing engine that computes a total quote price per job.
- An email inbox (IMAP) with a reply composer (SMTP) and a "Quote This"
  button that jumps straight into the matching job.
- A "Module6121 AI bridge" that exports AI-resolved part names/roles to a
  flat CSV/JSON every time a job is classified or viewed, for your VBA macro
  to read.

```
webapp/
  backend/     FastAPI app (Python)
  frontend/    Vite + React + TypeScript + Tailwind UI
  vba/         Paste-in VBA helper module for Module6121 integration
```

## Important: Module6121.bas

This app can **export** AI-resolved part names/prices for a VBA macro to
consume (see "Module6121 bridge" below and `vba/AI_Bridge_Import.bas`), but
it does not modify your actual `Module6121.bas` because that file has not
made it into this repo/environment yet (file uploads in this chat don't land
in the cloud agent's filesystem). To get the AI wired directly into your real
macro's Subs/named ranges:

- Paste the contents of `Module6121.bas` into the chat, **or**
- Add it to this git repo (e.g. `vba/Module6121.bas`) and push it.

Once I can see the actual cell references / named ranges Module6121 uses to
place part names, I can wire `webapp/vba/AI_Bridge_Import.bas`'s
`WriteRowToQuote` routine directly into it instead of leaving it as a
documented placeholder.

## Running locally

### 1. Backend (FastAPI)

```bash
cd webapp/backend
pip install -r requirements.txt   # fastapi, uvicorn, python-multipart, pydantic
python3 seed_demo_data.py         # optional: seeds a real T001015 demo job
python3 -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

### 2. Frontend (Vite + React)

```bash
cd webapp/frontend
npm install
npm run dev -- --host
```

Open the printed URL (defaults to `http://localhost:5173`). The dev server
proxies `/api/*` to `http://localhost:8000`.

### Production build

```bash
cd webapp/frontend
npm run build     # outputs frontend/dist
```

Serve `frontend/dist` with any static host, or add a `StaticFiles` mount to
`backend/app/main.py` to serve it directly from FastAPI.

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

`backend/app/pricing.py` ships **placeholder** rates -- there is no real CMS
price book in this repo. Edit rates from the Settings page in the UI (they
persist to `CMS_PRICING_CONFIG`), or replace `DEFAULT_RATES` directly. Each
role has a pricing mode:

- `flat` -- fixed price per unit (hardware, latch locks, pins)
- `per_cuin` -- rate x (Thickness x Width x Length), a rough material-volume
  proxy for plates
- `per_inch` -- rate x Length, for rails

Note: latch-lock/PLC/safety-strap hardware currently classifies with
`quote=False` by default (dozens of individual fasteners inside one latch
assembly would otherwise wildly inflate the placeholder total). Flip that in
`geometry_classifier/qwen_classify_xt_csv.py` if you'd rather quote it, ideally
per-assembly rather than per-fastener.

## Module6121 bridge

Every time a job's quote sheet is requested (`GET /api/jobs/{id}/quote-sheet`,
which the Quote Detail page calls on load), the backend writes:

```
<CMS_VBA_BRIDGE_DIR>/<JobID>_part_names.csv
<CMS_VBA_BRIDGE_DIR>/<JobID>_part_names.json
```

CSV columns: `Index, Component, Role, ResolvedName, Quote, Price, SecondaryPartingLine`.

See `vba/AI_Bridge_Import.bas` for a ready-to-paste VBA subroutine
(`CMS_AI_ImportPartNames`) that reads this CSV and either writes it to an
`AI_Import` worksheet (works out of the box) or can be wired into your real
Module6121 quote-population logic (needs your actual macro -- see above).

## Known limitations / honesty notes

- **Email** only activates once IMAP/SMTP secrets are set; there is no fake
  inbox data -- it fails closed with a clear "connect your email" state.
- **Pricing** is placeholder math, not CMS's real price book.
- **"Quote This" from an email** uses a simple job-token regex
  (`[A-Z]{1,2}\d{4,6}`, e.g. `J8420`, `T001015`, `C18606`) matched against
  known job IDs in the subject/attachment names. It's a heuristic, not
  guaranteed matching.
- **No STL files ship in this repo** (none exist in the source workspace);
  upload one from the Quote Detail page's "3D Model" tab to see the viewer.
- **Module6121.bas integration is a documented placeholder** until the real
  macro file is shared (see above).
