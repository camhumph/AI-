# AGENTS.md

## Cursor Cloud specific instructions

### What this repo is
CMS mold-quoting system. Two very different halves:

1. **SolidWorks VBA quoting macro** — `Module6121.bas` / `.swb` / `.swp`, launched by
   `CMS_Launcher.vbs` + `RunSolidWorksMacro.ps1`, plus `cms_visual_inspect.ps1`.
   This is the main product but it is **Windows + SolidWorks + Excel COM only**. It
   cannot be built, run, or tested on the Linux cloud VM (no SolidWorks / no COM).
   Treat changes to the `.bas` macro as **static-only** here; they must be validated
   on a Windows machine with SolidWorks installed.

2. **Python AI stack** — this is the part that actually runs on Linux and is what
   "the AI" refers to:
   - `server/cms_local_mold_ai.py` — FastAPI web app that runs a YOLO model on an
     uploaded mold image and returns detections (`/`, `/health`, `/upload`, `/predict-json`).
   - `geometry_classifier/` — Qwen/Ollama-based CAD-CSV part classifier + dataset tools.
   - `cms_price_lookup.py`, `cms_gmail_search.py`, `yolo_box_editor.py` — supporting CLIs.
   - `models/*.pt` — trained YOLO weights. `yolo26s.pt` (repo root) is the generic fallback.

### Running the Python AI stack (Linux)
- Use the virtualenv at `venv/` (git-ignored). Activate with `source venv/bin/activate`.
- Run the AI server: `uvicorn cms_local_mold_ai:app --host 127.0.0.1 --port 8077`
  (from the folder containing the module, e.g. `server/`).
- **Gotcha (hardcoded Windows paths):** the Python files hardcode `C:\CMS_AI` and
  `C:\CMS_Local_Workspace`. On Linux `Path(r"C:\CMS_AI")` becomes a **relative** folder
  literally named `C:\CMS_AI` created under the current working directory. The server
  looks for its model at `C:\CMS_AI/models/cms_mold_yolo26.pt` and otherwise falls back
  to `yolo26s.pt`. To exercise the real model without editing source, run from a scratch
  dir that contains a `C:\CMS_AI/models/cms_mold_yolo26.pt` (copy one of `models/*.pt`).
  Running these scripts from the repo root will drop a stray `C:\...` folder — delete it.
- **Model views are specialized:** `cms_mold_right_view8.pt` detects strongly on
  RIGHT-view renders; `cms_mold_yolo26.pt` and the full-view models are weaker/undertrained.
  Pick a model that matches the render view when testing detections.

### Lint / test / build
- There is no configured linter, test suite, `requirements.txt`, or build step.
  Use `python -m py_compile <files>` as the syntax/lint check.
- Model inference is CPU-only in this environment (torch CPU wheels).
