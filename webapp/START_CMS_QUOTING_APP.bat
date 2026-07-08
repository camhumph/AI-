@echo off
rem ============================================================
rem CMS AI Quoting - LOCAL-ONLY start script
rem
rem Binds to 127.0.0.1 so the app is reachable ONLY from this
rem machine (never the network/internet). Module6121 talks to it
rem at http://127.0.0.1:8000 on the same PC.
rem
rem First-time setup (once):
rem   cd webapp\backend  and run:  pip install -r requirements.txt
rem   cd webapp\frontend and run:  npm install && npm run build
rem ============================================================
cd /d "%~dp0backend"

rem Bridge exports go where Module6121's offline fallback also looks
rem (AI_BRIDGE_FILE_DIR in the macro), so the two always stay in sync.
set CMS_VBA_BRIDGE_DIR=C:\CMS_Local_Workspace\AI_Bridge

rem Point the app at the real job folders on this machine (uncomment to use):
rem set CMS_JOBS_ROOT=C:\CMS_Local_Workspace\AI_Jobs

echo Starting CMS AI Quoting on http://127.0.0.1:8000 (local machine only)...
start "" http://127.0.0.1:8000
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
