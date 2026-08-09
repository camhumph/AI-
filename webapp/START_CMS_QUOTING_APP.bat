@echo off
rem ============================================================
rem CMS AI Quoting - LOCAL-ONLY start script
rem
rem Start sequence (rebuild UI + start):
rem   cd C:\CMS_AI\webapp\frontend
rem   npm run build
rem   cd C:\CMS_AI\webapp
rem   START_CMS_QUOTING_APP.bat
rem
rem Or double-click:  webapp\START_SEQUENCE.bat
rem
rem Binds to 127.0.0.1 so the app is reachable ONLY from this
rem machine (never the network/internet). Module6121 talks to it
rem at http://127.0.0.1:8000 on the same PC.
rem
rem First-time setup (once):
rem   cd webapp\backend  and run:  pip install -r requirements.txt
rem   cd webapp\frontend and run:  npm install && npm run build
rem
rem Launcher diagnostics (if a quote gets stuck):
rem   C:\CMS_Local_Workspace\CMS_Quote_Log.txt
rem   C:\CMS_Local_Workspace\cms_launcher_status.txt
rem   C:\CMS_Local_Workspace\cms_macro_started.txt
rem   C:\CMS_Local_Workspace\cms_macro_error.txt
rem   C:\CMS_Local_Workspace\cms_macro_status.txt
rem ============================================================
cd /d "%~dp0backend"

rem Bridge exports go where Module6121's offline fallback also looks
set CMS_VBA_BRIDGE_DIR=C:\CMS_Local_Workspace\AI_Bridge

rem Email credentials + pricing config (Settings page writes here)
set CMS_DATA_DIR=C:\CMS_Local_Workspace\cms_data

rem Folder browser for C-number quote jobs (month folders live under Downloads)
rem
rem This share is only reachable on the company wifi. You do NOT need to change
rem it when working off the network: the app also searches your local Downloads,
rem OneDrive Downloads, Desktop and C:\CMS_Local_Workspace automatically, opens
rem the folder picker on whichever one answers, and shows the share crossed out
rem while it is offline. Reconnect and restart the app to get it back.
rem
rem To add another place to look, semicolon-separated:
rem   set CMS_WORKSPACE_EXTRA_ROOTS=D:\Jobs;E:\FromCustomer
set CMS_WORKSPACE_ROOT=\\Mycloudex2ultra\mexico\Downloads

rem PIN the job registry explicitly.
rem
rem This must never be left to the default. config.py derives JOBS_ROOT from
rem CMS_DATA_DIR when CMS_JOBS_ROOT is unset, so starting the backend any other
rem way (plain "python start_cms.py", an IDE, a service) silently pointed the
rem Quotes list at webapp\backend\data\jobs -- 3 folders instead of the 13 in
rem cms_data\jobs. Every job "disappeared" and nothing had been deleted.
rem
rem Do NOT set this to C:\CMS_Local_Workspace itself. That is the macro's
rem staging root, and quote_pipeline.stage_job_to_local_workspace rmtree's
rem C:\CMS_Local_Workspace\C##### on every re-quote -- which would delete the
rem registry folder, meta.json, classification.json, images and models with it.
set CMS_JOBS_ROOT=C:\CMS_Local_Workspace\cms_data\jobs

rem Training scan folder (BMS + standard jobs for Settings ^> Run Training Scan)
set CMS_TRAINING_ROOT=C:\Users\lenovo\Downloads\TRAINING

rem SolidWorks 2023 only — "(3)" install; plain SOLIDWORKS path is 2025 on this PC
set CMS_SOLIDWORKS_EXE=C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (3)\SLDWORKS.EXE
set CMS_SOLIDWORKS_PROGID=SldWorks.Application.31

echo.
echo CMS AI Quoting - start sequence reminder:
echo   cd C:\CMS_AI\webapp\frontend
echo   npm run build
echo   cd C:\CMS_AI\webapp
echo   START_CMS_QUOTING_APP.bat
echo.
echo Checking Python...
where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.12+ and run:
    echo   pip install -r requirements.txt
    pause
    exit /b 1
  )
  set PYTHON=py -3
) else (
  set PYTHON=python
)

echo Installing backend dependencies if needed...
%PYTHON% -m pip install -q -r requirements.txt
if errorlevel 1 (
  echo ERROR: pip install failed. Run manually: pip install -r requirements.txt
  pause
  exit /b 1
)

rem The Datum STL engine is a single JS file that has to sit at
rem webapp\frontend\src\lib\datumEngine.js for the bundle to resolve it.
rem --if-missing makes this a no-op once installed, so it is safe on every start.
rem If the file cannot be found it warns and continues: the app still runs, and
rem the Machining tab falls back to the server's coarse pattern estimate.
echo Checking Datum STL engine...
%PYTHON% "%~dp0..\install_datum_engine.py" --if-missing

if not exist "..\frontend\dist\index.html" (
  echo.
  echo WARNING: Frontend not built yet. Run:
  echo   cd C:\CMS_AI\webapp\frontend
  echo   npm run build
  echo Or double-click: webapp\START_SEQUENCE.bat
  echo.
  echo The API will still work at http://127.0.0.1:8000/api/health
  echo.
)

echo Starting CMS AI Quoting on http://127.0.0.1:8000 (local machine only)...
echo The browser opens automatically once the server is ready.
echo If a quote sticks, open the red status text in the app or:
echo   C:\CMS_Local_Workspace\CMS_Quote_Log.txt
echo Press Ctrl+C in this window to stop the server.
echo.
%PYTHON% start_cms.py
if errorlevel 1 pause
