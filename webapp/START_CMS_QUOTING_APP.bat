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
set CMS_WORKSPACE_ROOT=\\Mycloudex2ultra\mexico\Downloads

rem Point the app at the real job folders on this machine (uncomment to use):
rem set CMS_JOBS_ROOT=C:\CMS_Local_Workspace\AI_Jobs

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
