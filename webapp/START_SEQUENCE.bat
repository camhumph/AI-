@echo off
rem ============================================================
rem CMS AI Quoting - BUILD FRONTEND THEN START
rem
rem Exact start sequence:
rem   cd C:\CMS_AI\webapp\frontend
rem   npm run build
rem   cd C:\CMS_AI\webapp
rem   START_CMS_QUOTING_APP.bat
rem
rem This script does that from wherever the repo lives.
rem ============================================================
setlocal
cd /d "%~dp0"

echo.
echo ===== CMS AI Quoting start sequence =====
echo 1^) cd webapp\frontend
echo 2^) npm run build
echo 3^) cd webapp
echo 4^) START_CMS_QUOTING_APP.bat
echo ========================================
echo.

cd /d "%~dp0frontend"
if not exist "package.json" (
  echo ERROR: frontend\package.json not found. Are you in the CMS_AI repo?
  pause
  exit /b 1
)

where npm >nul 2>&1
if errorlevel 1 (
  echo ERROR: npm not found. Install Node.js, then re-open this window.
  pause
  exit /b 1
)

if not exist "node_modules\" (
  echo npm install ^(first time^)...
  call npm install
  if errorlevel 1 (
    echo ERROR: npm install failed.
    pause
    exit /b 1
  )
)

echo Building frontend ^(npm run build^)...
call npm run build
if errorlevel 1 (
  echo ERROR: npm run build failed.
  pause
  exit /b 1
)

cd /d "%~dp0"
echo.
echo Starting app via START_CMS_QUOTING_APP.bat ...
echo.
call "%~dp0START_CMS_QUOTING_APP.bat"
