@echo off
rem ============================================================
rem CMS AI Quoting - build and run, one step.
rem
rem Double-click this, or:
rem     cd /d C:\CMS_AI\webapp
rem     RUN.bat
rem
rem Order matters and this is why the file exists:
rem
rem   1. Install the Datum STL engine FIRST. The frontend imports it
rem      statically, so if it is missing when Vite bundles you get
rem      UNRESOLVED_IMPORT and no app at all. Installing before the build is
rem      the whole point.
rem   2. Build the frontend.
rem   3. Start the backend, which serves the built frontend.
rem
rem A stub engine is committed at src\lib\datumEngine.js so the build cannot
rem break even if step 1 finds nothing -- the Machining tab just falls back to
rem the server's coarse estimate instead.
rem ============================================================
setlocal

echo.
echo ============================================================
echo  CMS AI Quoting
echo ============================================================
echo.

rem ---- pick a Python -----------------------------------------
where python >nul 2>&1
if errorlevel 1 (
  where py >nul 2>&1
  if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.12+ and try again.
    pause
    exit /b 1
  )
  set "PYTHON=py -3"
) else (
  set "PYTHON=python"
)

rem ---- 1. Datum STL engine, before the bundler needs it ------
echo [1/3] Datum STL engine...
%PYTHON% "%~dp0..\install_datum_engine.py" --if-missing
echo.

rem ---- 2. frontend --------------------------------------------
echo [2/3] Building frontend...
pushd "%~dp0frontend"
if not exist "node_modules" (
  echo       node_modules missing - running npm install first...
  call npm install
  if errorlevel 1 (
    echo ERROR: npm install failed.
    popd
    pause
    exit /b 1
  )
)
call npm run build
if errorlevel 1 (
  echo.
  echo ERROR: frontend build failed. The app was NOT started.
  echo        Scroll up for the first error - that is the real one.
  popd
  pause
  exit /b 1
)
popd
echo.

rem ---- 3. run -------------------------------------------------
echo [3/3] Starting server...
echo.
call "%~dp0START_CMS_QUOTING_APP.bat"

endlocal
