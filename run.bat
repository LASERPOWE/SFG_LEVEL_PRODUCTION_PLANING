@echo off
REM Double-click this file to start the dashboard.
setlocal
cd /d "%~dp0"

echo ============================================================
echo  Laser SFG Production Planning Dashboard
echo ============================================================
echo.

REM Pick the best Python on this machine.  Prefer the 'py' launcher
REM (recommended on Windows), then fall back to 'python'.
where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py"
) else (
  set "PY=python"
)

echo Using interpreter: %PY%
%PY% --version
echo.

echo Installing/updating dependencies (Flask, requests)...
%PY% -m pip install -r requirements.txt --disable-pip-version-check --no-warn-script-location
echo.

echo ============================================================
echo Starting server. Browser will open in a few seconds.
echo Press Ctrl+C in this window to stop the server.
echo ============================================================
echo.

REM Open browser after a short delay so the server has time to bind
start "" /b cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:5000"

%PY% app.py

echo.
echo ============================================================
echo Server stopped. Press any key to close this window.
echo ============================================================
pause >nul
