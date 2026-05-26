@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  Laser SFG Dashboard - Local HTTP Server
echo ============================================================
echo.
echo Opening: http://localhost:8080/index.html
echo Press Ctrl+C in this window to stop the server.
echo.

start "" "http://localhost:8080/index.html"

node serve_index.js
