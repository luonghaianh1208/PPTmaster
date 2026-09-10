@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\vi\pptmaster.ps1" -Action update
set "RC=%ERRORLEVEL%"
echo.
pause
exit /b %RC%
