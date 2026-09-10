@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\vi\pptmaster.ps1" -Action check
set "RC=%ERRORLEVEL%"
echo.
pause
exit /b %RC%
