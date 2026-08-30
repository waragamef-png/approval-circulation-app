@echo off
setlocal

set "PROJECT_DIR=%~dp0."
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%PROJECT_DIR%\scripts\stop-local-app.ps1" -ProjectPath "%PROJECT_DIR%"

endlocal
