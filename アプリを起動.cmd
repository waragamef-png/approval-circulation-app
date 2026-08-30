@echo off
setlocal

set "PROJECT_DIR=%~dp0."
set "APP_NODE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"

if not exist "%APP_NODE%" set "APP_NODE=node.exe"

powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%PROJECT_DIR%\scripts\start-local-app.ps1" -ProjectPath "%PROJECT_DIR%" -NodePath "%APP_NODE%"

endlocal
