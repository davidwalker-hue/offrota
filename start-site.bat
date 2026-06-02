@echo off
setlocal
cd /d "%~dp0"
if "%ADMIN_PASSWORD%"=="" set "ADMIN_PASSWORD=ChangeMe2026!"
"C:\Users\DavidWalker\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" server.js
