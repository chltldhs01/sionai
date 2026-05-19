@echo off
setlocal
cd /d "%~dp0"
echo Starting Bymom weekly report generator...
"C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" server.mjs
