@echo off
setlocal
cd /d "%~dp0"

if not exist "tools\cloudflared.exe" (
  echo cloudflared.exe not found in tools folder.
  echo Run the setup step first or ask Codex to download it again.
  pause
  exit /b 1
)

start "Bymom Report Server" cmd /k "cd /d %~dp0 && C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe server.mjs"
timeout /t 3 >nul
start "Bymom Public Tunnel" cmd /k "cd /d %~dp0 && tools\cloudflared.exe tunnel --url http://localhost:3030 --no-autoupdate --logfile tunnel-live.log"
