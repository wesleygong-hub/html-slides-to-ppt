@echo off
setlocal
cd /d "%~dp0"
call npx playwright install chromium
if errorlevel 1 (
  echo [ERROR] Chromium installation failed.
  pause
  exit /b 1
)
echo Chromium installation complete.
pause
endlocal
