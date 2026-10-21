@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-ONGEN-Excel.ps1"
if errorlevel 1 (
  echo.
  echo ONGEN build failed. Review ONGEN-build.log in the newest output\build-* folder.
  pause
  exit /b 1
)
echo.
echo ONGEN Excel suite is ready in the newest output\build-* folder.
pause
