@echo off
setlocal
title PC Serial Scanner - Yakastsi / rgb.by
cd /d "%~dp0"

echo.
echo  ============================================================
echo   PC Serial Scanner
echo   Yakastsi - rgb.by
echo  ============================================================
echo.
echo  Collecting hardware serial numbers...
echo  Please wait. Do not close this window.
echo.

where powershell.exe >nul 2>&1
if errorlevel 1 (
  echo  ERROR: powershell.exe not found.
  echo  This tool needs Windows PowerShell.
  echo.
  pause
  exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0PC-Serial-Scanner.ps1"
set ERR=%ERRORLEVEL%

echo.
if not "%ERR%"=="0" (
  echo  ERROR. Exit code: %ERR%
  echo  Try: right-click this .bat -^> Run as administrator.
) else (
  echo  Done. Reports saved in the Reports folder.
)
echo.
pause
endlocal & exit /b %ERR%
