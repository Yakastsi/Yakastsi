@echo off
chcp 65001 >nul
title PC Serial Scanner — Yakastsi / rgb.by
cd /d "%~dp0"

echo.
echo  ============================================================
echo   PC Serial Scanner
echo   Yakastsi — rgb.by
echo  ============================================================
echo.
echo  Сбор серийных номеров комплектующих...
echo  Не закрывайте это окно.
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PC-Serial-Scanner.ps1"
set ERR=%ERRORLEVEL%

echo.
if %ERR% NEQ 0 (
  echo  Ошибка выполнения. Код: %ERR%
  echo  Попробуйте запустить от имени администратора.
) else (
  echo  Готово. Отчёты сохранены в папке Reports.
)
echo.
pause
exit /b %ERR%
