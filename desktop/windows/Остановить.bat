@echo off
chcp 65001 >nul
if not exist "%~dp0files\stop.ps1" (
  echo.
  echo   Не найдена папка files рядом с этим файлом. Сначала распакуйте архив целиком.
  echo.
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0files\stop.ps1"
