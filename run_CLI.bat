@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
title BOM Matcher v0.07 - CLI
echo.
echo   BOM Matcher v0.07  -  CLI
echo   Folder: %~dp0
echo.
python "%~dp0bom_core_0.07.py" %*
if errorlevel 1 (
  echo.
  echo   [ERROR] Failed. Run install_requirements.bat first.
)
echo.
pause
