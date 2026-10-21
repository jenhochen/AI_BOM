@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
set "PY=python"
if exist "%~dp0python_home.txt" set /p PY=<"%~dp0python_home.txt"
set "PYW=pythonw"
if exist "%~dp0python_home.txt" set "PYW=%PY:python.exe=pythonw.exe%"
title BOM Matcher v0.10 - CLI
echo.
echo   BOM Matcher v0.10  -  CLI
echo   Folder: %~dp0
echo.
"%PY%" "%~dp0bom_core_0.10.py" %*
if errorlevel 1 (
  echo.
  echo   [ERROR] Failed. Run install_requirements.bat first.
)
echo.
pause
