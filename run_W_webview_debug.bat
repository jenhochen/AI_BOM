@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
set "PY=python"
if exist "%~dp0python_home.txt" set /p PY=<"%~dp0python_home.txt"
set "PYW=pythonw"
if exist "%~dp0python_home.txt" set "PYW=%PY:python.exe=pythonw.exe%"
title BOM Matcher v0.10 - W / pywebview (debug)
set PYWEBVIEW_LOG=debug
echo.
echo   BOM Matcher v0.10  -  D  (debug mode, console kept open)
echo.
"%PY%" "%~dp0bom_matcher_W_0.10.py"
echo.
echo   Window closed. If there is an error above, please screenshot it.
pause
