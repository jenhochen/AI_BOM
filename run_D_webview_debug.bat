@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
title BOM Matcher v0.07 - D / pywebview (debug)
set PYWEBVIEW_LOG=debug
echo.
echo   BOM Matcher v0.07  -  D  (debug mode, console kept open)
echo.
python "%~dp0bom_matcher_D_0.07.py"
echo.
echo   Window closed. If there is an error above, please screenshot it.
pause
