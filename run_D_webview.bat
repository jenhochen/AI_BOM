@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
title BOM Matcher v0.07 - D / pywebview
start "" pythonw "%~dp0bom_matcher_D_0.07.py"
