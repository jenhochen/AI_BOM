@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
set "PY=python"
if exist "%~dp0python_home.txt" set /p PY=<"%~dp0python_home.txt"
set "PYW=pythonw"
if exist "%~dp0python_home.txt" set "PYW=%PY:python.exe=pythonw.exe%"
title BOM Matcher v0.10 - Install requirements
echo.
echo ================================================
echo   BOM Matcher v0.10   Install requirements
echo   Folder: %~dp0
echo ================================================
echo.
"%PY%" --version
if errorlevel 1 (
  echo.
  echo   [ERROR] Python not found.
  echo           Install Python 3.10+ and tick
  echo           "Add python.exe to PATH".
  echo           https://www.python.org/downloads/
  echo.
  pause
  exit /b 1
)
echo.
echo   Installing...
echo.
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install -r "%~dp0requirements.txt"
set "STCRED=%USERPROFILE%\.streamlit"
if not exist "%STCRED%" mkdir "%STCRED%" >nul 2>&1
if not exist "%STCRED%\credentials.toml" (
  echo [general]>"%STCRED%\credentials.toml"
  echo email = "">>"%STCRED%\credentials.toml"
)
echo.
echo ================================================
echo   Check result:
echo ================================================
"%PY%" -c "import pandas,openpyxl,PIL,importlib.metadata as m;print('  [OK] core      pandas',m.version('pandas'),'/ openpyxl',m.version('openpyxl'),'/ pillow',m.version('pillow'))" 2>nul || echo   [--] core      pandas/openpyxl/pillow NOT installed
"%PY%" -c "import streamlit,importlib.metadata as m;print('  [OK] plan C    streamlit',m.version('streamlit'))" 2>nul || echo   [--] plan C    streamlit NOT installed
"%PY%" -c "import webview,importlib.metadata as m;print('  [OK] plan D    pywebview',m.version('pywebview'))" 2>nul || echo   [--] plan D    pywebview NOT installed
echo.
echo   Next:
echo     run_CLI.bat           check the core first
echo     run_C_streamlit.bat   plan C  (browser)
echo     run_D_webview.bat     plan D  (native window)
echo.
pause
