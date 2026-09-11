@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
set "STCRED=%USERPROFILE%\.streamlit"
if not exist "%STCRED%" mkdir "%STCRED%" >nul 2>&1
if not exist "%STCRED%\credentials.toml" (
  echo [general]>"%STCRED%\credentials.toml"
  echo email = "">>"%STCRED%\credentials.toml"
)
title BOM Matcher v0.07 - C / Streamlit
echo.
echo   BOM Matcher v0.07  -  C  (Streamlit web UI)
echo   Folder: %~dp0
echo   URL   : http://localhost:8601
echo.
echo   Keep this window open while using the app.
echo   Close it to stop the server.
echo.
python -m streamlit run "%~dp0bom_matcher_C_0.07.py" --server.port 8601 --server.headless false
if errorlevel 1 (
  echo.
  echo   [ERROR] Failed to start.
  echo          - Run install_requirements.bat first, or
  echo          - Port 8601 may be in use; try:
  echo            python -m streamlit run "%~dp0bom_matcher_C_0.07.py" --server.port 8602
)
echo.
pause
