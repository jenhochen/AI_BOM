"""
產生 BOM Matcher v0.07 的 .bat 啟動檔。

三個硬性規則（這次閃退就是全部違反）：
  1. CRLF 換行  —— cmd.exe 用位元組位移讀批次檔，LF-only 會從行中間開始執行
  2. 純 ASCII   —— .bat 內不放中文，避免 cmd 用哪個 codepage 讀檔的不確定性
  3. 絕對路徑   —— 用 %~dp0 組出完整路徑，就算 cd 失效也能跑
"""
from pathlib import Path

DST = Path(r"F:\AI_BOM")

HEAD = (
    "@echo off\r\n"
    "chcp 65001 >nul 2>&1\r\n"
    "set PYTHONIOENCODING=utf-8\r\n"
    "cd /d \"%~dp0\"\r\n"
)

# Streamlit 首次執行會互動詢問 email 並停在 "Email:" 等輸入，
# 在雙擊 .bat 的情境下看起來就像卡死。預先寫一個空的 credentials.toml
# 跳過詢問（這正是 Streamlit 自己在使用者按 Enter 後會建立的檔案）。
SKIP_STREAMLIT_WELCOME = (
    "set \"STCRED=%USERPROFILE%\\.streamlit\"\r\n"
    "if not exist \"%STCRED%\" mkdir \"%STCRED%\" >nul 2>&1\r\n"
    "if not exist \"%STCRED%\\credentials.toml\" (\r\n"
    "  echo [general]>\"%STCRED%\\credentials.toml\"\r\n"
    "  echo email = \"\">>\"%STCRED%\\credentials.toml\"\r\n"
    ")\r\n"
)

BATS = {}

BATS["run_CLI.bat"] = HEAD + (
    "title BOM Matcher v0.07 - CLI\r\n"
    "echo.\r\n"
    "echo   BOM Matcher v0.07  -  CLI\r\n"
    "echo   Folder: %~dp0\r\n"
    "echo.\r\n"
    "python \"%~dp0bom_core_0.07.py\" %*\r\n"
    "if errorlevel 1 (\r\n"
    "  echo.\r\n"
    "  echo   [ERROR] Failed. Run install_requirements.bat first.\r\n"
    ")\r\n"
    "echo.\r\n"
    "pause\r\n"
)

BATS["run_C_streamlit.bat"] = HEAD + SKIP_STREAMLIT_WELCOME + (
    "title BOM Matcher v0.07 - C / Streamlit\r\n"
    "echo.\r\n"
    "echo   BOM Matcher v0.07  -  C  (Streamlit web UI)\r\n"
    "echo   Folder: %~dp0\r\n"
    "echo   URL   : http://localhost:8601\r\n"
    "echo.\r\n"
    "echo   Keep this window open while using the app.\r\n"
    "echo   Close it to stop the server.\r\n"
    "echo.\r\n"
    "python -m streamlit run \"%~dp0bom_matcher_C_0.07.py\""
    " --server.port 8601 --server.headless false\r\n"
    "if errorlevel 1 (\r\n"
    "  echo.\r\n"
    "  echo   [ERROR] Failed to start.\r\n"
    "  echo          - Run install_requirements.bat first, or\r\n"
    "  echo          - Port 8601 may be in use; try:\r\n"
    "  echo            python -m streamlit run \"%~dp0bom_matcher_C_0.07.py\""
    " --server.port 8602\r\n"
    ")\r\n"
    "echo.\r\n"
    "pause\r\n"
)

BATS["run_D_webview.bat"] = HEAD + (
    "title BOM Matcher v0.07 - D / pywebview\r\n"
    "start \"\" pythonw \"%~dp0bom_matcher_D_0.07.py\"\r\n"
)

BATS["run_D_webview_debug.bat"] = HEAD + (
    "title BOM Matcher v0.07 - D / pywebview (debug)\r\n"
    "set PYWEBVIEW_LOG=debug\r\n"
    "echo.\r\n"
    "echo   BOM Matcher v0.07  -  D  (debug mode, console kept open)\r\n"
    "echo.\r\n"
    "python \"%~dp0bom_matcher_D_0.07.py\"\r\n"
    "echo.\r\n"
    "echo   Window closed. If there is an error above, please screenshot it.\r\n"
    "pause\r\n"
)

BATS["install_requirements.bat"] = HEAD + (
    "title BOM Matcher v0.07 - Install requirements\r\n"
    "echo.\r\n"
    "echo ================================================\r\n"
    "echo   BOM Matcher v0.07   Install requirements\r\n"
    "echo   Folder: %~dp0\r\n"
    "echo ================================================\r\n"
    "echo.\r\n"
    "python --version\r\n"
    "if errorlevel 1 (\r\n"
    "  echo.\r\n"
    "  echo   [ERROR] Python not found.\r\n"
    "  echo           Install Python 3.10+ and tick\r\n"
    "  echo           \"Add python.exe to PATH\".\r\n"
    "  echo           https://www.python.org/downloads/\r\n"
    "  echo.\r\n"
    "  pause\r\n"
    "  exit /b 1\r\n"
    ")\r\n"
    "echo.\r\n"
    "echo   Installing...\r\n"
    "echo.\r\n"
    "python -m pip install --upgrade pip\r\n"
    "python -m pip install -r \"%~dp0requirements.txt\"\r\n"
) + SKIP_STREAMLIT_WELCOME + (
    "echo.\r\n"
    "echo ================================================\r\n"
    "echo   Check result:\r\n"
    "echo ================================================\r\n"
    # 版號一律走 importlib.metadata：pywebview 6.x 沒有 webview.__version__，
    # 用 webview.__version__ 會 AttributeError，害 || 分支誤報 NOT installed。
    "python -c \"import pandas,openpyxl,PIL,importlib.metadata as m;"
    "print('  [OK] core      pandas',m.version('pandas'),'/ openpyxl',"
    "m.version('openpyxl'),'/ pillow',m.version('pillow'))\""
    " 2>nul || echo   [--] core      pandas/openpyxl/pillow NOT installed\r\n"
    "python -c \"import streamlit,importlib.metadata as m;"
    "print('  [OK] plan C    streamlit',m.version('streamlit'))\""
    " 2>nul || echo   [--] plan C    streamlit NOT installed\r\n"
    "python -c \"import webview,importlib.metadata as m;"
    "print('  [OK] plan D    pywebview',m.version('pywebview'))\""
    " 2>nul || echo   [--] plan D    pywebview NOT installed\r\n"
    "echo.\r\n"
    "echo   Next:\r\n"
    "echo     run_CLI.bat           check the core first\r\n"
    "echo     run_C_streamlit.bat   plan C  (browser)\r\n"
    "echo     run_D_webview.bat     plan D  (native window)\r\n"
    "echo.\r\n"
    "pause\r\n"
)

for name, body in BATS.items():
    non_ascii = [c for c in body if ord(c) > 127]
    assert not non_ascii, f"{name} 含非 ASCII 字元: {non_ascii}"
    assert "\n" not in body.replace("\r\n", ""), f"{name} 有落單的 LF"
    p = DST / name
    p.write_bytes(body.encode("ascii"))
    b = p.read_bytes()
    crlf = b.count(b"\r\n")
    lone_lf = sum(1 for i, ch in enumerate(b)
                  if ch == 0x0A and (i == 0 or b[i - 1] != 0x0D))
    print(f"  {name:<30} {len(b):>5} bytes  CRLF={crlf:<3} lone-LF={lone_lf}  "
          f"ASCII={'OK' if max(b) < 128 else 'FAIL'}")
print("\n完成，共 %d 個 .bat" % len(BATS))
