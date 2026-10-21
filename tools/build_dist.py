#!/usr/bin/env python3
"""
tools/build_dist.py  —  產出 BOM Matcher 的兩種交付包

    dist/BOM_Matcher_v<VER>_完整離線安裝版/   + .zip   目標機器免上網
    dist/BOM_Matcher_v<VER>_精簡版/           + .zip   目標機器要上網

規格對齊 F:\\AI_ini_allin1 的 ADATA SSD Firmware Tool：
  完整離線安裝版 = 精簡版 + python\\（Python 安裝檔）+ wheels\\（離線套件）
  兩者都有 安裝.bat + 安裝說明.txt

用法
----
    python tools/build_dist.py              產出兩包（需要 dist/wheels_cache）
    python tools/build_dist.py --slim-only  只產出精簡版
    python tools/build_dist.py --refresh-wheels
                                            重新下載 wheels_cache 再打包

.bat 的三條硬規則（v0.07 交付事故的教訓，見 HANDOVER）
  1. CRLF 換行
  2. chcp 65001 必須在任何中文出現之前執行，且檔案存成 UTF-8 無 BOM
  3. 用絕對路徑 %~dp0
"""

import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VER = "0.10"
PY_VER = "3.14.4"
PY_EXE = f"python-{PY_VER}-amd64.exe"
PY_URL = f"https://www.python.org/ftp/python/{PY_VER}/{PY_EXE}"

# 從 Firmware Tool 的離線包借用同一份 Python 安裝檔（兩個專案版本一致）
FW_PY = Path(r"F:\AI_ini_allin1\dist"
             r"\ADATA_SSD_Firmware_Tool_v1.44_完整離線安裝版\python") / PY_EXE

DIST = ROOT / "dist"
WHEELS_CACHE = DIST / "wheels_cache"

PROGRAM_FILES = [
    f"bom_core_{VER}.py",
    f"bom_matcher_C_{VER}.py",
    f"bom_matcher_D_{VER}.py",
    "run_CLI.bat",
    "run_C_streamlit.bat",
    "run_D_webview.bat",
    "run_D_webview_debug.bat",
    "requirements.txt",
    f"changelog_BOM_matcher_v{VER}.txt",
    "HANDOVER_bom_matcher.md",
    "pcb.jpg",
    "adata.png",          # 品牌 LOGO（C/D 兩案的左上角）
]
SOP = [f"BOM_Matcher_v{VER}_操作SOP.pptx",
       f"BOM_Matcher_v{VER}_操作SOP.html"]
SAMPLE = {
    "TXT": ["SM8008_M2_D4X2_4G178_C261C_V1C_20261012.BOM.txt",
            "S-M2P42280SMH086-20260812_1.BOM"],
    # v0.09：Location 對照表一併附上，讓交付包可以直接示範雙面板模式
    "ExcelDB": ["BOM_20261014.xlsx", "SM8008_M2_Location_TOP_BOT.xlsx"],
}
EXTRA_KNOWLEDGE = "Layout and BOM Knowledge_20260624.xlsx"   # 放進 ExcelDB\


# ════════════════════════════════════════════════════════════════════════════
# .bat / .txt 產生（CRLF，UTF-8 無 BOM，chcp 先行）
# ════════════════════════════════════════════════════════════════════════════

HEAD = (
    "@echo off\r\n"
    "chcp 65001 >nul 2>&1\r\n"
    "set PYTHONIOENCODING=utf-8\r\n"
    "cd /d \"%~dp0\"\r\n"
    "setlocal\r\n"
)

CHECK_TAIL = (
    "echo.\r\n"
    "echo ============================================================\r\n"
    "echo   檢查結果\r\n"
    "echo ------------------------------------------------------------\r\n"
    "\"%PYEXE%\" -c \"import pandas,openpyxl,PIL,importlib.metadata as m;"
    "print('  [OK] 核心     pandas',m.version('pandas'),'/ openpyxl',"
    "m.version('openpyxl'),'/ pillow',m.version('pillow'))\""
    " 2>nul || echo   [--] 核心     未安裝\r\n"
    "\"%PYEXE%\" -c \"import streamlit,importlib.metadata as m;"
    "print('  [OK] C 案     streamlit',m.version('streamlit'))\""
    " 2>nul || echo   [--] C 案     streamlit 未安裝\r\n"
    "\"%PYEXE%\" -c \"import webview,importlib.metadata as m;"
    "print('  [OK] D 案     pywebview',m.version('pywebview'))\""
    " 2>nul || echo   [--] D 案     pywebview 未安裝\r\n"
    "echo ============================================================\r\n"
    "echo.\r\n"
    "echo   接著可以執行：\r\n"
    "echo     run_CLI.bat           先用命令列確認核心正常\r\n"
    "echo     run_C_streamlit.bat   C 案（瀏覽器介面）\r\n"
    "echo     run_D_webview.bat     D 案（原生視窗）\r\n"
    "echo.\r\n"
    "pause\r\n"
)

STREAMLIT_CRED = (
    "set \"STCRED=%USERPROFILE%\\.streamlit\"\r\n"
    "if not exist \"%STCRED%\" mkdir \"%STCRED%\" >nul 2>&1\r\n"
    "if not exist \"%STCRED%\\credentials.toml\" (\r\n"
    "  echo [general]>\"%STCRED%\\credentials.toml\"\r\n"
    "  echo email = \"\">>\"%STCRED%\\credentials.toml\"\r\n"
    ")\r\n"
)


def install_bat(offline: bool) -> str:
    """
    v0.09 修訂：離線包必須檢查 Python 版本，不能只看「有沒有 python」。

    新機器實測失敗（2026-10-19）：目標機器裝的是 Microsoft Store 的
    Python 3.12，舊版 安裝.bat 看到 where python 有結果就跳過安裝，直接拿
    為 3.14 編譯的 wheels 去裝 ——
        ERROR: Could not find a version that satisfies the requirement
               cffi>=1.17 (from clr-loader) (from versions: none)
    47 個 wheels 裡有 10 個是 cp314 專用的二進位檔，換個 Python 版本就廢了。

    改法：
      1. 比對 PATH 上的 python 版本是否等於 wheels 的目標版本
      2. 不等於就裝內附的，並用「絕對路徑」接著做，不依賴 PATH
         （Store 版的 app execution alias 常常排在 PATH 前面）
      3. 把實際使用的直譯器寫進 python_home.txt，run_*.bat 會讀它
    """
    pydir = "Python" + PY_VER.rsplit(".", 1)[0].replace(".", "")   # 3.14.4 -> Python314
    want = PY_VER.rsplit(".", 1)[0]                                # 3.14
    kind = "完整離線安裝版．不需要網路" if offline else "精簡版．需要網路"
    VP = ('python -c "import sys;print(str(sys.version_info[0])'
          '+chr(46)+str(sys.version_info[1]))" 2^>nul')
    EP = 'python -c "import sys;print(sys.executable)" 2^>nul'
    NL = "\r\n"
    L = []

    def w(s):
        L.append(s + NL)

    w('title BOM Matcher v%s - 安裝' % VER)
    w('echo ============================================================')
    w('echo   BOM Matcher v%s  安裝（%s）' % (VER, kind))
    w('echo ============================================================')
    w('echo.')
    w('set "PYEXE="')
    w('set "FOUND="')
    w('for /f "delims=" %%%%v in (\'%s\') do set "FOUND=%%%%v"' % VP)

    if offline:
        w('set "BUNDLED=%%LOCALAPPDATA%%\\Programs\\Python\\%s\\python.exe"' % pydir)
        w('echo [1/3] 檢查 Python（這個包的套件是為 %s 編譯的）' % want)
        w('if defined FOUND echo       目前 PATH 上的版本：%FOUND%')
        w('if not defined FOUND echo       PATH 上找不到 Python')
        w('if "%%FOUND%%"=="%s" (' % want)
        w('    for /f "delims=" %%%%p in (\'%s\') do set "PYEXE=%%%%p"' % EP)
        w('    echo       版本相符，不用另外安裝。')
        w('    goto PKGS')
        w(')')
        w('if exist "%BUNDLED%" (')
        w('    echo       已經裝過內附的 Python %s，直接使用。' % want)
        w('    set "PYEXE=%BUNDLED%"')
        w('    goto PKGS')
        w(')')
        w('echo [2/3] 安裝內附的 Python %s ...' % PY_VER)
        w('if not exist "python\\%s" (' % PY_EXE)
        w('    echo.')
        w('    echo [錯誤] 找不到 python\\%s，這個包不完整。' % PY_EXE)
        w('    pause')
        w('    exit /b 1')
        w(')')
        w('echo       安裝中，請稍候（約 1 分鐘，不會動到你原本的 Python）...')
        w('"python\\%s" /passive InstallAllUsers=0 PrependPath=1 Include_test=0' % PY_EXE)
        w('if not exist "%BUNDLED%" (')
        w('    echo.')
        w('    echo [錯誤] Python 安裝完成但找不到：')
        w('    echo        %BUNDLED%')
        w('    echo        請把這個訊息回報給提供者。')
        w('    pause')
        w('    exit /b 1')
        w(')')
        w('set "PYEXE=%BUNDLED%"')
    else:
        w('echo [1/3] 檢查 Python（需要 3.10 以上）')
        # 精簡版從 PyPI 抓，不會有離線包那種「wheels 綁死版本」的問題，
        # 但也因此什麼版本都收 —— 太舊的 Python 裝得起來卻跑不動，
        # 所以這裡擋一道最低版本。
        w('set "TOOOLD="')
        w('if defined FOUND python -c '
          '"import sys;sys.exit(0 if sys.version_info>=(3,10) else 1)" '
          '2>nul || set "TOOOLD=1"')
        w('if defined TOOOLD (')
        w('    echo.')
        w('    echo [錯誤] 偵測到的 Python 版本是 %FOUND%，太舊了。')
        w('    echo        本程式需要 3.10 以上。請先移除或改用新版 Python，')
        w('    echo        或改用「完整離線安裝版」交付包（自帶 %s）。' % PY_VER)
        w('    pause')
        w('    exit /b 1')
        w(')')
        w('if defined FOUND (')
        w('    for /f "delims=" %%%%p in (\'%s\') do set "PYEXE=%%%%p"' % EP)
        w('    echo       目前版本：%FOUND%')
        w('    goto PKGS')
        w(')')
        w('echo [2/3] 下載並安裝 Python %s ...' % PY_VER)
        w('echo       來源：%s' % PY_URL)
        w('if not exist "%%TEMP%%\\%s" (' % PY_EXE)
        w('    curl -L --fail -o "%%TEMP%%\\%s" "%s"' % (PY_EXE, PY_URL))
        w('    if errorlevel 1 (')
        w('        echo.')
        w('        echo [錯誤] 下載失敗。請確認這台機器連得上網際網路，')
        w('        echo        或改用「完整離線安裝版」交付包。')
        w('        pause')
        w('        exit /b 1')
        w('    )')
        w(')')
        w('"%%TEMP%%\\%s" /passive InstallAllUsers=0 PrependPath=1 Include_test=0' % PY_EXE)
        w('set "PYEXE=%%LOCALAPPDATA%%\\Programs\\Python\\%s\\python.exe"' % pydir)
        w('if not exist "%PYEXE%" set "PYEXE=python"')

    w(':PKGS')
    w('echo.')
    w('echo       使用的直譯器：%PYEXE%')
    w('"%PYEXE%" --version')
    w('if errorlevel 1 (')
    w('    echo.')
    w('    echo [錯誤] 這個 Python 不能執行，請回報給提供者。')
    w('    pause')
    w('    exit /b 1')
    w(')')
    w('echo %PYEXE%>"%~dp0python_home.txt"')
    w('echo.')
    if offline:
        w('echo [3/3] 安裝相依套件（從 wheels 資料夾，不連網）...')
        w('if not exist "wheels" (')
        w('    echo.')
        w('    echo [錯誤] 找不到 wheels 資料夾，這個包不完整。')
        w('    pause')
        w('    exit /b 1')
        w(')')
        w('"%PYEXE%" -m pip install --no-index --find-links wheels -r requirements.txt')
    else:
        w('echo [3/3] 下載並安裝相依套件...')
        w('"%PYEXE%" -m pip install -r requirements.txt')
    w('if errorlevel 1 (')
    w('    echo.')
    if offline:
        w('    echo [錯誤] 套件安裝失敗。請把上面的訊息回報給提供者。')
        w('    echo        （這個包的 wheels 是為 Python %s 編譯的）' % want)
    else:
        w('    echo [錯誤] 套件安裝失敗。連不上 PyPI 的話請改用「完整離線安裝版」。')
    w('    pause')
    w('    exit /b 1')
    w(')')

    return HEAD + "".join(L) + STREAMLIT_CRED + CHECK_TAIL


def readme_txt(offline: bool) -> str:
    kind = "【完整離線安裝版】安裝時不需要網路" if offline else "【精簡版】安裝時需要網路"
    net = ("""  ✔ 這個包自帶 Python 安裝檔與全部套件，安裝過程完全不需要網路。
     適合產線機器或連不上外網的環境。"""
           if offline else
           """  ⚠️ 安裝時需要連得上網際網路（從 python.org 抓 Python、
     從 PyPI 抓套件）。產線機器連不上外網請改用完整離線安裝版。""")
    size = "約 120 MB" if offline else "約 250 KB"
    return f"""BOM Matcher  v{VER}
{kind}
交付包大小：{size}
================================================================

Altium Text BOM + Excel 零件資料庫  →  ADATA E-BOM 格式 Excel

本包含兩套介面（C 案 / D 案）與一套 CLI，共用同一份核心邏輯，
產出的 Excel 完全相同。附可直接執行的範例資料。


■ 安裝

  1. 把整個資料夾複製到目標機器（路徑不限，程式以自己所在目錄為準）
  2. 執行「安裝.bat」
  3. 如果它說「Python 剛裝好，請重開視窗」，就關掉視窗、
     再執行一次「安裝.bat」把套件裝完

{net}


■ 啟動（三選一）

  run_CLI.bat            命令列，不需要 streamlit / pywebview
                         建議第一次先跑這個確認核心正常

  run_C_streamlit.bat    C 案．瀏覽器介面
                         http://localhost:8601
                         使用期間請勿關掉那個黑色主控台視窗
                         ★ 不要直接雙擊 bom_matcher_C_{VER}.py，不會開網頁

  run_D_webview.bat      D 案．原生視窗，不需要瀏覽器
  run_D_webview_debug.bat  同上，但保留主控台（出問題時用這支）

  兩個介面的標題列會標出是哪支 .bat 啟動的，同時開也分得出來：
      BOM Matcher v{VER}  [run_C_streamlit.bat]
      BOM Matcher v{VER}  [run_D_webview.bat]


■ 操作流程（C / D 兩案相同）

  1. 檔案設定      確認自動偵測的四個檔案
                   不對就用「瀏覽…」或「📂 目錄瀏覽器」自己選
  2. 執行比對
  3. 比對結果      檢查表格，可直接改 P/N / RC/EC / Description /
                   Qty / Location
  4. Knowledge     逐筆確認，不適用的按「忽略此筆」
  5. 表頭資訊      填 PCB板號 / 供應商 / IC Type 等（可留空）
  6. 產出 E-BOM    輸出到 output\\


■ 資料夾規則

  程式會掃描自己所在的資料夾與下列子資料夾：

      TXT\\           Altium Text BOM（*.txt 或 *.BOM）
      ExcelDB\\       零件資料庫 + Knowledge（*.xlsx）
      screenshot\\    PCB 圖片（*.jpg / *.png，pcb.* 優先）
      output\\        產出的 E-BOM（不存在則輸出到程式資料夾）

  ★ 自動偵測是「看內容」不是「看檔名」：
      Text BOM     要有一列是 TAB 分隔的表頭，欄位含
                   Item / Quantity / Reference
      零件資料庫   第 1 列要有 P/N 與 Description
      Knowledge    第 1 列要有 品名 與 包裝

  ★ 每一類候選都依「最後修改時間」排序，最新的排第一個、預設選中。


■ 本包附的範例資料

  TXT\\SM8008_M2_D4X2_4G178_C261C_V1C_20261012.BOM.txt   112 項（預設）
  TXT\\S-M2P42280SMH086-20260812_1.BOM                    44 項
  ExcelDB\\BOM_20261014.xlsx                              零件庫 153 筆
  ExcelDB\\Layout and BOM Knowledge_20260624.xlsx         Knowledge 6 筆
  pcb.jpg                                                PCB 圖片

  跑 run_CLI.bat 應該看到：

      解析 Text BOM: 112 項
      讀取零件資料庫: 153 筆
      載入 Knowledge: 6 筆
      載入 Location 對照表: 381 筆（Top 167 / Bot 214）→ 雙面板模式
      比對完成: 命中 55 / 未命中 58 / 替代料群組 29 / Knowledge 命中 4 列
      113 項（含PCB）→ 151 列（替代料拆出 38 列）  |  Total Qty = 384

  C 案、D 案跑同一份資料，數字必須完全一樣。


■ v{VER} 的重點

  • Location-Top / Location-Bottom 兩欄的每個 Location 加上 (T) / (B)
        (T)BC26,(T)C76,(T)C81        (B)BC28,(B)C55,(B)C56
  • R8 電容降壓選料：容值與封裝一致、但 DB 耐壓低於原始設計時可以選，
    Description 自動加註「(原始設計為10V, 請再確認)」，畫面標 🔴、Excel 標紫底。
    SM8008 命中 45 → 55（+10 項電容），Total Qty 不變。
    只放寬電容耐壓 —— 電流與誤差不放寬，那是功能性錯誤。
  • C/D 畫面上的 Location 欄也帶 (T)/(B)，與輸出 Excel 一致
  • 輸出路徑已經有同名檔時會先問，不會無聲蓋掉
  • 左上角換成 ADATA LOGO；C 案右上角的 Deploy 按鈕已拿掉
  • 用語統一：位號 → Location、面別 → Top/Bot、chip → 篩選鈕
  • 操作 SOP 改成標註式版面：截圖上標 ①②③，右側說明同號對應
  • 雙面板：給 Location 對照表就多一欄 Location-Bottom Layer；
    PCB 圖片放兩張（檔名含 top / bot）會自動切成上下兩半
  • Location 對不上時三層警告：日誌、Excel 淡紅底、紅字寫出是哪幾個

  完整清單見 changelog_BOM_matcher_v{VER}.txt


■ 雙面板怎麼用（選用）

  1. Location 對照表放進 ExcelDB\\（或程式資料夾）
     第一欄＝Location，第二欄＝Top/Bot（TOP / BOT）
     xlsx 為主要格式，txt 亦可
  2. PCB 圖片放兩張。檔名含 top / bot 會自動配對，手動指定任何兩個
     檔案也可以，不受檔名限制。Top 與 Bot 選到同一張時會提醒確認
     例：screenshot\\pcb_top.jpg、screenshot\\pcb_bot.jpg
  3. 執行比對 —— 日誌出現「→ 雙面板模式」就對了

  沒有 Location 對照表就是單面板，輸出與 v0.08 完全相同。

  ★ 對照表與 BOM 必須是同一個改版。本包附的 SM8008 範例刻意保留了
    6 個對不上的 Location，可以看到警告長什麼樣子。


■ 常見問題

  Q. 雙擊 .bat 一閃就關
     A. 沒裝 Python 或沒加入 PATH。先跑「安裝.bat」看訊息。

  Q. C 案瀏覽器打不開
     A. 8601 埠被占用。改執行：
        python -m streamlit run bom_matcher_C_{VER}.py --server.port 8602

  Q. D 案視窗開不起來
     A. 用 run_D_webview_debug.bat 看訊息。
        多半是缺 Microsoft Edge WebView2 Runtime（Win11 內建）。
        https://developer.microsoft.com/microsoft-edge/webview2/

  Q. 說找不到零件資料庫
     A. 檢查 xlsx 第 1 列是不是 P/N 與 Description 這兩個字。

================================================================
"""


# ════════════════════════════════════════════════════════════════════════════
# 打包
# ════════════════════════════════════════════════════════════════════════════

def write_crlf(path: Path, text: str):
    """CRLF + UTF-8 無 BOM。不可用 Write 工具或一般 open() 的預設值。"""
    path.write_bytes(text.replace("\r\n", "\n").replace("\n", "\r\n")
                     .encode("utf-8"))


def build(offline: bool) -> Path:
    name = f"BOM_Matcher_v{VER}_" + ("完整離線安裝版" if offline else "精簡版")
    dst = DIST / name
    if dst.exists():
        shutil.rmtree(dst)
    for sub in (".streamlit", "TXT", "ExcelDB", "output", "screenshot",
                "tools", "docs"):
        (dst / sub).mkdir(parents=True, exist_ok=True)

    for f in PROGRAM_FILES:
        shutil.copy2(ROOT / f, dst / f)
    shutil.copy2(ROOT / ".streamlit" / "config.toml", dst / ".streamlit")
    shutil.copy2(ROOT / "tools" / "gen_bats.py", dst / "tools")
    shutil.copy2(ROOT / "tools" / "build_dist.py", dst / "tools")
    for s in SOP:
        f = ROOT / "docs" / s
        if f.exists():
            shutil.copy2(f, dst / "docs" / s)

    for sub, files in SAMPLE.items():
        for f in files:
            src = ROOT / sub / f
            if src.exists():
                shutil.copy2(src, dst / sub / f)
    kn = ROOT / EXTRA_KNOWLEDGE
    if kn.exists():
        shutil.copy2(kn, dst / "ExcelDB" / EXTRA_KNOWLEDGE)

    (dst / "output" / "_readme.txt").write_text(
        "產出的 E-BOM 會放在這個資料夾。這個說明檔可以刪掉。\n", encoding="utf-8")
    (dst / "screenshot" / "_readme.txt").write_text(
        "PCB 圖片（jpg/png）可以放這裡，或直接放程式資料夾。\n", encoding="utf-8")

    write_crlf(dst / "安裝.bat", install_bat(offline))
    write_crlf(dst / "安裝說明.txt", readme_txt(offline))

    if offline:
        (dst / "python").mkdir(exist_ok=True)
        if not FW_PY.exists():
            raise SystemExit(f"找不到 Python 安裝檔：{FW_PY}")
        shutil.copy2(FW_PY, dst / "python" / PY_EXE)
        if not WHEELS_CACHE.is_dir() or not any(WHEELS_CACHE.iterdir()):
            raise SystemExit("wheels_cache 是空的，先跑 --refresh-wheels")
        shutil.copytree(WHEELS_CACHE, dst / "wheels")

    zpath = DIST / f"{name}.zip"
    if zpath.exists():
        zpath.unlink()
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for p in sorted(dst.rglob("*")):
            if p.is_file():
                z.write(p, Path(name) / p.relative_to(dst))

    n = sum(1 for p in dst.rglob("*") if p.is_file())
    mb = zpath.stat().st_size / 1024 / 1024
    print(f"  {name:<40} {n:>4} 檔   {mb:>7.1f} MB")
    return zpath


def refresh_wheels():
    WHEELS_CACHE.mkdir(parents=True, exist_ok=True)
    print("  下載 wheels（含 setuptools/wheel，proxy_tools 只有 sdist）...")
    subprocess.run([sys.executable, "-m", "pip", "download",
                    "-r", str(ROOT / "requirements.txt"), "setuptools", "wheel",
                    "-d", str(WHEELS_CACHE)], check=True)


def main():
    ap = argparse.ArgumentParser(description=f"BOM Matcher v{VER} 交付包")
    ap.add_argument("--slim-only", action="store_true")
    ap.add_argument("--refresh-wheels", action="store_true")
    a = ap.parse_args()

    DIST.mkdir(exist_ok=True)
    if a.refresh_wheels:
        refresh_wheels()

    print(f"BOM Matcher v{VER} 交付包")
    build(False)
    if not a.slim_only:
        build(True)
    print("\n完成 →", DIST)


if __name__ == "__main__":
    main()
