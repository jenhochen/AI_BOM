#!/usr/bin/env python3
"""
tools/build_dist.py  —  產出 BOM Matcher 的三種交付包（v0.11 起多 W免安裝版）

    dist/BOM_Matcher_v<VER>_完整離線安裝版/   + .zip   目標機器免上網
    dist/BOM_Matcher_v<VER>_精簡版/           + .zip   目標機器要上網
    dist/BOM_Matcher_v<VER>_W免安裝版/        + .zip   只有 W 案 exe，免 Python（v0.11）

規格對齊 F:\\AI_ini_allin1 的 ADATA SSD Firmware Tool：
  完整離線安裝版 = 精簡版 + python\\（Python 安裝檔）+ wheels\\（離線套件）
  兩者都有 安裝.bat + 安裝說明.txt

用法
----
    python tools/build_dist.py              產出三包（需要 dist/wheels_cache 與 pyinstaller）
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
VER = "0.11"
PY_VER = "3.14.4"
PY_EXE = f"python-{PY_VER}-amd64.exe"
PY_URL = f"https://www.python.org/ftp/python/{PY_VER}/{PY_EXE}"

# 從 Firmware Tool 的離線包借用同一份 Python 安裝檔（兩個專案版本一致）
FW_PY = Path(r"F:\AI_ini_allin1\dist"
             r"\ADATA_SSD_Firmware_Tool_v1.44_完整離線安裝版\python") / PY_EXE


def find_py_installer() -> Path:
    """v0.11：換了機器之後 F:\\AI_ini_allin1 不一定在。依序找：
    Firmware Tool 離線包 → dist\\python_cache\\ → 本專案舊版離線包的 python\\。"""
    cands = [FW_PY, Path(__file__).resolve().parent.parent / "dist" / "python_cache" / PY_EXE]
    cands += sorted((Path(__file__).resolve().parent.parent / "dist")
                    .glob(f"BOM_Matcher_v*_完整離線安裝版/python/{PY_EXE}"), reverse=True)
    for c in cands:
        if c.is_file():
            return c
    raise SystemExit("找不到 Python 安裝檔 " + PY_EXE + "，找過：\n  "
                     + "\n  ".join(map(str, cands)) + "\n下載：" + PY_URL)

DIST = ROOT / "dist"
WHEELS_CACHE = DIST / "wheels_cache"

PROGRAM_FILES = [
    f"bom_core_{VER}.py",
    f"bom_matcher_S_{VER}.py",
    f"bom_matcher_W_{VER}.py",
    "run_CLI.bat",
    "run_S_streamlit.bat",
    "run_W_webview.bat",
    "run_W_webview_debug.bat",
    "requirements.txt",
    f"changelog_BOM_matcher_v{VER}.txt",
    "HANDOVER_bom_matcher.md",
    "pcb.jpg",
    "adata.png",          # 品牌 LOGO（C/D 兩案的左上角）
]
# 2026-10-21 起 HTML 版停止維護，交付包只給 PPTX。
def _latest_sop() -> list:
    """2026-11-13 起 SOP 改成「進版時詢問」才更新，所以 SOP 的版號可能落後程式。
    拿 docs\\ 裡版號最大的那一份，不要只找目前 VER 的檔名（找不到就靜默漏放）。"""
    import re as _re
    found = []
    for f in (Path(__file__).resolve().parent.parent / "docs").glob("BOM_Matcher_v*_操作SOP.pptx"):
        m = _re.search(r"_v(\d+)\.(\d+)_", f.name)
        if m:
            found.append(((int(m.group(1)), int(m.group(2))), f.name))
    return [max(found)[1]] if found else []


SOP = _latest_sop()
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
    "print('  [OK] S 案     streamlit',m.version('streamlit'))\""
    " 2>nul || echo   [--] S 案     streamlit 未安裝\r\n"
    "\"%PYEXE%\" -c \"import webview,importlib.metadata as m;"
    "print('  [OK] W 案     pywebview',m.version('pywebview'))\""
    " 2>nul || echo   [--] W 案     pywebview 未安裝\r\n"
    "echo ============================================================\r\n"
    "echo.\r\n"
    "echo   接著可以執行：\r\n"
    "echo     run_CLI.bat           先用命令列確認核心正常\r\n"
    "echo     run_S_streamlit.bat   S 案（瀏覽器介面）\r\n"
    "echo     run_W_webview.bat     W 案（原生視窗）\r\n"
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

本包含兩套介面（S 案 / W 案）與一套 CLI，共用同一份核心邏輯，
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

  run_S_streamlit.bat    S 案．瀏覽器介面
                         http://localhost:8601
                         使用期間請勿關掉那個黑色主控台視窗
                         ★ 不要直接雙擊 bom_matcher_S_{VER}.py，不會開網頁

  run_W_webview.bat      W 案．原生視窗，不需要瀏覽器
  run_W_webview_debug.bat  同上，但保留主控台（出問題時用這支）

  兩個介面的標題列會標出是哪支 .bat 啟動的，同時開也分得出來：
      BOM Matcher v{VER}  [run_S_streamlit.bat]
      BOM Matcher v{VER}  [run_W_webview.bat]


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

  S 案、W 案跑同一份資料，數字必須完全一樣。


■ v{VER} 的重點

  • W 案多一種交付方式：「W免安裝版」—— 單一 exe，不用裝 Python 或任何套件
        BOM_Matcher_W_v{VER}.exe，解壓縮後雙擊即可（需要 Edge WebView2）
        產出的 Excel 與 .py 版逐格比對 0 差異
  • 修正 W 案開啟時檔案欄位全是「（未找到）」（v0.10 就有的問題，
    當時按「▶ 執行比對」仍可運作，所以沒被發現）
  • 比對邏輯沒有變動：三份 BOM 回歸零退步，SM8008 命中 55、Total Qty 384
  • 操作 SOP 截圖全部重拍；W 案改用真實視窗擷取
  • v0.10 的功能全部沿用：Location (T)/(B)、R8 電容降壓選料、雙面板 Top/Bot、
    同名檔覆蓋確認、ADATA LOGO

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

  Q. S 案瀏覽器打不開
     A. 8601 埠被占用。改執行：
        python -m streamlit run bom_matcher_S_{VER}.py --server.port 8602

  Q. W 案視窗開不起來
     A. 用 run_W_webview_debug.bat 看訊息。
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
        src = find_py_installer()
        cache = DIST / "python_cache" / PY_EXE      # 留一份，舊版包刪掉也還找得到
        if src != cache and not cache.is_file():
            cache.parent.mkdir(exist_ok=True)
            shutil.copy2(src, cache)
        shutil.copy2(src, dst / "python" / PY_EXE)
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


# ════════════════════════════════════════════════════════════════════════════
# v0.11：W 案免安裝版（單一 exe，使用者不用裝 Python 或任何套件）
# ════════════════════════════════════════════════════════════════════════════

EXE_NAME = f"BOM_Matcher_W_v{VER}.exe"


def exe_readme() -> str:
    return rf"""BOM Matcher  v{VER}  W 案免安裝版
================================================================

這一包不需要安裝 Python 或任何套件，直接雙擊 {EXE_NAME} 就能用。

■ 使用方式
  1. 整個資料夾解壓縮到任意位置（不要直接在 zip 裡面執行）
  2. 雙擊 {EXE_NAME}
     第一次啟動約 3～8 秒（程式要先自我解壓縮），之後就是正常的視窗
  3. 操作方式與 W 案相同，見 docs\BOM_Matcher_v{VER}_操作SOP.pptx

■ 資料夾結構（程式以 exe 所在的資料夾為工作目錄）
      {EXE_NAME}     主程式
      TXT\           Altium Text BOM
      ExcelDB\       零件資料庫 + Knowledge + Location 對照表（*.xlsx）
      screenshot\    PCB 圖片（jpg/png）
      output\        產出的 E-BOM
  附的範例資料可以直接按「▶ 執行比對」試跑。

■ 需要的系統元件
  Microsoft Edge WebView2 Runtime —— Windows 11 內建，Windows 10 多半已有。
  雙擊後沒有視窗出現的話，到下列網址安裝 Evergreen Runtime：
  https://developer.microsoft.com/microsoft-edge/webview2/

■ 常見問題
  ・防毒軟體跳出警告：單一 exe 會在啟動時解壓到暫存資料夾，部分防毒會誤判，
    請 IT 將此 exe 加入白名單。
  ・找不到檔案：確認 TXT\ ExcelDB\ 與 exe 放在同一層。
  ・要用 S 案（瀏覽器版）或 CLI：改用「精簡版」或「完整離線安裝版」。
"""


def build_exe_pkg() -> Path | None:
    exe = DIST / "exe" / EXE_NAME
    if not exe.is_file():
        print(f"  [略過] 找不到 {exe.name}，先跑 python tools/build_exe.py")
        return None
    name = f"BOM_Matcher_v{VER}_W免安裝版"
    dst = DIST / name
    if dst.exists():
        shutil.rmtree(dst)
    for sub in ("TXT", "ExcelDB", "output", "screenshot", "docs"):
        (dst / sub).mkdir(parents=True, exist_ok=True)
    shutil.copy2(exe, dst / EXE_NAME)
    shutil.copy2(ROOT / "pcb.jpg", dst / "screenshot" / "pcb.jpg")
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
    write_crlf(dst / "使用說明.txt", exe_readme())

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


def stage_for_teams():
    """
    把要放上 Teams 的兩個檔案集中到 dist\\_上傳Teams\\。

    Teams 的「03_AI Tool.蜂鳥助手 › 已共用」沒有同步到本機，程式碰不到，
    所以只能先備妥、由人拖曳上傳。哪天那個文件庫同步下來了，
    這裡改成直接複製過去就好（路徑記在 HANDOVER）。
    """
    out = DIST / "_上傳Teams"
    out.mkdir(exist_ok=True)
    want = [DIST / f"BOM_Matcher_v{VER}_完整離線安裝版.zip",
            DIST / f"BOM_Matcher_v{VER}_W免安裝版.zip"]      # v0.11
    want += [ROOT / "docs" / s for s in SOP]        # 最新一份 SOP（版號可能落後程式）
    for f in out.iterdir():          # 清掉舊版，免得拖到過期的檔
        if f.is_file():
            f.unlink()
    for f in want:
        if f.is_file():
            shutil.copy2(f, out / f.name)
        else:
            print(f"  [警告] 找不到 {f.name}，Teams 資料夾會缺這一個")
    print(f"\n  要上傳 Teams 的檔案已備妥 -> {out}")
    for f in sorted(out.iterdir()):
        print(f"    {f.name}  ({f.stat().st_size / 1048576:.1f} MB)")


def main():
    ap = argparse.ArgumentParser(description=f"BOM Matcher v{VER} 交付包")
    ap.add_argument("--slim-only", action="store_true")
    ap.add_argument("--refresh-wheels", action="store_true")
    ap.add_argument("--no-exe", action="store_true",
                    help="不重建 W 案 exe（沿用 dist/exe 裡現有的）")
    a = ap.parse_args()

    DIST.mkdir(exist_ok=True)
    if a.refresh_wheels:
        refresh_wheels()

    print(f"BOM Matcher v{VER} 交付包")
    build(False)
    if not a.slim_only:
        build(True)
        if not a.no_exe:     # v0.11：W 案單一 exe（PyInstaller）
            subprocess.run([sys.executable, str(ROOT / "tools" / "build_exe.py")],
                           check=True)
        build_exe_pkg()
        stage_for_teams()
    print("\n完成 →", DIST)


if __name__ == "__main__":
    main()
