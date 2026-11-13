"""
把 W 案（pywebview）打包成「單一 exe」，使用者不必安裝 Python 或任何套件。

    python tools/build_exe.py            # 產出 dist/exe/BOM_Matcher_W_v<VER>.exe
    python tools/build_exe.py --console  # 留著主控台視窗（除錯用）

需要（只有打包機要裝，使用者不用）：
    pip install pyinstaller

做法與坑
  1. core 是用 importlib 從檔案路徑載入的（檔名含小數點，不能 import），
     PyInstaller 的靜態分析看不到它 —— 所以：
       - bom_core_<VER>.py 以「資料檔」內嵌，執行時從 sys._MEIPASS 載入
       - core 用到的套件（pandas / openpyxl / PIL）要列進 hidden-import
  2. onefile 每次啟動會解壓到 %TEMP%\\_MEIxxxx，__file__ 指向那裡；
     程式目錄改用 sys.executable 的目錄（core 與 W 都有處理，見 v0.11 註解）。
  3. adata.png 也內嵌：exe 旁邊沒有這張圖時，LOGO 照樣顯示。
  4. 視窗程式（--windowed）沒有 stdout，print() 會直接略過、不會出錯。
  5. 使用者機器需要 Microsoft Edge WebView2 Runtime（Win11 內建、Win10 多半已有）。
     這是系統元件，不是 Python 套件，exe 無法內含。
  6. streamlit 等 S 案套件明確排除，exe 才不會膨脹。
"""
import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
from build_dist import VER  # noqa: E402  —— 版號只在 build_dist.py 改一處

NAME = f"BOM_Matcher_W_v{VER}"
OUT_DIR = ROOT / "dist" / "exe"
W_PY = ROOT / f"bom_matcher_W_{VER}.py"
CORE_PY = ROOT / f"bom_core_{VER}.py"
LOGO = ROOT / "adata.png"

HIDDEN = [
    "pandas", "openpyxl", "openpyxl.drawing.image", "openpyxl.styles",
    "openpyxl.utils", "PIL", "PIL.Image",
]
EXCLUDE = [
    "streamlit", "streamlit_option_menu", "altair", "pyarrow", "pydeck",
    "tornado", "matplotlib", "IPython", "jupyter", "tkinter", "pytest",
    "PyInstaller",
]


def make_icon(work: Path) -> Path | None:
    """adata.png（橫向）→ 置中貼到正方形白底 → 多尺寸 .ico。"""
    if not LOGO.is_file():
        return None
    from PIL import Image
    im = Image.open(LOGO).convert("RGBA")
    side = max(im.size)
    sq = Image.new("RGBA", (side, side), (255, 255, 255, 0))
    sq.paste(im, ((side - im.width) // 2, (side - im.height) // 2), im)
    ico = work / "app.ico"
    sq.save(ico, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64),
                        (128, 128), (256, 256)])
    return ico


def main():
    ap = argparse.ArgumentParser(description=f"BOM Matcher W v{VER} → 單一 exe")
    ap.add_argument("--console", action="store_true", help="保留主控台（除錯用）")
    a = ap.parse_args()

    for f in (W_PY, CORE_PY):
        if not f.is_file():
            raise SystemExit(f"找不到 {f}")

    work = Path(tempfile.mkdtemp(prefix="bm_exe_"))
    sep = ";"   # Windows 的 --add-data 分隔符
    cmd = [
        sys.executable, "-m", "PyInstaller", str(W_PY),
        "--name", NAME, "--onefile", "--noconfirm", "--clean",
        "--distpath", str(OUT_DIR),
        "--workpath", str(work / "build"),
        "--specpath", str(work),
        "--add-data", f"{CORE_PY}{sep}.",
        "--collect-submodules", "webview",
    ]
    cmd += ["--console"] if a.console else ["--windowed"]
    if LOGO.is_file():
        cmd += ["--add-data", f"{LOGO}{sep}."]
    ico = make_icon(work)
    if ico:
        cmd += ["--icon", str(ico)]
    for h in HIDDEN:
        cmd += ["--hidden-import", h]
    for e in EXCLUDE:
        cmd += ["--exclude-module", e]

    print(f"BOM Matcher W v{VER} → {OUT_DIR / (NAME + '.exe')}")
    r = subprocess.run(cmd, cwd=ROOT)
    shutil.rmtree(work, ignore_errors=True)
    if r.returncode:
        raise SystemExit(f"PyInstaller 失敗（exit {r.returncode}）")
    exe = OUT_DIR / f"{NAME}.exe"
    print(f"✓ {exe}  ({exe.stat().st_size / 1024 / 1024:.1f} MB)")


if __name__ == "__main__":
    main()
