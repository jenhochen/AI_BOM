"""
重拍 W 案（pywebview）的 SOP 截圖 —— 用「真的視窗」拍，不是模擬。

v0.10 的 W 截圖是用 page_html() ＋ 假的 window.pywebview.api 在 headless
瀏覽器裡拍，再把舊圖的標題列貼回去。那樣拍不到「真的啟動流程」：
v0.10 的 W 案啟動時頁面腳本中途丟錯、檔案欄位全部「（未找到）」，
SOP 的 W_01 就這樣把壞掉的畫面當正常畫面收進去了。

這支直接開正式視窗（1280x820，含 Windows 標題列），用 evaluate_js 切頁，
Win32 PrintWindow(PW_RENDERFULLCONTENT) 擷取整個視窗：
  - 視窗被別的視窗擋住也拍得到（不是螢幕截圖）
  - 尺寸與 v0.10 相同，content.js 的 marks 座標不必改

用法：
    python tools/shots_W.py            # 拍 W_01～W_07 到 docs/sop/shots/
    python tools/shots_W.py --out DIR  # 先拍到別的資料夾檢查
"""
import ctypes
import ctypes.wintypes as wt
import importlib.util
import sys
import threading
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
from build_dist import VER  # noqa: E402

OUT = Path(sys.argv[sys.argv.index("--out") + 1]) if "--out" in sys.argv \
    else ROOT / "docs" / "sop" / "shots"
W, H = 1280, 820

sys.argv = [sys.argv[0]]          # 不要把 --out 傳給 W 案
spec = importlib.util.spec_from_file_location("wmod", ROOT / f"bom_matcher_W_{VER}.py")
wm = importlib.util.module_from_spec(spec)
spec.loader.exec_module(wm)
import webview  # noqa: E402
from PIL import Image  # noqa: E402

user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
user32.SetProcessDPIAware()


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wt.DWORD), ("biWidth", wt.LONG), ("biHeight", wt.LONG),
                ("biPlanes", wt.WORD), ("biBitCount", wt.WORD),
                ("biCompression", wt.DWORD), ("biSizeImage", wt.DWORD),
                ("biXPelsPerMeter", wt.LONG), ("biYPelsPerMeter", wt.LONG),
                ("biClrUsed", wt.DWORD), ("biClrImportant", wt.DWORD)]


def grab(title: str) -> Image.Image:
    hwnd = user32.FindWindowW(None, title)
    if not hwnd:
        raise SystemExit(f"找不到視窗：{title}")
    r = wt.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    w, h = r.right - r.left, r.bottom - r.top
    hdc = user32.GetWindowDC(hwnd)
    mdc = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
    gdi32.SelectObject(mdc, bmp)
    user32.PrintWindow(hwnd, mdc, 2)            # PW_RENDERFULLCONTENT
    bi = BITMAPINFOHEADER(ctypes.sizeof(BITMAPINFOHEADER), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(bi), 0)
    gdi32.DeleteObject(bmp); gdi32.DeleteDC(mdc); user32.ReleaseDC(hwnd, hdc)
    img = Image.frombuffer("RGB", (w, h), buf, "raw", "BGRX", 0, 1)
    if img.size != (W, H):                       # 保持與 marks 座標同一個尺寸
        img = img.resize((W, H), Image.LANCZOS)
    return img


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    api = wm.Api()
    win = webview.create_window(wm.APP_TITLE, html=wm.page_html(), js_api=api,
                                width=W, height=H)
    wm.WIN = win
    js = win.evaluate_js
    done = []

    def shot(name, wait=1.2):
        time.sleep(wait)
        grab(wm.APP_TITLE).save(OUT / f"{name}.png")
        done.append(name)
        print(f"  ✓ {name}.png")

    def run():
        try:
            time.sleep(3)
            if not js("BOOTED && S.disc.txts.length > 0"):
                raise SystemExit("畫面沒有帶入自動偵測的檔案，不拍")
            shot("W_01_files")
            js("doMatch()")
            for _ in range(40):
                time.sleep(0.5)
                if js("S.rows.length") > 0:
                    break
            shot("W_02_rows", 2)
            # 06/07 要在產出之前拍，日誌才是比對結果（產出後日誌會換成輸出紀錄）
            js("S.filter='unk'; renderRows()"); shot("W_06_unknown")
            js("S.filter='chk'; renderRows()"); shot("W_07_r8")
            js("S.filter='all'; renderRows()")
            js("page('kn')");  shot("W_03_knowledge")
            js("page('hdr')"); shot("W_04_header")
            js("doExport(true)")
            time.sleep(4)
            js("page('out')"); shot("W_05_export")
        except BaseException as e:      # noqa: BLE001
            print("  [失敗]", e)
        finally:
            win.destroy()

    threading.Thread(target=run, daemon=True).start()
    webview.start()
    print(f"完成 {len(done)}/7 → {OUT}")
    return 0 if len(done) == 7 else 1


if __name__ == "__main__":
    sys.exit(main())
