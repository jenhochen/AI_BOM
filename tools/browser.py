"""
截圖／影片工具共用的瀏覽器設定。

2026-11-11 起改用 Microsoft Edge（Windows 內建，不必另外安裝 Chrome）。
環境變數 BM_BROWSER 可以指定其他 Chromium 系瀏覽器的完整路徑。

⚠ Edge headless 用 `--screenshot` 參數截圖時，畫面右緣會疊一個浮動圖示
  （不在 DOM 裡，CSS 藏不掉）。所以這裡的工具一律走 CDP 的
  Page.captureScreenshot —— 它只截網頁本身，不含瀏覽器自己的 UI。
  下面的 EXTRA_FLAGS 再把 Edge 的側邊欄／Copilot 等功能關掉，雙重保險。
"""
import os
from pathlib import Path

_CANDIDATES = [
    r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
]

# Edge 特有的介面元件，headless 截圖時全部關掉
EXTRA_FLAGS = [
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-extensions',
    '--disable-features=msEdgeSidebarV2,msHubApps,msEdgeDiscoverOnPage,'
    'msUndersideButton,msShowcaseCopilot,msEdgeTranslate,EdgeCollections,'
    'msImplicitSignin,msEdgeShoppingUI',
]


def browser_exe() -> str:
    env = os.environ.get('BM_BROWSER')
    if env and Path(env).is_file():
        return env
    for c in _CANDIDATES:
        if Path(c).is_file():
            return c
    raise SystemExit('找不到 Microsoft Edge（msedge.exe），'
                     '請安裝 Edge 或用環境變數 BM_BROWSER 指定瀏覽器路徑')
