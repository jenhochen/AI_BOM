"""
重拍 S 案（Streamlit）的 SOP 截圖。

Streamlit 的畫面要互動才看得到（按「執行比對」、切頁籤…），
但 headless Chrome 的 --screenshot 只能抓靜態頁面，所以這裡用 CDP：
  1. 開 headless Chrome 並打開 --remote-debugging-port
  2. 走 WebSocket 下 Runtime.evaluate 點按鈕、Page.captureScreenshot 抓圖

輸出 1440x942，和 W 案的 1280x820 一樣是「原始像素」，
content.js 的 marks 座標就是以這個尺寸為準。

用法：
    python -m streamlit run bom_matcher_S_0.10.py --server.port 8619 &
    python tools/shots_S.py 8619
"""
import asyncio
import base64
import json
import os
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import websockets

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8619
CDP = 9222
W, H = 1440, 942
OUT = Path(r'F:\AI_BOM\docs\sop\shots')
CHROME = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
PROFILE = Path(os.environ['TEMP']) / 'bm_shots_profile'


def chrome_targets():
    with urllib.request.urlopen(f'http://127.0.0.1:{CDP}/json') as r:
        return json.load(r)


async def run():
    proc = subprocess.Popen([
        CHROME, '--headless=new', '--disable-gpu', '--no-sandbox',
        '--hide-scrollbars', '--force-device-scale-factor=1',
        f'--window-size={W},{H}', f'--remote-debugging-port={CDP}',
        f'--user-data-dir={PROFILE}', f'http://localhost:{PORT}',
    ])
    try:
        ws_url = None
        for _ in range(40):
            time.sleep(0.5)
            try:
                pages = [t for t in chrome_targets()
                         if t['type'] == 'page' and 'localhost' in t.get('url', '')]
                if pages:
                    ws_url = pages[0]['webSocketDebuggerUrl']
                    break
            except Exception:
                pass
        if not ws_url:
            print('找不到 Chrome 分頁'); return

        async with websockets.connect(ws_url, max_size=64 * 1024 * 1024) as ws:
            n = 0

            async def cmd(method, params=None):
                nonlocal n
                n += 1
                await ws.send(json.dumps({'id': n, 'method': method,
                                          'params': params or {}}))
                while True:
                    m = json.loads(await ws.recv())
                    if m.get('id') == n:
                        return m.get('result', {})

            async def js(expr):
                r = await cmd('Runtime.evaluate',
                              {'expression': expr, 'awaitPromise': True,
                               'returnByValue': True})
                return r.get('result', {}).get('value')

            async def shot(name, wait=3):
                await asyncio.sleep(wait)
                r = await cmd('Page.captureScreenshot', {'format': 'png'})
                p = OUT / f'{name}.png'
                p.write_bytes(base64.b64decode(r['data']))
                print(f'  saved {name}  ({p.stat().st_size // 1024} KB)')

            await cmd('Page.enable')
            await cmd('Runtime.enable')
            await asyncio.sleep(12)          # Streamlit 首次載入 + 自動偵測

            CLICK = ("[...document.querySelectorAll('button')]"
                     ".filter(b=>b.innerText.includes({t}))[0]?.click()")

            # 側邊欄選單是 streamlit-option-menu 的自訂元件，整個包在
            # iframe 裡，JS 從外層抓不到 —— 只能用滑鼠事件按座標點。
            # 座標對應 1440x942 的畫面。
            NAVXY = {'檔案設定': (110, 193), '比對結果': (110, 243),
                     'Knowledge 命中': (110, 293), '表頭資訊': (110, 343),
                     '產出 E-BOM': (110, 393)}

            async def click_xy(x, y):
                for t in ('mousePressed', 'mouseReleased'):
                    await cmd('Input.dispatchMouseEvent',
                              {'type': t, 'x': x, 'y': y, 'button': 'left',
                               'clickCount': 1})

            async def nav(name):
                await click_xy(*NAVXY[name])
                await asyncio.sleep(4)

            await nav('檔案設定')
            await shot('S_01_files', 4)

            await js("[...document.querySelectorAll('details')]"
                     ".filter(d=>d.innerText.includes('目錄瀏覽器'))[0]"
                     "?.querySelector('summary')?.click()")
            await shot('S_02_browser', 5)

            await js("[...document.querySelectorAll('details')]"
                     ".filter(d=>d.open&&d.innerText.includes('目錄瀏覽器'))[0]"
                     "?.querySelector('summary')?.click()")
            await asyncio.sleep(2)
            await js(CLICK.format(t="'執行比對'"))
            await shot('S_03_rows', 14)

            await nav('產出 E-BOM')
            await js(CLICK.format(t="'確認覆蓋'"))
            await asyncio.sleep(2)
            await js(CLICK.format(t="'產出 Excel'"))
            await shot('S_04_export', 10)

            # 右上角的 ⋮ 也在 Streamlit 自己的元件裡，一樣用座標點
            await click_xy(1399, 29)
            await shot('S_05_menu', 3)
    finally:
        proc.terminate()


asyncio.run(run())
