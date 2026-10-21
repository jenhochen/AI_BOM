#!/usr/bin/env python3
"""
bom_matcher_C_0.10.py  v0.10  —  C 案：Streamlit 網頁版

啟動：
    run_C_streamlit.bat            （雙擊）
    streamlit run bom_matcher_C_0.10.py

工作目錄固定為「程式所在目錄」，不受啟動時的 cwd 影響。
"""

import io
import os
import sys
import importlib.util
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import streamlit as st


# ── 防呆：本檔不能直接雙擊 / python 執行，必須透過 streamlit run ──────────────
def _under_streamlit() -> bool:
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        try:   # suppress_warning 讓 bare mode 不噴 ScriptRunContext 警告
            return get_script_run_ctx(suppress_warning=True) is not None
        except TypeError:
            return get_script_run_ctx() is not None
    except Exception:
        return False


if not _under_streamlit():
    _here = Path(__file__).resolve().parent
    print()
    print('=' * 62)
    print('  這個檔案不能直接執行（雙擊或 python 執行都不行）。')
    print('  C 案是 Streamlit 網頁版，必須用 streamlit 啟動：')
    print()
    print(f'      雙擊  {_here / "run_C_streamlit.bat"}')
    print()
    print('  或在命令列輸入：')
    print(f'      python -m streamlit run "{Path(__file__).resolve()}"')
    print('=' * 62)
    print()
    try:
        input('按 Enter 關閉…')
    except (EOFError, KeyboardInterrupt):
        pass
    sys.exit(1)


# ── 載入核心（檔名含小數點，不能用一般 import）────────────────────────────────
APP_DIR = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location(
    'bom_core', APP_DIR / 'bom_core_0.10.py')
core = importlib.util.module_from_spec(_spec)
sys.modules['bom_core'] = core
_spec.loader.exec_module(core)

try:
    from streamlit_option_menu import option_menu
    _HAS_MENU = True
except ImportError:
    _HAS_MENU = False


# v0.10：標題列帶上啟動用的批次檔名稱，同時開 C/D 兩個介面時才分得出來
LAUNCHER = 'run_C_streamlit.bat'
APP_TITLE = f'BOM Matcher v{core.VERSION}  [{LAUNCHER}]'

st.set_page_config(page_title=APP_TITLE,
                   page_icon='◈', layout='wide',
                   initial_sidebar_state='expanded')

st.markdown("""
<style>
  .block-container{padding-top:2.2rem;padding-bottom:3rem;max-width:1500px}
  section[data-testid="stSidebar"]{background:#F0F2F6}
  div[data-testid="stMetricValue"]{font-size:1.9rem}
  .kn-note{background:#FFF6EC;border-left:3px solid #e8863a;border-radius:5px;
           padding:8px 11px;color:#6b4413;font-size:.86rem;line-height:1.55}
  .pathbox{background:#F0F2F6;border-radius:7px;padding:7px 11px;
           font-family:Consolas,monospace;font-size:.78rem;color:#43474f}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# state
# ════════════════════════════════════════════════════════════════════════════

def _init():
    if 'disc' in st.session_state:
        return
    d = core.discover(APP_DIR)
    st.session_state.disc   = d
    st.session_state.bom    = d['bom']
    st.session_state.db     = d['db']
    st.session_state.kn     = d['knowledge']
    st.session_state.pcb    = d['pcb']
    st.session_state.pcb_bot = d['pcb_bot']
    st.session_state.sides  = d['sides']
    st.session_state.out    = d['out']
    st.session_state.split  = True
    st.session_state.rows   = None
    st.session_state.log    = []
    st.session_state.hdr    = core.empty_header()
    st.session_state.result = ''

_init()


def log(msg):
    st.session_state.log.append(str(msg))


def _label(p: str) -> str:
    if not p:
        return '（無）'
    try:
        return str(Path(p).relative_to(APP_DIR))
    except ValueError:
        return p


# ── v0.10 新增：檔案瀏覽（補上 D 案本來就有、C 案缺的功能）────────────────
#
# 為什麼要做兩套：
#   1. 原生對話框  與 D 案一致，操作最直覺。但 Streamlit 的指令碼跑在背景
#      執行緒，tkinter 對話框偶爾會開在瀏覽器視窗後面，或在某些環境失效。
#   2. 目錄瀏覽器  純 Streamlit 元件，不依賴 tkinter，一定能用。
#
# v0.07 沒有這兩者，使用者只能接受自動偵測的預設值 —— 加上 discover() 當時
# 的排序 bug，結果 SM8008 的 BOM 選不到，產出的是舊範例檔。

FILE_FILTERS = {
    'bom': [('Text BOM', '*.txt;*.BOM'), ('All files', '*.*')],
    'db':  [('Excel', '*.xlsx'), ('All files', '*.*')],
    'kn':  [('Excel', '*.xlsx'), ('All files', '*.*')],
    'pcb': [('Image', '*.jpg;*.jpeg;*.png'), ('All files', '*.*')],
    'pcb_bot': [('Image', '*.jpg;*.jpeg;*.png'), ('All files', '*.*')],
    'sides': [('Excel', '*.xlsx'), ('Text', '*.txt'), ('All files', '*.*')],
    'out': [('Excel', '*.xlsx')],
}
EXT_FILTERS = {
    'bom': ('.txt', '.bom'),
    'db':  ('.xlsx',),
    'kn':  ('.xlsx',),
    'pcb': ('.jpg', '.jpeg', '.png'),
    'pcb_bot': ('.jpg', '.jpeg', '.png'),
    'sides': ('.xlsx', '.txt'),
    'out': ('.xlsx',),
}


def _native_dialog(kind: str, initial: str):
    """開原生檔案對話框。失敗回傳 None（呼叫端會退回目錄瀏覽器）。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        start = Path(initial).parent if initial else APP_DIR
        if not start.is_dir():
            start = APP_DIR
        if kind == 'out':
            r = filedialog.asksaveasfilename(
                parent=root, title='選擇輸出路徑', initialdir=str(start),
                initialfile=Path(initial).name if initial else '',
                defaultextension='.xlsx', filetypes=FILE_FILTERS[kind])
        else:
            r = filedialog.askopenfilename(
                parent=root, title='選擇檔案', initialdir=str(start),
                filetypes=FILE_FILTERS[kind])
        root.destroy()
        return r or ''
    except Exception as e:
        st.session_state.dlg_err = str(e)
        return None


def _dir_browser(kind: str, key: str):
    """純 Streamlit 目錄瀏覽器。回傳使用者選中的檔案路徑，沒選則 None。"""
    ck = f'cwd_{key}'
    if ck not in st.session_state:
        cur = st.session_state.get(key, '')
        base = Path(cur).parent if cur and Path(cur).parent.is_dir() else APP_DIR
        st.session_state[ck] = str(base)
    cwd = Path(st.session_state[ck])

    c1, c2, c3 = st.columns([1, 1, 4])
    if c1.button('⬆ 上一層', key=f'up_{key}', use_container_width=True):
        st.session_state[ck] = str(cwd.parent)
        st.rerun()
    if c2.button('🏠 程式目錄', key=f'home_{key}', use_container_width=True):
        st.session_state[ck] = str(APP_DIR)
        st.rerun()
    c3.text_input('目前位置', str(cwd), key=f'cwdbox_{key}',
                  label_visibility='collapsed', disabled=True)

    try:
        entries = sorted(cwd.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower()))
    except OSError as e:
        st.error(f'無法讀取目錄：{e}')
        return None

    dirs = [p for p in entries if p.is_dir() and not p.name.startswith('.')]
    exts = EXT_FILTERS[kind]
    files = [p for p in entries if p.is_file() and p.suffix.lower() in exts]

    dc, fc = st.columns(2)
    with dc:
        st.caption(f'資料夾（{len(dirs)}）')
        for d in dirs[:40]:
            if st.button(f'📁 {d.name}', key=f'd_{key}_{d.name}',
                         use_container_width=True):
                st.session_state[ck] = str(d)
                st.rerun()
    with fc:
        st.caption(f'符合的檔案（{len(files)}）')
        if not files:
            st.caption('（這個資料夾沒有符合的檔案）')
        for f in files[:40]:
            mt = datetime.fromtimestamp(f.stat().st_mtime).strftime('%Y-%m-%d')
            if st.button(f'📄 {f.name}　{mt}', key=f'f_{key}_{f.name}',
                         use_container_width=True):
                return str(f)
    return None


def _pick(title, key, options, help_txt='', kind=None):
    """下拉選單（自動偵測結果）＋ 瀏覽… ＋ 目錄瀏覽器 ＋ 手動輸入路徑。"""
    kind = kind or key
    # 上一輪「瀏覽…」或目錄瀏覽器選到的檔案，必須在 selectbox 建立「之前」
    # 才能套用。直接寫 st.session_state['sel_xxx'] 會噴
    #   StreamlitAPIException: cannot be modified after the widget ...
    # —— 實測按 PCB 圖片（BOT 面）的「瀏覽…」必爆。
    pend = st.session_state.pop(f'pend_{key}', None)
    if pend:
        st.session_state[key] = pend
        # 只把 session_state['sel_xxx'] 刪掉是沒用的 —— Streamlit 的
        # widget 值綁在「widget 身分」上，同一個 key 再建出來還是沿用舊值
        # （實測：目錄瀏覽器選了 pcb.jpg，下拉選單仍停在「手動輸入路徑」）。
        # 換掉 key 讓它變成新的 widget，index 才會生效。
        st.session_state[f'gen_{key}'] = st.session_state.get(f'gen_{key}', 0) + 1
    cur = st.session_state.get(key, '')
    opts = list(dict.fromkeys([o for o in options if o] + ([cur] if cur else [])))
    opts.append('— 手動輸入路徑 —')
    idx = opts.index(cur) if cur in opts else len(opts) - 1

    c1, c2 = st.columns([5, 1])
    with c1:
        sel = st.selectbox(
            title, opts, index=idx,
            format_func=lambda p: _label(p) if Path(p).is_absolute() else p,
            key=f'sel_{key}_{st.session_state.get(f"gen_{key}", 0)}',
            help=help_txt)
    with c2:
        st.write('')
        if st.button('瀏覽…', key=f'br_{key}', use_container_width=True):
            r = _native_dialog(kind, cur)
            if r:
                st.session_state[f'pend_{key}'] = r
                st.rerun()
            elif r is None:
                st.session_state[f'open_br_{key}'] = True
                st.rerun()

    if sel == '— 手動輸入路徑 —':
        sel = st.text_input(f'{title}（完整路徑）', value=cur, key=f'txt_{key}')

    with st.expander('📂 目錄瀏覽器', expanded=st.session_state.get(f'open_br_{key}', False)):
        if st.session_state.get('dlg_err'):
            st.warning(f'原生對話框無法使用，請用這裡選：{st.session_state.pop("dlg_err")}')
        got = _dir_browser(kind, key)
        if got:
            st.session_state[f'pend_{key}'] = got
            st.session_state[f'open_br_{key}'] = False
            st.rerun()

    st.session_state[key] = sel
    return sel


def _double(rows):
    """雙面板判定：和 core.write_excel 用同一個依據。"""
    return any(r.get('ref_bot') or r.get('ref_unk') for r in rows)


def _rows_to_df(rows):
    df = pd.DataFrame([{
        '狀態': ('⚠️' if r.get('ref_unk') else '🔴' if r.get('chk')
                 else '🟠' if r['kn']
                 else '🟡' if r['nc'] else '🟢' if r['group'] else '⚪'),
        # 替代料續列的 Item / Qty 留空。data_editor 會把 None/<NA> 顯示成
        # 字串 "None"，所以這兩欄一律用字串表示。
        'Item': str(r['item']) if r['item'] is not None else '',
        'RC/EC': r['rcec'],
        'P/N': r['pn'],
        'Description': r['desc'],
        'Qty': str(r['qty']) if r['qty'] not in (None, '') else '',
        # v0.10：畫面也顯示 (T)/(B)，與輸出 Excel 一致。改過的值寫回 ref_top /
        # ref_bot 時會帶著前綴，寫 Excel 時 tag_side() 冪等，不會疊成 (T)(T)。
        'Location-Top': (core.tag_side(r.get('ref_top'), 'T') if _double(rows)
                         else r['ref']),
        'Location-Bot': (core.tag_side(r.get('ref_bot', ''), 'B') if _double(rows)
                         else ''),
        'Location 未知': r.get('ref_unk', ''),
        'Knowledge Note': r['kn'],
    } for r in rows])
    if not _double(rows):
        df = df.drop(columns=['Location-Bot', 'Location 未知'])
    return df


def _df_to_rows(df, rows):
    """把編輯後的表格寫回 rows（依位置對應，不改變列數）。"""
    out = []
    for i, r in enumerate(rows):
        rec = df.iloc[i]
        n = dict(r)
        n['rcec'] = str(rec['RC/EC'] or '').strip()
        n['pn']   = str(rec['P/N'] or '').strip()
        n['desc'] = str(rec['Description'] or '')
        if 'Location-Bot' in rec.index:
            n['ref_top'] = str(rec['Location-Top'] or '')
            n['ref_bot'] = str(rec['Location-Bot'] or '')
        else:
            n['ref'] = str(rec['Location-Top'] or '')
        n['kn']   = str(rec['Knowledge Note'] or '')
        q = str(rec['Qty'] or '').strip()
        n['qty'] = int(float(q)) if q.replace('.', '', 1).isdigit() else None
        out.append(n)
    return out


def _style(df):
    def color(row):
        c = ('#FDE1E1' if row['狀態'] == '⚠️' else
             '#EDE9FE' if row['狀態'] == '🔴' else
             '#FFE0B2' if row['狀態'] == '🟠' else
             '#FFF2CC' if row['狀態'] == '🟡' else
             '#E2EFDA' if row['狀態'] == '🟢' else '')
        return [f'background-color:{c}'] * len(row) if c else [''] * len(row)
    return df.style.apply(color, axis=1)


def do_match():
    st.session_state.log = []
    try:
        rows, *_ = core.run_match(st.session_state.bom, st.session_state.db,
                                  st.session_state.kn,
                                  split_alt=st.session_state.split, log=log,
                                  sides_path=st.session_state.get('sides', ''))
        st.session_state.rows = rows
        st.session_state.result = ''
    except Exception as e:
        st.session_state.rows = None
        log(f'✗ 錯誤：{e}')


# ════════════════════════════════════════════════════════════════════════════
# sidebar
# ════════════════════════════════════════════════════════════════════════════

PAGES = ['檔案設定', '比對結果', 'Knowledge 命中', '表頭資訊', '產出 E-BOM']
ICONS = ['folder', 'table', 'lightbulb', 'card-heading', 'download']

# ── v0.10：品牌區 ADATA LOGO ＋ 彩蛋（比照 Firmware Tool）────────────────────
#  按住 Shift 連點三次 LOGO → 小蜜蜂遊戲。沒有開關、預設沒有任何提示。
#  用 CCv2 元件：LOGO 的 pointerdown 要在瀏覽器端處理（連點計數不必回 Python），
#  遊戲視窗掛在 document.body，不受側邊欄的版面與 Shadow DOM 樣式限制。
#  ⚠️ 監聽 pointerdown 不是 click、圖片要 draggable=false —— <img> 的原生拖曳
#     會把 click 吃掉（Firmware Tool 踩過的坑，見 bom_core 14b 節）。
_BRAND = st.components.v2.component(
    'bm_brand',
    html="""
<div class="br">
  <img class="lg" draggable="false" alt="ADATA">
  <div class="bg">B</div>
  <div><b>BOM Matcher</b><div class="sub"></div></div>
</div>
""",
    css="""
.br{display:flex;gap:9px;align-items:center;margin-bottom:2px;
    font-family:"Source Sans Pro",sans-serif;color:#31333F}
.br b{font-size:15px}
.sub{font-size:11px;color:#808495;white-space:nowrap}
.lg{width:44px;height:32px;object-fit:contain;background:#fff;border:1px solid #E6E8EE;
    border-radius:7px;padding:2px;box-sizing:border-box;
    -webkit-user-drag:none;user-select:none}
.bg{width:32px;height:32px;border-radius:8px;background:#4F46E5;color:#fff;
    display:flex;align-items:center;justify-content:center;font-weight:700}
""",
    js=core.BEE_GAME_JS + """
function eggClose() {
  const E = window.__bmEgg;
  if (!E) { return; }
  if (E.stop) { E.stop(); E.stop = null; }   // 收掉動畫與鍵盤監聽，否則方向鍵會繼續被吃
  E.el.style.display = 'none';
}

function eggOpen() {
  let E = window.__bmEgg;
  if (!E) {
    const el = document.createElement('div');
    el.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,.55);'
      + 'display:none;align-items:center;justify-content:center;z-index:1000000';
    el.innerHTML = '<div style="background:#0b1220;padding:10px;border-radius:14px;'
      + 'box-shadow:0 20px 50px rgba(0,0,0,.4)">'
      + '<div style="display:flex;justify-content:space-between;align-items:center;'
      + 'margin:0 2px 6px;font:600 11px Consolas,monospace;letter-spacing:2px;'
      + 'color:rgba(226,232,240,.55)"><span>BEE</span>'
      + '<button type="button" title="關閉" style="background:none;border:none;'
      + 'color:rgba(226,232,240,.75);font-size:15px;cursor:pointer;padding:0 4px">✕</button></div>'
      + '<canvas width="420" height="520" style="display:block;border-radius:8px"></canvas></div>';
    el.addEventListener('pointerdown', (e) => { if (e.target === el) { eggClose(); } });
    el.querySelector('button').addEventListener('click', eggClose);
    document.body.appendChild(el);
    E = window.__bmEgg = {el: el, stop: null};
  }
  E.el.style.display = 'flex';
  if (E.stop) { E.stop(); }
  E.stop = window.bmBeeGame(E.el.querySelector('canvas'), eggClose);
}

export default function (component) {
  const { data, parentElement } = component;
  const img = parentElement.querySelector('.lg');
  const badge = parentElement.querySelector('.bg');
  parentElement.querySelector('.sub').textContent = data.sub;
  if (!data.logo) { img.style.display = 'none'; return; }
  badge.style.display = 'none';
  if (img.getAttribute('src') !== data.logo) { img.setAttribute('src', data.logo); }
  if (img.__bmEgg) { return; }               // 每次 rerun 都會呼叫，監聽只掛一次
  img.__bmEgg = {n: 0, t: 0};
  img.addEventListener('pointerdown', (e) => {
    const c = img.__bmEgg;
    if (!e.shiftKey) { c.n = 0; return; }
    e.preventDefault();                      // Shift+點擊會選取文字
    const now = Date.now();
    if (now - c.t > data.ms) { c.n = 0; }
    c.t = now;
    if (++c.n >= data.n) { c.n = 0; eggOpen(); }
  });
}
""")

# v0.10：按「▶ 執行比對」之後自動跳到「比對結果」（與 D 案一致）。
# 按鈕在頁面主體、比側邊欄晚執行，所以不能當場切頁 —— 先存 goto_page，
# 下一次 rerun 在建立選單之前才套用。
_goto = st.session_state.pop('goto_page', None)

with st.sidebar:
    _BRAND(data={'logo': core.app_logo_uri(), 'sub': f'{LAUNCHER} · v{core.VERSION}',
                 'n': core.EGG_CLICKS, 'ms': core.EGG_WINDOW_MS}, key='brand')
    st.write('')

    if _HAS_MENU:
        page = option_menu(
            None, PAGES, icons=ICONS, default_index=1, key='nav',
            manual_select=(PAGES.index(_goto) if _goto in PAGES else None),
            styles={
                'container': {'padding': '0', 'background-color': '#F0F2F6'},
                'nav-link': {'font-size': '14px', 'padding': '9px 12px',
                             'border-radius': '8px', '--hover-color': '#e4e7ee'},
                'nav-link-selected': {'background-color': '#4F46E5'},
            })
    else:
        if _goto in PAGES:
            st.session_state['nav_radio'] = _goto
        page = st.radio('WORKFLOW', PAGES, index=1, key='nav_radio',
                        label_visibility='collapsed')

    st.divider()
    d = st.session_state.disc
    st.caption('自動偵測（程式所在目錄）')
    st.markdown(
        f"<div class='pathbox'>{APP_DIR}</div>", unsafe_allow_html=True)
    st.caption(f"TXT {len(d['txts'])} · DB {len(d['dbs'])} · "
               f"Knowledge {len(d['kns'])} · 圖片 {len(d['imgs'])}")
    if st.button('↺ 重新掃描目錄', use_container_width=True):
        # 只 pop 'disc' 不夠：st.selectbox 一旦建立了 widget key，之後的
        # rerun 會沿用「使用者上次選的值」並忽略傳進去的 index。
        # 結果是原本沒偵測到（值為空）的欄位，掃描後新增的檔案仍然選不到
        # —— 實測把 Location 對照表放進 ExcelDB\ 後按重新掃描，欄位還是空的。
        for k in [x for x in list(st.session_state)
                  if x.startswith(('sel_', 'txt_', 'cwd_', 'open_br_',
                                   'pend_', 'gen_'))]:
            st.session_state.pop(k, None)
        for k in ('disc', 'bom', 'db', 'kn', 'pcb', 'pcb_bot', 'sides'):
            st.session_state.pop(k, None)
        _init(); st.rerun()


# ════════════════════════════════════════════════════════════════════════════
# 1. 檔案設定
# ════════════════════════════════════════════════════════════════════════════

if page == '檔案設定':
    st.title('檔案設定')
    st.caption('預設值全部來自「程式所在目錄」及其 TXT\\ · ExcelDB\\ · '
               'screenshot\\ · output\\ 子資料夾。')

    c1, c2 = st.columns(2)
    with c1:
        _pick('Text BOM (.txt / .BOM)', 'bom', d['txts'],
              '候選依最後修改時間排序，最新的排第一個')
        _pick('零件資料庫 (.xlsx)', 'db', d['dbs'],
              '只列出第 1 列有 P/N + Description 的檔案，自動排除 E-BOM 產出檔')
    with c2:
        _pick('Knowledge DB (.xlsx)', 'kn', d['kns'],
              '只列出第 1 列有 品名 + 包裝 的檔案')
        _pick('PCB 圖片（單面板／TOP 面）', 'pcb', d['imgs'])

    st.divider()
    st.caption('雙面板（選填）—— 提供 Location 對照表就會切換成雙面板模式，'
               '輸出多一欄 Location-Bottom Layer')
    s1, s2 = st.columns(2)
    with s1:
        _pick('Location TOP/BOT 對照表', 'sides', d['sides_all'],
              'xlsx 為主、txt 為輔。第一欄＝Location，第二欄＝Top/Bot', kind='sides')
    with s2:
        _pick('PCB 圖片（BOT 面）', 'pcb_bot', d['imgs'], kind='pcb_bot',
              help_txt='自動偵測要檔名含 top / bot；手動指定任何檔案都可以，'
                       '不受檔名限制。Top 與 Bot 選到同一張時會提醒。')
    if st.session_state.get('sides'):
        st.success('雙面板模式：'
                   + Path(st.session_state.sides).name)
    else:
        st.info('單面板模式（沒有 Location 對照表），輸出與 v0.08 相同')

    oc1, oc2 = st.columns([5, 1])
    with oc1:
        st.session_state.out = st.text_input('輸出路徑', st.session_state.out)
    with oc2:
        st.write('')
        if st.button('瀏覽…', key='br_out', use_container_width=True):
            r = _native_dialog('out', st.session_state.out)
            if r:
                st.session_state.out = core.force_xlsx(r)
                st.rerun()
    st.session_state.split = st.toggle(
        '替代料拆列（一列只有一個 P/N）', value=st.session_state.split,
        help='關閉則回到 v0.06 的「61041115;61041238」單列格式')

    st.write('')
    if st.button('▶ 執行比對', type='primary'):
        do_match()
        st.session_state.goto_page = '比對結果'
        st.rerun()

    if st.session_state.log:
        st.code('\n'.join(st.session_state.log), language=None)


# ════════════════════════════════════════════════════════════════════════════
# 2. 比對結果
# ════════════════════════════════════════════════════════════════════════════

elif page == '比對結果':
    st.title('比對結果')
    st.caption('可直接在表格內修正 P/N、RC/EC、Description、Qty；'
               '修改會帶入產出的 Excel。')

    if st.session_state.rows is None:
        st.info('尚未執行比對。')
        if st.button('▶ 執行比對', type='primary'):
            do_match(); st.rerun()
    else:
        rows = st.session_state.rows
        s = core.stats(rows)
        m = st.columns(6)
        m[0].metric('總項次', s['items'], '含 PCB 列')
        m[1].metric('輸出列數', s['rows'], f"替代料 +{s['rows'] - s['items']}")
        m[2].metric('命中 P/N', s['hit'],
                    f"{s['hit'] / max(1, s['items'] - 1):.0%}")
        m[3].metric('未命中', s['miss'], '需人工補', delta_color='inverse')
        m[4].metric('Knowledge', s['kn'], '橘底標示')
        m[5].metric('Total Qty', s['qty'], '主列加總')

        st.write('')
        tab_e, tab_v = st.tabs(['✎ 編輯', '🎨 色底檢視（唯讀）'])
        df = _rows_to_df(rows)

        with tab_e:
            ed = st.data_editor(
                df, use_container_width=True, height=560, hide_index=True,
                disabled=['狀態', 'Item'],
                column_config={
                    '狀態': st.column_config.TextColumn(width='small'),
                    'Item': st.column_config.TextColumn(width='small'),
                    'RC/EC': st.column_config.SelectboxColumn(
                        options=['RC', 'EC'], width='small'),
                    'P/N': st.column_config.TextColumn(width='medium'),
                    'Description': st.column_config.TextColumn(width='large'),
                    'Qty': st.column_config.TextColumn(width='small'),
                    'Location-Top': st.column_config.TextColumn(width='medium'),
                    'Location-Bot': st.column_config.TextColumn(width='medium'),
                    'Location 未知': st.column_config.TextColumn(width='small'),
                    'Knowledge Note': st.column_config.TextColumn(width='large'),
                },
                key='editor')
            c1, c2 = st.columns([1, 5])
            if c1.button('儲存修改', type='primary'):
                st.session_state.rows = _df_to_rows(ed, rows)
                c2.success('已套用到待產出的資料')
            st.caption('🔴 **R8 降壓選料（請人工確認）**　🟠 Knowledge 命中　'
                       '🟡 原本為 NC　🟢 含替代料（已拆列）　⚪ 一般　'
                       '⚠️ 有 Location 查不到 Top/Bot　— 與輸出 Excel 的填色一致。'
                       'Item / Qty 空白的列 = 替代料續列，Qty 併入主列計算。')

        with tab_v:
            st.dataframe(_style(df), use_container_width=True,
                         height=560, hide_index=True)
            st.caption('Streamlit 的 st.data_editor 不套用 Styler 底色，'
                       '因此「可編輯」與「色底」分成兩個頁籤。')


# ════════════════════════════════════════════════════════════════════════════
# 3. Knowledge 命中
# ════════════════════════════════════════════════════════════════════════════

elif page == 'Knowledge 命中':
    st.title('Knowledge 命中')
    if st.session_state.rows is None:
        st.info('尚未執行比對。')
    else:
        hits = [(i, r) for i, r in enumerate(st.session_state.rows) if r['kn']]
        if not hits:
            st.success('沒有任何 Knowledge 命中。')
        st.caption(f'共 {len(hits)} 列命中。可逐筆確認，忽略後該列不會標橘底、'
                   'Note 也不會寫入 Excel。')
        for i, r in hits:
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                c1.markdown(f"**Item {r['item'] or '↳'} · {r['pn'] or '（無 P/N）'}**"
                            f"　<span style='color:#8b90a0'>{r['desc']}</span>",
                            unsafe_allow_html=True)
                if c2.button('✕ 忽略此筆', key=f'ign{i}'):
                    st.session_state.rows[i]['kn'] = ''
                    st.rerun()
                st.markdown(
                    f"<div class='kn-note'>{r['kn'].replace(chr(10), '<br>')}</div>",
                    unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# 4. 表頭資訊
# ════════════════════════════════════════════════════════════════════════════

elif page == '表頭資訊':
    st.title('表頭資訊')
    st.caption('填入 Excel rows 3–8 的欄位；留空則與 v0.06 相同（空白）。')
    h = st.session_state.hdr
    c1, c2 = st.columns(2)
    with c1:
        h['pcb_no']     = st.text_input('PCB板號', h['pcb_no'])
        h['pcb_pn']     = st.text_input('PCB料號', h['pcb_pn'])
        h['rcec_no']    = st.text_input('RC/EC NO.', h['rcec_no'])
        h['pcb_vendor'] = st.text_input('PCB供應商', h['pcb_vendor'])
        h['pcb_spec']   = st.text_input('PCB 規格', h['pcb_spec'])
    with c2:
        h['serial_no']  = st.text_input('發行序號', h['serial_no'])
        h['issue_date'] = st.text_input('發行(改版)日期', h['issue_date'])
        h['rev']        = st.text_input('目前版次', h['rev'])
        h['ic_type']    = st.text_input('Support IC Type', h['ic_type'])
    st.session_state.hdr = h


# ════════════════════════════════════════════════════════════════════════════
# 5. 產出
# ════════════════════════════════════════════════════════════════════════════

elif page == '產出 E-BOM':
    st.title('產出 E-BOM')
    if st.session_state.rows is None:
        st.info('尚未執行比對。請先到「檔案設定」執行。')
    else:
        s = core.stats(st.session_state.rows)
        st.markdown(f"<div class='pathbox'>{st.session_state.out}</div>",
                    unsafe_allow_html=True)
        st.caption(f"{s['items']} 項（含 PCB）→ {s['rows']} 列　|　"
                   f"Total Qty = {s['qty']}　|　PCB 圖片："
                   f"{_label(st.session_state.pcb)}")
        st.write('')
        # 同名檔要先確認再覆蓋（D 案用 confirm 對話框，這裡用勾選框）
        exists = Path(core.force_xlsx(st.session_state.out)).is_file()
        ok = True
        if exists:
            st.warning(f'這個檔案已經存在：{_label(core.force_xlsx(st.session_state.out))}')
            ok = st.checkbox('確認覆蓋', key='ovw')
        if st.button('▶ 產出 Excel', type='primary', disabled=not ok):
            st.session_state.log = []
            try:
                core.write_excel(st.session_state.rows, st.session_state.out,
                                 st.session_state.pcb or None,
                                 st.session_state.hdr, log=log,
                                 pcb_bot=st.session_state.get('pcb_bot') or None)
                st.session_state.result = st.session_state.out
            except Exception as e:
                log(f'✗ 錯誤：{e}')
                st.session_state.result = ''
            st.rerun()

        if st.session_state.result:
            st.success(f'已產出：{st.session_state.result}')
            with open(st.session_state.result, 'rb') as f:
                st.download_button('⤓ 下載 xlsx', f.read(),
                                   file_name=Path(st.session_state.result).name,
                                   mime='application/vnd.openxmlformats-'
                                        'officedocument.spreadsheetml.sheet')
        if st.session_state.log:
            st.code('\n'.join(st.session_state.log), language=None)
