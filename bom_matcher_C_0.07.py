#!/usr/bin/env python3
"""
bom_matcher_C_0.07.py  v0.07  —  C 案：Streamlit 網頁版

啟動：
    run_C_streamlit.bat            （雙擊）
    streamlit run bom_matcher_C_0.07.py

工作目錄固定為「程式所在目錄」，不受啟動時的 cwd 影響。
"""

import io
import sys
import importlib.util
from pathlib import Path
from datetime import date

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
    'bom_core', APP_DIR / 'bom_core_0.07.py')
core = importlib.util.module_from_spec(_spec)
sys.modules['bom_core'] = core
_spec.loader.exec_module(core)

try:
    from streamlit_option_menu import option_menu
    _HAS_MENU = True
except ImportError:
    _HAS_MENU = False


st.set_page_config(page_title=f'BOM Matcher v{core.VERSION}',
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


def _pick(title, key, options, help_txt=''):
    """下拉選單（自動偵測結果）＋ 手動輸入路徑。"""
    cur = st.session_state.get(key, '')
    opts = list(dict.fromkeys([o for o in options if o] + ([cur] if cur else [])))
    opts.append('— 手動輸入路徑 —')
    idx = opts.index(cur) if cur in opts else len(opts) - 1
    sel = st.selectbox(title, opts, index=idx,
                       format_func=lambda p: _label(p) if p.startswith(str(APP_DIR)) or Path(p).is_absolute() else p,
                       key=f'sel_{key}', help=help_txt)
    if sel == '— 手動輸入路徑 —':
        sel = st.text_input(f'{title}（完整路徑）', value=cur, key=f'txt_{key}')
    st.session_state[key] = sel
    return sel


def _rows_to_df(rows):
    df = pd.DataFrame([{
        '狀態': ('🟠' if r['kn'] else '🟡' if r['nc'] else '🟢' if r['group'] else '⚪'),
        # 替代料續列的 Item / Qty 留空。data_editor 會把 None/<NA> 顯示成
        # 字串 "None"，所以這兩欄一律用字串表示。
        'Item': str(r['item']) if r['item'] is not None else '',
        'RC/EC': r['rcec'],
        'P/N': r['pn'],
        'Description': r['desc'],
        'Qty': str(r['qty']) if r['qty'] not in (None, '') else '',
        'Location-Top': r['ref'],
        'Knowledge Note': r['kn'],
    } for r in rows])
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
        n['ref']  = str(rec['Location-Top'] or '')
        n['kn']   = str(rec['Knowledge Note'] or '')
        q = str(rec['Qty'] or '').strip()
        n['qty'] = int(float(q)) if q.replace('.', '', 1).isdigit() else None
        out.append(n)
    return out


def _style(df):
    def color(row):
        c = ('#FFE0B2' if row['狀態'] == '🟠' else
             '#FFF2CC' if row['狀態'] == '🟡' else
             '#E2EFDA' if row['狀態'] == '🟢' else '')
        return [f'background-color:{c}'] * len(row) if c else [''] * len(row)
    return df.style.apply(color, axis=1)


def do_match():
    st.session_state.log = []
    try:
        rows, *_ = core.run_match(st.session_state.bom, st.session_state.db,
                                  st.session_state.kn,
                                  split_alt=st.session_state.split, log=log)
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

with st.sidebar:
    st.markdown(
        f"<div style='display:flex;gap:9px;align-items:center;margin-bottom:2px'>"
        f"<div style='width:32px;height:32px;border-radius:8px;background:#4F46E5;"
        f"color:#fff;display:flex;align-items:center;justify-content:center;"
        f"font-weight:700'>B</div><div><b style='font-size:15px'>BOM Matcher</b>"
        f"<div style='font-size:11px;color:#808495'>E-BOM workspace · v{core.VERSION}"
        f"</div></div></div>", unsafe_allow_html=True)
    st.write('')

    if _HAS_MENU:
        page = option_menu(
            None, PAGES, icons=ICONS, default_index=1,
            styles={
                'container': {'padding': '0', 'background-color': '#F0F2F6'},
                'nav-link': {'font-size': '14px', 'padding': '9px 12px',
                             'border-radius': '8px', '--hover-color': '#e4e7ee'},
                'nav-link-selected': {'background-color': '#4F46E5'},
            })
    else:
        page = st.radio('WORKFLOW', PAGES, index=1, label_visibility='collapsed')

    st.divider()
    d = st.session_state.disc
    st.caption('自動偵測（程式所在目錄）')
    st.markdown(
        f"<div class='pathbox'>{APP_DIR}</div>", unsafe_allow_html=True)
    st.caption(f"TXT {len(d['txts'])} · DB {len(d['dbs'])} · "
               f"Knowledge {len(d['kns'])} · 圖片 {len(d['imgs'])}")
    if st.button('↺ 重新掃描目錄', use_container_width=True):
        for k in ('disc',):
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
        _pick('Text BOM (.txt)', 'bom', d['txts'])
        _pick('零件資料庫 (.xlsx)', 'db', d['dbs'],
              '只列出第 1 列有 P/N + Description 的檔案，自動排除 E-BOM 產出檔')
    with c2:
        _pick('Knowledge DB (.xlsx)', 'kn', d['kns'],
              '只列出第 1 列有 品名 + 包裝 的檔案')
        _pick('PCB 圖片', 'pcb', d['imgs'])

    st.session_state.out = st.text_input('輸出路徑', st.session_state.out)
    st.session_state.split = st.toggle(
        '替代料拆列（一列只有一個 P/N）', value=st.session_state.split,
        help='關閉則回到 v0.06 的「61041115;61041238」單列格式')

    st.write('')
    if st.button('▶ 執行比對', type='primary'):
        do_match(); st.rerun()

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
                    'Knowledge Note': st.column_config.TextColumn(width='large'),
                },
                key='editor')
            c1, c2 = st.columns([1, 5])
            if c1.button('儲存修改', type='primary'):
                st.session_state.rows = _df_to_rows(ed, rows)
                c2.success('已套用到待產出的資料')
            st.caption('🟠 Knowledge 命中　🟡 原本為 NC　🟢 含替代料（已拆列）　'
                       '⚪ 一般　— 與輸出 Excel 的填色一致。'
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
        if st.button('▶ 產出 Excel', type='primary'):
            st.session_state.log = []
            try:
                core.write_excel(st.session_state.rows, st.session_state.out,
                                 st.session_state.pcb or None,
                                 st.session_state.hdr, log=log)
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
