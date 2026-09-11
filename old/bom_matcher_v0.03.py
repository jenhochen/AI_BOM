#!/usr/bin/env python3
"""
bom_matcher.py  v0.03
Text BOM (Altium) + Excel Parts DB  →  ADATA E-BOM formatted Excel

雙擊執行 / 無參數：啟動 GUI（自動偵測同目錄 TXT / XLSX / pcb.jpg|png）
帶參數執行：CLI 模式

Usage (CLI)
-----------
    python bom_matcher.py <text_bom.txt> <parts_db.xlsx> [-o out.xlsx] [-s SheetName]
"""

import re
import sys
import argparse
import threading
from pathlib import Path
from datetime import date

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext
    _HAS_TK = True
except ImportError:
    _HAS_TK = False

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker

VERSION = "0.03"

_KNOWN_PKGS = {
    '0201','0402','0603','0805','1206',
    'SOD','SOT-723','SOT723','DFN',
    'BGA-247','BGA247','BGA-154','BGA154',
}

# P/N prefix → RC category (else EC)
_RC_PREFIXES = ('6103', '6104', '6900')


# ════════════════════════════════════════════════════════════════════════════════
# 1. Parse Altium text BOM
# ════════════════════════════════════════════════════════════════════════════════

def parse_text_bom(path: str) -> list[dict]:
    items, cur = [], None
    with open(path, encoding='utf-8', errors='replace') as f:
        in_data = False
        for raw in f:
            line = raw.rstrip('\n')
            if 'Item' in line and 'Quantity' in line and 'Reference' in line:
                in_data = True; continue
            if not in_data or line.strip().startswith('__'):
                continue
            if line.startswith('\t\t') and cur:
                extra = line.strip().rstrip(',')
                if extra:
                    cur['ref'] = cur['ref'].rstrip(',') + ',' + extra
                continue
            cols = line.split('\t')
            if len(cols) >= 5 and cols[0].strip().isdigit():
                if cur: items.append(cur)
                cur = {
                    'item': int(cols[0].strip()),
                    'qty':  int(cols[1].strip()),
                    'ref':  cols[2].strip().rstrip(','),
                    'part': cols[3].strip(),
                    'fp':   cols[4].strip(),
                }
    if cur: items.append(cur)
    return items


# ════════════════════════════════════════════════════════════════════════════════
# 2. Read parts database
# ════════════════════════════════════════════════════════════════════════════════

def read_db(path: str, sheet=None) -> pd.DataFrame:
    kw = {'sheet_name': sheet} if sheet else {'sheet_name': 0}
    df = pd.read_excel(path, **kw)
    df['P/N']         = df['P/N'].astype(str).str.strip()
    df['Description'] = df['Description'].astype(str).str.strip()
    return df


# ════════════════════════════════════════════════════════════════════════════════
# 3. NC helpers
# ════════════════════════════════════════════════════════════════════════════════

def is_nc(part: str) -> bool:
    return any(s.strip().upper() == 'NC' for s in re.split(r'[/\\]+', part))

def strip_nc(part: str) -> str:
    segs = [s for s in re.split(r'[/\\]+', part) if s.strip().upper() != 'NC']
    return '/'.join(segs).strip('/')


# ════════════════════════════════════════════════════════════════════════════════
# 4. Value & package helpers
# ════════════════════════════════════════════════════════════════════════════════

_RES_PATTERNS = [
    (re.compile(r'([\d.]+)\s*R\b', re.I), r'\1Ω'),
    (re.compile(r'([\d.]+)\s*K\b', re.I), r'\1KΩ'),
    (re.compile(r'([\d.]+)\s*M\b', re.I), r'\1MΩ'),
]

def normalize(val: str) -> str:
    v = re.sub(r'_.*$', '', val).strip()
    for pat, rep in _RES_PATTERNS:
        v2, n = pat.subn(rep, v)
        if n: return v2
    return v

def pkg_of(fp: str) -> str:
    m = re.search(r'(0\d{3})', fp)
    if m: return m.group(1)
    fu = fp.upper()
    for tag in ('SOD','SOT-723','SOT723','DFN','BGA-247','BGA247','BGA-154','BGA154'):
        if tag in fu: return tag
    return ''

def pkg_ok(pkg: str, desc: str) -> bool:
    if not pkg: return True
    return pkg.replace('-','').upper() in desc.replace('-','').upper()

def value_in_desc(val: str, desc: str) -> bool:
    pattern = r'(?<![0-9A-Za-z.])' + re.escape(val) + r'(?![0-9A-Za-z])'
    return bool(re.search(pattern, desc, re.IGNORECASE))


# ════════════════════════════════════════════════════════════════════════════════
# 5. Core matching
# ════════════════════════════════════════════════════════════════════════════════

def find_matches(part: str, fp: str, db: pd.DataFrame) -> list:
    val  = strip_nc(part)
    pkg  = pkg_of(fp)
    vu   = val.upper()
    nrm  = normalize(val)
    fpt  = fp[0].upper() if fp else ''
    seen, matches = set(), []

    def add(row):
        k = str(row['P/N'])
        if k not in seen:
            seen.add(k); matches.append(row)

    for _, row in db.iterrows():
        desc = str(row['Description'])
        du   = desc.upper()

        if fpt == 'C':
            if not pkg_ok(pkg, desc): continue
            if value_in_desc(val, desc): add(row)

        elif fpt == 'R':
            if not pkg_ok(pkg, desc): continue
            if value_in_desc(val, desc) or value_in_desc(nrm, desc): add(row)

        elif fpt == 'L':
            if not pkg_ok(pkg, desc): continue
            if 'BEAD' in vu and 'BEAD' in du:
                m = re.search(r'(\d+)\s*[Oo]hm', val)
                if m and f"{m.group(1)}Ω" in desc: add(row)
            else:
                m = re.match(r'([\d.]+uH)', val, re.I)
                if m and m.group(1).upper() in du and 'INDUCTOR' in du: add(row)

        else:
            if vu == 'TVS':
                if 'TVS' in du:
                    if 'SOD' not in du and pkg in ('0201','0402'): add(row)
                    elif 'SOD' in du and 'SOD' in fp.upper(): add(row)
                continue
            if value_in_desc(val, desc):
                add(row); continue
            tokens = [t for t in re.split(r'[_\-/\s]+', val) if len(t) >= 4]
            if any(t.upper() in du for t in tokens): add(row)

    return matches


def match_item(item: dict, db: pd.DataFrame) -> tuple:
    """Returns (pn_str, desc_str, is_nc_flag). P/N alternatives joined by ';'."""
    nc   = is_nc(item['part'])
    hits = find_matches(item['part'], item['fp'], db)

    if not hits:
        v = strip_nc(item['part'])
        p = pkg_of(item['fp'])
        display = p if p in _KNOWN_PKGS else item['fp']
        d = f"{v} {display}".strip()
        return '', d + (' (原本為NC)' if nc else ''), nc

    pn = str(hits[0]['P/N'])
    for h in hits[1:]:
        pn += f";{h['P/N']}"

    d = str(hits[0]['Description']).strip()
    return pn, d + (' (原本為NC)' if nc else ''), nc


def classify_rcec(pn: str) -> str:
    """RC: P/N starts with 6103(resistor)/6104(cap)/6900(paste). Else EC."""
    first = pn.split(';')[0].strip()
    return 'RC' if first.startswith(_RC_PREFIXES) else 'EC'


# ════════════════════════════════════════════════════════════════════════════════
# 6. Excel output – ADATA E-BOM template style
# ════════════════════════════════════════════════════════════════════════════════

def _s(style='thin') -> Side:
    return Side(style=style, color='000000')

def _border(l='thin', r='thin', t='thin', b='thin') -> Border:
    def sd(x): return _s(x) if x else Side()
    return Border(left=sd(l), right=sd(r), top=sd(t), bottom=sd(b))

def _apply(cell, value=None, font=None, fill=None, align=None, brd=None):
    if value is not None: cell.value = value
    if font:  cell.font      = font
    if fill:  cell.fill      = fill
    if align: cell.alignment = align
    if brd:   cell.border    = brd

def _lbl(name='Arial', size=14, bold=False):
    return Font(name=name, size=size, bold=bold)

def _ctr(wrap=False):
    return Alignment(horizontal='center', vertical='center', wrap_text=wrap)

def _lft(wrap=True):
    return Alignment(horizontal='left', vertical='center', wrap_text=wrap)

def _rgt():
    return Alignment(horizontal='right', vertical='center')


def _build_header(ws):
    """Rows 1-8: ADATA company header + PCB info table."""
    TBL_BRD  = _border()
    F_TITLE  = Font(name='Arial', size=14, bold=True)
    F_MAIN   = Font(name='Arial', size=24, bold=True)
    F_LBL_A  = Font(name='Arial',   size=14)
    F_LBL_K  = Font(name='標楷體', size=14)
    F_VAL    = Font(name='Arial',   size=14)

    # ── Row 1 : Company name ────────────────────────────────────────────────
    ws.merge_cells('A1:O1')
    _apply(ws['A1'], 'ADATA Technology Co., Ltd.', font=F_TITLE,
           align=Alignment(horizontal='left', vertical='center'),
           brd=_border(t='medium', b='thin', l='medium', r='medium'))
    ws.row_dimensions[1].height = 17.4

    # ── Row 2 : Main title ──────────────────────────────────────────────────
    ws.merge_cells('A2:L2')
    _apply(ws['A2'], 'M.2 SSD E-BOM', font=F_MAIN,
           align=_ctr(), brd=TBL_BRD)
    ws.row_dimensions[2].height = 30.6

    # ── Rows 3-8 : Info fields ──────────────────────────────────────────────
    rows_def = [
        # (row, left_label, right_label_K)
        (3, 'PCB板號：',        '發行序號：'),
        (4, 'PCB料號：',        '發行(改版)日期：'),
        (5, 'RC/EC NO.：',      '目前版次：'),
        (6, 'PCB供應商：',      None),
        (7, 'PCB 規格：',       '環保等級：'),
        (8, 'Support IC Type：', None),
    ]
    for row, lbl_a, lbl_k in rows_def:
        r = ws.row_dimensions[row]
        r.height = 39.6 if row == 5 else (39.6 if row == 7 else 19.8)
        if row == 8: r.height = 20.4

        ws.merge_cells(f'A{row}:C{row}')
        _apply(ws[f'A{row}'], lbl_a, font=F_LBL_A, align=_rgt(), brd=TBL_BRD)

        ws.merge_cells(f'D{row}:J{row}')
        # row 5 version display
        val_5 = '0' if row == 5 else None
        _apply(ws[f'D{row}'], val_5, font=F_VAL, align=_lft(False), brd=TBL_BRD)

        ws.merge_cells(f'K{row}:L{row}')
        if lbl_k:
            _apply(ws[f'K{row}'], lbl_k, font=F_LBL_K, align=_rgt(), brd=TBL_BRD)
        else:
            _apply(ws[f'K{row}'], None, brd=TBL_BRD)

        ws.merge_cells(f'M{row}:O{row}')
        if row == 5:
            _apply(ws[f'M{row}'], '0', font=F_VAL, align=_lft(False), brd=TBL_BRD)
        elif row == 7:
            _apply(ws[f'M{row}'],
                   '□A級 □S級 ■H級 □G級\n□OTHER 特殊客戶要求  :  ',
                   font=F_LBL_K, align=_lft(True), brd=TBL_BRD)
        else:
            _apply(ws[f'M{row}'], None, brd=TBL_BRD)

    # merge K7:L8 and M7:O8 for 環保等級 spanning two rows
    # (already merged K7:L7 above; need K7:L8 span)
    # unmerge K7:L7 and re-merge K7:L8
    ws.unmerge_cells('K7:L7')
    ws.merge_cells('K7:L8')
    _apply(ws['K7'], '環保等級：', font=F_LBL_K, align=_rgt(), brd=TBL_BRD)
    ws.unmerge_cells('M7:O7')
    ws.merge_cells('M7:O8')
    _apply(ws['M7'],
           '□A級 □S級 ■H級 □G級\n□OTHER 特殊客戶要求  :  ',
           font=F_LBL_K, align=_lft(True), brd=TBL_BRD)


def _add_pcb_image(ws, img_path):
    """Embed PCB image spanning rows 9-35 (0-indexed anchor: row 8→34, col 0→14)."""
    # Set row heights for image zone (rows 9-35)
    for r in range(9, 36):
        ws.row_dimensions[r].height = 28.2

    if not img_path:
        return
    p = Path(img_path)
    if not p.exists():
        print(f"[警告] 找不到 PCB 圖片: {img_path}")
        return
    try:
        img = XLImage(str(p))
        anchor = TwoCellAnchor(editAs='twoCell')
        anchor._from = AnchorMarker(col=0, colOff=0, row=8, rowOff=0)
        anchor.to    = AnchorMarker(col=14, colOff=0, row=34, rowOff=0)
        img.anchor   = anchor
        ws.add_image(img)
        print(f"✓ PCB 圖片已嵌入: {p.name}")
    except Exception as e:
        print(f"[警告] PCB 圖片嵌入失敗: {e}")


def _build_bom_header(ws, row: int):
    """Build BOM column header row."""
    F  = Font(name='Arial', size=12)
    BR = _border('medium', 'medium', 'thin', 'thin')

    headers = [
        ('A', 'Item',                 _ctr()),
        ('B', 'RC/EC',                _ctr()),
        ('C', 'P/N',                  _ctr()),
        ('D', 'Description',          _ctr()),
        ('E', 'Qty',                  _ctr()),
        ('F', 'Location-Top Layer',   _ctr()),
    ]
    for col, val, aln in headers:
        c = ws[f'{col}{row}']
        _apply(c, val, font=F, align=aln, brd=_border())

    ws.merge_cells(f'G{row}:O{row}')
    _apply(ws[f'G{row}'], 'Location-Bottom Layer',
           font=F, align=_ctr(), brd=_border('thin', 'medium', 'thin', 'thin'))
    ws.row_dimensions[row].height = 15.0


def _write_data_row(ws, row: int, item_no: int, rcec: str,
                    pn: str, desc: str, qty, ref: str, nc: bool):
    F_A  = Font(name='Arial', size=14)   # Item col (slightly larger like template)
    F_BD = Font(name='Arial', size=12)

    FL_NC  = PatternFill('solid', start_color='FFF2CC')  # yellow = NC
    FL_ALT = PatternFill('solid', start_color='E2EFDA')  # green  = alt P/N

    fill = FL_NC if nc else (FL_ALT if ';' in pn else None)

    data = [
        ('A', item_no, F_A,  _ctr()),
        ('B', rcec,    F_BD, _ctr()),
        ('C', pn,      F_BD, _lft()),
        ('D', desc,    F_BD, _lft()),
        ('E', qty,     F_BD, _ctr()),
        ('F', ref,     F_BD, _lft()),
    ]
    for col, val, fnt, aln in data:
        c = ws[f'{col}{row}']
        _apply(c, val, font=fnt, align=aln, brd=_border())
        if fill: c.fill = fill

    # Bottom Layer (merged G:O) – always empty for top-layer BOM
    ws.merge_cells(f'G{row}:O{row}')
    _apply(ws[f'G{row}'], None, font=F_BD,
           align=_lft(), brd=_border('thin', 'medium', 'thin', 'thin'))
    if fill:
        ws[f'G{row}'].fill = fill

    ws.row_dimensions[row].height = 17.4


def _write_total_row(ws, row: int, total_qty: int):
    F  = Font(name='Arial', size=12, bold=True)
    FL = PatternFill('solid', start_color='D9E1F2')

    ws.merge_cells(f'A{row}:B{row}')
    _apply(ws[f'A{row}'], 'Total', font=F,
           align=_ctr(), brd=_border(), fill=FL)
    for col in ['C', 'D']:
        c = ws[f'{col}{row}']
        _apply(c, None, font=F, brd=_border(), fill=FL)
    _apply(ws[f'E{row}'], total_qty, font=F,
           align=_ctr(), brd=_border(), fill=FL)
    ws.merge_cells(f'F{row}:O{row}')
    _apply(ws[f'F{row}'], None, font=F, brd=_border(), fill=FL)
    ws.row_dimensions[row].height = 17.4


def _write_footer(ws, row: int):
    """Remarks + approval section."""
    F_SM  = Font(name='Arial', size=10)
    F_LBL = Font(name='Arial', size=11)

    # blank row
    row += 1

    # Remark
    ws.merge_cells(f'A{row}:O{row}')
    _apply(ws[f'A{row}'],
           'Remark：Top Layer only BOM. NC items are listed for reference.',
           font=F_SM, align=_lft(False))
    ws.row_dimensions[row].height = 17.4
    row += 2

    # Legend
    FL_NC  = PatternFill('solid', start_color='FFF2CC')
    FL_ALT = PatternFill('solid', start_color='E2EFDA')
    ws.cell(row, 1, '圖例').font = Font(name='Arial', bold=True, size=9)
    for i, (label, fl, note) in enumerate([
        ('黃底', FL_NC,  '原本為 NC 的零件'),
        ('綠底', FL_ALT, '含替代料 (P/N 欄位以「;」分隔)'),
    ]):
        c = ws.cell(row + 1 + i, 1, label)
        c.fill = fl; c.font = Font(name='Arial', size=9)
        ws.cell(row + 1 + i, 2, note).font = Font(name='Arial', size=9)
    row += 4

    # Approval line
    ws.merge_cells(f'A{row}:B{row+1}')
    _apply(ws[f'A{row}'], '    Approval  :  ', font=F_LBL, align=_ctr())
    ws.merge_cells(f'C{row}:D{row+1}')
    _apply(ws[f'C{row}'], None)
    ws.merge_cells(f'E{row}:F{row+1}')
    _apply(ws[f'E{row}'], 'Rechecked: ', font=F_LBL, align=_ctr())
    ws.merge_cells(f'G{row}:H{row+1}')
    _apply(ws[f'G{row}'], ' Prepared:', font=F_LBL, align=_ctr())
    ws.merge_cells(f'I{row}:J{row+1}')
    ws.merge_cells(f'K{row}:O{row+1}')


def _set_col_widths(ws):
    widths = {
        'A': 5.1,  'B': 10.1, 'C': 29.1, 'D': 77.8, 'E': 9.2,
        'F': 54.3, 'G': 11.0, 'H': 16.0, 'I': 13.0, 'J': 13.0,
        'K':  9.0, 'L': 12.0, 'M': 13.3, 'N':  9.0, 'O': 13.0,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def write_excel(items: list, db: pd.DataFrame,
                out_path: str, pcb_img: str = None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'BOM'
    ws.sheet_view.showGridLines = False

    _set_col_widths(ws)
    _build_header(ws)
    _add_pcb_image(ws, pcb_img)

    BOM_HDR_ROW = 36
    _build_bom_header(ws, BOM_HDR_ROW)

    total_qty = 0
    data_row  = BOM_HDR_ROW + 1

    for seq, item in enumerate(items, 1):
        pn, desc, nc = match_item(item, db)
        qty = item['qty']
        ref = item['ref']
        total_qty += qty
        rcec = classify_rcec(pn)
        _write_data_row(ws, data_row, seq, rcec, pn, desc, qty, ref, nc)
        data_row += 1

    _write_total_row(ws, data_row, total_qty)
    _write_footer(ws, data_row)

    ws.freeze_panes = f'A{BOM_HDR_ROW + 1}'

    wb.save(out_path)
    print(f"✓ 輸出: {out_path}  |  {len(items)} 項  |  Total Qty = {total_qty}")


# ════════════════════════════════════════════════════════════════════════════════
# 7. Auto-detect helpers
# ════════════════════════════════════════════════════════════════════════════════

def _find_pcb_img(directory: str) -> str:
    """Look for pcb.jpg or pcb.png (case-insensitive) in directory."""
    d = Path(directory)
    for name in ('pcb.jpg', 'pcb.jpeg', 'pcb.png', 'PCB.jpg', 'PCB.JPG', 'PCB.png'):
        p = d / name
        if p.exists():
            return str(p)
    return ''


# ════════════════════════════════════════════════════════════════════════════════
# 8. GUI
# ════════════════════════════════════════════════════════════════════════════════

class _StdoutToText:
    def __init__(self, widget, root):
        self.widget = widget; self.root = root

    def write(self, s: str):
        def _insert():
            self.widget.configure(state='normal')
            self.widget.insert('end', s)
            self.widget.see('end')
            self.widget.configure(state='disabled')
        self.root.after(0, _insert)

    def flush(self): pass


class BomMatcherApp(tk.Tk if _HAS_TK else object):

    PAD = 8

    def __init__(self):
        super().__init__()
        self.title(f'BOM Matcher  v{VERSION}')
        self.resizable(True, True)
        self.minsize(720, 540)
        self._build_ui()
        self.after(100, self._auto_detect)

    def _build_ui(self):
        P = self.PAD
        self.configure(padx=P, pady=P)

        cwd_frame = tk.LabelFrame(self, text='目前工作目錄', padx=P, pady=4)
        cwd_frame.pack(fill='x', pady=(0, P))
        tk.Label(cwd_frame, text=str(Path.cwd()), anchor='w', fg='#333')\
            .pack(fill='x')

        file_frame = tk.LabelFrame(self, text='檔案設定', padx=P, pady=P)
        file_frame.pack(fill='x', pady=(0, P))
        file_frame.columnconfigure(1, weight=1)

        self.var_bom = tk.StringVar()
        self.var_db  = tk.StringVar()
        self.var_pcb = tk.StringVar()
        self.var_out = tk.StringVar()

        rows = [
            ('Text BOM (.txt)',    self.var_bom, self._browse_bom),
            ('Excel DB (.xlsx)',   self.var_db,  self._browse_db),
            ('PCB 圖片 (pcb.jpg)', self.var_pcb, self._browse_pcb),
            ('輸出 BOM (.xlsx)',   self.var_out, self._browse_out),
        ]
        for r, (label, var, cmd) in enumerate(rows):
            tk.Label(file_frame, text=label, anchor='e', width=18)\
                .grid(row=r, column=0, sticky='e', pady=3, padx=(0, 6))
            tk.Entry(file_frame, textvariable=var, width=58)\
                .grid(row=r, column=1, sticky='ew', pady=3)
            tk.Button(file_frame, text='瀏覽…', width=7, command=cmd)\
                .grid(row=r, column=2, padx=(6, 0), pady=3)

        btn_frame = tk.Frame(self)
        btn_frame.pack(fill='x', pady=(0, P))
        self.btn_run = tk.Button(
            btn_frame, text='▶  執行', width=16,
            bg='#1F4E79', fg='white', font=('Arial', 11, 'bold'),
            relief='flat', cursor='hand2', command=self._run,
        )
        self.btn_run.pack(side='right')

        log_frame = tk.LabelFrame(self, text='執行日誌', padx=P, pady=P)
        log_frame.pack(fill='both', expand=True)
        self.log_text = scrolledtext.ScrolledText(
            log_frame, state='disabled', height=12,
            font=('Consolas', 9), bg='#1e1e1e', fg='#d4d4d4',
            insertbackground='white', wrap='word',
        )
        self.log_text.pack(fill='both', expand=True)

    def _auto_detect(self):
        cwd   = Path.cwd()
        txts  = sorted(cwd.glob('*.txt'))
        xlsxs = sorted(cwd.glob('*.xlsx'))
        self._log(f'[掃描] 目錄：{cwd}\n')

        if len(txts) == 1:
            self._set_bom(str(txts[0]))
            self._log(f'[自動偵測] Text BOM → {txts[0].name}')
        elif len(txts) > 1:
            self._log(f'[提示] 找到 {len(txts)} 個 TXT，請手動選擇：')
            for f in txts: self._log(f'          {f.name}')
        else:
            self._log('[提示] 未找到 TXT BOM，請手動選擇')

        if len(xlsxs) == 1:
            self._set_db(str(xlsxs[0]))
            self._log(f'[自動偵測] Excel DB  → {xlsxs[0].name}')
        elif len(xlsxs) > 1:
            self._log(f'[提示] 找到 {len(xlsxs)} 個 XLSX，請手動選擇：')
            for f in xlsxs: self._log(f'          {f.name}')
        else:
            self._log('[提示] 未找到 Excel DB，請手動選擇')

        pcb = _find_pcb_img(str(cwd))
        if pcb:
            self.var_pcb.set(pcb)
            self._log(f'[自動偵測] PCB 圖片  → {Path(pcb).name}')
        else:
            self._log('[提示] 未找到 pcb.jpg/pcb.png（圖片為選填）')

        self._log('')

    def _set_bom(self, path: str):
        self.var_bom.set(path)
        if not self.var_out.get():
            p   = Path(path)
            out = p.parent / f'BOM_{p.stem}_{date.today():%Y%m%d}.xlsx'
            self.var_out.set(str(out))

    def _set_db(self, path: str):
        self.var_db.set(path)

    def _browse_bom(self):
        f = filedialog.askopenfilename(
            title='選擇 Text BOM', initialdir=Path.cwd(),
            filetypes=[('Text BOM', '*.txt'), ('All', '*.*')])
        if f: self._set_bom(f)

    def _browse_db(self):
        f = filedialog.askopenfilename(
            title='選擇 Excel 零件資料庫', initialdir=Path.cwd(),
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if f: self._set_db(f)

    def _browse_pcb(self):
        f = filedialog.askopenfilename(
            title='選擇 PCB 圖片', initialdir=Path.cwd(),
            filetypes=[('Image', '*.jpg *.jpeg *.png'), ('All', '*.*')])
        if f: self.var_pcb.set(f)

    def _browse_out(self):
        f = filedialog.asksaveasfilename(
            title='輸出 BOM 路徑', initialdir=Path.cwd(),
            defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if f: self.var_out.set(f)

    def _run(self):
        bom = self.var_bom.get().strip()
        db  = self.var_db.get().strip()
        out = self.var_out.get().strip()
        if not bom:
            messagebox.showerror('錯誤', '請選擇 Text BOM 檔案'); return
        if not db:
            messagebox.showerror('錯誤', '請選擇 Excel 零件資料庫'); return
        if not out:
            messagebox.showerror('錯誤', '請指定輸出路徑'); return
        self.btn_run.configure(state='disabled')
        self._log('=' * 55)
        threading.Thread(
            target=self._run_task,
            args=(bom, db, self.var_pcb.get().strip(), out),
            daemon=True,
        ).start()

    def _run_task(self, bom: str, db_path: str, pcb: str, out: str):
        old = sys.stdout
        sys.stdout = _StdoutToText(self.log_text, self)
        try:
            print(f'BOM Matcher v{VERSION}')
            print(f'  Text BOM : {bom}')
            print(f'  Excel DB : {db_path}')
            print(f'  PCB 圖片 : {pcb or "（未指定）"}')
            print(f'  Output   : {out}\n')
            items = parse_text_bom(bom)
            print(f'✓ 解析 Text BOM: {len(items)} 項')
            db_df = read_db(db_path)
            print(f'✓ 讀取零件資料庫: {len(db_df)} 筆')
            write_excel(items, db_df, out, pcb or None)
            self.after(0, lambda: messagebox.showinfo('完成', f'BOM 已產出：\n{out}'))
        except Exception as e:
            print(f'\n✗ 錯誤：{e}')
            import traceback; traceback.print_exc()
            self.after(0, lambda: messagebox.showerror('執行錯誤', str(e)))
        finally:
            sys.stdout = old
            self.after(0, lambda: self.btn_run.configure(state='normal'))
            self.after(0, lambda: self._log(''))

    def _log(self, msg: str):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')
        self.log_text.configure(state='disabled')


# ════════════════════════════════════════════════════════════════════════════════
# 9. Entry point
# ════════════════════════════════════════════════════════════════════════════════

def _cli():
    ap = argparse.ArgumentParser(
        description=f'BOM Matcher v{VERSION} — Text BOM + Excel DB → ADATA E-BOM Excel'
    )
    ap.add_argument('text_bom')
    ap.add_argument('excel_db')
    ap.add_argument('-o', '--output')
    ap.add_argument('-s', '--sheet')
    args = ap.parse_args()

    out = args.output or str(
        Path(args.text_bom).parent /
        f'BOM_{Path(args.text_bom).stem}_{date.today():%Y%m%d}.xlsx'
    )
    pcb = _find_pcb_img(str(Path(args.text_bom).parent))

    print(f'BOM Matcher v{VERSION}')
    print(f'  Text BOM : {args.text_bom}')
    print(f'  Excel DB : {args.excel_db}')
    print(f'  PCB 圖片 : {pcb or "（未找到 pcb.jpg/png）"}')
    print(f'  Output   : {out}\n')

    items = parse_text_bom(args.text_bom)
    print(f'✓ 解析 Text BOM: {len(items)} 項')
    db = read_db(args.excel_db, args.sheet)
    print(f'✓ 讀取零件資料庫: {len(db)} 筆')
    write_excel(items, db, out, pcb or None)


def main():
    if len(sys.argv) > 1:
        _cli()
    else:
        if not _HAS_TK:
            print("錯誤：GUI 需要 tkinter，請改用 CLI 模式或安裝 tkinter")
            sys.exit(1)
        BomMatcherApp().mainloop()


if __name__ == '__main__':
    main()
