#!/usr/bin/env python3
"""
bom_core_0.09.py  v0.09
BOM Matcher 核心模組（無 GUI）

Text BOM (Altium) + Excel Parts DB  →  ADATA E-BOM formatted Excel

本檔可獨立當 CLI 使用，也供下列前端 import：
    bom_matcher_C_0.09.py   Streamlit 網頁版
    bom_matcher_D_0.09.py   pywebview 原生視窗版

Usage (CLI)
-----------
    python bom_core_0.09.py                     # 全自動：掃描程式所在目錄
    python bom_core_0.09.py -o out.xlsx
    python bom_core_0.09.py BOM.txt db.xlsx -k Knowledge.xlsx -o out.xlsx
"""

import re
import sys
import argparse
from pathlib import Path
from datetime import date

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import range_boundaries

VERSION = "0.10"

# ── v0.09：輸出編碼自保 ──────────────────────────────────────────────────
# Windows 主控台預設 cp950，印不出 ✓ 這類字元。輸出被重導向到檔案或管線時
# （python bom_core_0.09.py > out.txt），Python 會改用 locale 編碼而噴
#     UnicodeEncodeError: 'cp950' codec can't encode character '✓'
# 把程式整個打斷。.bat 有設 PYTHONIOENCODING，但直接用 python 執行時沒有，
# 而且 python -I（isolated mode）會忽略所有 PYTHON* 環境變數。
# 所以不依賴環境變數，程式自己把 stdout/stderr 設成 UTF-8。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# ════════════════════════════════════════════════════════════════════════════════
# 0. 程式所在目錄（v0.09：全面取代 Path.cwd()）
# ════════════════════════════════════════════════════════════════════════════════

APP_DIR = Path(__file__).resolve().parent

# 依 F:\AI_BOM 現行資料夾結構的搜尋子目錄（找不到就略過）
SUB_TXT = ('TXT', 'txt')
SUB_DB  = ('ExcelDB', 'exceldb', 'DB', 'db')
SUB_IMG = ('screenshot', 'Screenshot', 'image', 'img')
SUB_OUT = ('output', 'Output', 'out')

# 產出檔的指紋：A1 是公司名 → 是 E-BOM 產出，不是零件資料庫
_OUTPUT_MARK = 'ADATA Technology Co., Ltd.'


_KNOWN_PKGS = {
    '0201', '0402', '0603', '0805', '1206',
    'SOD', 'SOT-723', 'SOT723', 'DFN',
    'BGA-247', 'BGA247', 'BGA-154', 'BGA154',
}
# RC class: resistors (6103*) and capacitors (6104*) only
_RC_PREFIXES = ('6103', '6104')

# Description-based RC detection patterns (for items with no P/N)
_CAP_VAL_START = re.compile(r'^\d+\.?\d*\s*(pF|nF|uF)', re.IGNORECASE)
_RES_WORD      = re.compile(r'\bRES\b',                  re.IGNORECASE)
_OHM_VAL       = re.compile(r'\d+\.?\d*\s*(Ω|ohm)',      re.IGNORECASE)


# ════════════════════════════════════════════════════════════════════════════════
# 1. Parse Altium text BOM
# ════════════════════════════════════════════════════════════════════════════════

def parse_text_bom(path: str) -> list:
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
    df.columns = [str(c).strip() for c in df.columns]
    if 'P/N' not in df.columns or 'Description' not in df.columns:
        raise ValueError(
            f"『{Path(path).name}』不是零件資料庫："
            f"第 1 列需要 'P/N' 與 'Description' 欄，實際為 {list(df.columns)[:6]}")
    df['P/N']         = df['P/N'].astype(str).str.strip()
    df['Description'] = df['Description'].astype(str).str.strip()
    # 濾掉空列 / NaN 列
    df = df[(df['P/N'] != '') & (df['P/N'].str.lower() != 'nan')]
    return df.reset_index(drop=True)


def is_parts_db(path) -> bool:
    """第 1 列必須同時有 P/N 與 Description，且不是 E-BOM 產出檔。"""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        wb.close()
    except Exception:
        return False
    vals = [str(v).strip() for v in first if v is not None]
    if not vals:
        return False
    if vals[0] == _OUTPUT_MARK:          # 是自己產出的 E-BOM
        return False
    up = [v.upper() for v in vals]
    return 'P/N' in up and 'DESCRIPTION' in up


def is_knowledge_db(path) -> bool:
    """
    第 1 列包含『品名』與『包裝』→ Knowledge 檔。
    v0.09 修正：不可只看檔名。PREVIEW_BOM_knowledge_v0.06.xlsx 檔名含
    knowledge 但其實是 E-BOM 產出檔，會載入 57 筆垃圾並產生大量假命中。
    """
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        wb.close()
    except Exception:
        return False
    vals = [str(v).strip() for v in first if v is not None]
    if not vals or vals[0] == _OUTPUT_MARK:      # 是自己產出的 E-BOM
        return False
    return '品名' in vals and '包裝' in vals


# ════════════════════════════════════════════════════════════════════════════════
# 3. NC helpers
# ════════════════════════════════════════════════════════════════════════════════

# v0.09：NC 的三種寫法
#   NC(47nF/10V)   括號式   <- SM8008 用這種，v0.08 完全認不出來
#   10pF/NC        斜線式   <- 舊 BOM 用這種
#   2.2uF_NC       底線式
# v0.08 只認斜線式，SM8008 有 26 項是括號式，後果是雙重的：
#   (1) 沒標黃底、沒有「(原本為NC)」
#   (2) strip_nc 沒剝掉 NC，拿整串 'NC(47nF/10V)' 去比對必定落空
#       —— 實測其中 10 項剝掉後 DB 裡確實有料
_NC_PAREN = re.compile(r'^\s*NC\s*\((.*)\)\s*$', re.IGNORECASE)


def is_nc(part: str) -> bool:
    if _NC_PAREN.match(part or ''):
        return True
    return any(s.strip().upper() == 'NC'
               for s in re.split(r'[/\\_]+', part or ''))


def strip_nc(part: str) -> str:
    m = _NC_PAREN.match(part or '')
    if m:
        part = m.group(1).strip()
    segs = [s for s in re.split(r'[/\\]+', part) if s.strip().upper() != 'NC']
    out = '/'.join(segs).strip('/')
    out = re.sub(r'_+NC\b', '', out, flags=re.IGNORECASE)
    out = re.sub(r'\bNC_+', '', out, flags=re.IGNORECASE)
    return out.strip()


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
    for tag in ('SOD', 'SOT-723', 'SOT723', 'DFN',
                'BGA-247', 'BGA247', 'BGA-154', 'BGA154'):
        if tag in fu: return tag
    return ''

def pkg_ok(pkg: str, desc: str) -> bool:
    if not pkg: return True
    return pkg.replace('-', '').upper() in desc.replace('-', '').upper()

def value_in_desc(val: str, desc: str) -> bool:
    pattern = r'(?<![0-9A-Za-z.])' + re.escape(val) + r'(?![0-9A-Za-z])'
    return bool(re.search(pattern, desc, re.IGNORECASE))


# ════════════════════════════════════════════════════════════════════════════════
# 4b. v0.09 新增：屬性式解析（R1–R6）
# ════════════════════════════════════════════════════════════════════════════════
#
# v0.08 是拿整串 Part 去 DB 的 Description 裡找子字串。SM8008 的 Part 欄位從
# 舊格式的「4.7uF」變成「4.7uF/10V」之後，整串永遠找不到 —— 32 項電容命中 0。
#
# 這種事換個案子還會再發生，所以不為單一元件打補丁，改成通用規則：
#   R2 切欄位   以 / _ 空白 逗號 切段，第一個「解析得出數值」的段落＝主值
#   R3 屬性辨識 其餘段落逐一認：耐壓 數字V／誤差 數字%／電流 數字A／直流阻抗 數字mR
#               認不得的段落「忽略但記錄」，不讓它害整筆比不到
#   R4 數值化   容值→pF、阻值→Ω、電感→uH，解決 0.47uF ≡ 470nF、4K7 ≡ 4.7K
#   R6 方向性   耐壓／電流 DB ≥ BOM；誤差 DB ≤ BOM

_CAP_MUL = {'P': 1.0, 'N': 1e3, 'U': 1e6, 'M': 1e9}
_RES_MUL = {'': 1.0, 'R': 1.0, 'K': 1e3, 'M': 1e6, 'G': 1e9}

_RE_CAP   = re.compile(r'^([\d.]+)\s*([pnum])F$', re.IGNORECASE)
_RE_RES   = re.compile(r'^([\d.]+)\s*([RKMG]?)(?:Ω|OHM)?$', re.IGNORECASE)
_RE_RES2  = re.compile(r'^(\d+)([RKMG])(\d+)$', re.IGNORECASE)     # 4K7 / 1R5
_RE_IND   = re.compile(r'^([\d.]+)\s*([num])H$', re.IGNORECASE)
_RE_VOLT  = re.compile(r'^[≧>=]*\s*([\d.]+)\s*V$', re.IGNORECASE)
_RE_TOL   = re.compile(r'^[±+\-]*\s*([\d.]+)\s*%$')
_RE_AMP   = re.compile(r'^([\d.]+)\s*A$', re.IGNORECASE)
_RE_DCR   = re.compile(r'^([\d.]+)\s*m[RΩ]$', re.IGNORECASE)
_SPLIT    = re.compile(r'[\/_,\s]+')


def cap_pf(s: str):
    """容值 → pF。0.47uF 與 470nF 都會得到 470000.0。"""
    m = _RE_CAP.fullmatch((s or '').strip())
    return float(m.group(1)) * _CAP_MUL[m.group(2).upper()] if m else None


def res_ohm(s: str):
    """阻值 → Ω。支援 200K / 0R / 1.15K / 4K7 / 300Ω。"""
    t = (s or '').strip().replace('Ω', '')
    m = _RE_RES2.fullmatch(t)
    if m:
        return (float(m.group(1)) + float(m.group(3)) / 10 ** len(m.group(3))) \
            * _RES_MUL[m.group(2).upper()]
    m = _RE_RES.fullmatch(t)
    if m and m.group(1):
        return float(m.group(1)) * _RES_MUL[m.group(2).upper()]
    return None


def ind_uh(s: str):
    """電感 → uH。"""
    m = _RE_IND.fullmatch((s or '').strip())
    if not m:
        return None
    mul = {'N': 1e-3, 'U': 1.0, 'M': 1e3}[m.group(2).upper()]
    return float(m.group(1)) * mul


def parse_part(part: str, kind: str) -> dict:
    """
    把 Part 欄位拆成屬性。kind = 'C' | 'R' | 'L'。

    回傳 {'val': 數值(已換算基本單位), 'raw': 主值原文,
          'volt': 耐壓V, 'tol': 誤差%, 'amp': 電流A, 'dcr': 直流阻抗mΩ,
          'rest': [無法辨識的段落]}
    val 為 None 代表「主值解析不出來」—— 這是 R7 診斷要抓的重點訊號。
    """
    conv = {'C': cap_pf, 'R': res_ohm, 'L': ind_uh}.get(kind)
    out = {'val': None, 'raw': '', 'volt': None, 'tol': None,
           'amp': None, 'dcr': None, 'rest': []}
    if not conv:
        return out
    segs = [s for s in _SPLIT.split(strip_nc(part or '')) if s]
    for i, s in enumerate(segs):
        if out['val'] is None:
            v = conv(s)
            if v is not None:
                out['val'], out['raw'] = v, s
                continue
        m = _RE_VOLT.fullmatch(s)
        if m and out['volt'] is None:
            out['volt'] = float(m.group(1)); continue
        m = _RE_TOL.fullmatch(s)
        if m and out['tol'] is None:
            out['tol'] = float(m.group(1)); continue
        m = _RE_AMP.fullmatch(s)
        if m and out['amp'] is None:
            out['amp'] = float(m.group(1)); continue
        m = _RE_DCR.fullmatch(s)
        if m and out['dcr'] is None:
            out['dcr'] = float(m.group(1)); continue
        if out['val'] is not None or i > 0:
            out['rest'].append(s)
    return out


# ── DB 側：描述解析成結構化索引（只做一次）──────────────────────────────────
_DB_CAP = re.compile(
    r'CAP\s+([\d.]+\s*[pnum]F)\s+(\d{4})\s*([≧>=]*\s*[\d.]+\s*V)?', re.IGNORECASE)
_DB_RES = re.compile(
    r'RES\s+(?:LF\s+)?([\d.]+\s*[KMG]?)\s*Ω\s+(\d{4})', re.IGNORECASE)
_DB_BEAD = re.compile(r'BEAD\s+([\d.]+)\s*Ω\s+(\d{4})', re.IGNORECASE)
_DB_IND = re.compile(r'INDUCTOR\s+([\d.]+\s*[num]H)', re.IGNORECASE)
_DB_TOL = re.compile(r'[±+\-]\s*([\d.]+)\s*%')
_DB_AMP = re.compile(r'([\d.]+)\s*A\b', re.IGNORECASE)


def build_db_index(db: pd.DataFrame) -> dict:
    """把 DB 描述解析成 {kind: {(值, 封裝): [(P/N, 屬性dict), ...]}}。"""
    idx = {'C': {}, 'R': {}, 'L': {}}
    for _, row in db.iterrows():
        pn, d = str(row['P/N']), str(row['Description'])
        tol = float(_DB_TOL.search(d).group(1)) if _DB_TOL.search(d) else None
        m = _DB_CAP.search(d)
        if m:
            v = cap_pf(m.group(1))
            if v is not None:
                volt = None
                if m.group(3):
                    mv = re.search(r'([\d.]+)', m.group(3))
                    volt = float(mv.group(1)) if mv else None
                idx['C'].setdefault((v, m.group(2)), []).append(
                    (pn, {'volt': volt, 'tol': tol, 'desc': d}))
            continue
        m = _DB_RES.search(d)
        if m:
            v = res_ohm(m.group(1))
            if v is not None:
                idx['R'].setdefault((v, m.group(2)), []).append(
                    (pn, {'volt': None, 'tol': tol, 'desc': d}))
            continue
        m = _DB_BEAD.search(d)
        if m:
            amp = float(_DB_AMP.search(d).group(1)) if _DB_AMP.search(d) else None
            idx['L'].setdefault((float(m.group(1)), m.group(2)), []).append(
                (pn, {'bead': True, 'amp': amp, 'tol': tol, 'desc': d}))
            continue
        m = _DB_IND.search(d)
        if m:
            v = ind_uh(m.group(1))
            if v is not None:
                amp = float(_DB_AMP.search(d).group(1)) if _DB_AMP.search(d) else None
                pkg = pkg_of(d)
                idx['L'].setdefault((v, pkg), []).append(
                    (pn, {'bead': False, 'amp': amp, 'tol': tol, 'desc': d}))
    return idx


def _bead_ohm(part: str):
    """FB/120R/55mR/2A → 120.0。BEAD 的阻值寫在第二段，第一段是 FB 標記。"""
    for s in _SPLIT.split(strip_nc(part or '')):
        if s.upper() in ('FB', 'BEAD'):
            continue
        v = res_ohm(s)
        if v is not None:
            return v
    return None


def _fmt_v(v) -> str:
    """10.0 -> '10'、6.3 -> '6.3'。"""
    return f'{v:g}'


def tag_side(refs, side: str) -> str:
    """
    v0.10 T1：Location-Top / Location-Bottom 欄的每個 Location 冠上 (T) / (B)。

        'R49,R50'  + 'T'  ->  '(T)R49,(T)R50'

    空字串原樣回傳（續列、PCB 列、查不到 Top/Bot 的列都會是空的）。
    """
    if not refs:
        return ''
    out = []
    for r in str(refs).split(','):
        r = r.strip()
        if not r:
            continue
        # 冪等：使用者在 C/D 介面手打過前綴時不要再疊一層
        out.append(r if _RE_TAGGED.match(r) else f'({side}){r}')
    return ','.join(out)


# ── R8：電容降壓選料（v0.10 新增）────────────────────────────────────────
# 容值與封裝一致、但 DB 耐壓低於設計值時可以選，Description 要加註記。
# 只放寬「電容的耐壓」—— 電流與誤差維持 R6 不放寬：
#   耐壓不足會縮短壽命但電路仍會動，看得到、可人工判斷；
#   誤差/電流不符是功能性錯誤，放寬等於讓程式替工程師做決定。
_RE_TAGGED = re.compile(r'^\([TB]\)', re.IGNORECASE)

R8_MARK = '請再確認'


def _r8_note(want_volt) -> str:
    return f' (原始設計為{_fmt_v(want_volt)}V, {R8_MARK})'


def structured_matches(part: str, fp: str, idx: dict, strict: bool = True) -> list:
    """
    用屬性比對。回傳 [(P/N, desc), ...]；無法解析或無命中都回空清單。
    strict=False 時跳過 R6 的方向性條件，只比「值＋封裝」——
    診斷用，拿來分辨「DB 根本沒這個值」還是「有但規格不符」。
    """
    kind = (fp or ' ')[0].upper()
    if kind not in idx:
        return []
    pkg = pkg_of(fp)
    if kind == 'L' and ('FB' in (part or '').upper() or 'BEAD' in (part or '').upper()):
        val, want = _bead_ohm(part), parse_part(part, 'R')
        cands = idx['L'].get((val, pkg), []) if val is not None else []
        cands = [c for c in cands if c[1].get('bead')]
    else:
        want = parse_part(part, kind)
        if want['val'] is None:
            return []
        cands = idx[kind].get((want['val'], pkg), [])

    out, low_volt = [], []
    for pn, attr in cands:
        if strict:
            if want.get('volt') is not None and attr.get('volt') is not None \
                    and attr['volt'] < want['volt']:
                # R6：DB 耐壓必須 ≥ BOM。R8 把電容的這一條留作候補。
                if kind == 'C':
                    low_volt.append((pn, attr))
                continue
            if want.get('amp') is not None and attr.get('amp') is not None \
                    and attr['amp'] < want['amp']:
                continue                  # R6：DB 電流必須 ≥ BOM（不放寬）
            if want.get('tol') is not None and attr.get('tol') is not None \
                    and attr['tol'] > want['tol']:
                continue                  # R6：DB 誤差必須 ≤ BOM（不放寬）
        out.append((pn, attr['desc']))

    # R8：規格完全符合的一個都沒有時，才考慮降壓品；耐壓最接近設計值的排前面
    if strict and not out and low_volt:
        low_volt.sort(key=lambda t: -(t[1].get('volt') or 0))
        note = _r8_note(want['volt'])
        out = [(pn, attr['desc'] + note) for pn, attr in low_volt]
    return out


# ── R7：未命中一律歸因 ──────────────────────────────────────────────────────
# v0.08 比不到就靜默留白，所以 Part 格式變了沒人發現。有了歸因，
# SM8008 一跑就會看到「32 項電容的主值全部解析失敗」，當場抓到。
D_HIT        = '命中'
D_NOVAL      = '主值解析失敗'
D_SPEC       = '規格不符（耐壓/電流/誤差）'
D_PKG        = '值有但封裝不符'
D_NOVALUE_DB = 'DB 沒有這個值'
D_NOTPARAM   = '非標準件（IC/連接器等）'


def diagnose(part: str, fp: str, idx: dict, hit: bool) -> str:
    if hit:
        return D_HIT
    kind = (fp or ' ')[0].upper()
    if kind not in idx:
        return D_NOTPARAM
    if kind == 'L' and ('FB' in (part or '').upper() or 'BEAD' in (part or '').upper()):
        val = _bead_ohm(part)
    else:
        val = parse_part(part, kind)['val']
    if val is None:
        return D_NOVAL
    # 值＋封裝都對、只是被 R6 擋掉 → 這是「DB 有料但規格不夠」，
    # 跟「DB 根本沒這個值」是兩回事，不能混為一談
    if structured_matches(part, fp, idx, strict=False):
        return D_SPEC
    if any(k[0] == val for k in idx[kind]):
        return D_PKG
    return D_NOVALUE_DB


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
                    if 'SOD' not in du and pkg in ('0201', '0402'): add(row)
                    elif 'SOD' in du and 'SOD' in fp.upper(): add(row)
                continue
            if value_in_desc(val, desc):
                add(row); continue
            tokens = [t for t in re.split(r'[_\-/\s]+', val) if len(t) >= 4]
            if any(t.upper() in du for t in tokens): add(row)

    return matches


def match_item(item: dict, db: pd.DataFrame, idx: dict = None) -> tuple:
    """
    v0.09：回傳 (candidates, nc)
      candidates = [(pn, desc), ...]  ← 每個元素代表一個 P/N（含替代料）
      無命中時 candidates = [('', '推測描述')]

    比對順序：先走屬性式（structured_matches），比不到才退回 v0.08 的
    子字串比對。這樣是「只增不減」—— 舊 BOM 走得通的路徑完全沒動，
    新格式則由屬性式接手。
    """
    nc = is_nc(item['part'])
    pairs = structured_matches(item['part'], item['fp'], idx) if idx else []

    if not pairs:
        hits = find_matches(item['part'], item['fp'], db)
        pairs = [(str(h['P/N']), str(h['Description']).strip()) for h in hits]

    if not pairs:
        v = strip_nc(item['part'])
        p = pkg_of(item['fp'])
        display = p if p in _KNOWN_PKGS else item['fp']
        return [('', f"{v} {display}".strip())], nc

    return pairs, nc


def classify_rcec(pn: str, desc: str = '', fp: str = '') -> str:
    """
    RC: P/N starts with 6103 (resistor) or 6104 (capacitor).
    If no P/N, detect from description text or footprint prefix (C/R).
    All else → EC.
    """
    first = pn.split(';')[0].strip()
    if first.startswith(_RC_PREFIXES):
        return 'RC'
    if not first:
        if _CAP_VAL_START.search(desc): return 'RC'
        if _RES_WORD.search(desc):      return 'RC'
        if _OHM_VAL.search(desc):       return 'RC'
        if fp and fp[0].upper() in ('C', 'R'): return 'RC'
    return 'EC'


# ════════════════════════════════════════════════════════════════════════════════
# 6. Knowledge loading & matching
# ════════════════════════════════════════════════════════════════════════════════

def load_knowledge(path: str) -> list:
    """Load Layout_and_BOM_Knowledge file → list of {pn, 品名, 包裝, note}."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    entries = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 5:
            continue
        _, pn, name, pkg, note = row[0], row[1], row[2], row[3], row[4]
        if not any([pn, name, pkg]):
            continue
        if isinstance(pn, (int, float)):
            pn = str(int(pn))
        entries.append({
            'pn':   str(pn).strip()   if pn   else '',
            '品名': str(name).strip() if name else '',
            '包裝': str(pkg).strip()  if pkg  else '',
            'note': str(note).strip() if note else '',
        })
    wb.close()
    return entries


def find_knowledge_hits(pn_str: str, desc_str: str, knowledge: list) -> str:
    """
    Return merged note (newline-separated, deduplicated) for all matching
    entries. Match triggers if any ONE of: PN / 品名 / 包裝 matches.
    """
    pn_str   = str(pn_str   or '')
    desc_str = str(desc_str or '')
    bom_pns  = {p.strip() for p in pn_str.split(';') if p.strip()}
    desc_nd  = desc_str.replace('-', '').upper()

    seen_notes, notes = set(), []
    for k in knowledge or []:
        hit = False
        if k['pn'] and k['pn'] in bom_pns:
            hit = True
        if not hit and k['品名'] and k['品名'].upper() in desc_str.upper():
            hit = True
        if not hit and k['包裝']:
            if k['包裝'].replace('-', '').upper() in desc_nd:
                hit = True
        if hit and k['note'] and k['note'] not in seen_notes:
            seen_notes.add(k['note'])
            notes.append(k['note'])
    return '\n'.join(notes)


# ════════════════════════════════════════════════════════════════════════════════
# 6b. v0.09 新增：Location TOP/BOT 對照表（雙面板）
# ════════════════════════════════════════════════════════════════════════════════
#
# 格式：第一欄＝Location，第二欄＝Top/Bot。
# 刻意「不看表頭名稱」—— 實際拿到的兩份檔案表頭就不一致：
#     TXT  : Location / T/B
#     XLSX : Location / T_B
# Top/Bot 接受 TOP / T / TOP SIDE / BOT / B / BOTTOM / BOTTOM SIDE。

_SIDE_TOP = {'TOP', 'T', 'TOPSIDE', 'TOP SIDE', 'F', 'FRONT'}
_SIDE_BOT = {'BOT', 'B', 'BOTTOM', 'BOTSIDE', 'BOTTOM SIDE', 'BACK'}


def _norm_side(v) -> str:
    s = str(v or '').strip().upper().replace('_', '').replace('-', '')
    if s in _SIDE_TOP: return 'TOP'
    if s in _SIDE_BOT: return 'BOT'
    return ''


def _read_text_any(path) -> str:
    """
    TXT 是選用格式，但還是要做編碼自動偵測：
    有 BOM 就照 BOM；沒有先試 UTF-8，失敗再退 cp950（繁中 Windows 預設）。
    """
    raw = Path(path).read_bytes()
    if raw[:2] in (b'\xff\xfe', b'\xfe\xff'):
        return raw.decode('utf-16')
    if raw[:3] == b'\xef\xbb\xbf':
        return raw[3:].decode('utf-8')
    for enc in ('utf-8', 'cp950', 'latin-1'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def load_sides(path: str) -> dict:
    """讀 Location 對照表（.xlsx 為主、.txt 為輔）→ {Location大寫: 'TOP'|'BOT'}。"""
    p = Path(path)
    pairs = []
    if p.suffix.lower() in ('.xlsx', '.xlsm'):
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        for row in wb.active.iter_rows(min_row=1, values_only=True):
            if row and len(row) >= 2:
                pairs.append((row[0], row[1]))
        wb.close()
    else:
        for ln in _read_text_any(p).splitlines():
            if not ln.strip():
                continue
            parts = re.split(r'\t|,|\s{2,}', ln)
            if len(parts) >= 2:
                pairs.append((parts[0], parts[1]))

    out = {}
    for ref, side in pairs:
        # Excel 可能把 Location 存成數字（例如 1E5 會變 100000.0）
        if isinstance(ref, float) and ref.is_integer():
            ref = int(ref)
        r = str(ref or '').strip().upper()
        s = _norm_side(side)
        if r and s:
            out[r] = s
    return out


def is_side_table(path) -> bool:
    """第二欄要有 TOP/BOT 這類值，至少 5 筆才算數（避免誤認別的 xlsx）。"""
    try:
        if Path(path).suffix.lower() in ('.xlsx', '.xlsm'):
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            n = 0
            for row in ws.iter_rows(min_row=1, max_row=40, values_only=True):
                if row and len(row) >= 2 and _norm_side(row[1]):
                    n += 1
            wb.close()
            return n >= 5
        head = _read_text_any(path)[:4096].splitlines()[:40]
        n = sum(1 for ln in head
                for parts in [re.split(r'\t|,|\s{2,}', ln)]
                if len(parts) >= 2 and _norm_side(parts[1]))
        return n >= 5
    except Exception:
        return False


def side_report(rows: list, sides: dict) -> dict:
    """統計 Location 對不上的狀況。"""
    bad, allrefs = [], set()
    for r in rows:
        if r.get('ref_unk'):
            bad.append({'item': r['item'], 'part': r['part'], 'qty': r['qty'],
                        'unk': r['ref_unk'].split(','),
                        'all': len(str(r['ref']).split(',')) if r['ref'] else 0})
        if r.get('ref'):
            allrefs |= {x.strip().upper() for x in str(r['ref']).split(',') if x.strip()}
    return {
        'bad': bad,
        'n_unknown': sum(len(b['unk']) for b in bad),
        'extra': sorted(set(sides) - allrefs),
        'top': sum(1 for v in sides.values() if v == 'TOP'),
        'bot': sum(1 for v in sides.values() if v == 'BOT'),
    }


# ════════════════════════════════════════════════════════════════════════════════
# 7. v0.09 新增：BOM 列建構（替代料拆列）
# ════════════════════════════════════════════════════════════════════════════════

def build_rows(items: list, db: pd.DataFrame, knowledge: list = None,
               split_alt: bool = True, sides: dict = None,
               diag: dict = None) -> list:
    """
    把解析後的 BOM items + 零件DB + Knowledge 轉成「可直接寫入 Excel 的列」。

    v0.09：一列只有一個 P/N。
      有替代料時（find_matches 回傳 2 筆以上）拆成多列：
        主列   item=序號   qty=數量  ref=位置   pn=主料
        續列   item=None   qty=None  ref=None   pn=替代料   alt=True
      Total Qty 只加總主列，不會因拆列而重複計算。

    回傳 list[dict]，欄位見 ROW_FIELDS。
    """
    idx = build_db_index(db)
    if diag is not None:
        diag.clear()

    rows = []
    for seq, item in enumerate(items, 1):
        cands, nc = match_item(item, db, idx)
        if diag is not None:
            why = diagnose(item['part'], item['fp'], idx, bool(cands[0][0]))
            diag.setdefault(why, []).append(
                (seq, item['part'], item['fp'], pkg_of(item['fp'])))
        if not split_alt:
            cands = [(';'.join(p for p, _ in cands if p), cands[0][1])]

        pn_all = ';'.join(p for p, _ in cands if p)
        rcec   = classify_rcec(pn_all, cands[0][1], item['fp'])
        group  = len(cands) > 1

        # v0.09 雙面板：把 Location 依 Top/Bot 拆開
        refs = [r.strip() for r in str(item['ref']).split(',') if r.strip()]
        if sides:
            r_top = [r for r in refs if sides.get(r.upper()) == 'TOP']
            r_bot = [r for r in refs if sides.get(r.upper()) == 'BOT']
            r_unk = [r for r in refs if r.upper() not in sides]
        else:
            r_top, r_bot, r_unk = refs, [], []

        for i, (pn, desc) in enumerate(cands):
            d = desc + (' (原本為NC)' if (nc and i == 0) else '')
            rows.append({
                'ref_top': ','.join(r_top) if i == 0 else '',
                'ref_bot': ','.join(r_bot) if i == 0 else '',
                'ref_unk': ','.join(r_unk) if i == 0 else '',
                'qty_top': len(r_top) if i == 0 else None,
                'qty_bot': len(r_bot) if i == 0 else None,
                'item':  seq if i == 0 else None,
                'seq':   seq,                      # 群組編號（拆列共用）
                'rcec':  rcec,
                'pn':    pn,
                'desc':  d,
                'chk':   R8_MARK in d,          # R8：降壓選料，需人工確認
                'qty':   item['qty'] if i == 0 else None,
                'ref':   item['ref'] if i == 0 else '',
                'nc':    nc,
                'alt':   i > 0,                    # 是替代料續列
                'group': group,                    # 屬於有替代料的群組
                'fp':    item['fp'],
                'part':  item['part'],
                'kn':    find_knowledge_hits(pn, d, knowledge) if knowledge else '',
            })

    # ── PCB 列：固定最後一項 ────────────────────────────────────────────────
    rows.append({
        'item': len(items) + 1, 'seq': len(items) + 1, 'rcec': 'EC',
        'pn': '', 'desc': 'PCB', 'qty': 1, 'ref': '',
        'ref_top': '', 'ref_bot': '', 'ref_unk': '',
        'qty_top': None, 'qty_bot': None,
        'nc': False, 'alt': False, 'group': False, 'chk': False,
        'fp': '', 'part': 'PCB',
        'kn': find_knowledge_hits('', 'PCB', knowledge) if knowledge else '',
    })
    return rows


ROW_FIELDS = ('item', 'seq', 'rcec', 'pn', 'desc', 'qty', 'ref',
              'ref_top', 'ref_bot', 'ref_unk', 'qty_top', 'qty_bot',
              'nc', 'alt', 'group', 'chk', 'fp', 'part', 'kn')


def total_qty(rows: list) -> int:
    """只加總主列（qty 非 None），替代料續列不重複計。"""
    return sum(int(r['qty']) for r in rows if r.get('qty') not in (None, ''))


def stats(rows: list) -> dict:
    main = [r for r in rows if not r['alt']]
    return {
        'items':   len(main),
        'rows':    len(rows),
        'hit':     sum(1 for r in main if r['pn']),
        'miss':    sum(1 for r in main if not r['pn']),
        'nc':      sum(1 for r in main if r['nc']),
        'altgrp':  len({r['seq'] for r in rows if r['group']}),
        'kn':      sum(1 for r in rows if r['kn']),
        'qty':     total_qty(rows),
    }


# ════════════════════════════════════════════════════════════════════════════════
# 8. Style helpers
# ════════════════════════════════════════════════════════════════════════════════

def _sd(style='thin'):
    return Side(style=style, color='000000')

def _brd(l='thin', r='thin', t='thin', b='thin') -> Border:
    def s(x): return _sd(x) if x else Side()
    return Border(left=s(l), right=s(r), top=s(t), bottom=s(b))

def _aln(h='left', v='center', wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


B_OUTER  = _brd('medium', 'medium', 'thin', 'thin')
B_TOPCAP = _brd('medium', 'medium', 'medium', 'thin')
B_BOTCAP = _brd('medium', 'medium', 'thin', 'medium')
B_BODY_L = _brd('medium', 'medium', 'thin', 'thin')
B_FOOT   = _brd('medium', 'medium', 'thin', 'thin')


def _set_box(ws, rng: str, value=None, font=None, align=None,
             fill=None, border: Border = None):
    """
    Merge cells (if range), distribute border correctly across all edge
    cells, and write value/font/align/fill to the anchor cell.
    """
    c1, r1, c2, r2 = range_boundaries(rng)
    if (c1, r1) != (c2, r2):
        ws.merge_cells(rng)

    if border is not None:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).border = Border(
                    left  =border.left   if c == c1 else Side(),
                    right =border.right  if c == c2 else Side(),
                    top   =border.top    if r == r1 else Side(),
                    bottom=border.bottom if r == r2 else Side(),
                )
    if fill is not None:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).fill = fill

    anchor = ws.cell(r1, c1)
    if value is not None: anchor.value     = value
    if font  is not None: anchor.font      = font
    if align is not None: anchor.alignment = align
    return anchor


# ════════════════════════════════════════════════════════════════════════════════
# 9. Header (rows 1-8)
# ════════════════════════════════════════════════════════════════════════════════

HEADER_FIELDS = ('pcb_no', 'serial_no', 'pcb_pn', 'issue_date',
                 'rcec_no', 'rev', 'pcb_vendor', 'pcb_spec', 'ic_type')

def empty_header() -> dict:
    d = {k: '' for k in HEADER_FIELDS}
    d['rev'] = '0'
    return d


def _build_header(ws, hi: dict = None):
    hi = {**empty_header(), **(hi or {})}
    FA = lambda sz, bold=False: Font(name='Arial',  size=sz, bold=bold)
    FK = lambda sz:             Font(name='標楷體', size=sz)
    R, L, C = _aln('right'), _aln('left'), _aln('center')

    ws.merge_cells('A1:O1')
    _set_box(ws, 'A1', 'ADATA Technology Co., Ltd.', FA(14, True), L)
    ws.row_dimensions[1].height = 17.4

    ws.merge_cells('A2:L2')
    _set_box(ws, 'A2', 'M.2 SSD E-BOM', FA(24, True), C)
    ws.row_dimensions[2].height = 30.6

    ws.row_dimensions[3].height = 19.8
    _set_box(ws, 'A3:C3', 'PCB板號：',  FA(14), R, border=B_TOPCAP)
    _set_box(ws, 'D3:J3', hi['pcb_no'] or None,    FA(14), L, border=B_TOPCAP)
    _set_box(ws, 'K3:L3', '發行序號：', FK(14), R, border=B_TOPCAP)
    _set_box(ws, 'M3:O3', hi['serial_no'] or None, FA(14), L, border=B_TOPCAP)

    std = [
        (4, 19.8, 'PCB料號：',   hi['pcb_pn'],     '發行(改版)日期：', hi['issue_date']),
        (5, 39.6, 'RC/EC NO.：', hi['rcec_no'],    '目前版次：',        hi['rev']),
        (6, 19.8, 'PCB供應商：', hi['pcb_vendor'], None,                None),
    ]
    for row, h, lbl_ac, val_dj, lbl_kl, val_mo in std:
        ws.row_dimensions[row].height = h
        _set_box(ws, f'A{row}:C{row}', lbl_ac,       FA(14), R, border=B_OUTER)
        _set_box(ws, f'D{row}:J{row}', val_dj or None, FA(14), L, border=B_OUTER)
        _set_box(ws, f'K{row}:L{row}', lbl_kl,       FK(14), R, border=B_OUTER)
        _set_box(ws, f'M{row}:O{row}', val_mo or None, FA(14), L, border=B_OUTER)

    ws.row_dimensions[7].height = 19.8
    ws.row_dimensions[8].height = 20.4
    _set_box(ws, 'A7:C7', 'PCB 規格：',        FA(14), R, border=B_OUTER)
    _set_box(ws, 'D7:J7', hi['pcb_spec'] or None, FA(14), L, border=B_OUTER)
    _set_box(ws, 'A8:C8', 'Support IC Type：', FA(14), R, border=B_BOTCAP)
    _set_box(ws, 'D8:J8', hi['ic_type'] or None,  FA(14), L, border=B_BOTCAP)

    box78 = _brd('medium', 'medium', 'thin', 'medium')
    _set_box(ws, 'K7:L8', '環保等級：', FK(14), R, border=box78)
    _set_box(ws, 'M7:O8',
             '□A級 □S級 ■H級 □G級\n□OTHER 特殊客戶要求  :  ',
             FK(14), _aln('left', 'center', True), border=box78)


# ════════════════════════════════════════════════════════════════════════════════
# 10. PCB image (rows 9-35)
# ════════════════════════════════════════════════════════════════════════════════

def _place_img(ws, path, anchor, target_h, label, log):
    p = Path(path)
    if not p.exists():
        log(f"[警告] 找不到 PCB 圖片: {path}")
        return
    try:
        img = XLImage(str(p))
        scale = target_h / img.height
        w = int(img.width * scale)
        img.height, img.width = target_h, w
        ws.add_image(img, anchor)
        log(f"✓ PCB 圖片已嵌入（{label}）: {p.name}  ({w}×{target_h}px)")
    except Exception as e:
        log(f"[警告] PCB 圖片嵌入失敗（{label}）: {e}")


def _add_pcb_image(ws, img_path, log=print, img_bot=None):
    """
    圖片區 rows 9-35（27 列）。

    單面板：一張圖佔滿整個區域（高 700px）—— 與 v0.08 完全相同
    雙面板：切上下兩半，TOP 放 rows 9-21、BOT 放 rows 22-35，各高 340px
    判斷依據＝有沒有傳 img_bot 進來（也就是有沒有載入 Location 對照表）
    """
    for r in range(9, 36):
        ws.row_dimensions[r].height = 28.2
        _set_box(ws, f'A{r}', None, border=_brd('medium', None, None, None))
        _set_box(ws, f'O{r}', None, border=_brd(None, 'medium', None, None))

    if img_bot:
        FL = Font(name='Arial', size=9, bold=True, color='4F46E5')
        _set_box(ws, 'A9',  'TOP', FL, _aln('left'),
                 border=_brd('medium', None, None, None))
        _set_box(ws, 'A22', 'BOT', FL, _aln('left'),
                 border=_brd('medium', None, 'thin', None))
        for c in range(2, 16):                       # 兩半之間畫一條分隔線
            cell = ws.cell(22, c)
            cell.border = Border(left=cell.border.left, right=cell.border.right,
                                 top=_sd('thin'), bottom=cell.border.bottom)
        if img_path:
            _place_img(ws, img_path, 'A9', 340, 'TOP', log)
        _place_img(ws, img_bot, 'A22', 340, 'BOT', log)
        return

    if not img_path:
        return
    _place_img(ws, img_path, 'A9', 700, '單面', log)


def _warn_images(double: bool, img_path, img_bot, log):
    """
    v0.10 T4：圖片設定的兩種狀況都要講清楚，而且不替使用者決定。

    檔名含 top / bot 只是「自動偵測」的規則，不是必要條件 ——
    使用者手動指定任何兩個檔案都算數，程式照用。
    """
    if double and not img_bot:
        log('[提醒] 這是雙面板，但只指定了一張 PCB 圖片，圖片區維持整塊不切半。')
        log('       自動偵測要兩張檔名含 top / bot 的圖（例：pcb_top.jpg / pcb_bot.jpg）；')
        log('       手動指定任何兩個檔案也可以，不受檔名限制。')
        return
    if img_path and img_bot and Path(img_path).resolve() == Path(img_bot).resolve():
        log(f'[提醒] Top 與 Bot 指到同一張圖片（{Path(img_bot).name}），'
            f'上下兩半會是同一張。')
        log('       確認無誤才繼續；若是選錯，請重新指定 Bot 面的圖。')


# ════════════════════════════════════════════════════════════════════════════════
# 11. BOM table
# ════════════════════════════════════════════════════════════════════════════════

def _build_bom_header(ws, row: int, double: bool = False):
    """
    Row 36 欄標題。

    單面板（v0.06 起）  F=Location-Top Layer，G:O=Knowledge Note
    雙面板（v0.09 新增）F=Location-Top Layer，G=Location-Bottom Layer，
                        H:O=Knowledge Note
    雙面板就是 ADATA 範本第 36 列的原始設計 —— v0.06 當時是單面板，
    才把 Location-Bottom Layer 改成 Knowledge Note。欄數維持 15（A:O），
    表頭、PCB 圖片區、外框線都不用動。
    """
    F  = Font(name='Arial', size=12, bold=True)
    BC = _aln('center')
    B  = _brd('medium', 'medium', 'medium', 'thin')
    for col, val in [('A', 'Item'), ('B', 'RC/EC'), ('C', 'P/N'),
                     ('D', 'Description'), ('E', 'Qty'),
                     ('F', 'Location-Top Layer')]:
        _set_box(ws, f'{col}{row}', val, F, BC, border=B)
    if double:
        _set_box(ws, f'G{row}', 'Location-Bottom Layer', F, BC, border=B)
        _set_box(ws, f'H{row}:O{row}', 'Knowledge Note', F, BC, border=B)
    else:
        _set_box(ws, f'G{row}:O{row}', 'Knowledge Note', F, BC, border=B)
    ws.row_dimensions[row].height = 15.0


FL_NC  = PatternFill('solid', start_color='FFF2CC')   # yellow  = NC
FL_ALT = PatternFill('solid', start_color='E2EFDA')   # green   = alt P/N group
FL_KN  = PatternFill('solid', start_color='FFE0B2')   # orange  = knowledge hit
FL_TOT = PatternFill('solid', start_color='D9E1F2')   # blue    = total
FL_BAD = PatternFill('solid', start_color='FDE1E1')   # red     = Location 查不到 Top/Bot
FL_CHK = PatternFill('solid', start_color='EDE9FE')   # purple  = R8 降壓選料，需人工確認


def _write_data_row(ws, row: int, r: dict, double: bool = False):
    F_A = Font(name='Arial', size=14)
    F_B = Font(name='Arial', size=12)
    F_N = Font(name='Arial', size=10)
    F_R = Font(name='Arial', size=10, color='C00000')

    kn_note = r.get('kn') or ''
    unk     = r.get('ref_unk') or ''
    # 填色優先序：Location 查不到 Top/Bot（紅） > R8 降壓選料（紫） >
    #             Knowledge（橘） > NC（黃） > 替代料（綠）
    fill = (FL_BAD if (double and unk)
            else FL_CHK if r.get('chk')
            else FL_KN  if kn_note
            else FL_NC  if r.get('nc')
            else FL_ALT if r.get('group')
            else None)

    qty = r.get('qty')
    if double:
        f_val = tag_side(r.get('ref_top'), 'T')
        g_val = tag_side(r.get('ref_bot'), 'B')
    else:
        f_val, g_val = r.get('ref') or '', None

    data = [
        ('A', r.get('item'),           F_A, _aln('center')),
        ('B', r.get('rcec') or '',     F_B, _aln('center')),
        ('C', r.get('pn') or '',       F_B, _aln('left')),
        ('D', r.get('desc') or '',     F_B, _aln('left', wrap=True)),
        ('E', qty if qty not in (None, '') else None, F_B, _aln('center')),
        ('F', f_val,                   F_B, _aln('left', wrap=True)),
    ]
    for col, val, fnt, aln in data:
        _set_box(ws, f'{col}{row}', val, fnt, aln, fill=fill, border=B_BODY_L)

    lines = max(1, kn_note.count('\n') + 1) if kn_note else 1
    if double:
        _set_box(ws, f'G{row}', g_val, F_B, _aln('left', wrap=True),
                 fill=fill, border=B_BODY_L)
        note = kn_note
        if unk:
            warn = f'⚠ {len(unk.split(","))} 個 Location 查不到 Top/Bot：{unk}'
            note = (note + '\n' + warn) if note else warn
            lines = max(lines, note.count('\n') + 1)
        _set_box(ws, f'H{row}:O{row}', note or None,
                 F_R if (unk and not kn_note) else F_N,
                 _aln('left', wrap=True), fill=fill, border=B_BODY_L)
    else:
        _set_box(ws, f'G{row}:O{row}', kn_note or None, F_N,
                 _aln('left', wrap=True), fill=fill, border=B_BODY_L)

    ws.row_dimensions[row].height = max(17.4, lines * 18.0)


def _write_total_row(ws, row: int, tq: int):
    F = Font(name='Arial', size=12, bold=True)
    _set_box(ws, f'A{row}:B{row}', 'Total', F, _aln('center'),
             fill=FL_TOT, border=B_BODY_L)
    _set_box(ws, f'C{row}', None, F, _aln('center'), fill=FL_TOT, border=B_BODY_L)
    _set_box(ws, f'D{row}', None, F, _aln('center'), fill=FL_TOT, border=B_BODY_L)
    _set_box(ws, f'E{row}', tq, F, _aln('center'), fill=FL_TOT, border=B_BODY_L)
    _set_box(ws, f'F{row}:O{row}', None, F, _aln('center'),
             fill=FL_TOT, border=B_BODY_L)
    ws.row_dimensions[row].height = 17.4


def _write_footer(ws, row: int, double: bool = False, chk: bool = False):
    F_SM  = Font(name='Arial', size=10)
    F_LBL = Font(name='Arial', size=11)

    row += 1
    _set_box(ws, f'A{row}:O{row}', None, border=B_FOOT)

    row += 1
    _set_box(ws, f'A{row}:O{row}',
             'Remark：Top Layer only BOM. NC items listed for reference. '
             '替代料已拆列，每列僅含一個 P/N；Qty 只記於主列，不重複計算。',
             F_SM, _aln('left'), border=B_FOOT)
    ws.row_dimensions[row].height = 14.0

    row += 1
    _set_box(ws, f'A{row}:O{row}', None, border=B_FOOT)

    row += 1
    _set_box(ws, f'A{row}', '圖例',
             Font(name='Arial', bold=True, size=9), border=B_FOOT)
    _set_box(ws, f'B{row}:O{row}', None, border=B_FOOT)

    legends = [
        ('黃底', FL_NC,  '原本為 NC 的零件'),
        ('綠底', FL_ALT, '含替代料（已拆列，Item/Qty/Location 只記於主列）'),
        ('橘底', FL_KN,  'Knowledge 命中（含注意事項）'),
    ]
    if chk:
        legends.append(
            ('紫底', FL_CHK, 'R8 降壓選料：DB 耐壓低於原始設計，Description 已註記，請人工確認'))
    if double:
        legends.append(
            ('紅底', FL_BAD, '有 Location 在對照表裡查不到 Top/Bot（請確認兩份檔案是否同版）'))
    for i, (label, fl, note) in enumerate(legends):
        r = row + 1 + i
        _set_box(ws, f'A{r}', label,
                 Font(name='Arial', size=9), fill=fl, border=B_FOOT)
        _set_box(ws, f'B{r}', note,
                 Font(name='Arial', size=9), border=B_FOOT)
        _set_box(ws, f'C{r}:O{r}', None, border=B_FOOT)
    row += len(legends)

    row += 1
    _set_box(ws, f'A{row}:O{row}', None, border=B_FOOT)   # blank divider
    row += 1
    _set_box(ws, f'A{row}:B{row+1}', '    Approval  :  ', F_LBL,
             _aln('center'), border=B_FOOT)
    _set_box(ws, f'C{row}:D{row+1}', None, border=B_FOOT)
    _set_box(ws, f'E{row}:F{row+1}', 'Rechecked: ', F_LBL,
             _aln('center'), border=B_FOOT)
    _set_box(ws, f'G{row}:H{row+1}', 'Prepared: ', F_LBL,
             _aln('center'), border=B_FOOT)
    _set_box(ws, f'I{row}:J{row+1}', None, border=B_FOOT)
    _set_box(ws, f'K{row}:O{row+1}', None, border=B_FOOT)

    last = row + 1
    for col in range(1, 16):
        c = ws.cell(last, col)
        c.border = Border(left=c.border.left, right=c.border.right,
                          top=c.border.top, bottom=_sd('medium'))


def _set_col_widths(ws, double: bool = False):
    w = {'A': 5.1,  'B': 10.1, 'C': 29.1, 'D': 77.8, 'E': 9.2,
         'F': 54.3, 'G': 11.0, 'H': 16.0, 'I': 13.0, 'J': 13.0,
         'K':  9.0, 'L': 12.0, 'M': 13.3, 'N':  9.0, 'O': 13.0}
    if double:
        # G 從 Knowledge Note 的第一欄變成 Location-Bottom，要給它跟 F 一樣寬
        w['F'] = w['G'] = 38.0
    for col, v in w.items():
        ws.column_dimensions[col].width = v


# ════════════════════════════════════════════════════════════════════════════════
# 12. Main write function
# ════════════════════════════════════════════════════════════════════════════════

BOM_HDR = 36     # BOM 欄標題所在列（沿用 ADATA E-BOM 範本）


def force_xlsx(path: str) -> str:
    """
    強制把輸出副檔名正規化成 .xlsx。

    v0.09 修正：使用者在輸出路徑打 .xls，openpyxl 仍然寫出 OOXML 格式，
    產生「內容是 xlsx、副檔名是 xls」的檔案。Excel 會跳「格式與副檔名
    不符」警告，openpyxl / pandas 則直接拒絕開啟：
        InvalidFileException: openpyxl does not support the old .xls format
    實際事故：output/SM8008/SM8008_adata.xls 無法被任何工具讀取。
    """
    p = Path(path)
    if p.suffix.lower() != '.xlsx':
        p = p.with_suffix('.xlsx')
    return str(p)


def write_excel(rows: list, out_path: str, pcb_img: str = None,
                header_info: dict = None, log=print, pcb_bot: str = None):
    """
    v0.09：輸入已建構好的 rows（可先經 GUI 人工修正），輸出 Excel。

    雙面板模式的判斷：rows 裡有任何一列帶 ref_top/ref_bot 資料就算雙面板。
    沒有 Location 對照表時 build_rows 不會填這兩欄，輸出與 v0.08 完全相同。
    """
    double = any(r.get('ref_bot') or r.get('ref_unk') for r in rows)
    fixed = force_xlsx(out_path)
    if fixed != out_path:
        log(f"[提醒] 輸出副檔名已自動更正為 .xlsx：{Path(fixed).name}")
    out_path = fixed
    if Path(out_path).is_file():
        log(f"[警告] 覆蓋既有檔案：{Path(out_path).name}")

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = 'BOM'
    ws.sheet_view.showGridLines = False

    _set_col_widths(ws, double)
    _build_header(ws, header_info)
    _add_pcb_image(ws, pcb_img, log=log, img_bot=pcb_bot if double else None)
    _warn_images(double, pcb_img, pcb_bot, log)
    _build_bom_header(ws, BOM_HDR, double)

    r = BOM_HDR + 1
    for row in rows:
        _write_data_row(ws, r, row, double)
        r += 1

    tq = total_qty(rows)
    _write_total_row(ws, r, tq)
    _write_footer(ws, r, double, any(x.get('chk') for x in rows))

    ws.freeze_panes = None

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    s = stats(rows)
    log(f"✓ 輸出: {out_path}（{'雙面板' if double else '單面板'}）")
    log(f"  {s['items']} 項（含PCB）→ {s['rows']} 列"
        f"（替代料拆出 {s['rows'] - s['items']} 列）  |  Total Qty = {tq}")
    return s


# ════════════════════════════════════════════════════════════════════════════════
# 13. Auto-detect（v0.09：以程式所在目錄為基準，含子資料夾）
# ════════════════════════════════════════════════════════════════════════════════

def _newest(paths: list) -> list:
    """
    依最後修改時間由新到舊排序。

    v0.09 修正：v0.07 的 discover() 取 txts[-1] 當預設值，而候選清單的順序
    是「先 glob *.txt 再 glob *.BOM」—— .BOM 檔永遠排在最後、永遠被選中，
    跟檔名和日期都無關。實際事故：2026-10-14 的 SM8008 選不到，反而選到
    2026-09-01 的舊範例檔，導致 C 案產出的是錯的 BOM。
    改為一律依 mtime 排序，最新的排第一個，由 discover() 取 [0]。
    """
    def mtime(p):
        try:
            return p.stat().st_mtime
        except OSError:
            return 0.0
    return sorted(paths, key=mtime, reverse=True)


def _dirs(base: Path, subs) -> list:
    """base + 存在的子資料夾。Windows 不分大小寫，需去重避免同一夾掃兩次。"""
    out, seen = [], set()
    for p in [base] + [base / s for s in subs]:
        if not p.is_dir():
            continue
        k = str(p.resolve()).lower()
        if k not in seen:
            seen.add(k); out.append(p)
    return out


def is_bom_txt(path) -> bool:
    """
    必須是 Altium 匯出的 Text BOM：表頭列以 TAB 分隔，且欄位中同時有
    Item / Quantity / Reference。

    v0.09 修正一：只用副檔名會把 changelog_*.txt 當成 BOM 檔。
    v0.09 修正二：只用「文字包含」仍會誤判 —— 本專案的 README 與 changelog
                  裡就寫著「內容須含 Bill Of Materials 或 Item/Quantity/
                  Reference」這句說明，結果說明文件自己命中了偵測規則。
                  改成要求 TAB 分隔的表頭列（與 parse_text_bom 實際依據一致）。
    """
    try:
        with open(path, encoding='utf-8', errors='replace') as f:
            head = f.read(8192)
    except Exception:
        return False
    for ln in head.splitlines():
        cols = [c.strip() for c in ln.split('\t')]
        if len(cols) >= 3 and 'Item' in cols and 'Quantity' in cols \
                and 'Reference' in cols:
            return True
    return False


# Altium 匯出的副檔名可能是 .txt 或 .BOM（Windows glob 不分大小寫）
_BOM_EXTS = ('*.txt', '*.BOM')


def find_bom_txts(base: Path = APP_DIR) -> list:
    found = []
    for d in _dirs(base, SUB_TXT):
        cands = []
        for ext in _BOM_EXTS:
            cands += sorted(d.glob(ext))
        found += [p for p in cands if is_bom_txt(p)]
    # 去重（Windows 不分大小寫）
    uniq, seen = [], set()
    for p in found:
        k = str(p).lower()
        if k not in seen:
            seen.add(k); uniq.append(p)
    return _newest(uniq)


def find_parts_dbs(base: Path = APP_DIR) -> list:
    found = []
    for d in _dirs(base, SUB_DB):
        for p in sorted(d.glob('*.xlsx')):
            if p.name.startswith('~$'):
                continue
            if is_parts_db(p):
                found.append(p)
    return _newest(found)


def find_knowledge_dbs(base: Path = APP_DIR) -> list:
    """只認表頭（項次|PN|品名|包裝|note），不認檔名 —— 見 is_knowledge_db。"""
    found = []
    for d in _dirs(base, SUB_DB):
        for p in sorted(d.glob('*.xlsx')):
            if p.name.startswith('~$'):
                continue
            if is_knowledge_db(p):
                found.append(p)
    return _newest(found)


_RE_TOP_IMG = re.compile(r'(^|[_\-])(top|t)([_\-]|$)', re.IGNORECASE)
_RE_BOT_IMG = re.compile(r'(^|[_\-])(bot|bottom|b)([_\-]|$)', re.IGNORECASE)


def find_pcb_pair(base: Path = APP_DIR) -> tuple:
    """
    雙面板用：找 TOP / BOT 兩張圖。
    檔名含 top / bot（可用 _ 或 - 分隔），例如
        pcb_top.jpg / pcb_bot.jpg、SM8008-TOP.png / SM8008-BOT.png
    兩張都找到才算數，否則回 ('', '') 讓程式走單面板。
    """
    top = bot = ''
    for p in find_pcb_imgs(base):
        s = Path(p).stem
        if not top and _RE_TOP_IMG.search(s):
            top = str(p)
        elif not bot and _RE_BOT_IMG.search(s):
            bot = str(p)
    return (top, bot) if (top and bot) else ('', '')


def find_side_tables(base: Path = APP_DIR) -> list:
    """Location TOP/BOT 對照表：xlsx 優先（必要格式），txt 其次（選用）。"""
    xs, ts = [], []
    for d in _dirs(base, SUB_DB + SUB_TXT):
        for p in sorted(d.glob('*.xlsx')):
            if not p.name.startswith('~$') and is_side_table(p):
                xs.append(p)
        for p in sorted(d.glob('*.txt')):
            if is_side_table(p):
                ts.append(p)
    return _newest(xs) + _newest(ts)


def find_pcb_imgs(base: Path = APP_DIR) -> list:
    found = []
    for d in _dirs(base, SUB_IMG):
        for ext in ('*.jpg', '*.jpeg', '*.png'):
            found += sorted(d.glob(ext))
    uniq, seen = [], set()
    for p in found:
        k = str(p).lower()
        if k not in seen:
            seen.add(k); uniq.append(p)
    # pcb.* 優先，其餘依最後修改時間由新到舊
    uniq = _newest(uniq)
    uniq.sort(key=lambda p: 0 if p.stem.lower() == 'pcb' else 1)
    return uniq


def output_dir(base: Path = APP_DIR) -> Path:
    for s in SUB_OUT:
        p = base / s
        if p.is_dir():
            return p
    return base


def default_out_name(bom_txt: str, base: Path = APP_DIR) -> str:
    stem = Path(bom_txt).stem if bom_txt else 'BOM'
    stem = re.sub(r'\.BOM$', '', stem, flags=re.I)   # xxx.BOM.txt → xxx
    return str(output_dir(base) / f'BOM_{stem}_{date.today():%Y%m%d}.xlsx')


def discover(base: Path = APP_DIR) -> dict:
    """一次回傳所有自動偵測結果（給三個前端共用）。"""
    txts = find_bom_txts(base)
    dbs  = find_parts_dbs(base)
    kns  = find_knowledge_dbs(base)
    imgs = find_pcb_imgs(base)
    # v0.09：候選清單一律「最新的排第一個」，預設值取 [0]
    sel_txt = str(txts[0]) if txts else ''
    sides = find_side_tables(base)
    ptop, pbot = find_pcb_pair(base)
    return {
        'base':      str(base),
        'txts':      [str(p) for p in txts],
        'dbs':       [str(p) for p in dbs],
        'kns':       [str(p) for p in kns],
        'imgs':      [str(p) for p in imgs],
        'sides_all': [str(p) for p in sides],
        'sides':     str(sides[0]) if sides else '',
        'pcb_top':   ptop,
        'pcb_bot':   pbot,
        'bom':       sel_txt,
        'db':        str(dbs[0])   if dbs  else '',
        'knowledge': str(kns[0])   if kns  else '',
        'pcb':       ptop or (str(imgs[0]) if imgs else ''),
        'out_dir':   str(output_dir(base)),
        'out':       default_out_name(sel_txt, base),
    }


# ════════════════════════════════════════════════════════════════════════════════
# 14. 一次跑完（給 GUI 呼叫）
# ════════════════════════════════════════════════════════════════════════════════

def run_match(bom_txt: str, db_path: str, kn_path: str = '',
              split_alt: bool = True, log=print, sides_path: str = '') -> tuple:
    """回傳 (rows, items, db_df, knowledge)。"""
    items = parse_text_bom(bom_txt)
    log(f'✓ 解析 Text BOM: {len(items)} 項')
    db = read_db(db_path)
    log(f'✓ 讀取零件資料庫: {len(db)} 筆')
    kn = None
    if kn_path and Path(kn_path).exists():
        kn = load_knowledge(kn_path)
        log(f'✓ 載入 Knowledge: {len(kn)} 筆')

    sides = None
    if sides_path and Path(sides_path).exists():
        sides = load_sides(sides_path)
        t = sum(1 for v in sides.values() if v == 'TOP')
        log(f'✓ 載入 Location 對照表: {len(sides)} 筆'
            f'（Top {t} / Bot {len(sides) - t}）→ 雙面板模式')

    diag = {}
    rows = build_rows(items, db, kn, split_alt=split_alt, sides=sides, diag=diag)
    s = stats(rows)
    log(f'✓ 比對完成: 命中 {s["hit"]} / 未命中 {s["miss"]} / '
        f'替代料群組 {s["altgrp"]} / Knowledge 命中 {s["kn"]} 列')
    report_r8(rows, log)
    report_diagnosis(diag, log)
    if sides:
        report_sides(rows, sides, log)
    return rows, items, db, kn


def report_r8(rows: list, log=print):
    """R8：降壓選料一律列出來，不能靜默帶過。"""
    bad = [r for r in rows if r.get('chk')]
    if not bad:
        return
    log(f'  [請人工確認] {len(bad)} 項電容採用「耐壓低於原始設計」的替代料：')
    for r in bad:
        log(f'           Item {r["seq"]:<4} {r["part"]:<18} -> {r["pn"]:<12} {r["desc"]}')
    log('           這些列在 Excel 標紫底，Description 已註記原始設計耐壓。')


def report_diagnosis(diag: dict, log=print):
    """
    R7：未命中一律歸因並印出。

    v0.08 比不到就靜默留白，所以 Part 欄位格式變了沒人發現 ——
    SM8008 的 32 項電容全部落空，卻要到人工核對產出才看得出來。
    """
    if not diag:
        return
    order = [D_HIT, D_NOVAL, D_SPEC, D_PKG, D_NOVALUE_DB, D_NOTPARAM]
    log('  未命中原因：')
    for k in order:
        v = diag.get(k)
        if not v or k == D_HIT:
            continue
        log(f'      {k:<16} {len(v):>3} 項'
            + (f'   例：{v[0][1]}' if v else ''))

    # 主值解析失敗＝格式可能又變了，要特別吵
    bad = diag.get(D_NOVAL, [])
    if bad:
        log(f'  [警告] 有 {len(bad)} 項「主值解析失敗」，'
            f'Part 欄位格式可能與預期不同，請檢查：')
        for seq, part, fp, _ in bad[:6]:
            log(f'           Item {seq:<4} {part!r:<28} fp={fp}')
        if len(bad) > 6:
            log(f'           …另有 {len(bad) - 6} 項')

    # 整類元件命中率 0% 的自我檢查
    byclass = {}
    for why, lst in diag.items():
        for seq, part, fp, _ in lst:
            k = (fp or ' ')[0].upper()
            k = k if k in 'CRL' else '其他'
            a, b = byclass.get(k, (0, 0))
            byclass[k] = (a + (1 if why == D_HIT else 0), b + 1)
    for k, (hit, tot) in sorted(byclass.items()):
        if tot >= 5 and hit == 0:
            log(f'  [警告] {k} 類 {tot} 項「全部未命中」，'
                f'八成不是 DB 沒料，而是比對規則對不上這份 BOM 的格式。')


def report_sides(rows: list, sides: dict, log=print):
    """雙面板：Location 對不上的三層處理之一 —— 日誌。"""
    rep = side_report(rows, sides)
    if rep['bad']:
        log(f'  [警告] 有 {rep["n_unknown"]} 個 Location 在對照表裡查不到 '
            f'Top/Bot，受影響的項目：')
        for b in rep['bad']:
            full = '　★ 這一項全部查不到，Top/Bot 都會是 0' \
                if len(b['unk']) == b['all'] else ''
            log(f'           Item {b["item"]:<4} {b["part"]:<18} '
                f'Qty {b["qty"]}  ->  {len(b["unk"])} 個查不到'
                f'（{",".join(b["unk"])}）{full}')
        log('           這些 Location 不會出現在 Top 或 Bot 欄，'
            '請確認兩份檔案是不是同一版。')
    if rep['extra']:
        log(f'  [提示] 對照表有 {len(rep["extra"])} 個 Location 在 BOM 裡找不到'
            f'（通常是不上件或改版殘留）：{",".join(rep["extra"][:20])}')


# ════════════════════════════════════════════════════════════════════════════════
# 15. CLI
# ════════════════════════════════════════════════════════════════════════════════

def _cli(argv=None):
    ap = argparse.ArgumentParser(
        description=f'BOM Matcher core v{VERSION}（工作目錄 = 程式所在目錄）')
    ap.add_argument('text_bom', nargs='?', help='未指定則自動偵測')
    ap.add_argument('excel_db', nargs='?', help='未指定則自動偵測')
    ap.add_argument('-o', '--output')
    ap.add_argument('-k', '--knowledge')
    ap.add_argument('-p', '--pcb', help='PCB 圖片（單面板）／TOP 面圖片')
    ap.add_argument('-b', '--pcb-bot', help='BOT 面圖片（雙面板）')
    ap.add_argument('-t', '--sides', help='Location TOP/BOT 對照表（xlsx 或 txt）')
    ap.add_argument('--single', action='store_true',
                    help='強制單面板，忽略 Location 對照表')
    ap.add_argument('--no-split', action='store_true',
                    help='關閉替代料拆列（回到 v0.06 的 a;b 單列格式）')
    args = ap.parse_args(argv)

    d = discover()
    bom = args.text_bom or d['bom']
    db  = args.excel_db or d['db']
    kn  = args.knowledge if args.knowledge is not None else d['knowledge']
    pcb = args.pcb       if args.pcb       is not None else d['pcb']
    sides = '' if args.single else (args.sides
                                    if args.sides is not None else d['sides'])
    pbot = args.pcb_bot if args.pcb_bot is not None else d['pcb_bot']
    out = args.output or default_out_name(bom)

    print(f'BOM Matcher v{VERSION}  (core / CLI)')
    print(f'  程式目錄   : {APP_DIR}')
    print(f'  Text BOM   : {bom or "（未找到）"}')
    print(f'  Excel DB   : {db  or "（未找到）"}')
    print(f'  Knowledge  : {kn  or "（未找到）"}')
    print(f'  PCB 圖片   : {pcb or "（未找到）"}')
    if sides:
        print(f'  BOT 面圖片 : {pbot or "（未找到，雙面板建議提供）"}')
        print(f'  Location表 : {sides}')
    print(f'  Output     : {out}')
    print(f'  面板       : {"雙面板" if sides else "單面板"}')
    print(f'  替代料拆列 : {"否" if args.no_split else "是"}\n')

    if not bom or not db:
        print('✗ 缺少 Text BOM 或零件資料庫，請用參數指定。')
        return 1

    rows, *_ = run_match(bom, db, kn, split_alt=not args.no_split,
                         sides_path=sides)
    write_excel(rows, out, pcb or None, pcb_bot=pbot or None)
    return 0


if __name__ == '__main__':
    sys.exit(_cli())
