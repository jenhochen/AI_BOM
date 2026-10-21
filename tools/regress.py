"""v0.09 vs v0.10 回歸：舊 BOM 的結果不可以退步。"""
import importlib.util
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
os.chdir(r"F:\AI_BOM")


def load(f, n):
    s = importlib.util.spec_from_file_location(n, f)
    m = importlib.util.module_from_spec(s)
    sys.modules[n] = m
    s.loader.exec_module(m)
    return m


v8 = load("old/bom_core_0.09.py", "v9")   # 舊版已搬到 old 目錄
v9 = load("bom_core_0.10.py", "v10")

CASES = [
    (r"TXT\S-M2P42280SMH086-20260812_1.BOM", r"ExcelDB\BOM_20260525.xlsx"),
    (r"TXT\S-SM2268XT_154BALL_20260515_FLASHX4_BOM.txt", r"ExcelDB\BOM_20260525.xlsx"),
    (r"TXT\SM8008_M2_D4X2_4G178_C261C_V1C_20261012.BOM.txt", r"ExcelDB\BOM_20261014.xlsx"),
]

print("%-42s %-14s %-14s %-12s %s" % ("BOM", "v0.09 命中", "v0.10 命中", "Total Qty", "判定"))
print("-" * 92)
bad = 0
for bom, db in CASES:
    r8 = v8.build_rows(v8.parse_text_bom(bom), v8.read_db(db), None, True)
    r9 = v9.build_rows(v9.parse_text_bom(bom), v9.read_db(db), None, True)
    s8, s9 = v8.stats(r8), v9.stats(r9)
    ok = s9["hit"] >= s8["hit"] and s9["qty"] == s8["qty"]
    bad += 0 if ok else 1
    print("%-42s %-14s %-14s %-12s %s" % (
        Path(bom).name[:40],
        "%d/%d" % (s8["hit"], s8["items"]),
        "%d/%d" % (s9["hit"], s9["items"]),
        "%s→%s" % (s8["qty"], s9["qty"]),
        "OK" if ok else "!! 退步"))
print("-" * 92)
print("退步案例:", bad)

print("\n=== SM8008 未命中歸因（R7）===")
bom, db = CASES[2]
diag = {}
v9.build_rows(v9.parse_text_bom(bom), v9.read_db(db), None, True, diag=diag)
for k, v in sorted(diag.items(), key=lambda x: -len(x[1])):
    print("   %-18s %3d 項" % (k, len(v)))
