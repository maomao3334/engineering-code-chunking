# -*- coding: utf-8 -*-
"""系统性核对：视觉读图的值 vs OCR 网格的值（多重集比较）。

思路：OCR 网格列错位、行粘连，但**数字本身大多识别对了**。
所以先比多重集（每个数值出现的次数）：
    一致 → 数值层面没漏没多，差异只在"行列归属"
    不一致 → 指出哪些值多了/少了，回去核对这些格

这比逐格比对高效得多，也不会被列错位误导。

用法：
    python compare_vision_vs_ocr.py --pdf 规范.pdf [--table 表4.2.10]
"""
import argparse
import json
import re
from collections import Counter
from pathlib import Path

_ap = argparse.ArgumentParser()
_ap.add_argument("--pdf", required=True)
_ap.add_argument("--table", default="表4.2.10")
_a = _ap.parse_args()

CASE = Path(_a.pdf).parent / "_pipeline"
cal = json.loads((CASE / "table_columns_calibrated.json").read_text(encoding="utf-8"))
tabs = json.loads((CASE / "tables_ocr.json").read_text(encoding="utf-8"))

c = cal[_a.table]
mine = Counter()
for vals in c["grid_rows"].values():
    for v in vals:
        if v is not None:
            mine[str(v)] += 1

t = [x for x in tabs if x["page"] == 12][0]
ocr = Counter()
for row in t["grid"]:
    for v in row:
        s = str(v).strip()
        if re.fullmatch(r"\d+(?:\.\d+)?", s):
            ocr[s] += 1
        # 处理粘连："1212" → 拆成 12,12；"2515" → 25,15
        elif re.fullmatch(r"(\d{2})(\d{2})", s):
            ocr[s[:2]] += 1
            ocr[s[2:]] += 1
        elif re.fullmatch(r"(\d)(\d{2})", s):
            ocr[s[:1]] += 1
            ocr[s[1:]] += 1

print(f"我读图数值: {sum(mine.values())} 个，{len(mine)} 种")
print(f"OCR 数值  : {sum(ocr.values())} 个，{len(ocr)} 种")
print()

only_mine = mine - ocr
only_ocr = ocr - mine

print("只在读图结果里（OCR 漏识或识别错）：")
if only_mine:
    for v, n in sorted(only_mine.items(), key=lambda x: (len(x[0]), x[0])):
        print(f"   值 {v} × {n}")
else:
    print("   无")

print("\n只在 OCR 里（我可能漏读，或 OCR 粘连误拆）：")
if only_ocr:
    for v, n in sorted(only_ocr.items(), key=lambda x: (len(x[0]), x[0])):
        print(f"   值 {v} × {n}")
else:
    print("   无")

same = mine == ocr
print(f"\n多重集完全一致: {'是 ✓' if same else '否'}")
if not same:
    print("  → 差异只说明数值层面的出入，行列归属仍需靠视觉读图确认")
