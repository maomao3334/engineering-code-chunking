# -*- coding: utf-8 -*-
"""把附录A转录与原图 OCR 网格并排输出，人工逐页扫错。

OCR 虽错位，但**文字内容**大多识别对了，可以当"拼写参照"：
若我的转录里出现 OCR 完全没有的字词组合，可能是读错。
"""
import json
import re
from pathlib import Path

import argparse
_ap = argparse.ArgumentParser()
_ap.add_argument("--pipeline", required=True, help="_pipeline 目录")
_a = _ap.parse_args()
CASE = Path(_a.pipeline)
A = json.loads((CASE / "appendix_a.json").read_text(encoding="utf-8"))
ocr = json.loads((CASE / "tables_ocr.json").read_text(encoding="utf-8"))

pages = ["p27_1.png", "p28_1.png", "p28_2.png", "p29_1.png", "p29_2.png",
         "p30_1.png", "p30_2.png", "p31_1.png", "p31_2.png", "p32_1.png"]

# 收集所有 OCR 文字
alltext = set()
for t in ocr:
    if t["img"] in pages:
        for row in t["grid"]:
            for c in row:
                s = re.sub(r"\s", "", str(c or ""))
                if s:
                    alltext.add(s)
blob = "".join(alltext)

print("检查：我的转录里，有哪些『部位/危险物』的关键词在 OCR 全文中找不到")
print("（找不到 ≠ 一定错，OCR 可能漏识，但值得逐个看）")
print("=" * 80)
suspect = []
for x in A:
    for field in ("part", "hazard"):
        v = re.sub(r"[，、；：（）()\s]", "", x[field])
        # 取 2~4 字的片段检查
        for i in range(0, max(1, len(v) - 2), 2):
            seg = v[i:i + 3]
            if len(seg) < 2:
                continue
            if seg not in blob:
                suspect.append((x["factory"], x["part"][:20], x["hazard"][:16], seg))
                break

seen = set()
for fac, part, hz, seg in suspect:
    k = (part, seg)
    if k in seen:
        continue
    seen.add(k)
    print(f"  [{fac}] {part} | 危险物={hz} | OCR 无「{seg}」")

print(f"\n可疑 {len(seen)} 处（需人工看原图确认）")
