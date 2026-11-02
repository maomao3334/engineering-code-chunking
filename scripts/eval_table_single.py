# -*- coding: utf-8 -*-
"""严格验证：不仅看"有没有这个数"，还要看"对象对不对"。

宽松判据会误判：问"明火与厂外道路"，匹配到"明火与码头 35m"也算过——
因为两者数值恰好都是 35。这种"值对但对象错"最危险。

严格判据：答案句里必须**同时**出现问题中的两个对象关键词，且数值正确。

用法：
    python eval_table_single.py --corpus 语料库.xlsx
    （不传则用环境变量 CORPUS_XLSX）
    下面 CASES 是针对 GB 50565-2010 表4.2.10 写的，换规范必须重写。
"""
import argparse
import json
import os
import re
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

_ap = argparse.ArgumentParser()
_ap.add_argument("--corpus", default=None, help="语料库 xlsx")
_a = _ap.parse_args()
if _a.corpus:
    os.environ["CORPUS_XLSX"] = str(Path(_a.corpus).resolve())

for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "NO_PROXY",
          "http_proxy", "https_proxy", "all_proxy", "no_proxy"):
    os.environ.pop(k, None)

from raglib import Retriever  # noqa: E402

r = Retriever(use_table_slices=True, use_expand=True, child_weight=0.8)

# (问题, 期望值, 对象A必须词, 对象B必须词)
CASES = [
    ("厂内主要道路与生产厂房（一、二级）甲类的防火间距？",
     "10", "厂内主要道路", "一、二级·甲类"),
    ("罐区甲、乙类泵或泵房与明火及散发火花地点的防火间距？",
     "30", "泵或泵房", "明火"),
    ("甲类仓库（储量≤10t）与厂内铁路的防火间距？",
     "30", "甲类仓库", "厂内铁路"),
    ("行政、生活建筑（三级）与码头的防火间距？",
     "40", "耐火等级三级", "码头"),
    ("生产厂房三级丁戊类与汽车装卸站的防火间距？",
     "15", "耐火等级三级·丁、戊类", "汽车装卸"),
    ("明火及散发火花地点与甲、乙类液体码头装卸区的防火间距？",
     "35", "明火", "码头装卸区"),
]

print(f"{'#':>2s} {'判据':6s} 说明")
print("-" * 90)
ok = 0
for i, (q, want, a, b) in enumerate(CASES, 1):
    hits = r.search_parent(q, k=5)
    ctx = "\n".join(h["context"] for h in hits)
    # 严格：同一句里必须同时有 a、b、want
    good = []
    for ln in ctx.split("\n"):
        if "防火间距不应小于" not in ln:
            continue
        if a in ln and b in ln and f" {want}m" in ln:
            good.append(ln.strip())
    # 反向检查：是否也出现了"同值但对象错"的句子（会造成误导）
    misleading = []
    for ln in ctx.split("\n"):
        if "防火间距不应小于" not in ln:
            continue
        if f" {want}m" in ln and a in ln and b not in ln:
            misleading.append(ln.strip())
    verdict = "✓" if good else "✗"
    ok += bool(good)
    print(f"{i:2d} {verdict:6s} {q[:36]}")
    if good:
        print(f"       正确句: {good[0][:110]}")
    else:
        print(f"       ✗ 无同时含「{a}」「{b}」且={want}m 的句子")
        if misleading:
            print(f"       ⚠ 但存在同值的其他对象句（会误导）: {misleading[0][:88]}")

print("-" * 90)
print(f"严格通过 {ok}/{len(CASES)}")
