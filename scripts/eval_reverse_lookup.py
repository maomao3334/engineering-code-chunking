# -*- coding: utf-8 -*-
"""验证枚举意图加权后，反向速查片能否进入 top-3。"""
import os
import sys

import argparse
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
_ap = argparse.ArgumentParser()
_ap.add_argument("--corpus", required=True, help="语料库 xlsx")
_a = _ap.parse_args()
os.environ["CORPUS_XLSX"] = str(Path(_a.corpus).resolve())
from raglib import Retriever  # noqa: E402

r = Retriever(use_table_slices=True, use_expand=True, child_weight=0.8)

QS = [
    ("哪些设施与厂内主要道路的防火间距是10m", "厂内主要道路"),
    ("哪些设施与厂内铁路的防火间距是10m", "厂内铁路"),
    ("与明火及散发火花地点间距为30m的有哪些", "明火及散发火花地点"),
]
print("枚举类问题：速查片是否进入 top-3")
print("=" * 74)
for q, key in QS:
    hits = r.search_parent(q, k=3)
    names = [h["parent"] for h in hits]
    # 判断 top-3 里有没有速查内容
    has = any("反向速查" in lv for h in hits for lv in h["hit_levels"])
    print(f"\n问：{q}")
    print(f"  top-3 条号: {names}   速查片命中: {'✓' if has else '✗'}")
    if has:
        for h in hits:
            if any("反向速查" in lv for lv in h["hit_levels"]):
                for ln in h["context"].split("\n"):
                    if key in ln and "={" in ln:
                        print(f"    → {ln.strip()[:120]}")
                break

print()
print("=" * 74)
print("普通定位问题：速查片不应干扰")
print("=" * 74)
for q, want in [("厂外铁路与甲、乙类可燃液体储罐（≤5000m³）的防火间距？", "35"),
                ("明火及散发火花地点与厂内主要道路的防火间距？", "10")]:
    hits = r.search_parent(q, k=3)
    ctx = "\n".join(h["context"] for h in hits)
    print(f"  {'✓' if f' {want}m' in ctx else '✗'} {q[:44]} → {want}m")
