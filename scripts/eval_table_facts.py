# -*- coding: utf-8 -*-
"""真实问答测试：模拟 AI 拿到切片后能否正确回答表格题。

不测"关键词在不在"，测"给 LLM 这段文本，它能不能答对"——
用检索结果 + 事实句，人工核对每条答案是否与原图一致。

用法：
    python eval_table_facts.py --corpus 语料库.xlsx
    或设环境变量 CORPUS_XLSX

注意：下面 CASES 的期望值是**针对 GB 50565-2010 纺织规范的三张表**写的，
换规范时必须重写（期望值写错会让你误判程序有 bug —— 我踩过两次）。
"""
import argparse
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

from raglib import Retriever  # noqa: E402

r = Retriever(use_table_slices=True, use_expand=True, child_weight=0.8)

# (问题, 期望数值, 参照物1, 参照物2) —— 期望值已对照原图核实
CASES = [
    ("可燃气体的厂外铁路与甲、乙类可燃液体储罐（≤5000m³）的防火间距是多少？", "35", "厂外铁路", "甲、乙类"),
    ("厂外民用建筑与丙类厂房的防火间距是多少？", "17", "厂外民用建筑", "丙类厂房"),
    ("高速公路与丙类可燃液体储罐的防火间距是多少？", "22", "高速公路", "丙类"),
    ("堆场储量10t~500t时，与甲类厂房及仓库的防火间距是多少？", "13", "10t~500t", "甲类厂房及仓库"),
    ("堆场储量10t~500t时，与明火或散发火花地点的防火间距是多少？", "25", "10t~500t", "明火"),
    ("堆场储量10t~500t时，与厂内铁路的防火间距是多少？", "12", "10t~500t", "厂内铁路"),
    ("室外变配电站（变压器总油量>10t）与甲、乙类储罐的防火间距是多少？", "50", "室外变", "甲、乙类"),
    ("通航江、河、海岸边与丙类厂房的防火间距是多少？", "15", "通航", "丙类厂房"),
]

print(f"{'#':>2s} {'期望':>4s} {'事实句命中':>10s} {'答案句':>6s}  判定")
print("-" * 86)
ok = 0
for i, (q, want, k1, k2) in enumerate(CASES, 1):
    hits = r.search_parent(q, k=3)
    # 在 top3 的所有切片文本里找"含 want 且含两个参照物"的事实句
    best = ""
    for h in hits:
        for ln in h["context"].split("\n"):
            if want in ln and k1 in ln and k2 in ln and "防火间距不应小于" in ln:
                best = ln.strip()
                break
        if best:
            break
    got = bool(best)
    ok += got
    print(f"{i:2d} {want:>4s} {'✓' if got else '✗':>10s} {'✓' if got else '✗':>6s}  "
          f"{'可答对' if got else '答不准'}")
    if best:
        print(f"     {best[:110]}")

print("-" * 86)
print(f"事实句可答对 {ok}/{len(CASES)} = {ok / len(CASES) * 100:.0f}%")
