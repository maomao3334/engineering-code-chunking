# -*- coding: utf-8 -*-
"""全量检索验收：覆盖条文 / 表格 / 附录 三类。

分类：
    T 条文类（正文查询）
    B 表格类（防火间距）
    A 附录类（火灾危险性分类，新增）
    反向类
"""
import os
import sys
from pathlib import Path

import argparse
sys.path.insert(0, str(Path(__file__).parent))
_ap = argparse.ArgumentParser()
_ap.add_argument("--corpus", required=True, help="语料库 xlsx")
_a = _ap.parse_args()
os.environ["CORPUS_XLSX"] = str(Path(_a.corpus).resolve())
from raglib import Retriever  # noqa: E402

r = Retriever(use_table_slices=True, use_expand=True, child_weight=0.8)

CASES = [
    # ---- T 条文类 ----
    ("T", "厂房的防火分区最大允许建筑面积是多少？", "面积"),
    ("T", "纺织厂房的耐火等级不应低于几级？", "耐火等级"),
    ("T", "消防车道的净宽度不应小于多少？", "4"),
    ("T", "室外消防用水量如何确定？", "用水量"),
    # ---- B 表格类 ----
    ("B", "厂外铁路与甲、乙类可燃液体储罐的防火间距？", "35"),
    ("B", "明火及散发火花地点与厂内主要道路的防火间距？", "10"),
    ("B", "松散可燃材料（10t~500t）与厂内次要道路的防火间距？", "5"),
    ("B", "行政生活建筑三级与码头装卸区的防火间距？", "40"),
    # ---- A 附录类 ----
    ("A", "涤纶属于几类火灾危险性？", "丙"),
    ("A", "粘胶纤维厂二硫化碳库属于几类？", "甲"),
    ("A", "涤纶成品库的火灾危险性？", "丙"),
    ("A", "腈纶生产的火灾危险性类别？", "丙"),
    ("A", "印染厂汽油气化室属于几类？", "甲"),
    ("A", "棉纺织厂棉粉尘属于几类？", "丙"),
    # ---- 反向类 ----
    ("R", "哪些生产环节属于甲类？", "甲"),
    ("R", "哪些设施与厂内铁路的防火间距是10m？", "10"),
    ("R", "与明火及散发火花地点间距为30m的有哪些？", "30"),
]

print(f"{'类':>2s} {'问题':44s} {'判定':>4s} top3")
print("-" * 92)
stat = {"T": [0, 0], "B": [0, 0], "A": [0, 0], "R": [0, 0]}
for cat, q, want in CASES:
    hits = r.search_parent(q, k=3)
    ctx = "\n".join(h["context"] for h in hits)
    ok = (want in ctx)
    stat[cat][1] += 1
    stat[cat][0] += ok
    print(f"{cat:>2s} {q[:42]:44s} {'✓' if ok else '✗':>4s} {[h['parent'] for h in hits]}")

print("-" * 92)
tot_ok = sum(v[0] for v in stat.values())
tot_n = sum(v[1] for v in stat.values())
for k, (o, n) in stat.items():
    name = {"T": "条文", "B": "表格", "A": "附录", "R": "反向"}[k]
    print(f"  {name}类: {o}/{n}")
print(f"  合计: {tot_ok}/{tot_n} = {tot_ok/tot_n*100:.0f}%")
