# -*- coding: utf-8 -*-
"""表格切片可解读性评测：不问"渲染好不好看"，只问"读取方能不能答对"。

三层判据：
    L1 找得到 —— 问题能否检索到对应表的切片
    L2 值在不在 —— 目标数字是否出现在切片文本里
    L3 读得懂 —— 这个数字的"行标签 + 列标题"是否也在文本里、且能对上

L3 才是关键：LLM 看到 `| 12 |` 无法判断它属于"厂内主要道路"还是"厂内次要道路"。

用法：
    python eval_table_readability.py --corpus 语料库.xlsx --cases cases.json

cases.json 格式（每条一个表格问答）：
    [{"no": "4.2.9", "q": "…与厂内主要道路的防火间距是多少？",
      "value": "12", "row": "10t~", "col": "厂内主要道路", "answer": "12m"}, ...]
"""
import argparse
import json
import os
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

ap0 = argparse.ArgumentParser(add_help=False)
ap0.add_argument("--corpus", default=None, help="语料库 xlsx（或用 CORPUS_XLSX）")
ap0.add_argument("--cases", default=None, help="用例 JSON；不给则用内置示例")
_a, _ = ap0.parse_known_args()
if _a.corpus:
    os.environ["CORPUS_XLSX"] = str(Path(_a.corpus).resolve())

from raglib import Retriever  # noqa: E402

r = Retriever(use_table_slices=True, use_expand=True, child_weight=0.8)
tbl = {s["parent"]: s["text"] for s in r.slices if "表（子块" in s["level"]}

# (条号, 问题, 期望值, 行标签, 列标题, 正确答案叙述)
CASES = [
    ("4.1.7", "厂外铁路与甲、乙类可燃液体罐区的防火间距是多少？", "35",
     "厂外铁路", "甲、乙类", "35m"),
    ("4.1.7", "厂外其他公路与丙类罐区的防火间距是多少？", "12",
     "厂外其他公路", "丙类", "12m"),
    ("4.1.7", "高速公路与甲类仓库的防火间距是多少？", "40",
     "高速公路", "甲类", "40m"),
    ("4.2.9", "堆场储量10t~500t时，与厂内主要道路的防火间距是多少？", "12",
     "10t~", "厂内主要道路", "12m"),
    ("4.2.9", "堆场储量10t~500t时，与厂内次要道路的防火间距是多少？", "5",
     "10t~", "厂内次要道路", "5m"),
    ("4.2.10", "生产厂房（甲类）与厂内主要道路的防火间距是多少？", "12",
     "甲类", "厂内主要道路", "12m"),
    ("4.2.10", "生产厂房（甲类）与铁路装卸设施的防火间距是多少？", "12",
     "甲类", "铁路装卸", "12m"),
]

print(f"{'#':>2s} {'条号':8s} {'L1找到':6s} {'L2值在':6s} {'L3行标签':8s} {'L3列标题':8s} 判定")
print("-" * 82)
res = []
for i, (no, q, val, rowk, colk, ans) in enumerate(CASES, 1):
    hits = r.search_parent(q, k=3)
    l1 = no in [h["parent"] for h in hits]
    txt = tbl.get(no, "")
    l2 = val in txt
    l3r = rowk in txt
    l3c = colk in txt
    ok = l1 and l2 and l3r and l3c
    res.append(ok)
    verdict = "✓ 读得懂" if ok else ("✗ 缺列标题" if l2 and not l3c else "✗ 缺信息")
    print(f"{i:2d} {no:8s} {'✓' if l1 else '✗':6s} {'✓' if l2 else '✗':6s} "
          f"{'✓' if l3r else '✗':8s} {'✓' if l3c else '✗':8s} {verdict}")

print("-" * 82)
ok_n = sum(res)
print(f"完全可解读 {ok_n}/{len(CASES)} = {ok_n / len(CASES) * 100:.0f}%")

# 按表统计
print("\n=== 按表统计 ===")
from collections import defaultdict
by = defaultdict(lambda: [0, 0])
for (no, *_rest), ok in zip(CASES, res):
    by[no][1] += 1
    by[no][0] += ok
for no, (o, t) in by.items():
    print(f"  {no:8s} {o}/{t} 可解读")
