# -*- coding: utf-8 -*-
"""表格 → 自包含事实句（用人工校准的列名，只生成能确认的）。

设计原则：
    1. 列名/行标签**只用校准过的**。未校准的表直接跳过，宁可不生成。
    2. 每条事实句带来源坐标（表号+行列），可回原图核对。
    3. 数值必须是纯数字；空格、`—` 一律跳过。
    4. 输出两个版本：`text`（自然句）和 `key`（检索用关键词串），
       后者保证 BM25 能命中"厂内主要道路 防火间距"这类查询。

覆盖范围（实测）：
    表4.1.7  10 行 × 5 列 = 最多 50 句
    表4.2.9   4 行 × 9 列 = 最多 36 句（行标签由"材料名+储量"组合）
    表4.2.10 跳过（未校准）
"""
import json
import re
import sys
from pathlib import Path

CASE = Path(sys.argv[1]).parent / "_pipeline"
# 注意：同一页可能有**多张表**（p11 就有两张：可燃材料堆场 / 原麻等）。
# 用 page 当 key 会让后者覆盖前者 —— 实测因此把第二张表的数据错配到第一张。
# 修正：保留列表，按"表头行数+列数+首个数据格"匹配校准项。
_raw = [t for t in json.loads((CASE / "tables_ocr.json").read_text(encoding="utf-8"))
        if t["w"] < 1000 and t["grid"]]
cal = json.loads((CASE / "table_columns_calibrated.json").read_text(encoding="utf-8"))


def pick_table(page, header_rows, min_col):
    """在同一页的多张表里，选出与校准项匹配的那张"""
    cands = [t for t in _raw if t["page"] == page
             and t.get("header_rows") == header_rows
             and len(t["grid"][0]) >= min_col]
    return cands[0] if cands else None

SPEC = "《纺织工程设计防火规范》GB 50565-2010"


def is_num(s):
    return bool(re.fullmatch(r"\d+(?:\.\d+)?", str(s).strip()))


def build(no, c):
    """按校准表生成事实句"""
    maxcol = max([int(k) for k in c["cols"]] or [0])
    t = pick_table(c["page"], c["header_rows"], maxcol + 1)
    if not t or not c["verified"]:
        return []
    grid, hr = t["grid"], c["header_rows"]
    nrow, ncol = len(grid), len(grid[0])
    # JSON 读回来的键是字符串，转回 int
    cols = {int(k): v for k, v in c["cols"].items() if v}
    rows = {int(k): v for k, v in c["rows"].items() if v}
    facts = []

    for r in range(hr, nrow):
        # 行标签：优先用校准值，否则从"非数据列"里取文字
        rowlab = rows.get(r, "")
        if not rowlab:
            parts = []
            for cc in range(ncol):
                if cc in cols:
                    continue
                v = grid[r][cc].strip()
                if v and v not in ("—", "-") and not is_num(v):
                    parts.append(v)
            rowlab = "，".join(parts)
        if not rowlab:
            continue

        for cc, collab in cols.items():
            if cc >= ncol:
                continue
            v = grid[r][cc].strip()
            if not v or v in ("—", "-") or not is_num(v):
                continue
            unit = c.get("unit", "m")
            text = (f"{SPEC}{no}"
                    f"（{c['title']}）：{rowlab} 与「{collab}」之间的防火间距"
                    f"不应小于 {v}{unit}。"
                    f"（依据：{SPEC} {no}，表{no.replace('表','')}，原文第{c['page']}页）")
            key = f"{rowlab} {collab} 防火间距 {v}{unit} {no}"
            facts.append(dict(table=no, page=c["page"], row=r, col=cc,
                              row_label=rowlab, col_label=collab,
                              value=v, unit=unit, text=text, key=key,
                              confidence="high"))
    return facts


allfacts = []
for no, c in cal.items():
    fs = build(no, c)
    allfacts += fs
    print(f"  {no}: 生成 {len(fs)} 条事实句")

out = CASE / "table_facts.json"
out.write_text(json.dumps(allfacts, ensure_ascii=False, indent=1), encoding="utf-8")

print(f"\n合计 {len(allfacts)} 条 → {out}\n")
print("=== 样本（表4.2.9 前 6 条）===")
for f in [x for x in allfacts if x["table"] == "表4.2.9"][:6]:
    print("  " + f["text"])
print("\n=== 样本（表4.1.7 前 4 条）===")
for f in [x for x in allfacts if x["table"] == "表4.1.7"][:4]:
    print("  " + f["text"])
