# -*- coding: utf-8 -*-
"""表格质量分级：自动判定每张表的还原可信度，写进质检报告。

为什么要分级：OCR + 坐标重建**不可能对所有表格都完美**。
实测《纺织工程设计防火规范》三张表：
    表4.1.7  横向表头      → 数据行全对，表头基本可读         A
    表4.2.9  竖排表头+合并 → 数据值基本对，表头被切成多行       B
    表4.2.10 超宽(42列)+合并→ 数据对，表头部分错位            B

如果统一报"表格已还原"，等于给使用者一个假的安全感。
分级 + 明确写清"哪一部分可信"，才是诚实的交付。
"""
import json
import re
import sys
from pathlib import Path

if len(sys.argv) < 2:
    sys.exit("用法：python grade_table_quality.py <规范.pdf>\n"
             "读取 <pdf目录>/_pipeline/tables_ocr.json，给每张表打 A/B/C 质量等级。\n"
             "A 可直接用 / B 数据可用表头需核对 / C 需人工重建")

CASE = Path(sys.argv[1]).parent / "_pipeline"
d = json.loads((CASE / "articles.json").read_text(encoding="utf-8"))
tabs = json.loads((CASE / "tables_ocr.json").read_text(encoding="utf-8"))


def grade(t):
    """返回 (等级, 理由)

    注意：不能只看"数字占比"——附录A/B 是"材料名称/火灾危险性类别"纯文字表，
    本来就没有数字，用数字判据会把它们全判成 C（实测踩过）。

    判据改为三类：
      A：结构整齐（空列少、列数适中），无论数字多少
      B：结构可用但有缺陷（竖排表头被拆行 / 列数过多 / 空列偏多）
      C：结构已散（空列过多 + 每行有效文字过少）
    """
    grid = t["grid"]
    if not grid:
        return "C", "OCR 无结果"

    n_row = len(grid)
    n_col = len(grid[0]) if grid else 0
    total = n_row * n_col or 1
    empty = sum(1 for row in grid for c in row if not c.strip())
    empty_ratio = empty / total
    filled = total - empty
    # 每行平均有效格数（衡量"行是否被拆碎"）
    per_row = filled / n_row if n_row else 0

    reasons = []
    # C：每行有效格太少（<1.2），说明文字被拆成一行一个
    if per_row < 1.2:
        return "C", f"每行仅 {per_row:.1f} 个有效格，表头可能被拆成逐字一行"
    # B：列数过多 或 空列偏多
    if n_col > 16:
        reasons.append(f"列数过多({n_col})")
    if empty_ratio > 0.60:
        reasons.append(f"空列偏多({empty_ratio:.0%})")
    if reasons:
        return "B", "；".join(reasons)
    return "A", f"每行 {per_row:.1f} 格，空列 {empty_ratio:.0%}"


print(f"{'表号':10s} {'页':>3s} {'行×列':>9s} {'数字占比':>8s} {'空列':>6s} 等级  说明")
print("-" * 78)
rows = []
for t in tabs:
    if t["w"] > 1000:          # 广告图
        continue
    g, why = grade(t)
    n_row, n_col = len(t["grid"]), (len(t["grid"][0]) if t["grid"] else 0)
    nums = sum(1 for r in t["grid"] for c in r if re.fullmatch(r"\d+(?:\.\d+)?", c.strip()))
    tot = n_row * n_col or 1
    empt = sum(1 for r in t["grid"] for c in r if not c.strip())
    rows.append(dict(page=t["page"], n_row=n_row, n_col=n_col, grade=g, why=why))
    print(f"  p{t['page']:<3d}    {n_row:>2d}×{n_col:<3d}  {nums/tot:>7.0%} {empt/tot:>6.0%}   {g}    {why}")

print("\n=== 建议写进质检报告的说明 ===")
for r in rows:
    label = {"A": "可直接使用", "B": "数据可用，表头需人工核对", "C": "需人工重建"}[r["grade"]]
    print(f"  p{r['page']}: {r['n_row']}行×{r['n_col']}列  等级 {r['grade']}（{label}）— {r['why']}")
