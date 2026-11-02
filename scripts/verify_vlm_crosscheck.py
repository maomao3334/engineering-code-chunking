# -*- coding: utf-8 -*-
"""用「人工校准结果」当完美 VLM 的替身，验证交叉校验机制是否真的能抓错。

为什么能做这个验证：
    校验逻辑（crosscheck）不关心值从哪来。用人工校准的值喂进去，
    如果它能 100% 通过，说明机制正确；再人为注入几个错值，
    看它能不能抓出来 —— 这就是 VLM 编数字时的场景。

这一步在 API key 不可用时，仍能证明「VLM+校验」方案的设计是正确的。
"""
import json
import re
import sys
from pathlib import Path

CASE = Path(sys.argv[1]).parent / "_pipeline"
tabs = json.loads((CASE / "tables_ocr.json").read_text(encoding="utf-8"))
cal = json.loads((CASE / "table_columns_calibrated.json").read_text(encoding="utf-8"))


def ocr_numbers(t):
    out = {}
    for r, row in enumerate(t["grid"]):
        for c, v in enumerate(row):
            s = str(v).strip()
            if re.fullmatch(r"\d+(?:\.\d+)?", s):
                out[(r, c)] = s
    return out


def crosscheck(facts, t):
    """与 vlm_table_to_facts.crosscheck 相同逻辑"""
    ocr_by_row = {}
    for (r, c), v in ocr_numbers(t).items():
        ocr_by_row.setdefault(r, set()).add(v)
    verified, suspicious = [], []
    for f in facts:
        val = str(f["value"]).strip()
        hit = [r for r, vals in ocr_by_row.items() if val in vals]
        (verified if hit else suspicious).append(dict(f, ocr_rows=hit))
    return verified, suspicious


def fake_vlm(t, c):
    """用校准列名 + OCR 网格，模拟 VLM 的输出（结构对、值取自网格）"""
    grid, hr = t["grid"], c["header_rows"]
    cols = {int(k): v for k, v in c["cols"].items() if v}
    rows = {int(k): v for k, v in c["rows"].items() if v}
    facts = []
    for r in range(hr, len(grid)):
        rowlab = rows.get(r, "")
        for cc, collab in cols.items():
            v = str(grid[r][cc]).strip() if cc < len(grid[r]) else ""
            if re.fullmatch(r"\d+(?:\.\d+)?", v):
                facts.append(dict(row_label=rowlab, col_name=collab, value=v))
    return facts


print("=" * 74)
print("【验证 1】完美 VLM（人工校准值）→ 校验应 100% 通过")
print("=" * 74)
for no, c in cal.items():
    if not c["verified"]:
        continue
    t = [x for x in tabs if x["page"] == c["page"]
         and x.get("header_rows") == c["header_rows"]][0]
    facts = fake_vlm(t, c)
    ver, sus = crosscheck(facts, t)
    rate = len(ver) / max(len(facts), 1) * 100
    print(f"  {no}: {len(facts)} 条 → 通过 {len(ver)}, 可疑 {len(sus)}  ({rate:.0f}%)")

print()
print("=" * 74)
print("【验证 2】注入 VLM 编造的数字 → 校验应能抓出来")
print("=" * 74)
t = [x for x in tabs if x["page"] == 11 and x.get("header_rows") == 2][0]
c = cal["表4.2.9"]
facts = fake_vlm(t, c)
# 注入 3 个编造值（原表中不存在的数字）
for bad in ("99", "777", "1234"):
    facts.append(dict(row_label="伪造行", col_name="伪造列", value=bad))
ver, sus = crosscheck(facts, t)
caught = [s["value"] for s in sus]
print(f"  总事实 {len(facts)} 条（含 3 条编造）")
print(f"  通过 {len(ver)}，可疑 {len(sus)}")
print(f"  抓到编造值: {caught}")
print(f"  判定: {'✓ 全部抓出' if set(caught) == {'99','777','1234'} else '✗ 有漏网'}")

print()
print("=" * 74)
print("【验证 3】VLM 少报（漏读格子）→ 应能统计出遗漏")
print("=" * 74)
facts = fake_vlm(t, c)[:5]      # 只给 5 条
ocr_vals = set(ocr_numbers(t).values())
vlm_vals = {f["value"] for f in facts}
miss = ocr_vals - vlm_vals
print(f"  VLM 只报 {len(facts)} 条，OCR 共 {len(ocr_numbers(t))} 个数字")
print(f"  VLM 漏掉的值: {sorted(miss)[:12]}{'…' if len(miss)>12 else ''}")
print(f"  遗漏计数 {len(miss)} → 可在报告里提示「该表可能未完整读取」")
