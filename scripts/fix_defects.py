# -*- coding: utf-8 -*-
"""按扫描报告修复语料库缺陷，输出新版本

修复内容：
  D03  燃烧性能等级 B1/B2/A1/A2 → B₁/B₂/A₁/A₂（下标，符合字段说明约定）
  D04  罗马数字 I/II/III → Ⅰ/Ⅱ/Ⅲ（限"X类""X级"语境）
  D05  °C → ℃（U+2103）
  D11  二级款 "1）" → "（1）"（补全括号，符合规范原文格式）

用法：python qa\\fix_defects.py [源文件] [目标文件]
"""
import hashlib
import re
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
ROOT = HERE.parent

# 燃烧性能等级：字母+数字，且后面跟"级"
RE_BURN = re.compile(r"(?<![A-Za-z0-9])([AB])([12])(?=\s*级)")
SUB = {"1": "₁", "2": "₂"}

# 罗马数字：仅当后跟"类"或"级"，避免误伤英文
RE_ROMAN = re.compile(r"(?<![A-Za-z\u2160-\u217f])(I{1,3})(?=\s*(?:类|级))")
ROMAN = {"I": "Ⅰ", "II": "Ⅱ", "III": "Ⅲ"}

# 二级款：行首 "1）" "2）" -> "（1）"
RE_SUBITEM = re.compile(r"(?m)^(\s*)(\d)）")

# 摄氏度
RE_C = re.compile(r"°C")


def fix_text(s):
    if not isinstance(s, str):
        return s, []
    notes = []
    n1 = len(RE_BURN.findall(s))
    if n1:
        s = RE_BURN.sub(lambda m: m.group(1) + SUB[m.group(2)], s)
        notes.append(f"B下标×{n1}")
    n2 = len(RE_ROMAN.findall(s))
    if n2:
        s = RE_ROMAN.sub(lambda m: ROMAN[m.group(1)], s)
        notes.append(f"罗马数字×{n2}")
    n3 = len(RE_C.findall(s))
    if n3:
        s = RE_C.sub("℃", s)
        notes.append(f"℃×{n3}")
    n4 = len(RE_SUBITEM.findall(s))
    if n4:
        s = RE_SUBITEM.sub(lambda m: f"{m.group(1)}（{m.group(2)}）", s)
        notes.append(f"二级款括号×{n4}")
    return s, notes


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else str(
        max(ROOT.glob("*语料库_v*.xlsx"),
            key=lambda p: int(re.search(r"_v(\d+)", p.name).group(1))))
    ver = int(re.search(r"_v(\d+)", Path(src).name).group(1))
    dst = sys.argv[2] if len(sys.argv) > 2 else str(
        Path(src).with_name(Path(src).name.replace(f"_v{ver}.xlsx", f"_v{ver + 1}.xlsx")))

    wb = openpyxl.load_workbook(src)
    log = {}

    for ws in wb.worksheets:
        head = [c.value for c in ws[1]]
        idx_text = [i for i, h in enumerate(head)
                    if isinstance(h, str) and any(k in h for k in
                       ("条文原文本体", "切片全文", "该层正文", "单元格值", "切片标题"))]
        idx_stat = {i: head[i] for i in range(len(head)) if isinstance(head[i], str)
                    and any(k in head[i] for k in ("字数", "SHA256", "估算token"))}
        changed = 0
        for row in ws.iter_rows(min_row=2):
            touched = False
            for i in idx_text + list(idx_stat):
                cell = row[i]
                new, notes = fix_text(cell.value)
                if new != cell.value:
                    if "SHA256" in str(head[i]):
                        continue
                    cell.value = new
                    touched = True
                    for n in notes:
                        log[n] = log.get(n, 0) + 1
            if touched:
                # 重算字数 / token / SHA256
                for i, name in idx_stat.items():
                    if "字数" in name and idx_text:
                        cell = row[i]
                        if cell.value is not None:
                            cell.value = len(str(row[idx_text[0]].value or ""))
                    elif "token" in name and idx_text:
                        cell = row[i]
                        if cell.value is not None:
                            cell.value = int(len(str(row[idx_text[0]].value or "")) * 1.6)
                    elif "SHA256" in name and idx_text:
                        cell = row[i]
                        if cell.value is not None:
                            cell.value = hashlib.sha256(
                                str(row[idx_text[0]].value or "").encode()).hexdigest()
                changed += 1
        if changed:
            log[f"[{ws.title}]"] = changed

    wb.save(dst)
    print(f"源：{Path(src).name}")
    print(f"新：{Path(dst).name}\n")
    for k, v in log.items():
        print(f"  {k}: {v}")
    print(f"\n已保存 {dst}")


if __name__ == "__main__":
    main()
