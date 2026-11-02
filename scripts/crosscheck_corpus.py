# -*- coding: utf-8 -*-
"""独立复核：不复用解析器，直接从 PDF 原文重算一遍，和语料库对账。

**为什么需要这一步**
    结构校验（validate_chunks）和缺陷扫描（scan_defects）都只看**语料库内部**：
    条号唯一、款号连续、括号配对……它们全都通过，内容仍可能已经脏了。
    实测踩过：某条正文后面被塞进 500 字表注，两项自动校验 100% 通过。

**怎么做到"独立"**
    用另一条代码路径（PDF 原始文字层 + 简单正则）重新数一遍条号，
    与语料库做集合对比。两条路径结论一致，才敢说没漏。

**查什么**
    1. 漏条：PDF 里有、语料库里没有
    2. 编造：语料库里有、PDF 里没有
    3. 断号：章内条号不连续
    4. 串条：正文里出现了别的条号
    5. 截断：正文结尾不是句读
    6. 对账：切片总数 = 父块 + 款 + 表

用法：
    python crosscheck_corpus.py --pdf 规范.pdf --xlsx 语料库.xlsx
"""
import argparse
import re
import sys

import fitz
from openpyxl import load_workbook

WM = re.compile(r"1190119|公众号|中国消防规范网")
PAGENO = re.compile(r"^\d{1,3}(\s*公众号\s*CN\d+)?$")


def pdf_lines(pdf):
    """PDF 原文行（去水印/页码），用于独立比对"""
    doc = fitz.open(pdf)
    out = []
    for i in range(doc.page_count):
        for ln in doc[i].get_text().split("\n"):
            s = ln.strip()
            if not s or WM.search(s) or PAGENO.match(s):
                continue
            out.append((i + 1, s))
    doc.close()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--xlsx", required=True)
    a = ap.parse_args()

    lines = pdf_lines(a.pdf)
    pdf_all = "".join(s for _, s in lines)

    # ---- PDF 侧：所有行首 X.Y.Z
    pdf_nos = {}
    for pg, s in lines:
        m = re.match(r"^(\d{1,2}\.\d{1,2}\.\d{1,3})\s+\S", s)
        if m:
            pdf_nos.setdefault(m.group(1), []).append(pg)

    # ---- 语料库侧
    wb = load_workbook(a.xlsx, read_only=True)
    ws = wb["语料库"]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[4]:
            rows.append(dict(no=str(r[4]), body=r[6] or ""))
    n_item = sum(1 for x in wb["列项明细"].iter_rows(min_row=2, values_only=True) if x[0])
    n_slice = sum(1 for x in wb["RAG切片"].iter_rows(min_row=2, values_only=True) if x[0])
    n_tbl = len({x[0] for x in wb["表格明细"].iter_rows(min_row=2, values_only=True) if x[0]})
    wb.close()

    kb = {r["no"] for r in rows}
    print(f"PDF 抓到条号 {len(pdf_nos)} 个（去重后）")
    print(f"语料库 {len(rows)} 条 / {n_item} 款 / {n_tbl} 表 / {n_slice} 切片\n")

    ok = True

    # 1 漏条
    miss = sorted(set(pdf_nos) - kb)
    print(f"1. 漏条（PDF 有、语料库无）: {miss or '无'}")
    ok &= not miss
    for n in miss:
        print(f"     出现在 p{pdf_nos[n]}")

    # 2 编造
    #   附录的条号（附录A/附录B）来自图片表格，PDF 文字层里抓不到编号，
    #   属于正常，不计入"编造"。
    fake = sorted(x for x in (kb - set(pdf_nos)) if not str(x).startswith("附录"))
    print(f"2. 编造（语料库有、PDF 无）: {fake or '无'}")
    ok &= not fake

    # 3 断号
    #   附录（附录A/附录B）不是 X.Y.Z 格式，跳过断号检查
    by_ch = {}
    for n in kb:
        p = str(n).split(".")
        if len(p) < 3 or not all(x.isdigit() for x in p):
            continue
        by_ch.setdefault(p[0], set()).add(int(p[2]))
    gaps = []
    for ch, nums in by_ch.items():
        exp = set(range(1, max(nums) + 1))
        d = sorted(exp - nums)
        if d:
            gaps.append((ch, d))
    print(f"3. 章内断号: {gaps or '无'}")
    ok &= not gaps

    # 4 串条（正文里出现别的条号）
    sus = []
    for r in rows:
        for m in re.finditer(r"(?<![条表附录])(\d{1,2}\.\d{1,2}\.\d{1,3})\s", r["body"]):
            if m.group(1) != r["no"]:
                sus.append((r["no"], m.group(1)))
                break
    print(f"4. 串条（正文含其他条号）: {sus or '无'}")
    if sus:
        print("     ↑ 可能是正常引用（\"应符合第X条\"），也可能是串条，需人工看")
        for x, y in sus[:8]:
            print(f"       {x} 含 {y}")

    # 5 截断
    BAD = tuple("。；：、）)1234567890")
    trunc = [r["no"] for r in rows if r["body"] and not r["body"].rstrip().endswith(BAD)]
    print(f"5. 疑似截断（结尾非句读）: {trunc or '无'}")
    ok &= not trunc

    # 6 抽样逐字
    print("\n6. 抽样逐字对照（最长/最短各 2 条）")
    srt = sorted(rows, key=lambda r: len(r["body"]))
    for r in srt[:2] + srt[-2:]:
        plain = re.sub(r"^\d+\s+", "", r["body"].replace("\n", ""))
        probe = plain[:32]
        found = probe in pdf_all or plain[:20] in pdf_all
        print(f"   {r['no']:8s} {len(r['body']):5d}字  "
              f"{'✓ 在PDF中' if found else '✗ 未找到，需人工看'}")
        if not found:
            print(f"       {probe}")
            ok = False

    # 7 对账
    #   注意：表格切片数 ≠ 表数量。一张大表会按行拆成多个切片，
    #   还会额外生成一个「反向速查」片。所以这里直接统计切片里以「表」开头的层级。
    import openpyxl as _ox
    _wb = _ox.load_workbook(str(a.xlsx), read_only=True)
    _ws = _wb["RAG切片"]
    n_tblslice = 0
    for _r in _ws.iter_rows(min_row=2, values_only=True):
        lv = str(_r[2] or "")
        if "父块" in lv:
            continue                      # 父块单独计（= 语料库行数）
        # 表格/附录子块都算作"非款子块"
        if lv.startswith("表（") or "附录" in lv:
            n_tblslice += 1
    _wb.close()
    exp = len(rows) + n_item + n_tblslice
    print(f"\n7. 切片对账: 期望 {len(rows)}父+{n_item}款+{n_tblslice}表切片"
          f"（{n_tbl} 张表按行拆分+速查）={exp}，"
          f"实际 {n_slice} → {'✓' if n_slice == exp else '✗'}")
    ok &= (n_slice == exp)

    print("\n" + ("✅ 独立复核通过" if ok else "❌ 复核发现问题，见上"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
