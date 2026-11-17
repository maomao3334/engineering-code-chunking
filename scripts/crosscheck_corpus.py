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
    6. 抽样逐字：最长/最短各 2 条，看正文是否真能在 PDF 里找到
    7. 对账：切片总数 = 父块 + 款 + 表

用法：
    python crosscheck_corpus.py --pdf 规范.pdf --xlsx 语料库.xlsx
    python crosscheck_corpus.py --pdf 规范.pdf --xlsx 语料库.xlsx --allow 5.1.1A

--allow：把「原件本身如此」的疑点（如原书某款漏印句号）列为已知项，不再判失败。
    注意：这只影响本脚本的通过/失败，不改语料库内容；每个 allow 都应在质检报告里有登记。

退出码：0 通过；1 发现问题。
"""
import argparse
import re
import sys

import fitz
from openpyxl import load_workbook

WM = re.compile(r"1190119|公众号|中国消防规范网")
PAGENO = re.compile(r"^\d{1,3}(\s*公众号\s*CN\d+)?$")

# PDF 侧条号：X.Y.Z，允许局部修订后缀（1.0.5A）
RE_PDF_NO = re.compile(r"^(\d{1,2}\.\d{1,2}\.\d{1,3}[A-Z]?)[\s　]+\S")
RE_NO = re.compile(r"^\d{1,2}\.\d{1,2}\.\d{1,3}[A-Z]?$")
RE_PH = re.compile(r"〔公式[^〕]*〕")


def norm(s):
    """去掉公式占位、空白与标点——占位符是我们自己插的，PDF 里当然没有；
    标点的全/半角差异（：vs :）在比对时也不该算错。"""
    s = RE_PH.sub("", str(s or ""))
    s = re.sub(r"\s+", "", s)
    return re.sub(r"[：:；;，,、。．\.·（）()“”\"'—\-]", "", s)


def pdf_lines(pdf):
    """PDF 原文行（去水印/页码），用于独立比对"""
    doc = fitz.open(pdf)
    out, alltext = [], []
    for i in range(doc.page_count):
        for ln in doc[i].get_text().split("\n"):
            s = ln.strip()
            if not s or WM.search(s) or PAGENO.match(s):
                continue
            out.append((i + 1, s))
            alltext.append(s)
    doc.close()
    return out, "".join(alltext)


def head_index(ws):
    r = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    return [str(x or "").strip() for x in (r or [])]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--allow", default="",
                    help="已知且已登记的疑点条号，逗号分隔（如 5.1.1A）；只对「截断」检查生效")
    a = ap.parse_args()
    allow = {x.strip() for x in a.allow.split(",") if x.strip()}

    lines, pdf_all = pdf_lines(a.pdf)
    pdf_ns = norm(pdf_all)

    # ---- PDF 侧：行首条号
    pdf_nos = {}
    for pg, s in lines:
        m = RE_PDF_NO.match(s)
        if m:
            pdf_nos.setdefault(m.group(1), []).append(pg)

    # ---- 语料库侧（按列名取，不依赖列序）
    wb = load_workbook(a.xlsx, read_only=True)
    ws = wb["语料库"]
    head = head_index(ws)
    i_no = head.index("条文号") if "条文号" in head else 4
    i_body = next((i for i, h in enumerate(head) if h.startswith("条文原文本体")),
                  head.index("条文原文本体") if "条文原文本体" in head else 6)
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if len(r) > i_no and r[i_no]:
            rows.append(dict(no=str(r[i_no]), body=r[i_body] or ""))
    n_item = sum(1 for x in wb["列项明细"].iter_rows(min_row=2, values_only=True) if x[0])
    n_slice = sum(1 for x in wb["RAG切片"].iter_rows(min_row=2, values_only=True) if x[0])
    n_tbl = len({x[0] for x in wb["表格明细"].iter_rows(min_row=2, values_only=True) if x[0]})
    slice_levels = [str(x[2] or "") for x in wb["RAG切片"].iter_rows(min_row=2, values_only=True)]
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
    #   附录整册（附录A）、附录条款（G.0.1）、附加说明没有 X.Y.Z 形式条号，
    #   PDF 文字层抓不到，属正常，不计入"编造"；只查 X.Y.Z[A-Z]? 形态的。
    fake = sorted(x for x in (kb - set(pdf_nos)) if RE_NO.match(str(x)))
    other = sorted(x for x in (kb - set(pdf_nos)) if not RE_NO.match(str(x)))
    print(f"2. 编造（语料库有、PDF 无）: {fake or '无'}")
    ok &= not fake
    if other:
        print(f"   附录/非 X.Y.Z 形态的条目（不参与比对）：{other[:10]}")

    # 3 断号
    by_ch = {}
    for n in kb:
        s = str(n)
        if not RE_NO.match(s):
            continue
        p = re.sub(r"[A-Z]$", "", s).split(".")
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
    #   两个坑：① 款号 1.0.4.1 会被切成 0.4.1；② "应符合第3.2.3条"是正常引用。
    #   所以要求左侧不是数字/点，且不是"第X条"的引用形态。
    RE_IN = re.compile(r"(?<![\d.条表附录])(\d{1,2}\.\d{1,2}\.\d{1,3}[A-Z]?)(?![\d.])")
    sus = []
    for r in rows:
        for m in RE_IN.finditer(r["body"]):
            if m.group(1) == r["no"]:
                continue
            head_ctx = r["body"][max(0, m.start() - 3):m.start()]
            tail_ctx = r["body"][m.end():m.end() + 2]
            if head_ctx.endswith("第") and tail_ctx.startswith("条"):
                continue                      # 正常引用
            sus.append((r["no"], m.group(1)))
            break
    print(f"4. 串条（正文含其他条号）: {sus or '无'}")
    if sus:
        print("     ↑ 可能是正常引用（\"应符合第X条\"），也可能是串条，需人工看")
        for x, y in sus[:8]:
            print(f"       {x} 含 {y}")

    # 5 截断（--allow 里的条号降级为提示）
    #   只查条文（X.Y.Z[A-Z]?）：附录整册的记录以表题/表注收尾，本来就不是句子。
    BAD = tuple("。；：、）)1234567890")
    trunc = [r["no"] for r in rows
             if RE_NO.match(str(r["no"])) and r["body"]
             and not r["body"].rstrip().endswith(BAD)]
    trunc_bad = [x for x in trunc if x not in allow]
    trunc_ok = [x for x in trunc if x in allow]
    print(f"5. 疑似截断（结尾非句读）: {trunc_bad or '无'}"
          + (f"　（已登记的已知项：{trunc_ok}）" if trunc_ok else ""))
    ok &= not trunc_bad

    # 6 抽样逐字（去公式占位 + 去空白与标点后比对；占位/全半角差异不算错）
    #   注意：含公式的条文，正文里"计算："到"式中"之间是我们插的占位符，
    #   在 PDF 里被公式编号图像隔断，连续子串必然找不到——所以按占位符切段，
    #   取最长的一段做探针，而不是拿整段开头硬比。
    print("\n6. 抽样逐字对照（最长/最短各 2 条）")
    srt = sorted(rows, key=lambda r: len(r["body"]))
    for r in srt[:2] + srt[-2:]:
        raw = re.sub(r"^\d+\s+", "", r["body"])
        segs = [norm(x) for x in RE_PH.split(raw)]
        segs = [x for x in segs if x]
        if not segs or max(len(x) for x in segs) < 12:
            print(f"   {r['no']:8s} {len(r['body']):5d}字  － 过短/整段为占位，跳过")
            continue
        probe = max(segs, key=len)
        found = any(len(p) >= 12 and p in pdf_ns for p in (probe[:22], probe[:16]))
        print(f"   {r['no']:8s} {len(r['body']):5d}字  "
              f"{'✓ 在PDF中' if found else '✗ 未找到，需人工看'}")
        if not found:
            print(f"       {probe[:40]}")
            ok = False

    # 7 对账（表格/附录子块都算"非款子块"）
    n_tblslice = sum(1 for lv in slice_levels
                     if "父块" not in lv and (lv.startswith("表（") or "附录" in lv))
    exp = len(rows) + n_item + n_tblslice
    print(f"\n7. 切片对账: 期望 {len(rows)}父+{n_item}款+{n_tblslice}表切片"
          f"（{n_tbl} 张表按行拆分+速查）={exp}，"
          f"实际 {n_slice} → {'✓' if n_slice == exp else '✗'}")
    ok &= (n_slice == exp)

    print("\n" + ("✅ 独立复核通过" if ok else "❌ 复核发现问题，见上"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
