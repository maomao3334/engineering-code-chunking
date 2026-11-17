# -*- coding: utf-8 -*-
"""交付前 schema 体检：页签齐不齐、必备列在不在、复审要用的元数据有没有。

为什么单独一个脚本（而不是塞进 validate_chunks.py）：
    validate_chunks.py 查的是"结构对不对"（切片 ID、款数、父子关系）；
    这里查的是"字段契约全不全"——少了「表格明细.来源」，切片照样对，
    但复审工作台算不出不确定度，会静默降级成关键词猜测。这类问题要在源头拦住。

用法：
    python lint_schema.py 语料库.xlsx
    python lint_schema.py 语料库.xlsx --strict          # 警告也算失败
    python lint_schema.py 语料库.xlsx --json out.json   # 机器可读

退出码：0 通过（可能有警告）；1 有错误（或 --strict 下有警告）。
"""
import argparse
import json
import re
import sys

try:
    import openpyxl
except ImportError:
    sys.exit("需要 openpyxl: pip install openpyxl")

# 必备页签：缺一个就没法做复审/导出
REQUIRED_SHEETS = {
    "语料库": "条文主表，一切导出的源头",
    "列项明细": "款级明细（铁律 1 的子块来源）",
    "表格明细": "表格单元格（铁律 2 的独立块来源）",
    "章节结构": "章节层级与统计",
    "RAG切片": "切片正文（交付物本体）",
    "字段说明": "字段定义与边界（证明知道自己在做什么）",
    "质检报告": "自动校验结果 + 存疑清单（复审工作台的输入）",
}
OPTIONAL_SHEETS = {
    "公式明细": "有公式时必选",
    "图表清单": "有图时必选",
    "修订标记": "有局部修订下划线时必选",
    "页级元素": "强烈建议：图题/表注单列，防止污染款正文",
    "前置文件": "建议：封面/公告/目次",
}

# 页签 → (必备列, 建议列)
COLUMNS = {
    "语料库": (
        ["规范全称", "规范编号", "章", "节", "条文号", "条文原文本体",
         "是否强条", "条文类型", "PDF起页", "PDF末页", "块ID", "列项数", "字数"],
        ["关联表格", "关联附录", "引用条文", "小标题来源", "删改状态", "修订标记", "审核状态"],
    ),
    "列项明细": (["条号", "列项路径", "该层正文"], ["条号来源", "PDF起页", "所属条小标题"]),
    "表格明细": (
        ["表号", "表标题", "所属条文", "行号", "列序号", "列标题", "单元格值",
         # 来源/交叉校验是复审工作台算不确定度的输入；缺了就只能靠关键词猜，
         # 等于"图片表格"和"文字层表格"分不出来 —— 按错误处理，不是建议。
         "来源", "交叉校验"],
        ["来源说明", "转录提示"],
    ),
    "章节结构": (["层级", "编号", "标题", "条文数"], ["PDF起页", "列项数", "字数", "条文号范围"]),
    "RAG切片": (["切片ID", "父条号", "切片层级", "切片全文"], ["正文字符数", "PDF起页", "PDF末页"]),
    "公式明细": (["公式编号", "所属条文", "规范化表达式", "可信度"], ["PDF页", "存疑说明"]),
    "修订标记": (["条文号", "PDF页", "被下划线的原文行"], ["被下划线文字（按段拆分）", "下划线x区间"]),
    "页级元素": (["条文号", "元素类型", "原文行"], ["PDF页"]),
    "图表清单": (["图表号", "图表名称"], ["所属条文", "PDF页", "内容说明"]),
}

KEY_SHEET_MUST_HAVE_ROWS = {"语料库": 10, "RAG切片": 10}


def head_of(ws):
    rows = ws.iter_rows(min_row=1, max_row=1, values_only=True)
    r = next(rows, None)
    return [str(x or "").strip() for x in (r or [])]


def has(head, name):
    return any(h == name or h.startswith(name) for h in head)


def main():
    ap = argparse.ArgumentParser(description="交付前 schema 体检")
    ap.add_argument("xlsx")
    ap.add_argument("--strict", action="store_true", help="警告也算失败")
    ap.add_argument("--json", dest="json_out", default="", help="机器可读结果")
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.xlsx, read_only=True, data_only=True)
    errs, warns, oks = [], [], []

    for name, why in REQUIRED_SHEETS.items():
        if name in wb.sheetnames:
            oks.append(f"页签「{name}」存在")
        else:
            errs.append(f"缺必备页签「{name}」（{why}）")
    for name, why in OPTIONAL_SHEETS.items():
        if name in wb.sheetnames:
            oks.append(f"页签「{name}」存在")
        else:
            warns.append(f"缺建议页签「{name}」（{why}）")

    # 列检查
    for name, (must, nice) in COLUMNS.items():
        if name not in wb.sheetnames:
            continue
        head = head_of(wb[name])
        miss = [c for c in must if not has(head, c)]
        if miss:
            errs.append(f"「{name}」缺必备列：{miss}（现有 {len(head)} 列）")
        else:
            oks.append(f"「{name}」必备列齐全（{len(head)} 列）")
        miss2 = [c for c in nice if not has(head, c)]
        if miss2:
            warns.append(f"「{name}」缺建议列：{miss2}")

    # 行数：核心页签不能是空壳
    for name, least in KEY_SHEET_MUST_HAVE_ROWS.items():
        if name in wb.sheetnames:
            n = wb[name].max_row - 1
            if n < least:
                errs.append(f"「{name}」只有 {n} 行（至少应有 {least} 行），疑似只建了空壳")
            else:
                oks.append(f"「{name}」{n} 行")

    # 表格明细必须有来源标注（v2 契约）
    if "表格明细" in wb.sheetnames:
        head = head_of(wb["表格明细"])
        ws = wb["表格明细"]
        if has(head, "来源"):
            i = next(i for i, h in enumerate(head) if h == "来源" or h.startswith("来源"))
            vals = {str(r[i] or "").strip() for r in ws.iter_rows(min_row=2, values_only=True)
                    if len(r) > i and str(r[i] or "").strip()}
            if not vals:
                errs.append("「表格明细.来源」全空：分不清文字层表格与图片转录表格，复审无法排序")
            elif not vals <= {"文字层", "图片读图"}:
                warns.append(f"「表格明细.来源」出现约定外的取值：{sorted(vals - {'文字层', '图片读图'})}")
            else:
                oks.append(f"「表格明细.来源」取值规范：{sorted(vals)}")

    # 质检报告要有存疑段（复审工作台读它算严重度）
    if "质检报告" in wb.sheetnames:
        head = head_of(wb["质检报告"])
        rows = list(wb["质检报告"].iter_rows(min_row=2, values_only=True))
        if not rows or all(str(x or "").strip() in ("", "占位") for x in rows[0]):
            errs.append("「质检报告」还是占位行：必须先跑质检，复审工作台才算得出严重度")
        else:
            i = next((i for i, h in enumerate(head) if h.startswith("检查维度")), None)
            g = []
            if i is not None:
                g = [r for r in rows if len(r) > i and str(r[i] or "").startswith("G")]
            if g:
                oks.append(f"「质检报告」存疑段 {len(g)} 条")
            else:
                warns.append("「质检报告」没有 G 段存疑行：复审队列会缺少「存疑严重度」这一档加分")

    # 页级元素若存在，款正文不应含页级元素（这里只做提示，硬校验在 validate_chunks）
    if "页级元素" in wb.sheetnames and "列项明细" in wb.sheetnames:
        warns_n = 0
        for r in wb["列项明细"].iter_rows(min_row=2, values_only=True):
            if not r:
                continue
            body = str(r[2] if len(r) > 2 else "")
            if re.match(r"^\s*(?:图\s*\d|表\s*\d|附图|附表|注\s*[：:])", body):
                warns_n += 1
        if warns_n:
            warns.append(f"「列项明细」有 {warns_n} 个款以图题/表注开头，疑似页级元素污染（详见 validate_chunks）")

    print(f"文件：{a.xlsx}")
    print(f"页签：{len(wb.sheetnames)} 个 → {wb.sheetnames}\n")
    for x in oks:
        print(f"  ✅ {x}")
    for x in warns:
        print(f"  ⚠️  {x}")
    for x in errs:
        print(f"  ❌ {x}")
    print()
    if a.json_out:
        with open(a.json_out, "w", encoding="utf-8") as fh:
            json.dump(dict(xlsx=a.xlsx, sheets=wb.sheetnames, ok=len(oks),
                           warnings=warns, errors=errs), fh, ensure_ascii=False, indent=1)

    if errs:
        print(f"❌ schema 体检未通过（{len(errs)} 项错误，{len(warns)} 项警告）")
        return 1
    if a.strict and warns:
        print(f"❌ --strict：{len(warns)} 项警告视为失败")
        return 1
    print(f"✅ schema 体检通过（{len(warns)} 项警告）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
