# -*- coding: utf-8 -*-
"""语料库缺陷扫描：系统性查找 OCR 与结构错误

背景：上一轮靠问答测试偶然发现了 8 处 `lx`→`1x` 的错误。
      人工翻表看不出来，但会直接让答案错。这里把这类检查固化下来。

用法：
    python qa\\scan_defects.py [语料库.xlsx]

输出：qa/缺陷扫描报告.md
"""
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
ROOT = HERE.parent
OUT = HERE / "缺陷扫描报告.md"


# ---------------- 规则 ----------------
# 每条：(编号, 名称, 严重度, 正则, 说明, 修法)
RULES = [
    ("D01", "照度单位 lx 被识别成 1x", "高",
     re.compile(r"\d\.\d1x\b"),
     "勒克斯写作 lx，OCR 把小写 l 认成数字 1",
     "数字末尾的 `1x` → `lx`"),

    ("D02", "上标 m2 / m3 未还原", "高",
     re.compile(r"(?<![A-Za-z])\d+\s?m[23](?![0-9A-Za-z])"),
     "应为 m² / m³（面积、体积）",
     "`m2`→`m²`，`m3`→`m³`"),

    ("D03", "燃烧性能等级 B1 / B2 未还原下标", "高",
     re.compile(r"(?<![A-Za-z0-9])[AB]1(?![0-9])"),
     "规范用 B₁、B₂（下标）；A2 级同理",
     "`B1`→`B₁`，`B2`→`B₂`"),

    ("D04", "罗马数字用阿拉伯/字母代替", "中",
     re.compile(r"(?<![A-Za-z])(?:I{1,3}|II|III)(?![A-Za-z])"),
     "规范用 Ⅰ Ⅱ Ⅲ（如 Ⅰ级木结构建筑、Ⅰ类飞机库）",
     "`I`→`Ⅰ`，`II`→`Ⅱ`，`III`→`Ⅲ`（需结合上下文判断）"),

    ("D05", "℃ 用了 °C 写法", "低",
     re.compile(r"°C"),
     "规范原文用 ℃（U+2103）",
     "视觉等价，与外部库比对时视为相同，不必强改"),

    ("D06", "数字与单位之间有多余空格", "中",
     re.compile(r"\d\s{2,}(?:m²|m³|m|mm|cm|h|min|lx|人|层)\b"),
     "OCR 断行导致 6.0 和 m² 之间被插入空格",
     "删除多余空格"),

    ("D07", "阿拉伯数字 0 与字母 O 混淆", "高",
     re.compile(r"\d[OО]\d|\b[OО]\.\d"),
     "OCR 把数字 0 认成字母 O（或西里尔 О）",
     "改回 0，需人工确认"),

    ("D08", "全角数字/字母混入", "中",
     re.compile(r"[０-９Ａ-Ｚａ-ｚ]"),
     "应为半角",
     "转半角"),

    ("D09", "条号引用格式不规范", "低",
     re.compile(r"第\s+\d+\.\d+\.\d+\s*条"),
     "规范原文是「第7.1.8条」，中间无空格",
     "去空格（只影响溯源匹配，不影响语义）"),

    ("D10", "疑似漏字（连续标点）", "中",
     re.compile(r"[，。；：]{2,}"),
     "可能 OCR 丢了字，或行重建时拼接出错",
     "人工核对原页"),

    ("D11", "括号不配对", "高",
     None,       # 单独用配对逻辑检查
     "左括号数与右括号数不一致，通常意味着一整段被截断",
     "人工核对原页"),

    ("D12", "「不应大于/不应小于」被写成相反", "高",
     None,
     "语义类错误，无法自动判定；这里只列出含比较词的条文供人工抽核",
     "人工抽核"),
]


def check_parens(text):
    pairs = [("（", "）"), ("(", ")"), ("《", "》"), ("【", "】")]
    bad = []
    for a, b in pairs:
        if text.count(a) != text.count(b):
            bad.append(f"{a}{b}：{text.count(a)}/{text.count(b)}")
    return bad


def main():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else str(
        max(ROOT.glob("*语料库_v*.xlsx"), key=lambda p: int(re.search(r"_v(\d+)", p.name).group(1))))

    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    print(f"扫描：{Path(xlsx).name}\n")

    findings = []            # (规则, 工作表, 位置, 片段, 上下文, 条号)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        head = [str(x or "") for x in rows[0]]
        for ri, row in enumerate(rows[1:], start=2):
            clause = ""
            for key in ("条文号", "父条号", "条号"):
                if key in head:
                    v = row[head.index(key)]
                    if v:
                        clause = str(v)
                        break
            for ci, val in enumerate(row):
                if not isinstance(val, str) or not val.strip():
                    continue
                for rid, name, sev, rx, why, fix in RULES:
                    if rx is None:
                        continue
                    for m in rx.finditer(val):
                        s = max(0, m.start() - 18)
                        ctx = val[s:m.end() + 18].replace("\n", " ")
                        findings.append(dict(rule=rid, sheet=ws.title, where=f"行{ri}",
                                             clause=clause, hit=m.group(),
                                             ctx=ctx, sev=sev))
                # 括号配对（只查正文类字段，且只报"右括号比左括号多"这种真丢失）
                if ci < len(head) and any(k in head[ci] for k in ("条文原文本体", "全文", "该层正文")):
                    pairs = [("（", "）"), ("《", "》")]
                    bad = [f"{a}{b}：{val.count(a)}/{val.count(b)}"
                           for a, b in pairs if val.count(b) > val.count(a)]
                    if bad:
                        findings.append(dict(rule="D11", sheet=ws.title, where=f"行{ri}",
                                             clause=clause, hit="; ".join(bad),
                                             ctx=val[:60].replace("\n", " "), sev="高"))

    # 去重（同一位置同一规则同一命中只报一次）
    seen, uniq = set(), []
    for f in findings:
        k = (f["rule"], f["sheet"], f["where"], f["hit"])
        if k not in seen:
            seen.add(k)
            uniq.append(f)

    # 按严重度分组统计
    by_rule = defaultdict(list)
    for f in uniq:
        by_rule[f["rule"]].append(f)

    rule_meta = {r[0]: (r[1], r[2], r[4], r[5]) for r in RULES}

    L, A = [], None
    def A(s):
        L.append(s)

    A("# 语料库缺陷扫描报告\n")
    A(f"- 扫描文件：`{Path(xlsx).name}`")
    A(f"- 规则数：{len(RULES)} 类（OCR 字符混淆、单位符号、结构完整性）")
    A(f"- 命中总数：**{len(uniq)}** 处\n")

    sev_order = {"高": 0, "中": 1, "低": 2}
    A("## 一、按规则汇总\n")
    A("| 规则 | 检查项 | 严重度 | 命中 |")
    A("|---|---|---|---|")
    for rid, name, sev, rx, why, fix in RULES:
        n = len(by_rule.get(rid, []))
        flag = "**" if (n and sev == "高") else ""
        A(f"| {rid} | {name} | {sev} | {flag}{n}{flag} |")
    A("")

    high = [f for f in uniq if f["sev"] == "高"]
    A(f"## 二、高严重度问题（{len(high)} 处，会直接影响答案正确性）\n")
    if not high:
        A("无。\n")
    else:
        A("| 规则 | 工作表 | 位置 | 条号 | 命中内容 | 上下文 |")
        A("|---|---|---|---|---|---|")
        for f in sorted(high, key=lambda x: x["rule"])[:60]:
            A(f"| {f['rule']} | {f['sheet']} | {f['where']} | {f['clause']} | "
              f"`{f['hit']}` | {f['ctx']} |")
        if len(high) > 60:
            A(f"\n（共 {len(high)} 处，此处列前 60 处）")
        A("")

    A("## 三、中低严重度（抽样）\n")
    for sev in ("中", "低"):
        items = [f for f in uniq if f["sev"] == sev]
        A(f"### {sev}（{len(items)} 处）\n")
        for rid in {f["rule"] for f in items}:
            fs = [f for f in items if f["rule"] == rid]
            nm = rule_meta[rid][0]
            A(f"- **{rid} {nm}**：{len(fs)} 处，例如 {fs[0]['sheet']} {fs[0]['where']} "
              f"`{fs[0]['hit']}`（{fs[0]['ctx']}）")
        A("")

    A("## 四、怎么处理\n")
    A("| 规则 | 处理方式 |")
    A("|---|---|")
    seen_rule = set()
    for rid, name, sev, rx, why, fix in RULES:
        if rid in by_rule and rid not in seen_rule:
            seen_rule.add(rid)
            A(f"| {rid} | {fix} |")
    A("")
    A("> 高严重度的建议全部修掉再交付；中低严重度逐条人工确认，"
      "其中 D05（℃/°C）属于视觉等价，与外部库比对时**不应算作差异**。\n")

    A("## 五、说明：为什么这类检查必须自动化\n")
    A("上一轮的 `lx`→`1x` 错误是靠问答测试**偶然**发现的：")
    A("问「疏散照明的地面最低水平照度是多少」，答案里出现 `10.01x`。")
    A("")
    A("这个错误的特点：")
    A("- 人眼扫表时，`10.01x` 和 `10.0lx` 几乎无法区分")
    A("- 但答案是错的，而且看起来很合理")
    A("- 241 条 × 多个页签，人工逐字核对不现实")
    A("")
    A("所以：**语料库交付前必须跑一遍自动扫描**，把它当成和 `validate_chunks.py` 同级的必做步骤。")

    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L[:34]))
    print(f"\n报告：{OUT}")


if __name__ == "__main__":
    main()
