# -*- coding: utf-8 -*-
"""把填好的「人工复审工作台」回收成修正清单（overrides.json）+ 复审汇总。

为什么需要它
------------
复审人只会填「结论 + 复核意见」，维护人需要的是**可执行的结构化修正清单**。
中间这一步如果靠人转述，必然丢件、串行、漏依据。所以固化成一个脚本：

    工作台（人工填） → overrides.json（机器读） → 建库脚本回灌 → 重跑校验

用法：
    python collect_review_results.py --workbench 人工复审工作台.xlsx
    python collect_review_results.py --workbench 工作台.xlsx --corpus 语料库.xlsx \
                                     --out overrides.json --report 复审结果汇总.md
    python collect_review_results.py --workbench 工作台.xlsx --require-complete   # 还有待核就 exit 1

复核意见的写法约定（脚本会自动抽取，写不清楚也能原样带出）：
    「改为：<新文本>；依据：<纸质版第X页 / 原图某处>」
    只写「改为：…」也可以；只写自由文字则整段落到 overrides 的 note 里，交人工判断。
"""
import argparse
import json
import re
import sys
from collections import Counter

try:
    import openpyxl
except ImportError:
    sys.exit("需要 openpyxl: pip install openpyxl")

VERDICTS = ("待核", "通过", "需修改", "存疑保留")
RE_FIX = re.compile(r"(?:改为|应改为|改成|修正为)\s*[:：]\s*(.+?)(?=(?:[；;]\s*(?:依据|理由|根据))|$)", re.S)
RE_BASIS = re.compile(r"(?:依据|理由|根据)\s*[:：]\s*(.+?)$", re.S)


def find_head(ws):
    rows = ws.iter_rows(min_row=1, max_row=1, values_only=True)
    r = next(rows, None)
    return [str(x or "").strip() for x in (r or [])]


def pick(head, rec, *names, default=""):
    for n in names:
        if n in head:
            i = head.index(n)
            if i < len(rec) and rec[i] is not None:
                return str(rec[i]).strip()
        for i, h in enumerate(head):
            if h.startswith(n) and i < len(rec) and rec[i] is not None:
                return str(rec[i]).strip()
    return default


def main():
    ap = argparse.ArgumentParser(description="回收人工复审结果 → overrides.json")
    ap.add_argument("--workbench", required=True, help="填好的复审工作台 xlsx")
    ap.add_argument("--out", default="overrides.json", help="修正清单输出路径")
    ap.add_argument("--report", default="复审结果汇总.md", help="汇总报告输出路径（空串则不写）")
    ap.add_argument("--corpus", default="", help="语料库 xlsx（可选：用于补全原文）")
    ap.add_argument("--require-complete", action="store_true",
                    help="存在「待核」时以非零退出（用于卡验收）")
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.workbench, data_only=True)
    sheet = next((s for s in wb.sheetnames if s.startswith("01")), None) or \
        next((s for s in wb.sheetnames if "队列" in s), None)
    if not sheet:
        sys.exit(f"找不到复审队列页签（现有：{wb.sheetnames}）")
    ws = wb[sheet]
    head = find_head(ws)
    if "结论" not in head and not any(h.startswith("结论") for h in head):
        sys.exit(f"「{sheet}」没有「结论」列，表头：{head}")

    recs = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not any(str(x or "").strip() for x in row):
            continue
        verdict = pick(head, row, "结论") or "待核"
        rec = dict(
            seq=pick(head, row, "序"),
            prio=pick(head, row, "优先级"),
            score=pick(head, row, "不确定度"),
            kind=pick(head, row, "对象类型"),
            where=pick(head, row, "对象定位"),
            parent=pick(head, row, "所属条文", "父条"),
            md_file=pick(head, row, "md文件"),
            md_loc=pick(head, row, "md内位置"),
            verdict=verdict,
            reviewer=pick(head, row, "复核人"),
            date=pick(head, row, "日期"),
            note=pick(head, row, "复核意见"),
            content=pick(head, row, "待判读内容"),
        )
        recs.append(rec)
    wb.close()

    # ---------- 统计 ----------
    by_verdict = Counter(r["verdict"] for r in recs)
    by_kind = Counter((r["kind"], r["verdict"]) for r in recs)
    fix_list, keep_list, pending = [], [], []
    for r in recs:
        if r["verdict"] == "需修改":
            m = RE_FIX.search(r["note"] or "")
            b = RE_BASIS.search(r["note"] or "")
            fix_list.append(dict(
                kind=r["kind"], where=r["where"], parent=r["parent"],
                md_file=r["md_file"], md_loc=r["md_loc"],
                original=(r["content"] or "")[:300],
                corrected=(m.group(1).strip() if m else ""),
                basis=(b.group(1).strip() if b else ""),
                note=r["note"], reviewer=r["reviewer"], date=r["date"],
                priority=r["prio"], uncertainty=r["score"],
            ))
        elif r["verdict"] == "存疑保留":
            keep_list.append(dict(kind=r["kind"], where=r["where"], note=r["note"],
                                  md_file=r["md_file"], md_loc=r["md_loc"]))
        elif r["verdict"] == "待核":
            pending.append(dict(kind=r["kind"], where=r["where"], priority=r["prio"]))

    unknown = [r for r in recs if r["verdict"] not in VERDICTS]
    payload = dict(
        workbench=a.workbench,
        total=len(recs),
        by_verdict=dict(by_verdict),
        fixes=fix_list,
        keep_as_is=keep_list,
        pending=pending,
        unknown_verdicts=[dict(where=r["where"], verdict=r["verdict"]) for r in unknown],
    )
    with open(a.out, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)

    # ---------- 报告 ----------
    if a.report:
        L = ["# 人工复审结果汇总\n",
             f"- 工作台：`{a.workbench}`",
             f"- 对象总数：**{len(recs)}**",
             "- 结论分布：" + "　".join(f"{v} {by_verdict.get(v, 0)}" for v in VERDICTS),
             f"- 待执行修正：**{len(fix_list)}** 项；存疑保留：{len(keep_list)} 项；"
             f"未判读：{len(pending)} 项\n"]
        if unknown:
            L.append("> ⚠️ 有 %d 行的结论不在四取值内（%s），已原样跳过\n"
                     % (len(unknown), sorted({r["verdict"] for r in unknown})))

        L.append("## 一、按对象类型 × 结论\n")
        L.append("| 类型 | 对象数 | " + " | ".join(VERDICTS) + " |")
        L.append("|---" * (2 + len(VERDICTS)) + "|")
        for kind in sorted({r["kind"] for r in recs}):
            n = sum(1 for r in recs if r["kind"] == kind)
            cells = " | ".join(str(by_kind.get((kind, v), 0)) for v in VERDICTS)
            L.append(f"| {kind} | {n} | {cells} |")
        L.append("")

        if fix_list:
            L.append("## 二、需修改清单（回灌用）\n")
            L.append("| # | 类型 | 定位 | md 位置 | 复核意见 | 抽取到的「改为」 | 依据 |")
            L.append("|---|---|---|---|---|---|---|")
            for i, f in enumerate(fix_list, 1):
                L.append(f"| {i} | {f['kind']} | {f['where']} | "
                         f"`{f['md_file']}` {f['md_loc']} | {(f['note'] or '').replace('|', '／')[:80]} | "
                         f"{(f['corrected'] or '—')[:40]} | {(f['basis'] or '—')[:40]} |")
            L.append("")
            no_fix = [f for f in fix_list if not f["corrected"]]
            if no_fix:
                L.append(f"> ⚠️ 有 {len(no_fix)} 条「需修改」没按 `改为：<新文本>` 写，"
                         f"脚本抽不出新文本，需人工补：{', '.join(x['where'] for x in no_fix[:8])}\n")
        if keep_list:
            L.append("## 三、存疑保留（不改原文，仅登记）\n")
            for k in keep_list:
                L.append(f"- **{k['kind']} {k['where']}**：{k['note'] or '（未写理由）'}")
            L.append("")
        if pending:
            L.append("## 四、尚未判读\n")
            L.append(f"共 {len(pending)} 项：" +
                     "、".join(f"{p['kind']}{p['where']}({p['priority']})" for p in pending[:20]) +
                     ("……" if len(pending) > 20 else ""))
            L.append("")

        L.append("## 五、回灌方式\n")
        L.append("1. 维护人读 `%s` 的 `fixes[]`，按类型改对应脚本：条文→建库脚本修正表、"
                 "表格→表格转录数据、公式→公式数据；\n" % a.out)
        L.append("2. 改完重跑：建库 → `lint_schema.py` → `validate_chunks.py` → "
                 "`scan_defects.py` → `crosscheck_corpus.py`；")
        L.append("3. 四项指标（漏条 0 / 编造 0 / 正文保真 100% / 切片对账一致）不退化才算回灌成功。")
        with open(a.report, "w", encoding="utf-8") as fh:
            fh.write("\n".join(L))

    print(f"复审对象 {len(recs)} 个｜" + "　".join(f"{v} {by_verdict.get(v, 0)}" for v in VERDICTS))
    print(f"需修改 {len(fix_list)} 项（抽到新文本 {sum(1 for f in fix_list if f['corrected'])} 项）｜"
          f"存疑保留 {len(keep_list)} 项｜未判读 {len(pending)} 项")
    print(f"修正清单：{a.out}")
    if a.report:
        print(f"汇总报告：{a.report}")

    if a.require_complete and pending:
        print(f"❌ 还有 {len(pending)} 个对象未判读（--require-complete）")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
