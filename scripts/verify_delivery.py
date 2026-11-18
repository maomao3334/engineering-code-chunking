# -*- coding: utf-8 -*-
"""一条命令跑完交付前四关，汇总"四个关键数字 + 退出码"。

为什么要有它：四关是四个脚本，靠人记顺序、靠眼看输出，迟早漏跑一个。
这个脚本不改变任何检查逻辑，只做编排与汇总——它报的每个数字都来自子脚本的原始输出。

四关（任一失败即非零退出）：
    1. schema 体检     lint_schema.py       页签/列/来源标注齐不齐
    2. 结构校验        validate_chunks.py   切片 ID、款数、父块含全款、款正文洁净度、公式占位
    3. 缺陷扫描        scan_defects.py      12 类字符缺陷（高严重度即失败）
    4. 独立交叉复核    crosscheck_corpus.py 换代码路径重算：漏条/编造/断号/串条/截断/对账
   （可选）复审工作台   build_review_workbench.py   生成按不确定度排序的复审队列

用法：
    python verify_delivery.py --corpus GB50193-93_2010_规范语料库.xlsx
    python verify_delivery.py --corpus 语料库.xlsx --pdf 规范.pdf          # 含第 4 关
    python verify_delivery.py --corpus 语料库.xlsx --skip-workbench --json 验收.json
"""
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent


def run(cmd):
    r = subprocess.run([sys.executable] + cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def grab(pattern, text, default="—"):
    m = re.search(pattern, text)
    return m.group(1) if m else default


def main():
    ap = argparse.ArgumentParser(description="交付前一键验收（四关）")
    ap.add_argument("--corpus", required=True, help="语料库 xlsx")
    ap.add_argument("--pdf", default="", help="源 PDF（提供则跑第 4 关独立交叉复核）")
    ap.add_argument("--workdir", default="", help="中间产物目录（默认 <语料库目录>/_验收）")
    ap.add_argument("--skip-workbench", action="store_true", help="不生成复审工作台")
    ap.add_argument("--skip-crosscheck", action="store_true", help="不跑独立交叉复核")
    ap.add_argument("--allow", default="",
                    help="独立复核的已知疑点条号，逗号分隔（如 5.1.1A），透传给 crosscheck_corpus")
    ap.add_argument("--json", dest="json_out", default="", help="机器可读验收结果")
    ap.add_argument("--export-dir", default="",
                    help="导出目录（提供则跑第⑥关：md 表格/公式格式体检 + 目录命名）")
    a = ap.parse_args()

    if not os.path.isfile(a.corpus):
        sys.exit(f"找不到 {a.corpus}")
    work = Path(a.workdir) if a.workdir else Path(a.corpus).parent / "_验收"
    work.mkdir(parents=True, exist_ok=True)

    steps = []          # (名称, ok, 摘要, 原始输出尾部)

    # 1. schema
    code, out = run([str(HERE / "lint_schema.py"), a.corpus])
    steps.append(("① schema 体检", code == 0,
                  f"错误 {out.count('❌')} 项 / 警告 {out.count('⚠️')} 项", out))

    # 2. 结构
    code, out = run([str(HERE / "validate_chunks.py"), a.corpus])
    kuan = grab(r"款检查：([^\n]+)", out)
    formula = grab(r"公式占位：([^\n]+)", out)
    slices = grab(r"切片：(\d+ = 父块 \d+ \+ 子块 \d+[^\n]*)", out)
    steps.append(("② 结构校验", code == 0,
                  f"切片 {slices}｜{kuan}｜{formula}", out))

    # 3. 缺陷扫描
    dscan = work / "缺陷扫描报告.md"
    djson = work / "defects.json"
    code, out = run([str(HERE / "scan_defects.py"), a.corpus,
                     "--out", str(dscan), "--json", str(djson)])
    total = grab(r"命中总数：\*\*(\d+)\*\*", out, "?")
    high = grab(r"## 二、高严重度问题（(\d+) 处", out, "?")
    steps.append(("③ 缺陷扫描", code == 0, f"命中 {total} 处（高严重度 {high} 处）", out))

    # 4. 独立交叉复核（需要 PDF）
    if a.pdf and not a.skip_crosscheck:
        cmd = [str(HERE / "crosscheck_corpus.py"), "--pdf", a.pdf, "--xlsx", a.corpus]
        if a.allow:
            cmd += ["--allow", a.allow]
        code, out = run(cmd)
        miss = grab(r"1\. 漏条（PDF 有、语料库无）:\s*(\S+)", out)
        fake = grab(r"2\. 编造（语料库有、PDF 无）:\s*(\S+)", out)
        recon = grab(r"7\. 切片对账:\s*([^\n]+)", out)
        steps.append(("④ 独立交叉复核", code == 0,
                      f"漏条 {miss} / 编造 {fake}｜{recon}", out))
    elif a.pdf:
        steps.append(("④ 独立交叉复核", True, "已按参数跳过", ""))
    else:
        steps.append(("④ 独立交叉复核", None, "未提供 --pdf，跳过（交付前建议补上）", ""))

    # 5. 复审工作台
    wb_out = work / "人工复审工作台.xlsx"
    mf_out = work / "review_manifest.json"
    if not a.skip_workbench:
        code, out = run([str(HERE / "build_review_workbench.py"), a.corpus,
                         "--out", str(wb_out), "--manifest", str(mf_out)])
        objs = grab(r"复审对象 (\d+) 个", out, "?")
        dist = grab(r"复审对象 \d+ 个｜([^\n]+)", out, "—")
        md_hit = grab(r"（(\d+)/\d+ 有 md 定位）", out, "—")
        steps.append(("⑤ 复审工作台", code == 0,
                      f"对象 {objs} 个｜{dist}", out))
    else:
        steps.append(("⑤ 复审工作台", None, "已按参数跳过", ""))

    # 6. 导出目录的 md 格式体检（表格空行 / 公式定界 / 目录命名）
    #    为什么单独一关：这两类问题**不影响检索召回**，只在人打开 md 时暴露，
    #    所以前四关全绿也可能带着它们交付（实测踩过：Obsidian 里表不渲染）。
    if a.export_dir:
        ed = Path(a.export_dir)
        if not ed.is_dir():
            steps.append(("⑥ md 格式体检", False, f"目录不存在：{ed}", ""))
        else:
            c1, o1 = run([str(HERE / "md_normalize.py"), str(ed), "--check"])
            c2, o2 = run([str(HERE / "normalize_formulas.py"), str(ed), "--check"])
            bad_t = grab(r"共 (\d+) 个文件有问题", o1, "0")
            bad_f = grab(r"共 (\d+) 个文件有问题", o2, "0")
            # 目录命名：应以规范编号命名（含 GB/GBT 编号），不能叫「导出」「output」
            name_ok = bool(re.search(r"GB\s*/?\s*T?\s*\d", ed.name))
            ok6 = (c1 == 0 and c2 == 0 and name_ok)
            brief = (f"表格异常 {bad_t} 个｜公式异常 {bad_f} 个｜"
                     f"目录名 {'合规' if name_ok else '未含国标编号'}")
            steps.append(("⑥ md 格式体检", ok6, brief, o1 + "\n" + o2))
    else:
        steps.append(("⑥ md 格式体检", None, "未提供 --export-dir，跳过", ""))

    # ---------- 汇总 ----------
    W = 78
    print("=" * W)
    print(f"交付验收：{Path(a.corpus).name}")
    print("=" * W)
    failed = 0
    for name, ok, brief, out in steps:
        mark = "✅" if ok is True else ("⚠️ " if ok is None else "❌")
        if ok is False:
            failed += 1
        print(f"{mark} {name}　{brief}")
        if ok is False:
            tail = [ln for ln in out.strip().splitlines() if ln.strip()][-6:]
            for ln in tail:
                print(f"      {ln}")
    print("-" * W)
    if failed:
        print(f"❌ 验收未通过：{failed} 关失败（明细见 {work}）")
    else:
        print("✅ 核心四关通过：" + "；".join(s[2] for s in steps[:4]))
        rest = [s for s in steps[4:] if s[1] is True]
        if rest:
            print("   " + "；".join(f"{s[0]} {s[2]}" for s in rest))
        print(f"   产物目录：{work}")
    print("=" * W)

    payload = dict(corpus=a.corpus, pdf=a.pdf, workdir=str(work),
                   failed=failed,
                   steps=[dict(name=n, ok=o, brief=b) for n, o, b, _ in steps])
    if a.json_out:
        Path(a.json_out).write_text(json.dumps(payload, ensure_ascii=False, indent=1),
                                    encoding="utf-8")
        print(f"机器可读：{a.json_out}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
