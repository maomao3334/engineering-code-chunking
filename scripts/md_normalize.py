# -*- coding: utf-8 -*-
"""Markdown 可读性规范化：让导出的 .md 在 Obsidian / Typora / VS Code 里**直接能看表**

为什么需要它（实测踩过）：
    `raglib` 渲染出来的切片长这样——

        **表3.2.1 单层、多层民用建筑…**（源：PDF 第 7 页）
        | 建筑物及场所 | 建筑规模性质 | … |
        |---|---|---|

    标题行和表体之间**没有空行**。GFM（Obsidian 用的就是 GFM）规定
    **表格不能打断段落**，所以这一坨会被当成一个段落，表根本不渲染，
    必须手动敲一个空行才显示成表。

    另外「数据行」是伪表格：

        行标签 | 1 | 2 | 3
        候机楼 | A | A | B1

    行首没有竖线，也不是合法 GFM 表格 —— 在 Obsidian 里就是一堆竖线文本。

本脚本干四件事（幂等，跑两遍结果一样）：
    1. 真表格块：**前后补空行**；缺分隔行就补；列数不齐就补 `—`
    2. 伪表格块（`行标签 | 1 | 2` 这种）→ 转成合法 GFM 表格
    3. 元数据头 / 表标题 / 注： / 【转录提示】 前后补空行
    4. 连续空行压成一行

用法：
    python md_normalize.py <目录或文件>                 # 只体检，不改
    python md_normalize.py <目录或文件> --in-place      # 就地修
    python md_normalize.py <目录> --out 修好的目录       # 另存
    python md_normalize.py <目录> --check               # 只报问题数，退出码非零=有问题
"""
import argparse
import re
import sys
from pathlib import Path

# Windows 控制台默认 GBK，会因 ✅/❌ 抛 UnicodeEncodeError（实测踩过）
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:
        pass

# ---------------------------------------------------------------- 基础判定
SEP_CELL = re.compile(r"^\s*:?-{2,}:?\s*$")


def _cells(line):
    """把一行按竖线拆成单元格（去掉首尾空单元）。"""
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def _is_real_table_line(line):
    """以竖线开头的行 = 真表格行（GFM 要求行首有 `|`）。"""
    return line.lstrip().startswith("|")


def _is_sep(line):
    if not _is_real_table_line(line):
        return False
    cs = _cells(line)
    return bool(cs) and all(SEP_CELL.match(c) for c in cs)


def _looks_like_pseudo_row(line):
    """`行标签 | 1 | 2 | 3`：没有行首竖线、但有 ≥2 个竖线、每格都不长。

    为什么要限长度：正文里偶尔出现「A｜B」这类跨格符号，
    但中文正文的竖线两侧通常有空格且格很长，不会被误判。
    """
    if _is_real_table_line(line):
        return False
    if line.count("|") < 2:
        return False
    cs = _cells(line)
    if len(cs) < 2:
        return False
    if any(len(c) > 40 for c in cs):
        return False
    # 全空的格子占比过高说明不是表
    if sum(1 for c in cs if not c) > len(cs) // 2:
        return False
    return True


def _sep_line(n):
    return "|" + "|".join(["---"] * max(n, 1)) + "|"


def _pad_row(row, n):
    cs = _cells(row)
    if len(cs) < n:
        cs += ["—"] * (n - len(cs))
    elif len(cs) > n:
        cs = cs[:n]
    return "| " + " | ".join(cs) + " |"


# ---------------------------------------------------------------- 主变换
def normalize_text(text):
    """返回 (新文本, 统计dict)。幂等。"""
    stats = dict(blank=0, table=0, pseudo=0, sep=0, pad=0)
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    # --- 1. 连续空行压成一行（先做，后面补空行才不会被误删）
    out = []
    for ln in lines:
        if ln.strip() == "":
            if out and out[-1].strip() == "":
                continue
            out.append("")
        else:
            out.append(ln)
    lines = out

    # --- 2. 伪表格 → 真表格（连续 ≥2 行、竖线数一致）
    i = 0
    conv = []
    while i < len(lines):
        if _looks_like_pseudo_row(lines[i]):
            j = i
            while j < len(lines) and _looks_like_pseudo_row(lines[j]):
                j += 1
            grp = lines[i:j]
            counts = {ln.count("|") for ln in grp}
            if len(grp) >= 2 and len(counts) == 1:
                n = len(_cells(grp[0]))
                conv.append(_pad_row(grp[0], n))
                conv.append(_sep_line(n))
                for ln in grp[1:]:
                    conv.append(_pad_row(ln, n))
                stats["pseudo"] += 1
                i = j
                continue
        conv.append(lines[i])
        i += 1
    lines = conv

    # --- 3. 真表格块：补分隔行、列数对齐、前后空行
    i = 0
    out = []
    while i < len(lines):
        if _is_real_table_line(lines[i]):
            j = i
            while j < len(lines) and _is_real_table_line(lines[j]):
                j += 1
            block = lines[i:j]
            stats["table"] += 1
            # 分隔行必须在第 2 行
            if len(block) >= 1 and not (len(block) >= 2 and _is_sep(block[1])):
                n0 = len(_cells(block[0]))
                block = [block[0], _sep_line(n0)] + block[1:]
                stats["sep"] += 1
            # 列数对齐（以分隔行为准，最长者）
            n = max(len(_cells(b)) for b in block)
            fixed = []
            for k, b in enumerate(block):
                if k == 1:
                    fixed.append(b if len(_cells(b)) == n else _sep_line(n))
                else:
                    if len(_cells(b)) != n:
                        stats["pad"] += 1
                    fixed.append(_pad_row(b, n))
            block = fixed
            # 前空行
            if out and out[-1].strip() != "":
                out.append("")
                stats["blank"] += 1
            out.extend(block)
            # 后空行
            if j < len(lines) and lines[j].strip() != "":
                out.append("")
                stats["blank"] += 1
            i = j
            continue
        out.append(lines[i])
        i += 1
    lines = out

    # --- 4. 元数据头 / 表标题 / 注： / 转录提示 前后空行
    KEY = re.compile(r"^(\*\*表|表\s*\d+(\.\d+)*\s|注：|【转录提示】|【表|【附表)")
    META = re.compile(r"^《[^》]+》.*｜")
    out = []
    for k, ln in enumerate(lines):
        need_before = False
        if k > 0 and lines[k - 1].strip() != "":
            if (KEY.match(ln.strip()) or META.match(ln.strip())
                    or re.match(r"^\*\*附表", ln.strip())):
                need_before = True
        if need_before:
            out.append("")
            stats["blank"] += 1
        out.append(ln)
        # 表标题后必须有空行（下一行是表体时尤其重要）
        if k + 1 < len(lines) and lines[k + 1].strip() != "":
            s = ln.strip()
            if (re.match(r"^\*\*表", s) or re.match(r"^\*\*附表", s)
                    or re.match(r"^【表", s) or re.match(r"^【附表", s)
                    or re.match(r"^注：", s) or re.match(r"^【转录提示】", s)
                    or META.match(s)):
                out.append("")
                stats["blank"] += 1
    lines = out

    # --- 5. 收尾：去首尾空行、连续空行再压一次
    text = "\n".join(lines).strip("\n") + "\n"
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text, stats


# ---------------------------------------------------------------- 体检
def check_text(text):
    """返回问题列表（用于 --check）。"""
    probs = []
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if _is_real_table_line(ln):
            prev = lines[i - 1] if i else ""
            if i and prev.strip() != "" and not _is_real_table_line(prev):
                probs.append((i + 1, "表格前缺空行"))
            if i and _is_real_table_line(prev) is False and _is_sep(ln):
                probs.append((i + 1, "分隔行位置不对（应为表格第2行）"))
            break_flag = False
            j = i
            while j < len(lines) and _is_real_table_line(lines[j]):
                j += 1
            if j < len(lines) and lines[j].strip() != "":
                probs.append((j + 1, "表格后缺空行"))
            # 缺分隔行
            if i + 1 < len(lines) and not _is_sep(lines[i + 1]):
                probs.append((i + 2, "表格缺分隔行 |---|"))
            break
    for i, ln in enumerate(lines):
        if _looks_like_pseudo_row(ln):
            probs.append((i + 1, "伪表格（行首缺竖线，Obsidian 不渲染）"))
    return probs


def main():
    ap = argparse.ArgumentParser(description="Markdown 可读性规范化（表格空行/伪表格转真表）")
    ap.add_argument("path", help=".md 文件或目录")
    ap.add_argument("--in-place", action="store_true", help="就地修改（默认只预览统计）")
    ap.add_argument("--out", default="", help="另存目录（与 --in-place 二选一）")
    ap.add_argument("--check", action="store_true", help="只体检，不修改；有问题退出码 1")
    ap.add_argument("--quiet", action="store_true", help="只打印汇总")
    a = ap.parse_args()

    p = Path(a.path)
    if not p.exists():
        sys.exit(f"找不到 {p}")
    files = [p] if p.is_file() else sorted(p.rglob("*.md"))
    if not files:
        sys.exit(f"{p} 下没有 .md")

    tot = dict(blank=0, table=0, pseudo=0, sep=0, pad=0, files=0, changed=0)
    probs_all = []
    outdir = Path(a.out) if a.out else None
    if outdir:
        outdir.mkdir(parents=True, exist_ok=True)

    for f in files:
        raw = f.read_text(encoding="utf-8")
        if a.check:
            probs = check_text(raw)
            if probs:
                probs_all.append((f, probs))
                if not a.quiet:
                    print(f"\n❌ {f.name}（{len(probs)} 处）")
                    for ln, msg in probs[:8]:
                        print(f"   L{ln}: {msg}")
            continue
        new, st = normalize_text(raw)
        tot["files"] += 1
        for k in ("blank", "table", "pseudo", "sep", "pad"):
            tot[k] += st[k]
        if new != raw:
            tot["changed"] += 1
            if a.in_place:
                f.write_text(new, encoding="utf-8")
            elif outdir:
                (outdir / f.name).write_text(new, encoding="utf-8")
            if not a.quiet and not outdir:
                print(f"  {f.name}: 空行+{st['blank']} 表{st['table']} "
                      f"伪表{st['pseudo']} 补分隔{st['sep']} 补列{st['pad']}")

    if a.check:
        if probs_all:
            print(f"\n共 {len(probs_all)} 个文件有问题（共 "
                  f"{sum(len(x[1]) for x in probs_all)} 处）")
            print("修法：python md_normalize.py <该目录> --in-place")
            return 1
        print(f"✅ {len(files)} 个 md 表格格式全部合规")
        return 0

    mode = "已就地修改" if a.in_place else ("已另存到 " + str(outdir) if outdir else "预览（未写盘）")
    print(f"\n{mode}：{tot['files']} 个文件，改动 {tot['changed']} 个")
    print(f"  补空行 {tot['blank']}｜规范化表格 {tot['table']} 张｜"
          f"伪表格→真表 {tot['pseudo']} 处｜补分隔行 {tot['sep']}｜补列 {tot['pad']}")
    if not a.in_place and not outdir:
        print("  要写盘：加 --in-place（就地）或 --out <目录>（另存）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
