# -*- coding: utf-8 -*-
"""生成「人工复审工作台」：把语料库里**确定性不足**的对象排成一条队列，最不确定的排最上面。

为什么按不确定度排、而不按条号排
---------------------------------
人工判读是整条交付链上最贵的资源。按条号顺序逐条看，等于把 90% 的注意力花在
「文字层直接抽取、独立复核已通过」的低风险条文上；而真正会出错的地方
（图片表格读数、公式还原、原书自相矛盾处、几何识别的修订标记）被摊平在几百行里，
很容易被"看到后面看麻了"漏掉。所以队列必须按**出错概率**降序，而不是按文档顺序。

不确定度怎么算（0-100，可解释、可复核）
--------------------------------------
每个对象带若干「不确定来源」标签，每类标签给固定分，叠加后截断到 100：

  来源可靠性（最高 +40）
    图片表格视觉转录        +30    文字层抽取          0
    公式由渲染图还原        +30    几何识别（下划线）  +25
    扫描页截断/污损         +20

  已发现的内部矛盾（最高 +25）
    原书自相矛盾（表值 vs 表注/量级）  +25
    与他人交叉校验出分歧              +20

  已登记的存疑（最高 +45）
    严重度高 +45 / 中 +30 / 低 +15

  结构性脆弱点（累加，最高 +30）
    编号非 X.Y.Z（A/B 后缀）  +8     跨页条文          +6
    结尾非句读                +10    已删除条          +12
    公式占位（不可引用）      +10    页级元素混排      +5
    附录整册（无款级结构）    +8     切片形态待验      +5

  复核覆盖缺口（最高 +25）
    该方法原理上覆盖不到（如修订漏标）  +25
    无独立交叉校验                      +15

优先级分档：P0 ≥85（必须先看）/ P1 60-84 / P2 35-59 / P3 <35（可批量扫过）

用法
----
    python scripts/build_review_workbench.py 语料库.xlsx
    python scripts/build_review_workbench.py 语料库.xlsx --out 复审工作台.xlsx
    python scripts/build_review_workbench.py 语料库.xlsx --top-slices 20

只读 xlsx，不改语料库；输出一个独立工作簿。缺页签时自动降级（只有「语料库」「RAG切片」也能跑）。
"""
import argparse
import os
import re
import sys

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:  # pragma: no cover
    sys.exit("需要 openpyxl：pip install openpyxl")

VERDICTS = ["待核", "通过", "需修改", "存疑保留"]
PRIORITY_FILL = {
    "P0": PatternFill("solid", fgColor="F8CBAD"),   # 红：必须先看
    "P1": PatternFill("solid", fgColor="FFE699"),   # 黄
    "P2": PatternFill("solid", fgColor="E2EFDA"),   # 绿
    "P3": PatternFill("solid", fgColor="F2F2F2"),   # 灰
}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
TEXT = Font(size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

NAT = re.compile(r"(\d+)")


def natkey(s):
    """自然排序：3.2.10 排在 3.2.9 之后；附录A 排在附录B 之前。"""
    s = str(s or "")
    out = []
    for part in NAT.split(s):
        if part.isdigit():
            out.append((1, int(part), ""))
        else:
            out.append((0, 0, part))
    return out


def load(wb, name):
    if name not in wb.sheetnames:
        return [], []
    rows = list(wb[name].iter_rows(values_only=True))
    if not rows:
        return [], []
    head = [str(x or "").strip() for x in rows[0]]
    recs = [dict(zip(head, r)) for r in rows[1:] if any(v is not None and v != "" for v in r)]
    return head, recs


def sval(rec, *names, default=""):
    """按候选列名取值（不同版本语料库列名可能略有差异）。"""
    for n in names:
        if n in rec and rec[n] is not None and str(rec[n]).strip() != "":
            return str(rec[n]).strip()
    return default


# ------------------------------------------------- 导出 md 定位（复核时直接翻文件）
def build_md_index(xlsx_path, md_dir_arg=""):
    """定位导出的 markdown 目录，建立「条号 → 文件 / 行」索引。

    返回 (root_label, files, catalog)：
      root_label  用于展示的相对根（如 `导出_平台导入/markdown`）
      files       {文件名: [行文本, ...]}
      catalog     {条号: 文件名}（优先读 json/catalog.json，其次按 `NNN_条号.md` 推断）
    找不到 md 目录时返回 ("", {}, {})，工作台该列留空，不报错。
    """
    base = os.path.dirname(os.path.abspath(xlsx_path))
    cands = []
    if md_dir_arg:
        cands.append(os.path.abspath(md_dir_arg))
    for sub in ("导出_平台导入/markdown", "markdown", "export/markdown",
                "导出_平台导入/Markdown", "md"):
        cands.append(os.path.join(base, sub.replace("/", os.sep)))
    root = next((c for c in cands if os.path.isdir(c)), "")
    if not root:
        return "", {}, {}

    files = {}
    for fn in sorted(os.listdir(root)):
        if not fn.lower().endswith(".md"):
            continue
        try:
            with open(os.path.join(root, fn), encoding="utf-8", errors="replace") as fh:
                files[fn] = fh.read().splitlines()
        except OSError:
            continue

    catalog = {}
    for jf in (os.path.join(os.path.dirname(root), "json", "catalog.json"),
               os.path.join(root, "catalog.json")):
        if not os.path.isfile(jf):
            continue
        try:
            import json
            with open(jf, encoding="utf-8") as fh:
                for it in json.load(fh):
                    no, fn = it.get("no"), it.get("file")
                    if no and fn in files:
                        catalog[str(no)] = fn
        except (OSError, ValueError, AttributeError, TypeError):
            pass
        if catalog:
            break
    if not catalog:                     # 退化：从文件名 `NNN_条号.md` 推
        for fn in files:
            m = re.match(r"^\d+_(.+)\.md$", fn)
            if m:
                catalog[m.group(1)] = fn

    label = os.path.relpath(root, base).replace(os.sep, "/")
    return label, files, catalog


def _find(files, catalog, no, pattern, limit=3):
    """在某条对应的 md 里找首个匹配行号（1 基）。pattern 为已编译正则。"""
    fn = catalog.get(str(no))
    if not fn:
        return ""
    for i, ln in enumerate(files.get(fn, []), start=1):
        if pattern.search(ln):
            return f"L{i}"
    return ""


def locate_md(kind, obj, files, catalog, root_label):
    """把复审对象映射到「导出 md 的文件 + 行号」。返回 (文件相对路径, 位置说明)。"""
    if not files or not catalog:
        return "", ""

    def path(fn):
        return f"{root_label}/{fn}" if root_label else fn

    # --- 条文 / 款：父条号的 md
    if kind == "条文":
        no = obj["where"]
        fn = catalog.get(no)
        if not fn:
            return "", ""
        lines = files[fn]
        body = ""
        for i, ln in enumerate(lines, start=1):
            if re.match(r"^\s*" + re.escape(no) + r"(?![.\d])", ln):
                body = f"L{i}"
                break
        return path(fn), f"元数据头 L1；正文 {body or 'L3 起'}"

    # --- 表格：`**表号 …**` 起始，到表格结束
    if kind == "表格":
        tno = obj["where"]
        fn = catalog.get(str(obj.get("parent") or ""))
        if not fn:
            return "", ""
        lines = files[fn]
        start = next((i for i, ln in enumerate(lines, start=1)
                      if tno in ln and ln.lstrip().startswith("**")), None)
        if not start:
            start = next((i for i, ln in enumerate(lines, start=1) if tno in ln), None)
        if not start:
            return path(fn), ""
        end, seen_row = start, False
        for i in range(start, len(lines)):          # start 是 0 基的下一行
            ln = lines[i].strip()
            if ln.startswith("|"):
                seen_row = True
                end = i + 1
            elif seen_row:
                break
        return path(fn), f"表块 L{start}–L{end}"

    # --- 公式：正文里的 `〔公式（x）〕` 占位
    if kind == "公式":
        fid = obj["where"]
        num = fid.strip("（）() ")
        fn = catalog.get(str(obj.get("parent") or ""))
        if not fn:
            return "", ""
        pat = re.compile(r"公式\s*[（(]\s*" + re.escape(num))
        hits = [i for i, ln in enumerate(files[fn], start=1) if pat.search(ln)]
        return path(fn), ("占位行 " + "、".join(f"L{i}" for i in hits[:3])) if hits else ""

    # --- 修订标记：按下划线文字在 md 正文里找
    if kind == "修订标记":
        no = obj["where"]
        fn = catalog.get(no)
        if not fn:
            return "", ""
        segs = [s for s in re.split(r"[｜|]", obj["content"]) if s.strip()]
        keys = []
        for s in segs:
            s = re.sub(r"^\[段\d+\]", "", s).strip()
            s = re.sub(r"^" + re.escape(no) + r"[：:]\s*", "", s)
            s = re.sub(r"[｜|].*$", "", s)
            if len(s) >= 4:
                keys.append(s[:12])
        line = ""
        for key in keys:
            line = _find(files, catalog, no, re.compile(re.escape(key)))
            if line:
                break
        if not line:                       # 定位不到就给正文起点，别留空
            line = _find(files, catalog, no,
                         re.compile(r"^\s*" + re.escape(no) + r"(?![.\d])")) or "L3 起"
        return path(fn), line

    # --- 图：图号出现在 md 正文（题注或正文引用）
    if kind == "图":
        fid = obj["where"]
        fn = catalog.get(str(obj.get("parent") or ""))
        if not fn:
            return "", ""
        hits = [i for i, ln in enumerate(files[fn], start=1) if fid in ln]
        return path(fn), ("出现于 " + "、".join(f"L{i}" for i in hits[:3])) if hits else ""

    # --- 切片：按层级回到父 md
    if kind == "切片":
        sid = obj["where"]
        no = str(obj.get("parent") or "")
        fn = catalog.get(no)
        if not fn:
            return "", ""
        tail = sid.split("#", 1)[1] if "#" in sid else ""
        if tail.startswith("表"):
            sub = dict(obj)
            sub["where"] = tail
            return locate_md("表格", sub, files, catalog, root_label)
        if tail and re.match(r"^\d+\.\d+\.\d+[A-Z]?\.\d+$", tail):
            return path(fn), _find(files, catalog, no,
                                   re.compile(r"^\s*" + re.escape(tail) + r"(?![\d])"))
        return locate_md("条文", dict(obj, where=no), files, catalog, root_label)

    return "", ""


# ---------------------------------------------------------------- 打分
def score_formulas(formulas):
    objs = []
    for f in formulas:
        fid = sval(f, "公式编号", "公式号")
        conf = sval(f, "可信度", "可信度（高/中/低）")
        caveat = sval(f, "存疑说明", "疑点")
        parent = sval(f, "所属条文", "条号")
        page = sval(f, "PDF页", "页")
        score = 60
        tags = []
        if re.search(r"低|待人工|待复核|存疑", conf):
            score += 35
            tags.append("公式由渲染图还原·可信度低（+35）")
        elif re.search(r"中", conf):
            score += 20
            tags.append("公式按式中变量推定·可信度中（+20）")
        else:
            score -= 5
            tags.append("公式已用「式中」校验")
        if caveat:
            score += 5
            tags.append("已登记存疑（+5）")
        objs.append(dict(
            score=min(100, score), kind="公式", where=fid,
            parent=parent, page=page,
            content=f"{fid}｜所属条文 {parent}｜{sval(f, '规范化表达式（整理，非原文逐字）', '表达式')}",
            why="；".join(tags) + (f"　存疑：{caveat[:80]}" if caveat else ""),
        ))
    return objs


def score_tables(cells, page_of_parent):
    """单元格行 → 按表聚合。转录提示里含 图片/OCR/读图/渲染/扫描 ⇒ 图片表格。"""
    by_table = {}
    for c in cells:
        tno = sval(c, "表号")
        by_table.setdefault(tno, []).append(c)
    objs = []
    for tno, rows in by_table.items():
        caveat = " ".join(sval(r, "转录提示", "存疑说明") for r in rows)
        vals = [sval(r, "单元格值") for r in rows]
        src = sval(rows[0], "来源", "数据来源")
        cross = sval(rows[0], "交叉校验", "校验")
        score, tags = 25, []
        is_image = False
        if "文字层" in src or "非图片" in caveat:
            tags.append("文字层抽取")
        elif "图片" in src:
            is_image = True
        elif re.search(r"图片|OCR|读图|渲染|扫描|视觉", caveat):
            is_image = True
        elif not src and not caveat:
            tags.append("来源未标注·按文字层处理")
        if is_image:
            score += 35
            tags.append("表体为图片·视觉转录（+35）")
        # 有交叉校验就压住一部分不确定度；没有则加分
        if is_image and not cross and not re.search(r"OCR|ocr|比对|交叉|校验", caveat + src):
            score += 15
            tags.append("未见独立交叉校验（+15）")
        elif is_image and cross:
            score -= 5
            tags.append("有独立交叉校验（-5）")
        trunc = sum(1 for v in vals if re.search(r"[.．]$|\?$|…$|□|�", v))
        if trunc:
            add = min(20, 5 * trunc)
            score += add
            tags.append(f"疑似截断/污损 {trunc} 处（+{add}）")
        if re.search(r"不符|矛盾|异常|量级|疑原|排印", caveat):
            score += 25
            tags.append("原表内部不自洽（+25）")
        elif re.search(r"不一致|不统一|不规整", caveat):
            score += 10
            tags.append("原表格式不统一·需确认非转录错（+10）")
        if re.search(r"截断|污损|缺笔|不可辨|仅可辨|模糊|噪点", caveat):
            score += 10
            tags.append("扫描质量差/字符不可辨（+10）")
        if len(rows) >= 40:
            score += 5
            tags.append("长表·转录累积风险（+5）")
        title = sval(rows[0], "表标题")
        parent = sval(rows[0], "所属条文")
        # 预览前 3 行，方便判读时不打开语料库也能看出格式
        by_row = {}
        for r in rows:
            try:
                ri = int(float(sval(r, "行号") or 0))
            except ValueError:
                ri = 0
            by_row.setdefault(ri, []).append(sval(r, "单元格值"))
        preview = "；".join(" ｜ ".join(v for v in by_row[k] if v != "")
                           for k in sorted(by_row)[:3])
        objs.append(dict(
            score=min(100, score), kind="表格", where=tno,
            parent=parent, page=page_of_parent.get(parent, ""),
            content=f"{tno} {title}（整表 {len(by_row)} 行 / {len(rows)} 单元格；"
                    f"所属：{parent or '附录'}）｜预览：{preview[:150]}",
            why="；".join(tags),
        ))
    return objs


def score_clauses(corpus, kuan, suspect_sev, pagelevel_count):
    kuan_by_parent = {}
    for k in kuan:
        kuan_by_parent.setdefault(sval(k, "条号"), 0)
        kuan_by_parent[sval(k, "条号")] += 1
    objs = []
    for r in corpus:
        no = sval(r, "条文号")
        body = sval(r, "条文原文本体")
        typ = sval(r, "条文类型")
        p1 = sval(r, "PDF起页")
        p2 = sval(r, "PDF末页")
        score, tags = 20, []
        sev = suspect_sev.get(no)
        if sev:
            add = {"高": 45, "中": 30, "低": 15}.get(sev, 15)
            score += add
            tags.append(f"已登记存疑·严重度{sev}（+{add}）")
        if body and not re.search(r"[。；：？！”』）\]】]$|^（此条删除）", body.rstrip()):
            score += 10
            tags.append("结尾非句读·疑截断（+10）")
        if p1 and p2 and p1 != p2:
            score += 6
            tags.append(f"跨页 {p1}-{p2}（+6）")
        if re.search(r"\d[A-Z]$", no):
            score += 8
            tags.append("A/B 后缀编号·非 X.Y.Z（+8）")
        if "此条删除" in body:
            score += 12
            tags.append("已删除条·易被误判为漏条（+12）")
        if "〔公式" in body:
            score += 10
            tags.append("含公式占位·不可逐字引用（+10）")
        pc = pagelevel_count.get(no, 0)
        if pc:
            score += 5
            tags.append(f"含页级元素 {pc} 行·图题表注混排（+5）")
        if typ == "附录" or no.startswith("附录"):
            score += 8
            tags.append("附录整册·无款级结构（+8）")
        if typ == "术语" and re.search(r"[A-Za-z]{4,}", body):
            score += 5
            tags.append("术语条含英文原名（+5）")
        objs.append(dict(
            score=min(100, score), kind="条文", where=no,
            parent=no, page=p1,
            content=f"{no}（{sval(r, '章')}／{sval(r, '节')}，{len(body)} 字）"
                    + (f"｜{body[:110]}…" if len(body) > 110 else f"｜{body}"),
            why="；".join(tags) or "文字层抽取·结构正常",
        ))
    return objs


def score_revisions(revs):
    objs = []
    for r in revs:
        no = sval(r, "条文号")
        line = sval(r, "被下划线的原文行")
        segs = sval(r, "被下划线文字（按段拆分）", "被下划线文字")
        score, tags = 30, ["几何识别下划线·非文字层（+25）"]
        score += 25
        nseg = segs.count("[段") or 1
        if nseg > 1:
            score += 10
            tags.append(f"同行 {nseg} 段下划线·曾拼接出错（+10）")
        if len(re.sub(r"\s", "", line)) <= 24:
            score += 10
            tags.append("短行·易误判归属（+10）")
        objs.append(dict(
            score=min(100, score), kind="修订标记", where=no,
            parent=no, page=sval(r, "PDF页"),
            content=f"{no}：{line[:110]}｜下划线：{segs[:80]}",
            why="；".join(tags),
        ))
    return objs


def score_figures(figures):
    """图：只登记题注、未做数值化。风险在"图里的信息答不出来"，不在切分。"""
    objs = []
    for f in figures:
        fid = sval(f, "图表号", "图号")
        desc = sval(f, "内容说明", "说明")
        score, tags = 30, ["图未数值化·仅登记题注（+30）"]
        if re.search(r"诺模图|换算图|曲线|查图|坐标", desc):
            score += 20
            tags.append("图内曲线/诺模图·数值类问题答不了（+20）")
        if re.search(r"示意|布置", desc):
            score += 8
            tags.append("示意图·尺寸关系靠文字描述（+8）")
        objs.append(dict(
            score=min(100, score), kind="图", where=fid,
            parent=sval(f, "所属条文", "父条"), page=sval(f, "PDF页"),
            content=f"{fid} {sval(f, '图表名称', '图名')}（所属：{sval(f, '所属条文')}）｜{desc[:130]}",
            why="；".join(tags),
        ))
    return objs


def score_structure(struct):
    """结构核对：章/节/附录 的层级与数量。风险低，但缺了就没法证明"没漏一章"。"""
    objs = []
    for r in struct:
        lvl = sval(r, "层级")
        no = sval(r, "编号")
        title = sval(r, "标题")
        score, tags = 25, ["结构核对·成本低（+25）"]
        if lvl == "附录":
            score += 8
            tags.append("附录整册（+8）")
        if sval(r, "列项数") == "0" and lvl == "章":
            score += 5
            tags.append("该章无款级结构（+5）")
        objs.append(dict(
            score=min(100, score), kind="结构", where=no,
            parent=no, page=sval(r, "PDF起页"),
            content=f"{lvl} {no} {title}：条文 {sval(r, '条文数')} 条 / 款 {sval(r, '列项数')} 个 / "
                    f"{sval(r, '字数')} 字 / 范围 {sval(r, '条文号范围')}",
            why="；".join(tags),
        ))
    return objs


def score_slices(slices, obj_by_parent, top):
    """只给风险最高的父条/表配「切片形态待验」对象，且刻意低于父对象本身：
    正文还没判对之前，先验交付形态没有意义，避免队列顶部被切片挤占。"""
    out = []
    parents = sorted(
        [o for o in obj_by_parent.values()],
        key=lambda o: -o["score"])[: max(0, top)]
    seen = set()
    for o in parents:
        for s in slices:
            sid = sval(s, "切片ID")
            if sid in seen or sval(s, "父条号") != o["where"]:
                continue
            seen.add(sid)
            txt = sval(s, "切片全文")
            out.append(dict(
                score=max(30, int(o["score"] * 0.72)), kind="切片", where=sid,
                parent=o["where"], page=sval(s, "PDF起页"),
                content=f"{sval(s, '切片层级')}｜{sval(s, '正文字符数')} 字｜"
                        + txt[:110].replace("\n", " ⏎ "),
                why=f"继承父对象风险（{o['kind']} {o['where']} {o['score']}）×0.72："
                    f"验元数据头/全款/引用闭合",
            ))
            break
    return out


def priority(score):
    if score >= 82:
        return "P0"
    if score >= 60:
        return "P1"
    if score >= 35:
        return "P2"
    return "P3"


def main():
    ap = argparse.ArgumentParser(description="生成按不确定度降序排列的人工复审工作台")
    ap.add_argument("xlsx", help="语料库工作簿")
    ap.add_argument("--out", default="", help="输出路径，默认 <xlsx目录>/人工复审工作台.xlsx")
    ap.add_argument("--top-slices", type=int, default=12,
                    help="为风险最高的 N 个对象各配 1 个切片形态复核项（默认 12）")
    ap.add_argument("--md-dir", default="",
                    help="导出的 markdown 目录；默认自动找 <xlsx目录>/导出_平台导入/markdown")
    ap.add_argument("--only", default="",
                    help="只导出指定优先级，逗号分隔（如 P0 或 P0,P1）")
    ap.add_argument("--kind", default="",
                    help="只导出指定对象类型，逗号分隔（条文/表格/公式/修订标记/图/切片/结构）")
    ap.add_argument("--manifest", default="",
                    help="另存机器可读 JSON（打分依据 + 存疑清单），默认不写")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()

    if not os.path.isfile(a.xlsx):
        sys.exit(f"找不到 {a.xlsx}")
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.xlsx)),
                                "人工复审工作台.xlsx")

    wb = openpyxl.load_workbook(a.xlsx, read_only=True, data_only=True)
    _, corpus = load(wb, "语料库")
    _, kuan = load(wb, "列项明细")
    _, cells = load(wb, "表格明细")
    _, formulas = load(wb, "公式明细")
    _, revs = load(wb, "修订标记")
    _, slices = load(wb, "RAG切片")
    _, qa = load(wb, "质检报告")
    _, pls = load(wb, "页级元素")
    _, struct = load(wb, "章节结构")
    _, figures = load(wb, "图表清单")

    if not corpus:
        sys.exit("「语料库」页签为空或不存在，无法生成复审队列")

    # 质检报告 G 段「存疑」→ 严重度字典
    suspect_sev = {}
    for r in qa:
        dim = sval(r, "检查维度")
        if not dim.startswith("G") and "存疑" not in dim:
            continue
        no = sval(r, "检查项")
        m = re.search(r"高|中|低", dim)
        if no and m:
            cur = suspect_sev.get(no)
            order = {"高": 3, "中": 2, "低": 1}
            if not cur or order[m.group()] > order[cur]:
                suspect_sev[no] = m.group()

    pagelevel = {}
    for r in pls:
        no = sval(r, "条文号")
        pagelevel[no] = pagelevel.get(no, 0) + 1

    page_of_parent = {sval(r, "条文号"): sval(r, "PDF起页") for r in corpus}

    objs = []
    objs += score_formulas(formulas)
    objs += score_tables(cells, page_of_parent)
    objs += score_clauses(corpus, kuan, suspect_sev, pagelevel)
    objs += score_revisions(revs)
    objs += score_figures(figures)
    objs += score_structure(struct)
    by_parent = {o["where"]: o for o in objs if o["kind"] in ("条文", "表格")}
    objs += score_slices(slices, by_parent, a.top_slices)

    KIND_ORDER = {"条文": 0, "修订标记": 1, "公式": 2, "表格": 3, "图": 4, "切片": 5, "结构": 6}
    objs.sort(key=lambda o: (-o["score"], KIND_ORDER.get(o["kind"], 9), natkey(o["where"])))
    for o in objs:
        o["prio"] = priority(o["score"])

    # 定位导出的 markdown（复审时可直接翻文件）
    md_label, md_files, md_catalog = build_md_index(a.xlsx, a.md_dir)
    md_hit = 0
    for o in objs:
        try:
            f, loc = locate_md(o["kind"], o, md_files, md_catalog, md_label)
        except Exception:                       # 定位失败不影响主流程
            f, loc = "", ""
        o["md_file"], o["md_loc"] = f, loc
        md_hit += 1 if f else 0

    n_all = len(objs)
    only = {s.strip().upper() for s in (a.only or "").split(",") if s.strip()}
    kinds = {s.strip() for s in (a.kind or "").split(",") if s.strip()}
    if only:
        objs = [o for o in objs if o["prio"] in only]
    if kinds:
        objs = [o for o in objs if o["kind"] in kinds]
    if not objs:
        sys.exit("过滤后没有对象（检查 --only / --kind 取值）")
    filtered = len(objs) != n_all

    if a.manifest:
        import json
        mf = dict(
            xlsx=os.path.basename(a.xlsx),
            total=n_all, exported=len(objs),
            filter=dict(only=sorted(only), kind=sorted(kinds)),
            md_root=md_label,
            priority={p: sum(1 for o in objs if o["prio"] == p)
                      for p in ("P0", "P1", "P2", "P3")},
            suspects=[dict(no=k, severity=v) for k, v in sorted(suspect_sev.items())],
            items=[dict(prio=o["prio"], score=o["score"], kind=o["kind"], where=o["where"],
                        parent=o["parent"], page=o["page"], why=o["why"],
                        md_file=o.get("md_file", ""), md_loc=o.get("md_loc", ""))
                   for o in objs],
        )
        with open(a.manifest, "w", encoding="utf-8") as fh:
            json.dump(mf, fh, ensure_ascii=False, indent=1)

    # ------------------------------------------------------------ 写工作簿
    wb2 = openpyxl.Workbook()
    wb2.remove(wb2.active)

    # 00 说明
    ws = wb2.create_sheet("00-复审说明")
    GUIDE = [
        ("", ""),
        ("一、这张表怎么看", ""),
        ("", "队列按**不确定度降序**排列：最可能出错的对象在最上面。请不要按条号顺序从头翻——"
             "上面 20 行就是这部规范里最不确定的 20 个地方，先判它们。"),
        ("", "每行 4 个结论（下拉）：通过 / 需修改 / 存疑保留 / 待核。拿不准就选「存疑保留」并写一句为什么。"),
        ("", "「为什么排这里」列写的是打分依据，可直接作为复审记录的依据；不同意排序可以改，但别忽略标签。"),
        ("", ""),
        ("二、每类对象去哪里比对（两种入口，任选）", ""),
        ("① 用导出的 md（最直接）", "队列的「md文件」「md内位置」两列已给出**具体文件与行号**，直接打开即可："
                                    "例如 `导出_平台导入/markdown/043_3.3.3.md` 的 L25 起是 表3.3.3 的表块。"
                                    "目录约定：`NNN_条号.md`（附表/附图在所属附录的 md 内），映射表见 `json/catalog.json`。"
                                    "没有 md 目录时这两列留空，改用 ②。"),
        ("② 用 xlsx 页签", "语料库／表格明细／公式明细／修订标记／图表清单／RAG切片／章节结构 按对象定位筛选。"),
        ("", ""),
        ("三、按类型怎么比对", ""),
        ("条文", "md 正文 L3 起（或 xlsx 语料库）。对照原件重点看：数字、否定词（应/不应）、单位、条号引用。"),
        ("表格", "md 里 `**表号 …**` 起的表块（xlsx 表格明细）。图片表格另见 imgs/ 下的渲染图；核小数位、量级、`—`≠`0`。"),
        ("公式", "md 里 `〔公式（x）〕` 所在行（表达式本身在 xlsx 公式明细）；对照 formulas/ 渲染图核系数/指数/运算符。"),
        ("修订标记", "md 里对应条文行（xlsx 修订标记）；对照原件下划线，判断是否确为修订内容，并留意漏标。"),
        ("图", "md 里出现图号的行（xlsx 图表清单）；核题注完整性，确认图内信息是否必须数值化。"),
        ("切片", "回到父条的 md（xlsx RAG切片）；验元数据头 L1、父块含全款、引用闭合。"),
        ("结构", "xlsx 章节结构（md 是条文级导出，不含章级记录）；对照原件目次核层级、顺序、条号连续性。"),
        ("", ""),
        ("四、不确定度怎么算（0-100，可复核）", ""),
        ("来源可靠性", "图片表格视觉转录 +30；公式由渲染图还原 +30；几何识别下划线 +25；扫描页截断/污损 +20；文字层抽取 0"),
        ("内部矛盾", "原表自相矛盾（值 vs 注／量级） +25；与他人交叉校验出分歧 +20"),
        ("已登记存疑", "严重度高 +45 / 中 +30 / 低 +15"),
        ("结构性脆弱点", "A/B 后缀编号 +8；跨页 +6；结尾非句读 +10；已删除条 +12；公式占位 +10；页级元素混排 +5；附录整册 +8"),
        ("复核覆盖缺口", "方法原理覆盖不到（如修订漏标）+25；无独立交叉校验 +15"),
        ("优先级", "P0 ≥82 必须先看；P1 60-81；P2 35-59；P3 <35 可批量扫过"),
        ("", ""),
        ("五、四个结论的含义", ""),
        ("通过", "已比对，与权威原件一致（半角/全角、空格差异不算不一致）。"),
        ("需修改", "与原件不一致，或确系排印错误需加注；**必须**在「复核意见」写清改成什么、依据哪一处。"),
        ("存疑保留", "原件可疑但证据不足以断定。本库原则是**不改原文**，只登记。"),
        ("待核", "尚未判读（初始值）。"),
        ("", ""),
        ("六、改哪里（不要直接改语料库 xlsx）", ""),
        ("", "xlsx 是脚本产物，手改会被下次重跑覆盖。三种回灌方式："),
        ("A. 改脚本", "条文修正 → 建库脚本；表格数值 → 表格转录表；公式 → 公式表。改完重跑建库 + 三步校验，"
                      "必须确认「漏条 0 / 编造 0 / 保真 100% / 切片对账一致」不退化。"),
        ("B. 写覆盖表", "复审人不动代码时：回传「对象定位 + 原文 + 应改为 + 依据」给维护人统一回灌。"),
        ("C. 只加注", "把疑点写进「质检报告」，检索侧用同义词兜（如 A||B||C），最适合 RAG。"),
        ("", ""),
        ("七、验收线", ""),
        ("", "「02-汇总」里待核数归零；每条「需修改」都写明改法与依据；重跑三步校验四项指标不退化。"),
        ("", "自动校验只能证明「结构没漏、字没串行、切片对得上」，证明不了「与原件逐字一致」——"
             "所以这一关不能省，也不能用「自动校验全过」代替。"),
    ]
    for i, (k, v) in enumerate(GUIDE, start=1):
        ws.cell(row=i, column=1, value=k).font = Font(
            bold=bool(k), size=10, color="1F4E79" if k.startswith(("一", "二", "三", "四", "五")) else "000000")
        c = ws.cell(row=i, column=2, value=v)
        c.font, c.alignment = TEXT, WRAP
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 120

    # 01 队列
    ws = wb2.create_sheet("01-复审队列")
    head = ["序", "优先级", "不确定度", "对象类型", "对象定位", "所属条文/父条",
            "PDF页", "待判读内容", "为什么排这里（打分依据）",
            "md文件（导出包内路径）", "md内位置",
            "结论", "复核人", "日期", "复核意见（要改就写清怎么改）"]
    ws.append(head)
    for j in range(1, len(head) + 1):
        c = ws.cell(row=1, column=j)
        c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, CENTER, BORDER
    for i, o in enumerate(objs, start=2):
        ws.append([i - 1, o["prio"], o["score"], o["kind"], o["where"], o["parent"],
                   o["page"], o["content"], o["why"],
                   o.get("md_file", ""), o.get("md_loc", ""),
                   "待核", "", "", ""])
        for j in range(1, len(head) + 1):
            c = ws.cell(row=i, column=j)
            c.font, c.border = TEXT, BORDER
            c.alignment = CENTER if j <= 4 or j in (7, 11, 12, 13, 14) else WRAP
        ws.cell(row=i, column=2).fill = PRIORITY_FILL[o["prio"]]
        ws.cell(row=i, column=12).fill = PatternFill("solid", fgColor="FFF2CC")
    for col, w in zip("ABCDEFGHIJKLMNO",
                      [5, 8, 9, 9, 18, 14, 7, 62, 40, 42, 22, 10, 8, 10, 30]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 34
    dv = DataValidation(type="list", formula1='"' + ",".join(VERDICTS) + '"',
                        allow_blank=True, showDropDown=False)
    dv.error = "请从下拉中选择"
    ws.add_data_validation(dv)
    dv.add(f"L2:L{len(objs) + 1}")

    # 02 汇总
    ws = wb2.create_sheet("02-汇总")
    ws["A1"] = "复审进度（打开后自动重算）"
    ws["A1"].font = Font(bold=True, size=12, color="1F4E79")
    r = 3
    ws.cell(row=r, column=1, value="按优先级").font = Font(bold=True, size=10)
    r += 1
    for j, h in enumerate(["优先级", "对象数", "待核", "通过", "需修改", "存疑保留"], start=1):
        c = ws.cell(row=r, column=j, value=h)
        c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, CENTER, BORDER
    r += 1
    for p in ("P0", "P1", "P2", "P3"):
        ws.cell(row=r, column=1, value=p)
        ws.cell(row=r, column=2, value=f'=COUNTIF(\'01-复审队列\'!B:B,"{p}")')
        for j, v in zip(range(3, 7), VERDICTS):
            ws.cell(row=r, column=j,
                    value=f'=COUNTIFS(\'01-复审队列\'!B:B,"{p}",\'01-复审队列\'!L:L,"{v}")')
        r += 1
    ws.cell(row=r, column=1, value="合计").font = Font(bold=True, size=10)
    for j in range(2, 7):
        L = get_column_letter(j)
        ws.cell(row=r, column=j, value=f"=SUM({L}{r - 4}:{L}{r - 1})").font = Font(bold=True, size=10)
    r += 2
    ws.cell(row=r, column=1, value="按对象类型").font = Font(bold=True, size=10)
    r += 1
    for j, h in enumerate(["类型", "对象数", "待核", "通过", "需修改", "存疑保留"], start=1):
        c = ws.cell(row=r, column=j, value=h)
        c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, CENTER, BORDER
    r += 1
    start = r
    for kind in sorted({o["kind"] for o in objs}, key=lambda k: KIND_ORDER.get(k, 9)):
        ws.cell(row=r, column=1, value=kind)
        ws.cell(row=r, column=2, value=f'=COUNTIF(\'01-复审队列\'!D:D,"{kind}")')
        for j, v in zip(range(3, 7), VERDICTS):
            ws.cell(row=r, column=j,
                    value=f'=COUNTIFS(\'01-复审队列\'!D:D,"{kind}",\'01-复审队列\'!L:L,"{v}")')
        r += 1
    ws.cell(row=r, column=1, value="合计").font = Font(bold=True, size=10)
    for j in range(2, 7):
        L = get_column_letter(j)
        ws.cell(row=r, column=j, value=f"=SUM({L}{start}:{L}{r - 1})").font = Font(bold=True, size=10)
    for col, w in zip("ABCDEF", [14, 10, 10, 10, 10, 12]):
        ws.column_dimensions[col].width = w
    for rr in ws.iter_rows(min_row=3, max_row=r):
        for c in rr:
            c.border = BORDER
            if c.font.size != 10:
                c.font = TEXT
    r += 2
    c = ws.cell(row=r, column=1, value="验收线：待核 = 0，且每条「需修改」都在复核意见里写明了改法与依据。")
    c.font = Font(size=10, italic=True, color="C00000")
    if filtered:
        r += 1
        cond = []
        if only:
            cond.append("优先级 " + "/".join(sorted(only)))
        if kinds:
            cond.append("类型 " + "/".join(sorted(kinds)))
        c = ws.cell(row=r, column=1,
                    value=f"⚠ 本表是过滤导出（{'；'.join(cond)}），仅 {len(objs)}/{n_all} 个对象；"
                          f"整体进度看全量工作台。")
        c.font = Font(size=10, italic=True, color="ED7D31")

    wb2.calculation.fullCalcOnLoad = True
    wb2.save(out)

    if not a.quiet:
        cnt = {}
        for o in objs:
            cnt[o["prio"]] = cnt.get(o["prio"], 0) + 1
        print(f"已生成：{out}" + (f"（过滤后 {len(objs)}/{n_all}）" if filtered else ""))
        print(f"复审对象 {len(objs)} 个｜" + "　".join(f"{p} {cnt.get(p, 0)}" for p in ("P0", "P1", "P2", "P3")))
        print("队列前 5 项（最不确定）：")
        for o in objs[:5]:
            print(f"  {o['prio']} {o['score']:>3} {o['kind']} {o['where']}｜{o['why'][:56]}")
        if a.manifest:
            print(f"机器可读清单：{a.manifest}")


if __name__ == "__main__":
    main()
