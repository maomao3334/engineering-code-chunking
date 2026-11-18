# -*- coding: utf-8 -*-
"""源 PDF 缺陷自诊断：在切分之前先查出「这份 PDF 本身就坏了」的地方

为什么需要（实测踩过）：
    GB 50222-1995 的 表3.4.1 切完没有数值，切片里写「表体与表3.2.1 完全相同」。
    肉眼看 PDF 第 10 页，表3.4.1 清清楚楚有数据 —— 一度以为是切分/OCR 的锅。
    用 PyMuPDF 一查：p7 和 p10 **引用的是同一个图片对象 xref=84**（md5 相同）。
    也就是说这份 PDF 里 表3.4.1 的图，本来就是 表3.2.1 的图 ——
    **源文件缺陷**，换任何切分方法都救不回来，只能换源或人工补录。

    以前这类问题要等到「导出后发现表是空的」才被发现，返工成本极高。
    这个脚本把它提前到**切分之前**：一条命令列出全部可疑点。

能查出的四类（P0/P1/P2）：
    P0 表图跨页复用    同一 image xref / 同一 md5 出现在 ≥2 页或多处
                       → 至少一处表内容错位（本次 表3.4.1 就是这个）
    P1 表号存在但无表  正文写了「按表X.Y.Z」，该页却既没大图也没表格文本
                       → 表可能是矢量绘制、或整页扫描、或表被裁掉
    P1 整页扫描        单张图覆盖 >60% 页面积 → 无文字层，必须 OCR/VLM
    P2 无文字层        全文字符数≈0 → 公式/符号只能靠视觉转录

用法：
    python diagnose_source_defects.py 规范.pdf
    python diagnose_source_defects.py 规范.pdf --md 源缺陷清单.md --json 源缺陷.json
    python diagnose_source_defects.py 规范.pdf --min-width 300 --quiet
"""
import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:
        pass

try:
    import fitz
except ImportError:
    sys.exit("需要 PyMuPDF：pip install pymupdf")

TBL_NO = re.compile(r"表\s*([A-Z]?\d+(?:\s*[\.\-]\s*\d+)*)")


def _table_anchor(page, img_rect, cache={}):
    """在图上方找最近的一个「表X.Y.Z」，作为这张图所属的表号"""
    key = id(page)
    if key not in cache:
        items = []
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, txt = b[0], b[1], b[2], b[3], b[4]
            for m in TBL_NO.finditer(txt or ""):
                items.append((y0, m.group(0).replace(" ", "")))
        cache[key] = sorted(items)
    best = None
    for y0, name in cache[key]:
        if y0 <= img_rect.y0 + 8:
            d = img_rect.y0 - y0
            if 0 <= d <= 260 and (best is None or d < best[0]):
                best = (d, name)
    return best[1] if best else ""


def diagnose(pdf_path, min_width=400):
    doc = fitz.open(pdf_path)
    pages = []
    for i in range(doc.page_count):
        pg = doc[i]
        txt = pg.get_text()
        big = []
        for im in pg.get_images(full=True):
            xref, w, h = im[0], im[2], im[3]
            if w < min_width:
                continue
            try:
                data = doc.extract_image(xref)["image"]
                md5 = hashlib.md5(data).hexdigest()
            except Exception:
                md5 = ""
            rects = pg.get_image_rects(xref)
            big.append(dict(xref=xref, w=w, h=h, md5=md5,
                            rects=[(r.x0, r.y0, r.x1, r.y1) for r in rects],
                            table=_table_anchor(pg, rects[0]) if rects else ""))
        pages.append(dict(page=i + 1, chars=len(txt.strip()), text=txt, big=big))

    defects = []

    # ---- P0 表图跨页复用
    by_xref, by_md5 = {}, {}
    for p in pages:
        for b in p["big"]:
            by_xref.setdefault(b["xref"], []).append((p["page"], b))
            if b["md5"]:
                by_md5.setdefault(b["md5"], []).append((p["page"], b))
    seen = set()
    for md5, uses in by_md5.items():
        if len(uses) < 2:
            continue
        key = tuple(sorted(u[0] for u in uses))
        if key in seen:
            continue
        seen.add(key)
        names = [u[1]["table"] for u in uses if u[1]["table"]]
        if len(set(names)) < 2:
            # 同一个表跨页（长表续页）不算缺陷
            continue
        defects.append(dict(
            level="P0", kind="表图跨页复用",
            detail=f"同一图片（xref={uses[0][1]['xref']}，md5={md5[:12]}，"
                   f"{uses[0][1]['w']}x{uses[0][1]['h']}）被挂到不同表号下",
            where="、".join(f"p{u[0]}→{u[1]['table'] or '未识别表号'}" for u in uses),
            advice="至少一处表内容错位。换一份干净 PDF；或人工补录该表并标注来源；"
                   "或保留「源缺陷·无数据」标注（最安全）"))

    # ---- P1 表号有、表没有
    for p in pages:
        names = sorted(set(m.group(0).replace(" ", "") for m in TBL_NO.finditer(p["text"])))
        if not names or p["big"]:
            continue
        if "|" in p["text"] and p["text"].count("|") > 12:
            continue          # 文字层里已经是表格
        defects.append(dict(
            level="P1", kind="表号存在但无表",
            detail=f"正文提到 {'、'.join(names[:4])}，但该页既无大图也无表格文本",
            where=f"p{p['page']}（{p['chars']} 字）",
            advice="可能是矢量绘制表格/表格被裁掉/整页扫描。"
                   "先放大该页目视确认，再决定走文字抽取还是视觉转录"))

    # ---- P1 整页扫描（25 页的扫描件不要报 25 条，合并成 1 条）
    scan_pages = []
    for p in pages:
        for b in p["big"]:
            if not b["rects"]:
                continue
            x0, y0, x1, y1 = b["rects"][0]
            area = max(0.0, (x1 - x0)) * max(0.0, (y1 - y0))
            if area <= 0:
                continue
            parea = 842 * 595
            if area / parea > 0.6 and p["chars"] < 60:
                scan_pages.append(p["page"])
                break
    if scan_pages:
        if len(scan_pages) >= 3:
            show = "、".join(f"p{x}" for x in scan_pages[:6])
            more = f" 等 {len(scan_pages)} 页" if len(scan_pages) > 6 else ""
            defects.append(dict(
                level="P1", kind="整页扫描",
                detail=f"{len(scan_pages)} 页为整页图片（文字层 <60 字/页）",
                where=show + more,
                advice="整本必须 OCR/视觉转录；公式与符号按「双写法」落库"
                       "（$…$ 渲染式 + 纯文本对照），否则平台读不懂"))
        else:
            for pno in scan_pages:
                defects.append(dict(
                    level="P1", kind="整页扫描",
                    detail="单张图覆盖 >60% 页面积，文字层几乎为空",
                    where=f"p{pno}",
                    advice="该页必须 OCR/视觉转录；公式按「双写法」落库"))

    # ---- P2 无文字层
    total = sum(p["chars"] for p in pages)
    if total < 40 * doc.page_count * 0.2:
        defects.append(dict(
            level="P2", kind="基本无文字层",
            detail=f"全文仅 {total} 字（{doc.page_count} 页），文字层不可用",
            where="全文",
            advice="整本走 OCR/视觉转录；公式务必同时输出 $…$ 与纯文本两版"))

    doc.close()
    return pages, defects


def render_md(pdf_path, pages, defects):
    n0 = sum(1 for d in defects if d["level"] == "P0")
    n1 = sum(1 for d in defects if d["level"] == "P1")
    n2 = sum(1 for d in defects if d["level"] == "P2")
    L = [f"# 源 PDF 缺陷清单", "",
         f"- 源文件：`{Path(pdf_path).name}`",
         f"- 页数：{len(pages)}｜页均文字：{sum(p['chars'] for p in pages) // max(1, len(pages))} 字",
         f"- 命中：**P0 {n0} 处 / P1 {n1} 处 / P2 {n2} 处**", ""]
    if not defects:
        L += ["✅ 未发现源缺陷。可以进入切分流程。", ""]
        return "\n".join(L)
    order = {"P0": 0, "P1": 1, "P2": 2}
    for d in sorted(defects, key=lambda x: order[x["level"]]):
        L += [f"## {d['level']}　{d['kind']}", "",
              f"- 位置：{d['where']}",
              f"- 现象：{d['detail']}",
              f"- 处理：{d['advice']}", ""]
    if n0:
        L += ["---", "",
              "## P0 必须当场决策（三选一，写进交付说明）", "",
              "1. **换源**：找同一规范的干净 PDF（推荐）",
              "2. **人工补录**：从纸质版/官方条文说明逐格核对，补进 `overrides/`，"
              "切片里标注「人工补录，源见…」",
              "3. **标注无数据**：切片写明「源 PDF 该表为错误图片，无法取得数值，"
              "请查纸质版」，绝不套用相邻表的数据", ""]
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="源 PDF 缺陷自诊断（表图复用/无表/扫描页）")
    ap.add_argument("pdf")
    ap.add_argument("--md", default="", help="输出 Markdown 报告")
    ap.add_argument("--json", dest="json_out", default="", help="输出 JSON")
    ap.add_argument("--min-width", type=int, default=400,
                    help="多大像素宽才当「表格图」（默认 400）")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()

    if not Path(a.pdf).is_file():
        sys.exit(f"找不到 {a.pdf}")

    pages, defects = diagnose(a.pdf, a.min_width)
    md = render_md(a.pdf, pages, defects)

    if a.md:
        Path(a.md).write_text(md, encoding="utf-8")
    if a.json_out:
        Path(a.json_out).write_text(
            json.dumps(dict(pdf=a.pdf, pages=len(pages), defects=defects),
                       ensure_ascii=False, indent=1), encoding="utf-8")

    if not a.quiet:
        print(md)
        if a.md:
            print(f"报告：{a.md}")
        if a.json_out:
            print(f"JSON：{a.json_out}")
    p0 = sum(1 for d in defects if d["level"] == "P0")
    return 1 if p0 else 0


if __name__ == "__main__":
    sys.exit(main())
