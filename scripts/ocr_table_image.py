# -*- coding: utf-8 -*-
"""阶段3（v4）：表格 OCR —— 固定网格，不搞 colspan 跳跃。

前三版教训：
    v1 只按 y 聚类 → 单元格内换行被拆成两行
    v2 所有框 x0 聚类 → 表头合并单元格搅乱列边界，12 列里 6 列空
    v3 用数据框定列（对！）但引入 colspan 跳跃 → 表头跨列把后面的列吃掉，
       输出只剩 3 列

v4 定案：**网格固定，格子只填不跳**
    1. 列：用"数据框"（纯数字/倍杆）的 x 中心聚类定列，实测极规整
       （6 列：248/328/394/426/454/514）。
    2. 最左再加一列"行标签列"，边界取第一数据列中心 - 40（实测标签 x1≤198，中心 248）。
    3. 行：用数据框 y 中心聚类定行。
    4. 每个框按 **x 中心** 就近落到一列、按 y 中心就近落到一行，**不跳格**。
       合并单元格的代价是文字落到起始列，但**列结构不会被破坏**。
"""
import json
import re
import sys
from pathlib import Path

import fitz
from rapidocr_onnxruntime import RapidOCR

if len(sys.argv) < 2:
    sys.exit("用法：python ocr_table_image.py <规范.pdf>\n"
             "把 PDF 里的大图（表格）OCR 后按坐标重建为行列网格，\n"
             "输出到 <pdf目录>/_pipeline/tables_ocr.json")

PDF = sys.argv[1]
CASE = Path(PDF).parent / "_pipeline"
IMG_DIR = CASE / "table_imgs"
IMG_DIR.mkdir(exist_ok=True)
ocr = RapidOCR()

ROMAN_FIX = [
    (re.compile(r"(?<![A-Za-z])I(?=[、，,．.])"), "Ⅰ"),
    (re.compile(r"(?<![A-Za-z])II(?=[、，,．.])"), "Ⅱ"),
    (re.compile(r"(?<![A-Za-z])III(?=[、，,．.])"), "Ⅲ"),
]
OCR_FIX = [
    (re.compile(r"^注！$"), "注1"),
    (re.compile(r"^[(（]000>$"), "≤5000m³"),
    (re.compile(r"<501$"), "≤50t"),
]


def fix_text(s):
    for pat, rep in ROMAN_FIX + OCR_FIX:
        s = pat.sub(rep, s)
    return re.sub(r"Ⅰ、级", "Ⅰ、Ⅱ级", s)


def cluster_anchor(vals, gap):
    order = sorted(range(len(vals)), key=lambda i: vals[i])
    lab = [0] * len(vals)
    cur, anchor = 0, vals[order[0]]
    for k in range(1, len(order)):
        v = vals[order[k]]
        if v - anchor > gap:
            cur += 1
            anchor = v
        lab[order[k]] = cur
    mins = {}
    for i, c in enumerate(lab):
        mins[c] = min(mins.get(c, 1e9), vals[i])
    remap = {c: i for i, c in enumerate(sorted(mins, key=lambda c: mins[c]))}
    return [remap[c] for c in lab], len(mins)


def is_data_cell(t):
    s = t.strip()
    return bool(re.fullmatch(r"\d+(?:\.\d+)?", s)) or "倍" in s


def rebuild(boxes):
    if not boxes:
        return [], dict(n_col=0, n_row=0, cols=[], rows=[])

    data = [b for b in boxes if is_data_cell(b[4])]
    anchors = data if len(data) >= 4 else boxes

    # ---- 列中心（数据框）
    dx = [(b[0] + b[2]) / 2 for b in anchors]
    lab_c, n_data_col = cluster_anchor(dx, 20)
    groups = {}
    for c, x in zip(lab_c, dx):
        groups.setdefault(c, []).append(x)
    centers = [sum(groups[c]) / len(groups[c]) for c in sorted(groups)]

    # 加行标签列：中心取「第一数据列中心 - 80」（实测标签居中在 ~60~100）
    all_centers = [centers[0] - 90] + centers

    # ---- 行中心（数据框）
    dy = [(b[1] + b[3]) / 2 for b in anchors]
    lab_r, n_row = cluster_anchor(dy, 21)
    rgroups = {}
    for c, y in zip(lab_r, dy):
        rgroups.setdefault(c, []).append(y)
    rcenters = [sum(rgroups[c]) / len(rgroups[c]) for c in sorted(rgroups)]

    # ---- 表头区：第一个数据行中心之上的文字框，是表头（双层合并单元格）。
    #      不把它并进第一条数据——实测会得到 "…1.厂外民用建筑" 这种混合格。
    #      做法：表头按 y 再聚成若干行，放在数据行之前。
    first_data_y = min(rcenters)
    head_boxes = [b for b in boxes if (b[1] + b[3]) / 2 < first_data_y - 8]
    head_rows = []
    if head_boxes:
        hy = [(b[1] + b[3]) / 2 for b in head_boxes]
        hl, nh = cluster_anchor(hy, 16)
        hgroups = {}
        for c, y in zip(hl, hy):
            hgroups.setdefault(c, []).append(y)
        hcenters = [sum(hgroups[c]) / len(hgroups[c]) for c in sorted(hgroups)]

        def hci(xc):
            return min(range(len(all_centers)),
                       key=lambda i: abs(all_centers[i] - xc))

        hcells = {}
        for x0, y0, x1, y1, t in head_boxes:
            xc, yc = (x0 + x1) / 2, (y0 + y1) / 2
            ri = min(range(len(hcenters)), key=lambda i: abs(hcenters[i] - yc))
            hcells.setdefault((ri, hci(xc)), []).append((y0, fix_text(t)))
        for r in range(len(hcenters)):
            head_rows.append([
                "".join(p[1] for p in sorted(hcells.get((r, c), []), key=lambda p: p[0]))
                for c in range(len(all_centers))])
        # 表头行不参与数据行匹配
        boxes = [b for b in boxes if (b[1] + b[3]) / 2 >= first_data_y - 8]

    # ---- 就近落格
    cells = {}
    for x0, y0, x1, y1, t in boxes:
        xc, yc = (x0 + x1) / 2, (y0 + y1) / 2
        ci = min(range(len(all_centers)), key=lambda i: abs(all_centers[i] - xc))
        ri = min(range(len(rcenters)), key=lambda i: abs(rcenters[i] - yc))
        cells.setdefault((ri, ci), []).append((y0, fix_text(t)))

    grid = list(head_rows)
    for r in range(n_row):
        row = []
        for c in range(len(all_centers)):
            parts = sorted(cells.get((r, c), []), key=lambda p: p[0])
            row.append("".join(p[1] for p in parts))
        grid.append(row)

    return grid, dict(n_col=len(all_centers), n_row=len(grid),
                      cols=[round(x) for x in all_centers],
                      rows=[round(y) for y in rcenters])


def main():
    doc = fitz.open(PDF)
    tables = []
    for idx in range(doc.page_count):
        for j, im in enumerate(doc[idx].get_images()):
            xref, w, h = im[0], im[2], im[3]
            if w < 300 or h < 200:
                continue
            pix = fitz.Pixmap(doc, xref)
            if pix.n > 4:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            p = IMG_DIR / f"p{idx + 1}_{j}.png"
            pix.save(p)
            result, _ = ocr(str(p))
            if not result:
                continue
            boxes = []
            for box, text, score in result:
                xs = [q[0] for q in box]
                ys = [q[1] for q in box]
                boxes.append((min(xs), min(ys), max(xs), max(ys), text.strip()))
            grid, meta = rebuild(boxes)
            tables.append(dict(page=idx + 1, img=p.name, w=w, h=h, grid=grid, meta=meta))
            print(f"  p{idx + 1} 图{j}: {w}x{h} 框{len(boxes)} → "
                  f"{meta['n_row']}行×{meta['n_col']}列")
    doc.close()
    out = CASE / "tables_ocr.json"
    out.write_text(json.dumps(tables, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n共 {len(tables)} 张 → {out}")


if __name__ == "__main__":
    main()
