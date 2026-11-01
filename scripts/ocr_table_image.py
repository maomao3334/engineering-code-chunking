# -*- coding: utf-8 -*-
"""阶段3（v8，定案）：结构取 pred_html（已验证正确），文字按单元格边界精确回填。

教训：不要自己解释 logic_points 的坐标约定——踩了两次（顺序反了导致行列膨胀）。
      rapid_table 的 **pred_html 是可靠的**（已实测：rowspan/colspan 正确、表头层级正确）。
      所以：
        结构  ← 解析 pred_html 的 <tr>/<td rowspan colspan>
        文字  ← 按 cell_bboxes 多边形回填 OCR 框（避免默认输出的跨框粘连）

两层配合：HTML 给网格形状，几何回填给每格的干净文字。
"""
import json
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

import fitz
import numpy as np
from rapid_table import RapidTable, RapidTableInput

if len(sys.argv) < 2:
    sys.exit("用法：python ocr_table_image.py <规范.pdf>\n"
             "结构用 rapid_table(SLANet-Plus) 的 pred_html，文字按单元格几何回填。\n"
             "依赖：pip install rapid_table rapidocr")

PDF = sys.argv[1]
CASE = Path(PDF).parent / "_pipeline"
IMG_DIR = CASE / "table_imgs"
IMG_DIR.mkdir(exist_ok=True)

TEXT_FIX = [
    (re.compile(r"^序台$"), "序号"),
    (re.compile(r"^一内、"), "内、"),
    # 上标（D02）：25000m2 → 25000m³；5000m3 → 5000m³
    (re.compile(r"(\d)m2(?![0-9])"), r"\1m²"),
    (re.compile(r"(\d)m3(?![0-9])"), r"\1m³"),
    # 罗马数字（D04）：I、级 → Ⅰ、Ⅱ级；单独 I、 → Ⅰ、
    (re.compile(r"(?<![A-Za-z])I、级"), "Ⅰ、Ⅱ级"),
    (re.compile(r"(?<![A-Za-z])I、"), "Ⅰ、"),
    (re.compile(r"(?<![A-Za-z])II(?=[、，,．.])"), "Ⅱ"),
]


def fix_text(s):
    for pat, rep in TEXT_FIX:
        s = pat.sub(rep, s)
    # 括号配对：合并表头常出现 "（），" 这种
    if s.count("（") != s.count("）"):
        if s.count("（") < s.count("）"):
            s = "（" * (s.count("）") - s.count("（")) + s
        else:
            s = s + "）" * (s.count("（") - s.count("）"))
    return s


class HTMLTable(HTMLParser):
    """解析 pred_html → 单元格列表 [(sr, sc, rs, cs)]"""

    def __init__(self):
        super().__init__()
        self.rows, self._row, self._cell = [], None, None

    def handle_starttag(self, tag, a):
        a = dict(a)
        if tag == "tr":
            self._row = []
        elif tag == "td" and self._row is not None:
            self._cell = (int(a.get("rowspan", 1) or 1),
                          int(a.get("colspan", 1) or 1))

    def handle_endtag(self, tag):
        if tag == "td" and self._cell is not None and self._row is not None:
            self._row.append(self._cell)
            self._cell = None
        elif tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def layout(self):
        """展开成 (n_row, n_col, [(sr, sc, rs, cs)])"""
        if not self.rows:
            return 0, 0, []
        n_col = max(sum(cs for _, cs in r) for r in self.rows)
        n_row = len(self.rows)
        occ = [[False] * n_col for _ in range(n_row)]
        out = []
        for r, row in enumerate(self.rows):
            c = 0
            for rs, cs in row:
                while c < n_col and occ[r][c]:
                    c += 1
                if c >= n_col:
                    break
                for dr in range(rs):
                    for dc in range(cs):
                        if r + dr < n_row and c + dc < n_col:
                            occ[r + dr][c + dc] = True
                out.append((r, c, rs, cs))
                c += cs
        return n_row, n_col, out


def poly_bounds(p):
    a = np.array(p, dtype=float).reshape(-1, 2)
    return a[:, 0].min(), a[:, 1].min(), a[:, 0].max(), a[:, 1].max()


def build_table(img, engine):
    res = engine(str(img))
    html = res.pred_htmls[0] if len(res.pred_htmls) else ""
    parser = HTMLTable()
    parser.feed(html)
    n_row, n_col, cells = parser.layout()
    if not cells:
        return [], [], 0

    polys = res.cell_bboxes[0] if len(res.cell_bboxes) else []
    polys = [poly_bounds(p) for p in polys]

    # OCR 框
    ocr_out = engine.ocr_engine(str(img))
    boxes = []
    if getattr(ocr_out, "boxes", None) is not None:
        for box, txt in zip(ocr_out.boxes, ocr_out.txts):
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            boxes.append(dict(cx=(min(xs) + max(xs)) / 2,
                              cy=(min(ys) + max(ys)) / 2,
                              y0=min(ys), x0=min(xs),
                              t=str(txt).strip()))

    # HTML 单元格数应与多边形数一致；不一致则退回"多边形顺序"对应
    grid = [["" for _ in range(n_col)] for _ in range(n_row)]
    if len(polys) == len(cells):
        areas = [(x1 - x0) * (y1 - y0) for (x0, y0, x1, y1) in polys]

        def owner_of(b):
            """框中心落在哪些单元格里，返回面积最大的那个（父格优先）"""
            hit = [i for i, (ax0, ay0, ax1, ay1) in enumerate(polys)
                   if ax0 - 2 <= b["cx"] <= ax1 + 2
                   and ay0 - 2 <= b["cy"] <= ay1 + 2]
            return max(hit, key=lambda i: areas[i]) if hit else None

        for i, ((sr, sc, rs, cs), (x0, y0, x1, y1)) in enumerate(zip(cells, polys)):
            parts = []
            for b in boxes:
                if owner_of(b) == i:
                    parts.append(b)
            parts.sort(key=lambda b: (round(b["y0"]), b["x0"]))
            grid[sr][sc] = fix_text("".join(p["t"] for p in parts))
    else:
        # 兜底：HTML 文字（会粘连，但至少有内容）
        plain = re.findall(r"<td[^>]*>(.*?)</td>", html, re.S)
        idx = 0
        for sr, sc, rs, cs in cells:
            if idx < len(plain):
                grid[sr][sc] = fix_text(re.sub(r"<[^>]+>|\s+", "", plain[idx]))
                idx += 1

    spans = [[sr, sc, rs, cs, grid[sr][sc]] for sr, sc, rs, cs in cells
             if rs > 1 or cs > 1]

    def is_num(s):
        return bool(re.fullmatch(r"\d+(?:\.\d+)?", s.strip()))

    # 数据行判定：行里有"行标签"（1. / 2. / 序号）**或**≥2 个纯数字。
    # 只看数字会把「注1 注1 25 — 17」这种数据行误判成表头（实测表4.1.7）。
    ROWLABEL = re.compile(r"^\s*\d+[.、]\s*\S")

    hrows = n_row
    for i, row in enumerate(grid):
        nnum = sum(1 for x in row if is_num(x))
        has_label = any(ROWLABEL.match(str(x)) for x in row)
        if nnum >= 2 or has_label:
            hrows = i
            break
    return grid, spans, hrows


def main():
    engine = RapidTable(RapidTableInput())
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
            grid, spans, hrows = build_table(p, engine)
            tables.append(dict(page=idx + 1, img=p.name, w=w, h=h, grid=grid,
                               spans=spans, header_rows=hrows,
                               n_row=len(grid), n_col=len(grid[0]) if grid else 0))
            print(f"  p{idx + 1} 图{j}: {w}x{h} → {len(grid)}行×"
                  f"{len(grid[0]) if grid else 0}列, 表头{hrows}行, 合并{len(spans)}")
    doc.close()
    out = CASE / "tables_ocr.json"
    out.write_text(json.dumps(tables, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n共 {len(tables)} 张 → {out}")


if __name__ == "__main__":
    main()
