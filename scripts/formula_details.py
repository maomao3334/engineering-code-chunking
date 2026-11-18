# -*- coding: utf-8 -*-
"""
公式明细 ↔ 正文 ↔ 切片 的落地与校验（v1.10.2 新增）
=====================================================

解决的现实问题
--------------
GB 50193-93 这类"扫描件重建"的规范，PDF 文字层把公式拆成单字符、运算符
（= + × √）还用 Symbol 字体私有区码位存放，**表达式无法从文字顺序还原**。
v1.10.0/1.10.1 的做法是：正文留 `〔公式（x）〕` 占位、表达式只进 xlsx「公式明细」。
代价是平台上检索到了条文也算不出一一公式不在切片里。

v1.10.2 起改为"分级写回正文"：

  · 高可信公式  → 写入正文与切片（md 用 MathJax + 纯文本双写法）
                 但**必须带标记**：〔公式（x）·整理式·非原文逐字·可信度 高〕
  · 中/低可信公式 → 正文只留〔公式（x）·待核·未引入表达式〕+ 一句话原因，
                 表达式**不进正文**，避免把 AI 还原式当条文原文引用
  · 原文列（语料库.条文原文本体）**始终保留裸占位**，保持原文保真与交叉校验可比

三条护栏（本脚本 --check 实现）
------------------------------
  G1 分级   ：高可信必须落到切片/md；中低可信必须只有占位、不得出现表达式
  G2 双写法 ：md 里每个 $$ 块后面必须紧跟 `> 纯文本：`，且 LaTeX 语法配平
  G3 对账   ：正文裸占位 ↔ 公式明细 双向一致；切片/md 里的标记块也都能对上明细
  G4 完整性 ：公式所在的 md / 父块切片里，必须同时含该条的「式中」变量定义
              （否则检索到公式也代入不了）

公式明细表（xlsx 页签「公式明细」）推荐列
---------------------------------------
  公式编号 / 所属条文 / PDF页 / 规范化表达式（整理，非原文逐字）
  / LaTeX / 纯文本 / 正文处理 / 可信度 / 存疑说明 / 渲染图

  正文处理取值：写入正文 | 待核占位
  （缺该列时按"有 LaTeX 且可信度以「高」开头 → 写入"推断，向后兼容旧库）

用法
----
  # 1) 纯校验（四关之外的公式专项，可单独跑，也可被 validate_chunks 调用）
  python formula_details.py --xlsx 语料库.xlsx [--md 导出/markdown] --check

  # 2) 把公式注入 md（导出脚本未内置注入时用；幂等）
  python formula_details.py --xlsx 语料库.xlsx --md 导出/markdown --in-place --style both
  python formula_details.py --xlsx 语料库.xlsx --md 导出/markdown --in-place --style plain

  # 3) 只打印某条条文注入后的效果（调试）
  python formula_details.py --xlsx 语料库.xlsx --preview 3.2.3
"""
from __future__ import annotations

import argparse
import os
import re
import sys

import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import normalize_formulas as nf  # 同目录；复用 latex_to_plain 与语法检查

sys.stdout.reconfigure(encoding="utf-8")

# ---------------------------------------------------------------- 常量

SHEET = "公式明细"
CORPUS_SHEET = "语料库"
SLICE_SHEET = "RAG切片"
BODY_COL = "条文原文本体"
PARENT_COL = "父条号"
SLICE_TEXT_COL = "切片全文"

# 裸占位：〔公式（3.2.3-1）〕
RE_PLACEHOLDER = re.compile(r"〔公式\s*[（(]\s*([^）)〕]+?)\s*[）)]\s*〕")
# 已注入标记：〔公式（3.2.3-1）·整理式…〕 / 〔公式（4.0.5）·待核…〕
RE_MARKED = re.compile(r"〔公式\s*[（(]\s*([^）)〕]+?)\s*[）)][^〕]*〕")

MARK_WRITE = "〔公式（{fid}）·整理式·非原文逐字·可信度{level}〕"
MARK_PENDING = "〔公式（{fid}）·待核·未引入表达式〕"
NOTE_PENDING = "（该公式存疑：{reason}；表达式待人工核对后再引入正文，切勿直接引用）"

ACTION_WRITE = "写入正文"
ACTION_PENDING = "待核占位"

# ---------------------------------------------------------------- 基础工具


def fid_key(s: str) -> str:
    """把 `（3.2.3-1）` / `(4.03)` / `3.2.3-1` 统一成 `3.2.3-1`。"""
    s = (s or "").strip()
    s = s.strip("（）()")
    return re.sub(r"\s+", "", s)


def level_of(status: str) -> str:
    s = (status or "").strip()
    for ch in "高中低":
        if s.startswith(ch):
            return ch
    return "中"


def norm_ws(s: str) -> str:
    """比对用：去掉所有空白。"""
    return re.sub(r"\s+", "", s or "")


def short_reason(caveat: str, limit: int = 60) -> str:
    t = re.sub(r"[*_`\\]", "", (caveat or "").strip())
    t = re.sub(r"\s+", "", t)
    return t if len(t) <= limit else t[:limit] + "…"


RE_LATEX_CMD = re.compile(r"\\[a-zA-Z]{2,}")


def has_pending_leak(text: str, f: "Formula", window: int = 420) -> bool:
    """待核公式的标记之后，是否紧跟着出现了表达式（$$ 块或 LaTeX 命令）。

    「待核」的语义是"表达式不引入正文"，所以不能用 latex 字段有没有值来判
    —— 恰恰相反，待核公式本来就没有 latex；要判的是**标记附近多出了数学内容**。
    """
    pat = re.compile(r"〔公式\s*[（(]\s*" + re.escape(f.fid) + r"\s*[）)][^〕]*〕")
    for m in pat.finditer(text or ""):
        tail = (text or "")[m.end(): m.end() + window]
        nxt = RE_MARKED.search(tail)
        if nxt:
            tail = tail[:nxt.start()]
        if "$$" in tail or RE_LATEX_CMD.search(tail):
            return True
    return False


# ---------------------------------------------------------------- 数据表


class Formula:
    __slots__ = ("fid", "parent", "page", "expr", "latex", "plain", "status",
                 "caveat", "img", "action", "level")

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k, ""))
        self.fid = fid_key(self.fid)
        self.parent = (self.parent or "").strip()
        self.level = level_of(self.status)
        if self.action not in (ACTION_WRITE, ACTION_PENDING):
            # 向后兼容：没写"正文处理"列时按可信度 + 是否有 LaTeX 推断
            if self.latex.strip() and self.level == "高":
                self.action = ACTION_WRITE
            else:
                self.action = ACTION_PENDING
        if self.action == ACTION_WRITE and not self.plain.strip() and self.latex.strip():
            self.plain = nf.latex_to_plain(self.latex)
        if self.action == ACTION_WRITE and not self.latex.strip():
            # 只有纯文本也能用（少数平台不渲染），但 md 双写法会缺一行 → 降级待核
            self.action = ACTION_PENDING

    @property
    def is_write(self) -> bool:
        return self.action == ACTION_WRITE

    def render(self, mode: str = "both") -> str:
        """mode: both=MathJax+纯文本（md）；plain=只纯文本（切片/纯文本平台）。

        md 模式尾部补空行：否则 `> 纯文本：…` 后面紧跟的「式中…」会被 Markdown
        当成 blockquote 的惰性续行吞进引用块（v1.10.2 实测踩到）。
        """
        if self.is_write:
            mark = MARK_WRITE.format(fid=self.fid, level=self.level)
            if mode == "plain":
                return f"{mark}{self.plain}"
            return (f"{mark}\n\n$$\n{self.latex}\n$$\n\n"
                    f"> 纯文本：{self.plain}\n\n")
        mark = MARK_PENDING.format(fid=self.fid)
        note = NOTE_PENDING.format(reason=short_reason(self.caveat))
        return f"{mark}\n{note}\n\n" if mode != "plain" else f"{mark}{note}"

    def display_plain(self) -> str:
        return self.plain if self.is_write else ""


class FormulaTable:
    def __init__(self, formulas=None):
        self.by_fid: dict[str, Formula] = {}
        for f in formulas or []:
            self.by_fid[f.fid] = f

    def __bool__(self):
        return bool(self.by_fid)

    def __len__(self):
        return len(self.by_fid)

    # ---- 载入

    @classmethod
    def from_xlsx(cls, path: str) -> "FormulaTable":
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            return cls.from_workbook(wb)
        finally:
            wb.close()

    @classmethod
    def from_workbook(cls, wb) -> "FormulaTable":
        if SHEET not in wb.sheetnames:
            return cls()
        rows = list(wb[SHEET].iter_rows(values_only=True))
        if not rows:
            return cls()
        head = [(str(x or "").strip()) for x in rows[0]]

        def g(row, *names):
            for n in names:
                if n in head:
                    i = head.index(n)
                    v = row[i] if i < len(row) else ""
                    return "" if v is None else str(v).strip()
            return ""

        out = []
        for r in rows[1:]:
            if not any(r):
                continue
            fid = g(r, "公式编号")
            if not fid:
                continue
            out.append(Formula(
                fid=fid,
                parent=g(r, "所属条文"),
                page=g(r, "PDF页"),
                expr=g(r, "规范化表达式（整理，非原文逐字）", "规范化表达式", "表达式"),
                latex=g(r, "LaTeX", "LaTeX 渲染式"),
                plain=g(r, "纯文本", "纯文本对照"),
                status=g(r, "可信度"),
                caveat=g(r, "存疑说明", "存疑"),
                img=g(r, "渲染图", "渲染图文件"),
                action=g(r, "正文处理", "是否引入正文"),
            ))
        return cls(out)

    @classmethod
    def from_data(cls, data) -> "FormulaTable":
        """data: [dict]，键同 Formula。"""
        return cls([Formula(**d) for d in data])

    # ---- 注入

    def inject(self, text: str, mode: str = "both"):
        """把 text 里的裸占位替换成渲染块。返回 (新文本, 统计 dict)。"""
        if not text:
            return text, {"replaced": 0, "orphan": []}
        stats = {"replaced": 0, "orphan": []}

        def rep(m):
            key = fid_key(m.group(1))
            f = self.by_fid.get(key)
            if not f:
                stats["orphan"].append(key)
                return m.group(0)
            stats["replaced"] += 1
            return f.render(mode)

        return RE_PLACEHOLDER.sub(rep, text), stats

    # ---- 校验

    def check(self, xlsx_path: str, md_dir: str | None = None):
        errors: list[str] = []
        warns: list[str] = []
        info: dict = {"total": len(self.by_fid),
                      "write": sum(1 for f in self.by_fid.values() if f.is_write),
                      "pending": sum(1 for f in self.by_fid.values() if not f.is_write)}

        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
        try:
            # ---------- 语料库原文里的裸占位 ----------
            body_by_parent: dict[str, str] = {}
            if CORPUS_SHEET in wb.sheetnames:
                rows = list(wb[CORPUS_SHEET].iter_rows(values_only=True))
                head = [str(x or "").strip() for x in rows[0]]
                i_no = head.index("条文号") if "条文号" in head else None
                i_body = head.index(BODY_COL) if BODY_COL in head else None
                if i_no is not None and i_body is not None:
                    for r in rows[1:]:
                        if not any(r):
                            continue
                        no = str(r[i_no] or "").strip()
                        body = str(r[i_body] or "")
                        body_by_parent[no] = body
            ph_seen: dict[str, int] = {}
            for no, body in body_by_parent.items():
                for m in RE_PLACEHOLDER.finditer(body):
                    ph_seen[fid_key(m.group(1))] = ph_seen.get(fid_key(m.group(1)), 0) + 1
            for key in ph_seen:
                if key not in self.by_fid:
                    errors.append(f"G3 正文裸占位〔公式（{key}）〕在「{SHEET}」中没有对应行")
            for key, f in self.by_fid.items():
                if key not in ph_seen and not f.is_write:
                    warns.append(f"G3 公式明细登记了（{key}）但正文无占位（可能只出现在表格/附录）")
            info["placeholders"] = sum(ph_seen.values())

            # ---------- RAG切片 ----------
            slices: list[tuple[str, str, str]] = []   # (父条号, 层级, 全文)
            if SLICE_SHEET in wb.sheetnames:
                rows = list(wb[SLICE_SHEET].iter_rows(values_only=True))
                head = [str(x or "").strip() for x in rows[0]]
                idx = {n: head.index(n) for n in (PARENT_COL, SLICE_TEXT_COL) if n in head}
                i_lv = head.index("切片层级") if "切片层级" in head else None
                if len(idx) == 2:
                    for r in rows[1:]:
                        if not any(r):
                            continue
                        slices.append((str(r[idx[PARENT_COL]] or "").strip(),
                                       str(r[i_lv] or "") if i_lv is not None else "",
                                       str(r[idx[SLICE_TEXT_COL]] or "")))

            slice_all = "\n".join(s[2] for s in slices) if slices else ""
            # 父块切片文本：公式必须落在**它所属条**的父块里（不能靠别的条的
            # 「引用条文」闭合段蹭到全文检索里蒙混过关）
            parent_text: dict[str, str] = {}
            for pno, lv, txt in slices:
                if "父块" in lv:
                    parent_text[pno] = parent_text.get(pno, "") + "\n" + txt
            if slices and self.by_fid:
                # G1 分级：切片里不得残留裸占位
                left = [fid_key(m.group(1)) for m in RE_PLACEHOLDER.finditer(slice_all)]
                if left:
                    errors.append(f"G1 切片中仍有 {len(left)} 个裸占位未注入公式：{sorted(set(left))[:6]}")
                # G3 对账：标记块必须能对上明细
                marked = {}
                for m in RE_MARKED.finditer(slice_all):
                    key = fid_key(m.group(1))
                    marked[key] = marked.get(key, 0) + 1
                for key in marked:
                    if key not in self.by_fid:
                        errors.append(f"G3 切片标记〔公式（{key}）…〕在「{SHEET}」中无对应行")
                # G1 分级（按所属条定位）+ G4 完整性
                for key, f in self.by_fid.items():
                    ptext = parent_text.get(f.parent, "")
                    if f.is_write:
                        if not ptext:
                            errors.append(f"G1 公式（{key}）所属条 {f.parent} 没有父块切片")
                        elif norm_ws(f.plain) not in norm_ws(ptext):
                            errors.append(
                                f"G1 高可信公式（{key}）不在所属条 {f.parent} 的父块切片里"
                                f"（检索命中该条也拿不到公式）")
                        if ptext and "式中" in body_by_parent.get(f.parent, "") and "式中" not in ptext:
                            errors.append(f"G4 父块切片 {f.parent} 有公式（{key}）但缺「式中」变量定义")
                    else:
                        if ptext and "待核" not in ptext:
                            errors.append(f"G1 待核公式（{key}）的所属条 {f.parent} 父块里没有「待核」标记")
                        if norm_ws(f.latex) and norm_ws(f.latex) in norm_ws(ptext):
                            errors.append(f"G1 待核公式（{key}）的表达式泄漏进了父块切片")
                        if ptext and has_pending_leak(ptext, f):
                            errors.append(f"G1 待核公式（{key}）标记后出现了表达式（$$ 或 LaTeX 命令），违反「不引入正文」")

            info["slices"] = len(slices)

            # ---------- md ----------
            if md_dir and os.path.isdir(md_dir):
                md_files = [os.path.join(md_dir, n) for n in sorted(os.listdir(md_dir))
                            if n.lower().endswith(".md")]
                joined = []
                for p in md_files:
                    with open(p, encoding="utf-8") as fh:
                        joined.append((p, fh.read()))
                allmd = "\n".join(t for _, t in joined)
                info["md_files"] = len(joined)
                n_dd = allmd.count("$$") // 2
                n_plain = len(re.findall(r"^>\s*纯文本[：:]", allmd, flags=re.M))
                info["md_display_blocks"] = n_dd
                info["md_plain_lines"] = n_plain
                if n_dd != n_plain:
                    errors.append(f"G2 md 里 $$ 块 {n_dd} 个，`> 纯文本：` 行 {n_plain} 行，对不上")

                # 逐 md：批注标记 ↔ 明细；G4 完整性（式中同文件）
                for p, txt in joined:
                    for m in RE_PLACEHOLDER.finditer(txt):
                        errors.append(f"G2 {os.path.basename(p)} 仍含裸占位（{fid_key(m.group(1))}），未做双写法")
                        break
                    for m in RE_MARKED.finditer(txt):
                        key = fid_key(m.group(1))
                        f = self.by_fid.get(key)
                        if not f:
                            errors.append(f"G3 {os.path.basename(p)} 的标记〔公式（{key}）〕在明细里没有")
                            continue
                        if f.is_write and norm_ws(f.plain) not in norm_ws(txt):
                            errors.append(f"G1 {os.path.basename(p)} 含（{key}）标记但没有表达式")
                        if (not f.is_write) and "待核" not in m.group(0):
                            errors.append(f"G1 {os.path.basename(p)} 的低可信公式（{key}）未标「待核」")
                        if (not f.is_write) and has_pending_leak(txt, f):
                            errors.append(f"G1 {os.path.basename(p)} 的待核公式（{key}）后面跟了表达式")
                        # G4：该条原文有「式中」时，公式必须与它同文件
                        body = body_by_parent.get(f.parent, "")
                        if "式中" in body and "式中" not in txt:
                            errors.append(f"G4 {os.path.basename(p)} 有公式（{key}）但缺「式中」变量定义")
                # G2：LaTeX 语法 + 裸 LaTeX
                for p, txt in joined:
                    try:
                        issues = nf.check_text(txt)
                    except Exception as e:      # noqa: BLE001
                        warns.append(f"G2 {os.path.basename(p)} 语法检查异常：{e}")
                        continue
                    for it in issues[:5]:
                        errors.append(f"G2 {os.path.basename(p)} 公式/表格问题：{it}")
        finally:
            wb.close()
        return errors, warns, info


# ---------------------------------------------------------------- CLI


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="公式明细 ↔ 正文/切片/md 的落地与校验")
    ap.add_argument("--xlsx", required=False, help="语料库 xlsx（含「公式明细」页签）")
    ap.add_argument("--md", help="导出 markdown 目录")
    ap.add_argument("--check", action="store_true", help="只校验，不写文件")
    ap.add_argument("--in-place", action="store_true", help="就地写入 md（仅 md 模式支持）")
    ap.add_argument("--style", choices=["both", "plain"], default="both",
                    help="md 注入写法：both=MathJax+纯文本（默认）；plain=只纯文本")
    ap.add_argument("--preview", help="打印某条条文注入后的效果（传条号）")
    ap.add_argument("--json", help="把校验结果写成 JSON")
    a = ap.parse_args(argv)

    if not a.xlsx:
        ap.error("需要 --xlsx 指定语料库（用来读「公式明细」）")
    table = FormulaTable.from_xlsx(a.xlsx)
    if not table:
        print(f"⚠ 没有读到公式明细：{a.xlsx}")
        return 0

    print(f"公式明细：{len(table)} 条 ｜ 写入正文 {sum(1 for f in table.by_fid.values() if f.is_write)}"
          f" ｜ 待核占位 {sum(1 for f in table.by_fid.values() if not f.is_write)}")

    if a.preview:
        wb = openpyxl.load_workbook(a.xlsx, read_only=True, data_only=True)
        text = ""
        try:
            rows = list(wb[CORPUS_SHEET].iter_rows(values_only=True))
            head = [str(x or "").strip() for x in rows[0]]
            i_no, i_body = head.index("条文号"), head.index(BODY_COL)
            for r in rows[1:]:
                if str(r[i_no] or "").strip() == a.preview:
                    text = str(r[i_body] or "")
                    break
        finally:
            wb.close()
        if not text:
            print(f"⚠ 没找到条文 {a.preview}")
            return 2
        injected, st = table.inject(text, a.style)
        print(f"\n--- 原文（占位态） ---\n{text}\n\n--- 注入后（{a.style}） ---\n{injected}\n"
              f"--- 替换 {st['replaced']} 处，孤儿 {st['orphan']} ---")
        return 0

    if a.check or not a.in_place:
        errors, warns, info = table.check(a.xlsx, a.md)
        print(f"统计：{info}")
        for w in warns:
            print(f"  ⚠ {w}")
        for e in errors:
            print(f"  ✗ {e}")
        if a.json:
            import json
            with open(a.json, "w", encoding="utf-8") as f:
                json.dump({"errors": errors, "warns": warns, "info": info},
                          f, ensure_ascii=False, indent=1)
        if errors:
            print(f"\n公式专项校验未通过：{len(errors)} 项错误")
            return 1
        print("\n✅ 公式专项校验通过（分级/双写法/对账/完整性）")
        return 0

    # ---- 注入 md
    if not a.md:
        ap.error("--in-place 需要同时给 --md")
    changed = 0
    total_rep = 0
    orphans = set()
    for name in sorted(os.listdir(a.md)):
        if not name.lower().endswith(".md"):
            continue
        p = os.path.join(a.md, name)
        with open(p, encoding="utf-8") as fh:
            txt = fh.read()
        new, st = table.inject(txt, a.style)
        total_rep += st["replaced"]
        orphans.update(st["orphan"])
        if new != txt:
            with open(p, "w", encoding="utf-8", newline="") as fh:
                fh.write(new)
            changed += 1
    print(f"注入完成：{changed} 个 md 被修改，共替换 {total_rep} 处公式")
    if orphans:
        print(f"⚠ 明细里没有的公式编号：{sorted(orphans)}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
