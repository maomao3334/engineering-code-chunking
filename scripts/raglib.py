# -*- coding: utf-8 -*-
"""
父子块检索库（零成本版，不调用任何 API）

数据源：工作区内版本号最大的 *语料库_v*.xlsx 「RAG切片」页
        （可用环境变量 CORPUS_XLSX 覆盖）
        241 个父块（整条）+ 520 个子块（各款）= 761 个切片

两种检索模式，用来验证会上争的那个问题：
  flat   —— 平铺检索，谁分高给谁（= 直接向量库的默认行为，会漏款）
  parent —— 命中任一款，就把整条父块（含全部款 + 章/节/条 头）一起交出去

分词：中文按字二元组（bigram），数字/单位/条号按词，不依赖 jieba。
打分：BM25（手写，不依赖 rank_bm25）。
"""
import math
import os
import re
from collections import Counter, defaultdict
from pathlib import Path


def _find_corpus():
    """定位语料库：环境变量 CORPUS_XLSX 优先，否则取工作区里版本号最大的那个。

    不再硬编码文件名——语料库会随修复迭代版本（v2→v3→v4…），
    硬编码会导致"改了数据但程序还在读旧文件"这种极难排查的问题。
    """
    env = os.environ.get("CORPUS_XLSX")
    if env and Path(env).exists():
        return str(Path(env))
    root = Path(__file__).resolve().parent.parent
    cands = []
    for p in root.glob("*语料库_v*.xlsx"):
        m = re.search(r"_v(\d+)\.xlsx$", p.name)
        if m:
            cands.append((int(m.group(1)), p))
    if not cands:
        raise FileNotFoundError(f"在 {root} 找不到 *语料库_v*.xlsx")
    return str(max(cands)[1])


XLSX = _find_corpus()

_CJK = re.compile(r"[\u4e00-\u9fff]+")
_ALNUM = re.compile(r"[a-z0-9]+(?:\.[0-9]+)*")
_REF = re.compile(r"第(\d+\.\d+\.\d+)条")

# 「使用期/管理类」问题的标记词。这类问题的答案在第12章「使用与维护」或
# 《高层民用建筑消防安全管理规定》里，而不是第3章的设计条文里。
USE_MARK = ("占用", "被占", "占着", "占了", "堵", "堆放", "堆了", "堆东西",
            "停放", "停车", "充电", "检查", "处罚", "罚款", "违法", "物业",
            "管理", "维护", "拒不改正", "怎么处理", "怎么办")


def _is_use_clause(header):
    h = header or ""
    return ("第12章" in h) or ("使用与维护" in h) or ("应急管理部令" in h)


def tokenize(s: str):
    """中文 bigram + 数字/单位/编号 token"""
    s = (s or "").lower()
    toks = []
    for m in _CJK.finditer(s):
        run = m.group()
        if len(run) == 1:
            toks.append(run)
        else:
            toks.extend(run[i:i + 2] for i in range(len(run) - 1))
    for m in _ALNUM.finditer(s):
        toks.append(m.group())
    return toks


def load_slices():
    """读 RAG切片 页，拆出元数据头 / 正文"""
    import openpyxl
    wb = openpyxl.load_workbook(XLSX, read_only=True)
    ws = wb["RAG切片"]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        sid, parent, level, title, text = r[0], r[1], r[2], r[3], r[4]
        if not sid or not text:
            continue
        text = str(text)
        if "\n" in text:
            header, body = text.split("\n", 1)
        else:
            header, body = text, ""
        out.append(dict(
            id=str(sid), parent=str(parent), level=str(level), title=str(title or ""),
            text=text, header=header, body=body,
            chars=r[5] or 0, page_from=r[7], page_to=r[8],
        ))
    return out


def load_tables():
    """读「表格明细」，返回 (条号->表号, 表号->{title,header,rows})

    表在 v2 库里是"孤岛"：只存在「表格明细」页签，没有进切片。
    这里把它读出来，好挂回父块上下文。
    """
    import openpyxl
    wb = openpyxl.load_workbook(XLSX, read_only=True)
    ws = wb["语料库"]
    link = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[14]:
            link[str(r[4])] = str(r[14])

    ws = wb["表格明细"]
    cells = defaultdict(list)          # 表号 -> [(行号, 列序号, 列标题, 值)]
    title = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        tno = str(r[0])
        title[tno] = str(r[1] or "")
        cells[tno].append((int(r[4]), int(r[5]), str(r[6] or ""), str(r[7] or "")))

    tables = {}
    for tno, cs in cells.items():
        heads = {}
        rows = defaultdict(dict)
        for rn, cn, head, val in cs:
            heads.setdefault(cn, head)
            rows[rn][cn] = val
        ncol = max(heads) if heads else 0
        # 表头就是第 1 行（OCR 认定的表头行）；数据从第 2 行起，
        # 否则表头会被打印两次（一次作 header，一次作数据行）
        first = min(rows) if rows else 1
        header = [rows[first].get(i, heads.get(i, f"列{i}"))
                  for i in range(1, ncol + 1)]
        body = [[rows[rn].get(i, "") for i in range(1, ncol + 1)]
                for rn in sorted(rows) if rn != first]
        tables[tno] = dict(title=title.get(tno, ""), header=header, rows=body)
    return link, tables


def load_meta():
    """条号 -> 规范全称/编号/章/节/小标题"""
    import openpyxl
    wb = openpyxl.load_workbook(XLSX, read_only=True)
    ws = wb["语料库"]
    meta = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        meta[str(r[4])] = dict(name=str(r[0]), code=str(r[1]), zhang=str(r[2]),
                               jie=str(r[3]), title=str(r[5] or ""))
    return meta


def build_table_slices(meta, link, tables):
    """把表格做成独立子块（切片ID: 编号::条号#表号），父条号指回条文

    这就是硬伤 A 的正解：表格本来在切片体系之外，检索永远捞不到。
    """
    out = []
    for tid, tno in link.items():
        t = tables.get(tno)
        m = meta.get(tid)
        if not t or not m:
            continue
        header = (f"《{m['name']}》{m['code']}｜{m['zhang']}｜{m['jie']}｜"
                  f"第{tid}条｜{t['title'] or tno}｜表格")
        L = ["| " + " | ".join(t["header"]) + " |",
             "|" + "|".join(["---"] * len(t["header"])) + "|"]
        for row in t["rows"]:
            L.append("| " + " | ".join(x or "—" for x in row) + " |")
        body = "\n".join(L)
        out.append(dict(
            id=f"{m['code'].replace(' ', '')}::{tid}#{tno}", parent=tid,
            level=f"表（子块·{tno}）", title=t["title"] or tno,
            text=header + "\n" + body, header=header, body=body,
            chars=len(body), page_from=None, page_to=None,
        ))
    return out


def load_synonyms():
    import json
    from pathlib import Path
    p = Path(__file__).parent / "synonyms.json"
    if not p.exists():
        return []
    return json.loads(p.read_text(encoding="utf-8"))["口语映射"]


def expand_query(q):
    """本地 query 改写：口语 -> 规范术语（零成本版）"""
    extra = []
    for rule in load_synonyms():
        if any(t in q for t in rule["触发"]):
            extra.extend(rule["扩展"])
    return list(dict.fromkeys(extra))


class BM25:
    def __init__(self, docs_tokens, k1=1.5, b=0.75):
        self.N = len(docs_tokens)
        self.k1, self.b = k1, b
        self.tf, self.dl, self.df = [], [], Counter()
        for toks in docs_tokens:
            c = Counter(toks)
            self.tf.append(c)
            self.dl.append(sum(c.values()))
            for t in c:
                self.df[t] += 1
        self.avgdl = (sum(self.dl) / self.N) if self.N else 0.0
        self.idf = {t: math.log(1 + (self.N - d + 0.5) / (d + 0.5)) for t, d in self.df.items()}

    def scores(self, qtoks):
        sc = [0.0] * self.N
        for t in set(qtoks):
            idf = self.idf.get(t)
            if idf is None:
                continue
            for i in range(self.N):
                f = self.tf[i].get(t, 0)
                if f:
                    denom = f + self.k1 * (1 - self.b + self.b * self.dl[i] / self.avgdl)
                    sc[i] += idf * f * (self.k1 + 1) / denom
        return sc


class Retriever:
    def __init__(self, weighted_header=1.0, use_table_slices=False, use_expand=False,
                 child_weight=0.8, use_boost=1.35):
        self.use_expand = use_expand
        self.use_table_slices = use_table_slices
        self.child_weight = child_weight   # 子块分数折算（款太短，太容易"局部匹配"压过正确的条）
        self.use_boost = use_boost         # 「使用期/管理类」问题的章节路由加权
        self.link, self.tables = load_tables()
        base = load_slices()
        self.table_slices = build_table_slices(load_meta(), self.link, self.tables)
        self.slices = base + (self.table_slices if use_table_slices else [])
        self.parents = {}          # 父条号 -> 父块 slice
        self.children = defaultdict(list)
        for s in self.slices:
            if "父块" in s["level"]:
                self.parents[s["parent"]] = s
            else:
                self.children[s["parent"]].append(s)
        self.children = dict(self.children)
        # 索引 = 正文 + 元数据头（头里含 章/节 标题，对"第7章""疏散"这类问法有用）
        docs = []
        for s in self.slices:
            toks = tokenize(s["body"])
            if weighted_header > 0:
                htoks = tokenize(s["header"])
                toks += htoks * max(1, int(round(weighted_header)))
            docs.append(toks)
        self.bm = BM25(docs)
        self.by_id = {s["id"]: i for i, s in enumerate(self.slices)}

    # ---------- 查询侧改写 ----------
    def _scores(self, query, expand_weight=0.6):
        """原问句正常打分；扩展出的规范术语按 expand_weight 加权叠加"""
        sc = self.bm.scores(tokenize(query))
        if self.use_expand:
            ex = expand_query(query)
            if ex:
                ex_sc = self.bm.scores(tokenize(" ".join(ex)))
                sc = [a + expand_weight * b for a, b in zip(sc, ex_sc)]
        if self.use_boost > 1 and any(m in query for m in USE_MARK):
            # 使用期/管理类问题：第12章（使用与维护）和法规条文加权。
            # 不加权的话，"消防车道"会被第3章一堆设计条文抢走，问「被私家车占了怎么办」
            # 却答成"该沿两条长边设消防车道"——这是实测踩过的坑。
            sc = [v * self.use_boost if _is_use_clause(s["header"]) else v
                  for v, s in zip(sc, self.slices)]
        return sc

    # ---------- 两种模式 ----------
    def search_flat(self, query, k=5):
        sc = self._scores(query)
        order = sorted(range(len(self.slices)), key=lambda i: -sc[i])[:k]
        return [dict(slice=self.slices[i], score=round(sc[i], 3)) for i in order if sc[i] > 0]

    def search_parent(self, query, k=5):
        """子块命中 -> 回溯父块整条（含全部款）"""
        sc = self._scores(query)
        best = defaultdict(float)
        hit_child = defaultdict(list)
        for i, s in enumerate(self.slices):
            p = s["parent"]
            lv = s["level"]
            # 只有"款"是碎片，要折算；父块和表格是完整信息单元，不折算
            frag = ("父块" not in lv) and ("表" not in lv)
            v = sc[i] * self.child_weight if frag else sc[i]
            if v > best[p]:
                best[p] = v
            if sc[i] > 0 and "父块" not in lv:
                hit_child[p].append(lv)
        order = sorted(best.items(), key=lambda x: -x[1])[:k]
        out = []
        for p, s0 in order:
            if s0 <= 0 or p not in self.parents:
                continue
            out.append(dict(
                slice=self.parents[p], score=round(s0, 3),
                parent=p, hit_levels=hit_child.get(p, []),
                context=self.render(p),
            ))
        return out

    # ---------- 拼接上下文 ----------
    def render(self, parent):
        """父块全文（已含章/节/条元数据头 + 全部款）"""
        return self.parents[parent]["text"] if parent in self.parents else ""

    def table_md(self, parent):
        """该条引用的表，渲染成「列索引 + 数据行」。

        为什么要列索引：OCR 出来的竖排表头（"厂内主要道路"）会被拆成多行碎片，
        数据行 `| 25 | 12 | 12 |` 里的 12 无法判断属于哪一列 —— **数据读不懂**。
        做法：把表头区按列纵向拼接复原标题，先打印「列号→标题」映射，
        再打印数据行，读取方就能按列序对齐。

        列数对齐：每行列数常不一致（合并单元格被切开），一律补齐到 ncol。
        """
        tno = self.link.get(parent)
        t = self.tables.get(tno) if tno else None
        if not t:
            return ""
        # 原始网格 = header 行 + 数据行（load_tables 已把第 1 行当 header）
        grid = [list(t["header"])] + [list(r) for r in t["rows"]]
        ncol = len(t["header"]) or (len(grid[0]) if grid else 0)

        # 数据行起点：第一行含 ≥2 个纯数字
        def is_num(c):
            return bool(re.fullmatch(r"\d+(?:\.\d+)?", str(c).strip()))

        data_start = len(grid)
        for i, row in enumerate(grid):
            if sum(1 for c in row if is_num(c)) >= 2:
                data_start = i
                break

        # 按列纵向拼表头
        cols = []
        for c in range(ncol):
            parts = []
            for row in grid[:data_start]:
                v = str(row[c]).strip() if c < len(row) else ""
                if v and v not in ("—", "-", "None"):
                    parts.append(v)
            title = "".join(parts)
            # OCR 重复：同一短语被识别两次（"厂内厂内主要道路道路"）
            for _ in range(3):
                dup = re.match(r"^(.{2,5})\1", title)
                if dup:
                    title = dup.group(1) + title[len(dup.group(0)):]
                dup2 = re.search(r"(.{2,5})\1$", title)
                if dup2:
                    title = title[:dup2.start()] + dup2.group(1)
            cols.append(title)

        # 注意：竖排表头的碎片（"道路厂内厂内主要道路道路" / "次要"）**没有修复**。
        # 试过用正则把「上一列溢出到本列的碎片」拼回去，在表4.2.9 上失败——
        # 本质是在猜数据。这类表只能做到「数值在」，列对应关系需人工核对，
        # 已在「表格质量」页签逐表标注。详见 references/pdf-to-corpus-walkthrough.md。

        L = [f"【{t['title'] or tno}】"]
        named = [(i + 1, c) for i, c in enumerate(cols) if c]
        if named:
            L.append("列索引（数据行按此列序对齐）：")
            L.append("  " + "；".join(f"列{i}={c}" for i, c in named))

        body = grid[data_start:] if data_start < len(grid) else grid
        for row in body:
            cells = [str(x or "—") for x in row[:ncol]]
            cells += ["—"] * (ncol - len(cells))
            L.append("| " + " | ".join(cells) + " |")
        return "\n".join(L)

    def refs_of(self, parent):
        """该条正文里交叉引用的其他条（第X.Y.Z条）"""
        p = self.parents.get(parent)
        if not p:
            return []
        return [x for x in dict.fromkeys(_REF.findall(p["body"])) if x != parent]

    def render_expanded(self, parent, with_refs=True):
        """父块（头 + 正文 + 全部款）+ 关联表格 + 被引用的条文

        最后一项是优化项：规范里 17 条条文写着"除应符合第X条外"，
        不把被引用的条带上，答案就是悬空的。
        """
        p = self.parents.get(parent)
        if not p:
            return ""
        parts = [p["header"], p["body"]]
        tmd = self.table_md(parent)
        if tmd:
            parts.append(tmd)
        if with_refs:
            for ref in self.refs_of(parent):
                rp = self.parents.get(ref)
                if not rp:
                    continue
                parts.append(f"【关联条文 {ref}｜{rp['title']}】")
                parts.append(rp["body"])
                rt = self.table_md(ref)
                if rt:
                    parts.append(rt)
        return "\n".join(x for x in parts if x)


if __name__ == "__main__":
    r = Retriever()
    print("切片", len(r.slices), "父块", len(r.parents), "有款的父块", len(r.children))
    for q in ["消防电梯前室的使用面积最小是多少", "建筑防火的总体目标包含哪几个方面"]:
        print("\n问：", q)
        for m in ("flat", "parent"):
            hits = getattr(r, f"search_{m}")(q, k=3)
            top = hits[0]["slice"] if hits else None
            print(f"  [{m}] top1 = {top['parent'] if top else '-'} {top['level'] if top else ''} {top['title'] if top else ''}")
