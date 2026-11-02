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
        # 找不到时**不抛错**：build_table_slices 等纯函数不依赖 XLSX，
        # 生成流水线（先算切片、后写 xlsx）需要能加载本模块。
        # 真正用到 XLSX 的 load_* 会自己报错。
        return None
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


def load_facts(structured=False):
    """读表格事实 → {表号: [事实句 或 dict, ...]}

    structured=True 时优先读 _pipeline/table_facts.json（带 row/row_label/col_label），
    按行拆切片需要这些字段；读不到则回退到 xlsx 页签的纯字符串。
    """
    if structured:
        import json as _json
        if not XLSX:
            return {}
        p = Path(XLSX).parent / "_pipeline" / "table_facts.json"
        if p.exists():
            try:
                raw = _json.loads(p.read_text(encoding="utf-8"))
                out2 = defaultdict(list)
                for f in raw:
                    out2[str(f["table"])].append(f)
                return dict(out2)
            except Exception:
                pass

    # 回退：读 xlsx 的「表格事实」页签（纯字符串，无行/列字段）
    import openpyxl
    try:
        wb = openpyxl.load_workbook(XLSX, read_only=True)
    except Exception:
        return {}
    if "表格事实" not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb["表格事实"]
    out = defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] and r[6]:
            out[str(r[0])].append(str(r[6]))
    wb.close()
    return dict(out)


def load_calibrated():
    """读人工/视觉校准的表格结构 → {表号: {cols, row_labels, grid_rows}}

    优先从语料库的「表格事实」页签反推（列标签/行标签都在里面）；
    没有显式 grid_rows 时返回空，调用方回退到 OCR 网格并标注不可靠。
    """
    import json as _json
    if not XLSX:
        return {}
    p = Path(XLSX).parent / "_pipeline" / "table_columns_calibrated.json"
    if not p.exists():
        return {}
    try:
        return _json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


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


MAX_TABLE_FACTS = 40      # 单表切片里最多列多少条事实句（超过则均匀抽样）


def build_table_slices(meta, link, tables, facts=None, calibrated=None):
    """把表格做成独立子块，父条号指回条文。

    两种切片粒度（按事实数自动选）：

    **少事实（≤40 条）**：一张表 = 一个子块
        ID: 编号::条号#表号

    **多事实（>40 条，如 4.2.10 有 140 条）**：**按行拆成多个子块**
        ID: 编号::条号#表号#r行号
        为什么不抽样：实测抽样 40/140 后严格问答从 6/6 掉到 2/6 ——
        答案被抽掉的格子就永远查不到了。
        按行拆的好处：每片仍自包含（一行的 15 个值都在），
        切得更细反而让检索更准（BM25 不会被无关行稀释）。
    """
    facts = facts or {}

    out = []
    for tid, tno in link.items():
        t = tables.get(tno)
        m = meta.get(tid)
        if not t or not m:
            continue
        base_hdr = (f"《{m['name']}》{m['code']}｜{m['zhang']}｜{m['jie']}｜"
                    f"第{tid}条｜{t['title'] or tno}｜表格")
        fs = facts.get(tno) or []
        cal = (calibrated or {}).get(tno) or {}
        grows = cal.get("grid_rows")
        # 反向速查片需要这三样（在校准数据里）
        grows_data = grows
        ccols = {int(k): v for k, v in (cal.get("cols") or {}).items() if v}
        rlabels_data = {int(k): v for k, v in (cal.get("row_labels") or {}).items()}

        # ---- 事实多 → 按行拆成多个子块
        #      facts 可能是纯字符串（从 xlsx 读）或 dict（从 JSON 读），两种都处理
        if len(fs) > MAX_TABLE_FACTS:
            from collections import OrderedDict
            byrow = OrderedDict()
            for i, f in enumerate(fs):
                if isinstance(f, dict):
                    r, txt, lab = f.get("row", i), f.get("text", ""), f.get("row_label", "")
                else:
                    # 字符串形态：用序号当行号（每 MAX 条归一组，避免无法分组）
                    r, txt, lab = i // 15, str(f), ""
                byrow.setdefault(r, []).append((txt, lab))
            for r, rowitems in byrow.items():
                lab = next((x[1] for x in rowitems if x[1]), "")
                hdr = f"{base_hdr}#r{r}｜{lab}"
                LL = [f"本片为「{t['title'] or tno}」中「{lab}」这一行；"
                      f"下表每条均为：本行对象 与 该列对象 之间的防火间距（单位 m）。"]
                for txt, _ in rowitems:
                    LL.append("  · " + txt)
                LL.append(f"  ⚠ 以上事实仅限本表所列对象（{t['title']}）；"
                          f"若问题对象不在本表范围内，请勿套用。")
                LL.append(f"  依据：《{m['name']}》{m['code']} 第{tid}条，原文第"
                          f"{cal.get('page', '')}页。")
                body = "\n".join(LL)
                out.append(dict(
                    id=f"{m['code'].replace(' ', '')}::{tid}#{tno}#r{r}",
                    parent=tid, level=f"表（子块·{tno}·r{r}）",
                    title=f"{t['title'] or tno}（{lab}）",
                    text=hdr + "\n" + body, header=hdr, body=body,
                    chars=len(body), page_from=None, page_to=None))

            # ---- 额外生成一个「速查汇总」分片，专供**反向检索**
            #      为什么需要：按行拆片后，"哪些对象与X的间距是10m" 这类反向问题
            #      需要覆盖全部行分片，而 top-k 只取少数块 → 必然漏行。
            #      实测：「厂内主要道路」列 5 个 10m 恰好都落在高分片里才答全，
            #      换「厂内铁路」列就会漏。
            #      汇总片把所有「列 → 值 → 行清单」倒排出来，一张片就能答全。
            if grows_data and ccols:
                summ = []
                kk = sorted(ccols)
                # 行标签太长（30~50 字），全部内联会让速查片膨胀到 4600+ 字。
                # 改用「行号引用」：先给行清单，再用编号指代。
                rl = sorted(int(x) for x in grows_data)
                summ.append("行号对照（下文用 R序号 指代）：")
                for i, r in enumerate(rl):
                    summ.append(f"  R{i}={rlabels_data.get(r, '')}")
                summ.append("")
                summ.append("列 → 值 → 对象行号：")
                for ci, col in enumerate(kk):
                    buckets = {}
                    for i, r in enumerate(rl):
                        vals = grows_data[str(r)]
                        v = vals[ci] if ci < len(vals) else None
                        if v is None:
                            continue
                        buckets.setdefault(str(v), []).append(f"R{i}")
                    if not buckets:
                        continue
                    parts = []
                    for v in sorted(buckets, key=lambda x: (len(x), x)):
                        parts.append(f"{v}={{{','.join(buckets[v])}}}")
                    summ.append(f"  「{ccols[col]}」：" + "；".join(parts))
                if summ:
                    hdr2 = base_hdr + "｜反向速查（列→值→行号）"
                    # 反向问题（"哪些对象与X的间距是Y"）的查询词主要是
                    # 表名 + 列名 + 数值。速查片本身短，BM25 里长度归一化会吃亏
                    # （实测排第 14 名）。所以开头把关键词**显式重复一遍**，
                    # 提高词频与命中率。
                    _keys = "；".join(ccols[k] for k in sorted(ccols))
                    _title = t["title"] or tno
                    body2 = (
                        f"反向速查表：{_title}。\n"
                        f"本表全部列名：{_keys}。\n"
                        f"本表全部行对象："
                        + "、".join(rlabels_data.get(r, "") for r in rl) + "。\n"
                        f"（上列关键词用于反向检索："
                        f"哪些对象与某列的某项数值对应、某列等于某值的行有哪些。）\n\n"
                        "本片按「列 → 数值 → 等于此值的行号」倒排，"
                        "把行号换回下面的行名称即可作答。\n"
                        + "\n".join(summ) +
                        f"\n  ⚠ 仅限本表（{_title}）；某值下未出现的行号表示该格无要求。\n"
                        f"  依据：《{m['name']}》{m['code']} 第{tid}条，"
                        f"原文第{cal.get('page', '')}页。")
                    out.append(dict(
                        id=f"{m['code'].replace(' ', '')}::{tid}#{tno}#sum",
                        parent=tid, level=f"表（反向速查·{tno}）",
                        title=f"{t['title'] or tno}（反向速查：列→值→对象）",
                        text=hdr2 + "\n" + body2, header=hdr2, body=body2,
                        chars=len(body2), page_from=None, page_to=None))
            continue

        # ---- 事实少 → 整表一个子块
        header = base_hdr
        L = []
        if fs:
            L.append(f"本片为「{t['title'] or tno}」全表；"
                     f"下表每条均为：行对象 与 列对象 之间的防火间距（单位 m）。")
            L.append(f"事实（共 {len(fs)} 条，可直接引用）：")
            for s in fs:
                txt = s.get("text", "") if isinstance(s, dict) else str(s)
                L.append("  · " + txt)
            L.append(f"  ⚠ 以上事实仅限本表所列对象（{t['title']}）；"
                     f"若问题对象不在本表范围内，请勿套用。")

        # 数据行：有视觉校准的 grid_rows 就用它，**不输出 OCR 坏网格**
        #   实测 OCR 网格把 "12 12" 粘成 "1212"、把行标签混进数据列，
        #   和上面的事实句直接矛盾。
        if grows:
            keys = sorted((int(k) for k in cal.get("cols", {})), key=int)
            rlabels = {int(k): v for k, v in (cal.get("row_labels") or {}).items()}
            L.append("数据行（视觉读图 + 逐块核对，未列出的格为「—」）：")
            L.append("  行标签 | " + " | ".join(str(k) for k in keys))
            for r in sorted(int(x) for x in grows):
                vals = grows[str(r)]
                cells = [str(vals[i]) if i < len(vals) and vals[i] is not None else "—"
                         for i in range(len(keys))]
                L.append(f"  {rlabels.get(r, '')[:30]} | " + " | ".join(cells))
        else:
            L.append("原始数据行（OCR 网格，列对齐可能有偏差，以事实句为准）：")
            L.append("| " + " | ".join(t["header"]) + " |")
            L.append("|" + "|".join(["---"] * len(t["header"])) + "|")
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
        self.facts = load_facts(structured=True)   # 表号 -> [fact dict]，带 row/col 便于按行拆片
        self.calibrated = load_calibrated()   # 表号 -> {cols, row_labels, grid_rows}
        base = load_slices()
        self.table_slices = build_table_slices(load_meta(), self.link, self.tables, self.facts, self.calibrated)

        # 去重：语料库的「RAG切片」页签里**已含**表格子块（生成时写进去的），
        # 这里又 build 一遍 → 同一 ID 出现两份，一份含事实句一份不含。
        # 实测后果：检索有时命中旧的那份，事实句拿不到 → 表格题答不准。
        # 处理：以 runtime 构建的 table_slices 为准（它含事实句），丢弃 base 里的表格块。
        if use_table_slices:
            base = [s for s in base if "表（子块" not in s["level"]]
            self.slices = base + self.table_slices
        else:
            self.slices = base
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
    def render(self, parent, with_table=True, with_refs=False):
        """父块全文（章/节/条元数据头 + 全部款）+ 关联表格

        **必须带上关联表格**：表格是条文的一部分。
        实测踩过：条文写"不应小于表4.1.7 的规定"，但检索只返回这一句，
        表里的事实句/数值全拿不到 → 表格题 0/8 答对。

        表格内容**直接用已构建的 table_slices**（含按行分片与视觉校准数据），
        不要再走 table_md 从 tables 重新渲染 —— 实测那样会漏掉分片、
        且用回 OCR 坏网格，与事实句自相矛盾。
        """
        base = self.parents[parent]["text"] if parent in self.parents else ""
        if not base:
            return ""
        parts = [base]

        # 款 + 附录子块：都在 children 里。
        #   原来 render 只拼 base，导致附录父块只有 351 字摘要，
        #   问"二硫化碳库属于几类"取不到具体条目（实测 12/14 附录题失败）。
        for k in self.children.get(parent, []):
            if "表（" in k["level"]:
                continue                      # 表格行分片在下面单独加
            parts.append(k["body"])

        if with_table:
            ts = [s for s in self.table_slices if s["parent"] == parent]
            if ts:
                parts.extend(s["body"] for s in ts)
            else:
                tm = self.table_md(parent)
                if tm:
                    parts.append(tm)
        if with_refs:
            for rno in self.refs_of(parent):
                rp = self.parents.get(rno)
                if rp:
                    parts.append(f"【被引用条文 {rno}】\n" + rp["text"])
        return "\n".join(parts)

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

        # 事实句优先（人工校准列名后生成，读取方直接读句子即可）
        #   **不抽样**：检索侧已按行拆片控制单块大小，
        #   这里若再抽样会让答案整段丢失（实测抽样后严格问答 6/6 -> 2/6）。
        fs = self.facts.get(tno) or []
        if fs:
            L.append(f"事实（共 {len(fs)} 条，可直接引用）：")
            for s in fs:
                txt = s.get("text", "") if isinstance(s, dict) else str(s)
                L.append("  · " + txt)
            # 防止误用：事实句只适用于**本表对象**。
            # 实测踩过：问"生产厂房与厂内主要道路"时，检索把堆场表(4.2.9)的
            # 事实句给了 AI，用户会误以为那就是答案。
            L.append(f"  ⚠ 以上事实仅限本表所列对象（{t['title']}）；"
                     f"若问题对象不在本表范围内，请勿套用，应另查其他表或说明查不到。")

        # 列索引 / 数据行：
        #   · 有校准数据就用它（视觉读图逐格核对过），**不输出 OCR 坏网格**
        #     实测：OCR 列索引是 "可燃液体罐区甲、乙类(总储量(00)"、
        #     "（（甲、乙类2由）仓库除外）" 这种碎片，与上面正确的事实句矛盾。
        #   · 没有校准时才回退 OCR，并明确标注不可靠。
        cal = getattr(self, "calibrated", {}).get(tno) or {}
        grows = cal.get("grid_rows")
        ccols = {int(k): v for k, v in (cal.get("cols") or {}).items() if v}
        if grows:
            keys = sorted(ccols) or sorted((int(k) for k in cal.get("cols", {})), key=int)
            rlabels = {int(k): v for k, v in (cal.get("row_labels") or {}).items()}
            if ccols:
                L.append("列索引（已校准）：")
                L.append("  " + "；".join(f"列{k}={ccols[k]}" for k in sorted(ccols)))

            # 文字格：有的行不是米数，而是「1.5倍杆(塔)高度」这类表述。
            # 必须**填回到数据行里**，否则该行显示成 | — | — | — | — | —，
            # 读取方会误以为这一行"全部无要求"（实测踩过）。
            # text_cells 的键形如 "col1"/"col2"，要转成真实列号。
            tc_raw = cal.get("text_cells") or {}
            tc_by_row = {}
            for r, d in tc_raw.items():
                m2 = {}
                for kk, vv in d.items():
                    kk2 = str(kk)
                    if kk2.startswith("col"):
                        kk2 = kk2[3:]
                    try:
                        m2[int(kk2)] = vv
                    except ValueError:
                        continue
                tc_by_row[int(r)] = m2

            L.append("数据行（视觉读图 + 逐块核对；"
                     "「—」=原表无要求，非数字格按文字给出）：")
            L.append("  行标签 | " + " | ".join(str(k) for k in keys))
            for r in sorted(int(x) for x in grows):
                vals = grows[str(r)]
                tc = tc_by_row.get(r, {})
                cells = []
                for i in range(len(keys)):
                    k = keys[i]
                    if k in tc:
                        cells.append(str(tc[k]))       # 文字格优先
                    else:
                        v = vals[i] if i < len(vals) else None
                        cells.append(str(v) if v is not None else "—")
                L.append(f"  {rlabels.get(r, '')} | " + " | ".join(cells))
            return "\n".join(L)

        # 无校准数据：回退 OCR 网格，并明确标注不可靠
        named = [(i + 1, c) for i, c in enumerate(cols) if c]
        if named:
            L.append("列索引（OCR 原样，列对齐可能有偏差）：")
            L.append("  " + "；".join(f"列{i}={c}" for i, c in named))
        body = grid[data_start:] if data_start < len(grid) else grid
        L.append("原始数据行（OCR 网格，以事实句为准）：")
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
        """父块（头 + 正文 + 全部款）+ 关联表格 + 附录明细 + 被引用的条文

        · 款必须带上：条文正文常只有一句"应符合下列规定"，细节全在款里。
        · 附录/表格子块必须带上：附录父块只是摘要，118 条明细在子块里；
          不带的话导出 md 只有 900 字，查"涤纶属于几类"根本查不到（实测踩过）。
        · 被引用条文：规范里 17 条写"除应符合第X条外"，不带就是悬空答案。
        """
        p = self.parents.get(parent)
        if not p:
            return ""
        parts = [p["header"], p["body"]]

        # 款/附录子块（都在 children 里）——只排除表格行分片
        for k in self.children.get(parent, []):
            if "表（" in k["level"]:
                continue
            parts.append(k["body"])

        tmd = self.table_md(parent)
        if tmd:
            parts.append(tmd)

        # 附录明细 / 表格行分片 / 反向速查
        #   注意：table_md() 已经渲染过一次表格内容；若这里再拼 table_slices，
        #   表格会出现两遍（实测 4.1.7 的「事实（共」和「数据行」各出现 2 次）。
        #   所以：**只有该 parent 的 table_md 为空时**才用 table_slices，
        #   附录没有 table_md，正好走这条路。
        if not tmd:
            for ts in self.table_slices:
                if ts["parent"] == parent:
                    parts.append(ts["body"])

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
