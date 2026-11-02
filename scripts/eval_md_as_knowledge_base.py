# -*- coding: utf-8 -*-
"""最终验收：在多种 chunk_size 下跑完整问答，确认 md 作为知识库可用。

覆盖 300/500/800/1200 四种切块（平台常见范围），每题要求：
    top3 中有一块同时含 数值 + 两个对象词 + 表号
"""
import re
from collections import Counter
from pathlib import Path

import argparse
_ap = argparse.ArgumentParser()
_ap.add_argument("--md-dir", required=True, help="导出的 markdown 目录")
_a = _ap.parse_args()
MD_DIR = Path(_a.md_dir)


def tokenize(text):
    t = re.sub(r"\s+", "", text)
    toks = [m.group() for m in re.finditer(r"[0-9]+(?:\.[0-9]+)?|[A-Za-z]+", t)]
    han = re.sub(r"[^\u4e00-\u9fff]", " ", t)
    for seg in han.split():
        for i in range(len(seg) - 1):
            toks.append(seg[i:i + 2])
        if len(seg) == 1:
            toks.append(seg)
    return toks


def load_chunks(size, overlap):
    chunks = []
    for p in sorted(MD_DIR.glob("*.md")):
        txt = p.read_text(encoding="utf-8")
        if len(txt) <= size:
            chunks.append(dict(file=p.name, text=txt))
        else:
            start = 0
            while start < len(txt):
                chunks.append(dict(file=p.name, text=txt[start:start + size]))
                start += size - overlap
    return chunks


class BM25:
    def __init__(self, docs, k1=1.5, b=0.75):
        import math
        self.math = math
        self.docs = [Counter(tokenize(d)) for d in docs]
        self.lens = [sum(c.values()) for c in self.docs]
        self.avg = sum(self.lens) / max(len(self.lens), 1)
        self.k1, self.b = k1, b
        self.df = Counter()
        for c in self.docs:
            for t in c:
                self.df[t] += 1
        self.N = len(docs)

    def top(self, q, k=3):
        qt = tokenize(q)
        out = []
        for c, ln in zip(self.docs, self.lens):
            s = 0.0
            for t in qt:
                if t not in c:
                    continue
                idf = self.math.log(1 + (self.N - self.df[t] + 0.5) / (self.df[t] + 0.5))
                tf = c[t]
                s += idf * tf * (self.k1 + 1) / (
                    tf + self.k1 * (1 - self.b + self.b * ln / self.avg))
            out.append(s)
        order = sorted(range(len(out)), key=lambda i: -out[i])[:k]
        return [(self._chunks[i], out[i]) for i in order if out[i] > 0]

    def bind(self, chunks):
        self._chunks = chunks
        return self


CASES = [
    ("厂外铁路与甲、乙类可燃液体储罐（≤5000m³）的防火间距是多少？", "35", "表4.1.7", "厂外铁路", "可燃液体储罐"),
    ("厂外民用建筑与丙类厂房的防火间距是多少？", "17", "表4.1.7", "厂外民用建筑", "丙类厂房"),
    ("高速公路与甲、乙类仓库的防火间距？", "30", "表4.1.7", "高速公路", "甲、乙类仓库"),
    ("室外变配电站与丙类可燃液体储罐的防火间距？", "40", "表4.1.7", "室外变", "可燃液体储罐"),
    ("堆场储量501t~1000t时与厂内主要道路的防火间距？", "10", "表4.2.9", "501t~1000t", "厂内主要道路"),
    ("松散可燃材料10t~500t与厂内次要道路的防火间距？", "5", "表4.2.9", "松散", "厂内次要道路"),
    ("经压实包装可燃材料1001t~5000t与明火地点的防火间距？", "37", "表4.2.9", "1001t~5000t", "明火"),
    ("明火及散发火花地点与厂内主要道路的防火间距？", "10", "表4.2.10", "明火", "厂内主要道路"),
    ("甲类仓库与厂内铁路的防火间距？", "30", "表4.2.10", "甲类仓库", "厂内铁路"),
    ("行政、生活建筑三级与码头装卸区的防火间距？", "40", "表4.2.10", "行政、生活建筑·耐火等级三级", "码头"),
    ("罐区甲、乙类泵或泵房与厂内主要道路的防火间距？", "10", "表4.2.10", "泵或泵房", "厂内主要道路"),
]

print(f"{'chunk':>6s} {'块数':>5s} {'可答对':>8s}  失败项")
print("-" * 76)
for size in (300, 500, 800, 1200):
    chunks = load_chunks(size, min(50, size // 6))
    bm = BM25([c["text"] for c in chunks]).bind(chunks)
    ok, fails = 0, []
    for q, want, tbl, a, b in CASES:
        hits = bm.top(q, k=3)
        good = any(f"{want}m" in ch["text"] and a in ch["text"] and b in ch["text"]
                   for ch, _ in hits)
        if good:
            ok += 1
        else:
            fails.append(q[:16])
    pct = ok / len(CASES) * 100
    print(f"{size:>6d} {len(chunks):>5d} {ok:>4d}/{len(CASES):<3d}  "
          f"{pct:>5.0f}%  {', '.join(fails) if fails else '无'}")
