# -*- coding: utf-8 -*-
"""压力测试：找 md 作为知识库的真实缺陷。

上一版 9/9 太宽松（只查"某块同时含数字+两词"）。这版测真正的风险点：

  T1 切块断裂：同一行的事实句被切到两块，两边都不完整
  T2 上下文丢失：块里只有事实句、没有表名/单位，答出来不知道是哪张表
  T3 易混淆：同一数值出现在多个对象，检索能区分吗
  T4 小 chunk_size：平台默认可能 500 字，更碎
  T5 只给文件名不含章节号时（有的平台索引文件名）
"""
import re
import sys
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

    def scores(self, q):
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
        return out


print("=" * 80)
print("T1/T2 切块断裂与上下文丢失（chunk=800）")
print("=" * 80)
chunks = load_chunks(800, 100)
bm = BM25([c["text"] for c in chunks])

# 检查每个块的头部是否带表名（上下文）
no_ctx = []
for c in chunks:
    t = c["text"]
    if "防火间距不应小于" not in t:
        continue
    # 该块是否含表名或表号
    if not re.search(r"表\s*4\.\d+\.\d+|纺织工业工厂总平面布置|可燃材料堆场|厂外建筑物或其他设施", t):
        no_ctx.append(c)
print(f"含事实句但**不含任何表名/表号**的块: {len(no_ctx)} 个")
for c in no_ctx[:5]:
    print(f"   {c['file'][:24]:26s} 开头: {c['text'][:60]!r}")

# 检查事实句是否被切断
broken = 0
samples = []
for c in chunks:
    t = c["text"]
    # 半截句：以「与「」开头或以「的防火间距」结尾但没数值
    if re.search(r"^[^。\n]*的防火间距$", t.strip()[-30:]):
        broken += 1
        samples.append((c["file"], t[-60:]))
print(f"\n以「的防火间距」结尾（可能被切断）的块: {broken} 个")
for f, s in samples[:5]:
    print(f"   {f[:24]:26s} 尾部: {s!r}")

print()
print("=" * 80)
print("T3 易混淆：同一数值对应多个对象，检索能否区分")
print("=" * 80)
# 4.1.7 里 35 出现多次；4.2.10 里 30 出现极多次
CONFUSE = [
    ("厂外铁路与甲、乙类可燃液体储罐的防火间距？", "35", "厂外铁路", "可燃液体储罐"),
    ("厂外铁路与甲、乙类厂房的防火间距？", "30", "厂外铁路", "甲、乙类厂房"),
    ("明火及散发火花地点与甲类仓库的防火间距？", "30", "明火", "甲类仓库"),
    ("明火及散发火花地点与厂内铁路的防火间距？", "20", "明火", "厂内铁路"),
]
for q, want, a, b in CONFUSE:
    sc = bm.scores(q)
    order = sorted(range(len(sc)), key=lambda i: -sc[i])[:3]
    hit = None
    for i in order:
        t = chunks[i]["text"]
        if f"{want}m" in t and a in t and b in t:
            hit = chunks[i]["file"]
            break
    print(f"  {'✓' if hit else '✗'} {q[:36]:38s} 期望{want}m  "
          f"{'命中 ' + hit[:18] if hit else '未命中'}")

print()
print("=" * 80)
print("T4 更小的 chunk_size（平台可能 500 字）")
print("=" * 80)
for size in (500, 300):
    ch = load_chunks(size, 50)
    b2 = BM25([c["text"] for c in ch])
    n_no_ctx = sum(1 for c in ch
                   if "防火间距不应小于" in c["text"]
                   and not re.search(r"表\s*4\.\d+\.\d+|纺织工业工厂|可燃材料堆场|厂外建筑物", c["text"]))
    print(f"  chunk={size}: {len(ch)} 块，其中 {n_no_ctx} 块含事实句但无表名"
          f"（{n_no_ctx / max(len(ch),1) * 100:.0f}%）")
