# -*- coding: utf-8 -*-
"""诊断"为什么没命中期望条文"——把排名、分数、token 交集打出来。

不要靠猜排名。每个"没命中"的题都跑一遍这个，三种输出对应三种修法：

  · 期望排第 2~3、分数与第 1 差 <5%      → 被"近义条款"压住（多为原则条款）
  · 期望排 5 名外、问句∩期望 非空         → 排序权重问题，调权重/加同义词
  · 期望排 5 名外、问句∩期望 为空         → 词法检索天花板，只能靠同义词/别名

用法：
    python diagnose_misses.py --rag <rag目录> --q "问题" --want 7.4.1
    python diagnose_misses.py --rag . --cases cases.json     # 批量

cases.json 格式：
    [{"q": "我这几间铺子只有一部楼梯，最多能做几层？", "want": "7.4.1"}, ...]
"""
import argparse
import json
import sys
from pathlib import Path


def load_rag(rag_dir):
    sys.path.insert(0, str(rag_dir))
    import raglib
    return raglib


def diagnose(R, raglib, q, want, k=5):
    hits = R.search_parent(q, k=max(k, 10))
    rank = next((i + 1 for i, h in enumerate(hits) if h["parent"] == want), None)

    print("=" * 74)
    print(f"问：{q}")
    print(f"  扩展词：{raglib.expand_query(q)}")
    for i, h in enumerate(hits[:k], 1):
        mark = "  <<< 期望" if h["parent"] == want else ""
        t = h["slice"]["title"]
        print(f"  {i}. {h['score']:7.2f}  {h['parent']:10s} {t[:26]}{mark}")
    print(f"  期望 {want} 排名：{rank}")

    p = R.parents.get(want)
    if not p:
        print(f"  !! {want} 不在语料库里（先确认它切进来了）")
        return dict(q=q, want=want, rank=rank, hits=[h["parent"] for h in hits[:k]])

    qs = set(raglib.tokenize(q))
    ps = set(raglib.tokenize(p["body"] + p["header"]))
    inter = sorted(qs & ps)
    print(f"  问句∩期望：{inter or '【空——词法检索无解，只能靠同义词】'}")
    print(f"  期望标题：{p['title']}   正文 {len(p['body'])} 字")

    if rank and rank > 1:
        first = hits[0]["parent"]
        pf = R.parents.get(first)
        if pf:
            fs = set(raglib.tokenize(pf["body"] + pf["header"]))
            print(f"  第1名 {first}（{pf['title'][:20]}）问句∩它：{sorted(qs & fs)}")
            want_score = next(h["score"] for h in hits if h["parent"] == want)
            gap = (hits[0]["score"] - want_score) / max(hits[0]["score"], 1e-6)
            print(f"  分差：{gap * 100:.1f}%"
                  + ("   ← 差距很小，多半是被近义/原则条款压住" if gap < 0.05 else ""))

    return dict(q=q, want=want, rank=rank, overlap=inter,
                hits=[h["parent"] for h in hits[:k]])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rag", required=True, help="含 raglib.py 的目录")
    ap.add_argument("--q", help="单个问题")
    ap.add_argument("--want", help="期望条文号")
    ap.add_argument("--cases", help="批量：JSON 文件 [{q,want}, ...]")
    ap.add_argument("-k", type=int, default=5)
    a = ap.parse_args()

    raglib = load_rag(Path(a.rag).resolve())
    R = raglib.Retriever(use_table_slices=True, use_expand=True, child_weight=0.8)

    if a.cases:
        cases = json.loads(Path(a.cases).read_text(encoding="utf-8"))
        for c in cases:
            diagnose(R, raglib, c["q"], c["want"], a.k)
    elif a.q and a.want:
        diagnose(R, raglib, a.q, a.want, a.k)
    else:
        ap.error("给 --q/--want，或 --cases cases.json")


if __name__ == "__main__":
    main()
