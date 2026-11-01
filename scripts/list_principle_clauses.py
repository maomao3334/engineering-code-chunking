# -*- coding: utf-8 -*-
"""列出「原则/引述条款」清单——检索降权规则的依据。

判定规则（两条同时满足）：
  1. 正文命中原则句式：相适应 / 尚应符合 / 依照…规定 / 应符合本规范
  2. 几乎没有具体数值（数字字符 ≤ 1 个）

为什么这类条款要降权：它们永远不可能是具体问题的答案。
  问"最多几层"，答案是表里的数字，不可能是"应与层数相适应"。

实测（GB 55037-2022 + 高层民用建筑消防安全管理规定，292 条）：
  命中 10 条，占 3.4%。规律：GB 每章第一节第一条几乎都是"适应性"条款
  （2.1.1 / 5.1.1 / 7.1.1 / 8.1.1），另外还有法规里的责任/费用划分条。

用法：
    python list_principle_clauses.py --rag <含 raglib.py 的目录>
"""
import argparse
import re
import sys
from pathlib import Path

PRINCIPLE_RE = re.compile(r"(相适应|尚应符合|依照.{0,8}规定|应符合本规范)")
NUM_RE = re.compile(r"\d")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rag", required=True, help="含 raglib.py 的目录")
    ap.add_argument("--pattern", default=None, help="自定义原则句式正则（可选）")
    a = ap.parse_args()

    sys.path.insert(0, str(Path(a.rag).resolve()))
    from raglib import Retriever  # noqa: E402

    pre = re.compile(a.pattern) if a.pattern else PRINCIPLE_RE
    r = Retriever()

    rows = []
    for k, v in r.parents.items():
        b = v["body"]
        if pre.search(b) and len(NUM_RE.findall(b)) <= 1:
            rows.append((k, v["title"], len(b), b[:70].replace("\n", " ")))

    print(f"共 {len(rows)} 条 / 总 {len(r.parents)} 条"
          f"（占 {len(rows) / max(len(r.parents), 1) * 100:.1f}%）\n")
    for k, t, n, s in rows:
        print(f"{k:10s} {t[:26]:28s} {n:4d}字  {s}")

    print("\n建议：这一类在检索时降权（实测 ×0.35 或 ×0.5 效果相同，取 0.35）。")
    print("不要降到 0.1 以下，也不要删除——用户真问\"总则怎么规定\"时还要能检索到。")


if __name__ == "__main__":
    main()
