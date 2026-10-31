# -*- coding: utf-8 -*-
"""
校验规范语料库切片是否完整可用。

用法：
    python validate_chunks.py <语料库.xlsx>
    python validate_chunks.py <语料库.xlsx> --sheet RAG切片 --corpus 语料库

检查项（任一失败即非零退出）：
    1. 切片 ID 唯一
    2. 每个子块的父条号都能在父块表里找到（孤儿检查）
    3. 父块的「列项数」与子块数量一致
    4. 每个切片全文首行含元数据头（用 ｜ 分隔）
    5. 条文号在语料库里存在（悬空父条检查）
    6. 表格块能被某个条文引用

只依赖 openpyxl。不修改输入文件。
"""
import argparse
import re
import sys
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:
    sys.exit("需要 openpyxl: pip install openpyxl")

HEADER_SEP = "｜"
RE_CLAUSE = re.compile(r"^\d+\.\d+\.\d+$")


def load(wb, sheet):
    if sheet not in wb.sheetnames:
        sys.exit(f"找不到页签：{sheet}（现有：{wb.sheetnames}）")
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit(f"页签 {sheet} 是空的")
    head = [str(x or "").strip() for x in rows[0]]
    return head, [dict(zip(head, r)) for r in rows[1:] if any(r)]


def col(row, name, default=""):
    """按前缀取字段值。真实表格的表头常带括号说明（如「切片全文（含元数据头，可直接嵌入）」）。"""
    if name in row:
        return row[name]
    for k, v in row.items():
        if k.startswith(name):
            return v
    return default


def has_col(head, name):
    return any(h == name or h.startswith(name) for h in head)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--sheet", default="RAG切片")
    ap.add_argument("--corpus", default="语料库")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.xlsx, read_only=True, data_only=True)
    errs, warns = [], []

    # ---- 语料库 ----
    chead, corpus = load(wb, a.corpus)
    need = ["条文号", "章", "节", "条文原文本体", "列项数"]
    miss = [c for c in need if not has_col(chead, c)]
    if miss:
        sys.exit(f"「{a.corpus}」缺字段：{miss}\n现有字段：{chead}")

    clause_ids = [str(col(r, "条文号")) for r in corpus]
    dup = [k for k, v in Counter(clause_ids).items() if v > 1]
    if dup:
        errs.append(f"语料库里条文号重复：{dup[:10]}")
    bad_no = [c for c in clause_ids if not RE_CLAUSE.match(c)]
    if bad_no:
        warns.append(f"条文号格式异常（非 X.Y.Z）：{bad_no[:10]}")

    # 章/节 结构
    no_jie = [col(r, "条文号") for r in corpus
              if str(col(r, "节") or "").strip() in ("", "0 节", "0")]
    if no_jie:
        errs.append(f"「节」字段为空或写成 0 节（应写'未设节'）：{no_jie[:10]}")

    for r in corpus:
        body = str(col(r, "条文原文本体") or "")
        if "表" in body and "|" in body:
            warn = "{} 正文里疑似内联了表格（表格应独立成块）".format(col(r, "条文号"))
            warns.append(warn)
            break

    # ---- 切片 ----
    shead, slices = load(wb, a.sheet)
    need = ["切片ID", "父条号", "切片层级", "切片全文"]
    miss = [c for c in need if not has_col(shead, c)]
    if miss:
        sys.exit(f"「{a.sheet}」缺字段：{miss}\n现有字段：{shead}")

    sid = [str(col(r, "切片ID")) for r in slices]
    dup = [k for k, v in Counter(sid).items() if v > 1]
    if dup:
        errs.append(f"切片ID重复：{dup[:10]}")

    parents = {}
    children = defaultdict(list)
    tables = []
    for r in slices:
        lv = str(col(r, "切片层级") or "")
        p = str(col(r, "父条号") or "")
        if "父块" in lv:
            parents[p] = r
        elif "表" in lv:
            tables.append(r)
            children[p].append(r)
        else:
            children[p].append(r)

    # 孤儿
    orphan = [p for p in children if p not in parents]
    if orphan:
        errs.append(f"子块找不到父块（孤儿）：{orphan[:10]}")

    # 悬空父条（父块不在语料库里）
    dangling = [p for p in parents if p not in set(clause_ids)]
    if dangling:
        errs.append(f"父条不在语料库中：{dangling[:10]}")

    # 列项数 vs 子块数（只比「款」，不含表格块）
    mismatch = []
    for p, pr in parents.items():
        n_ku = len([c for c in children.get(p, []) if "表" not in str(col(c, "切片层级"))])
        declared = col(pr, "列项数")
        if declared is None or str(declared).strip() == "":
            continue
        try:
            declared = int(float(declared))
        except (TypeError, ValueError):
            continue
        if declared != n_ku:
            mismatch.append(f"{p}: 声明 {declared} 款，实际 {n_ku} 个款切片")
    if mismatch:
        errs.append(f"款数不一致（{len(mismatch)} 处）：" + "; ".join(mismatch[:8]))

    # 元数据头
    no_head = [str(col(r, "切片ID")) for r in slices
               if HEADER_SEP not in str(col(r, "切片全文") or "").split("\n")[0]]
    if no_head:
        errs.append(f"切片全文首行缺元数据头：{no_head[:10]}")

    # 款号连续性：子块序号必须是 1..n，不能跳号/重号
    RE_KUAN = re.compile(r"·\s*(\d+)\s*[）)]\s*$")
    kuan_nums = defaultdict(list)
    for r in slices:
        lv = str(col(r, "切片层级") or "")
        if "父块" in lv or "表" in lv:
            continue
        m = RE_KUAN.search(lv)
        if m:
            kuan_nums[str(col(r, "父条号") or "")].append(int(m.group(1)))
    gap = []
    for p, nums in kuan_nums.items():
        s = sorted(nums)
        if s != list(range(1, len(s) + 1)):
            dup_n = [x for x in set(s) if s.count(x) > 1]
            miss = [x for x in range(1, len(s) + 1) if x not in s]
            gap.append(f"{p}: 序号{s}" + (f" 缺{miss}" if miss else "")
                       + (f" 重{dup_n}" if dup_n else ""))
    if gap:
        errs.append(f"款号不连续（{len(gap)} 处）：" + "; ".join(gap[:8]))

    # 表格块是否有条文引用
    if tables:
        refd = set()
        for r in corpus:
            if str(col(r, "关联表格") or "").strip():
                refd.add(str(col(r, "条文号")))
        lost = [str(col(t, "切片ID")) for t in tables if str(col(t, "父条号")) not in refd]
        if lost:
            warns.append(f"表格块的父条未在语料库「关联表格」里登记：{lost[:10]}")

    # ---- 报告 ----
    print(f"文件：{a.xlsx}")
    print(f"语料库：{len(corpus)} 条")
    print(f"切片：{len(slices)} = 父块 {len(parents)} + 子块 {len(slices) - len(parents)}"
          f"（其中表格块 {len(tables)}）")
    print(f"有款的父条：{len(children)}")
    print()

    for w in warns:
        print(f"[警告] {w}")
    if errs:
        for e in errs:
            print(f"[错误] {e}")
        print(f"\n❌ 未通过（{len(errs)} 项错误，{len(warns)} 项警告）")
        return 1
    print(f"✅ 全部通过（{len(warns)} 项警告）")
    return 0


if __name__ == "__main__":
    sys.exit(main())

