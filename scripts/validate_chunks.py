# -*- coding: utf-8 -*-
"""
校验规范语料库切片是否完整可用。

用法：
    python validate_chunks.py <语料库.xlsx>
    python validate_chunks.py <语料库.xlsx> --sheet RAG切片 --corpus 语料库
    python validate_chunks.py <语料库.xlsx> --allow-suffix "A,B" --kuan 列项明细

检查项（任一失败即非零退出）：
    1. 切片 ID 唯一
    2. 每个子块的父条号都能在父块表里找到（孤儿检查）
    3. 父块的「列项数」与子块数量一致
    4. 每个切片全文首行含元数据头（用 ｜ 分隔）
    5. 条文号在语料库里存在（悬空父条检查）
    6. 表格块能被某个条文引用
    7. 【新增】父块正文必须含该条全部款（铁律 1 的落地校验）
    8. 【新增】款正文洁净度：款不得以 图/表/注/附图/附表 开头（页级元素污染检查）
    9. 【新增】公式占位与「公式明细」一一对应
    9b.【v1.10.2】公式分级落地：高可信公式必须写入切片（带「整理式·非原文逐字」标记与表达式），
         中/低可信只留「待核·未引入表达式」占位且表达式不得出现在切片，切片不得残留裸占位
        （明细有「LaTeX」或「正文处理」列时才启用；逐项诊断由 formula_details.py 提供）

只依赖 openpyxl。不修改输入文件。
"""
import argparse
import os
import re
import sys
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:
    sys.exit("需要 openpyxl: pip install openpyxl")

HEADER_SEP = "｜"
RE_CLAUSE = re.compile(r"^\d+\.\d+\.\d+$")
RE_CLAUSE_SUFFIX = re.compile(r"^(\d+\.\d+\.\d+)([A-Z])$")
# 附录内的条款号（附录G 的 G.0.1）与「附加说明」不是 X.Y.Z 形式，属合法形态
RE_APPX_CLAUSE = re.compile(r"^[A-J]\.\d+\.\d+$")
NON_CLAUSE_LABELS = {"附加说明", "条文说明"}
# 款正文里不该出现的"行首"内容（页级元素应单列，不得折进款）
RE_PAGELEVEL_LINE = re.compile(r"^\s*(?:图\s*\d+\.\d+\.\d+|表\s*\d+\.\d+\.\d+|附图\s*[A-Z]|附表\s*[A-Z]|注\s*[：:])")
RE_FORMULA_PH = re.compile(r"〔公式\s*[（(]([^）)]+)[）)]\s*〕")



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
    ap.add_argument("--kuan", default="列项明细", help="款级页签名（无此页签则跳过款相关检查）")
    ap.add_argument("--allow-suffix", default="A,B",
                    help="允许的条号后缀（局部修订新增条，如 1.0.5A）。默认 A,B；传空串关闭")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()
    allow_suffix = {s.strip().upper() for s in (a.allow_suffix or "").split(",") if s.strip()}

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
    bad_no, suffix_hit = [], []
    for c in clause_ids:
        if RE_CLAUSE.match(c) or c.startswith("附录") or RE_APPX_CLAUSE.match(c) \
                or c in NON_CLAUSE_LABELS:
            continue
        m = RE_CLAUSE_SUFFIX.match(c)
        if m:
            if m.group(2).upper() in allow_suffix:
                suffix_hit.append(c)
                continue
            bad_no.append(c)
        else:
            bad_no.append(c)
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
        elif "表" in lv or "附录" in lv:
            # 附录子块（附录A（子块·聚酯）等）与表格块同类：独立信息单元，不算"款"
            tables.append(r)
            children[p].append(r)
        else:
            children[p].append(r)

    # 孤儿
    orphan = [p for p in children if p not in parents]
    if orphan:
        errs.append(f"子块找不到父块（孤儿）：{orphan[:10]}")

    # 悬空父条（父块不在语料库里）
    #   附录A/附录B 在语料库页签里有对应行（章=附录），属于合法父条。
    clause_ids = set(clause_ids) | {"附录A", "附录B"}
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
        # 附录子块不来自「关联表格」列，属正常，按父条号排除
        lost = [str(col(t, "切片ID")) for t in tables
                if str(col(t, "父条号")) not in refd
                and not str(col(t, "父条号")).startswith("附录")]
        if lost:
            warns.append(f"表格块的父条未在语料库「关联表格」里登记：{lost[:10]}")

    # ---- 7/8. 款相关：款正文洁净度 + 父块含全部款（铁律 1）
    kuan_rows = []
    if a.kuan and a.kuan in wb.sheetnames:
        khead, kuan_rows = load(wb, a.kuan)
    dirty, lost_kuan = [], []
    if kuan_rows:
        if not has_col(khead, "条号") or not has_col(khead, "该层正文"):
            warns.append(f"「{a.kuan}」缺「条号」或「该层正文」列，跳过款检查")
        else:
            body_by_no = {}
            for r in corpus:
                no = str(col(r, "条文号"))
                body_by_no[no] = re.sub(r"\s+", "", str(col(r, "条文原文本体") or ""))
            for k in kuan_rows:
                no = str(col(k, "条号"))
                kno = str(col(k, "列项路径") or col(k, "款号") or "")
                ktext = str(col(k, "该层正文") or "")
                # 8. 款正文不得以页级元素开头（图题/表题/注）
                for ln in ktext.split("\n"):
                    if RE_PAGELEVEL_LINE.match(ln):
                        dirty.append(f"{no} 款{kno}: {ln.strip()[:26]}")
                        break
                # 7. 父块正文必须含该款（款号 + 内容片段）
                body = body_by_no.get(no)
                if body is None:
                    lost_kuan.append(f"{no}: 语料库里没有这一条")
                    continue
                if kno and re.sub(r"\s+", "", kno) not in body:
                    lost_kuan.append(f"{no}: 父块正文缺款号 {kno}")
                elif ktext and re.sub(r"\s+", "", ktext)[:12] not in body:
                    lost_kuan.append(f"{no}: 父块正文缺款 {kno} 的正文")
    if dirty:
        errs.append(f"款正文混入页级元素（图题/表题/表注应单列，不得折进款）{len(dirty)} 处："
                    + "; ".join(dirty[:6]))
    if lost_kuan:
        errs.append(f"父块正文未含全部款（铁律 1 失败）{len(lost_kuan)} 处："
                    + "; ".join(lost_kuan[:6]))

    # ---- 9. 公式占位 ↔「公式明细」一一对应
    n_ph = n_detail = 0
    if "公式明细" in wb.sheetnames:
        fhead, frows = load(wb, "公式明细")
        if has_col(fhead, "公式编号"):
            def norm(s):
                return re.sub(r"[\s（）()]", "", str(s or ""))
            detail = Counter(norm(col(f, "公式编号")) for f in frows if norm(col(f, "公式编号")))
            ph = Counter()
            for r in corpus:
                for m in RE_FORMULA_PH.finditer(str(col(r, "条文原文本体") or "")):
                    ph[norm(m.group(1))] += 1
            n_ph, n_detail = sum(ph.values()), sum(detail.values())
            orphan_ph = [k for k in ph if k not in detail]
            if orphan_ph:
                errs.append(f"正文有公式占位但「公式明细」里没有对应行：{orphan_ph[:8]}")
            only_detail = [k for k in detail if k not in ph]
            if only_detail:
                warns.append(f"「公式明细」有登记但正文无占位（可能公式在表格/附录里，需确认）：{only_detail[:8]}")
            if ph != detail:
                warns.append(f"公式占位与明细条数不一致：占位 {n_ph} / 明细 {n_detail}")

    # ---- 9b. v1.10.2 公式分级落地（高可信写入切片 / 中低可信只留待核 / 无裸占位）
    f_info = {}
    if "公式明细" in wb.sheetnames:
        fhead2, _frows2 = load(wb, "公式明细")
        if has_col(fhead2, "正文处理") or has_col(fhead2, "LaTeX"):
            try:
                import formula_details as fd
            except ImportError:
                sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                import formula_details as fd
            ft = fd.FormulaTable.from_xlsx(a.xlsx)
            f_errs, f_warns, f_info = ft.check(a.xlsx)   # 切片级校验（md 级由 md 体检/导出闸门负责）
            errs.extend(f_errs)
            warns.extend(f_warns)

    # ---- 报告 ----
    print(f"文件：{a.xlsx}")
    print(f"语料库：{len(corpus)} 条" +
          (f"（其中 A/B 后缀新增条 {len(suffix_hit)} 个，已按 --allow-suffix 放行）" if suffix_hit else ""))
    print(f"切片：{len(slices)} = 父块 {len(parents)} + 子块 {len(slices) - len(parents)}"
          f"（其中表格块 {len(tables)}）")
    print(f"有款的父条：{len(children)}")
    if kuan_rows:
        print(f"款检查：{len(kuan_rows)} 款｜款正文污染 {len(dirty)} 处｜父块缺款 {len(lost_kuan)} 处")
    if n_ph or n_detail:
        print(f"公式占位：正文 {n_ph} 个 ｜「公式明细」{n_detail} 行")
    if f_info:
        print(f"公式分级：共 {f_info.get('total')} 条（写入正文 {f_info.get('write')} / "
              f"待核占位 {f_info.get('pending')}）｜切片 {f_info.get('slices')} 个已校验落地")
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

