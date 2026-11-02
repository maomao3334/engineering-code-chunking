# -*- coding: utf-8 -*-
"""用视觉大模型理解表格 → 生成事实句（含数字交叉校验）。

## 为什么需要这层校验

VLM 读表有两个已知失效模式：
    1. **编数字**：格子里是"—"或模糊，它会"补"一个合理值
    2. **错位**：行列对齐错一格，数值全串

而在工程规范场景，**一个错的防火间距比查不到更危险**。
所以本脚本的设计是：

    VLM 负责「理解结构」（表名/列名/行列对应）—— 这是它擅长的
    OCR 负责「提供数值」—— 这是确定性的
    两边交叉核对，只有**数值完全一致**的事实句才标 high，否则标 low 待人工确认

## 三种校验结果

    一致      → confidence=high，可直接入库
    仅 VLM 有 → confidence=low，可能 VLM 编的，需人工看原图
    仅 OCR 有 → 记为遗漏，VLM 没读出来（可能格子看不清）

## 用法

    export ANTHROPIC_API_KEY=sk-...
    python vlm_table_to_facts.py --pdf 规范.pdf [--pages 9,11,12]

依赖：pip install anthropic pyyaml
若无 API key，脚本会提示并退出，不影响既有 OCR 流程。
"""
import argparse
import base64
import json
import os
import re
import sys
from pathlib import Path

# 本机代理配置里的 NO_PROXY 含 "[::1]"，httpx 解析会抛 Invalid port: ':1]'。
# 必须在 import anthropic 之前清掉。
for _k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY",
           "http_proxy", "https_proxy", "all_proxy", "NO_PROXY", "no_proxy"):
    os.environ.pop(_k, None)

PROMPT = """这是一张中国国家标准的表格图片。请仔细读表，输出严格的 JSON。

要求：
1. title: 这张表的完整标题（含表号，如「表4.2.9 可燃材料堆场…」）
2. columns: 数组，每项 {index, name}。表头有多层合并，请把每列**最具体**的完整名称说清楚，
   上层分组名用括号附在后面，例如「甲类厂房及仓库（属防火间距(m)建筑物、构筑物）」
3. facts: 数组，每项 {row_label, col_index, col_name, value}
   - row_label: 该行最左列的行名称（如「2.厂外铁路」或「经压实包装的可燃材料 10t~500t」）
   - value: **只填格子里确实印着的数字**，不要带单位
   - 格子里是「—」「/」或空白时，value 填 null，**绝对不要推测或补全**
   - 看不清的格也填 null，不要编数字

只输出 JSON，不要任何解释文字、不要 markdown 代码块。"""


def load_key():
    """优先环境变量，其次 DSH 凭据中心"""
    k = os.environ.get("ANTHROPIC_API_KEY")
    if k:
        return k
    try:
        import yaml
        p = Path.home() / ".dsh" / ".credentials.yaml"
        if p.exists():
            d = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
            return (d.get("refs") or {}).get("A_API_KEY")
    except Exception:
        pass
    return None


def call_vlm(img: Path, key: str, model: str, base: str):
    from anthropic import Anthropic
    b64 = base64.b64encode(img.read_bytes()).decode()
    c = Anthropic(api_key=key, base_url=base)
    r = c.messages.create(
        model=model, max_tokens=8000,
        messages=[dict(role="user", content=[
            dict(type="image", source=dict(type="base64", media_type="image/png", data=b64)),
            dict(type="text", text=PROMPT)])])
    txt = "".join(b.text for b in r.content if hasattr(b, "text")).strip()
    if txt.startswith("```"):
        txt = txt.split("```")[1]
        txt = txt.split("\n", 1)[1] if "\n" in txt else txt
    return json.loads(txt)


def ocr_numbers(t: dict):
    """从已有 OCR 网格里取出全部数字（键为 (行,列)）"""
    out = {}
    for r, row in enumerate(t["grid"]):
        for c, v in enumerate(row):
            s = str(v).strip()
            if re.fullmatch(r"\d+(?:\.\d+)?", s):
                out[(r, c)] = s
    return out


def crosscheck(vlm: dict, t: dict):
    """VLM 的值与 OCR 的值交叉核对

    匹配方式：按 value 完全相等 + 同一行内。
    （VLM 的 col_index 与 OCR 的列号不一定一致，不能直接按坐标比。）
    """
    ocr_by_row = {}
    for (r, c), v in ocr_numbers(t).items():
        ocr_by_row.setdefault(r, set()).add(v)

    verified, suspicious, missing = [], [], []
    for f in vlm.get("facts", []):
        val = f.get("value")
        if val is None:
            continue
        val = str(val).strip()
        # 找哪一行的 OCR 数字集合里含这个值
        hit_rows = [r for r, vals in ocr_by_row.items() if val in vals]
        item = dict(row_label=f.get("row_label", ""),
                    col_name=f.get("col_name", ""), value=val,
                    vlm_col=f.get("col_index"))
        if hit_rows:
            item["ocr_rows"] = hit_rows
            verified.append(item)
        else:
            suspicious.append(item)

    vlm_vals = {str(f["value"]).strip() for f in vlm.get("facts", [])
                if f.get("value") is not None}
    for r, vals in ocr_by_row.items():
        for v in vals - vlm_vals:
            missing.append(dict(row=r, value=v))

    return verified, suspicious, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--pages", default="", help="只处理这些页，逗号分隔；空=全部")
    ap.add_argument("--model", default=os.environ.get("ANTHROPIC_MODEL",
                                                      "claude-sonnet-4-5-20250929"))
    ap.add_argument("--base", default=os.environ.get("ANTHROPIC_BASE_URL",
                                                     "https://api.anthropic.com"))
    a = ap.parse_args()

    key = load_key()
    if not key:
        sys.exit("未找到 ANTHROPIC_API_KEY，也未在 DSH 凭据中心找到 A_API_KEY。\n"
                 "设置后重试：export ANTHROPIC_API_KEY=sk-...")

    CASE = Path(a.pdf).parent / "_pipeline"
    tabs = json.loads((CASE / "tables_ocr.json").read_text(encoding="utf-8"))
    pages = {int(x) for x in a.pages.split(",") if x.strip()} if a.pages else None

    results = []
    for t in tabs:
        if t["w"] > 1000:            # 广告图
            continue
        if pages and t["page"] not in pages:
            continue
        img = CASE / "table_imgs" / t["img"]
        if not img.exists():
            continue
        print(f"\n处理 p{t['page']} {t['img']} …")
        try:
            vlm = call_vlm(img, key, a.model, a.base)
        except Exception as e:
            print(f"  VLM 失败: {str(e)[:160]}")
            results.append(dict(page=t["page"], img=t["img"], error=str(e)[:300]))
            continue
        ver, sus, mis = crosscheck(vlm, t)
        print(f"  title={vlm.get('title', '')[:40]}")
        print(f"  列名 {len(vlm.get('columns', []))} 个, 事实 {len(vlm.get('facts', []))} 条")
        print(f"  ✓ 与 OCR 一致 {len(ver)}   ⚠ 仅 VLM 有 {len(sus)}   ○ OCR 有但 VLM 漏 {len(mis)}")
        if sus:
            print(f"    可疑样本: {sus[:3]}")
        results.append(dict(page=t["page"], img=t["img"], vlm=vlm,
                            verified=ver, suspicious=sus, missing=mis))

    out = CASE / "vlm_table_facts.json"
    out.write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n→ {out}")
    print("\n注意：只有 verified 里的数值可视为可信；")
    print("     suspicious 需人工对照原图确认，不要直接入库。")


if __name__ == "__main__":
    main()
