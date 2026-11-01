# -*- coding: utf-8 -*-
"""把语料库导出成「任何 RAG 平台都能直接吃」的格式

核心思路（很关键）：
    我们的切片中位 175 字、最长 912 字。而主流 RAG 平台的默认分段长度
    是 500~1000 字。所以——**只要一个文件装一个切片，平台怎么切都是它自己**。
    不需要 API 导入，不需要平台支持父子结构，上传文件夹就行。

输出（三种，按需用）：

    markdown/           241 个 .md 文件，一个文件 = 一条完整条文（含全部款 + 关联表格）
                        → 适合：RAGFlow / MaxKB / AnythingLLM / FastGPT / Dify / 任何平台的"上传文件"
                        → 文件名带条号，排序、检索、溯源都方便

    分段文本/全库.md     单文件，用 `## 第X条` 分隔
                        → 适合：支持"按标题分段"的平台（RAGFlow 的 Title Chunker 等）

    json/slices.json    结构化 JSON（父块 + 子块 + 元数据）
                        → 适合：FastGPT pushData / Dify segments / 自己写脚本

用法：
    python export\\export_for_platform.py
    python export\\export_for_platform.py --out D:\\我的导出
"""
import argparse
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).parent
ROOT = HERE.parent
# raglib.py 与本脚本同目录（skill 自带），优先这里；也兼容项目里 rag/ 的布局
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "rag"))


def safe_name(s):
    """文件名里不能有的字符换成下划线"""
    return re.sub(r'[\\/:*?"<>|\s]+', "_", s).strip("_")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=str(HERE / "导出"),
                    help="输出目录（默认 export/导出）")
    ap.add_argument("--corpus", default=None, help="指定语料库 xlsx")
    a = ap.parse_args()

    if a.corpus:
        import os
        os.environ["CORPUS_XLSX"] = a.corpus

    from raglib import Retriever
    r = Retriever(use_table_slices=True, use_expand=True, child_weight=0.8)

    out = Path(a.out)
    md_dir = out / "markdown"
    txt_dir = out / "分段文本"
    json_dir = out / "json"
    for d in (md_dir, txt_dir, json_dir):
        d.mkdir(parents=True, exist_ok=True)

    # ---------- 1. 一条条文一个 .md 文件 ----------
    single = []
    for no, p in sorted(r.parents.items(), key=lambda kv: _sortkey(kv[0])):
        text = r.render_expanded(no)          # 含所有款 + 关联表 + 被引用条
        title = p["title"] or ""
        fname = f"{no} {safe_name(title)}.md"
        (md_dir / fname).write_text(text, encoding="utf-8")

        single.append({
            "条号": no,
            "标题": title,
            "章": p["header"].split("｜")[1] if "｜" in p["header"] else "",
            "节": p["header"].split("｜")[2] if p["header"].count("｜") >= 2 else "",
            "款数": len([c for c in r.children.get(no, []) if "表" not in c["level"]]),
            "有表格": bool(r.table_md(no)),
            "字符数": len(text),
            "文件名": fname,
        })

    # ---------- 2. 单文件，按二级标题分段 ----------
    lines = [
        "# 建筑防火通用规范 GB 55037-2022",
        "",
        "> 依据 GB 55037-2022 整理的条文库。全部条文为强制性条文。",
        "> 共 241 条，按「章—节—条」组织。每条下含全部款与关联表格。",
        "",
    ]
    cur_ch = None
    for no, p in sorted(r.parents.items(), key=lambda kv: _sortkey(kv[0])):
        ch = p["header"].split("｜")[1] if "｜" in p["header"] else ""
        if ch != cur_ch:
            lines += [f"# {ch}", ""]
            cur_ch = ch
        jie = p["header"].split("｜")[2] if p["header"].count("｜") >= 2 else ""
        lines += [f"## 第{no}条 {p['title']}", "", f"*{ch}｜{jie}*", ""]
        lines += [p["body"], ""]
        tmd = r.table_md(no)
        if tmd:
            lines += [tmd, ""]
        refs = r.refs_of(no)
        if refs:
            lines += ["**关联条文：**" + "、".join(f"第{x}条" for x in refs), ""]
        lines += ["---", ""]
    (txt_dir / "全库.md").write_text("\n".join(lines), encoding="utf-8")

    # ---------- 3. JSON ----------
    slices = []
    for no, p in sorted(r.parents.items(), key=lambda kv: _sortkey(kv[0])):
        slices.append({
            "id": f"GB55037-2022::{no}",
            "parent": no,
            "level": "条",
            "title": p["title"],
            "text": r.render_expanded(no),
            "metadata": {
                "规范编号": "GB 55037-2022",
                "章": p["header"].split("｜")[1] if "｜" in p["header"] else "",
                "节": p["header"].split("｜")[2] if p["header"].count("｜") >= 2 else "",
                "条号": no,
                "小标题": p["title"],
                "是否强条": "是",
            },
        })
    for no, kids in r.children.items():
        for c in kids:
            if "表" in c["level"]:
                continue
            slices.append({
                "id": c["id"], "parent": no, "level": "款",
                "title": c["title"], "text": c["text"],
                "metadata": {"规范编号": "GB 55037-2022", "条号": no,
                             "层级": c["level"]},
            })
    (json_dir / "slices.json").write_text(
        json.dumps(slices, ensure_ascii=False, indent=1), encoding="utf-8")

    # FastGPT pushData 格式（q=父块全文，indexes=子块）
    fastgpt = []
    for no, p in sorted(r.parents.items(), key=lambda kv: _sortkey(kv[0])):
        kids = [c for c in r.children.get(no, []) if "表" not in c["level"]]
        fastgpt.append({
            "q": r.render_expanded(no),
            "indexes": [{"text": c["text"]} for c in kids] or
                       [{"text": r.render_expanded(no)}],
            "metadata": {"条号": no, "标题": p["title"]},
        })
    (json_dir / "fastgpt_pushdata.json").write_text(
        json.dumps(fastgpt, ensure_ascii=False, indent=1), encoding="utf-8")

    # ---------- 说明 ----------
    readme = f"""# 语料库导出包

由 `export/export_for_platform.py` 生成，源语料库：`{Path(r.__class__.__module__).name if False else 'v5'}`

## 目录

| 目录 | 内容 | 数量 |
|---|---|---|
| `markdown/` | 一条条文一个 .md 文件 | {len(single)} 个 |
| `分段文本/全库.md` | 单文件，`## 第X条` 分隔 | 1 个 |
| `json/slices.json` | 结构化切片（父块+子块+元数据） | {len(slices)} 条 |
| `json/fastgpt_pushdata.json` | FastGPT pushData 格式 | {len(fastgpt)} 条 |

## 怎么用

### 方式 A：上传 markdown 文件夹（最通用）

几乎所有 RAG 平台都支持上传 .md 或 .txt 文件。把 `markdown/` 目录下的
{len(single)} 个文件全部上传即可。

**为什么这样能保持切片完整**：每个文件只有约 175 字（最长 912 字），
小于平台默认的分段长度（通常 500~1000 字）。所以平台分段之后，
**一个文件恰好等于一个 chunk**，我们的父子结构被完整保留。

适用：RAGFlow、MaxKB、AnythingLLM、FastGPT、Dify、Cherry Studio、WeKnora 等，
以及任何支持"上传文件建知识库"的工具。

### 方式 B：单文件

上传 `分段文本/全库.md`。文件里用 `## 第X条` 作为分段标记，
支持"按标题分段 / Title Chunker"的平台会直接按条切开。

### 方式 C：JSON（需要 API 导入）

- `json/slices.json`：通用格式，自己写脚本灌向量库或任意平台
- `json/fastgpt_pushdata.json`：直接 POST 到 FastGPT 的
  `/core/dataset/data/pushData`（每次最多 200 条，分 {max(1, (len(fastgpt)+199)//200)} 次发）

## 内容说明

每个文件的结构：

```
《建筑防火通用规范》GB 55037-2022｜7 安全疏散与避难设施｜7.4 公共建筑｜第7.4.1条｜公共建筑安全出口与疏散楼梯｜强制性条文
（条文正文，含全部款）
【表7.4.1 …】
（表格，Markdown 格式）
```

第一行是元数据头，**这一行是溯源的关键** —— 模型读到它才知道自己在引用哪一条。

## 校验

导出后可用于核对：条号总数应为 241，带表格的应为 4 条。
"""
    (out / "README.md").write_text(readme, encoding="utf-8")

    # ---------- 报告 ----------
    print(f"输出目录：{out}\n")
    print(f"  markdown/           {len(single)} 个文件")
    print(f"  分段文本/全库.md     {(txt_dir / '全库.md').stat().st_size:,} 字节")
    print(f"  json/slices.json    {len(slices)} 条切片")
    print(f"  json/fastgpt_...    {len(fastgpt)} 条（需分 {(len(fastgpt)+199)//200} 次推送）")
    print()
    sizes = [s["字符数"] for s in single]
    print(f"单个文件字符数：最短 {min(sizes)}，中位 {sorted(sizes)[len(sizes)//2]}，最长 {max(sizes)}")
    print(f"  → 中位 {sorted(sizes)[len(sizes)//2]} 字，远小于平台默认分段长度（500~1000），")
    print(f"    所以上传后一个文件就是一个 chunk，不会被切碎。")
    print()
    n_table = sum(1 for s in single if s["有表格"])
    print(f"带表格的条文：{n_table} 条（表格已渲染成 Markdown 内联进正文）")


def _sortkey(no):
    return tuple(int(x) for x in no.split("."))


if __name__ == "__main__":
    main()
