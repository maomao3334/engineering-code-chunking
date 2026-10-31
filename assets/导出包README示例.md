# 语料库导出包

由 `export/export_for_platform.py` 生成，源语料库：`v5`

## 目录

| 目录 | 内容 | 数量 |
|---|---|---|
| `markdown/` | 一条条文一个 .md 文件 | 241 个 |
| `分段文本/全库.md` | 单文件，`## 第X条` 分隔 | 1 个 |
| `json/slices.json` | 结构化切片（父块+子块+元数据） | 761 条 |
| `json/fastgpt_pushdata.json` | FastGPT pushData 格式 | 241 条 |

## 怎么用

### 方式 A：上传 markdown 文件夹（最通用）

几乎所有 RAG 平台都支持上传 .md 或 .txt 文件。把 `markdown/` 目录下的
241 个文件全部上传即可。

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
  `/core/dataset/data/pushData`（每次最多 200 条，分 2 次发）

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
