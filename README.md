# engineering-code-chunking

把**中国工程建设标准**（GB/GBJ/JGJ 等）的 PDF / OCR 文本，切成**可溯源到条文号**的层级语料库
（章 → 节 → 条 → 款 + 表格），并生成父子块 RAG 切片。

> 一句话原理：**条文是父块，款是子块。检索命中任一款，就把整条连同全部款、章节头和表格一起交出去。**

这是给 AI Agent（DeepSeek Harness / Claude Skills 格式）用的 skill 包，也可以当普通方法论 + 脚本库直接读。

## 为什么需要它

通用切分库（Chonkie / LlamaIndex text splitter / 平台自带分块）都是把文本切成一堆**平级 chunk**。
而规范里的 `1 2 3 4` 是 `2.1.2` 的**子款**——一旦切成平级，问"建筑防火的目标有哪几个"，
只捞到一款就答，答案必然漏。

实测（同一份原文、同一检索器、同一组题，只换切法）：

| 切分方式 | 切片数 | 命中@3 | **全款交付率** |
|---|---|---|---|
| 正则平切 | 644 | 18/20 | **0%** |
| 固定长度切 | 299 | 19/20 | 79% |
| 只切到条 | 241 | 19/20 | 100% |
| **父子块（本方法）** | 760 | 18/20 | **100%** |

**关键结论：命中率证明不了切法好坏（四种都在 18~19/20），全款交付率才是分水岭。
而向量化修不了漏款——那是结构问题，不是检索算法问题。**

## 三条铁律

1. **数字层级不是兄弟，是父子**：`1 2 3 4` 属于 `2.1.2`，不是和它平级的四条
2. **表格必须独立成块并挂回条文**：正文只写"应符合表7.4.1的规定"，表体不进索引就永远搜不到数字
3. **跨条引用要闭合**：`第4.1.4条` 这类引用不带上被引条文，答案是悬空的

## 目录

```
SKILL.md                                  主文档（给 agent 读的入口）
references/
  pipeline.md                             OCR → 解析 → 出表的流水线
  chunk-schema.md                         语料库表结构（8 个页签、列定义）
  edge-cases.md                           边界情况（跨页条、款号缺失、表格跨页…）
  prompt-engineering.md                   切分提示词工程（为批量切 200 部准备）
  quality-gates.md                        质检门（6 类结构检查）
  evaluation.md                           检索评测方法
  export-and-deploy.md                    导出到平台 / 免安装打包
  multi-standard-and-acceptance.md       并第二部规范 + 20 题验收自动跑分
  retrieval-tuning.md                     检索调参实录：什么有效、什么试过没用
  acceptance-design.md                    验收题设计：期望值先核实
scripts/
  validate_chunks.py                      结构校验
  scan_defects.py                         12 类 OCR 字符缺陷扫描
  fix_defects.py                          自动修复
  compare_chunking.py                     四种切分策略对比实验
  export_for_platform.py                  导出 markdown/json/分段文本
  diagnose_misses.py                      诊断"为什么没命中"（排名/分数/token 交集）
  list_principle_clauses.py               列出原则/引述条款（检索降权依据）
assets/                                   示例产出（评测集、缺陷报告、导出说明）
```

## 用法

**作为 skill**：把整个目录放到 `~/.agents/skills/`（DSH）或对应 skills 目录，agent 会自动加载 `SKILL.md`。

**当脚本用**：

```bash
python scripts/validate_chunks.py 语料库.xlsx       # 交付前必跑
python scripts/scan_defects.py [语料库.xlsx]         # 12 类字符缺陷
python scripts/compare_chunking.py --xlsx ...        # 切分方法对比实验

# 检索没命中时，先诊断再调参，不要凭直觉改权重
python scripts/diagnose_misses.py --rag ./rag --q "问题" --want 7.4.1
python scripts/list_principle_clauses.py --rag ./rag  # 该降权的条款清单
```

## 实测数据（本方法在真实规范上的表现）

- 语料库：GB 55037-2022（241 条）+《高层民用建筑消防安全管理规定》（51 条）= **292 条 / 940 切片**
- 多款条文全款交付：**100%**
- 20 道验收题自动跑分：**20/20**
- 与参考库逐字比对：236/241 一致，5 条差异全部可解释

### 含糊问法的短题（商家口吻，12 题）

- 10 题准确 / 2 题正确拒答
- 原则条款降权修复 1 题（7.4.1 从第 2 名到第 1 名），**其余 12 道回归题分数一字节未变**
- 2 题确认是**验收期望值写错**（不是检索问题）——见 `acceptance-design.md`

## 适用范围 / 不适用

✅ GB / GBJ / JGJ 等有 `1` `1.0.1` `1 2 3` 编号层级的工程建设规范
❌ 科普教材、条文说明、案例文书、执法文书等**没有编号层级**的自由文本（那些该用语义切分）

## 版本

- **v1.2.0**（当前）：新增「检索调参实录」（什么有效、什么试过没用）与「验收题设计」（期望值先核实）；
  补充两个可复用诊断脚本 `diagnose_misses.py` / `list_principle_clauses.py`
- v1.1.0：新增"并第二部规范（含法规）"与"20 题验收自动跑分"

## 许可

本仓库内的方法与代码可自由使用、修改、商用。其中引用的国家标准条文片段版权归原发布机构，
仅作方法演示与格式示例，请以正式出版文本为准。
