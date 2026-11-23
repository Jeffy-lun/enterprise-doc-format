---
name: enterprise-doc-format
description: "When the user writes, restyles, or reviews a business document (annual report, board paper, proposal, white paper, memo, SOP, contract summary, resume) and expects it to read well and look professionally typeset — especially Chinese or mixed CJK/Latin Word/PowerPoint/Excel/PDF output. Also use when asked to 'check formatting', '审一下排版', '这份文档哪里别扭', '按规范检查文档', '这张表怎么这么难看', or to fix inconsistent fonts/line spacing/table borders/gridlines/number formats. Covers prose rules (voice, banned words, units, numbers, headings, AI tells), typographic rules (type scale, CJK leading, tracking, measure, table rules, page setup), and comes with scripts that audit .docx, .xlsx and .pptx automatically."
description_zh: "企业办公文件（年报/董事会材料/方案/白皮书/汇报/制度）的书写规范与排印规范检查与重排；可用脚本自动审 .docx 并一键套用公司版式。"
description_en: "Write, audit, and restyle business documents: prose rules plus CJK-aware typography, with scripts that check and reformat .docx files."
version: 1.2.0
allowed-tools: Bash, Read, Write, Edit
metadata:
  emoji: "📄"
  category: documents
  requires:
    bins:
      - python
    python:
      - python-docx
      - lxml
  tags:
    - typography
    - writing
    - docx
    - office
    - cjk
---

# 企业办公文件：书写规范 + 排印规范

一份企业文档出问题，永远出在两处：**话没写对**（主语漂、数字矛盾、术语不统一、AI 腔），**字没排对**（行距贴死、字号乱、表格网格线糊成一片）。这套技能把两件事合成一条流程，并且能自动检查。

## 何时用

用户在做任何**要被别人读到、可能被打印或转发**的企业文件时：

- 年报、半年报、董事会/经营分析材料、投资与融资文件
- 方案书、白皮书、尽调回复、投标文件
- 内部汇报、周报月报、复盘、制度与 SOP、会议纪要
- 合同要点、对外函件、简历、对外 PDF
- Word / PowerPoint / Excel / PDF 任一格式

触发词：写报告、改文档、排版、审排版、格式优化、文风检查、按规范检查、把这份文档弄得像样一点。

## 工作流（四步，不要跳）

### 1. 定体裁与版式

先问清三件事（用户没说就按默认推，别连着问）：

| 体裁 | 版式预设 | 典型体积 |
|---|---|---|
| 董事会报告 / 年报 | `board` | 5–60 页，A4 纵向 |
| 白皮书 / 方案书 | `whitepaper` | 10–40 页，A4 纵向 |
| 内部汇报 / 周报 | `memo` | 1–5 页，A4 纵向 |

预设值见 `references/house-styles.md`。**不要临场发明字号和行距**——预设已经过校验。

### 2. 先写，按书写规范

`references/writing.md` 是硬规则，重点是这几条（全量在那里）：

- 一句一个指标；数字必须能勾稽（收入 = 分项之和，同比 = 本期/上期 − 1）
- 事实自相矛盾是最贵的错：「净利润 +8.7%」与「利润下滑」不能并存
- 主语统一：全篇只用「公司」或只用「我们」，不混
- 模糊量词要么删，要么给数字：「大幅增长」→「增长 12.4%」
- 单位与口径写一次就要一致：`亿元`/`万元` 不混，`个百分点` 不写成 `pt` 或 `%`
- 不用破折号当标点、不用直引号、不用三个点；中文不用斜体
- 绝对化断言（唯一、最、第一、全部）没有数据支撑就删

### 3. 再排，按排印规范

`references/typography.md` 是硬规则。**中文的三条最容易错**：

- 大标题行高 **1.3–1.4**（拉丁那套 1.0–1.2 会让中文撞行），正文 **1.7–1.8**
- 负字距**只给拉丁**；中文一律 `0`
- 全大写拉丁必须正字距 `0.06–0.1em`

其余：字号阶梯 1.2–1.25 且全篇 ≤ 7 档、正文中文 20–35 字/行、不用两端对齐、表格只留「表头细线 + 表尾粗线」、数字右对齐。

### 4. 自动检查，改到零严重

```bash
python scripts/audit_docx.py 报告.docx            # Word / Markdown / 纯文本
python scripts/audit_docx.py 报告.docx --json     # 机读，给 agent 迭代
python scripts/audit_xlsx.py 报表.xlsx            # Excel 工作簿
python scripts/audit_xlsx.py 报表.xlsx --sheet 利润表
python scripts/audit_pptx.py 汇报.pptx            # PowerPoint
python scripts/audit_pptx.py 汇报.pptx --slide 3
python scripts/audit_docx.py notes.md             # 纯文本/ Markdown 只跑书写规则
```

输出格式固定为 `文件:位置 编号 问题 → 改法`，例如：

```
报告.docx:p12 T4 标题 18pt、行距 1.15（中文大标题需 1.3–1.4，会撞行）→ 行距固定 25pt
报告.docx:p7  W2 直引号 "设备卖出去…" → 中文双引号「」
报告.docx:t2r1 W4 数字与单位之间缺空格「48.62亿元」→「48.62 亿元」
```

**严重项必须清零**：`T1 T4 T8 T9 W1 W2 W4 W5 W6`。提示项（`T13` 无页码等）按体裁判断。

已有文档想直接收口：

```bash
python scripts/restyle_docx.py 报告.docx --style board          # 另存 报告-排版版.docx
python scripts/restyle_docx.py 报告.docx --style board --inplace # 就地改（先备份）
```

重排只动**样式层**：字体、字号、固定行距、段间距、对齐、表格线、页脚页码。不改文字内容——文字的事回第 2 步。

## 交付前自检

- [ ] `audit_docx.py` 的严重项为 0
- [ ] 数字能对上：分项之和、同比、跨表同名指标一致
- [ ] 全篇同一套字号（≤ 7 档）、同一套字体（≤ 2 族）
- [ ] 大标题没有撞行；正文没有一行超过 40 汉字
- [ ] 表格没有网格线；数字右对齐且位数对齐
- [ ] **Excel**：网格线已关、数字有格式（千分位 / 百分比）、列宽不截断、比率没存成小数、表头加粗且有单位声明
- [ ] **PowerPoint**：一页一个主张、正文 ≥16pt、一页 ≤6 条 / ≤120 字、图表无网格线、图片有 alt
- [ ] 有页码；打印预览无孤行、表格不跨页
- [ ] 没有 `—`、`...`、`"`、`6.7 %`、`48.62亿元`

## 与其他技能的关系

- 长文/杂志式版式（拉引、编辑级排版）→ 先看 `references/typography.md` 的「编辑级」一节
- 营销文案（首页、落地页、广告）→ 用 `copywriting`，本技能只管企业文件
- 图表与可视化 → 用 `archify` / `diagram`，本技能只管文字与表格
- PPT 与 Excel 的具体参数 → `references/office-params.md`

## 参考文件

| 文件 | 内容 |
|---|---|
| `references/writing.md` | 书写规范全量（语气、结构、数字、单位、标题、AI 腔、禁用词） |
| `references/typography.md` | 排印规范全量（阶梯、行高、字距、行长、表格、分页、CJK 覆盖） |
| `references/house-styles.md` | 三套可直接用的企业版式（pt 值、表格线、页眉页脚） |
| `references/office-params.md` | Word / Excel / PPT / PDF 各自的落地参数与坑 |
| `references/checklist.md` | 检查矩阵（编号 ↔ 规则 ↔ 严重度），含 Excel 的 X、PowerPoint 的 P |
| `scripts/audit_docx.py` | 审计 .docx / .md / .txt，输出问题清单 |
| `scripts/audit_xlsx.py` | 审计 .xlsx / .xlsm（网格线、数字格式、列宽、混列、打印设置） |
| `scripts/audit_pptx.py` | 审计 .pptx（一页一条主张、字号下限、念稿、图表网格线、alt 文本） |
| `scripts/text_rules.py` | 书写规则共用实现（三个审计器都走它，避免规则分叉） |
| `scripts/restyle_docx.py` | 按预设重排 .docx 样式层 |
| `install/DEFAULT-INVOKE.md` | 让各 agent（Claude / DSH / 豆包 / WorkBuddy）默认调用 |
