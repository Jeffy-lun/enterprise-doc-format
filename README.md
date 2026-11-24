# enterprise-doc-format

企业办公文件的**书写规范 + 排印规范**，合并自两套来源：

- [Vercel Writing Guidelines](https://github.com/vercel-labs/writing-guidelines)（prose 规则：语气、结构、数字、单位、标点、AI 腔）
- OpenDesign `craft/typography.md` + `typography-hierarchy(.editorial)`（排印规则：字号阶梯、CJK 行高、字距、行长、表格、分页）

两套规则原本是给英文开发者文档和网页设计用的。这个技能把它们重写成**中文企业文件可用**的版本，加了四件事：

1. **自动审计（Word）**：`scripts/audit_docx.py` 读一份 .docx，输出 `文件:位置 编号 问题 → 改法`，严重项 > 0 时退出码 1。
2. **自动审计（Excel）**：`scripts/audit_xlsx.py` 查网格线、数字格式、列宽截断、数字列夹文字、比率存成小数、冻结与打印设置 —— 企业报表的四个常见死因。
3. **自动审计（PowerPoint）**：`scripts/audit_pptx.py` 查一页几个主张、正文是否 ≥16pt、是否在念稿、图表网格线、图片 alt —— 判断「这页能不能在 5 秒内读到主张」。
4. **一键重排**：`scripts/restyle_docx.py` 按预设（`board` / `whitepaper` / `memo`）改样式层——字体、字号、固定行距、段距、对齐、表格线、页码。不改一个字。

## 一分钟上手

```bash
# Word 审计（人读）
python scripts/audit_docx.py 年报.docx

# Word 审计（给 agent 迭代用）
python scripts/audit_docx.py 年报.docx --json

# Excel 审计
python scripts/audit_xlsx.py 报表.xlsx

# PowerPoint 审计
python scripts/audit_pptx.py 汇报.pptx

# 按董事会版式重排（另存 年报-排版版.docx）
python scripts/restyle_docx.py 年报.docx --style board
```

依赖：`python-docx`、`openpyxl`、`python-pptx`、`lxml`（`pip install python-docx openpyxl python-pptx`）。

## 输出长什么样

```
年报.docx:p2  T1 中文正文 10.5pt、行距 1.00（18 处不足） → 行距改固定值 19pt（1.8 倍）
年报.docx:p1  T4 标题 26pt、行距 1.00，中文会撞行（1 处） → 行距改固定值 35pt（1.35 倍）
年报.docx:t0  T9 表格画了网格线 → 关掉 insideH/insideV 与外框，只留表头细线 + 表尾粗线
年报.docx:p16 W1 用破折号当标点（1 处） → 改冒号/逗号/句号，或重写
年报.docx:p2  W4 数字与单位之间缺空格（1 处） → 加半角空格：48.62 亿元
年报.docx:all W11 「公司」与「我们/笔者」同篇混用 → 全篇统一一个主语

严重 8 条 / 提示 7 条　（严重项需清零）
```

真实效果（同一份年报，重排前后）：

| | 重排前 | 重排后 |
|---|---|---|
| 严重项 | 8 | 4 |
| 修掉的 | — | T1 行距、T4 标题撞行、T8 两端对齐、T9 表格网格、T6 全大写字距、T13 无页码、T15 多倍行距 |
| 剩下的 | — | 全是文字问题（W1 破折号、W4 数字贴单位、W11 主语混用、W12 pt/个百分点） |

重排只管样式层，所以剩下的必然是要动笔改的那些——这个分工是刻意的。

Excel 同理（同一份财务报表，两种排法）：

| | 未排印 | 精排 |
|---|---|---|
| 严重项 | **8** | **0** |
| 命中的问题 | X1 网格线、X2 General 格式、X3 数字左对齐、X4 列宽截断、X8 全网格边框、X9 表头无区分、X11 无打印设置 + 文字类 W4/W8 | — |

PowerPoint 同理（同一套内容，两种做法）：

| | 未排印 | 精排 |
|---|---|---|
| 严重项 | **7** | **0** |
| 命中的问题 | P1 正文 12pt、P2 一页 9 条、P3 一页 153 字在念稿、P7 中文斜体、P10 全靠文本框、P11 4:3 画幅、P12 图片无 alt、P14 缺标题 + P9 图表网格线 | — |

## 内容

```
SKILL.md                     技能说明 + 四步工作流
references/
  writing.md                 书写规范（W 编号，可对号查）
  typography.md              排印规范（含中文覆盖、编辑级排版）
  house-styles.md            三套版式（board / whitepaper / memo）的具体 pt 值
  office-params.md           Word / Excel / PPT / PDF 各自的落地参数与坑
  checklist.md               检查矩阵（W / T / X 三组编号）
scripts/
  audit_docx.py              审计 .docx / .md / .txt
  audit_xlsx.py              审计 .xlsx / .xlsm
  audit_pptx.py              审计 .pptx
  text_rules.py              书写规则共用实现（三个审计器都走它）
  restyle_docx.py            按预设重排 .docx 样式层
install/DEFAULT-INVOKE.md    让各 agent 默认调用
evals/                       fixture 生成器 + 回归测试
```

## 一句话装进 agent（WorkBuddy / 豆包 / 任何能跑命令的助手）

把下面整段复制给你的 agent（WorkBuddy、豆包、Cursor 里的助手都行），它会自己下载、解压、装好并立刻按这套规范工作：

```text
请从这里下载并安装一个写作与排版技能：https://github.com/Jeffy-lun/enterprise-doc-format/releases/download/v1.0.0/enterprise-doc-format-1.0.0.zip —— 解压后把 enterprise-doc-format 整个目录放到你的技能目录（WorkBuddy 为 %USERPROFILE%\.workbuddy\skills-marketplace\skills\，其他平台放到你的 skills 目录），确认目录里有 SKILL.md 和 scripts/audit_docx.py；装不了文件时（例如没有文件写入权限），就把 SKILL.md 全文读进上下文当作常驻规范；无论哪种方式，装完先回我一句「enterprise-doc-format 已安装 + 版本号」，然后把它作为此后所有企业文档（年报、董事会材料、方案书、白皮书、汇报、周报、制度、SOP、会议纪要、对外函件、简历）的书写与排版规范执行：先按 references/house-styles.md 定版式，内容遵 references/writing.md，排版遵 references/typography.md，交付前运行 python scripts/audit_docx.py <文件>，严重项必须为 0 才算完成。
```

它会做四件事：下载 → 落到技能目录（或降级为读入上下文）→ 自报版本 → 之后所有文档都按规范走并跑审计。装完想确认生效，随便让它写一份带表格的汇报，看中文行距是不是 1.7–1.8、有没有跑 `audit_docx.py`。

## 安装（手动）

### Claude Code / Claude Desktop（Agent Skills）

```bash
git clone <this-repo> ~/.claude/skills/enterprise-doc-format
# 或直接拷目录
```

### 通用 agent（AGENTS.md 生态、Cursor、Codex 等）

```bash
cp -r enterprise-doc-format ~/.agents/skills/
```

然后把 `install/DEFAULT-INVOKE.md` 里的相应片段贴进 `AGENTS.md`（用户级或项目级）。

### WorkBuddy

把整个目录放进技能市场目录，或用「安装技能」上传 zip：

```
%USERPROFILE%\.workbuddy\skills-marketplace\skills\enterprise-doc-format\
```

### 豆包 / 其他平台（只吃 zip 或纯提示词）

1. 上传 `dist/enterprise-doc-format-1.0.0.zip`（技能/工具安装入口）
2. 平台不支持技能包时：把 `SKILL.md` + `references/writing.md` + `references/typography.md` 三段贴进「提示词 / 知识库」，把 `install/DEFAULT-INVOKE.md` 的默认调用句放在最前面
3. 脚本能力需要能跑 Python 的环境；纯提示词模式下只能靠 `references/checklist.md` 人工过

### DSH（DeepSeek Harness）

技能目录加到会话技能路径下即可；本机已装到 `~/.agents/skills/enterprise-doc-format`。

## 默认调用

见 `install/DEFAULT-INVOKE.md`——里面有四种形态的现成句子（AGENTS.md、WorkBuddy 默认技能、豆包智能体提示词、Cursor rules）。

## 版本

`1.2.1` — 新增 T17（图片段落用固定行距 → 导出 PDF 图变空白）；`references/office-params.md` 补齐四个渲染级坑（图片段落行距、`tcBorders` 顺序、白色边框盖参考线、列宽三处设全）。

`1.2.0` — 新增 PowerPoint 审计（`audit_pptx.py`，P1–P15）：一页一主张、正文 ≥16pt、念稿、图表网格线、alt 文本、画幅；evals 增加 pptx fixture，CI 三格式全跑。

`1.1.0` — 新增 Excel 审计（`audit_xlsx.py`，X1–X15），书写规则抽成 `text_rules.py` 供多个审计器共用；evals 增加 xlsx fixture。

`1.0.0` — 首次发布。规则来源锁定：writing-guidelines `83e2316`；OpenDesign craft 规则库同期快照。

后续加规则时：改 `references/*.md` → 在 `checklist.md` 加编号 → 在 `audit_docx.py` 加检测 → 跑 `evals/evals.json` 里的用例 → 升 version。

## 许可

规则文本来自上游两个 MIT 项目，改写部分同 MIT。
