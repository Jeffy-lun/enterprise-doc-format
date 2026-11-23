# 让 agent 默认调用这个技能

技能装好只是「可用」。要让它**默认被调用**，得在 agent 的常驻指令里写一句。下面五种形态按平台选一段。

---

## 0. 连安装一起交给 agent（最快）

把这一整段丢给 WorkBuddy / 豆包 / 任何能跑命令的助手，它会自己下载、安装、自报版本，之后按规范工作：

```text
请从这里下载并安装一个写作与排版技能：https://github.com/Jeffy-lun/enterprise-doc-format/releases/download/v1.0.0/enterprise-doc-format-1.0.0.zip —— 解压后把 enterprise-doc-format 整个目录放到你的技能目录（WorkBuddy 为 %USERPROFILE%\.workbuddy\skills-marketplace\skills\，其他平台放到你的 skills 目录），确认目录里有 SKILL.md 和 scripts/audit_docx.py；装不了文件时（例如没有文件写入权限），就把 SKILL.md 全文读进上下文当作常驻规范；无论哪种方式，装完先回我一句「enterprise-doc-format 已安装 + 版本号」，然后把它作为此后所有企业文档（年报、董事会材料、方案书、白皮书、汇报、周报、制度、SOP、会议纪要、对外函件、简历）的书写与排版规范执行：先按 references/house-styles.md 定版式，内容遵 references/writing.md，排版遵 references/typography.md，交付前运行 python scripts/audit_docx.py <文件>，严重项必须为 0 才算完成。
```

---

## 1. AGENTS.md（通用：Claude Code、DSH、Codex、Cursor、Windsurf 等都读）

贴到用户级 `~/.agents/AGENTS.md` 或项目根的 `AGENTS.md`：

```markdown
## 文档格式

写、改、审任何企业办公文件（年报、董事会材料、方案书、白皮书、汇报、制度、纪要、对外函件、简历）
之前，先加载 `enterprise-doc-format` 技能并按它走：

1. 先定体裁（board / whitepaper / memo），从 `references/house-styles.md` 取版式，不临场发明字号
2. 内容按 `references/writing.md` 写：一句一个指标、数字要能勾稽、主语统一、模糊量词给数
3. 排版按 `references/typography.md`：中文大标题行高 1.3–1.4、正文 1.7–1.8、中文零字距、正文 20–35 字/行、表格只留表头细线 + 表尾粗线
4. 交付前跑 `python scripts/audit_docx.py <文件>`，**严重项清零**才算完成

已经有 .docx 时，先 `python scripts/restyle_docx.py <文件> --style board` 收口样式层，再跑审计。
```

## 2. CLAUDE.md（Claude Code 项目级）

同一段贴到项目 `CLAUDE.md`，或用户级 `~/.claude/CLAUDE.md`。项目里已有别的规则时放在最前面。

## 3. WorkBuddy：设为默认技能

安装后在该技能详情里打开「默认启用 / 自动调用」。若客户端没有开关，把下面这句加进「全局指令 / 默认提示词」：

```
涉及文档写作与排版时，自动使用 enterprise-doc-format 技能：先按 house-styles 定版式，
内容遵 writing.md，排版遵 typography.md，产出后必须跑 audit_docx.py 并且严重项为 0。
```

## 4. 豆包 / 其他只吃提示词的平台

在「智能体 → 人设与回复逻辑（系统提示词）」最前面加：

```
你在写任何办公文件（报告、方案、汇报、制度、函件）时，一律遵循以下两套规范：
【书写】一句一个指标；数字必须能勾稽（分项之和=合计、本期/上期-1=同比）；全篇主语统一
（只用「公司」或只用「我们」）；模糊量词（大幅、显著、较多、约）必须给数字；单位与口径全篇一致
（亿元/万元不混，变动写「个百分点」不写 pt 或 %）；不用破折号当标点、不用直引号、不用三个点；
中文不用斜体；绝对化断言（唯一、第一、彻底）无数据支撑就删。
【排版】字号不超过 7 档、比例 1.2–1.25；中文正文行高 1.7–1.8，大标题 1.3–1.4（禁止沿用拉丁的
1.0–1.2）；中文字距一律 0，全大写拉丁加 0.06–0.1em；正文一行 20–35 汉字，不用两端对齐，
用段间距而非首行缩进；表格不画网格线，只留表头细线 + 表尾粗线，数字右对齐、单位写表头；
正式文件有页码，章标题不与正文分离。
```

需要脚本能力时，提示豆包「调用 Python 工具运行 audit_docx.py」。

## 5. Cursor / Copilot rules（`.cursor/rules/*.mdc`、`.github/copilot-instructions.md`）

```markdown
写文档时遵循 enterprise-doc-format 技能：版式取 references/house-styles.md，
文字遵 references/writing.md，排版遵 references/typography.md；
保存 .docx 前跑 scripts/audit_docx.py，严重项必须为 0。
```

---

## 验证默认调用是否生效

随便让 agent 写一份带表格的汇报，看它有没有做这三件事：

1. 报出用了哪套版式（「按 board 排」）
2. 中文行距是否 1.7–1.8（不是 1.15）
3. 交付前是否跑了 `audit_docx.py`

三条缺一条，说明默认调用没生效，检查常驻指令是否被后续指令覆盖。
