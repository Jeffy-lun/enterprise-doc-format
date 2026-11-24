# 各格式落地参数与坑

排印规则要落到具体格式里，每个格式有自己的单位、自己的坑。

---

## Word（.docx）

### 必须做

| 项 | 做法 |
|---|---|
| 行距 | `paragraph_format.line_spacing = Pt(22)` + `line_spacing_rule = WD_LINE_SPACING.EXACTLY`（**固定值**，不是 `line_spacing = 1.8`） |
| 中文字体 | 除了 `run.font.name`，还要写 `w:eastAsia`，否则中文回退到默认宋体 |
| 字距 | `w:spacing w:val` 单位是 1/20 pt：9pt × 0.08em = 0.72pt → `val="14"` |
| 表格线 | `tblBorders` 全 `nil`，再逐格 `tcBorders` 画需要的边 |
| 不跨页 | 行 `w:cantSplit`；整表除末行 `keep_with_next` |
| 页码 | 页脚用 `fldSimple` 域 `PAGE`；封面页 `w:titlePg` 单独不显示 |

### 三个必踩的坑

1. **多倍行距不可靠**：`line_spacing = 1.8` 生成 `auto` 行距，由渲染器按自己的字体行高再放大一次，换机器/换渲染器行距就变。中文长文一律固定值。
2. **OOXML 子元素有顺序**：`w:spacing` 必须排在 `w:sz` **之前**，`w:tblBorders` 必须在 `w:tblLayout` **之前**，`w:pBdr` 必须在 `w:spacing` **之前**。顺序错了 Word/LibreOffice **静默忽略**——字距不生效、表格线关不掉、页脚框线去不掉，都是这个原因。写 OOXML 时用「按 schema 顺序插入」的辅助函数，不要直接 `append`。
3. **内置样式带私货**：`Title`、`Heading 1`、`Quote` 样式自带 Calibri / Times 斜体 / 蓝色。要改样式本身（`doc.styles['Heading 1'].font...` + `rFonts.eastAsia`），不能只改段落 run，否则导出 PDF 会暴露无关字体。

### 插图与表格的四个渲染级坑（都由真实文档踩出来）

**坑一：含图片的段落不能用固定行距，否则图被裁成一行高。**
症状：Word 里图是好的，导出 PDF 后图片位置一片空白，图注却正常。原因：正文样式是 `Pt(21) + EXACTLY`，图片段落继承了它，渲染器把 10cm 高的图裁到 21 磅，看着就是空白。修法：

```python
pf.line_spacing = None                       # 不要沿用 Normal 的固定值
pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
p.add_run().add_picture(path, width=Cm(6.9))
```

图注段落可以继续用固定行距（那是文字）。同一份文档里「正文固定行距 + 插图」必然碰撞，写完图必须导一次 PDF 看。

**坑二：`tcBorders` 子元素顺序错了，整条边框设置被丢弃。**
schema 顺序是 `top → left → bottom → right`（不是 top/left/right/bottom）。写错时渲染器**忽略整个 `tcBorders`**，表格回退成默认网格——你会以为「设了 nil 怎么还有线」。

**坑三：LibreOffice 会给「无边框表格」画浅灰参考线。**
`tblBorders` 全 `nil` 也会画，导出 PDF 里就是一层灰格子，正好破坏三线表。两个办法合用：

- 四边给**白色细线**（`val=single, sz=2, color=FFFFFF`）盖掉参考线，视觉等同无边框；
- 需要显示的横线改用**单元格段落的底边框**（`w:pBdr/w:bottom`）画，位置与粗细完全可控，渲染器不能覆盖。

**坑四：列宽总和必须等于版心，且要在三个地方设。**
A4 左右各 3.0cm 时版心 15cm。列宽总和 15.2cm 就会出现列错位、三线断裂。三处都要给：

```python
tblW = sum(widths) / 2.54 * 1440        # 表级总宽，单位 twips
table.columns[i].width = Cm(widths[i])  # 列
cell.width = Cm(widths[i])              # 单元格
```

只设 `cell.width` 在 `tblLayout=fixed` 下不生效，列宽会按 `tblGrid` 走。

### 数字
表格数字右对齐 + 等宽数字；金额 `#,##0`；比率 `0.0%`；变动 `+0.0%;-0.0%`，百分点用 `+0.0" pt"` 并把值乘 100。

---

## Excel（.xlsx）

| 项 | 做法 |
|---|---|
| 网格线 | `ws.sheet_view.showGridLines = False` |
| 列宽 | 项目列 28–32，数值列 14–17，前面留 1 列 3.0 宽做装订留白 |
| 行高 | 标题 32–34、表头 22–24、数据 19–21；空一行做区块间隔 |
| 数字格式 | `#,##0;(#,##0)`（负数括号）、`0.0%`、`+0.0%` |
| 对齐 | 项目列左 + 层级 `indent`；数字右；全部垂直居中 |
| 表格线 | 无边框，只用 `cell.border.bottom`：表头 thin `#111`，合计行 above thin，表尾 medium `#111` |
| 冻结 | `ws.freeze_panes` 到数据首行 |
| 打印 | `fitToWidth=1`、`fitToHeight=0`、`print_area`、页脚左右两端文字、水平居中 |

**坑**：Excel 没有「字间距」，全大写的仪式感只能用字号 + 字重 + 灰色表达；百分比格式的单元格如果存的是百分点（1.4）而格式是 `0.0%`，会显示 140%。

---

## PowerPoint（.pptx）

| 项 | 做法 |
|---|---|
| 字号 | 一页一个主张：标题 28–40pt、正文 16–20pt、注释 10–12pt；正文不小于 16pt |
| 行距 | 1.2–1.35（幻灯片是扫读，不是长读） |
| 行数 | 一页正文 ≤ 6 行、每行 ≤ 24 汉字 |
| 对齐 | 左对齐；数字右对齐 |
| 图表 | 去掉网格线与图例框，直接标数值 |

**坑**：幻灯片的中文行距不要用固定值（内容一多会裁字），用 1.25–1.35 倍。

---

## PDF

| 项 | 做法 |
|---|---|
| 分页 | 章起新页；表格/图/拉引不跨页；页末无孤行 |
| 书签 | 章标题必须用 `Heading 1` 样式（并改掉它的字体字号），导出后自动生成 PDF 大纲 |
| 页眉页脚 | 内页页眉 + 页脚页码；封面 `titlePg` 不显示 |
| 字体 | 必须全部子集嵌入（导出后检查 `/BaseFont` 里的 `XXXXXX+` 前缀） |
| 元数据 | 标题、作者、主题、关键词都写上 |

**坑**：Word 里「直接另存 PDF」会把内置样式的字体一次性带出来（Calibri、等线、Times 斜体），检查 PDF 的字体列表比检查 Word 更快暴露样式层污染。
