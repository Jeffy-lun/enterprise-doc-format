#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""生成审计用的两份 fixture：bad.docx（满身问题）与 good.docx（合规基准）。

不依赖任何外部文件，纯 python-docx 现场造。供 evals/run_evals.py 与 CI 使用。

    python evals/make_fixtures.py
"""
import os

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "fixtures")

BODY_CN = "微软雅黑"
HEAD_CN = "宋体"
LATIN = "Segoe UI"


def font(run, size, *, cn=BODY_CN, lat=LATIN, bold=False, color="111111",
         tracking=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = lat
    run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(W + "rFonts")
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(W + "ascii", lat)
    rFonts.set(W + "hAnsi", lat)
    rFonts.set(W + "eastAsia", cn)
    if tracking:
        sp = OxmlElement("w:spacing")
        sp.set(W + "val", str(int(round(tracking * size * 20))))
        rPr.append(sp)


def para(doc, text, *, size=11.5, role="body", lead=21, before=0, after=8,
         align=WD_ALIGN_PARAGRAPH.LEFT, cn=BODY_CN, bold=False,
         color="111111", tracking=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = Pt(lead)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.first_line_indent = Pt(0)
    p.alignment = align
    font(p.add_run(text), size, cn=cn, bold=bold, color=color,
         tracking=tracking)
    return p


def page_margins(doc, t=2.5, b=2.4, l=3.0, r=3.0):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(t), Cm(b)
    sec.left_margin, sec.right_margin = Cm(l), Cm(r)


def footer_page_number(doc, size=9):
    sec = doc.sections[0]
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run("第 "), size, color="555555")
    fld = OxmlElement("w:fldSimple")
    fld.set(W + "instr", " PAGE ")
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz"); sz.set(W + "val", str(int(size * 2)))
    col = OxmlElement("w:color"); col.set(W + "val", "555555")
    rPr.append(col); rPr.append(sz)
    t = OxmlElement("w:t"); t.text = "1"
    r.append(rPr); r.append(t)
    fld.append(r)
    p._p.append(fld)
    font(p.add_run(" 页"), size, color="555555")


def borderless_table(doc, rows):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    tblPr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(W + "val", "nil")
        borders.append(el)
    tblPr.append(borders)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            cell.width = Cm(3.6)
            p = cell.paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.LEFT if ci == 0
                           else WD_ALIGN_PARAGRAPH.RIGHT)
            font(p.add_run(str(val)), 9 if ri == 0 else 10.5,
                 color="555555" if ri == 0 else "111111", bold=ri == 0)
            tcPr = cell._tc.get_or_add_tcPr()
            tcB = OxmlElement("w:tcBorders")
            for edge in ("top", "left", "right"):
                el = OxmlElement(f"w:{edge}")
                el.set(W + "val", "nil")
                tcB.append(el)
            b = OxmlElement("w:bottom")
            b.set(W + "val", "single")
            b.set(W + "sz", "6" if ri == 0 else ("12" if ri == len(rows) - 1 else "4"))
            b.set(W + "color", "111111" if ri in (0, len(rows) - 1) else "CBC9C3")
            tcB.append(b)
            tcPr.append(tcB)
    return t


def make_bad(path):
    doc = Document()
    page_margins(doc)
    style = doc.styles["Normal"]
    style.font.size = Pt(10.5)
    style.font.name = "Times New Roman"
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    style.paragraph_format.line_spacing = 1.0
    style.paragraph_format.space_after = Pt(0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    font(p.add_run("ANNUAL REPORT 2025"), 10.5, cn="宋体", lat="Times New Roman")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    font(p.add_run("2025年年度报告"), 26, cn="宋体", lat="Times New Roman", bold=True)

    body = (
        "2025年，公司实现营业收入48.62亿元，同比增长12.4%；归属于母公司股东的"
        "净利润5.31亿元，同比增长8.7%。增速比2024年低3.1个百分点。"
    )
    for text in (
        body,
        "我们没有把利润下滑归因于外部环境就算完。三季度的产线改造把精密传动件的"
        "单位工时降了9%，这是今年唯一值得写进报告的自救动作。",
        "设备卖出去才是服务的开始——这是今年的核心判断。",
    ):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.0
        font(p.add_run(text), 10.5, cn="宋体", lat="Times New Roman")

    # 带网格线的表格
    t = doc.add_table(rows=4, cols=3)
    t.style = "Table Grid"
    data = [("项目", "2025年", "2024年"),
            ("营业收入（亿元）", "48.62", "43.25"),
            ("毛利率", "31.6%", "30.2%"),
            ("资产负债率", "38.4%", "41.0%")]
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            font(p.add_run(val), 10.5, cn="宋体", lat="Times New Roman")

    doc.save(path)


def make_good(path):
    doc = Document()
    page_margins(doc)
    normal = doc.styles["Normal"]
    normal.font.size = Pt(11.5)
    normal.font.name = LATIN
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_CN)

    para(doc, "ANNUAL REPORT 2025", size=9, lead=14, after=0,
         color="555555", tracking=0.08)
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(24)
    h.paragraph_format.space_after = Pt(14)
    h.paragraph_format.line_spacing = Pt(41)
    h.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    font(h.add_run("2025 年年度报告"), 30, cn=HEAD_CN, bold=True)

    para(doc, "营收 48.62 亿元，同比增 12.4%；越南工厂 6 月投产，"
              "海外收入占比升到 21.4%。", size=13, lead=20, after=14)

    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(24)
    h1.paragraph_format.space_after = Pt(8)
    h1.paragraph_format.line_spacing = Pt(25)
    h1.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    font(h1.add_run("致股东"), 18, cn=HEAD_CN, bold=True)

    para(doc, "2025 年，公司实现营业收入 48.62 亿元，同比增长 12.4%。"
              "归属于母公司股东的净利润 5.31 亿元，同比增长 8.7%，低于营收增速。")
    para(doc, "公司没有把增速放缓全部归因于外部环境。三季度完成产线改造，"
              "精密传动件的单位工时下降 9%，四季度毛利率回升到 32.4%。")

    para(doc, "金额单位见表内", size=9, lead=14, after=4)
    borderless_table(doc, [("项目", "2025 年", "2024 年"),
                           ("营业收入（亿元）", "48.62", "43.25"),
                           ("毛利率", "31.6%", "30.2%"),
                           ("资产负债率", "38.4%", "41.0%")])
    para(doc, "注：数据未经审计，以年度报告全文为准。", size=9, lead=14,
         before=6, after=0)

    footer_page_number(doc)
    doc.save(path)


def make_bad_xlsx(path):
    """Excel 的典型样子：网格线开着、General 格式、数字左对齐、列宽截断、
    表头不加粗、没冻结没打印区、同列数字文字混排。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "利润表"
    ws["A1"] = "恒启精密制造股份有限公司 2025 年年度财务数据（未经审计，单位：万元）"

    data = [("项目", "2025年", "2024年", "变动"),
            ("营业收入（亿元）", 48.62, 43.25, 0.124),
            ("毛利率", 0.316, 0.302, 0.014),
            ("归母净利润（亿元）", 5.31, 4.88, 0.087),
            ("资产负债率", 0.384, 0.410, -0.026)]
    for ri, row in enumerate(data, start=3):
        for ci, val in enumerate(row, start=1):
            c = ws.cell(ri, ci, val)
            c.alignment = Alignment(horizontal="left")
            c.border = Border(*(Side(style="thin"),) * 4)
            c.font = Font(name="宋体", size=11)
    ws.column_dimensions["A"].width = 8.43
    ws.cell(6, 3, "见附注")          # 数字列中间夹一个文字 → X12
    ws.cell(9, 1, "备注：营业收入48.62亿元，同比增12.4%")
    ws.cell(10, 1, "这是今年唯一值得写进报告的自救动作")
    wb.save(path)


def make_good_xlsx(path):
    """按 references/typography.md 第 10 节排的报表：关网格线、数字右对齐带格式、
    只有表头细线 + 表尾粗线、冻结、打印区、单位声明。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.properties import PageSetupProperties

    wb = Workbook()
    ws = wb.active
    ws.title = "利润表"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    for col in "CDE":
        ws.column_dimensions[col].width = 16

    def put(r, c, v, *, size=10.5, bold=False, color="111111", fmt=None,
            align="left"):
        cell = ws.cell(r, c, v)
        cell.font = Font(name="Segoe UI", size=size, bold=bold, color=color)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        if fmt:
            cell.number_format = fmt
        return cell

    put(2, 2, "恒启精密制造股份有限公司", size=9, bold=True, color="555555")
    put(3, 2, "利润表｜2025 年年度财务数据", size=18, bold=True)
    put(4, 2, "报告期 2025-01-01 至 2025-12-31　金额单位：人民币万元", size=9,
        color="555555")

    head = ("项目", "2025 年", "2024 年", "变动")
    for i, h in enumerate(head, start=2):
        cell = put(6, i, h, size=9, bold=True, color="555555",
                   align="left" if i == 2 else "right")
        cell.border = Border(bottom=Side(style="thin", color="111111"))

    rows = [("营业收入（亿元）", 48.62, 43.25, 0.124, "#,##0"),
            ("毛利率", 0.316, 0.302, 0.014, "0.0%"),
            ("归母净利润（亿元）", 5.31, 4.88, 0.087, "#,##0"),
            ("资产负债率", 0.384, 0.410, -0.026, "0.0%")]
    for ri, (label, a, b, yoy, fmt) in enumerate(rows, start=7):
        ws.row_dimensions[ri].height = 21
        put(ri, 2, label)
        put(ri, 3, a, fmt=fmt, align="right")
        put(ri, 4, b, fmt=fmt, align="right")
        put(ri, 5, round(yoy * 100, 1) if fmt in ("0.0%",) else yoy,
            fmt='+0.0" pt"' if fmt == "0.0%" else "+0.0%;-0.0%", align="right")
        last = ri == 6 + len(rows)
        for ci in range(2, 6):
            ws.cell(ri, ci).border = Border(
                bottom=Side(style="medium" if last else "hair", color="111111"))

    put(12, 2, "注：数据未经审计，以年度报告全文为准。", size=9, color="555555")
    ws.freeze_panes = "C7"
    ws.print_area = "A1:E12"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    wb.save(path)


def _pixel():
    """造一张 8×8 PNG 供 fixture 用（图片无替代文本才能测 P12）。"""
    from PIL import Image

    path = os.path.join(OUT, "_pixel.png")
    if not os.path.exists(path):
        os.makedirs(OUT, exist_ok=True)
        Image.new("RGB", (8, 8), (200, 200, 200)).save(path)
    return path


def make_bad_pptx(path):
    """幻灯片的典型样子：4:3、正文 12pt、一页 9 条要点、一页念稿、
    中文斜体、三种字体、图表带网格线、图片没替代文本。"""
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Emu, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(9144000), Emu(6858000)  # 4:3

    # 第 1 页：标题 + 9 条要点 + 12pt
    s1 = prs.slides.add_slide(prs.slide_layouts[1])
    s1.shapes.title.text = "2025年经营情况"
    body = s1.placeholders[1].text_frame
    body.text = "营业收入48.62亿元"
    for line in ("归母净利润5.31亿元", "经营现金流7.05亿元", "研发投入3.28亿元",
                 "毛利率31.6%", "海外收入10.41亿元", "越南工厂投产",
                 "在手订单21.4亿元", "分红1.20亿元"):
        body.add_paragraph().text = line
    for p in body.paragraphs:
        for r in p.runs:
            r.font.size = Pt(12)
            r.font.name = "宋体"

    # 第 2 页：念稿（200+ 字）+ 中文斜体
    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "业务回顾"
    tb = s2.placeholders[1].text_frame
    tb.text = ("2025 年公司实现营业收入 48.62 亿元，同比增长 12.4%；归属于母公司股东的"
               "净利润 5.31 亿元，同比增长 8.7%，低于营收增速。二季度钢材价格上行，"
               "上年 640 万元的汇兑收益又转为本年 2,180 万元损失，净利增速因此被压住。"
               "公司没有把增速放缓全部归因于外部环境，三季度完成产线改造。")
    for p in tb.paragraphs:
        for r in p.runs:
            r.font.size = Pt(12)
            r.font.name = "宋体"
            r.font.italic = True

    # 第 3 页：标题 + 图表（带网格线）+ 图片（无替代文本）+ 另一种字体
    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    s3.shapes.title.text = "收入结构"
    data = CategoryChartData()
    data.categories = ["传动部件", "系统集成", "运维服务"]
    data.add_series("营收", (26.80, 15.42, 6.40))
    s3.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(600000),
                        Emu(1600000), Emu(5000000), Emu(3000000), data)
    s3.shapes.add_picture(_pixel(), Emu(6100000), Emu(1600000),
                          Emu(1200000), Emu(1200000))
    # 图片替代文本留空（真实文件里最常见的样子）
    for shp in s3.shapes:
        if getattr(shp, "shape_type", None) is not None:
            from pptx.enum.shapes import MSO_SHAPE_TYPE
            if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shp._element.nvPicPr.cNvPr.set("descr", "")
    # 一张四边都有框的表格
    tbl = s3.shapes.add_table(4, 3, Emu(600000), Emu(4800000),
                              Emu(5000000), Emu(1200000)).table
    for ri, row in enumerate((("项目", "2025", "2024"),
                              ("营收（亿元）", "48.62", "43.25"),
                              ("净利（亿元）", "5.31", "4.88"),
                              ("毛利率", "31.6%", "30.2%"))):
        for ci, val in enumerate(row):
            tbl.cell(ri, ci).text = val

    # 第 4 页：无标题，正文用普通文本框（没有大纲结构）
    s4 = prs.slides.add_slide(prs.slide_layouts[6])
    for i in range(4):
        box = s4.shapes.add_textbox(Emu(600000), Emu(600000 + i * 800000),
                                    Emu(6000000), Emu(600000))
        para = box.text_frame.paragraphs[0]
        run = para.add_run()
        run.text = f"第 {i + 1} 条：公司口径为 48.62亿元，未经审计"
        run.font.size = Pt(11)
        run.font.name = "Calibri"

    prs.save(path)


def make_good_pptx(path):
    """一页一个主张：16:9、标题 32pt、正文 18pt、每页 ≤6 条 ≤120 字、
    图表无网格线、用版式占位符、全大写字距。"""
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Emu, Pt

    prs = Presentation()                      # 默认 4:3，显式改成 16:9
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)

    def set_run(run, size, bold=False, name="微软雅黑", color="111111",
                spc=None):
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = name
        run.font.italic = False
        rPr = run._r.get_or_add_rPr()
        rPr.set("lang", "zh-CN")
        if spc:
            rPr.set("spc", str(spc))

    # 第 1 页：标题 + 3 条
    s1 = prs.slides.add_slide(prs.slide_layouts[1])
    s1.shapes.title.text = "2025 年营收 48.62 亿元，增长 12.4%"
    for p in s1.shapes.title.text_frame.paragraphs:
        for r in p.runs:
            set_run(r, 32, bold=True, name="宋体")
    tf = s1.placeholders[1].text_frame
    tf.text = "净利 5.31 亿元，同比 +8.7%，低于营收增速"
    for line in ("海外收入占比 21.4%（上年 17.8%）", "在手订单 21.4 亿元，覆盖 6 个月"):
        tf.add_paragraph().text = line
    for p in tf.paragraphs:
        for r in p.runs:
            set_run(r, 18)

    # 第 2 页：图表，无网格线
    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.shapes.title.text = "三条业务线，运维毛利率最高（41.8%）"
    for p in s2.shapes.title.text_frame.paragraphs:
        for r in p.runs:
            set_run(r, 28, bold=True, name="宋体")
    data = CategoryChartData()
    data.categories = ["传动部件", "系统集成", "运维服务"]
    data.add_series("营收（亿元）", (26.80, 15.42, 6.40))
    ch = s2.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, 609600, 1600200,
                             5486400, 3200400, data).chart
    ch.has_legend = False
    ch.value_axis.has_major_gridlines = False
    ch.value_axis.visible = False

    # 第 3 页：全大写 kicker 带字距
    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    s3.shapes.title.text = "2026 年的承诺"
    for p in s3.shapes.title.text_frame.paragraphs:
        for r in p.runs:
            set_run(r, 28, bold=True, name="宋体")
    kicker = s3.shapes.add_textbox(609600, 3200400, 5486400, 457200)
    kp = kicker.text_frame.paragraphs[0]
    kr = kp.add_run()
    kr.text = "OUTLOOK 2026"
    set_run(kr, 12, spc=800)

    prs.save(path)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    make_bad(os.path.join(OUT, "bad.docx"))
    make_good(os.path.join(OUT, "good.docx"))
    make_bad_xlsx(os.path.join(OUT, "bad.xlsx"))
    make_good_xlsx(os.path.join(OUT, "good.xlsx"))
    make_bad_pptx(os.path.join(OUT, "bad.pptx"))
    make_good_pptx(os.path.join(OUT, "good.pptx"))
    print("fixtures:", OUT)
