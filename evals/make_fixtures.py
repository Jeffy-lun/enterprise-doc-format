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


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    make_bad(os.path.join(OUT, "bad.docx"))
    make_good(os.path.join(OUT, "good.docx"))
    print("fixtures:", OUT)
