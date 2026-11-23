#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""按企业版式重排 .docx 的样式层：字体、字号、固定行距、段距、对齐、表格线、页码。

只动样式与段落格式，不改一个字。文字的问题交给 references/writing.md 和 audit_docx.py。

用法：
    python restyle_docx.py 报告.docx --style board
    python restyle_docx.py 报告.docx --style memo --out 报告-排版版.docx
    python restyle_docx.py 报告.docx --style board --inplace   # 先自己备份
"""
import argparse
import os
import sys

PRESETS = {
    # 名字: 页边距(上,下,左,右) / 各角色(字号pt, 行距pt, 段前pt, 段后pt, 字体)
    "board": {
        "margins": (2.5, 2.4, 3.0, 3.0),
        "display": (30, 41, 24, 14, "serif"),
        "h1": (18, 25, 24, 8, "serif"),
        "h2": (14, 20, 16, 6, "serif"),
        "h3": (12, 17, 12, 4, "sans"),
        "body": (11.5, 21, 0, 8, "sans"),
        "small": (10, 15, 0, 4, "sans"),
        "label": (9, 14, 0, 0, "sans"),
    },
    "whitepaper": {
        "margins": (2.5, 2.4, 2.8, 2.8),
        "display": (34, 46, 28, 16, "serif"),
        "h1": (20, 27, 28, 10, "serif"),
        "h2": (14, 20, 18, 6, "serif"),
        "h3": (12, 17, 14, 4, "sans"),
        "body": (11, 20, 0, 8, "sans"),
        "small": (9.5, 14, 0, 4, "sans"),
        "label": (9, 14, 0, 0, "sans"),
    },
    "memo": {
        "margins": (2.4, 2.2, 2.6, 2.6),
        "display": (20, 28, 0, 10, "serif"),
        "h1": (13, 18, 18, 6, "sans"),
        "h2": (12, 17, 14, 4, "sans"),
        "h3": (11, 16, 12, 4, "sans"),
        "body": (10.5, 19, 0, 6, "sans"),
        "small": (9, 14, 0, 4, "sans"),
        "label": (8.5, 13, 0, 0, "sans"),
    },
}

SERIF_EA, SERIF_LAT = "宋体", "Times New Roman"
SANS_EA, SANS_LAT = "微软雅黑", "Segoe UI"
INK = "111111"
MUTED = "555555"
HAIR = "CBC9C3"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
             "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
             "noProof", "snapToGrid", "vanish", "webHidden", "color",
             "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
             "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
               "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
               "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
               "tblCellMar", "tblLook", "tblCaption", "tblDescription"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
              "shd", "noWrap", "tcMar", "textDirection", "tcFitText",
              "vAlign", "hideMark"]


def insert_in_order(parent, element, order):
    from docx.oxml import OxmlElement  # noqa: F401
    tag = element.tag.split("}")[1]
    idx = order.index(tag)
    for child in list(parent):
        ctag = child.tag.split("}")[1]
        if ctag in order and order.index(ctag) > idx:
            child.addprevious(element)
            return element
    parent.append(element)
    return element


def apply_run_font(run, size, role_font):
    from docx.shared import Pt, RGBColor
    from docx.oxml import OxmlElement

    run.font.size = Pt(size)
    ea, lat = (SERIF_EA, SERIF_LAT) if role_font == "serif" else (SANS_EA, SANS_LAT)
    run.font.name = lat
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(W + "rFonts")
    if rFonts is None:
        rFonts = insert_in_order(rPr, OxmlElement("w:rFonts"), RPR_ORDER)
    rFonts.set(W + "ascii", lat)
    rFonts.set(W + "hAnsi", lat)
    rFonts.set(W + "eastAsia", ea)
    run.italic = False                      # 中文不用斜体
    run.font.color.rgb = RGBColor.from_string(INK)


def style_role(style_name):
    n = (style_name or "").strip().lower()
    if n in ("title",):
        return "display"
    if n.startswith("heading 1"):
        return "h1"
    if n.startswith("heading 2"):
        return "h2"
    if n.startswith("heading 3") or n.startswith("heading 4"):
        return "h3"
    return "body"


def restyle_style(style, spec):
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.shared import Pt, RGBColor
    from docx.oxml import OxmlElement

    size, lead, before, after, fam = spec
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor.from_string(INK)
    ea, lat = (SERIF_EA, SERIF_LAT) if fam == "serif" else (SANS_EA, SANS_LAT)
    style.font.name = lat
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(W + "rFonts")
    if rFonts is None:
        rFonts = insert_in_order(rPr, OxmlElement("w:rFonts"), RPR_ORDER)
    rFonts.set(W + "ascii", lat)
    rFonts.set(W + "hAnsi", lat)
    rFonts.set(W + "eastAsia", ea)
    pf = style.paragraph_format
    pf.line_spacing = Pt(lead)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.first_line_indent = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT


def style_table(table, label_size):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor
    from docx.oxml import OxmlElement
    import re

    tblPr = table._tbl.tblPr
    old = tblPr.find(W + "tblBorders")
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(W + "val", "nil")
        borders.append(el)
    insert_in_order(tblPr, borders, TBLPR_ORDER)

    num = re.compile(r"^[+\-−]?\s*[\d,]+(\.\d+)?\s*(%|pt)?$")
    rows = table.rows
    for ri, row in enumerate(rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(W + "cantSplit") is None:
            trPr.append(OxmlElement("w:cantSplit"))
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            old = tcPr.find(W + "tcBorders")
            if old is not None:
                tcPr.remove(old)
            tcB = OxmlElement("w:tcBorders")
            for edge in ("top", "left", "bottom", "right"):
                el = OxmlElement(f"w:{edge}")
                el.set(W + "val", "nil")
                tcB.append(el)
            insert_in_order(tcPr, tcB, TCPR_ORDER)

            is_num = bool(num.match(cell.text.strip().replace(" ", ""))) \
                and not cell.text.strip().endswith("年")
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.first_line_indent = Pt(0)
                p.alignment = (WD_ALIGN_PARAGRAPH.RIGHT if is_num
                               else WD_ALIGN_PARAGRAPH.LEFT)
                for r in p.runs:
                    r.italic = False
                    if ri == 0:
                        r.font.size = Pt(label_size)
                        r.bold = True
                        r.font.color.rgb = RGBColor.from_string(MUTED)

            if ri == 0:
                el = tcB.find(W + "bottom")
                el.set(W + "val", "single"); el.set(W + "sz", "6")
                el.set(W + "color", INK)
            elif ri == len(rows) - 1:
                el = tcB.find(W + "bottom")
                el.set(W + "val", "single"); el.set(W + "sz", "12")
                el.set(W + "color", INK)
            else:
                el = tcB.find(W + "bottom")
                el.set(W + "val", "single"); el.set(W + "sz", "4")
                el.set(W + "color", HAIR)


def add_page_number_footer(section, label_size):
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.shared import Pt, RGBColor
    from docx.oxml import OxmlElement

    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    for child in list(p._p):
        if child.tag == W + "r" or child.tag == W + "fldSimple":
            p._p.remove(child)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(label_size + 5)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY

    def run(text):
        r = p.add_run(text)
        r.font.size = Pt(label_size)
        r.font.color.rgb = RGBColor.from_string(MUTED)
        r.font.name = SANS_LAT
        rPr = r._element.get_or_add_rPr()
        rFonts = rPr.find(W + "rFonts")
        if rFonts is None:
            rFonts = insert_in_order(rPr, OxmlElement("w:rFonts"), RPR_ORDER)
        rFonts.set(W + "eastAsia", SANS_EA)
        return r

    def field(instr, placeholder):
        fld = OxmlElement("w:fldSimple")
        fld.set(W + "instr", instr)
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz"); sz.set(W + "val", str(int(label_size * 2)))
        col = OxmlElement("w:color"); col.set(W + "val", MUTED)
        rPr.append(col); rPr.append(sz)
        t = OxmlElement("w:t"); t.text = placeholder
        r.append(rPr); r.append(t)
        fld.append(r)
        p._p.append(fld)

    run("第 ")
    field(" PAGE ", "1")
    run(" 页 / 共 ")
    field(" NUMPAGES ", "1")
    run(" 页")


def main():
    ap = argparse.ArgumentParser(description="按企业版式重排 docx 样式层")
    ap.add_argument("path")
    ap.add_argument("--style", default="board", choices=sorted(PRESETS))
    ap.add_argument("--out")
    ap.add_argument("--inplace", action="store_true")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    try:
        from docx import Document
        from docx.shared import Cm, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    except ImportError:
        print("需要 python-docx：pip install python-docx", file=sys.stderr)
        return 2

    preset = PRESETS[args.style]
    doc = Document(args.path)

    # 页面
    for sec in doc.sections:
        t, b, l, r = preset["margins"]
        sec.top_margin, sec.bottom_margin = Cm(t), Cm(b)
        sec.left_margin, sec.right_margin = Cm(l), Cm(r)
        add_page_number_footer(sec, preset["label"][0])

    # 样式
    wanted = {"Normal": "body", "Title": "display", "Subtitle": "display",
              "Heading 1": "h1", "Heading 2": "h2",
              "Heading 3": "h3", "Heading 4": "h3", "Quote": "body"}
    touched = []
    for name, role in wanted.items():
        try:
            sty = doc.styles[name]
        except KeyError:
            continue
        restyle_style(sty, preset[role][:4] + (preset[role][4],))
        # 内置 Title/Heading 样式自带段落下边框（那条蓝线），删掉
        pPr = sty.element.find(W + "pPr")
        if pPr is not None:
            pBdr = pPr.find(W + "pBdr")
            if pBdr is not None:
                pPr.remove(pBdr)
        touched.append(name)
    # 标题字重：display/h1/h2 加粗，body 常规
    for name in ("Title", "Heading 1", "Heading 2", "Heading 3", "Heading 4"):
        try:
            doc.styles[name].font.bold = True
        except KeyError:
            pass
    try:
        doc.styles["Normal"].font.bold = False
    except KeyError:
        pass

    # 段落与 run
    n_body = n_head = 0
    for p in doc.paragraphs:
        role = style_role(p.style.name if p.style else "")
        size, lead, before, after, fam = preset[role]
        pf = p.paragraph_format
        pf.line_spacing = Pt(lead)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.first_line_indent = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if role.startswith("h") or role == "display":
            pf.keep_with_next = True
            n_head += 1
        else:
            n_body += 1
        for r in p.runs:
            apply_run_font(r, size, fam)
            if role != "body" and r.text.strip():
                r.bold = True          # 标题字重交给样式层，run 也要顶上
            # 全大写拉丁加正字距（0.08em → 1/20 pt）
            if r.text.strip() and r.text.strip().isascii() \
                    and r.text.strip().isupper() and len(r.text.strip()) >= 4:
                from docx.oxml import OxmlElement as _OE
                rPr = r._element.get_or_add_rPr()
                old = rPr.find(W + "spacing")
                if old is not None:
                    rPr.remove(old)
                sp = _OE("w:spacing")
                sp.set(W + "val", str(int(round(0.08 * size * 20))))
                insert_in_order(rPr, sp, RPR_ORDER)

    for t in doc.tables:
        style_table(t, preset["label"][0])

    out = args.path
    if args.inplace:
        root, _ = os.path.splitext(args.path)
        backup = root + "-备份.docx"
        os.replace(args.path, backup) if not os.path.exists(backup) else None
        out = args.path
    elif args.out:
        out = args.out
    else:
        root, _ = os.path.splitext(args.path)
        out = root + "-排版版.docx"

    doc.save(out)
    if not args.quiet:
        print(f"样式已重排：{args.style}　正文段 {n_body} / 标题段 {n_head} / "
              f"表格 {len(doc.tables)}")
        print(f"样式：{', '.join(touched)}")
        print(f"输出：{out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
