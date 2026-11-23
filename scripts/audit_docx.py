#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""企业文档审计：书写规范（W）+ 排印规范（T）。书写规则共用 text_rules.py。

用法：
    python audit_docx.py 报告.docx
    python audit_docx.py 报告.docx --json
    python audit_docx.py notes.md            # 纯文本只跑书写规则
    python audit_docx.py 报告.docx --max 200 # 最多报多少条

输出：`文件:位置 编号 问题 → 改法`，末尾一行统计。严重项 > 0 时退出码 1。
"""
import argparse
import json
import re
import sys

from text_rules import (CJK, NUM, Finding, SEV_CRITICAL, SEV_HINT,
                        check_writing as check_writing_shared)


# ── 文档抽取 ────────────────────────────────────────────────────
class Doc:
    def __init__(self):
        self.paragraphs = []       # [(loc, text, style, para_obj)]
        self.tables = []           # [(loc, [cell_text], table_obj)]
        self.section = None
        self.doc = None
        self.kind = "text"

    def all_text(self):
        out = [p[1] for p in self.paragraphs]
        for _, cells, _ in self.tables:
            out.extend(cells)
        return "\n".join(out)


def load_docx(path):
    from docx import Document

    d = Doc()
    d.kind = "docx"
    d.doc = Document(path)
    d.section = d.doc.sections[0]
    for i, p in enumerate(d.doc.paragraphs):
        d.paragraphs.append((f"p{i}", p.text, p.style.name if p.style else "", p))
    for ti, t in enumerate(d.doc.tables):
        cells = []
        for ri, row in enumerate(t.rows):
            for ci, cell in enumerate(row.cells):
                if cell.text.strip():
                    cells.append(cell.text.strip())
        d.tables.append((f"t{ti}", cells, t))
    return d


def load_text(path):
    d = Doc()
    d.kind = "text"
    with open(path, encoding="utf-8", errors="replace") as fh:
        for i, line in enumerate(fh, 1):
            if line.strip():
                d.paragraphs.append((f"L{i}", line.rstrip("\n"), "", None))
    return d


# ── 排印检查 ───────────────────────────────────────────────────
def eff_size(p, style_name=None, default=None):
    """段落有效字号：run 显式 > 段落样式 > Normal。"""
    for r in p.runs:
        if r.font.size:
            return r.font.size.pt
    try:
        st = p.style
        while st is not None:
            if st.font is not None and st.font.size:
                return st.font.size.pt
            st = st.base_style
    except Exception:
        pass
    return default


def leading_ratio(p, size):
    """行距倍数：段落 > 段落样式 > Normal 样式；固定值按 pt/字号 折算。"""
    from docx.enum.text import WD_LINE_SPACING

    pf = p.paragraph_format
    ls = pf.line_spacing
    rule = pf.line_spacing_rule
    if ls is None:
        try:
            st = p.style
            while st is not None and ls is None:
                ls = st.paragraph_format.line_spacing
                rule = st.paragraph_format.line_spacing_rule
                st = st.base_style
        except Exception:
            pass
    if ls is None:
        return None
    if rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
        return ls.pt / size if size else None
    return float(ls)


def body_size(paras):
    """正文基准字号：按「段落数」取众数，只在正文样式段落里统计。"""
    from collections import Counter

    counter = Counter()
    for _, text, style, p in paras:
        if style.startswith("Heading") or style in ("Title", "Subtitle"):
            continue
        if not CJK.search(text) or len(text) < 15 or p is None:
            continue
        sz = eff_size(p, style)
        if sz:
            counter[round(sz * 2) / 2] += 1
    return counter.most_common(1)[0][0] if counter else 11.0


def preceding_declares_unit(table):
    """表前一段（正文顺序）是否写了「单位」。"""
    from docx.text.paragraph import Paragraph

    prev = table._tbl.getprevious()
    while prev is not None:
        if prev.tag.endswith("}p"):
            text = Paragraph(prev, table._parent).text
            if text.strip():
                return "单位" in text
        elif prev.tag.endswith("}tbl"):
            return False
        prev = prev.getprevious()
    return False


def check_typography(d, out):
    if d.kind != "docx":
        return
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

    sec = d.section
    width_pt = (sec.page_width - sec.left_margin - sec.right_margin) / 12700.0
    size = body_size(d.paragraphs)
    cpl = width_pt / size if size else 0

    # T7 行长
    if CJK.search(d.all_text()):
        if cpl > 40:
            out.append(Finding("layout", "T7",
                               f"版心 {width_pt:.0f}pt ÷ 正文 {size:g}pt ≈ 一行 {cpl:.0f} 汉字",
                               f"收版心到 {size * 35:.0f}pt 左右（约一行 35 字），别缩字号",
                               SEV_CRITICAL))
        elif 0 < cpl < 20:
            out.append(Finding("layout", "T7",
                               f"一行只有 {cpl:.0f} 汉字，偏窄",
                               "略放宽版心或加大正文", SEV_HINT))

    # T1 / T4 / T8 / T11 / T15
    short_lead, big_lead, justify, indent, multi = [], [], [], [], []
    for loc, text, style, p in d.paragraphs:
        if not text.strip() or p is None:
            continue
        is_head = style.startswith("Heading") or style in ("Title", "Subtitle")
        psize = eff_size(p, style, size)
        ratio = leading_ratio(p, psize)
        if ratio is not None:
            # 大标题判定按字号，不只按样式：Normal 里手工敲的 26pt 标题也算
            if psize >= 18 and ratio < 1.3:
                big_lead.append((loc, psize, ratio))
            is_body_size = abs(psize - size) <= 0.6
            if not is_head and is_body_size and CJK.search(text) \
                    and ratio < 1.6 and len(text) > 10:
                short_lead.append((loc, psize, ratio))
        if p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            justify.append(loc)
        fli = p.paragraph_format.first_line_indent
        if fli and fli.pt > 0:
            indent.append(loc)
        if not is_head and CJK.search(text) and len(text) > 10:
            if p.paragraph_format.line_spacing_rule != WD_LINE_SPACING.EXACTLY:
                multi.append(loc)

    if short_lead:
        loc, ps, r = short_lead[0]
        out.append(Finding(loc, "T1",
                           f"中文正文 {ps:g}pt、行距 {r:.2f}（{len(short_lead)} 处不足）",
                           f"行距改固定值 {round(ps * 1.8)}pt（1.8 倍）", SEV_CRITICAL))
    if big_lead:
        loc, ps, r = big_lead[0]
        out.append(Finding(loc, "T4",
                           f"标题 {ps:g}pt、行距 {r:.2f}，中文会撞行（{len(big_lead)} 处）",
                           f"行距改固定值 {round(ps * 1.35)}pt（1.35 倍）", SEV_CRITICAL))
    if justify:
        out.append(Finding(justify[0], "T8",
                           f"两端对齐（{len(justify)} 处），中文会出现河沟",
                           "改左对齐", SEV_CRITICAL))
    if indent:
        out.append(Finding(indent[0], "T11",
                           f"首行缩进 {len(indent)} 处", "缩进归零，用段后间距分段",
                           SEV_HINT))
    if multi:
        out.append(Finding(multi[0], "T15",
                           f"{len(multi)} 段用多倍行距，跨渲染器会变",
                           "改固定值行距", SEV_HINT))

    # T3 字号档数
    sizes = set()
    for _, _, _, p in d.paragraphs:
        if p is None:
            continue
        for r in p.runs:
            if r.font.size and r.text.strip():
                sizes.add(round(r.font.size.pt * 2) / 2)
    if len(sizes) > 7:
        out.append(Finding("all", "T3",
                           f"全篇 {len(sizes)} 档字号：{sorted(sizes)}",
                           "收到 7 档以内（建议 5 档）+1", SEV_CRITICAL))

    # T5 / T6 / T14 字符级
    neg, caps, italic = [], [], []
    for loc, text, style, p in d.paragraphs:
        if p is None:
            continue
        for r in p.runs:
            if not r.text.strip():
                continue
            rPr = r._element.rPr
            spacing = rPr.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing") \
                if rPr is not None else None
            val = int(spacing.get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")) \
                if spacing is not None else 0
            if val < 0 and CJK.search(r.text):
                neg.append(loc)
            if r.text.strip().isascii() and r.text.strip().isupper() \
                    and len(r.text.strip()) >= 4 and val <= 0:
                caps.append((loc, r.text.strip()))
            if r.italic and CJK.search(r.text):
                italic.append(loc)
    if neg:
        out.append(Finding(neg[0], "T5", f"中文套了负字距（{len(neg)} 处）",
                           "中文一律 0，负字距只给拉丁", SEV_CRITICAL))
    if caps:
        loc, txt = caps[0]
        out.append(Finding(loc, "T6", f"全大写「{txt}」无正字距",
                           "加 0.08em（9pt 字 = 0.72pt）", SEV_HINT))
    if italic:
        out.append(Finding(italic[0], "T14", f"中文斜体（{len(italic)} 处）",
                           "取消斜体，用字重或颜色强调", SEV_CRITICAL))

    # T9 / T10 / T16 表格
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for loc, cells, t in d.tables:
        tblPr = t._tbl.tblPr
        borders = tblPr.find(W + "tblBorders") if tblPr is not None else None
        gridded = False
        if borders is not None:
            for edge in ("insideH", "insideV", "top", "left", "bottom", "right"):
                el = borders.find(W + edge)
                if el is not None and el.get(W + "val") not in ("nil", "none"):
                    gridded = True
        elif t.style is not None and "Grid" in (t.style.name or ""):
            # 没有显式 tblBorders 时，Table Grid 样式才会真的画出网格
            gridded = True
        if gridded and len(t.rows) >= 2:
            out.append(Finding(loc, "T9", "表格画了网格线",
                               "关掉 insideH/insideV 与外框，只留表头细线 + 表尾粗线",
                               SEV_CRITICAL))
        if not re.search(r"单位", " ".join(cells)) and len(t.rows) >= 2:
            numeric = sum(1 for c in cells if NUM.match(c.replace(" ", "")))
            if numeric >= 3 and not preceding_declares_unit(t):
                out.append(Finding(loc, "T16", "表格没有单位声明",
                                   "表头或表注写「金额单位：万元」", SEV_HINT))
        for ci in range(len(t.columns)):
            col = [t.cell(ri, ci) for ri in range(len(t.rows))]
            nums = [c for c in col[1:] if NUM.match(c.text.strip().replace(" ", ""))]
            if len(nums) >= 3:
                lefts = [c for c in nums
                         if c.paragraphs[0].alignment in
                         (None, WD_ALIGN_PARAGRAPH.LEFT)]
                if lefts:
                    out.append(Finding(f"{loc}c{ci}", "T10",
                                       f"数字列左对齐（{len(lefts)} 格）",
                                       "数字右对齐，位数才能对齐", SEV_HINT))
                    break

    # T12 字体族（run 显式 + 样式层）
    fams = set()
    for p in d.doc.paragraphs:
        for r in p.runs:
            if r.font.name:
                fams.add(r.font.name)
        try:
            st = p.style
            while st is not None:
                if st.font is not None and st.font.name:
                    fams.add(st.font.name)
                st = st.base_style
        except Exception:
            pass
    for t in d.doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.font.name:
                            fams.add(r.font.name)
    if len(fams) >= 3:
        sev = SEV_CRITICAL if len(fams) >= 4 else SEV_HINT
        out.append(Finding("all", "T12",
                           f"{len(fams)} 个字体：{sorted(fams)[:6]}",
                           "收到 2 族（标题族 + 正文族）；内置样式自带的字体要改掉",
                           sev))

    # T13 页码
    has_page = any("PAGE" in p._p.xml for p in sec.footer.paragraphs)
    if not has_page:
        out.append(Finding("footer", "T13", "页脚没有页码", "加「第 x 页 / 共 y 页」",
                           SEV_HINT))


# ── 书写检查（转交 text_rules） ─────────────────────────────────
def check_writing(d, out):
    items = [(loc, text) for loc, text, _, _ in d.paragraphs]
    for loc, cells, _ in d.tables:
        for i, c in enumerate(cells):
            items.append((f"{loc}#{i}", c))

    paragraphs = []
    headings = []
    for loc, text, style, p in d.paragraphs:
        is_head = style.startswith("Heading") or style in ("Title", "Subtitle")
        if not is_head:
            paragraphs.append((loc, text))
        elif text.strip() and len(text.strip()) <= 8:
            headings.append((loc, text))
        if text.strip() and len(text.strip()) <= 8 and not is_head \
                and p is not None and p.runs and p.runs[0].font.size \
                and p.runs[0].font.size.pt >= 14:
            headings.append((loc, text))

    check_writing_shared(items, out, paragraphs=paragraphs, headings=headings)


# ── 主流程 ─────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="企业文档书写 + 排印审计")
    ap.add_argument("path")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--max", type=int, default=100)
    args = ap.parse_args()

    is_docx = args.path.lower().endswith(".docx")
    ext = args.path.lower().rsplit(".", 1)[-1] if "." in args.path else ""
    if not is_docx and ext not in ("md", "mdx", "txt", "markdown"):
        print(f"只支持 .docx / .md / .txt（收到 .{ext}）；"
              f"PDF 请先转成 docx，或导出为 .md 再审计；"
              f"Excel 请用 audit_xlsx.py", file=sys.stderr)
        return 2
    if is_docx:
        try:
            d = load_docx(args.path)
        except ImportError:
            print("需要 python-docx：pip install python-docx", file=sys.stderr)
            return 2
        except Exception as exc:
            print(f"打不开 docx：{exc}", file=sys.stderr)
            return 2
    else:
        d = load_text(args.path)

    findings = []
    check_typography(d, findings)
    check_writing(d, findings)

    order = {SEV_CRITICAL: 0, SEV_HINT: 1}
    findings.sort(key=lambda f: (order[f.sev], f.code))
    findings = findings[:args.max]

    if args.json:
        print(json.dumps(
            {"file": args.path, "kind": d.kind,
             "critical": sum(1 for f in findings if f.sev == SEV_CRITICAL),
             "hint": sum(1 for f in findings if f.sev == SEV_HINT),
             "findings": [f.as_dict() for f in findings]},
            ensure_ascii=False, indent=2))
    else:
        name = args.path.replace("\\", "/").split("/")[-1]
        for f in findings:
            print(f.line(name))
        c = sum(1 for f in findings if f.sev == SEV_CRITICAL)
        h = sum(1 for f in findings if f.sev == SEV_HINT)
        print(f"\n严重 {c} 条 / 提示 {h} 条"
              f"{'　（严重项需清零）' if c else '　✓ 可以交付'}")

    return 1 if any(f.sev == SEV_CRITICAL for f in findings) else 0


if __name__ == "__main__":
    sys.exit(main())
