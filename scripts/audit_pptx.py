#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""PowerPoint 审计：排印（P）+ 书写（W，共用 text_rules.py）。

用法：
    python audit_pptx.py 汇报.pptx
    python audit_pptx.py 汇报.pptx --json
    python audit_pptx.py 汇报.pptx --slide 3      # 只审第 3 页

评估的是「这页能不能让人在 5 秒内读到主张」：字号够不够大、一页塞了几条、
有没有把稿子念出来、图表网格线有没有糊住数据。
"""
import argparse
import json
import sys

from text_rules import (CJK, Finding, SEV_CRITICAL, SEV_HINT,
                        check_writing as check_writing_shared)

MIN_BODY_PT = 16.0        # 幻灯片正文下限（投影/截图都要看得清）
MIN_TITLE_PT = 24.0       # 标题下限
MAX_BULLETS = 6           # 一页要点上限
MAX_CHARS = 120           # 一页非标题字数上限（超过就是在念稿）
MAX_SIZES = 5             # 一份幻灯片的字号档数上限


def _eff_size(run, para):
    for obj in (run, para):
        try:
            if obj is not None and obj.font is not None and obj.font.size:
                return obj.font.size.pt
        except Exception:
            continue
    return None


def _eff_name(run, para):
    for obj in (run, para):
        try:
            if obj is not None and obj.font is not None and obj.font.name:
                return obj.font.name
        except Exception:
            continue
    return None


def _spc(run):
    """字符间距，单位 1/100 pt（ooxml a:rPr@spc）。"""
    try:
        rPr = run._r.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
        if rPr is None:
            return 0
        return int(rPr.get("spc") or 0)
    except Exception:
        return 0


def _is_title(shape):
    try:
        if shape.is_placeholder:
            from pptx.enum.shapes import PP_PLACEHOLDER
            return shape.placeholder_format.type in (
                PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
    except Exception:
        pass
    return False


def audit_slide(slide, index, out):
    tag = f"S{index}"
    shapes = list(slide.shapes)
    title_shape = slide.shapes.title
    body_shapes = [s for s in shapes if s.has_text_frame and not _is_title(s)]

    # P14 缺标题
    if title_shape is None or not title_shape.text_frame.text.strip():
        out.append(Finding(tag, "P14", "这一页没有标题",
                           "每页一个主张，标题写结论不写名词", SEV_HINT))
    else:
        sizes = [s for p in title_shape.text_frame.paragraphs
                 for r in p.runs for s in [_eff_size(r, p)] if s]
        if sizes and min(sizes) < MIN_TITLE_PT:
            out.append(Finding(tag, "P4",
                               f"标题只有 {min(sizes):g}pt（< {MIN_TITLE_PT:g}）",
                               f"标题 ≥ {MIN_TITLE_PT:g}pt，与正文拉开级差", SEV_HINT))

    # P1 正文字号 / P2 要点数 / P3 字数
    body_sizes, bullets, chars = [], 0, 0
    for s in body_shapes:
        for p in s.text_frame.paragraphs:
            text = "".join(r.text for r in p.runs).strip()
            if not text:
                continue
            bullets += 1
            chars += len(text)
            for r in p.runs:
                sz = _eff_size(r, p)
                # 只把含中文的 run 当正文：全大写的 kicker、页码这类装饰小字不算
                if sz and CJK.search(r.text):
                    body_sizes.append(sz)
    if body_sizes and min(body_sizes) < MIN_BODY_PT:
        out.append(Finding(tag, "P1",
                           f"正文最小 {min(body_sizes):g}pt（< {MIN_BODY_PT:g}）",
                           f"正文 ≥ {MIN_BODY_PT:g}pt；放不下就删内容，不缩字",
                           SEV_CRITICAL))
    if bullets > MAX_BULLETS:
        out.append(Finding(tag, "P2", f"这一页 {bullets} 条要点（> {MAX_BULLETS}）",
                           "一页一个主张；拆页或把细节移进口头说明", SEV_CRITICAL))
    if chars > MAX_CHARS:
        out.append(Finding(tag, "P3", f"这一页正文 {chars} 字（> {MAX_CHARS}）",
                           "幻灯片是提示物，不是文档；长句移进备注", SEV_CRITICAL))

    # P15 层级过深
    for s in body_shapes:
        deep = [p for p in s.text_frame.paragraphs if p.level and p.level > 2]
        if deep:
            out.append(Finding(tag, "P15", f"列表层级到第 {max(p.level for p in deep) + 1} 级",
                               "最多两级；再深就是结构问题", SEV_HINT))
            break

    # P9 图表网格线 / P13 表格边框 / P12 图片替代文本
    for s in shapes:
        if getattr(s, "has_chart", False) and s.has_chart:
            try:
                if s.chart.value_axis.has_major_gridlines:
                    out.append(Finding(tag, "P9", "图表带水平网格线",
                                       "关掉网格线，直接标数值", SEV_HINT))
            except Exception:
                pass
        if getattr(s, "has_table", False) and s.has_table:
            xml = s._element.xml
            if "<a:lnL" in xml and "<a:lnR" in xml and "<a:lnT" in xml \
                    and "<a:lnB" in xml and len(s.table.rows) >= 3:
                out.append(Finding(tag, "P13", "表格四边都有框线",
                                   "只留表头下细线与表尾粗线", SEV_HINT))
        if getattr(s, "shape_type", None) is not None:
            from pptx.enum.shapes import MSO_SHAPE_TYPE
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    descr = s._element.nvPicPr.cNvPr.get("descr")
                except Exception:
                    descr = None
                if not descr:
                    out.append(Finding(tag, "P12", "图片没有替代文本",
                                       "加 alt 文本，投影与无障碍都需要", SEV_HINT))

    # P10 全靠普通文本框（没有大纲结构）
    text_shapes = [s for s in shapes if s.has_text_frame and s.text_frame.text.strip()]
    loose = [s for s in text_shapes if not s.is_placeholder]
    if text_shapes and len(loose) / len(text_shapes) > 0.6 and len(loose) >= 3:
        out.append(Finding(tag, "P10",
                           f"{len(loose)}/{len(text_shapes)} 个文本框不是占位符",
                           "用版式占位符，否则没有大纲、改母版也不生效", SEV_HINT))


def audit_deck(prs, out):
    # P11 画幅
    try:
        ratio = prs.slide_width / prs.slide_height
        if abs(ratio - 16 / 9) > 0.03:
            out.append(Finding("deck", "P11", f"画幅 {ratio:.2f}:1（不是 16:9）",
                               "16:9 是现代投影与屏幕的默认", SEV_HINT))
    except Exception:
        pass

    # P5 字号档数 / P6 字体族 / P7 中文斜体 / P8 全大写字距
    sizes, fams = set(), set()
    italic_cn, caps_no_spc = [], []
    for i, slide in enumerate(prs.slides, 1):
        for s in slide.shapes:
            if not s.has_text_frame:
                continue
            for p in s.text_frame.paragraphs:
                for r in p.runs:
                    if not r.text.strip():
                        continue
                    sz = _eff_size(r, p)
                    if sz:
                        sizes.add(round(sz * 2) / 2)
                    name = _eff_name(r, p)
                    if name:
                        fams.add(name)
                    if r.font.italic and CJK.search(r.text):
                        italic_cn.append(f"S{i}")
                    t = r.text.strip()
                    if t.isascii() and t.isupper() and len(t) >= 4 and _spc(r) <= 0:
                        caps_no_spc.append((f"S{i}", t))

    if len(sizes) > MAX_SIZES:
        out.append(Finding("deck", "P5",
                           f"全篇 {len(sizes)} 档字号：{sorted(sizes)}",
                           f"幻灯片的层级比文档少，收到 {MAX_SIZES} 档以内", SEV_HINT))
    if len(fams) >= 3:
        sev = SEV_CRITICAL if len(fams) >= 4 else SEV_HINT
        out.append(Finding("deck", "P6", f"{len(fams)} 个字体：{sorted(fams)[:5]}",
                           "收到 2 族（标题一、正文一）", sev))
    if italic_cn:
        out.append(Finding(italic_cn[0], "P7", f"中文斜体（{len(italic_cn)} 处）",
                           "取消斜体，用字重或颜色强调", SEV_CRITICAL))
    if caps_no_spc:
        loc, txt = caps_no_spc[0]
        out.append(Finding(loc, "P8", f"全大写「{txt}」无字距",
                           "加 0.06–0.1em（a:rPr@spc = 600~1000）", SEV_HINT))


def collect_text(prs):
    items, paragraphs, headings = [], [], []
    for i, slide in enumerate(prs.slides, 1):
        for j, s in enumerate(slide.shapes, 1):
            if not s.has_text_frame:
                continue
            loc = f"S{i}shape{j}"
            is_title = _is_title(s)
            for k, p in enumerate(s.text_frame.paragraphs):
                text = "".join(r.text for r in p.runs)
                if not text.strip():
                    continue
                items.append((loc, text))
                if len(text) > 8:
                    paragraphs.append((loc, text))
                if is_title and len(text.strip()) <= 8:
                    headings.append((loc, text))
            # 表格文字
        for j, s in enumerate(slide.shapes, 1):
            if getattr(s, "has_table", False) and s.has_table:
                for row in s.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            items.append((f"S{i}table{j}", cell.text))
        if slide.has_notes_slide:
            note = slide.notes_slide.notes_text_frame.text
            if note.strip():
                items.append((f"S{i}notes", note))
    return items, paragraphs, headings


def main():
    ap = argparse.ArgumentParser(description="PowerPoint 排印 + 书写审计")
    ap.add_argument("path")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--slide", type=int, help="只审第 N 页（从 1 开始）")
    ap.add_argument("--max", type=int, default=100)
    args = ap.parse_args()

    if not args.path.lower().endswith(".pptx"):
        print("只支持 .pptx；.ppt 请先另存为 .pptx", file=sys.stderr)
        return 2

    try:
        from pptx import Presentation
    except ImportError:
        print("需要 python-pptx：pip install python-pptx", file=sys.stderr)
        return 2

    prs = Presentation(args.path)
    slides = list(prs.slides)
    findings = []
    if args.slide:
        if not 1 <= args.slide <= len(slides):
            print(f"第 {args.slide} 页不存在（共 {len(slides)} 页）", file=sys.stderr)
            return 2
        audit_slide(slides[args.slide - 1], args.slide, findings)
    else:
        for i, slide in enumerate(slides, 1):
            audit_slide(slide, i, findings)
        audit_deck(prs, findings)

    items, paragraphs, headings = collect_text(prs)
    check_writing_shared(items, findings, paragraphs=paragraphs, headings=headings)

    order = {SEV_CRITICAL: 0, SEV_HINT: 1}
    findings.sort(key=lambda f: (order[f.sev], f.code))
    findings = findings[:args.max]

    if args.json:
        print(json.dumps(
            {"file": args.path, "kind": "pptx", "slides": len(slides),
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
