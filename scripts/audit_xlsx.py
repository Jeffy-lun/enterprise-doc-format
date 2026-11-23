#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Excel 工作簿审计：排印（X）+ 书写（W，共用 text_rules.py）。

用法：
    python audit_xlsx.py 报表.xlsx
    python audit_xlsx.py 报表.xlsx --json
    python audit_xlsx.py 报表.xlsx --sheet "利润表"   # 只审一个表

企业报表的通病是「网格线糊成一片、数字没格式、列宽截断、比率存成小数」，
这四件事 X1/X2/X4/X13 各自都能单独毁掉一张表的可读性。
"""
import argparse
import json
import re
import sys

from text_rules import (CJK, NUM, Finding, SEV_CRITICAL, SEV_HINT,
                        check_writing as check_writing_shared)

DEFAULT_COL_WIDTH = 8.43      # Excel 默认列宽（字符）
DEFAULT_ROW_HEIGHT = 15.0     # Excel 默认行高（pt）


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_text_number(v):
    return isinstance(v, str) and bool(re.fullmatch(
        r"[-+]?[\d,]+(\.\d+)?%?", v.strip())) and any(ch.isdigit() for ch in v)


def _span_cols(ws, cell):
    """该单元格横向跨了几列（合并区域的主格才会 > 1）。"""
    try:
        for rng in ws.merged_cells.ranges:
            if cell.coordinate == rng.coord.split(":")[0]:
                return rng.max_col - rng.min_col + 1
    except Exception:
        pass
    return 1


def audit_sheet(ws, out, only_numeric_rows=True):
    from openpyxl.cell.cell import MergedCell

    name = ws.title
    rows = list(ws.iter_rows())
    if not rows:
        return

    # 收集统计
    numeric_cells, text_cells = [], []
    fams, sizes = set(), set()
    n_merged = len(ws.merged_cells.ranges)
    n_cells = sum(len(r) for r in rows)
    border4 = 0
    left_nums = []

    for row in rows:
        for c in row:
            if isinstance(c, MergedCell) or c.value is None:
                continue
            if _is_number(c.value):
                numeric_cells.append(c)
                if c.alignment is not None and c.alignment.horizontal == "left":
                    left_nums.append(c)
            elif isinstance(c.value, str):
                if c.value.strip():
                    text_cells.append(c)
            if c.font and c.font.name:
                fams.add(c.font.name)
            if c.font and c.font.size:
                sizes.add(round(c.font.size, 1))
            b = c.border
            if b and all(getattr(b, e) is not None and getattr(b, e).style
                         for e in ("left", "right", "top", "bottom")):
                border4 += 1

    # X1 网格线
    grid_off = ws.sheet_view.showGridLines is False
    if not grid_off and len(numeric_cells) >= 10:
        out.append(Finding(f"{name}!", "X1",
                           f"网格线未关闭（{len(numeric_cells)} 个数值单元格被格子切开）",
                           "sheet_view.showGridLines = False，只用表头细线 + 表尾粗线",
                           SEV_CRITICAL))

    # X2 数字格式
    general = [c for c in numeric_cells if (c.number_format or "General") == "General"]
    big = [c for c in general if abs(float(c.value)) >= 1000]
    ratio = [c for c in general if 0 < abs(float(c.value)) < 1]
    if big:
        out.append(Finding(f"{name}!{big[0].coordinate}", "X2",
                           f"{len(big)} 个数值是 General 格式（如 {big[0].value}）",
                           "金额用 #,##0，比率用 0.0%，变动用 +0.0%;-0.0%",
                           SEV_CRITICAL))
    elif ratio:
        out.append(Finding(f"{name}!{ratio[0].coordinate}", "X2",
                           f"{len(ratio)} 个比率是 General（如 {ratio[0].value}，读者要自己乘 100）",
                           "设百分比格式 0.0%", SEV_CRITICAL))

    # X3 数字左对齐
    if len(left_nums) >= 3:
        out.append(Finding(f"{name}!{left_nums[0].coordinate}", "X3",
                           f"{len(left_nums)} 个数字左对齐",
                           "数字右对齐，位数才能竖着对", SEV_HINT))

    # X4 列宽
    for col_letter, dim in ws.column_dimensions.items():
        if dim.width is None:
            continue
        need = 0
        for row in rows:
            for c in row:
                if isinstance(c, MergedCell):
                    continue
                if c.column_letter == col_letter and c.value is not None:
                    if _span_cols(ws, c) > 1:
                        continue          # 横向合并的标题格不代表列宽需求
                    txt = str(c.value)
                    right = ws.cell(row=c.row, column=c.column + 1)
                    if right.value is None:
                        continue          # 右侧为空时文本会溢出显示，不算截断
                    n = sum(2 if CJK.search(ch) else 1 for ch in txt)
                    need = max(need, n)
        if need and need > dim.width * 1.05 + 1:
            out.append(Finding(f"{name}!{col_letter}1", "X4",
                               f"列宽 {dim.width:.1f} 放不下最长内容（约需 {need * 0.95:.1f}）",
                               "加宽列或允许换行，别把内容截成 ####", SEV_CRITICAL))
            break

    # X5 行高
    small_rows = []
    for idx in range(len(rows)):
        r_dim = ws.row_dimensions.get(idx + 1)
        h = r_dim.height if r_dim is not None and r_dim.height else DEFAULT_ROW_HEIGHT
        sizes_in_row = [c.font.size for c in rows[idx]
                        if c.value is not None and c.font and c.font.size]
        if sizes_in_row and max(sizes_in_row) >= 12 and h < max(sizes_in_row) * 1.2:
            small_rows.append(idx + 1)
    if small_rows:
        out.append(Finding(f"{name}!A{small_rows[0]}", "X5",
                           f"{len(small_rows)} 行行高不够（大字号被压）",
                           "表头 22–24pt、数据行 19–21pt", SEV_HINT))

    # X6 字体族
    if len(fams) >= 3:
        sev = SEV_CRITICAL if len(fams) >= 4 else SEV_HINT
        out.append(Finding(f"{name}!", "X6", f"{len(fams)} 个字体：{sorted(fams)[:5]}",
                           "收到 2 族（中文一、数字一）", sev))

    # X7 字号
    if len(sizes) > 7:
        out.append(Finding(f"{name}!", "X7", f"{len(sizes)} 档字号：{sorted(sizes)}",
                           "收到 7 档以内", SEV_HINT))
    tiny = [s for s in sizes if s < 9]
    if tiny:
        out.append(Finding(f"{name}!", "X7", f"有 {tiny} pt 的小字",
                           "正文不小于 9pt，打印会糊", SEV_HINT))

    # X8 全网格边框
    if n_cells and border4 / n_cells > 0.3 and border4 >= 8:
        out.append(Finding(f"{name}!", "X8",
                           f"{border4}/{n_cells} 个单元格四边都有框",
                           "只留表头下细线 + 表尾粗线，行间用极浅细线或不用",
                           SEV_CRITICAL))

    # X9 表头无区分
    head = [c for c in rows[0] if c.value is not None]
    if head:
        bold = sum(1 for c in head if c.font and c.font.bold)
        if bold == 0 and len(rows) >= 3:
            out.append(Finding(f"{name}!{head[0].coordinate}", "X9",
                               "表头没加粗、与数据同字重",
                               "表头加粗 + 9pt 灰，或用一条细线收口", SEV_HINT))

    # X10 冻结
    if len(rows) > 20 and not ws.freeze_panes:
        out.append(Finding(f"{name}!", "X10", f"{len(rows)} 行数据但没有冻结表头",
                           "ws.freeze_panes 到数据首行", SEV_HINT))

    # X11 打印设置
    if len(numeric_cells) >= 10:
        if not ws.print_area:
            out.append(Finding(f"{name}!", "X11", "没有设置打印区域",
                               "设 print_area 与 fitToWidth=1，否则打印会横向断页",
                               SEV_HINT))
        elif not ws.page_setup.fitToWidth:
            out.append(Finding(f"{name}!", "X11", "打印未按宽度缩放",
                               "page_setup.fitToWidth = 1", SEV_HINT))

    # X12 同列类型混杂：只有「数字中间夹文字」才算坏，表头/表注那种上下结构不算
    ncol = max(len(r) for r in rows)
    for ci in range(ncol):
        seq = []
        for ri in range(len(rows)):
            if ci >= len(rows[ri]):
                continue
            c = rows[ri][ci]
            if isinstance(c, MergedCell) or c.value is None:
                continue
            seq.append(c)
        nums = [c for c in seq if _is_number(c.value)]
        if len(nums) < 3:
            continue
        first_num = next(i for i, c in enumerate(seq) if _is_number(c.value))
        interleaved = [c for c in seq[first_num:]
                       if isinstance(c.value, str) and c.value.strip()
                       and not _is_text_number(c.value)]
        if interleaved:
            out.append(Finding(f"{name}!{interleaved[0].coordinate}", "X12",
                               f"数字列中间夹了 {len(interleaved)} 个文字（求和会被跳过）",
                               "文字备注挪出数据列，或另起一行注", SEV_CRITICAL))
            break

    # X13 文本型数字 / 比率存错
    textnums = [c for c in text_cells if _is_text_number(c.value)]
    if len(textnums) >= 2:
        out.append(Finding(f"{name}!{textnums[0].coordinate}", "X13",
                           f"{len(textnums)} 个数字存成了文本（无法参与计算）",
                           "改成数值单元格 + 数字格式", SEV_CRITICAL))
    # 只有「纯百分比格式 + 值超过 1.5」才说明是百分点没除 100；
    # 带正负号或文字的格式（如 +0.0%;-0.0%）是同比变动，733% 是合法值。
    plain_pct = {"0%", "0.0%", "0.00%"}
    pct_bad = [c for c in numeric_cells
               if (c.number_format or "").strip() in plain_pct
               and abs(float(c.value)) > 1.5]
    if pct_bad:
        out.append(Finding(f"{name}!{pct_bad[0].coordinate}", "X13",
                           f"{len(pct_bad)} 个单元格是百分点值配了百分比格式（会显示成 140%）",
                           "值除以 100，或改用 +0.0\" pt\" 格式", SEV_CRITICAL))

    # X14 合并单元格
    if n_cells and n_merged >= 5 and n_merged * 12 > n_cells:
        out.append(Finding(f"{name}!", "X14", f"{n_merged} 处合并单元格",
                           "合并会破坏排序与取数，标题区少量可以", SEV_HINT))

    # X15 单位声明
    flat = " ".join(c.value for c in text_cells if isinstance(c.value, str))
    if len(numeric_cells) >= 5 and not re.search(r"单位", flat):
        out.append(Finding(f"{name}!", "X15", "没有单位声明",
                           "表头或表注写「金额单位：万元」", SEV_HINT))


def check_writing(ws_list, out):
    """把单元格文本喂给共用的书写规则。"""
    items, paragraphs, headings = [], [], []
    for ws in ws_list:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.strip():
                    loc = f"{ws.title}!{c.coordinate}"
                    items.append((loc, v))
                    if len(v) > 8:
                        paragraphs.append((loc, v))
                    if len(v.strip()) <= 8 and c.font and c.font.bold:
                        headings.append((loc, v))
    if items:
        check_writing_shared(items, out, paragraphs=paragraphs, headings=headings)


def main():
    ap = argparse.ArgumentParser(description="Excel 工作簿排印 + 书写审计")
    ap.add_argument("path")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--sheet", help="只审指定工作表")
    ap.add_argument("--max", type=int, default=100)
    args = ap.parse_args()

    if not args.path.lower().endswith((".xlsx", ".xlsm")):
        print("只支持 .xlsx / .xlsm；.xls 请先另存为 .xlsx", file=sys.stderr)
        return 2

    try:
        from openpyxl import load_workbook
    except ImportError:
        print("需要 openpyxl：pip install openpyxl", file=sys.stderr)
        return 2

    wb = load_workbook(args.path, data_only=False)
    sheets = [wb[args.sheet]] if args.sheet and args.sheet in wb.sheetnames \
        else [wb[n] for n in wb.sheetnames]

    findings = []
    for ws in sheets:
        audit_sheet(ws, findings)
    check_writing(sheets, findings)

    order = {SEV_CRITICAL: 0, SEV_HINT: 1}
    findings.sort(key=lambda f: (order[f.sev], f.code))
    findings = findings[:args.max]

    if args.json:
        print(json.dumps(
            {"file": args.path, "kind": "xlsx", "sheets": [s.title for s in sheets],
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
