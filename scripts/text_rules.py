#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""书写规范（W 编号）的共用实现。

docx 与 xlsx 审计都走这里，避免规则分叉。调用方把「可检查的文本」整理成
`[(位置, 文本), ...]` 传进来，位置用文件内的定位串（如 p12 / t0r1c2 / Sheet!A1）。
"""
import re

CJK = re.compile(r"[\u4e00-\u9fff]")
NUM = re.compile(r"^[+\-−]?\s*[\d,]+(\.\d+)?\s*%?$")
UNIT_GLUE = re.compile(
    r"\d(亿元|万元|万台|万套|万个|万人|万吨|个月|个百分点|条|台|套|人|元|吨|公里|平方米)")
PCT_SPACE = re.compile(r"\d\s+%")
LATIN_GLUE = re.compile(r"[\u4e00-\u9fff][A-Za-z]|[A-Za-z][\u4e00-\u9fff]")
DOT3 = re.compile(r"\.\.\.")
DASH = re.compile(r"[—–]|(?<!\d)--(?!\d)")
STRAIGHT_Q = re.compile(r"[\"']")

BANNED = ["赋能", "闭环", "抓手", "一站式", "极致", "大幅", "显著", "轻松",
          "一键", "完美", "领先", "全面提升", "深度赋能", "助力", "无缝"]
ABSOLUTE = ["唯一", "最好", "第一名", "彻底", "零缺陷"]
VAGUE = ["较多", "较少", "较快", "较慢", "明显", "左右", "大约", "略有"]
GENERIC_HEAD = ["概览", "概述", "结语", "小结", "其他", "更多", "详情", "备注",
                "综述", "简介"]

SEV_CRITICAL, SEV_HINT = "严重", "提示"


class Finding:
    def __init__(self, loc, code, msg, fix, sev):
        self.loc, self.code, self.msg, self.fix, self.sev = loc, code, msg, fix, sev

    def line(self, fname):
        return f"{fname}:{self.loc} {self.code} {self.msg} → {self.fix}"

    def as_dict(self):
        return {"loc": self.loc, "code": self.code, "issue": self.msg,
                "fix": self.fix, "severity": self.sev}


def _find(items, rep, code, msg, fix, sev, out, limit=1):
    n = 0
    first = None
    for loc, text in items:
        for m in rep.finditer(text):
            n += 1
            if first is None:
                first = (loc, text[max(0, m.start() - 8):m.start() + 12])
            if limit and n >= limit:
                break
        if limit and n >= limit:
            break
    if n:
        out.append(Finding(first[0], code, f"{msg}（{n} 处）",
                           f"{fix}　例：「{first[1]}」", sev))
    return n


def _merge_by_code(out, code, n_words, head):
    hits = [f for f in out if f.code == code]
    if len(hits) > 1:
        keep = hits[0]
        keep.msg = f"{head}{len(hits)} 种：" + "、".join(
            f.msg.split("「")[1].split("」")[0] for f in hits)
        for f in hits[1:]:
            out.remove(f)


def check_writing(items, out, *, paragraphs=None, headings=None,
                  full_text=None):
    """items: [(loc, text)] —— 所有可检查文本（含表格单元格）。
    paragraphs: 段落列表（用于 W9 段落过长），元素同 items。
    headings: [(loc, text)] —— 标题类文本（用于 W10 泛化标题）。
    full_text: 未给则用 items 拼。
    """
    full = full_text if full_text is not None else "\n".join(t for _, t in items)

    _find(items, DASH, "W1", "用破折号当标点", "改冒号/逗号/句号，或重写",
          SEV_CRITICAL, out)
    _find(items, STRAIGHT_Q, "W2", "直引号", "中文用「」或“”", SEV_CRITICAL, out)
    _find(items, DOT3, "W3", "三个点", "用省略号 ……", SEV_HINT, out)
    _find(items, UNIT_GLUE, "W4", "数字与单位之间缺空格",
          "加半角空格：48.62 亿元", SEV_CRITICAL, out)
    _find(items, PCT_SPACE, "W5", "百分号前多一个空格", "改 6.7%", SEV_CRITICAL, out)
    _find(items, LATIN_GLUE, "W6", "中英混排缺空格", "汉字与拉丁之间加空格",
          SEV_HINT, out)

    for word in BANNED:
        _find(items, re.compile(re.escape(word)), "W7", f"禁用词「{word}」",
              "删掉或换成具体描述", SEV_CRITICAL, out)
    _merge_by_code(out, "W7", len(BANNED), "禁用词 ")

    for word in ABSOLUTE:
        _find(items, re.compile(re.escape(word)), "W8", f"绝对化断言「{word}」",
              "无数据支撑就删", SEV_HINT, out)
    _merge_by_code(out, "W8", len(ABSOLUTE), "绝对化断言 ")

    for word in VAGUE:
        _find(items, re.compile(re.escape(word)), "W13", f"模糊量词「{word}」",
              "给数字，或删", SEV_HINT, out)
    _merge_by_code(out, "W13", len(VAGUE), "模糊量词 ")

    # W9 段落过长（只对段落生效，表格单元格不参与）
    # 含三个以上顿号的句子视为列举，列举读起来不累，不计入长句。
    def _sents(text):
        return [s for s in re.split(r"[。！？；]", text) if s.strip()]

    for loc, text in (paragraphs if paragraphs is not None else []):
        if len(text) < 30:
            continue
        if len(_sents(text)) > 4:
            out.append(Finding(loc, "W9", f"段落 {len(_sents(text))} 句（> 4）",
                               "拆段，一段一个意思", SEV_HINT))
            break
    for loc, text in (paragraphs if paragraphs is not None else []):
        long_s = max((len(s) for s in _sents(text)
                      if s.count("、") < 3), default=0)
        if long_s > 60:
            out.append(Finding(loc, "W9", f"单句 {long_s} 字（> 60）",
                               "拆句，一句一个概念", SEV_HINT))
            break

    # W10 泛化标题
    for loc, text in (headings if headings is not None else []):
        t = text.strip()
        if not t or len(t) > 8:
            continue
        if any(t.endswith(g) or t == g for g in GENERIC_HEAD):
            out.append(Finding(loc, "W10", f"泛化标题「{t}」",
                               "改成能预判内容的标题", SEV_HINT))
            break

    # W11 主语混用
    if "公司" in full and re.search(r"我们|笔者|本人", full):
        out.append(Finding("all", "W11", "「公司」与「我们/笔者」同篇混用",
                           "全篇统一一个主语", SEV_CRITICAL))

    # W12 单位口径
    if "亿元" in full and "万元" in full:
        out.append(Finding("all", "W12", "同篇出现「亿元」与「万元」",
                           "检查同一指标有没有换过单位；并用时在表注里声明",
                           SEV_HINT))
    if "个百分点" in full and re.search(r"\d\s*pt", full):
        out.append(Finding("all", "W12", "「个百分点」与「pt」混用",
                           "统一写「个百分点」", SEV_CRITICAL))
