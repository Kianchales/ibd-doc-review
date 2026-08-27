#!/usr/bin/env python3
"""IPO 文档基本格式核对 v2（只读，不改文件）。

按「问题类型 × 严重程度」组织：HIGH=错误（须改）/ MEDIUM=警告 / LOW=提示。

组别与核对项：
  text  文字类：
        heading_seq 标题层级序号连续性（含 第X节/第X章、问题X 自定义编号）   HIGH
        terms       用词规范性（错别字/异形词；支持外部清单扩展）            MEDIUM
        dates       日期写法统一（中文/分隔符/斜杠/连写/英文月缩写/年月）      MEDIUM
        spaces      多余空格与重复标点                                    HIGH
        abbr        释义简称统一（冲突/前置使用/未定义复用/引号风格）          MEDIUM-LOW
        geo         国家城市表述合规（--geo-file 外部清单驱动）              HIGH
  data  数据类：
        amounts     金额千分位与两位小数（豁免 %、文号/编码）                MEDIUM
        consistency 同名指标数值前后一致                                  MEDIUM
        calc        表格合计行求和 / 占比列合计≈100%                       HIGH
        cross_table 跨表同名科目数值比对                                  MEDIUM
  table 表格类：
        table_font  字号体系（五号21pt/小五18pt，其余违规）                 HIGH
        table_align 数字单元格右对齐                                    MEDIUM
        table_empty 空单元格                                           LOW
        table_na    「不适用」标记统一性                                 MEDIIUM

用法：
  python check_content.py --input <docx> [--output 报告.md]
      [--checks all | text | data | table | heading_seq,calc,...]
"""

import argparse
import datetime
import glob
import json
import os
import re
import sys
import zipfile
from collections import Counter

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TERM_RULES_FILE = os.path.join(SKILL_DIR, "references", "term_rules.json")

# ---------------------------------------------------------------- 基础解析


def load_docx(path):
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            out = {}
            if "word/document.xml" in names:
                out["word/document.xml"] = z.read("word/document.xml").decode("utf-8", errors="ignore")
            return out or None
    except Exception as e:
        print(f"[ERROR] 读取失败 {path}: {e}")
        return None


PARA_RE = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.S)
CELL_RE = re.compile(r"<w:tc\b[^>]*>.*?</w:tc>", re.S)
TBL_RE = re.compile(r"<w:tbl\b[^>]*>.*?</w:tbl>", re.S)
ROW_RE = re.compile(r"<w:tr\b[^>]*>.*?</w:tr>", re.S)


def _text_of(fragment):
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", fragment))


class Issue:
    def __init__(self, check_id, check_name, location, snippet, problem,
                 suggestion="", severity="MEDIUM"):
        self.check_id = check_id
        self.check_name = check_name
        self.location = location
        self.snippet = snippet
        self.problem = problem
        self.suggestion = suggestion
        self.severity = severity


SEV_ORDER = ["HIGH", "MEDIUM", "LOW"]
SEV_LABEL = {"HIGH": "错误", "MEDIUM": "警告", "LOW": "提示"}


def _para_info(body):
    ppr_m = re.search(r"<w:pPr\b[^>]*>.*?</w:pPr>", body, re.S)
    ppr = ppr_m.group(0) if ppr_m else None
    st = re.search(r'<w:pStyle w:val="([^"]+)"', ppr or "")
    jc = re.search(r'<w:jc w:val="([^"]+)"', ppr or "")
    return {
        "style": st.group(1) if st else None,
        "align": jc.group(1) if jc else None,
        "text": "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body)).strip(),
    }


def extract_structure(doc):
    items = []
    cell_spans = [(m.start(), m.end()) for m in CELL_RE.finditer(doc)]
    tbl_spans = [(m.start(), m.end()) for m in TBL_RE.finditer(doc)]
    para_spans = [(m.start(), m.end()) for m in PARA_RE.finditer(doc)]

    def table_no(pos):
        return sum(1 for ms, me in tbl_spans if ms <= pos)

    for ps, pe in para_spans:
        in_cell = any(cs <= ps and pe <= ce for cs, ce in cell_spans)
        info = _para_info(doc[ps:pe])
        if in_cell:
            info["kind"] = "cell"
            info["table_no"] = table_no(ps)
        else:
            info["kind"] = "body"
        items.append((info["kind"], info))
    return items


def full_text_of(doc):
    return "\n".join(_text_of(m.group(0)) for m in PARA_RE.finditer(doc))


def parse_tables(doc):
    tables = []
    for tn, tm in enumerate(TBL_RE.finditer(doc), 1):
        rows = []
        for rm in ROW_RE.finditer(tm.group(0)):
            cells = []
            for cm in CELL_RE.finditer(rm.group(0)):
                cxml = cm.group(0)
                sizes = sorted({int(v) for v in re.findall(r'<w:sz w:val="(\d+)"', cxml)})
                paras = [_para_info(pb.group(0)) for pb in PARA_RE.finditer(cxml)]
                raw_text = _text_of(cxml).strip()
                cells.append({"text": raw_text,
                              "raw": "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", cxml)),
                              "sizes": sizes, "paras": paras})
            rows.append(cells)
        head_snips = []
        for row in rows[:2]:
            for c in row:
                if c["text"]:
                    head_snips.append(c["text"][:14])
            if head_snips:
                break
        tables.append({"no": tn, "rows": rows,
                       "head": ("｜".join(head_snips[:3])) if head_snips else ""})
    return tables


# ---------------------------------------------------------------- 中文数字与序号

_CN_UNITS = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
             "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def cn_to_int(s):
    s = s.strip()
    if not s:
        return None
    if s == "十":
        return 10
    if len(s) == 1:
        return _CN_UNITS.get(s)
    if s.startswith("十"):
        rest = s[1:]
        return 10 + (_CN_UNITS.get(rest, 0) if rest else 0)
    if "十" in s:
        a, _, b = s.partition("十")
        av = _CN_UNITS.get(a)
        bv = _CN_UNITS.get(b, 0) if b else 0
        if av is None:
            return None
        return av * 10 + bv
    if s.endswith("百"):
        av = _CN_UNITS.get(s[:-1])
        return av * 100 if av is not None else None
    return _CN_UNITS.get(s)


CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"

CN_CHAPTER_RE = re.compile(r"^第\s*([一二三四五六七八九十百]+|\d+)\s*[章节篇]")
PROBLEM_NUM_RE = re.compile(r"^问题\s*(\d{1,3})(?=$|[.．:：、\s（(])")
LEVEL_RES = [
    ("L1", re.compile(r"^([一二三四五六七八九十百]+)、"), "cn"),
    ("L2", re.compile(r"^[（(]\s*([一二三四五六七八九十百]+)\s*[）)]"), "cn"),
    ("L3", re.compile(r"^(\d{1,3})\s*[、．]"), "num"),
    ("L4", re.compile(r"^[（(]\s*(\d{1,3})\s*[）)]"), "num"),
    ("L5", re.compile(r"^(\d{1,3})\s*[）)]"), "num"),
    ("L6", re.compile(r"^([①-⑳])"), "circled"),
    ("L7", re.compile(r"^([A-Z])\s*[、．.]"), "alpha_up"),
    ("L8", re.compile(r"^([a-z])\s*[、．.]"), "alpha_lo"),
]
LEVEL_NAMES = {
    "C0": "章/节编号（第X节/第X章）",
    "P0": "问题编号（问题X）",
    "L1": "一级（一、）", "L2": "二级（（一））", "L3": "三级（1、）",
    "L4": "四级（（1））", "L5": "五级（1））", "L6": "六级（①）",
    "L7": "七级（A、）", "L8": "八级（a、）",
}
ALL_LEVELS = list(LEVEL_NAMES.keys())


def parse_numbering(text):
    """识别段落开头序号 → (level, value:int, raw)。"""
    m = CN_CHAPTER_RE.match(text)
    if m:
        raw = m.group(1)
        val = int(raw) if raw.isdigit() else cn_to_int(raw)
        if val and val > 0:
            return ("C0", val, raw)
    m = PROBLEM_NUM_RE.match(text)
    if m:
        return ("P0", int(m.group(1)), m.group(1))
    for level, pat, kind in LEVEL_RES:
        m = pat.match(text)
        if m:
            raw = m.group(1)
            if kind == "cn":
                val = cn_to_int(raw)
            elif kind == "circled":
                val = CIRCLED.index(raw) + 1
            elif kind == "alpha_up":
                val = ord(raw) - ord("A") + 1
            elif kind == "alpha_lo":
                val = ord(raw) - ord("a") + 1
            else:
                val = int(raw)
            if val is not None and val > 0:
                return (level, val, raw)
    return None


def _disp_num(level, val):
    if level == "C0":
        return str(val)
    if level == "P0":
        return str(val)
    if level in ("L1", "L2"):
        if val <= 10:
            return ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"][val]
        tens, ones = divmod(val, 10)
        s = "十" if tens == 1 else _CN_UNITS.get(tens, str(tens)) + "十"
        if ones:
            s += _CN_UNITS.get(ones, str(ones))
        return s
    return str(val)


# ---------------------------------------------------------------- 核对项：文字类


def check_heading_seq(items):
    """标题层级序号连续性 v3：
    - 跳过目录条目（以 1~4 位纯数字结尾 = 带页码特征的行）
    - 同层必须递增 +1（跳号/重号/倒退报 HIGH）
    - 进入更深层级：重新起 1；回到更高层级：沿该层自身上次值 +1 继续"""
    issues = []
    prev_li = -1
    prev_val = None
    last_at_level = {}
    seq_no = 0
    toc_tail_re = re.compile(r"\s*[0-9]{1,4}\s*$")
    for kind, info in items:
        if kind != "body" or not info["text"]:
            continue
        text = info["text"]
        # 引号包裹的引用内容不参与序号核对（2026-08-27 裁定：引用条款/承诺原文
        # 内部的「1.」「（一）」等为其自身格式，不应与正文序号链混排）
        stripped = text.strip()
        if stripped[:1] in ('"', "'", "\u201c", "\u201d", "\u300c", "\u300d") or \
           (stripped.startswith('"') and stripped.endswith('"')):
            continue
        # 长段落且含成对引号 → 引用条款/承诺原文，其内部序号不参与正文链
        if len(stripped) > 80 and re.search(r'[“"\u300c][^“"\u300d]{6,}[”"\u300d]', stripped):
            continue
        parsed = parse_numbering(text)
        if not parsed:
            continue
        # 目录条目：以纯数字结尾（页码特征）
        if toc_tail_re.search(text):
            continue
        seq_no += 1
        level, val, raw = parsed
        name = LEVEL_NAMES[level]
        loc = f"{name}#{seq_no}「{text[:22]}」"
        li = ALL_LEVELS.index(level)

        def disp(v):
            return _disp_num(level, v) if v else str(v)

        def note(expected):
            return f"序号「{raw}」，应调整为目标「{disp(expected)}」"

        problem = None
        suggestion = ""
        severity = "HIGH"

        if prev_li == -1:
            expected = 1
            if val != 1:
                problem = f"全文首个序号为「{raw}」（应为起始值 1/一）"
        elif li == prev_li:
            expected = prev_val + 1
            if val != expected:
                kind_str = ("跳号" if val > expected else
                            "重号" if val == prev_val else "倒退")
                problem = (f"序号「{raw}」，前一号为「{_disp_num(level, prev_val)}」，{kind_str}")
                suggestion = f"调整为「{_disp_num(level, expected)}」"
        elif li > prev_li:
            expected = 1
            if val != 1:
                problem = (f"进入更深层级时首个序号为「{raw}」（通常应为起始值 1/一）；"
                           f"若本意是与上层衔接请检查上文是否缺号")
                suggestion = f"如需延续请调整为「{_disp_num(level, expected)}」"
                severity = "LOW"
        else:
            base = last_at_level.get(level)
            expected = (base + 1) if base is not None else 1
            if val != expected:
                kind_str = "跳号" if val > expected else "倒退/重号"
                problem = (f"回到该层级时序号「{raw}」，上一次该层为"
                           f"「{_disp_num(level, base)}」，{kind_str}——若为新小节重新起号可忽略")
                suggestion = f"如需延续请调整为「{_disp_num(level, expected)}」"
                severity = "LOW"

        if problem:
            issues.append(Issue(
                "heading_seq", name, loc, text[:40], problem,
                suggestion if suggestion else f"目标序号 {expected}", severity))
        prev_li = li
        prev_val = val
        last_at_level[level] = val
    return issues


# ---- terms 用词规范

def load_external_rules(path):
    """外部术语规则清单 JSON：[{"pattern":"...","problem":"...","suggestion":"..."}]。"""
    if path and os.path.isfile(path):
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            return [(r.get("pattern", ""), r.get("problem", ""), r.get("suggestion", ""))
                    for r in data if isinstance(r, dict) and r.get("pattern")]
        except Exception as e:
            print(f"[WARN] 术语清单读取失败({path}): {e}")
    return []


BUILTIN_TERM_RULES = [
    {"pattern": r"帐[面户务套]", "problem": "'帐'应为'账'（账面/账户/账务/账套）",
     "suggestion": "统一使用'账'"},
    {"pattern": r"(?<!可)其它(?=[^\u4e00-\u9fff]|$)", "problem": "'其它'宜规范为'其他'",
     "suggestion": "统一使用'其他'"},
]


def check_terms(doc_text, term_rules=None):
    rules = list(BUILTIN_TERM_RULES)
    rules += [(r["pattern"], r["problem"], r["suggestion"]) for r in (term_rules or [])]
    issues = []
    reported = Counter()
    for rule in rules:
        pat, problem, suggestion = rule
        try:
            rx = re.compile(pat)
        except re.error:
            continue
        hits = list(rx.finditer(doc_text))
        shown = 0
        for m in hits:
            reported[pat] += 1
            shown += 1
            if shown > 3:
                break
            near_l = max(0, m.start() - 12)
            snippet = doc_text[near_l:m.end() + 12].replace("\n", "")
            issues.append(Issue("terms", "用词规范",
                                f"全文出现 {len(hits)} 次，示例：「…{snippet[:26]}…」",
                                f"{problem}（命中「{m.group(0)[:16]}」）",
                                suggestion or "", "MEDIUM"))
    return issues


# ---- dates

MONTH_EN = "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec"
DATE_FORMS = [
    ("中文 年月日（2024年6月30日）", re.compile(r"20\d{2}年\d{1,2}月\d{1,2}日?")),
    ("中文 年月（2024年6月）", re.compile(r"20\d{2}年\d{1,2}月(?![\d一二三四五六七八九十]{1,3}日)")),
    ("横线年月日（2024-06-30）", re.compile(r"(?<![\d./-])20\d{2}-\d{1,2}-\d{1,2}(?![\d.-])")),
    ("点分隔年月日（2024.6.30）", re.compile(r"(?<![\d./-])20\d{2}\.\d{1,2}\.\d{1,2}(?![\d.-])")),
    ("斜杠年月日（2024/06/30）", re.compile(r"(?<![\d./-])20\d{2}/\d{1,2}/\d{1,2}(?![\d./-])")),
    ("横线年月（2024-06）", re.compile(r"(?<![\d./-])20\d{2}-\d{2}(?![-\d])")),
    ("连写式 YYYYMMDD（20240630）", re.compile(r"(?<![\d.])(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])(?!\d)")),
    ("英文月缩写（Jun 2024 / Jun 30, 2024）",
     re.compile(rf"\b(?:{MONTH_EN})\.?,?(?:\s+\d{{1,2}},)?\s*20\d{{2}}\b|\b20\d{{2}},?\s+(?:{MONTH_EN})\.?\s+\d{{1,2}}\b", re.I)),
    ("美式 月/日/年（06/30/2024）", re.compile(r"(?<![\d/.])(?:0?[1-9]|1[0-2])/(?:0?[1-9]|[12]\d|3[01])/20\d{2}(?![\d/.])")),
    ("英文 年月（2024-05）", re.compile(r"(?<![\d./-])20\d{2}-(?:0[1-9]|1[0-2])(?![-\d])(?=\D|$)")),
]


def check_dates(items):
    counters = Counter()
    samples = {}
    for i, (kind, info) in enumerate(items):
        t = info["text"]
        if not t:
            continue
        for fname, pat in DATE_FORMS:
            hits = pat.findall(t)
            if hits:
                counters[fname] += len(hits)
                samples.setdefault(fname, []).append((i, kind, info, t))

    issues = []
    if not counters:
        return issues
    # 2026-08-27 裁定：中文「年月日/年月」为最正式表述，可与另一种统一格式共存——
    # 主导判定排除中文式，其余形式内部再统一
    excl = {"中文 年月日（2024年6月30日）", "中文 年月（2024年6月）"}
    dom_items = [(f, c) for f, c in counters.most_common() if f not in excl]
    dominant, dom_n = dom_items[0] if dom_items else (None, 0)
    others = {f: c for f, c in counters.items() if f != dominant and f not in excl}
    if dominant is not None and others:
        others_str = "、".join(f"{f}×{c}" for f, c in sorted(others.items(), key=lambda x: -x[1]))
        issues.append(Issue(
            "dates", "日期写法统一", "全文",
            f"主导「{dominant}」×{dom_n}；其余 {others_str}",
            f"建议全文统一为「{dominant}」，少数派明细如下", "MEDIUM"))
        for fname, lst in samples.items():
            # 中文「X年X月X日」为最正式表述（2026-08-27 裁定），永不列入少数派
            if fname in excl or fname == dominant or len(issues) > 120:
                continue
            shown_loc = set()
            for (i, kind, info, t) in lst:
                pat = dict(DATE_FORMS)[fname]
                m = pat.search(t)
                if not m:
                    continue
                ctx_l = max(0, m.start() - 12)
                loc = ("表格" if kind == "cell" else "正文") + \
                      (f"表{info['table_no']}" if kind == "cell" else f"#{i + 1}") + \
                      f"「…{t[ctx_l:m.end() + 8]}…」"
                if loc in shown_loc:
                    continue
                shown_loc.add(loc)
                issues.append(Issue("dates", "日期写法统一", loc, m.group(0),
                                    "非主流写法", f"建议改为「{dominant}」", "LOW"))
                if len(shown_loc) >= 8:
                    break

    sep_counter = Counter()
    sep_sample = {}
    for i, (kind, info) in enumerate(items):
        t = info["text"]
        for m in re.finditer(r"20\d{2}\s*([．.\-/])\s*\d{1,2}", t):
            sp = "." if m.group(1) in ".．" else m.group(1)
            key = "点分隔(.)" if sp == "." else "横线分隔(-)" if sp == "-" else f"其他({sp})"
            sep_counter[key] += 1
            sep_sample.setdefault(key, i)
    if len(sep_counter) > 1:
        issues.append(Issue(
            "dates", "日期写法统一", "全文",
            "分隔符不统一：" + "、".join(f"{k}×{v}" for k, v in sep_counter.most_common()),
            "建议统一分隔符",
            f"各形式首现位置示例：{sorted(set(sep_sample.values()))}", "LOW"))
    return issues


# ---- spaces 多余空格与重复标点

RE_MULTI_SPACE_CJK = re.compile(r"[\u4e00-\u9fff]  +[\u4e00-\u9fff]")
RE_DUP_CN_PUNCT = re.compile(r"([。，；：？！、])\1+")
RE_DUP_HALF_PUNCT = re.compile(r"[,.;:!?]{2,}")
RE_PUNCT_MIX = re.compile(r"。\.|\.。|，,|,，|；;|::")


def check_spaces(items):
    issues = []
    for i, (kind, info) in enumerate(items):
        t = info["text"]
        if not t:
            continue
        layout_spacing = bool(re.search(r"年\s{2,}月\s{2,}日", t) or
                              re.search(r"目\s{2,}录", t))
        # 排版性豁免（2026-08-27 实测）：签署页「年 月 日」、目录页「目 录」为规范排版；
        # 少量短词宽间隔（如签署页人名「邱  嵩」）降为提示级
        for m in RE_MULTI_SPACE_CJK.finditer(t):
            ctx_l = max(0, m.start() - 10)
            snippet = t[ctx_l:m.end() + 10]
            if layout_spacing or re.search(r"目\s{2,}录", t):
                continue
            # 短词宽间隔（签署页人名/对齐排版，2026-08-27 裁定）完全忽略
            if len(re.sub(r"\s", "", t)) <= 6:
                continue
            issues.append(Issue(
                "spaces", "多余空格/标点", "正文" + f"#{i + 1}" +
                f"「…{snippet[:24]}…」", m.group(0),
                "中文之间出现连续空格",
                "确认是否为刻意排版；否则删除多余空格", "HIGH"))
        for m in RE_DUP_CN_PUNCT.finditer(t):
            ctx_l = max(0, m.start() - 10)
            issues.append(Issue(
                "spaces", "多余空格/标点", "正文" + f"#{i + 1}" +
                f"「…{t[ctx_l:m.end() + 10]}…」", m.group(0),
                f"全角标点重复：「{m.group(0)}」", "去重标点", "HIGH"))
        for m in RE_PUNCT_MIX.finditer(t):
            ctx_l = max(0, m.start() - 10)
            issues.append(Issue(
                "spaces", "多余空格/标点", "正文" + f"#{i + 1}" +
                f"「…{t[ctx_l:m.end() + 10]}…」", m.group(0),
                "中英标点混排", "按语境保留一个并统一全半角", "HIGH"))
        for m in RE_DUP_HALF_PUNCT.finditer(t):
            if re.fullmatch(r"\.{3}|…+", m.group(0)):
                continue  # 省略号
            # 缩写固定用法豁免（2026-08-27 裁定）：Co.,Ltd. / Inc., 等
            if m.group(0)[0] == "." and m.start() > 0 and t[m.start() - 1].isalpha():
                continue
            ctx_l = max(0, m.start() - 10)
            issues.append(Issue(
                "spaces", "多余空格/标点", "正文" + f"#{i + 1}" +
                f"「…{t[ctx_l:m.end() + 10]}…」", m.group(0),
                "连续半角标点重复", "修正标点", "MEDIUM"))
    return issues


# ---- abbr 释义简称

ABBR_DEF_RE = re.compile(
    r"([\u4e00-\u9fffA-Za-z0-9][\u4e00-\u9fffA-Za-z0-9·]{1,39}"
    r"(?:股份有限公司|有限公司|有限责任公司|公司|企业|集团|中心|基金|计划|银行|证券)?"
    r"\s*[（(]\s*以下简称\s*[「\"'『]?([^」\"'』()）]{1,25})[」\"'』]?\s*[）)])")


def check_abbr(doc_text):
    issues = []
    defs = {}
    quote_styles = Counter()
    for m in ABBR_DEF_RE.finditer(doc_text):
        full, short = m.group(1).strip(), m.group(2).strip()
        defs.setdefault(short, []).append((full, m.start()))
        qm = re.search(r"以下简称\s*([「\"'『])", m.group(0))
        if qm:
            quote_styles[qm.group(1)] += 1

    # 同名简称多全称 → 冲突
    for short, pairs in sorted(defs.items()):
        uniq = sorted({p[0] for p in pairs})
        if len(uniq) > 1:
            issues.append(Issue(
                "abbr", "释义简称", f"简称「{short}」（偏移{pairs[0][1]} 起）",
                " / ".join(uniq[:4]),
                f"同名简称对应 {len(uniq)} 个不同全称，存在指向冲突风险",
                "请核实区分口径或更换简称", "MEDIUM"))

        defined_at = min(p[1] for p in pairs)
        full_name = pairs[0][0]

        # 定义前使用（问题级）
        pre_count = doc_text.count(short, 0, defined_at)
        if pre_count > 0:
            issues.append(Issue(
                "abbr", "释义简称", f"简称「{short}」首次定义前",
                short, f"首次定义之前已独立使用 {pre_count} 次",
                "投行惯例简称应在首次全称处即时定义，请核实前置出现是否需补定义或改写",
                "MEDIUM"))

        # 定义后全称复用 ≥3（提示级）
        later_count = doc_text.count(full_name, defined_at + len(pairs[0][0]))
        if later_count >= 3:
            tail = doc_text.find(full_name, defined_at)
            issues.append(Issue(
                "abbr", "释义简称", f"约偏移{tail}", full_name,
                f"已定义简称后仍以全称出现 {later_count} 次",
                f"建议统一切换为简称「{short}」", "LOW"))

    # 未定义使用的疑似简称：括号内纯中文短语，未定义但正文独立复用 ≥2 次（提示级）
    PAREN_SHORT_RE = re.compile(r"[（(]([\u4e00-\u9fff]{2,8})[）)]")
    PAREN_SKIP = {"以下简称", "转回", "转销", "续上表", "承上表"}
    KEYWORD_EXCLUDE = ("万元", "亿元", "年度", "期间", "所得税", "情况")
    # 常见词白名单（2026-08-27 裁定：专业术语/地名/常规表述/数字无需定义，不视为疑似未定义简称）
    ABBR_WHITELIST = {
        # 投行专业术语
        "独立董事", "草案", "主承销商", "董事长", "副董事长", "监事", "监事会",
        "监事会主席", "职工代表监事", "股东大会", "董事会", "董事会秘书", "董秘",
        "总经理", "财务总监", "独立财务顾问", "保荐机构", "联席保荐机构",
        "承销商", "律师事务所", "会计师事务所", "律师", "会计师",
        "审计委员会", "提名委员会", "薪酬与考核委员会", "战略委员会", "审核委员会",
        "执行董事", "非执行董事", "高级管理人员", "核心技术人员", "控股股东",
        "实际控制人", "关联方", "关联交易", "募集资金", "募集资金投资项目",
        "招股说明书", "上市规则", "公司章程", "公司法", "证券法",
        # 地名
        "上海", "北京", "深圳", "广州", "杭州", "南京", "成都", "重庆", "武汉",
        "西安", "天津", "苏州", "宁波", "青岛", "厦门", "长沙", "郑州", "济南",
        "合肥", "福州", "昆明", "大连", "无锡", "佛山", "东莞", "珠海", "中山",
        "境外", "境内", "中国大陆", "香港", "澳门", "台湾",
        # 常规表述
        "一级", "二级", "三级", "四级", "五级", "六级", "七级", "八级",
        "高级", "中级", "初级", "个人", "单位", "金额", "数量", "比例",
        "发行人", "公司", "集团", "企业", "有限合伙", "合伙企业",
    }
    CN_NUM_WORD = re.compile(r"^[一二三四五六七八九十百千两零]{2,8}$")  # 十三/十四/三十一等数字词
    paren_counts = Counter()
    for m in PAREN_SHORT_RE.finditer(doc_text):
        phrase = m.group(1)
        if phrase in PAREN_SKIP or phrase in defs or phrase in ABBR_WHITELIST:
            continue
        if any(k in phrase for k in KEYWORD_EXCLUDE):
            continue
        if CN_NUM_WORD.match(phrase):
            continue
        paren_counts[phrase] += 1
    for phrase, pc in paren_counts.most_common():
        reuse = doc_text.count(phrase) - pc
        if pc >= 1 and reuse >= 2:
            issues.append(Issue(
                "abbr", "释义简称", "全文", phrase,
                f"括号注释短语「{phrase}」（{pc} 处）在正文独立复用 {reuse} 次，未见「以下简称」定义——疑似未定义简称",
                "若作为简称请补规范定义；若非简称可忽略", "LOW"))

    if len(quote_styles) > 1:
        issues.append(Issue(
            "abbr", "释义简称", "全文",
            "、".join(f"{k}×{v}" for k, v in quote_styles.most_common()),
            "「以下简称」引号风格不统一", "建议全文统一引号风格", "LOW"))
    return issues


# ---- geo 国家/城市表述合规（外部清单驱动）

SENSITIVE_TERMS_FILE = os.path.join(SKILL_DIR, "references", "sensitive_terms.json")
LEVEL_MAP = {"CRITICAL": "HIGH", "IMPORTANT": "MEDIUM", "MINOR": "LOW"}


def load_geo_rules(geo_file):
    """敏感词清单：默认加载 skill 内置 references/sensitive_terms.json；
    --geo-file 提供的清单会**追加**进来。支持 [{pattern|term, level, note, suggestion}]。"""
    rules = []
    sources = [SENSITIVE_TERMS_FILE]
    if geo_file and geo_file not in ("", "/dev/null"):
        sources.insert(0, geo_file)
    for src in sources:
        if src and os.path.isfile(src):
            try:
                with open(src, encoding="utf-8") as f:
                    data = json.load(f)
                for g in data:
                    if isinstance(g, dict) and (g.get("pattern") or g.get("term")):
                        rules.append({
                            "pattern": g.get("pattern") or re.escape(g["term"]),
                            "level": g.get("level", "CRITICAL"),
                            "note": g.get("note", ""),
                            "suggestion": g.get("suggestion", ""),
                        })
            except Exception as e:
                print(f"[WARN] 敏感词清单读取失败({src}): {e}")
    return rules


def check_geo(doc_text, geo_rules):
    """国家/城市表述合规：清单驱动，pattern 支持 JSON 正则；CRITICAL=HIGH，
    IMPORTANT=MEDIUM，MINOR=LOW。逐处命中报告位置与上下文。"""
    issues = []
    for rule in (geo_rules or []):
        try:
            rx = re.compile(rule["pattern"])
        except re.error as e:
            print(f"[WARN] 敏感词规则编译失败({rule['pattern'][:20]}): {e}")
            continue
        severity = LEVEL_MAP.get(rule.get("level", "CRITICAL"), "HIGH")
        matches = list(rx.finditer(doc_text))
        if not matches:
            continue
        shown = 0
        for m in matches:
            ctx_l = max(0, m.start() - 12)
            snippet = doc_text[ctx_l:m.end() + 14].replace("\n", "")
            loc = f"全文（首现于「…{snippet[:30]}…」）" if shown == 0 else f"全文（第 {shown + 1} 处）"
            issues.append(Issue(
                "geo", "国家/地区表述合规", loc,
                m.group(0)[:40],
                f"{rule.get('level')}级：{rule.get('note', '')}（累计 {len(matches)} 处）".strip(),
                rule.get("suggestion", ""), severity))
            shown += 1
            if shown >= 3:   # 每条规则最多列 3 处明细
                break
    return issues


# ---- amounts 金额千分位两位小数（豁免 % 与文号/编码）

RE_NO_THOUSANDS = re.compile(r"(?<![\d,.%])(\d{5,9})(?=元|万元|亿元|[^\d]|$)")
RE_BAD_DECIMALS = re.compile(r"\d{1,3}(?:,\d{3})+\.(\d)(?![\d%])")
DOC_CODE_RE = re.compile(r"^\d{6,8}$")


def _is_percent_after(text, end):
    return text[end:].lstrip(" ").startswith("%")


def check_amounts(items):
    issues = []
    seen = set()
    for i, (kind, info) in enumerate(items):
        text = info["text"]
        if not text:
            continue
        work = re.sub(r"\d+(?:[,.]\d+)*\s*%",
                      lambda mm: "%" * len(mm.group(0)), text)
        where = ("表格" + f"表{info['table_no']}") if kind == "cell" else (f"正文#{i + 1}")

        for m in RE_NO_THOUSANDS.finditer(work):
            frag = m.group(0)
            key = (where, frag, m.start())
            if key in seen:
                continue
            seen.add(key)
            if DOC_CODE_RE.match(frag):
                continue  # 文号/编码豁免（含日期连写，由 dates 项处理）
            # 标准号豁免（2026-08-27 裁定）：GB/GB/T/ISO/Q/DB… 后接数字串属标准编号非金额
            prefix = text[max(0, m.start() - 10):m.start()]
            if re.search(r"[A-Za-z]{1,4}(?:/[A-Za-z]{1,4})?\s*$", prefix):
                continue
            # 地址要素豁免（2026-08-27 裁定）：…路6号105室-40627（集中办公区）等地址数字段
            if re.search(r"[路街号室栋层弄巷门]", prefix):
                continue
            # 连字符/波浪线分隔的编号段豁免（如 105室-40627）
            if prefix.rstrip()[-1:] in ("-", "—", "~"):
                continue
            # 编码串豁免：数字后紧跟字母（如 84025S11267R0SC）
            after_ch = text[m.end()] if m.end() < len(text) else ""
            if after_ch.isalpha():
                continue
            # # 前缀编号豁免（如 #72897）
            if prefix.rstrip()[-1:] == "#":
                continue
            # 案号/文号豁免：数字后接「号」（如 民初11336号）
            if text[m.end():m.end() + 2].lstrip().startswith("号"):
                continue
            note = ""
            if len(frag) >= 8:
                note = "（长数字串，若为文号/编码可忽略本条）"
            ctx_l = max(0, m.start() - 12)
            snippet = text[ctx_l:m.end() + 10].replace("\n", "")
            issues.append(Issue(
                "amounts", "金额数字（千分位/两位小数）", where + f"「…{snippet[:26]}…」",
                frag, "位数较多的数字未加千分位分隔符" + note, "添加千分位并核对是否需保留两位小数",
                "MEDIUM"))

        for m in RE_BAD_DECIMALS.finditer(work):
            frag_raw = text[m.start():m.end()] if m.start() < len(text) else ""
            frag = work[m.start():m.end()]
            key = (where, frag, m.start())
            if key in seen:
                continue
            seen.add(key)
            if _is_percent_after(text, m.end()):
                continue  # 百分比豁免
            ctx_l = max(0, m.start() - 12)
            snippet = text[ctx_l:m.end() + 10].replace("\n", "")
            issues.append(Issue(
                "amounts", "金额数字（千分位/两位小数）", where + f"「…{snippet[:26]}…」",
                frag, "带千分位的金额仅保留一位小数", "补齐至两位小数（如 .50）", "MEDIUM"))
    return issues


# ---- punctuation 中英文标点（前后字符判定法）

HALF_PUNCTS = ",.;:?!()\"'"
FULL_PUNCTS = "，。；：？！（）"


def _char_class(c):
    if c and "\u4e00" <= c <= "\u9fff":
        return "C"
    if c and c.isascii() and (c.isalpha() or c.isdigit()):
        return "E"
    return "O"


def _neighbor_class(s, k, direction):
    if direction < 0:
        rng = range(k - 1, -1, -1)
    else:
        rng = range(k + 1, len(s))
    for j in rng:
        c = s[j]
        if not c.isspace():
            return _char_class(c), j
    return None, None


FULL_HALF_MAP = {"，": ",", "。": ".", "；": ";", "：": ":",
                 "？": "?", "！": "!", "（": "(", "）": ")"}
HALF_FULL_MAP = {",": "，", ".": "。", ";": "；", ":": "：",
                 "?": "？", "!": "！", "(": "（", ")": "）"}


def check_punctuation(items):
    """标点规则（按前后字符类型判定）：
    - 半角标点的相邻非空字符任一侧为中文 → 应为全角；
    - 全角标点两侧均为英文/数字 → 应为半角；
    - 其余间隔场景默认中文标点不报。
    中文语境直排引号单独提示。"""
    issues = []
    half_set = set(HALF_PUNCTS)
    full_set = set(FULL_PUNCTS)

    def add(kind, i, info, s, start, end, problem, suggestion, severity="HIGH"):
        ctx_l = max(0, start - 10)
        snippet = s[ctx_l:end + 10].replace("\n", "")
        loc = ("表格" + f"表{info['table_no']}" if kind == "cell"
               else f"正文#{i + 1}") + f"「…{snippet[:24]}…」"
        issues.append(Issue("punctuation", "中英文标点", loc,
                            s[start:end], problem, suggestion, severity))

    for i, (kind, info) in enumerate(items):
        t = info["text"]
        if not t:
            continue
        # 英文缩写内部句点豁免（2026-08-27 裁定）：Inc./Ltd./Co. 后接中文等
        abbrev_re = re.compile(
            r"\b(?:Inc|Ltd|Co|Corp|No|Vol|Dr|Mr|Ms|St)\.\S{0,4}"
            r"|\b[A-Z]{1,3}(?:\.[A-Z]{1,3})+\.\S{0,2}")
        abbrev_spans = [(m.start(), m.end()) for m in abbrev_re.finditer(t)]

        def in_abbrev(pos):
            return any(a <= pos <= b for a, b in abbrev_spans)

        for k, ch in enumerate(t):
            if ch in half_set:
                pc, _pj = _neighbor_class(t, k, -1)
                nc, _nj = _neighbor_class(t, k, 1)
                if pc == "C" or nc == "C":
                    if ch == "." and in_abbrev(k):
                        continue
                    # 序号/编号/日期句点豁免（2026-08-27 裁定）：「1.公司」「2.如」「2024.6.30」
                    if ch == "." and k > 0 and t[k - 1].isdigit():
                        continue
                    add(kind, i, info, t, k, k + 1,
                        f"中文语境使用半角标点「{ch}」",
                        f"改为全角「{HALF_FULL_MAP.get(ch, '对应全角标点')}」")
            # 2026-08-27 裁定：英英之间全角标点不再报——投行文件中文阅读习惯下
            # （如「CPU，GPU」并列、简称「（GPT）」括注）全角即为规范；仅成段英文例外

        # 提示级：中文语境直排双引号
        dq_hits = [m.start() for m in re.finditer(r'"', t)]
        cn_near = any((_char_class(t[pos - 1]) == "C") if pos > 0 else False
                      for pos in dq_hits)
        if dq_hits and cn_near:
            add(kind, i, info, t, dq_hits[0], dq_hits[0] + 1,
                f"中文语境使用直排引号 \" （共 {len(dq_hits)} 处）",
                "建议改用「」或全角弯引号“”", "LOW")
    return issues


# ---- consistency 数值前后一致（同名指标不同值）

METRIC_VAL_RE = re.compile(
    r"([\u4e00-\u9fffA-Za-z][\u4e00-\u9fffA-Za-z0-9（）]{1,19}?(?:营业收入|净利润|净利总额|归母净利润|总资产|总负债|净资产|所有者权益|货币资金|应收账款|存货|总股本|股本))"
    r"[^\d\-]{0,6}(-?[\d,]+\.[\d]{2}|-?\d{4,}(?:,\d{3})*(?:\.\d+)?)\s*(万元|亿元|元|万|%)")


def check_consistency(items):
    metrics = {}
    issues = []
    for i, (kind, info) in enumerate(items):
        t = info["text"]
        if not t:
            continue
        for m in METRIC_VAL_RE.finditer(t):
            name, val, unit = m.group(1), m.group(2), m.group(3)
            # 千分位去掉再比较数值；保持展示为原文片段
            norm_val = val.replace(",", "").rstrip(".")
            try:
                fv = float(norm_val)
                if unit == "万元":
                    fv *= 10000
                elif unit == "亿元":
                    fv *= 100000000
            except ValueError:
                continue
            key = name[-14:]
            metrics.setdefault((key, unit), []).append(
                {"pos": i, "raw": m.group(0)[:40], "val": fv})
    for (name_key, unit), records in metrics.items():
        uniq_vals = {round(r["val"], 4) for r in records}
        if len(uniq_vals) > 1:
            vals_disp = " / ".join(str(r["val"]) + unit for r in records[:6])
            positions = ", ".join(f"#{r['pos'] + 1}段" for r in records[:6])
            issues.append(Issue(
                "consistency", "指标数值前后一致",
                f"指标「{name_key}」（单位{unit}，出现于 {positions}）",
                vals_disp,
                f"同名指标在不同位置数值不同（共 {len(records)} 处、{len(uniq_vals)} 个不同取值），疑似前后不一致——可能是口径差异，请人工核实",
                "核实口径一致后统一数值，或在表述中明确口径差异", "MEDIUM"))
    return issues


# ---- calc 表格合计行求和 / 占比列合计≈100%

def _to_num(s):
    s = s.replace(",", "").replace("%", "").strip().rstrip("。")
    try:
        return float(s)
    except ValueError:
        return None


def check_calc(tables):
    issues = []
    for tbl in tables:
        no = tbl["no"]
        rows = tbl["rows"]
        texts = [[c["text"] for c in row] for row in rows]
        tbl_loc = f"表{no}" + (f"（首行『{tbl['head']}』）" if tbl.get("head") else "")

        # (a) 合计行求和校验（其余数据行列和 vs 合计行；×100 差异视为单位口径问题）
        for ri, trow in enumerate(texts):
            if not any(("合计" in c) or ("总计" in c) for c in trow if isinstance(c, str)):
                continue
            header_rows = 1 if ri > 0 else 0
            data_rows = [r2 for ri2, r2 in enumerate(texts)
                         if ri2 >= header_rows and ri2 != ri and any(c.strip() for c in r2)]
            # 排除小计/中间汇总行（2026-08-27 裁定：小计行含数值，作分项会重复计数）
            data_rows = [r2 for r2 in data_rows
                         if not any(("小计" in c) or ("其中" in c) or ("减：" in c)
                                    or ("加：" in c) or ("剔除" in c)
                                    for c in r2 if isinstance(c, str))]
            if len(data_rows) < 2:
                continue
            bad_cols = []
            for ci in range(len(trow)):
                col_total = _to_num(trow[ci])
                if col_total is None:
                    continue
                parts = []
                ok_all = True
                for drow in data_rows:
                    v = _to_num(drow[ci]) if ci < len(drow) else None
                    if v is None:
                        ok_all = False
                        break
                    parts.append(v)
                if not ok_all or not parts:
                    continue  # 该列存在非数值单元格，跳过（无法可靠求和）
                s = round(sum(parts), 4)
                tol = max(0.02, abs(col_total) * 0.001)
                if abs(s - col_total) > tol:
                    bad_cols.append((ci, s, col_total))
            if not bad_cols:
                continue
            # 单位口径差异识别：仅分项之和恰为合计值的 100 倍（小数 vs 百分数）时降级
            unit_issue = []
            real_bad = []
            for bc in bad_cols:
                ci, s, tot = bc
                if s and abs(s - tot * 100) <= max(0.02, abs(tot * 100) * 0.001):
                    unit_issue.append(bc)   # 小数 vs 百分数 口径
                else:
                    real_bad.append(bc)
            if unit_issue:
                ci, s, tot = unit_issue[0]
                issues.append(Issue(
                    "calc", "表格合计与占比计算校验",
                    f"{tbl_loc} 合计行第{ci + 1}列",
                    f"分项之和 {s:g} 为合计值 {tot:g} 的 100 倍",
                    "疑似单位/口径不一致（分项与合计数量级差异过大），请人工核实",
                    "MEDIUM"))
            for ci, s, tot in real_bad[:3]:
                issues.append(Issue(
                    "calc", "表格合计与占比计算校验",
                    f"{tbl_loc} 合计行（第{ri + 1}行）第{ci + 1}列",
                    f"分项之和 {s:g} ≠ 合计值 {tot:g}",
                    "疑似计算错误或分项有遗漏（容差已计入四舍五入），请人工复核",
                    "重新计算合计或补充分项", "HIGH"))

        # (b) 占比列合计 ≈100%
        ncols = max((len(r) for r in texts), default=0)
        for ci in range(ncols):
            pct_vals = []
            has_header = False
            for ri, trow in enumerate(texts):
                if any(("小计" in c) or ("其中" in c) for c in trow if isinstance(c, str)):
                    pct_vals.append(None)
                    continue
                if ci >= len(trow):
                    pct_vals.append(None)
                    continue
                txt = trow[ci].strip()
                if any(k in txt for k in ("比例", "占比", "%")):
                    has_header = True
                v = _to_num(txt)
                if "%" in txt or (txt.endswith("%")):
                    v = _to_num(txt.replace("%", ""))
                    pct_vals.append(v if v is None else v)
                    continue
                if "占比" in (texts[0][ci] if texts and ci < len(texts[0]) else ""):
                    pct_vals.append(v)
                else:
                    pct_vals.append(None)
            nums = [v for v in pct_vals if v is not None]
            if has_header and len(nums) >= 3:
                ssum = round(sum(nums), 2)
                if abs(ssum - 100.0) > 1.0:
                    issues.append(Issue(
                        "calc", "表格合计与占比计算校验",
                        f"{tbl_loc} 第{ci + 1}列（占比列）",
                        f"占比合计 {ssum:g}% ≠ 100%（±1%）",
                        "疑似占比计算错误或有遗漏项，请人工复核", "HIGH"))
    return issues


def c_ok_text(row, ci):
    try:
        return isinstance(row[ci], str) and bool(row[ci].strip()) and \
            not any(k in row[ci] for k in ("单位", "注"))
    except Exception:
        return False


# ---- cross_table 同名科目跨表数值比对

def check_cross_table(tables):
    first_col_values = {}
    for tbl in tables:
        seen_in_this_table = set()
        for row in tbl["rows"]:
            if not row:
                continue
            key = row[0]["text"]
            if not key or key in seen_in_this_table:
                continue
            seen_in_this_table.add(key)
            vals = tuple(_to_num(c["text"]) for c in row[1:])
            first_col_values.setdefault(key, []).append((tbl["no"], vals))

    issues = []
    for key, occurrences in sorted(first_col_values.items()):
        distinct = {(vals) for _, vals in occurrences if all(v is not None for v in vals)}
        tables_involved = [no for no, _ in occurrences]
        if len(distinct) > 1 and len(tables_involved) > 1:
            disp = " / ".join(str(list(v))[:60] for _, v in occurrences[:4])
            issues.append(Issue(
                "cross_table", "跨表同名科目勾稽", f"科目「{key}」出现在表 {'、'.join('表'+str(n) for n in tables_involved)}",
                f"各行数值不一致：{disp}",
                "疑似勾稽关系差异——可能是口径/期间不同属正常，请人工核实", "MEDIUM"))
    return issues


# ---- table_* 表格类

FONT_OK_SIZES = {21, 18}   # 半点值：21 = 10.5pt 五号；18 = 9pt 小五


def _is_formal_data_table(tbl):
    """正式数据表判定（2026-08-27 裁定）：行数≤2（封面提示框）或整表无数字
    （签字页/人名表）不参照正文表格规范——豁免字号与对齐检查。"""
    rows = tbl["rows"]
    if len(rows) <= 2:
        return False
    joined = "".join(c["text"] for row in rows for c in row)
    return bool(re.search(r"\d", joined))


def check_table_font(tables):
    issues = []
    for tbl in tables:
        if not _is_formal_data_table(tbl):
            continue
        bad_by_size = Counter()
        example = None
        for ri, row in enumerate(tbl["rows"]):
            for ci, cell in enumerate(row):
                if not cell["text"]:
                    continue
                for sz in cell["sizes"]:
                    if sz in FONT_OK_SIZES:
                        continue
                    pt = sz / 2
                    bad_by_size[f"{pt:g}pt(sz={sz})"] += 1
                    if example is None:
                        example = (ri + 1, ci + 1, cell["text"][:16], f"{pt:g}pt")
        if bad_by_size:
            detail = "、".join(f"{k}×{v}" for k, v in bad_by_size.most_common())
            ex_txt = (f"，如 行{example[0]} 列{example[1]}「{example[2]}」为 {example[3]}"
                      if example else "")
            issues.append(Issue(
                "table_font", "表格字号体系",
                f"表{tbl['no']}" + (f"（首行『{tbl['head']}』）" if tbl.get("head") else ""),
                detail,
                "IPO 表格字号应统一为五号（10.5pt）；放不下可用小五（9pt）" + ex_txt,
                "将违规字号调整为五号或小五", "HIGH"))
    return issues


def _is_numeric_like(cell):
    t = cell["text"]
    if not t:
        return False
    core = re.sub(r"[（）()\u4e00-\u9fff]", "", t)
    return bool(re.fullmatch(r"-?\d{1,3}(?:,\d{3})+(?:\.\d+)?|-?\d{4,}(?:\.\d+)?", core))


def check_table_align(tables):
    issues = []
    examples = Counter()
    sample_cell = None
    for tbl in tables:
        if not _is_formal_data_table(tbl):
            continue
        for ri, row in enumerate(tbl["rows"]):
            for ci, cell in enumerate(row):
                if ci == 0 or not _is_numeric_like(cell):
                    continue
                non_right = [p for p in cell["paras"]
                             if p["align"] not in ("right", None)]
                if non_right:
                    examples[(tbl["no"],)] += 1
                    if sample_cell is None:
                        sample_cell = (tbl["no"], ri + 1, ci + 1, cell["text"][:16])
    if examples:
        detail = "、".join(f"表{k}×{v}" for k, v in examples.most_common())
        sc = (f"，如 表{sample_cell[0]} 行{sample_cell[1]} 列{sample_cell[2]}"
              f"「{sample_cell[3]}」") if sample_cell else ""
        issues.append(Issue(
            "table_align", "表格数字右对齐", detail + sc,
            "部分数字单元格未右对齐",
            "表格内会计数字建议右对齐（参考 check_content 的规则来源：rules.md 表格 v2）", "MEDIUM"))
    return issues


def check_table_empty(tables):
    """空单元格（2026-08-27 裁定：不再逐表报警告，改为全文 1 条汇总提示）。"""
    issues = []
    total_all = 0
    affected = 0
    table_names = []
    for tbl in tables:
        empty_rows = Counter()
        for ri, row in enumerate(tbl["rows"]):
            for ci, cell in enumerate(row):
                if cell["text"] == "":
                    empty_rows[(ri + 1)] += 1
        if empty_rows:
            total = sum(empty_rows.values())
            total_all += total
            affected += 1
            table_names.append(f"表{tbl['no']}" + (f"（首行『{tbl['head']}』）" if tbl.get("head") else ""))
    if total_all:
        detail_show = "、".join(table_names[:8]) + ("…等" if len(table_names) > 8 else "")
        issues.append(Issue(
            "table_empty", "表格空单元格（汇总提示）",
            f"共 {affected} 张表、{total_all} 个空单元格：" + detail_show,
            f"空单元格共 {total_all} 个（{affected} 张表）",
                "如为无内容建议统一以「—」填充；确属留白可忽略", "LOW"))
    return issues


NA_TOKENS_GROUPS = {
    "长破折号": ["—"],
    "短横线": ["-", "－"],
    "斜杠": ["/"],
    "文字类": ["不适用", "N/A", "N.A.", "无"],
}


def check_table_na(tables):
    issues = []
    for tbl in tables:
        na_counter = Counter()
        for row in tbl["rows"]:
            for cell in row:
                t = cell["text"]
                if t in [tok for toks in NA_TOKENS_GROUPS.values() for tok in toks]:
                    na_counter[t] += 1
        distinct = list(na_counter.keys())
        if len(distinct) > 1:
            merged = {}
            for gname, tokens in NA_TOKENS_GROUPS.items():
                used = [k for k in distinct if k in tokens]
                c = sum(v for k, v in na_counter.items() if k in tokens)
                if c:
                    merged[gname] = (c, used)
            if len(merged) > 1:
                detail = "、".join(f"{g}{'/'.join(tk)}×{c}" for g, (c, tk) in merged.items())
                issues.append(Issue(
                    "table_na", "表格填写（标记统一性）",
                    f"表{tbl['no']}" + (f"（首行『{tbl['head']}』）" if tbl.get("head") else ""), detail,
                    "同一表内「不适用/无内容」标记符号混用",
                    "建议全表统一一种标记（通常「—」或「不适用」）", "MEDIUM"))
    return issues


# ---------------------------------------------------------------- 登记 & 主流程

CHECK_REGISTRY = [
    {"id": "heading_seq", "group": "text", "name": "标题层级序号连续性（跳号/重号/倒退）", "severity": "HIGH"},
    {"id": "terms", "group": "text", "name": "用词规范性（错别字/异形词）", "severity": "MEDIUM"},
    {"id": "dates", "group": "text", "name": "日期写法统一（十种形式识别）", "severity": "MEDIUM"},
    {"id": "spaces", "group": "text", "name": "多余空格与重复标点", "severity": "HIGH"},
    {"id": "punctuation", "group": "text", "name": "中英文标点（前后字符判定）", "severity": "HIGH"},
    {"id": "abbr", "group": "text", "name": "释义简称统一（含未定义使用检出）", "severity": "MEDIUM"},
    {"id": "geo", "group": "text", "name": "国家/城市表述合规（外部清单）", "severity": "HIGH"},
    {"id": "amounts", "group": "data", "name": "金额千分位与两位小数", "severity": "MEDIUM"},
    {"id": "consistency", "group": "data", "name": "指标数值前后一致", "severity": "MEDIUM"},
    {"id": "calc", "group": "data", "name": "表格合计与占比计算校验", "severity": "HIGH"},
    {"id": "cross_table", "group": "data", "name": "跨表同名科目勾稽比对", "severity": "MEDIUM"},
    {"id": "table_font", "group": "table", "name": "表格字号体系（五号/小五）", "severity": "HIGH"},
    {"id": "table_align", "group": "table", "name": "表格数字右对齐", "severity": "MEDIUM"},
    {"id": "table_empty", "group": "table", "name": "表格空单元格", "severity": "LOW"},
    {"id": "table_na", "group": "table", "name": "不适用标记统一性", "severity": "MEDIUM"},
]

GROUPS = {"text": [], "data": [], "table": []}
for _item in CHECK_REGISTRY:
    GROUPS[_item["group"]].append(_item["id"])
CHECK_BY_ID = {_item["id"]: _item for _item in CHECK_REGISTRY}


def resolve_checks(spec):
    """all / 组名 / 核对项 id（可混合逗号分隔）。返回有序 id 集合。"""
    chosen = []
    if spec in (None, "", "all"):
        return [c["id"] for c in CHECK_REGISTRY]
    tokens = [tk.strip() for tk in spec.split(",") if tk.strip()]
    for tk in tokens:
        if tk in GROUPS:
            chosen.extend(GROUPS[tk])
        elif tk in CHECK_BY_ID:
            chosen.append(tk)
        else:
            print(f"[WARN] 未知核对项/组名: {tk}")
    ordered = [c["id"] for c in CHECK_REGISTRY if c["id"] in set(chosen)]
    return ordered


def run(input_path, check_ids, geo_file=None, terms_file=None):
    xmls = load_docx(input_path)
    if not xmls:
        return None
    doc = xmls["word/document.xml"]
    items = extract_structure(doc)
    tables = parse_tables(doc)
    full_text = full_text_of(doc)
    geo_rules = load_geo_rules(geo_file)
    term_rules = load_external_rules(terms_file)

    runners = {
        "heading_seq": lambda: check_heading_seq(items),
        "terms": lambda: check_terms(full_text, term_rules),
        "dates": lambda: check_dates(items),
        "spaces": lambda: check_spaces(items),
        "punctuation": lambda: check_punctuation(items),
        "abbr": lambda: check_abbr(full_text),
        "geo": lambda: check_geo(full_text, geo_rules),
        "amounts": lambda: check_amounts(items),
        "consistency": lambda: check_consistency(items),
        "calc": lambda: check_calc(tables),
        "cross_table": lambda: check_cross_table(tables),
        "table_font": lambda: check_table_font(tables),
        "table_align": lambda: check_table_align(tables),
        "table_empty": lambda: check_table_empty(tables),
        "table_na": lambda: check_table_na(tables),
    }
    results = {}
    for cid in check_ids:
        results[cid] = runners[cid]() if cid in runners else []
    return base_name(input_path), results


def base_name(p):
    return os.path.basename(p)


def render_report(base, check_ids, results):
    lines = ["# 投行基本格式核对报告", ""]
    lines.append(f"- **核对对象**：`{base}`")
    lines.append("- **性质**：只读核对，未对文件做任何修改")
    lines.append("- 生成：ibd-doc-review `check_content.py` v2")
    lines.append("")
    cnt = {s: 0 for s in SEV_ORDER}
    by_check_sev = {}
    for cid in check_ids:
        for it in results[cid]:
            cnt[it.severity] = cnt.get(it.severity, 0) + 1
            by_check_sev.setdefault(cid, Counter())[it.severity] += 1

    lines.append("## 核对总览（按严重程度）")
    lines.append("")
    lines.append("| 组别 | 核对项 | 错误(HIGH) | 警告(MED) | 提示(LOW) |")
    lines.append("|------|--------|-----------|-----------|-----------|")
    group_names = {"text": "文字类", "data": "数据类", "table": "表格类"}
    last_group = None
    for item in CHECK_REGISTRY:
        cid = item["id"]
        if cid not in results:
            continue
        glabel = group_names[item["group"]]
        last_group = item["group"]
        sevc = by_check_sev.get(cid, Counter())
        lines.append(f"| {glabel} | {item['name']} | {sevc.get('HIGH', 0)} | "
                     f"{sevc.get('MEDIUM', 0)} | {sevc.get('LOW', 0)} |")
    lines.append(f"| **合计** | — | **{cnt.get('HIGH', 0)}** | **{cnt.get('MEDIUM', 0)}** | "
                 f"**{cnt.get('LOW', 0)}** |")
    lines.append("")
    lines.append("> 严重程度：HIGH=错误（必须修正）；MEDIUM=警告（大概率需修正）；LOW=提示（人工酌情）。")
    lines.append("> 标注「疑似」的条目为机器初筛结果，需人工复核定性。")
    lines.append("")

    sev_block = {"HIGH": "### 🔴 错误（HIGH，须修正）",
                 "MEDIUM": "### 🟡 警告（MEDIUM，建议修正）",
                 "LOW": "### 🟢 提示（LOW，人工酌情）"}

    for sev in SEV_ORDER:
        block_items = [(cid, it) for cid in check_ids for it in results[cid]
                       if it.severity == sev]
        lines.append(sev_block[sev])
        lines.append("")
        if not block_items:
            lines.append("无。")
            lines.append("")
            continue
        lines.append("| 核对项 | 位置 | 原文/对象 | 问题 | 建议 |")
        lines.append("|---|------|----------|------|------|")
        esc = lambda s: s.replace("|", "\\|").replace("\n", " ")
        block_items.sort(key=lambda x: CHECK_BY_ID.get(x[0], {}).get("name", ""))
        for idx, (cid, it) in enumerate(block_items, 1):
            cname = CHECK_BY_ID.get(cid, {}).get("name", cid)
            lines.append(f"| {idx} | {cname} | {esc(it.location)} | {esc(it.snippet)} | "
                         f"{esc(it.problem)} | {esc(it.suggestion)} |")
        lines.append("")
    return "\n".join(lines)


def resolve_files(path):
    if os.path.isdir(path):
        return sorted(glob.glob(os.path.join(path, "*.docx")))
    if os.path.isfile(path):
        return [path]
    return glob.glob(path)


def main():
    ap = argparse.ArgumentParser(description="IPO 文档基本格式核对 v2（只读，不改文件）")
    ap.add_argument("--input", required=True, help="docx 文件路径（支持 glob）")
    ap.add_argument("--output", default=None, help="核对报告 md 输出路径（默认 <input>_格式核对报告.md）")
    ap.add_argument("--checks", default="all",
                    help="all 或 组名(text/data/table) 或核对项 id，可组合逗号分隔，"
                         "如 --checks text,data 或 --checks calc,cross_table")
    ap.add_argument("--geo-file", default=None,
                    help="国家/城市敏感词清单 JSON：[{\"term\":\"...\",\"note\":\"...\",\"suggestion\":\"...\"}]")
    ap.add_argument("--terms-file", default=None,
                    help="术语规则清单 JSON 扩展：[{\"pattern\":\"...\",\"problem\":\"...\",\"suggestion\":\"...\"}]")
    args = ap.parse_args()

    files = resolve_files(args.input)
    if not files:
        print(f"[ERROR] 未找到 docx: {args.input}")
        sys.exit(1)

    check_ids = resolve_checks(args.checks)
    exit_ok = True
    for f in files:
        result = run(f, check_ids, geo_file=args.geo_file, terms_file=args.terms_file)
        if result is None:
            exit_ok = False
            continue
        base, results = result
        report = render_report(base, check_ids, results)
        out = args.output or os.path.splitext(f)[0] + "_格式核对报告.md"
        with open(out, "w", encoding="utf-8") as fh:
            fh.write(report)

        summary_parts = []
        sev_cnt = Counter()
        for cid in check_ids:
            for it in results[cid]:
                sev_cnt[it.severity] += 1
        for sev in SEV_ORDER:
            label = SEV_LABEL[sev]
            mark = "✅ 0" if sev_cnt.get(sev, 0) == 0 else f"⚠️ {sev_cnt[sev]}"
            summary_parts.append(f"{label}: {mark}")
        print(f"\n=== 格式核对：{os.path.basename(f)} ===")
        for line_txt in summary_parts:
            print(f"  {line_txt}")
        print(f"  → 报告已写入：{out}")

    sys.exit(0 if exit_ok else 2)


if __name__ == "__main__":
    main()
