#!/usr/bin/env python3
"""IPO 文档基本格式核对（只读，不改文件）。

六大核对项：
  1. 标题层级序号：序号链连续统一，无跳号、无重号
  2. 金额数字：千分位分隔符 + 保留两位小数
  3. 日期写法：多种写法并存时报主导写法与其他实例
  4. 释义/简称：定义唯一性、定义后全称复用提示、引号风格
  5. 中英文标点：中文语境半角标点、字母数字间全角标点、括号混用
  6. 表格填写：空单元格、不适用符号统一性、单元格首尾空格

用法：
  python check_content.py --input <docx> [--output <核对报告.md>] [--checks 123456]

输出：控制台摘要 + Markdown 核对报告（默认 <input>_格式核对报告.md）
"""

import argparse
import glob
import os
import re
import sys
import zipfile
from collections import Counter

# ---------------------------------------------------------------- 基础解析


def load_docx(path):
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            out = {}
            if "word/document.xml" in names:
                out["word/document.xml"] = z.read("word/document.xml").decode("utf-8", errors="ignore")
            if "word/settings.xml" in names:
                out["word/settings.xml"] = z.read("word/settings.xml").decode("utf-8", errors="ignore")
            return out or None
    except Exception as e:
        print(f"[ERROR] 读取失败 {path}: {e}")
        return None


PARA_RE = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.S)
CELL_RE = re.compile(r"<w:tc\b[^>]*>.*?</w:tc>", re.S)
TBL_RE = re.compile(r"<w:tbl\b[^>]*>.*?</w:tbl>", re.S)


def _text_of(fragment):
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", fragment))


def extract_structure(doc):
    """提取文档结构序列 [(kind, text, extra)]：
    kind ∈ body（正文段落）/ cell（表格单元格）；cell 的 extra 为表格序号（从 1 起）。
    顺序：按文档中各元素出现位置排列；表格内段落统一归为 cell。"""
    items = []
    tbl_spans = [(m.start(), m.end()) for m in TBL_RE.finditer(doc)]
    cell_spans = [(m.start(), m.end()) for m in CELL_RE.finditer(doc)]
    para_spans = [(m.start(), m.end()) for m in PARA_RE.finditer(doc)]

    def table_no(pos):
        return sum(1 for ms, me in tbl_spans if ms <= pos)

    consumed = set()
    events = []  # (start, end, kind, extra)
    # 表格整体占位的事件用于吞掉内部段落判定
    for ps, pe in para_spans:
        key = (ps, pe)
        if key in consumed:
            continue
        in_cell = any(cs <= ps and pe <= ce for cs, ce in cell_spans)
        t = _text_of(doc[ps:pe]).strip()
        if in_cell:
            events.append((ps, pe, "cell", table_no(ps), t))
        else:
            events.append((ps, pe, "body", None, t))
        consumed.add(key)
    events.sort(key=lambda e: e[0])
    return [(kind, text, extra) for _, _, kind, extra, text in events]


class Issue:
    def __init__(self, check, location, snippet, problem, suggestion=""):
        self.check = check
        self.location = location
        self.snippet = snippet
        self.problem = problem
        self.suggestion = suggestion


# ---------------------------------------------------------------- 1. 标题层级序号

_CN_UNITS = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
             "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def cn_to_int(s):
    """中文数字（简体，支持 一~九十九）转 int；失败返回 None。"""
    if not s:
        return None
    s = s.strip()
    if s == "十":
        return 10
    if len(s) == 1:
        v = _CN_UNITS.get(s)
        return v
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
    return _CN_UNITS.get(s)


CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"

NUM_PATTERNS = [
    ("L1", re.compile(r"^([一二三四五六七八九十百]+)、")),
    ("L2", re.compile(r"^[（(]\s*([一二三四五六七八九十百]+)\s*[）)]")),
    ("L3", re.compile(r"^(\d{1,3})\s*[、．]")),
    ("L4", re.compile(r"^[（(]\s*(\d{1,3})\s*[）)]")),
    ("L5", re.compile(r"^(\d{1,3})\s*[）)]")),
    ("L6", re.compile(r"^([①-⑳])")),
    ("L7", re.compile(r"^([A-Z])\s*[、．.]")),
    ("L8", re.compile(r"^([a-z])\s*[、．.]")),
]


def parse_numbering(text):
    """识别段落开头序号，返回 (level, 序号数值 int) 或 None。"""
    for idx, (level, pat) in enumerate(NUM_PATTERNS):
        m = pat.match(text)
        if m:
            raw = m.group(1)
            if level in ("L1", "L2"):
                val = cn_to_int(raw)
            elif level == "L6":
                val = CIRCLED.index(raw) + 1
            elif level in ("L7", "L8"):
                val = ord(raw) - (ord("A") if level == "L7" else ord("a")) + 1
            else:
                val = int(raw)
            if val is None or val <= 0:
                return None
            return (level, val, raw)
    return None


LEVEL_NAMES = {"L1": "一级（一、）", "L2": "二级（（一））", "L3": "三级（1、）",
               "L4": "四级（（1））", "L5": "五级（1））", "L6": "六级（①）",
               "L7": "七级（A、）", "L8": "八级（a、）"}
ALL_LEVELS = list(LEVEL_NAMES.keys())


def _disp_num(level, val):
    """层级内数值按其书写体系显示：中文层（一、/（一））转中文数字，其余阿拉伯。"""
    if level in ("L1", "L2"):
        if val <= 10:
            return ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"][val]
        tens, ones = divmod(val, 10)
        s = "十" if tens == 1 else _CN_UNITS.get(tens, str(tens)) + "十"
        if ones:
            s += _CN_UNITS.get(ones, str(ones))
        return s
    return str(val)


def check_heading_sequence(items):
    """层次感知的序号连续性核对：同层须递增 +1；进入更低层重新起 1；回到高层按其自身序列续。"""
    issues = []
    seq_no = 0
    prev_level_i = -1
    prev_num = None
    last_at_level = {}
    for kind, text, extra in items:
        if kind != "body" or not text:
            continue
        parsed = parse_numbering(text)
        if not parsed:
            continue
        seq_no += 1
        level, val, raw = parsed
        loc = f"序号段#{seq_no}「{text[:18]}」"
        li = ALL_LEVELS.index(level)
        disp_prev = _disp_num(level, prev_num) if prev_num else None
        disp_exp = _disp_num(level, (prev_num or 0) + 1)
        if prev_level_i == -1:
            if val != 1:
                issues.append(Issue("标题序号", loc, text[:40],
                                    f"{LEVEL_NAMES[level]}首个序号为 {raw}（应为起始值 1/一）"))
        elif li == prev_level_i:
            expected = (prev_num or 0) + 1
            if val != expected:
                kind_str = "重号" if val == prev_num else ("跳号" if val > expected else "倒退/重号")
                exp_disp = _disp_num(level, expected)
                issues.append(Issue("标题序号", loc, text[:40],
                                    f"{LEVEL_NAMES[level]}序号「{raw}」，前一号为「{disp_prev}」，{kind_str}（应连续，期望「{exp_disp}」）"))
        else:
            # 层级切换：更低层重新起 1；更高层按其自身上一值续
            if li > prev_level_i:
                if val != 1:
                    issues.append(Issue("标题序号", loc, text[:40],
                                        f"降入{LEVEL_NAMES[level]}，首个序号为 {raw}（应为起始值 1/一）"))
            else:
                last_here = last_at_level.get(level)
                if last_here is not None and val != last_here + 1:
                    kind_str = "重号" if val == last_here else ("跳号" if val > last_here + 1 else "倒退/重号")
                    issues.append(Issue("标题序号", loc, text[:40],
                                        f"回升至{LEVEL_NAMES[level]}，序号 {raw}，上次该级为 {last_here}，{kind_str}"))
        prev_level_i = li
        prev_num = val
        last_at_level[level] = val
    return issues


# ---------------------------------------------------------------- 2. 金额数字

RE_NO_THOUSANDS = re.compile(r"(?<![\d,.%])(\d{5,9})(?=元|万元|亿元|[^\d]|$)")
RE_BAD_DECIMALS = re.compile(r"\d{1,3}(?:,\d{3})+\.(\d)(?![\d])")
RE_DOC_CODE = re.compile(r"^\d{6,8}$")  # 豁免：文号/编码类 6~8 位纯数字


def _is_percent_after(text, end):
    """命中片段之后（允许一个空格）紧跟百分号 → 比例，豁免。"""
    tail = text[end:].lstrip(" ")
    return tail.startswith("%")


def check_amounts(items):
    issues = []
    counter = 0
    seen_positions = set()
    for i, (kind, text, extra) in enumerate(items):
        if not text:
            continue
        # 预剥离百分比片段（豁免：比例数值不作千分位/两位小数要求）
        work_text = re.sub(r"\d+(?:[,.]\d+)*\s*%", lambda mm: "%" * len(mm.group(0)), text)
        for pat, tag, msg in (
            (RE_NO_THOUSANDS, "无千分位", "5 位以上数字未加千分位分隔符"),
            (RE_BAD_DECIMALS, "小数位不足", "带千分位的金额小数位不足两位"),
        ):
            for m in pat.finditer(work_text):
                frag = m.group(0)
                key = (i, frag, m.start())
                if key in seen_positions:
                    continue
                seen_positions.add(key)
                counter += 1
                ctx_l = max(0, m.start() - 14)
                snippet = text[ctx_l:m.end() + 8]
                note = ""
                if tag == "无千分位":
                    if RE_DOC_CODE.match(frag):
                        note = ""
                        counter -= 1
                        seen_positions.discard(key)
                        continue  # 豁免：6~8 位视为文号/编码（含日期连写，由第 3 项核对处理）
                    if len(frag) >= 8:
                        note = "（长数字串，若为文号/编码可忽略本条）"
                loc = ("表格" if kind == "cell" else "正文") + \
                      (f"表{extra}" if kind == "cell" else f"第{i + 1}段") + f"「…{snippet[:26]}…」"
                issues.append(Issue("金额数字", loc, frag, msg, note))
    return issues


# ---------------------------------------------------------------- 3. 日期写法

RE_D_SEQ = re.compile(r"(?<![\d.])(20\d{2}\s*[．.\-/]\s*\d{1,2}\s*[．.\-/]\s*\d{1,2})(?!\d)")
RE_D_COMPACT = re.compile(r"(?<![\d.])((?:19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01]))(?!\d)")
RE_D_CN = re.compile(r"(20\d{2}年\d{1,2}月\d{1,2}日?)")
RE_D_SLASH = re.compile(r"(?<![\d./])(20\d{2}/\d{1,2}/\d{1,2})(?![\d./])")


def normalize_seq(s):
    parts = re.split(r"\s*[．.\-/]\s*", s.strip())
    if len(parts) != 3:
        return None
    y, mo, d = parts
    try:
        return f"{int(y)}-{int(mo):02d}-{int(d):02d}"
    except ValueError:
        return None


def norm_compact(s):
    if len(s) != 8:
        return None
    try:
        return f"{int(s[:4])}-{int(s[4:6]):02d}-{int(s[6:]):02d}"
    except ValueError:
        return None


def norm_cn(s):
    m = re.match(r"(20\d{2})年(\d{1,2})月(\d{1,2})日?", s)
    if not m:
        return None
    return f"{int(m.group(1))}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"


def check_dates(items):
    forms = {
        "连写式 YYYYMMDD": RE_D_COMPACT,
        "点/横线分隔式": RE_D_SEQ,
        "斜杠式 YYYY/M/D": RE_D_SLASH,
        "中文式 年月日": RE_D_CN,
    }
    counters = Counter()
    samples = {}
    detail_rows = []
    n = 0
    for i, (kind, text, extra) in enumerate(items):
        if not text:
            continue
        for fname, pat in forms.items():
            for m in pat.finditer(text):
                raw = m.group(0)
                counters[fname] += 1
                n += 1
                samples.setdefault(fname, []).append((i, kind, extra, raw))
    if n == 0:
        return []
    dominant, dom_count = counters.most_common(1)[0]
    issues = []
    inconsistent_forms = {f: c for f, c in counters.items() if f != dominant and c > 0}
    if len(counters) > 1:
        issues.append(Issue(
            "日期写法", "全文",
            "、".join(f"{f}×{c}" for f, c in counters.most_common()),
            f"存在多种日期写法；主导为「{dominant}」（{dom_count} 处），其余共 {sum(inconsistent_forms.values())} 处",
            f"建议统一为「{dominant}」并列出少数派清单供修改"))

    # 分隔符内部一致性："."、"-"、"–" 是否混用
    seps = Counter()
    sep_samples = {}
    for i, (kind, text, extra) in enumerate(items):
        if not text:
            continue
        for m in RE_D_SEQ.finditer(text):
            sep = re.search(r"[．.\-/]", m.group(0))
            if sep:
                sp = "." if sep.group(0) in ".．" else sep.group(0)
                key = "点分隔(.)" if sp == "." else ("横线分隔(-)" if sp == "-" else f"其他({sp})")
                seps[key] += 1
                sep_samples.setdefault(key, (i, kind, extra, m.group(0)))
    if len(seps) > 1:
        issues.append(Issue("日期写法", "全文",
                            "、".join(f"{k}×{v}" for k, v in seps.most_common()),
                            "同为分隔符式日期但分隔符不统一", "建议统一分隔符"))

    # 少数派实例明细
    for fname, lst in samples.items():
        if fname == dominant:
            continue
        for (i, kind, extra, raw) in lst[:40]:
            loc = ("表格" if kind == "cell" else "正文") + (f"表{extra}" if kind == "cell" else f"第{i + 1}段")
            issues.append(Issue("日期写法", loc, raw,
                                f"非主导写法（主导：{dominant}）", "建议改为统一写法"))
    return issues


# ---------------------------------------------------------------- 4. 释义/简称

ABBR_DEF_RE = re.compile(
    r"([\u4e00-\u9fffA-Za-z0-9][\u4e00-\u9fffA-Za-z0-9·]{1,39}"
    r"(?:股份有限公司|有限公司|有限责任公司|公司|企业|集团|中心|基金|计划|银行|证券)?"
    r"\s*[（(]\s*以下简称\s*[「\"'『]?([^」\"'』()）]{1,25})[」\"'』]?\s*[）)])")


def check_abbr(doc_text):
    issues = []
    defs = {}       # 简称 -> [(全称, 出现位置offset)]
    quote_styles = Counter()
    for m in ABBR_DEF_RE.finditer(doc_text):
        full, short = m.group(1).strip(), m.group(2).strip()
        defs.setdefault(short, []).append((full, m.start()))
        qm = re.search(r"以下简称\s*([「\"'『])", m.group(0))
        if qm:
            quote_styles[qm.group(1)] += 1

    # 同一简称对应不同全称
    for short, pairs in sorted(defs.items()):
        uniq = sorted({p[0] for p in pairs})
        if len(uniq) > 1:
            first_loc = f"偏移{pairs[0][1]}"
            issues.append(Issue("释义简称", f"简称「{short}」（{first_loc} 起）",
                                " / ".join(uniq[:4]),
                                f"同名简称对应 {len(uniq)} 个不同全称", "请核实是否存在指向冲突"))
        # 定义之后全称复用
        defined_at = min(p[1] for p in pairs) + len(pairs[0][0])
        full_name = pairs[0][0]
        later_count = doc_text.count(full_name, defined_at)
        if later_count >= 3:
            tail = doc_text.find(full_name, defined_at)
            issues.append(Issue("释义简称", f"定义「{short}」之后（约偏移{tail}）",
                                full_name,
                                f"已定义简称，其后仍以全称出现 {later_count} 次",
                                f"建议统一切换为简称「{short}」"))

    # 引号风格混用
    if len(quote_styles) > 1:
        issues.append(Issue("释义简称", "全文",
                            "、".join(f"{k}×{v}" for k, v in quote_styles.most_common()),
                            "「以下简称」引号风格不统一", "建议全文统一引号风格"))
    return issues


# ---------------------------------------------------------------- 5. 中英文标点

CJK = r"\u4e00-\u9fff"

HALF_PUNCTS = ",.;:?!()\"'"
FULL_PUNCTS = "，。；：？！（）"


def _char_class(c):
    """C=中文；E=英文字母/数字；O=其他。"""
    if c and "\u4e00" <= c <= "\u9fff":
        return "C"
    if c and c.isascii() and (c.isalpha() or c.isdigit()):
        return "E"
    return "O"


def _neighbor_class(s, k, direction):
    """取位置 k 前/后第一个非空白字符的类别；越界返回 None。"""
    if direction < 0:
        rng = range(k - 1, -1, -1)
    else:
        rng = range(k + 1, len(s))
    for j in rng:
        c = s[j]
        if not c.isspace():
            return _char_class(c), j
    return None, None


def check_punctuation(items):
    """标点规则（2026-08-27 用户裁定：按前后字符类型判定）：
    - 半角标点 [,.;:?!()] 的相邻非空字符任一侧为中文 → 应为全角；
    - 全角标点 [，。；：？！] 相邻非空字符两侧均为英文/数字 → 应为半角；
    - 其余间隔场景默认中文标点，不报。
    引号直排形态单独提示。"""
    issues = []
    half_set = set(HALF_PUNCTS)
    full_set = set(FULL_PUNCTS)

    def add(kind, i, extra, s, start, end, problem, suggestion):
        ctx_l = max(0, start - 10)
        snippet = s[ctx_l:end + 10].replace("\n", "")
        loc = ("表格" if kind == "cell" else "正文") + \
              (f"表{extra}" if kind == "cell" else f"第{i + 1}段") + f"「…{snippet[:24]}…」"
        issues.append(Issue("中英标点", loc, s[start:end], problem, suggestion))

    for i, (kind, text, extra) in enumerate(items):
        if not text:
            continue
        # 规则一：半角标点任一侧为中文 → 应全角
        for k, ch in enumerate(text):
            if ch in half_set:
                pc, _pj = _neighbor_class(text, k, -1)
                nc, _nj = _neighbor_class(text, k, 1)
                if pc == "C" or nc == "C":
                    problem = f"中文语境使用半角标点「{ch}」"
                    full_map = {",": "，", ".": "。", ";": "；", ":": "：",
                                "?": "？", "!": "！", "(": "（", ")": "）", "\"": "“”"}
                    suggestion = f"改为全角「{full_map.get(ch, '对应全角标点')}」"
                    add(kind, i, extra, text, k, k + 1, problem, suggestion)

        # 规则二：全角标点两侧均为英文/数字 → 应半角
        for k, ch in enumerate(text):
            if ch in full_set:
                pc, pj = _neighbor_class(text, k, -1)
                nc, nj = _neighbor_class(text, k, 1)
                if pc == "E" and nc == "E":
                    half_map = {"，": ",", "。": ".", "；": ";", "：": ":",
                                "？": "?", "！": "!", "（": "(", "）": ")"}
                    problem = f"英文/数字之间误用全角标点「{ch}」"
                    suggestion = f"改为半角「{half_map.get(ch, ',')}」"
                    add(kind, i, extra, text, k, k + 1, problem, suggestion)

        # 提示级：中文语境直排双引号
        dq = re.compile(r'"')
        hits = list(dq.finditer(text))
        cn_near = any(
            (_char_class(text[m.start() - 1]) == "C") if m.start() > 0 else False
            for m in hits
        )
        if hits and cn_near:
            add(kind, i, extra, text, hits[0].start(), hits[0].end(),
                f"中文语境使用直排引号 \" （共 {len(hits)} 处）", "建议改用「」或全角弯引号“”")
    return issues


# ---------------------------------------------------------------- 6. 表格填写

NA_TOKENS = ["—", "－", "-", "/", "N/A", "N.A.", "不适用", "无"]


def check_tables(doc):
    issues = []
    stats = []
    for tn, tbl in enumerate(TBL_RE.finditer(doc), 1):
        tbl_xml = tbl.group(0)
        rows = re.findall(r"<w:tr\b[^>]*>.*?</w:tr>", tbl_xml, re.S)
        empty_cells = 0
        space_cells = 0
        na_counter = Counter()
        total_cells = 0
        for row in rows:
            cells = CELL_RE.findall(row)
            for c in cells:
                t = _text_of(c).strip()
                total_cells += 1
                raw_inner = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", c))
                if t == "":
                    empty_cells += 1
                elif raw_inner != raw_inner.strip():
                    space_cells += 1
                if t in NA_TOKENS:
                    na_counter[t] += 1
        notes = []
        if empty_cells:
            notes.append(f"空单元格 {empty_cells} 个")
        if space_cells:
            issues.append(Issue("表格填写", f"表{tn}", "",
                                f"{space_cells} 个单元格文本带首尾空格", "建议去除首尾空格"))
        used_na = {k: v for k, v in na_counter.items()}
        distinct_na = list(used_na.keys())
        if len(distinct_na) > 1:
            groups = {"长破折号": ["—"], "短横线": ["-", "－"], "斜杠": ["/"], "文字类": ["不适用", "N/A", "N.A.", "无"]}
            merged = {}
            for gname, tokens in groups.items():
                c = sum(v for k, v in used_na.items() if k in tokens)
                if c:
                    merged[gname] = (c, [k for k in distinct_na if k in tokens])
            if len(merged) > 1:
                issues.append(Issue("表格填写", f"表{tn}",
                                    "、".join(f"{g}{'/'.join(tk)}×{c}" for g, (c, tk) in merged.items()),
                                    "「不适用/无内容」标记符号在同一表内不统一",
                                    "建议全表统一一种标记（通常「—」或「不适用」）"))
        if notes:
            stats.append((tn, notes))
        elif total_cells:
            stats.append((tn, [f"共 {total_cells} 单元格，未见填写问题"]))
    return issues, stats


# ---------------------------------------------------------------- 主流程


def run_checks(input_path, checks=("1", "2", "3", "4", "5", "6")):
    xmls = load_docx(input_path)
    if not xmls:
        print("[ERROR] 无法读取 docx")
        return None
    doc = xmls["word/document.xml"]
    all_issues = []
    sections = {}

    def do(no, title, fn):
        if no in checks:
            res = fn()
            sections[f"{no}. {title}"] = res
            all_issues.extend(res)

    base = os.path.basename(input_path)

    def items_all():
        return [(k, t, x) for k, t, x in extract_structure(doc)]

    if "1" in checks:
        do("1", "标题层级序号（连续性/重号）", lambda: check_heading_sequence(items_all()))
    if "2" in checks:
        do("2", "金额数字（千分位/两位小数）", lambda: check_amounts(items_all()))
    if "3" in checks:
        do("3", "日期写法统一", lambda: check_dates(items_all()))
    if "4" in checks:
        full_text = "\n".join(_text_of(m.group(0)) for m in PARA_RE.finditer(doc)) + "\n" + \
                    "\n".join(_text_of(c.group(0)) for c in CELL_RE.finditer(doc))
        do("4", "释义/简称统一", lambda: check_abbr(full_text))
    if "5" in checks:
        do("5", "中英文标点", lambda: check_punctuation(items_all()))
    if "6" in checks:
        table_issues, table_stats = check_tables(doc)
        sections["6. 表格填写"] = table_issues
        for no_, notes_ in table_stats:
            if any("问题" in nn for nn in notes_):
                pass
        all_issues.extend(table_issues)
        sections["_table_stats"] = table_stats
    return base, sections, all_issues


def render_report(base, sections, all_issues):
    lines = [f"# 投行基本格式核对报告", ""]
    lines.append(f"- **核对对象**：`{base}`")
    lines.append("- **性质**：只读核对，未对文件做任何修改")
    lines.append("- 生成：ipo-doc-formatting `check_content.py`")
    lines.append("")
    lines.append("## 核对总览")
    lines.append("")
    lines.append("| 核对项 | 问题/提示条数 |")
    lines.append("|--------|--------------|")
    for name, res in sections.items():
        if name == "_table_stats":
            continue
        label = name.split(". ", 1)[-1]
        lines.append(f"| {label} | {len(res)} |")
    by_check = Counter(i.check for i in all_issues)
    lines.append("| **合计** | **{}** |".format(len(all_issues)))
    lines.append("")
    lines.append("> 说明：标注「疑似/请人工确认」的条目为机器初筛结果，需人工复核定性。")
    lines.append("")

    for name, res in sections.items():
        if name == "_table_stats":
            tlines = []
            for tn, notes in res:
                flagged = any("问题" in nn for nn in notes)
                flag = " ⚠️" if flagged else " ✅"
                tlines.append(f"- 表{tn}{flag}：" + "；".join(notes))
            if tlines:
                lines.append("### 表格逐表状态")
                lines.extend(tlines)
                lines.append("")
            continue
        lines.append(f"## {name}")
        lines.append("")
        if not res:
            lines.append("未发现问题。")
            lines.append("")
            continue
        lines.append("| # | 位置 | 原文/对象 | 问题 | 建议 |")
        lines.append("|---|------|----------|------|------|")
        for idx, it in enumerate(res, 1):
            esc = lambda s: s.replace("|", "\\|").replace("\n", " ")
            lines.append(f"| {idx} | {esc(it.location)} | {esc(it.snippet)} | "
                         f"{esc(it.problem)} | {esc(it.suggestion)} |")
        lines.append("")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="IPO 文档基本格式核对（只读，六大项）")
    ap.add_argument("--input", required=True, help="docx 文件路径（支持 glob）")
    ap.add_argument("--output", default=None, help="核对报告 md 输出路径（默认 <input>_格式核对报告.md）")
    ap.add_argument("--checks", default="123456",
                    help="启用哪些核对项，默认 123456：1标题序号 2金额 3日期 4释义简称 5中英标点 6表格")
    args = ap.parse_args()

    files = resolve_files(args.input)
    if not files:
        print(f"[ERROR] 未找到 docx: {args.input}")
        sys.exit(1)

    overall_ok = True
    for f in files:
        result = run_checks(f, checks=tuple(args.checks))
        if result is None:
            overall_ok = False
            continue
        base, sections, all_issues = result
        report = render_report(base, sections, all_issues)
        out = args.output or os.path.splitext(f)[0] + "_格式核对报告.md"
        with open(out, "w", encoding="utf-8") as fh:
            fh.write(report)
        print(f"\n=== 格式核对：{os.path.basename(f)} ===")
        for name, res in sections.items():
            if name == "_table_stats":
                continue
            mark = "✅" if not res else f"⚠️ {len(res)} 条"
            print(f"  {name}: {mark}")
        print(f"  合计问题/提示：{len(all_issues)} 条")
        # 控制台输出问题明细（每项最多 8 条）
        shown = 0
        for it in all_issues:
            if shown >= 24:
                print(f"  …… 其余 {len(all_issues) - shown} 条见报告")
                break
            note = f"（{it.suggestion}）" if it.suggestion else ""
            print(f"  [{it.check}] {it.location} | {it.snippet[:30]} | {it.problem}{note}")
            shown += 1
        print(f"  → 报告已写入：{out}")

    sys.exit(0 if overall_ok else 2)


def resolve_files(path):
    if os.path.isdir(path):
        return sorted(glob.glob(os.path.join(path, "*.docx")))
    if os.path.isfile(path):
        return [path]
    hits = glob.glob(path)
    return hits if hits else []


if __name__ == "__main__":
    main()
