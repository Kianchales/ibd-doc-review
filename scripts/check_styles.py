#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_styles.py — IPO 文档样式应用检查脚本（ipo-doc-audit skill 附带）
功能：
  - 成品文档模式（默认）：统计 docx 中 pStyle 使用分布，校验必备样式、裸段落、空段落（段落间禁止空行）、标题跳级。
  - 模板/样式库模式（--mode template）：校验 styles.xml 中 000-009 / 0011+001 / a4-a6 样式定义完整性。
用法：
  python check_styles.py --input <docx路径或目录> --scenario 招股书|反馈回复|报告 [--mode document|template]
  python check_styles.py --input <套样式后.docx> --verify-content <原文.docx>   # 内容完整性校验（严禁修改原文内容）
说明：
  - 只依赖标准库 zipfile/re，无第三方依赖。
  - 中文文件名：优先用 glob 兜底（Git Bash 直传中文参数可能乱码）。
  - 场景决定「一级样式」：招股书/报告 → 001；反馈回复 → 0011+001（监管问题黑体）。
"""
import argparse
import datetime
import glob
import os
import re
import sys
import zipfile

# 各场景必备样式：正文 + 一级（必检）；002 及以上由跳级检测兜底
REQUIRED = {
    "招股书": ["000", "001"],
    "报告": ["000", "001"],
    "反馈回复": ["000", "0011", "001"],
}
# 标题样式→大纲级别（跳级检测）
OUTLINE = {"001": 0, "0011": 0, "002": 1, "003": 2, "004": 3, "005": 4, "006": 5, "007": 6}
HEADING_STYLES = set(OUTLINE.keys())
# 样式库模式：styles.xml 应包含的样式 ID
LIBRARY = ["000", "001", "002", "003", "004", "005", "006", "007", "008", "009", "a4", "a5", "a6"]
LIBRARY_EXTRA = {"反馈回复": ["0011", "001"]}
# 序号开头段落判定（双维度算法）：序号前缀模式
# 注意：半角点须后跟空白（区分 "1. 标题" 与 "78.50" 小数）；表格内数字另行排除
NUMBER_PAT = re.compile(
    r"^[（(]\s*[一二三四五六七八九十百\d]+\s*[）)]"   # （一）（1）(一)(1)
    r"|^\d+\s*[、．]"                                # 1、1．
    r"|^\d+\.\s"                                    # 1. 标题（半角点+空白）
    r"|^[①-⑳]"                                      # ①-⑳
    r"|^[A-Za-z]\s*[、．.]"                          # A、a．a.
)


def resolve_files(path):
    """输入可为单个 docx 或目录；返回 docx 文件列表。"""
    if os.path.isdir(path):
        return sorted(glob.glob(os.path.join(path, "*.docx")))
    if os.path.isfile(path):
        return [path]
    hits = glob.glob(path)
    return hits if hits else []


def load_docx(docx_path):
    """读取 docx 内 XML，返回 dict(name->xml) 或 None。"""
    try:
        with zipfile.ZipFile(docx_path) as z:
            names = set(z.namelist())
            out = {}
            for key in ("word/document.xml", "word/styles.xml"):
                if key in names:
                    out[key] = z.read(key).decode("utf-8", errors="ignore")
            return out
    except Exception as e:
        print(f"  [ERROR] 读取失败: {e}")
        return None


def check_document(docx_path, scenario):
    print(f"\n=== {os.path.basename(docx_path)}（成品文档 · 场景: {scenario}）===")
    xmls = load_docx(docx_path)
    if not xmls or "word/document.xml" not in xmls:
        print("  [ERROR] 非标准 docx（无 word/document.xml）")
        return False
    doc = xmls["word/document.xml"]

    stats = {}
    bare = []
    heading_seq = []
    empty = []  # 空段落（段落间禁止空行/空段落）

    # 标记表格单元格（<w:tc>）内的段落起始位置——表格单元格内容走表格样式（a6/直接格式），
    # 不套 000-009 正文样式，不应计入「裸段落」误报；表格内空段亦不算违规
    table_p_offsets = set()
    for tc in re.finditer(r"<w:tc\b[^>]*>.*?</w:tc>", doc, re.S):
        for pm in re.finditer(r"<w:p\b", tc.group(0)):
            table_p_offsets.add(tc.start() + pm.start())

    for m in re.finditer(r"<w:p\b[^>]*>(.*?)</w:p>", doc, re.S):
        body = m.group(1)
        st = re.search(r'<w:pStyle w:val="([^"]+)"', body)
        style = st.group(1) if st else None
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body)).strip()
        stats[style] = stats.get(style, 0) + 1
        if style in HEADING_STYLES:
            heading_seq.append(style)
        if style is None and text and m.start() not in table_p_offsets:
            bare.append(text[:50])
        # 空段落：无任何文本内容（表格外）——间距应由样式 spacing 控制，空段落必须删除
        if not text and m.start() not in table_p_offsets:
            empty.append(m.start())

    # 自闭合空段落 <w:p/>（主循环正则只匹配配对形式，此处补检）
    for sc in re.finditer(r"<w:p\s*/>", doc):
        if sc.start() not in table_p_offsets:
            empty.append(sc.start())

    print("  pStyle 分布：")
    if not stats:
        print("    （无任何已命名样式段落 —— 未套用样式体系）")
    for s in sorted(stats, key=lambda x: (x is None, x)):
        print(f"    {s if s else '(裸)'}: {stats[s]}")

    ok = True
    req = REQUIRED.get(scenario, REQUIRED["报告"])
    missing = [s for s in req if stats.get(s, 0) == 0]
    if missing:
        print(f"  [FAIL] 缺少必备样式: {missing}")
        ok = False
    else:
        print(f"  [PASS] 必备样式齐全: {req}")

    if bare:
        print(f"  [WARN] {len(bare)} 个裸段落（未应用样式，建议补 pStyle）：")
        for b in bare[:5]:
            print(f"    - {b}")
    else:
        print("  [PASS] 无裸正文段落")

    # 空段落检查（段落间禁止空行/空段落）
    if empty:
        print(f"  [WARN] {len(empty)} 个空段落（无文字内容，禁止——"
              f"间距由样式 spacing 控制，空段落应删除）：")
    else:
        print("  [PASS] 无空段落（段落间距由样式 spacing 控制）")

    jumps = []
    for i in range(1, len(heading_seq)):
        prev, cur = heading_seq[i - 1], heading_seq[i]
        if prev in OUTLINE and cur in OUTLINE and OUTLINE[cur] > OUTLINE[prev] + 1:
            jumps.append((prev, cur))
    if jumps:
        print(f"  [WARN] 标题跳级: {jumps}（如 002 后直接 004，检查是否漏了 003）")
    else:
        print("  [PASS] 标题层级连续" + ("（无标题段落）" if not heading_seq else ""))

    print(f"  => {'通过' if ok else '需修正（见上方 FAIL/WARN）'}")
    return ok


def check_template(docx_path, scenario):
    print(f"\n=== {os.path.basename(docx_path)}（样式库 · 场景: {scenario}）===")
    xmls = load_docx(docx_path)
    if not xmls or "word/styles.xml" not in xmls:
        print("  [ERROR] 非标准 docx（无 word/styles.xml）")
        return False
    styles_xml = xmls["word/styles.xml"]
    present = set(re.findall(r'<w:style [^>]*w:styleId="([^"]+)"', styles_xml))

    need = list(LIBRARY) + LIBRARY_EXTRA.get(scenario, [])
    missing = [s for s in need if s not in present]
    if missing:
        print(f"  [FAIL] 样式库缺样式: {missing}")
        return False
    print(f"  [PASS] 样式库完整（{len(need)} 个样式全部存在）: {need}")
    return True


def extract_text(docx_path):
    """提取 docx 全部可见文本（按 <w:t> 顺序，含表格内文字），用于内容完整性对比。
    过滤空文本：合并单元格（gridSpan/vMerge）会产生空 <w:t> 占位，不算内容。"""
    xmls = load_docx(docx_path)
    if not xmls or "word/document.xml" not in xmls:
        return None
    doc = xmls["word/document.xml"]
    return [m.group(1) for m in re.finditer(r"<w:t[^>]*>([^<]*)</w:t>", doc) if m.group(1)]


def extract_para_texts(docx_path):
    """提取段落级拼接文本序列（每段拼接全部 <w:t>），用于内容完整性对比。
    与 run 结构无关：merge-runs 等 run 合并操作不影响对比结果（2026-08-27 修复：
    原 extract_text 按 <w:t> 逐 run 对比，run 合并后误报内容被修改）。
    过滤空段落（无文字内容）；表格内段落同样按 <w:p> 顺序提取。"""
    xmls = load_docx(docx_path)
    if not xmls or "word/document.xml" not in xmls:
        return None
    doc = xmls["word/document.xml"]
    texts = []
    for m in re.finditer(r"<w:p\b[^>]*>(.*?)</w:p>", doc, re.S):
        body = m.group(1)
        t = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body))
        if t.strip():
            texts.append(t)
    return texts


def _paras_with_xml(doc):
    """解析 document.xml 的段落，返回 [(完整段落XML, pPr片段或None, pStyle或None, 文本)]。"""
    out = []
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", doc, re.S):
        body = m.group(0)
        ppr_m = re.search(r"<w:pPr\b[^>]*>.*?</w:pPr>", body, re.S)
        ppr = ppr_m.group(0) if ppr_m else None
        style_m = re.search(r'<w:pStyle w:val="([^"]+)"', ppr or "")
        style = style_m.group(1) if style_m else None
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body)).strip()
        out.append((body, ppr, style, text))
    return out


def _rebuild_para_revision(body, new_style, author, date, rev_id):
    """把段落 body 的 pStyle 改为 new_style，并插入 w:pPrChange（快照原格式）形成 Word 修订。"""
    ppr_m = re.search(r"<w:pPr\b[^>]*>.*?</w:pPr>", body, re.S)
    if ppr_m:
        old_ppr = ppr_m.group(0)
        old_style_m = re.search(r'<w:pStyle w:val="([^"]+)"', old_ppr)
        if old_style_m:
            new_ppr = old_ppr[:old_style_m.start(1)] + new_style + old_ppr[old_style_m.end(1):]
        else:
            # 无 pStyle：在 pPr 开标签后插入 pStyle（2026-08-27 修复：原 replace("<w:pPr",...) 
            # 漏掉开标签的 ">"，导致 <w:pStyle/> 后残留多余 ">"，真实文档裸段落（有 pPr 无 pStyle）必现）
            open_m = re.search(r"<w:pPr\b[^>]*>", old_ppr)
            if open_m:
                new_ppr = old_ppr[:open_m.end()] + f'<w:pStyle w:val="{new_style}"/>' + old_ppr[open_m.end():]
            else:
                new_ppr = old_ppr
        change_snapshot = re.sub(r"<w:rPr>.*?</w:rPr>", "", old_ppr, flags=re.S)
        change = (f'<w:pPrChange w:id="{rev_id}" w:author="{author}" w:date="{date}">'
                  f"{change_snapshot}</w:pPrChange>")
        new_ppr = new_ppr[: -len("</w:pPr>")] + change + "</w:pPr>"
        return body[: ppr_m.start()] + new_ppr + body[ppr_m.end():]
    new_ppr = (f'<w:pPr><w:pStyle w:val="{new_style}"/>'
               f'<w:pPrChange w:id="{rev_id}" w:author="{author}" w:date="{date}">'
               f"<w:pPr/></w:pPrChange></w:pPr>")
    # 无原 pPr：把新 pPr 插入 <w:p> 开标签之后
    open_m = re.match(r"(<w:p\b[^>]*>)", body)
    if open_m:
        return body[: open_m.end()] + new_ppr + body[open_m.end():]
    return new_ppr + body


def make_revision(orig_path, styled_path, output_path):
    """生成 Word 修订稿（2026-08-27）：以原文档为基底，样式化文档为参照，
    对每个样式变化段落插入 w:pPrChange 修订（格式更改），并开启 trackChanges。
    打开后可在审阅面板看到「格式更改」修订，可接受/拒绝。"""
    orig_xmls = load_docx(orig_path)
    if not orig_xmls or "word/document.xml" not in orig_xmls:
        print(f"  [ERROR] 读取原文件失败: {orig_path}")
        return False
    orig_doc = orig_xmls["word/document.xml"]
    styled_doc = load_docx(styled_path).get("word/document.xml", "") if load_docx(styled_path) else ""

    o_paras = _paras_with_xml(orig_doc)
    s_paras = _paras_with_xml(styled_doc)
    # 非空文本对齐（内容未变，只改格式）
    o_seq = [(i, t) for i, (_, _, _, t) in enumerate(o_paras) if t]
    s_seq = [(i, t) for i, (_, _, _, t) in enumerate(s_paras) if t]
    if [t for _, t in o_seq] != [t for _, t in s_seq]:
        print("  [FAIL] 文本序列不一致——内容被修改（严禁修改原文内容）")
        return False

    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev_id = 0
    n_changes = 0
    for (oi, ot), (si, st) in zip(o_seq, s_seq):
        o_style = o_paras[oi][2]
        s_style = s_paras[si][2]
        if o_style != s_style:
            new_style = s_style if s_style else ""
            rev_id += 1
            body = o_paras[oi][0]
            new_body = _rebuild_para_revision(body, new_style, "ipo-doc-audit", now, rev_id)
            orig_doc = orig_doc.replace(body, new_body, 1)
            n_changes += 1

    # 开启修订模式（settings.xml 加 trackRevisions + revisionView；load_docx 不含 settings，直接读 zip）
    settings = ""
    with zipfile.ZipFile(orig_path) as z:
        if "word/settings.xml" in z.namelist():
            settings = z.read("word/settings.xml").decode("utf-8")
    if settings:
        # 2026-08-27 三次修复（OpenXmlValidator 实证）：
        #   a) 正确元素名是 w:trackRevisions，不是 w:trackChanges——后者在 CT_Settings schema 中
        #      根本不存在，Word 解析时忽略/容错，修订记录与显示全部失效；
        #   b) 合法位置是 w:bordersDoNotSurroundFooter 之后（CT_Settings 序列：
        #      ...bordersHeader→bordersFooter→revisionView→trackRevisions→defaultTabStop...），
        #      插到其他任何位置 OpenXmlValidator 均报 unexpected/invalid child；
        #   c) revisionView 必须带 w:formatting="1"，否则打开时格式更改标记默认隐藏。
        rv = ('<w:revisionView w:markup="1" w:comments="1" w:insDel="1"'
              ' w:formatting="1" w:inkAnnotations="1"/>')
        insert_block = ""
        if "<w:revisionView" not in settings:
            insert_block += rv
        if "<w:trackRevisions" not in settings:
            insert_block += "<w:trackRevisions/>"
        if insert_block:
            m = re.search(r"<w:bordersDoNotSurroundFooter\b[^>]*/>", settings) or \
                re.search(r"<w:bordersDoNotSurroundFooter\b[^>]*>.*?</w:bordersDoNotSurroundFooter>",
                          settings, re.S)
            if m:
                settings = settings[: m.end()] + insert_block + settings[m.end():]
            else:
                # 兜底：无 bordersFooter 时追加末尾（尽力而为，标准文档均有该元素）
                settings = settings.replace("</w:settings>", insert_block + "</w:settings>", 1)

    # 写回新 docx（复用原文件全部部件，替换 document.xml/settings.xml）
    try:
        with zipfile.ZipFile(orig_path) as zin:
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.namelist():
                    data = zin.read(item)
                    if item == "word/document.xml":
                        data = orig_doc.encode("utf-8")
                    elif item == "word/settings.xml" and settings:
                        data = settings.encode("utf-8")
                    zout.writestr(item, data)
    except Exception as e:
        print(f"  [ERROR] 写出修订稿失败: {e}")
        return False

    print(f"  → 修订稿已生成：{output_path}（{n_changes} 处格式更改修订，Word 审阅可接受/拒绝）")
    return True


def extract_paras(docx_path):
    """提取段落序列 [(样式, 文本, 是否表格内, 全局序号)]，用于 diff 对比。"""
    xmls = load_docx(docx_path)
    if not xmls or "word/document.xml" not in xmls:
        return None
    doc = xmls["word/document.xml"]
    table_p_offsets = set()
    for tc in re.finditer(r"<w:tc\b[^>]*>.*?</w:tc>", doc, re.S):
        for pm in re.finditer(r"<w:p\b", tc.group(0)):
            table_p_offsets.add(tc.start() + pm.start())
    paras = []
    for idx, m in enumerate(re.finditer(r"<w:p\b[^>]*>(.*?)</w:p>", doc, re.S)):
        body = m.group(1)
        st = re.search(r'<w:pStyle w:val="([^"]+)"', body)
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body)).strip()
        paras.append((st.group(1) if st else None, text, m.start() in table_p_offsets, idx))
    return paras


STYLE_DESC = {
    "000": "000 正文（宋体12pt，首行缩进2字符，1.5倍行距）",
    "001": "001 一级标题（黑体16pt，居中，分页前）",
    "0011": "0011 问题一级标题（黑体16pt，两端对齐）",
    "001q": "001 监管问题正文（黑体，缩进）",
    "002": "002 二级标题（黑体14pt）",
    "003": "003 三级标题（黑体12pt）",
    "004": "004 四级标题（加粗12pt）",
    "005": "005 五级标题（加粗12pt）",
    "006": "006 六级标题（默认12pt）",
    "007": "007 七级标题（默认12pt）",
    "008": "008 表格前单位行（右对齐，五号）",
    "009": "009 表格后备注行（五号，首行缩进）",
}


def style_label(style):
    if not style:
        return "裸段落（无 pStyle）"
    return STYLE_DESC.get(style, f"样式 {style}")


def cmd_diff(new_path, orig_path):
    """格式修改 diff（2026-08-27 交付物能力）：对比原文件与修正稿，
    生成「格式问题清单」（位置/原文/改成什么）+ 修改统计，输出控制台与 md 文件。
    前提：只改格式不改内容（文本序列应逐字一致）。"""
    print(f"\n=== 格式修改清单（{os.path.basename(new_path)} vs 原文 {os.path.basename(orig_path)}）===")
    p_orig = extract_paras(orig_path)
    p_new = extract_paras(new_path)
    if not p_orig or not p_new:
        print("  [ERROR] 文件读取失败，无法对比")
        return False

    # 空段落不计入内容（合并单元格等会产生空占位，非内容）；仅非空段落参与文本一致性
    non_empty_orig = [(s, t) for s, t, _, _ in p_orig if t]
    non_empty_new = [(s, t) for s, t, _, _ in p_new if t]
    t_orig = [t for _, t in non_empty_orig]
    t_new = [t for _, t in non_empty_new]
    if t_orig != t_new:
        print("  [FAIL] 文本序列不一致——内容被修改（严禁修改原文内容，铁律 0）")
        for i, (a, b) in enumerate(zip(t_orig, t_new)):
            if a != b:
                print(f"    第 {i + 1} 处：原文「{a[:30]}」 vs 现文「{b[:30]}」")
                break
        return False

    n_empty_orig = sum(1 for _, t, _, _ in p_orig if not t)
    n_empty_new = sum(1 for _, t, _, _ in p_new if not t)
    empty_delta = n_empty_orig - n_empty_new  # >0 删除了空段落

    changes = []  # (位置, 原文, 改成)
    for i, ((so, to), (sn, tn)) in enumerate(zip(non_empty_orig, non_empty_new)):
        if to != tn:
            continue  # 内容变化已 FAIL
        if so != sn:
            pos = f"第{i + 1}段「{to[:16]}」"
            changes.append((pos, style_label(so), style_label(sn)))
    if empty_delta > 0:
        changes.append(("全文（空段落）", f"空段落 {empty_delta} 个", "删除（段落间距由样式 spacing 控制）"))

    n_tbl_orig = len(re.findall(r"<w:tbl\b", _doc_xml(orig_path)))
    n_tbl_new = len(re.findall(r"<w:tbl\b", _doc_xml(new_path)))
    tbl_note = ""
    if n_tbl_orig or n_tbl_new:
        tbl_note = f"；表格 {n_tbl_orig}→{n_tbl_new} 张，样式已按模板 v2（100%宽/两级表头/纯黑边框/数字右对齐/合计加粗）应用"

    # 统计分类
    n_heading = sum(1 for _, _, s in changes if s and s.split()[0] in ("001", "0011", "002", "003", "004", "005", "006", "007", "001q"))
    n_body = sum(1 for _, _, s in changes if s and s.split()[0] in ("000", "008", "009"))
    n_bare = sum(1 for _, o, _ in changes if "裸段落" in o)
    n_total = len(changes)

    # 控制台输出
    print(f"  修改段落：{n_total} 处（标题样式 {n_heading} / 正文样式 {n_body} / 裸段落样式化 {n_bare} / 空段落处理 {max(empty_delta, 0)} 处）{tbl_note}")
    print(f"  总段落：{len(p_orig)}；文本一致性：{'PASS（内容未改动）' if t_orig == t_new else 'FAIL'}")

    # 输出 md 清单
    md_path = os.path.splitext(new_path)[0] + "_格式修改清单.md"
    lines = ["# 格式修改清单", ""]
    lines.append(f"- **修正稿**：`{os.path.basename(new_path)}`")
    lines.append(f"- **原文**：`{os.path.basename(orig_path)}`")
    lines.append(f"- 生成：ipo-doc-audit `check_styles.py --diff`（只改格式，不改内容）")
    lines.append("")
    lines.append("## 修改统计")
    lines.append("")
    lines.append("| 项 | 数量 |")
    lines.append("|----|------|")
    lines.append(f"| 总段落数 | {len(p_orig)} |")
    lines.append(f"| 修改段落合计 | **{n_total}** |")
    lines.append(f"| 其中：标题样式 | {n_heading} |")
    lines.append(f"| 正文样式 | {n_body} |")
    if empty_delta > 0:
        lines.append(f"| 空段落删除 | {empty_delta} |")
    if n_tbl_orig or n_tbl_new:
        lines.append(f"| 表格 | {n_tbl_orig} → {n_tbl_new} 张（样式按模板 v2 应用） |")
    lines.append(f"| 文本内容 | 与原文逐字一致（未增删改） |")
    lines.append("")
    lines.append("## 修改明细")
    lines.append("")
    lines.append("| # | 位置 | 原文 | 改成 |")
    lines.append("|---|------|------|------|")
    for idx, (pos, old, new) in enumerate(changes, 1):
        old_c = old.replace("|", "\\|")
        new_c = new.replace("|", "\\|")
        lines.append(f"| {idx} | {pos} | {old_c} | {new_c} |")
    lines.append("")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"  → 清单已写入：{md_path}")
    return True


def _doc_xml(docx_path):
    xmls = load_docx(docx_path)
    return xmls.get("word/document.xml", "") if xmls else ""


def check_numbering(docx_path):
    """序号段落核对（双维度算法辅助）：列出所有序号开头段落及上下文，
    供按 rules.md 判定「标题 vs 正文」：
      情况1（标题）：序号段短 + 后段独立正文展开 → 序号段用标题样式，后段 000
      情况2（列举正文）：序号段长 + 后段连续序号/无后段 → 序号段本身用 000
    输出为核对清单（WARN 级别，需人工按上下文定夺，不自动 FAIL）。"""
    print(f"\n=== 序号段落核对（{os.path.basename(docx_path)}）===")
    xmls = load_docx(docx_path)
    if not xmls or "word/document.xml" not in xmls:
        print("  [ERROR] 非标准 docx（无 word/document.xml）")
        return False
    doc = xmls["word/document.xml"]

    # 排除表格内段落（表格内容非文档序号结构，且数字/百分比易误配）
    table_p_offsets = set()
    for tc in re.finditer(r"<w:tc\b[^>]*>.*?</w:tc>", doc, re.S):
        for pm in re.finditer(r"<w:p\b", tc.group(0)):
            table_p_offsets.add(tc.start() + pm.start())

    paras = []
    for m in re.finditer(r"<w:p\b[^>]*>(.*?)</w:p>", doc, re.S):
        body = m.group(1)
        st = re.search(r'<w:pStyle w:val="([^"]+)"', body)
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body)).strip()
        paras.append((st.group(1) if st else None, text, m.start() in table_p_offsets))

    hits = []
    for i, (style, text, in_table) in enumerate(paras):
        if in_table:
            continue
        if text and NUMBER_PAT.match(text):
            hits.append((i, style, text))

    if not hits:
        print("  [INFO] 未发现序号开头段落（（一）/1、/（1）/① 等）")
        return True

    print(f"  共 {len(hits)} 个序号开头段落（按双维度算法核对：长短 + 后段是否分段）:")
    print("  # | 样式 | 字数 | 序号段落(前24字) | 后段(前24字) | 倾向")
    for idx, (i, style, text) in enumerate(hits, 1):
        nxt = paras[i + 1][1] if i + 1 < len(paras) else ""
        nxt_is_num = bool(nxt and NUMBER_PAT.match(nxt))
        short = len(text) <= 40
        if not nxt:
            lean = "情况2:列举正文(无后段)" if not short else "短句无后段→倾向正文"
        elif nxt_is_num:
            lean = "情况2:列举正文(后段连续序号)" if not short else "短句+连续序号→倾向正文"
        else:
            lean = "情况1:标题+展开(后段独立正文)" if short else "长句+后段展开→上下文主判,按分段定"
        style_s = style if style else "(裸)"
        print(f"  {idx:2d} | {style_s:6s} | {len(text):3d} | {text[:24]} | {nxt[:24]} | {lean}")
    print("  [HINT] 判定主判据=上下文是否分段；长短为辅助。详见 rules.md「序号段落判定」。")
    return True


def verify_content(docx_path, original_path):
    """内容完整性检查（只改格式、严禁修改原文内容）。
    套样式后的文本必须与原文逐字一致（含表格内文字、标点、空格、数字）。"""
    print(f"\n=== 内容完整性校验（{os.path.basename(docx_path)} vs 原文 {os.path.basename(original_path)}）===")
    t_orig = extract_para_texts(original_path)
    t_new = extract_para_texts(docx_path)
    if t_orig is None or t_new is None:
        print("  [ERROR] 文件读取失败，无法校验")
        return False
    if t_orig == t_new:
        print(f"  [PASS] 内容完整：{len(t_new)} 段文本与原文逐字一致（未增删改任何文字）")
        return True
    for i, (a, b) in enumerate(zip(t_orig, t_new)):
        if a != b:
            print(f"  [FAIL] 第 {i + 1} 处文本被修改：")
            print(f"    原文: {a[:50]}")
            print(f"    现文: {b[:50]}")
            break
    else:
        if len(t_orig) != len(t_new):
            print(f"  [FAIL] 文本段数量不一致：原文 {len(t_orig)} 段 → 现文 {len(t_new)} 段（有增删）")
    print("  [FAIL] 严禁修改原文内容——打回重做")
    return False


def main():
    ap = argparse.ArgumentParser(description="IPO 文档样式应用检查")
    ap.add_argument("--input", required=True, help="docx 文件路径或目录（支持 glob）")
    ap.add_argument("--scenario", choices=["招股书", "报告", "反馈回复"], default="报告")
    ap.add_argument("--mode", choices=["document", "template"], default="document",
                    help="document=成品文档校验（默认）；template=样式库完整性校验")
    ap.add_argument("--verify-content", metavar="原文.docx", default=None,
                    help="内容完整性校验：对比 --input 与原文文本是否逐字一致（严禁修改原文内容）")
    ap.add_argument("--check-numbering", action="store_true",
                    help="序号段落核对：列出序号开头段落（（一）/1、/（1）/①）及上下文，辅助判定标题 vs 正文")
    ap.add_argument("--diff", metavar="原文.docx", default=None,
                    help="格式修改 diff：对比 --input（修正稿）与原文，生成格式问题清单（位置/原文/改成什么）+ 统计，写入 <修正稿>_格式修改清单.md")
    ap.add_argument("--revise", metavar="原文件.docx", default=None,
                    help="生成 Word 修订稿：以原文件为基底，将 --input（样式化结果）的格式改动转为 Word 修订（w:pPrChange 格式更改），并开启 trackChanges；输出 <原文件>_修订稿.docx（--output 可覆盖）")
    ap.add_argument("--output", default=None, help="--revise 的输出路径（默认 <原文件>_修订稿.docx）")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    files = resolve_files(args.input)
    if not files:
        print(f"[ERROR] 未找到 docx: {args.input}")
        sys.exit(1)

    if args.diff:
        orig_files = resolve_files(args.diff)
        if not orig_files:
            print(f"[ERROR] 未找到原文: {args.diff}")
            sys.exit(1)
        all_ok = all(cmd_diff(f, orig_files[0]) for f in files)
        print(f"\n{'格式修改清单生成完成 ✅' if all_ok else '生成失败 ❌'}")
        sys.exit(0 if all_ok else 2)

    if args.check_numbering:
        all_ok = all(check_numbering(f) for f in files)
        print(f"\n{'序号段落核对完成 ✅' if all_ok else '核对异常 ❌'}")
        sys.exit(0 if all_ok else 2)

    if args.verify_content:
        orig_files = resolve_files(args.verify_content)
        if not orig_files:
            print(f"[ERROR] 未找到原文: {args.verify_content}")
            sys.exit(1)
        # 内容校验：--input 每个文件与原文对比（单文件对单文件）
        all_ok = True
        for f in files:
            all_ok &= verify_content(f, orig_files[0])
        print(f"\n{'内容完整性全部通过 ✅' if all_ok else '内容被修改 ❌（严禁修改原文内容）'}")
        sys.exit(0 if all_ok else 2)

    if args.revise:
        orig_files = resolve_files(args.revise)
        if not orig_files:
            print(f"[ERROR] 未找到原文件: {args.revise}")
            sys.exit(1)
        print(f"\n=== 生成 Word 修订稿（--input 为样式化结果，--revise 为原文件）===")
        all_ok = True
        for f in files:
            out = args.output or os.path.splitext(orig_files[0])[0] + "_修订稿.docx"
            all_ok &= make_revision(orig_files[0], f, out)
        print(f"\n{'修订稿生成完成 ✅' if all_ok else '生成失败 ❌'}")
        sys.exit(0 if all_ok else 2)

    fn = check_template if args.mode == "template" else check_document
    all_ok = all(fn(f, args.scenario) for f in files)
    print(f"\n{'全部通过 ✅' if all_ok else '存在需修正项 ❌'}")
    sys.exit(0 if all_ok else 2)


if __name__ == "__main__":
    main()
