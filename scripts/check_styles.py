#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_styles.py — IPO 文档样式应用检查脚本（ipo-doc-formatting skill 附带）
功能：
  - 成品文档模式（默认）：统计 docx 中 pStyle 使用分布，校验必备样式、裸段落、空段落（2026-08-25 裁定：段落间禁止空行）、标题跳级。
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
    empty = []  # 空段落（2026-08-25 用户裁定：段落间禁止空行/空段落）

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

    # 空段落检查（2026-08-25 用户裁定：段落间禁止空行/空段落）
    if empty:
        print(f"  [WARN] {len(empty)} 个空段落（无文字内容，2026-08-25 裁定禁止——"
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


def verify_content(docx_path, original_path):
    """内容完整性检查（2026-08-25 用户裁定：只改格式、严禁修改原文内容）。
    套样式后的文本必须与原文逐字一致（含表格内文字、标点、空格、数字）。"""
    print(f"\n=== 内容完整性校验（{os.path.basename(docx_path)} vs 原文 {os.path.basename(original_path)}）===")
    t_orig = extract_text(original_path)
    t_new = extract_text(docx_path)
    if t_orig is None or t_new is None:
        print("  [ERROR] 文件读取失败，无法校验")
        return False
    if t_orig == t_new:
        print(f"  [PASS] 内容完整：{len(t_new)} 段文本与原文逐字一致（未增删改任何文字）")
        return True
    # 定位首个差异
    for i, (a, b) in enumerate(zip(t_orig, t_new)):
        if a != b:
            print(f"  [FAIL] 第 {i + 1} 处文本被修改：")
            print(f"    原文: {a[:50]}")
            print(f"    现文: {b[:50]}")
            break
    else:
        if len(t_orig) != len(t_new):
            print(f"  [FAIL] 文本段数量不一致：原文 {len(t_orig)} 段 → 现文 {len(t_new)} 段（有增删）")
    print("  [FAIL] 严禁修改原文内容（2026-08-25 用户裁定）——打回重做")
    return False


def main():
    ap = argparse.ArgumentParser(description="IPO 文档样式应用检查")
    ap.add_argument("--input", required=True, help="docx 文件路径或目录（支持 glob）")
    ap.add_argument("--scenario", choices=["招股书", "报告", "反馈回复"], default="报告")
    ap.add_argument("--mode", choices=["document", "template"], default="document",
                    help="document=成品文档校验（默认）；template=样式库完整性校验")
    ap.add_argument("--verify-content", metavar="原文.docx", default=None,
                    help="内容完整性校验：对比 --input 与原文文本是否逐字一致（严禁修改原文内容）")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    files = resolve_files(args.input)
    if not files:
        print(f"[ERROR] 未找到 docx: {args.input}")
        sys.exit(1)

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

    fn = check_template if args.mode == "template" else check_document
    all_ok = all(fn(f, args.scenario) for f in files)
    print(f"\n{'全部通过 ✅' if all_ok else '存在需修正项 ❌'}")
    sys.exit(0 if all_ok else 2)


if __name__ == "__main__":
    main()
