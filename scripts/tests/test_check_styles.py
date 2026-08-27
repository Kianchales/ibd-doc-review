#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_check_styles.py — ibd-doc-review 校验脚本自测（标准流程强化）
覆盖：
  1. 合格反馈回复文档 → 通过（PASS）
  2. 不合格文档（无 pStyle）→ 失败（FAIL）
  3. 空段落检测（自闭合 <w:p/> 与配对 <w:p></w:p>）→ 检出（WARN）
  4. 模板/样式库模式（反馈回复 15 样式 / 报告 13 样式）→ 通过（PASS）
运行：python scripts/tests/test_check_styles.py
"""
import os
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile

SKILL_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
SCRIPT = os.path.join(SKILL_DIR, "scripts", "check_styles.py")
TEMPLATE_DIR = os.path.join(SKILL_DIR, "assets", "templates")

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def make_mini_docx(path, paras):
    """paras: [(style|None, text)]，生成最小 docx（仅 document.xml + 必需部件）。
    pStyle 按真实 Word 文档形态置于 w:pPr 内。"""
    body = []
    for style, text in paras:
        pstyle = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
        body.append(f"<w:p>{pstyle}<w:r><w:t>{text}</w:t></w:r></w:p>")
    doc = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}><w:body>{"".join(body)}<w:sectPr/></w:body></w:document>'
    )
    ct = (
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    settings = (
        '<?xml version="1.0"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:zoom w:percent="100"/><w:bordersDoNotSurroundFooter/>'
        '<w:defaultTabStop w:val="420"/></w:settings>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/settings.xml", settings)


def run_check(path, scenario="反馈回复", mode="document"):
    cmd = [sys.executable, SCRIPT, "--input", path, "--scenario", scenario, "--mode", mode]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode, proc.stdout


def run_verify(new_path, orig_path):
    """内容完整性校验（严禁修改原文内容）。"""
    cmd = [sys.executable, SCRIPT, "--input", new_path, "--verify-content", orig_path]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode, proc.stdout


CONTENT_SCRIPT = os.path.join(os.path.dirname(SCRIPT), "check_content.py")


def read_content_report(path):
    """读取 <path>_格式核对报告.md 内容。"""
    rp = os.path.splitext(path)[0] + "_格式核对报告.md"
    if os.path.exists(rp):
        with open(rp, encoding="utf-8") as f:
            return f.read()
    return ""


def run_content(path, checks=None):
    """投行格式核对（check_content.py，只读六大项）。"""
    cmd = [sys.executable, CONTENT_SCRIPT, "--input", path]
    if checks:
        cmd += ["--checks", checks]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode, proc.stdout


def run_diff(new_path, orig_path):
    """格式修改 diff（2026-08-27 交付物能力：清单+统计）。"""
    cmd = [sys.executable, SCRIPT, "--input", new_path, "--diff", orig_path]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode, proc.stdout


def run_revise(styled_path, orig_path, output_path):
    """生成 Word 修订稿（2026-08-27：w:pPrChange + trackChanges）。"""
    cmd = [sys.executable, SCRIPT, "--input", styled_path, "--revise", orig_path, "--output", output_path]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode, proc.stdout


def run_numbering(path):
    """序号段落核对（双维度算法辅助）。"""
    cmd = [sys.executable, SCRIPT, "--input", path, "--check-numbering"]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode, proc.stdout


class TestCheckStyles(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="ipo_fmt_test_")
        # 合格反馈回复文档
        cls.good = os.path.join(cls.tmp, "good.docx")
        make_mini_docx(cls.good, [
            ("0011", "问题1.关于主营业务收入确认准确性的核查"),
            ("001", "请发行人说明收入确认政策与合同条款约定的一致性。"),
            ("000", "【回复】"),
            ("002", "一、收入确认政策与合同条款一致性分析"),
            ("000", "报告期内，发行人按照企业会计准则的规定确认收入。"),
        ])
        # 不合格文档（无任何 pStyle）
        cls.bad = os.path.join(cls.tmp, "bad.docx")
        make_mini_docx(cls.bad, [(None, "裸段落一"), (None, "裸段落二")])
        # 含空段落（自闭合 + 配对）
        cls.empty = os.path.join(cls.tmp, "empty.docx")
        with zipfile.ZipFile(cls.good) as z:
            doc = z.read("word/document.xml").decode("utf-8")
        doc = doc.replace("</w:p>", "</w:p><w:p/>", 1)
        doc = doc.replace("</w:p>", "</w:p><w:p></w:p>", 1)
        with zipfile.ZipFile(cls.empty, "w", zipfile.ZIP_DEFLATED) as z2:
            for n in ["[Content_Types].xml", "_rels/.rels"]:
                with zipfile.ZipFile(cls.good) as z3:
                    z2.writestr(n, z3.read(n))
            z2.writestr("word/document.xml", doc)

    def test_good_reply_passes(self):
        code, out = run_check(self.good)
        self.assertEqual(code, 0, f"合格文档应通过，输出:\n{out}")
        self.assertIn("[PASS]", out)

    def test_bad_doc_fails(self):
        code, out = run_check(self.bad)
        self.assertNotEqual(code, 0, "不合格文档应失败")
        self.assertIn("[FAIL]", out)
        self.assertIn("裸段落", out)

    def test_empty_para_detected(self):
        code, out = run_check(self.empty)
        self.assertIn("空段落", out)
        # 自闭合 + 配对 = 2 个空段落
        m = re.search(r"(\d+) 个空段落", out)
        self.assertIsNotNone(m, f"应报空段落数量，输出:\n{out}")
        self.assertEqual(int(m.group(1)), 2, f"应为 2 个空段落，输出:\n{out}")

    def test_template_library_reply_ok(self):
        tpl = os.path.join(TEMPLATE_DIR, "反馈回复样式.docx")
        self.assertTrue(os.path.exists(tpl), "反馈回复样式.docx 应存在（assets/templates/）")
        code, out = run_check(tpl, scenario="反馈回复", mode="template")
        self.assertEqual(code, 0, f"样式库校验应通过，输出:\n{out}")

    def test_template_library_report_ok(self):
        tpl = os.path.join(TEMPLATE_DIR, "报告模板.docx")
        self.assertTrue(os.path.exists(tpl), "报告模板.docx 应存在（assets/templates/）")
        code, out = run_check(tpl, scenario="招股书", mode="template")
        self.assertEqual(code, 0, f"样式库校验应通过，输出:\n{out}")

    def test_table_paragraph_not_bare(self):
        """表格内段落不应被报为裸段落（修复回归测试）。"""
        # 构造含表格的文档：表格单元格段落 + 正文段落
        doc = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document {NS}><w:body>'
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格内文字</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            '<w:p><w:pPr><w:pStyle w:val="000"/></w:pPr><w:r><w:t>正文段落</w:t></w:r></w:p>'
            "<w:sectPr/></w:body></w:document>"
        )
        path = os.path.join(self.tmp, "table.docx")
        with zipfile.ZipFile(self.good) as z:
            ct = z.read("[Content_Types].xml")
            rels = z.read("_rels/.rels")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", rels)
            z.writestr("word/document.xml", doc)
        code, out = run_check(path, scenario="招股书")
        self.assertIn("[PASS] 无裸正文段落", out, f"表格内段落不应算裸段落，输出:\n{out}")

    def test_verify_content_same_passes(self):
        """内容完整性：相同内容 → PASS（只改格式不改内容）。"""
        code, out = run_verify(self.good, self.good)
        self.assertEqual(code, 0, f"相同内容应通过，输出:\n{out}")
        self.assertIn("[PASS]", out)

    def test_verify_content_modified_fails(self):
        """内容完整性：篡改文字 → FAIL（严禁修改原文内容）。"""
        path = os.path.join(self.tmp, "modified.docx")
        with zipfile.ZipFile(self.good) as z:
            doc = z.read("word/document.xml").decode("utf-8").replace("主营业务收入", "主营收入", 1)
            ct = z.read("[Content_Types].xml")
            rels = z.read("_rels/.rels")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", rels)
            z.writestr("word/document.xml", doc)
        code, out = run_verify(path, self.good)
        self.assertNotEqual(code, 0, "篡改内容应失败")
        self.assertIn("[FAIL]", out)

    def test_numbering_case1_and_case2(self):
        """序号段落核对（双维度算法）：情况1 标题+展开 vs 情况2 列举正文。"""
        # 情况1：短标题（一）后跟独立正文；情况2：长句 1、2、3、连续序号列举
        body_paras = [
            ("003", "（一）营业收入构成分析"),
            ("000", "发行人营业收入按产品类别划分，具体情况如下："),
            ("000", "1、精密焊接设备收入占比最高，报告期内分别为 72.00%、76.00% 和 78.50%，系发行人核心收入来源，客户结构稳定。"),
            ("000", "2、自动化装配线收入占比逐年提升，报告期内分别为 13.50%、12.80% 和 11.80%，订单金额及交付规模持续增长。"),
            ("000", "3、配套软件及运维服务收入占比相对稳定，报告期内维持在 10% 左右，毛利率较高。"),
        ]
        path = os.path.join(self.tmp, "numbering.docx")
        make_mini_docx(path, body_paras)
        code, out = run_numbering(path)
        self.assertEqual(code, 0, f"序号核对应正常，输出:\n{out}")
        self.assertIn("情况1:标题+展开", out, "（一）短标题+后段展开应为情况1")
        self.assertIn("情况2:列举正文", out, "1、2、3 长句连续序号应为情况2")

    def test_numbering_ignores_table_numbers(self):
        """序号核对：表格内数字/百分比不应被当作序号段落（修复回归）。"""
        # 表格内 78.50% 等数字不应进入序号核对
        body = (
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>78.50%</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            '<w:p><w:pPr><w:pStyle w:val="000"/></w:pPr><w:r><w:t>正文段落</w:t></w:r></w:p>'
        )
        path = os.path.join(self.tmp, "num_table.docx")
        with zipfile.ZipFile(self.good) as z:
            ct = z.read("[Content_Types].xml")
            rels = z.read("_rels/.rels")
            doc = z.read("word/document.xml").decode("utf-8")
        doc = doc.replace("<w:body>", f"<w:body>{body}", 1)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", rels)
            z.writestr("word/document.xml", doc)
        code, out = run_numbering(path)
        self.assertEqual(code, 0, f"序号核对应正常，输出:\n{out}")
        self.assertIn("未发现序号开头段落", out, "表格内百分比不应被识别为序号段落")

    def test_diff_same_style_no_changes(self):
        """格式修改 diff：内容与样式均相同 → 无修改。"""
        code, out = run_diff(self.good, self.good)
        self.assertEqual(code, 0, f"diff 应正常，输出:\n{out}")
        self.assertIn("修改段落：0 处", out, "相同文档应无修改")

    def test_diff_style_changes_generates_list(self):
        """格式修改 diff：样式变化 → 生成清单（位置/原文/改成什么）+ 统计。"""
        path = os.path.join(self.tmp, "styled_orig.docx")
        # 原文：裸段落；修正稿：带样式（内容一致）
        make_mini_docx(path, [(None, "一、收入确认政策分析"), (None, "报告期内，发行人按照准则确认收入。")])
        good = os.path.join(self.tmp, "styled_good.docx")
        make_mini_docx(good, [("002", "一、收入确认政策分析"), ("000", "报告期内，发行人按照准则确认收入。")])
        code, out = run_diff(good, path)
        self.assertEqual(code, 0, f"diff 应正常，输出:\n{out}")
        self.assertIn("修改段落：2 处", out, "两段样式变化应记录 2 处")
        self.assertIn("PASS（内容未改动）", out, "内容应一致")
        list_file = good.replace(".docx", "_格式修改清单.md")
        self.assertTrue(os.path.exists(list_file), "应生成格式修改清单 md 文件")
        with open(list_file, encoding="utf-8") as f:
            content = f.read()
        self.assertIn("## 修改统计", content)
        self.assertIn("## 修改明细", content)
        self.assertIn("002 二级标题", content, "清单应含新样式说明")
        os.remove(list_file)

    def test_revise_generates_word_revision(self):
        """生成 Word 修订稿：w:pPrChange 格式更改修订 + trackChanges。"""
        orig = os.path.join(self.tmp, "rev_orig.docx")
        styled = os.path.join(self.tmp, "rev_styled.docx")
        out = os.path.join(self.tmp, "rev_out.docx")
        make_mini_docx(orig, [(None, "一、收入确认政策分析"), (None, "报告期内，发行人按照准则确认收入。")])
        make_mini_docx(styled, [("002", "一、收入确认政策分析"), ("000", "报告期内，发行人按照准则确认收入。")])
        code, _ = run_revise(styled, orig, out)
        self.assertEqual(code, 0, "修订稿生成应成功")
        with zipfile.ZipFile(out) as z:
            doc = z.read("word/document.xml").decode("utf-8")
            settings = z.read("word/settings.xml").decode("utf-8")
        self.assertEqual(doc.count("<w:pPrChange "), 2, "两段样式变化应有 2 处格式更改修订")
        self.assertIn("<w:trackRevisions/>", settings, "应开启修订模式（正确元素名是 w:trackRevisions，"
                      "不存在 w:trackChanges——OpenXmlValidator 实证）")
        self.assertNotIn("w:trackChanges", settings, "不得含非法元素 w:trackChanges")
        self.assertIn('w:author="ibd-doc-review"', doc, "修订作者应标注")
        self.assertIn("w:formatting=\"1\"", settings,
                      "应默认显示格式更改修订（revisionView formatting=1），否则 Word 打开看不见标记")
        # CT_Settings 序列（OpenXmlValidator 实证唯一合法位置）：bordersDoNotSurroundFooter
        # → revisionView → trackRevisions → defaultTabStop
        self.assertLess(settings.index("<w:bordersDoNotSurroundFooter"),
                        settings.index("<w:revisionView"),
                        "revisionView 应在 bordersDoNotSurroundFooter 之后（CT_Settings 序列）")
        self.assertLess(settings.index("<w:revisionView"), settings.index("<w:trackRevisions/>"),
                        "revisionView 应紧邻 trackRevisions 之前")
        os.remove(out)

    def test_revise_bare_para_with_ppr_no_stray_gt(self):
        """回归：裸段落「有 pPr 但无 pStyle」生成修订时，不得残留多余 '>'（2026-08-27 修复）。
        真实文档裸段落 pPr 含 spacing/ind/rPr 直接格式，旧实现 replace('<w:pPr', ...) 漏掉
        开标签 '>' 导致 <w:pStyle/> 后多一个 '>'，XML 损坏。"""
        def make_with_ppr(path, paras):
            """paras: [(pstyle|None, text)]；None 也生成带 spacing/ind 直接格式的 pPr（模拟真实裸段落）。"""
            body = []
            for style, text in paras:
                if style:
                    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
                else:
                    ppr = ('<w:pPr><w:spacing w:line="360" w:lineRule="auto"/>'
                           '<w:ind w:firstLineChars="200" w:firstLine="480"/></w:pPr>')
                body.append(f"<w:p>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p>")
            doc = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:document {NS}><w:body>{"".join(body)}<w:sectPr/></w:body></w:document>'
            )
            with zipfile.ZipFile(self.good) as z:
                ct = z.read("[Content_Types].xml")
                rels = z.read("_rels/.rels")
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("[Content_Types].xml", ct)
                z.writestr("_rels/.rels", rels)
                z.writestr("word/document.xml", doc)

        orig = os.path.join(self.tmp, "rev_bare_orig.docx")
        styled = os.path.join(self.tmp, "rev_bare_styled.docx")
        out = os.path.join(self.tmp, "rev_bare_out.docx")
        make_with_ppr(orig, [(None, "一、收入确认政策分析"), (None, "报告期内，发行人按照准则确认收入。")])
        make_with_ppr(styled, [("002", "一、收入确认政策分析"), ("000", "报告期内，发行人按照准则确认收入。")])
        code, _ = run_revise(styled, orig, out)
        self.assertEqual(code, 0, "修订稿生成应成功")
        with zipfile.ZipFile(out) as z:
            doc = z.read("word/document.xml").decode("utf-8")
        self.assertEqual(doc.count("<w:pPrChange "), 2, "两段样式变化应有 2 处格式更改修订")
        self.assertNotIn("/>>", doc.replace("</w:pPrChange>", ""), "不得残留多余 '>'（/> 后紧跟 >）")
        self.assertIn('<w:pStyle w:val="002"/>', doc, "修订后段落应带新 pStyle")
        self.assertIn("<w:spacing w:line=\"360\"", doc, "原 pPr 直接格式应保留")
        os.remove(out)

    def test_content_check_heading_skip_and_amount(self):
        """格式核对（check_content.py）：标题序号跳号 + 金额千分位/两位小数检出。"""
        path = os.path.join(self.tmp, "content.docx")
        make_mini_docx(path, [
            ("002", "一、发行人基本情况"),
            ("000", "截至20260630，发行人总资产为 12345.6 万元。"),
            ("002", "三、股权结构"),
            ("000", "报告期内营业收入为 110,200.5万元。"),
        ])
        code, out = run_content(path, checks="heading_seq,amounts")
        self.assertEqual(code, 0, f"核对应正常完成，输出:\n{out}")
        report = read_content_report(path)
        self.assertIn("跳号", report, "「一、→ 三、」应检出跳号")
        self.assertIn("12345", report, "5 位以上数字未加千分位应检出")
        self.assertIn("110,200.5", report, "一位小数金额应检出")

    def test_content_punctuation_neighbor(self):
        """格式核对（check_content.py）：标点前后字符判定——中文语境半角报、英英间全角报。"""
        path = os.path.join(self.tmp, "punct_n.docx")
        make_mini_docx(path, [
            ("000", "XX股份有限公司(以下简称「发行人」)成立，持有ABC，DEF股权?"),
        ])
        code, out = run_content(path, checks="punctuation")
        self.assertEqual(code, 0, f"核对应正常完成，输出:\n{out}")
        report = read_content_report(path)
        self.assertIn("中文语境使用半角标点", report)
        # 2026-08-27 裁定：英英之间全角标点不再报（投行中文阅读习惯）
        self.assertNotIn("误用全角", report)

    def test_content_amount_exemptions(self):
        """格式核对豁免：比例%、文号/编码不报；仅真实金额检出。"""
        path = os.path.join(self.tmp, "exempt.docx")
        make_mini_docx(path, [
            ("000", "比例合计100%；证券代码123456；文号20260001；金额为12345.6万元。"),
        ])
        code, out = run_content(path, checks="amounts")
        self.assertEqual(code, 0, f"核对应正常完成，输出:\n{out}")
        report = read_content_report(path)
        self.assertIn("| 12345 |", report, "真实金额应检出")
        # 豁免项不应作为「问题对象」出现在明细表/明细行中
        for exempt in ("| 20260001 |", "| 123456 |", "| 100 |"):
            self.assertNotIn(exempt, report, f"{exempt} 应豁免（比例%/文号/编码）")

    def test_content_punctuation_neighbor_rule(self):
        """标点前后字符判定法：中文语境半角必报；英文/数字间半角不报；英英间全角报。"""
        bad_path = os.path.join(self.tmp, "punct_bad.docx")
        good_path = os.path.join(self.tmp, "punct_good.docx")
        make_mini_docx(bad_path, [
            ("000", "XX股份有限公司(以下简称「发行人」)成立，持有ABC，DEF股权?"),
        ])
        make_mini_docx(good_path, [
            ("000", "报告期为2023年1月1日（即12个月），收入1,234.56万元（占比5%）。"),
            ("000", "The Company held 1,200 shares as of Jun 30."),
        ])
        run_content(bad_path, checks="punctuation")
        run_content(good_path, checks="punctuation")
        rep_bad = read_content_report(bad_path)
        rep_good = read_content_report(good_path)
        self.assertIn("中文语境使用半角标点", rep_bad, "中文语境半角 ( , ? 应检出")
        # 2026-08-27 裁定：英英之间全角标点不再报（投行中文阅读习惯）
        self.assertNotIn("误用全角", rep_bad)
        for banned in ["中文语境使用半角标点", "误用全角"]:
            self.assertNotIn(banned, rep_good, f"规范写法不应报「{banned}」")


if __name__ == "__main__":
    unittest.main(verbosity=2)