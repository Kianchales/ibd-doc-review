#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_check_styles.py — ipo-doc-formatting 校验脚本自测（2026-08-25 标准流程强化）
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
    """paras: [(style|None, text)]，生成最小 docx（仅 document.xml + 必需部件）。"""
    body = []
    for style, text in paras:
        pstyle = f'<w:pStyle w:val="{style}"/>' if style else ""
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
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)


def run_check(path, scenario="反馈回复", mode="document"):
    cmd = [sys.executable, SCRIPT, "--input", path, "--scenario", scenario, "--mode", mode]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode, proc.stdout


def run_verify(new_path, orig_path):
    """内容完整性校验（严禁修改原文内容，2026-08-25 裁定）。"""
    cmd = [sys.executable, SCRIPT, "--input", new_path, "--verify-content", orig_path]
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


if __name__ == "__main__":
    unittest.main(verbosity=2)
