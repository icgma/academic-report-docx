#!/usr/bin/env python3
"""Comprehensive v1.1.0 Robustness Test Suite for `academic-report-docx`."""

from __future__ import annotations

import json
import os
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path

from lxml import etree

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = REPO_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import audit_report_prose as prose_linter
import latex_to_omml
import md_to_academic_docx as compiler
import verify_numbers as num_gate

W_NS = compiler.W_NS
M_NS = compiler.M_NS
WP_NS = compiler.WP_NS
W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"
WP = f"{{{WP_NS}}}"


def _make_png_bytes(width: int, height: int) -> bytes:
    row = b"\x00" + (b"\x20\x40\x60" * width)
    raw = row * height

    def _chunk(tag: bytes, data: bytes) -> bytes:
        crc = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", crc)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", zlib.compress(raw)) + _chunk(b"IEND", b"")


class AcademicReportDocxV11Tests(unittest.TestCase):
    """End-to-end tests verifying all 6 v1.1.0 robustness upgrades."""

    def test_01_context_aware_prose_linter_and_profiles(self) -> None:
        # 1. Quoted interview/archival speech + blockquote should be exempted from pronoun bans
        md_valid_quotes = """---
title: 基层治理访谈与面板分析汇报
---

访谈记录显示，基层经办人员普遍反映“我觉得跨部门流转比过去快了一半，你们现在线上就能核验”，与定量结果一致。

> 档案摘录（编号 A12）：“我部核查各省上报清册，尔等须于限内补齐原档。”

![趋势分布](chart.png)

图 1 跨部门流转用时年度分布趋势
"""
        self.assertEqual(prose_linter.audit_markdown(md_valid_quotes), [])

        # 2. Unquoted 1st/2nd-person pronoun MUST fail
        md_unquoted_pronoun = "本阶段我已经完成了全部清洗，请你确认变量口径。\n"
        issues_pronoun = prose_linter.audit_markdown(md_unquoted_pronoun)
        self.assertTrue(any("[人称代词违规]" in x for x in issues_pronoun), issues_pronoun)

        # 3. Domain profile switching (`core` vs `core,historical`)
        md_hist_term = "清代官员考核档案中包含大量双轨奏疏与题本记录。\n"
        self.assertEqual(prose_linter.audit_markdown(md_hist_term, profiles=["core"]), [])
        issues_hist = prose_linter.audit_markdown(md_hist_term, profiles=["core", "historical"])
        self.assertTrue(any("奏疏" in x for x in issues_hist), issues_hist)

        # 4. Misplaced figure caption ABOVE image MUST fail (`表题在表上、图题在图下`)
        md_bad_fig = "图 1 错误放置在图片上方的图题\n\n![趋势图](chart.png)\n"
        issues_fig = prose_linter.audit_markdown(md_bad_fig)
        self.assertTrue(any("[图题位置违规]" in x for x in issues_fig), issues_fig)

    def test_02_wide_table_landscape_and_two_row_spanned_headers(self) -> None:
        md_wide = """---
title: 多模型宽表自动横向分节与双层跨列表头测试
---

前序纵向正文段落说明基准样本构造逻辑与估计框架。

表 1 基础样本分布简表

| 区域分组 | 城市数量 | 观测数量 | 样本占比 |
| :--- | :---: | :---: | :---: |
| 东部地区 | 100 | 1,200 | 40.00% |
| 中西部地区 | 150 | 1,800 | 60.00% |

注：本表为 4 列常规窄表，应保持在纵向 A4 页面内。

表 2 分区域与分行政层级的多模型稳健性回归结果

| 解释变量 | 东部城市子样本 | > | > | 中西部城市子样本 | > | > | 全样本汇总 |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
|  | (1) 基线 | (2) 加控制 | (3) 工具变量 | (4) 基线 | (5) 加控制 | (6) 工具变量 | (7) 双固定效应 |
| 数字化指数（*digital*） | -1.842*** | -1.621*** | -1.705*** | -1.214** | -1.108** | -1.190** | -1.516*** |
|  | (0.291) | (0.275) | (0.312) | (0.388) | (0.364) | (0.402) | (0.264) |
| 观测数 N | 1,200 | 1,200 | 1,200 | 1,800 | 1,800 | 1,800 | 3,000 |

注：括号内为聚类稳健标准误；本表共 8 列，应自动切换至横向 A4 分节页并在表后切回纵向页面。

后续纵向正文段落继续讨论机制检验与异质性推论，页面应自动恢复为标准纵向 A4 排版。
"""
        for force_pure in ("0", "1"):
            with tempfile.TemporaryDirectory() as td:
                tmp = Path(td)
                md_file = tmp / "wide_report.md"
                out_docx = tmp / f"wide_out_{force_pure}.docx"
                md_file.write_text(md_wide, encoding="utf-8")

                env = os.environ.copy()
                env["FORCE_PURE_PYTHON_DOCX"] = force_pure
                res = subprocess.run(
                    [sys.executable, str(SCRIPTS_DIR / "md_to_academic_docx.py"), str(md_file), "-o", str(out_docx), "--verify"],
                    env=env,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    check=False,
                )
                self.assertEqual(res.returncode, 0, msg=res.stderr + "\n" + res.stdout)

                stats = compiler.inspect_academic_docx(out_docx)
                self.assertEqual(stats["tables"], 2)
                self.assertEqual(stats["landscape_sections"], 1)

                with zipfile.ZipFile(out_docx, "r") as zf:
                    doc = etree.fromstring(zf.read("word/document.xml"))
                sects = doc.findall(f".//{W}sectPr")
                orients = [s.find(f"{W}pgSz").get(f"{W}orient", "portrait") for s in sects]
                self.assertEqual(orients, ["portrait", "landscape", "portrait"])

                wide_tbl = doc.findall(f".//{W}tbl")[1]
                rows = wide_tbl.findall(f"{W}tr")
                hdr_count = sum(1 for r in rows if r.find(f"{W}trPr/{W}tblHeader") is not None)
                self.assertEqual(hdr_count, 2)
                spans = [
                    int(gs.get(f"{W}val"))
                    for gs in rows[0].findall(f"{W}tc/{W}tcPr/{W}gridSpan")
                ]
                self.assertEqual(spans, [3, 3])

    def test_03_figure_normalization_and_caption_below(self) -> None:
        svg_sample = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 900 450" width="900" height="450">
  <rect width="900" height="450" fill="#ffffff"/>
  <line x1="80" y1="380" x2="820" y2="120" stroke="#000000" stroke-width="4"/>
</svg>"""
        for force_pure in ("0", "1"):
            with tempfile.TemporaryDirectory() as td:
                tmp = Path(td)
                big_png = tmp / "oversized_2400x1200.png"
                big_png.write_bytes(_make_png_bytes(400, 200))
                vec_svg = tmp / "vector_trend.svg"
                vec_svg.write_text(svg_sample, encoding="utf-8")

                md_fig = f"""---
title: 插图宽度自适应与图下居中图题测试
---

正文说明事件研究法动态效应估计轨迹。

![动态效应图]({big_png.name})

图 1 改革前后跨部门响应时长动态效应估计

注：竖线代表 95% 置信区间，基期设定为改革前第 1 期。

![矢量示意图]({vec_svg.name})

图 2 行政审批数据跨部门流转路径示意
"""
                md_file = tmp / "fig_report.md"
                out_docx = tmp / f"fig_out_{force_pure}.docx"
                md_file.write_text(md_fig, encoding="utf-8")

                # Force oversized EMU inside pure check as well
                env = os.environ.copy()
                env["FORCE_PURE_PYTHON_DOCX"] = force_pure
                res = subprocess.run(
                    [sys.executable, str(SCRIPTS_DIR / "md_to_academic_docx.py"), str(md_file), "-o", str(out_docx), "--verify"],
                    env=env,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    check=False,
                )
                self.assertEqual(res.returncode, 0, msg=res.stderr + "\n" + res.stdout)
                stats = compiler.inspect_academic_docx(out_docx)
                self.assertEqual(stats["figures"], 2)

                with zipfile.ZipFile(out_docx, "r") as zf:
                    doc = etree.fromstring(zf.read("word/document.xml"))
                for ext_el in doc.findall(f".//{WP}extent"):
                    self.assertLessEqual(int(ext_el.get("cx", "0")), compiler.MAX_FIG_WIDTH_EMU)

    def test_04_direct_docx_restyle_and_windows_lock_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            messy_docx = tmp / "messy_colored.docx"
            clean_docx = tmp / "clean_target.docx"

            # Synthesize a messy Word document with Calibri, blue/red text, shading, and full grid borders
            doc_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:r="{compiler.R_NS}">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="等线"/><w:color w:val="2F5496"/></w:rPr><w:t>一、原始彩色网格表测试</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:rPr><w:color w:val="FF0000"/><w:highlight w:val="yellow"/><w:u w:val="single"/></w:rPr><w:t>本段带有红色字体、黄色高亮与下划线。</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="4472C4"/>
          <w:left w:val="single" w:sz="4" w:color="4472C4"/>
          <w:bottom w:val="single" w:sz="4" w:color="4472C4"/>
          <w:right w:val="single" w:sz="4" w:color="4472C4"/>
          <w:insideH w:val="single" w:sz="4" w:color="4472C4"/>
          <w:insideV w:val="single" w:sz="4" w:color="4472C4"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tr>
        <w:tc><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr><w:p><w:r><w:t>指标</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr><w:p><w:r><w:t>均值</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>响应时长</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>4.25</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>"""
            styles_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="等线"/><w:color w:val="2F5496"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:color w:val="2F5496"/></w:rPr></w:style>
</w:styles>"""
            rels_root = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
            doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
            ct_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""
            with zipfile.ZipFile(messy_docx, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("[Content_Types].xml", ct_xml)
                zf.writestr("_rels/.rels", rels_root)
                zf.writestr("word/_rels/document.xml.rels", doc_rels)
                zf.writestr("word/styles.xml", styles_xml)
                zf.writestr("word/document.xml", doc_xml)

            res1 = subprocess.run(
                [sys.executable, str(SCRIPTS_DIR / "md_to_academic_docx.py"), str(messy_docx), "--restyle", "-o", str(clean_docx), "--verify"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                check=False,
            )
            self.assertEqual(res1.returncode, 0, msg=res1.stderr)
            self.assertTrue(clean_docx.exists())

            # Simulate Windows Word/WPS `~$` lock file on `clean_target.docx`
            lock_file = tmp / f"~${clean_docx.name}"
            lock_file.write_bytes(b"locked")

            res2 = subprocess.run(
                [sys.executable, str(SCRIPTS_DIR / "md_to_academic_docx.py"), str(messy_docx), "--restyle", "-o", str(clean_docx), "--verify"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                check=False,
            )
            self.assertEqual(res2.returncode, 0, msg=res2.stderr)
            fallback_docx = tmp / "clean_target_updated.docx"
            self.assertTrue(fallback_docx.exists(), "Expected automatic _updated.docx fallback when ~$ lock exists!")

    def test_05_empirical_number_cross_verification(self) -> None:
        sample_md = (REPO_ROOT / "examples" / "sample_empirical_report.md").read_text(encoding="utf-8")
        self.assertEqual(num_gate.verify_markdown_numbers(sample_md), [])

        # Sign flip (`+1.539` vs `-1.539` in Table 2)
        flipped_md = sample_md.replace("$-1.539^{***}$", "$+1.539^{***}$")
        flip_issues = num_gate.verify_markdown_numbers(flipped_md)
        self.assertTrue(any("[正负号不一致]" in x for x in flip_issues), flip_issues)

        # Digit transposition (`-1.593` vs `-1.539` in Table 2)
        typo_md = sample_md.replace("$-1.539^{***}$", "$-1.593^{***}$")
        typo_issues = num_gate.verify_markdown_numbers(typo_md)
        self.assertTrue(any("[表文数值失配]" in x for x in typo_issues), typo_issues)

        # Ground-truth `--data-source` JSON validation
        gt_payload = {
            "sample": {"cities": 285, "raw_obs": 3420, "clean_obs": 3378, "pass_rate": 0.9965},
            "table1": [0.942, 1.0, 0.9185, 0.9976, 0.884, 0.9988, 0.7614, 0.9642],
            "table2": {
                "digital_gov": [-1.842, 0.291, -1.516, 0.264, -1.539, 0.278, -0.084, 0.245],
                "lag_resp": [0.312, 0.041, 0.308, 0.044, 0.319, 0.042],
                "ln_fiscal": [-0.418, 0.196, -0.352, 0.174, -0.367, 0.181, -0.349, 0.176],
                "broadband": [-0.625, 0.218, -0.514, 0.195, -0.529, 0.203, -0.508, 0.197],
                "N": [3257, 2972],
                "r2": [0.6412, 0.7185, 0.7204, 0.7011],
            },
        }
        with tempfile.TemporaryDirectory() as td:
            gt_path = Path(td) / "results.json"
            gt_path.write_text(json.dumps(gt_payload), encoding="utf-8")
            self.assertEqual(num_gate.verify_markdown_numbers(sample_md, gt_path), [])

            corrupted_table_md = sample_md.replace("0.7204", "0.8888")
            gt_issues = num_gate.verify_markdown_numbers(corrupted_table_md, gt_path)
            self.assertTrue(any("[数据源对账失败]" in x for x in gt_issues), gt_issues)

    def test_06_zero_pandoc_native_omml_and_five_agent_bundle_parity(self) -> None:
        # 1. Test rich LaTeX -> OMML AST directly
        omml_para = latex_to_omml.latex_to_omml_element(
            r"\hat{\beta}_1 = \frac{\sum_{i=1}^{N} (X_i - \bar{X}) Y_i}{\sqrt{\sum_{i=1}^{N} X_i^2}}",
            display=True,
        )
        self.assertEqual(omml_para.tag, f"{M}oMathPara")
        self.assertIsNotNone(omml_para.find(f".//{M}f"))
        self.assertIsNotNone(omml_para.find(f".//{M}nary"))
        self.assertIsNotNone(omml_para.find(f".//{M}rad"))
        self.assertIsNotNone(omml_para.find(f".//{M}acc"))

        # 2. Force pure-Python compilation on sample report and verify genuine `<m:oMath>` equations
        with tempfile.TemporaryDirectory() as td:
            out_docx = Path(td) / "pure_python_sample.docx"
            env = os.environ.copy()
            env["FORCE_PURE_PYTHON_DOCX"] = "1"
            res = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPTS_DIR / "md_to_academic_docx.py"),
                    str(REPO_ROOT / "examples" / "sample_empirical_report.md"),
                    "-o",
                    str(out_docx),
                    "--verify",
                ],
                env=env,
                capture_output=True,
                text=True,
                encoding="utf-8",
                check=False,
            )
            self.assertEqual(res.returncode, 0, msg=res.stderr)
            stats = compiler.inspect_academic_docx(out_docx)
            self.assertGreaterEqual(stats["equations"], 5, "Expected native <m:oMath> equations without Pandoc!")

        # 3. Validate 5-agent skill metadata & SVG well-formedness
        skill_text = (REPO_ROOT / "SKILL.md").read_text(encoding="utf-8")
        meta, _, _ = prose_linter.parse_frontmatter(skill_text)
        self.assertEqual(meta.get("name"), "academic-report-docx")
        self.assertIn("whenToUse", meta)
        self.assertEqual(meta.get("user-invocable"), "true")
        self.assertEqual(meta.get("disable-model-invocation"), "false")
        self.assertIn("display_name_zh", meta)
        self.assertIn("display_name_en", meta)
        self.assertTrue((REPO_ROOT / "agents" / "openai.yaml").exists())

        for svg_file in (REPO_ROOT / "assets").glob("*.svg"):
            etree.parse(str(svg_file))


if __name__ == "__main__":
    unittest.main()
