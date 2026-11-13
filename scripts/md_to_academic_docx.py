#!/usr/bin/env python3
"""Universal Markdown-to-Academic-DOCX Compiler & OpenXML Typesetter (`md_to_academic_docx.py`).

Compiles Markdown into publication-grade Chinese-English academic Word (`.docx`) documents:
- Chinese font: `宋体` (SimSun); English/Numbers font: `Times New Roman`; Math: `Cambria Math` (OMML)
- Color: 100% pure black (`#000000`) across all headings, body text, equations, tables, and page numbers
- Page layout: A4 (`21.0 cm x 29.7 cm`), `2.54 cm` top/bottom margins, `2.80 cm` left/right margins,
  centered `10pt` footer page numbers
- Body paragraphs: `12pt` (`小四`), `1.5x` line spacing (`w:line="360"`), `2-char` first-line indent, justified
- Tables: Standard GB/T 7713.1 Three-Line Academic Tables (`三段式三线表`: `1.5pt` top rule,
  `0.75pt` header-bottom rule, `1.5pt` bottom rule, zero vertical or inner horizontal rules, zero shading),
  paired with centered bold `10.5pt` table captions above and `9pt` table notes below.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"
R = f"{{{R_NS}}}"


def find_pandoc() -> str | None:
    """Locate Pandoc binary across PATH and standard Windows/macOS/Linux locations."""
    found = shutil.which("pandoc")
    if found:
        return found
    candidates = [
        Path(r"C:\Program Files\Pandoc\pandoc.exe"),
        Path.home() / "AppData" / "Local" / "Pandoc" / "pandoc.exe",
        Path("/opt/homebrew/bin/pandoc"),
        Path("/usr/local/bin/pandoc"),
        Path("/usr/bin/pandoc"),
    ]
    for c in candidates:
        if c.exists():
            return str(c)
    return None


def _set_rfonts_and_black(
    rpr: etree._Element,
    size_half_pt: int | None = None,
    bold: bool | None = None,
    is_math: bool = False,
) -> None:
    """Enforce Times New Roman + SimSun (`宋体`) and pure black `#000000` on `<w:rPr>`."""
    for tag in ("rStyle", "shd", "highlight", "bdr"):
        el = rpr.find(f"{W}{tag}")
        if el is not None and (tag != "rStyle" or el.get(f"{W}val") in ("48", "VerbatimChar", "Hyperlink", "20")):
            rpr.remove(el)

    rf = rpr.find(f"{W}rFonts")
    if rf is None:
        rf = etree.SubElement(rpr, f"{W}rFonts")
    for attr in list(rf.attrib.keys()):
        del rf.attrib[attr]
    if is_math:
        rf.set(f"{W}ascii", "Cambria Math")
        rf.set(f"{W}hAnsi", "Cambria Math")
        rf.set(f"{W}eastAsia", "宋体")
        rf.set(f"{W}cs", "Cambria Math")
    else:
        rf.set(f"{W}ascii", "Times New Roman")
        rf.set(f"{W}hAnsi", "Times New Roman")
        rf.set(f"{W}eastAsia", "宋体")
        rf.set(f"{W}cs", "Times New Roman")

    col = rpr.find(f"{W}color")
    if col is None:
        col = etree.SubElement(rpr, f"{W}color")
    for attr in list(col.attrib.keys()):
        del col.attrib[attr]
    col.set(f"{W}val", "000000")

    if size_half_pt is not None:
        for sz_tag in ("sz", "szCs"):
            sz_el = rpr.find(f"{W}{sz_tag}")
            if sz_el is None:
                sz_el = etree.SubElement(rpr, f"{W}{sz_tag}")
            sz_el.set(f"{W}val", str(size_half_pt))

    if bold is not None:
        for b_tag in ("b", "bCs"):
            b_el = rpr.find(f"{W}{b_tag}")
            if bold:
                if b_el is None:
                    b_el = etree.SubElement(rpr, f"{W}{b_tag}")
                b_el.attrib.pop(f"{W}val", None)
            elif b_el is not None:
                rpr.remove(b_el)


def _set_para_layout(
    ppr: etree._Element,
    jc: str = "both",
    before: int = 0,
    after: int = 0,
    line: int = 360,
    first_line_chars: int = 0,
    first_line_dxa: int = 0,
    keep_next: bool = False,
) -> None:
    """Configure paragraph alignment, line spacing, indentation, and clear borders/shading."""
    for tag in ("pBdr", "shd"):
        el = ppr.find(f"{W}{tag}")
        if el is not None:
            ppr.remove(el)

    jc_el = ppr.find(f"{W}jc")
    if jc_el is None:
        jc_el = etree.SubElement(ppr, f"{W}jc")
    jc_el.set(f"{W}val", jc)

    sp = ppr.find(f"{W}spacing")
    if sp is None:
        sp = etree.SubElement(ppr, f"{W}spacing")
    for attr in list(sp.attrib.keys()):
        del sp.attrib[attr]
    sp.set(f"{W}before", str(before))
    sp.set(f"{W}after", str(after))
    sp.set(f"{W}line", str(line))
    sp.set(f"{W}lineRule", "auto")

    ind = ppr.find(f"{W}ind")
    if ind is None:
        ind = etree.SubElement(ppr, f"{W}ind")
    for attr in list(ind.attrib.keys()):
        del ind.attrib[attr]
    if first_line_chars > 0:
        ind.set(f"{W}firstLineChars", str(first_line_chars))
        ind.set(f"{W}firstLine", str(first_line_dxa))
    else:
        ind.set(f"{W}left", "0")
        ind.set(f"{W}right", "0")
        ind.set(f"{W}firstLine", "0")

    kn = ppr.find(f"{W}keepNext")
    if keep_next:
        if kn is None:
            etree.SubElement(ppr, f"{W}keepNext")
    elif kn is not None:
        ppr.remove(kn)


def _format_three_line_table(tbl: etree._Element) -> None:
    """Format `<w:tbl>` into a textbook GB/T 7713.1 Academic Three-Line Table (`三段式三线表`)."""
    rows = tbl.findall(f"{W}tr")
    if not rows:
        return
    n_cols = max(len(r.findall(f"{W}tc")) for r in rows)
    cell_font_half_pt = 21 if n_cols <= 4 else 18  # 10.5pt (`五号`) or 9pt (`小五`) for wide tables

    tbl_pr = tbl.find(f"{W}tblPr")
    if tbl_pr is None:
        tbl_pr = etree.Element(f"{W}tblPr")
        tbl.insert(0, tbl_pr)
    for child in list(tbl_pr):
        tbl_pr.remove(child)

    tbl_w = etree.SubElement(tbl_pr, f"{W}tblW")
    tbl_w.set(f"{W}w", "5000")
    tbl_w.set(f"{W}type", "pct")

    tbl_jc = etree.SubElement(tbl_pr, f"{W}jc")
    tbl_jc.set(f"{W}val", "center")

    bdrs = etree.SubElement(tbl_pr, f"{W}tblBorders")
    for side, val, sz in (
        ("top", "single", "12"),
        ("left", "none", "0"),
        ("bottom", "single", "12"),
        ("right", "none", "0"),
        ("insideH", "none", "0"),
        ("insideV", "none", "0"),
    ):
        b = etree.SubElement(bdrs, f"{W}{side}")
        b.set(f"{W}val", val)
        if val != "none":
            b.set(f"{W}sz", sz)
            b.set(f"{W}space", "0")
            b.set(f"{W}color", "000000")

    cmar = etree.SubElement(tbl_pr, f"{W}tblCellMar")
    for side, dxa in (("top", "70"), ("left", "80"), ("bottom", "70"), ("right", "80")):
        m = etree.SubElement(cmar, f"{W}{side}")
        m.set(f"{W}w", dxa)
        m.set(f"{W}type", "dxa")

    n_rows = len(rows)
    for r_idx, tr in enumerate(rows):
        tr_pr = tr.find(f"{W}trPr")
        if tr_pr is None:
            tr_pr = etree.Element(f"{W}trPr")
            tr.insert(0, tr_pr)
        for child in list(tr_pr):
            tr_pr.remove(child)
        etree.SubElement(tr_pr, f"{W}cantSplit")
        if r_idx == 0:
            etree.SubElement(tr_pr, f"{W}tblHeader")

        cells = tr.findall(f"{W}tc")
        cell_texts = ["".join(tc.itertext()).strip() for tc in cells]
        is_panel_subhead = (
            r_idx > 0
            and len(cell_texts) > 1
            and bool(re.match(r"^[A-Z]\.\s", cell_texts[0]))
            and all(t == "" for t in cell_texts[1:])
        )

        for c_idx, tc in enumerate(cells):
            tc_pr = tc.find(f"{W}tcPr")
            if tc_pr is None:
                tc_pr = etree.Element(f"{W}tcPr")
                tc.insert(0, tc_pr)
            for tag in ("tcBorders", "shd", "vAlign"):
                old = tc_pr.find(f"{W}{tag}")
                if old is not None:
                    tc_pr.remove(old)

            tc_bdrs = etree.SubElement(tc_pr, f"{W}tcBorders")
            top_spec = ("single", "12") if r_idx == 0 else ("none", "0")
            if r_idx == 0:
                bot_spec = ("single", "6")
            elif r_idx == n_rows - 1:
                bot_spec = ("single", "12")
            else:
                bot_spec = ("none", "0")

            for side, (b_val, b_sz) in (
                ("top", top_spec),
                ("left", ("none", "0")),
                ("bottom", bot_spec),
                ("right", ("none", "0")),
            ):
                b = etree.SubElement(tc_bdrs, f"{W}{side}")
                b.set(f"{W}val", b_val)
                if b_val != "none":
                    b.set(f"{W}sz", b_sz)
                    b.set(f"{W}space", "0")
                    b.set(f"{W}color", "000000")

            valign = etree.SubElement(tc_pr, f"{W}vAlign")
            valign.set(f"{W}val", "center")

            if r_idx == 0:
                col_jc = "left" if c_idx == 0 else "center"
            elif c_idx == 0 or (n_cols == 4 and c_idx == 3):
                col_jc = "left"
            else:
                col_jc = "center"

            for p in tc.findall(f"{W}p"):
                ppr = p.find(f"{W}pPr")
                if ppr is None:
                    ppr = etree.Element(f"{W}pPr")
                    p.insert(0, ppr)
                _set_para_layout(ppr, jc=col_jc, before=20, after=20, line=260, first_line_chars=0)
                for r in p.findall(f"{W}r"):
                    rpr = r.find(f"{W}rPr")
                    if rpr is None:
                        rpr = etree.Element(f"{W}rPr")
                        r.insert(0, rpr)
                    force_bold = True if (r_idx == 0 or is_panel_subhead) else None
                    _set_rfonts_and_black(rpr, size_half_pt=cell_font_half_pt, bold=force_bold)
                for mr in p.findall(f".//{M}r"):
                    rpr = mr.find(f"{W}rPr")
                    if rpr is None:
                        rpr = etree.Element(f"{W}rPr")
                        mr.insert(0, rpr)
                    _set_rfonts_and_black(rpr, size_half_pt=cell_font_half_pt, bold=None, is_math=True)


def apply_academic_word_formatting(docx_path: Path) -> None:
    """Post-process OpenXML inside `docx_path` to enforce SimSun + Times New Roman, black, & 3-line tables."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(docx_path, "r") as zin:
            zin.extractall(tmp)

        styles_path = tmp / "word" / "styles.xml"
        styles_tree = etree.parse(str(styles_path))
        styles_root = styles_tree.getroot()

        sid_to_name: dict[str, str] = {}
        for st in styles_root.findall(f"{W}style"):
            sid = st.get(f"{W}styleId", "")
            nm_el = st.find(f"{W}name")
            nm = (nm_el.get(f"{W}val") or "").lower() if nm_el is not None else ""
            if sid and nm:
                sid_to_name[sid] = nm

        for rpr_def in styles_root.findall(f".//{W}rPrDefault/{W}rPr"):
            _set_rfonts_and_black(rpr_def, size_half_pt=24, bold=False)

        for st in styles_root.findall(f"{W}style"):
            nm = sid_to_name.get(st.get(f"{W}styleId", ""), "")
            rpr = st.find(f"{W}rPr")
            if rpr is None:
                rpr = etree.SubElement(st, f"{W}rPr")
            if nm == "title":
                _set_rfonts_and_black(rpr, size_half_pt=32, bold=True)
            elif nm == "subtitle":
                _set_rfonts_and_black(rpr, size_half_pt=22, bold=False)
            elif nm == "heading 1":
                _set_rfonts_and_black(rpr, size_half_pt=28, bold=True)
            elif nm in ("heading 2", "heading 3"):
                _set_rfonts_and_black(rpr, size_half_pt=24, bold=True)
            else:
                _set_rfonts_and_black(rpr, size_half_pt=24, bold=None)

            ppr = st.find(f"{W}pPr")
            if ppr is not None:
                for tag in ("pBdr", "shd"):
                    el = ppr.find(f"{W}{tag}")
                    if el is not None:
                        ppr.remove(el)

        styles_tree.write(str(styles_path), xml_declaration=True, encoding="UTF-8")

        doc_path = tmp / "word" / "document.xml"
        doc_tree = etree.parse(str(doc_path))
        doc_root = doc_tree.getroot()
        body = doc_root.find(f"{W}body")
        assert body is not None

        children = list(body)
        for idx, child in enumerate(children):
            if child.tag == f"{W}tbl":
                _format_three_line_table(child)
                continue
            if child.tag != f"{W}p":
                continue

            p_text = "".join(child.itertext()).strip()
            ppr = child.find(f"{W}pPr")
            if ppr is None:
                ppr = etree.Element(f"{W}pPr")
                child.insert(0, ppr)
            ps = ppr.find(f"{W}pStyle")
            sid = ps.get(f"{W}val") if ps is not None else ""
            sname = sid_to_name.get(sid, "")

            has_display_math = child.find(f"{M}oMathPara") is not None
            next_is_tbl = idx + 1 < len(children) and children[idx + 1].tag == f"{W}tbl"
            prev_is_tbl = idx > 0 and children[idx - 1].tag == f"{W}tbl"
            is_tbl_caption = next_is_tbl and bool(re.match(r"^表\s*\d+", p_text))
            is_tbl_note = prev_is_tbl and p_text.startswith("注：")

            if sname == "title":
                _set_para_layout(ppr, jc="center", before=80, after=120, line=300, first_line_chars=0, keep_next=True)
                target_sz, target_bold = 32, True
            elif sname == "subtitle":
                _set_para_layout(ppr, jc="center", before=0, after=240, line=300, first_line_chars=0, keep_next=True)
                target_sz, target_bold = 22, False
            elif sname == "heading 1":
                _set_para_layout(ppr, jc="left", before=240, after=120, line=360, first_line_chars=0, keep_next=True)
                target_sz, target_bold = 28, True
            elif sname in ("heading 2", "heading 3"):
                _set_para_layout(ppr, jc="left", before=160, after=80, line=360, first_line_chars=0, keep_next=True)
                target_sz, target_bold = 24, True
            elif is_tbl_caption:
                _set_para_layout(ppr, jc="center", before=180, after=60, line=300, first_line_chars=0, keep_next=True)
                target_sz, target_bold = 21, True
            elif is_tbl_note:
                _set_para_layout(ppr, jc="both", before=60, after=180, line=280, first_line_chars=0, keep_next=False)
                target_sz, target_bold = 18, False
            elif has_display_math:
                _set_para_layout(ppr, jc="center", before=80, after=80, line=300, first_line_chars=0, keep_next=False)
                target_sz, target_bold = 24, None
            else:
                num_pr = ppr.find(f"{W}numPr")
                if num_pr is not None:
                    ppr.remove(num_pr)
                _set_para_layout(
                    ppr,
                    jc="both",
                    before=0,
                    after=0,
                    line=360,
                    first_line_chars=200,
                    first_line_dxa=480,
                    keep_next=False,
                )
                target_sz, target_bold = 24, None

            for r in child.findall(f".//{W}r"):
                if any(anc.tag == f"{M}oMath" for anc in r.iterancestors()):
                    continue
                rpr = r.find(f"{W}rPr")
                if rpr is None:
                    rpr = etree.Element(f"{W}rPr")
                    r.insert(0, rpr)
                _set_rfonts_and_black(rpr, size_half_pt=target_sz, bold=target_bold)

            for mr in child.findall(f".//{M}r"):
                rpr = mr.find(f"{W}rPr")
                if rpr is None:
                    rpr = etree.Element(f"{W}rPr")
                    mr.insert(0, rpr)
                _set_rfonts_and_black(rpr, size_half_pt=target_sz, bold=None, is_math=True)

        footer_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W_NS}">
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>
        <w:color w:val="000000"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>
        <w:color w:val="000000"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> PAGE </w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>
        <w:color w:val="000000"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>
        <w:color w:val="000000"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>
        <w:color w:val="000000"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:ftr>
"""
        (tmp / "word" / "footer1.xml").write_text(footer_xml, encoding="utf-8")

        rels_path = tmp / "word" / "_rels" / "document.xml.rels"
        rels_tree = etree.parse(str(rels_path))
        rels_root = rels_tree.getroot()
        existing_ids = {el.get("Id", "") for el in rels_root}
        ftr_rid = "rIdAcademicFooter1"
        if ftr_rid not in existing_ids:
            rel_ns = rels_root.tag.split("}")[0] + "}"
            rel_el = etree.SubElement(rels_root, f"{rel_ns}Relationship")
            rel_el.set("Id", ftr_rid)
            rel_el.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
            rel_el.set("Target", "footer1.xml")
        rels_tree.write(str(rels_path), xml_declaration=True, encoding="UTF-8")

        ct_path = tmp / "[Content_Types].xml"
        ct_tree = etree.parse(str(ct_path))
        ct_root = ct_tree.getroot()
        ct_ns = ct_root.tag.split("}")[0] + "}"
        if not any(el.get("PartName") == "/word/footer1.xml" for el in ct_root):
            ov = etree.SubElement(ct_root, f"{ct_ns}Override")
            ov.set("PartName", "/word/footer1.xml")
            ov.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
        ct_tree.write(str(ct_path), xml_declaration=True, encoding="UTF-8")

        sect_pr = body.find(f"{W}sectPr")
        if sect_pr is None:
            sect_pr = etree.SubElement(body, f"{W}sectPr")
        for tag in ("pgSz", "pgMar", "docGrid", "footerReference"):
            for old in sect_pr.findall(f"{W}{tag}"):
                sect_pr.remove(old)

        ftr_ref = etree.Element(f"{W}footerReference")
        ftr_ref.set(f"{W}type", "default")
        ftr_ref.set(f"{R}id", ftr_rid)
        sect_pr.insert(0, ftr_ref)

        pg_sz = etree.SubElement(sect_pr, f"{W}pgSz")
        pg_sz.set(f"{W}w", "11906")
        pg_sz.set(f"{W}h", "16838")

        pg_mar = etree.SubElement(sect_pr, f"{W}pgMar")
        pg_mar.set(f"{W}top", "1440")
        pg_mar.set(f"{W}bottom", "1440")
        pg_mar.set(f"{W}left", "1588")
        pg_mar.set(f"{W}right", "1588")
        pg_mar.set(f"{W}header", "851")
        pg_mar.set(f"{W}footer", "992")
        pg_mar.set(f"{W}gutter", "0")

        doc_tree.write(str(doc_path), xml_declaration=True, encoding="UTF-8")

        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for fpath in sorted(tmp.rglob("*")):
                if fpath.is_file():
                    zout.write(fpath, fpath.relative_to(tmp).as_posix())


def _build_pure_python_base_docx(md_path: Path, docx_path: Path) -> None:
    """Fallback pure-Python Markdown -> OpenXML `.docx` builder when Pandoc is unavailable."""
    raw = md_path.read_text(encoding="utf-8")
    lines = raw.splitlines()
    title, subtitle = "", ""
    idx = 0
    if lines and lines[0].strip() == "---":
        idx = 1
        while idx < len(lines) and lines[idx].strip() != "---":
            if lines[idx].startswith("title:"):
                title = lines[idx].split(":", 1)[1].strip().strip("\"'")
            elif lines[idx].startswith("subtitle:"):
                subtitle = lines[idx].split(":", 1)[1].strip().strip("\"'")
            idx += 1
        idx += 1

    nsmap = {"w": W_NS, "m": M_NS, "r": R_NS}
    doc_el = etree.Element(f"{W}document", nsmap=nsmap)
    body = etree.SubElement(doc_el, f"{W}body")

    def add_para(text: str, style_id: str | None = None) -> None:
        p = etree.SubElement(body, f"{W}p")
        if style_id:
            ppr = etree.SubElement(p, f"{W}pPr")
            ps = etree.SubElement(ppr, f"{W}pStyle")
            ps.set(f"{W}val", style_id)
        parts = re.split(r"(\*[^\*]+\*)", text)
        for part in parts:
            if not part:
                continue
            r = etree.SubElement(p, f"{W}r")
            if part.startswith("*") and part.endswith("*") and len(part) > 2:
                rpr = etree.SubElement(r, f"{W}rPr")
                etree.SubElement(rpr, f"{W}i")
                part = part[1:-1]
            t = etree.SubElement(r, f"{W}t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = part

    if title:
        add_para(title, "Title")
    if subtitle:
        add_para(subtitle, "Subtitle")

    while idx < len(lines):
        s = lines[idx].strip()
        if not s:
            idx += 1
            continue
        if s.startswith("# "):
            add_para(s[2:].strip(), "Heading1")
            idx += 1
        elif s.startswith("## "):
            add_para(s[3:].strip(), "Heading2")
            idx += 1
        elif s.startswith("### "):
            add_para(s[4:].strip(), "Heading3")
            idx += 1
        elif s.startswith("|"):
            tbl = etree.SubElement(body, f"{W}tbl")
            while idx < len(lines) and lines[idx].strip().startswith("|"):
                row_line = lines[idx].strip()
                cells = [c.strip() for c in row_line.strip("|").split("|")]
                if not all(re.match(r"^:?-{2,}:?$", c) for c in cells):
                    tr = etree.SubElement(tbl, f"{W}tr")
                    for cell_txt in cells:
                        tc = etree.SubElement(tr, f"{W}tc")
                        p = etree.SubElement(tc, f"{W}p")
                        r = etree.SubElement(p, f"{W}r")
                        t = etree.SubElement(r, f"{W}t")
                        t.text = cell_txt.replace("*", "")
                idx += 1
        else:
            add_para(s)
            idx += 1

    etree.SubElement(body, f"{W}sectPr")
    styles_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr/></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:rPr/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:rPr/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:rPr/></w:style>
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
    docx_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct_xml)
        zf.writestr("_rels/.rels", rels_root)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/document.xml", etree.tostring(doc_el, xml_declaration=True, encoding="UTF-8"))


def verify_academic_docx(docx_path: Path) -> tuple[int, int]:
    """Run strict OpenXML assertions verifying fonts, pure black color, and 3-line tables."""
    with zipfile.ZipFile(docx_path, "r") as z:
        doc = etree.fromstring(z.read("word/document.xml"))
        styles = etree.fromstring(z.read("word/styles.xml"))
        ftr = etree.fromstring(z.read("word/footer1.xml"))

    colors = {el.get(f"{W}val") for root in (doc, styles, ftr) for el in root.findall(f".//{W}color")}
    if colors != {"000000"}:
        raise AssertionError(f"OpenXML color violation: expected {{'000000'}}, found {colors}")

    ascii_fonts: set[str | None] = set()
    ea_fonts: set[str | None] = set()
    for root in (doc, styles, ftr):
        for rf in root.findall(f".//{W}rFonts"):
            if any(anc.tag == f"{M}oMath" for anc in rf.iterancestors()):
                continue
            ascii_fonts.add(rf.get(f"{W}ascii"))
            ea_fonts.add(rf.get(f"{W}eastAsia"))
    if ascii_fonts != {"Times New Roman"}:
        raise AssertionError(f"OpenXML ASCII font violation: expected {{'Times New Roman'}}, found {ascii_fonts}")
    if ea_fonts != {"宋体"}:
        raise AssertionError(f"OpenXML EastAsia font violation: expected {{'宋体'}}, found {ea_fonts}")

    body = doc.find(f"{W}body")
    assert body is not None
    children = list(body)
    tbls = [c for c in children if c.tag == f"{W}tbl"]

    for idx, c in enumerate(children):
        if c.tag != f"{W}tbl":
            continue
        cap_txt = "".join(children[idx - 1].itertext()).strip() if idx > 0 else ""
        note_txt = "".join(children[idx + 1].itertext()).strip() if idx + 1 < len(children) else ""
        if not re.match(r"^表\s*\d+", cap_txt):
            raise AssertionError(f"Table missing centered '表 X' caption immediately above (found: '{cap_txt[:30]}')")
        if not note_txt.startswith("注："):
            raise AssertionError(f"Table missing '注：' note immediately below (found: '{note_txt[:30]}')")

        bdrs = c.find(f"{W}tblPr/{W}tblBorders")
        assert bdrs is not None
        assert bdrs.find(f"{W}top").get(f"{W}sz") == "12"
        assert bdrs.find(f"{W}bottom").get(f"{W}sz") == "12"
        for side in ("left", "right", "insideH", "insideV"):
            assert bdrs.find(f"{W}{side}").get(f"{W}val") == "none"

        rows = c.findall(f"{W}tr")
        r0_bot = rows[0].find(f"{W}tc/{W}tcPr/{W}tcBorders/{W}bottom")
        assert r0_bot is not None and r0_bot.get(f"{W}val") == "single" and r0_bot.get(f"{W}sz") == "6"
        if len(rows) > 2:
            r1_bot = rows[1].find(f"{W}tc/{W}tcPr/{W}tcBorders/{W}bottom")
            assert r1_bot is not None and r1_bot.get(f"{W}val") == "none"

    full_text = "".join(doc.itertext())
    return len(tbls), len(full_text)


def main() -> int:
    parser = argparse.ArgumentParser(description="Compile Markdown into strictly formatted Academic Word (.docx).")
    parser.add_argument("input_md", type=Path, help="Input Markdown file path.")
    parser.add_argument("-o", "--output", type=Path, default=None, help="Output .docx file path.")
    parser.add_argument("--skip-prose-audit", action="store_true", help="Skip prose style audit before compiling.")
    parser.add_argument("--verify", action="store_true", help="Run OpenXML verification assertions on output .docx.")
    args = parser.parse_args()

    if not args.input_md.exists():
        print(f"ERROR: Input file not found: {args.input_md}", file=sys.stderr)
        return 2

    if not args.skip_prose_audit:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        from audit_report_prose import audit_markdown

        issues = audit_markdown(args.input_md.read_text(encoding="utf-8"))
        if issues:
            print(f"Prose audit failed with {len(issues)} issue(s) (pass --skip-prose-audit to bypass):", file=sys.stderr)
            for item in issues:
                print(f"  - {item}", file=sys.stderr)
            return 1

    out_docx = args.output or args.input_md.with_suffix(".docx")
    out_docx.parent.mkdir(parents=True, exist_ok=True)

    pandoc_bin = None if os.environ.get("FORCE_PURE_PYTHON_DOCX") == "1" else find_pandoc()
    if pandoc_bin:
        subprocess.run(
            [pandoc_bin, str(args.input_md), "-f", "markdown+tex_math_dollars", "-o", str(out_docx)],
            check=True,
        )
        engine = f"Pandoc ({Path(pandoc_bin).name}) + OpenXML Post-Processor"
    else:
        _build_pure_python_base_docx(args.input_md, out_docx)
        engine = "Pure-Python OpenXML Builder + Post-Processor"

    apply_academic_word_formatting(out_docx)
    print(f"[OK] Compiled Academic Word DOCX via {engine}: {out_docx}")

    if args.verify:
        n_tbls, n_chars = verify_academic_docx(out_docx)
        print(
            f"[PASS] OpenXML Verification Gate passed: {n_tbls} Three-Line Table(s), "
            f"{n_chars:,} chars, 100% 宋体 + Times New Roman, pure black #000000."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
