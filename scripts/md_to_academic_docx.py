#!/usr/bin/env python3
"""Universal Markdown/DOCX-to-Academic-DOCX Compiler & OpenXML Typesetter (`md_to_academic_docx.py`, v1.1.0).

Compiles Markdown (or restyles existing messy `.docx` files via `--restyle`) into publication-grade
Chinese-English academic Word (`.docx`) documents:
- Fonts: Chinese `宋体` (SimSun); English/Numbers `Times New Roman`; Math `Cambria Math` (native OMML)
- Color: 100% pure black (`#000000`) across all headings, body text, hyperlinks, equations, tables, and footers
- Page layout: A4 Portrait (`21.0 cm x 29.7 cm`), auto-switching wide tables (`>= 7` cols or `<!-- landscape -->`)
  onto dedicated A4 Landscape (`横向分节页`) sections while keeping narrative pages Portrait
- Body paragraphs: `12pt` (`小四`), `1.5x` line spacing (`w:line="360"`), `2-char` first-line indent, justified
- Tables: Standard GB/T 7713.1 Three-Line Tables (`三段式三线表`: `1.5pt` top/bottom rules, `0.75pt` header rule,
  2-row `<w:gridSpan>` group headers with `0.5pt` `cmidrule`, stacked SE tightening, `表 X` above + `注：` below)
- Figures: Width-clamped (`<= 14.5 cm`), centered with zero first-line indent, `keepNext` locked to bold `10.5pt`
  `图 X` caption below (`表题在表上、图题在图下`), with automatic SVG-to-PNG fallback & zero-Pandoc image embedding
- Safeguards: Empirical number cross-verification (`verify_numbers.py`) + Windows Word/WPS `~$` lock fallback.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
import zlib
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"
R = f"{{{R_NS}}}"
WP = f"{{{WP_NS}}}"
A = f"{{{A_NS}}}"
PIC = f"{{{PIC_NS}}}"

LANDSCAPE_MARKER_TOKEN = "ACADEMIC_LANDSCAPE_SECTION_MARKER_V11"
MAX_FIG_WIDTH_EMU = 5_220_000   # 14.5 cm (fits cleanly inside 15.4 cm A4 Portrait text width)
MAX_FIG_HEIGHT_EMU = 6_480_000  # 18.0 cm
EMU_PER_PX_96DPI = 9525


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


def safe_resolve_output_path(desired_path: Path) -> tuple[Path, bool]:
    """Detect Windows Word/WPS `~$` lock files or exclusive OS write locks and fall back cleanly."""
    desired_path.parent.mkdir(parents=True, exist_ok=True)

    def _is_locked(target: Path) -> bool:
        lock_candidates = [target.parent / f"~${target.name}"]
        if len(target.name) > 2:
            lock_candidates.append(target.parent / f"~${target.name[2:]}")
        if any(lc.exists() for lc in lock_candidates):
            return True
        if target.exists():
            try:
                with target.open("r+b"):
                    pass
            except PermissionError:
                return True
        return False

    if not _is_locked(desired_path):
        return desired_path, False

    for suffix_idx in range(1, 50):
        tag = "_updated" if suffix_idx == 1 else f"_updated_{suffix_idx}"
        candidate = desired_path.with_name(f"{desired_path.stem}{tag}{desired_path.suffix}")
        if not _is_locked(candidate):
            return candidate, True

    raise PermissionError(f"Output path {desired_path} and fallback paths are locked by another process.")


def _read_image_dimensions_px(img_path: Path) -> tuple[int, int]:
    """Read `(width_px, height_px)` from PNG, JPEG, GIF, or SVG using Python stdlib only."""
    data = img_path.read_bytes()
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        w_px, h_px = struct.unpack(">II", data[16:24])
        return max(1, w_px), max(1, h_px)

    if data.startswith(b"GIF8") and len(data) >= 10:
        w_px, h_px = struct.unpack("<HH", data[6:10])
        return max(1, w_px), max(1, h_px)

    if data.startswith(b"\xff\xd8"):
        pos = 2
        while pos + 9 < len(data):
            if data[pos] != 0xFF:
                pos += 1
                continue
            marker = data[pos + 1]
            if marker == 0xD9 or marker == 0xDA:
                break
            seg_len = struct.unpack(">H", data[pos + 2 : pos + 4])[0]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                h_px, w_px = struct.unpack(">HH", data[pos + 5 : pos + 9])
                return max(1, w_px), max(1, h_px)
            pos += 2 + max(2, seg_len)

    if img_path.suffix.lower() == ".svg":
        try:
            svg_root = etree.fromstring(data)
            vb = svg_root.get("viewBox", "")
            if vb:
                parts = [float(x) for x in re.split(r"[\s,]+", vb.strip()) if x]
                if len(parts) == 4 and parts[2] > 0 and parts[3] > 0:
                    return round(parts[2]), round(parts[3])
            w_raw = re.sub(r"[^0-9.]", "", svg_root.get("width", "800")) or "800"
            h_raw = re.sub(r"[^0-9.]", "", svg_root.get("height", "450")) or "450"
            return max(1, int(float(w_raw))), max(1, int(float(h_raw)))
        except (etree.XMLSyntaxError, ValueError):
            return 800, 450

    return 800, 450


def _create_minimal_white_png(width_px: int, height_px: int) -> bytes:
    """Create a valid minimal RGB PNG canvas in pure stdlib (`zlib` + `struct`) as SVG fallback."""
    w_c = max(16, min(width_px, 600))
    h_c = max(16, min(height_px, 400))
    row = b"\x00" + (b"\xf8\xf9\xfa" * w_c)
    raw_scanlines = row * h_c

    def _chunk(ctype: bytes, payload: bytes) -> bytes:
        crc = zlib.crc32(ctype + payload) & 0xFFFFFFFF
        return struct.pack(">I", len(payload)) + ctype + payload + struct.pack(">I", crc)

    ihdr = struct.pack(">IIBBBBB", w_c, h_c, 8, 2, 0, 0, 0)
    idat = zlib.compress(raw_scanlines, level=6)
    return b"\x89PNG\r\n\x1a\n" + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", idat) + _chunk(b"IEND", b"")


def _try_rasterize_svg_to_png(svg_path: Path, out_png: Path) -> bool:
    """Attempt high-DPI SVG -> PNG conversion via installed Python or CLI rasterizers."""
    try:
        import cairosvg  # type: ignore[import-not-found]

        cairosvg.svg2png(url=str(svg_path), write_to=str(out_png), scale=2.0)
        if out_png.exists():
            return True
    except (ImportError, OSError, RuntimeError, ValueError):
        svg_rasterized = False
    else:
        svg_rasterized = out_png.exists()
    if svg_rasterized:
        return True

    for cmd_tpl in (
        ["resvg", "--zoom", "2", str(svg_path), str(out_png)],
        ["rsvg-convert", "-z", "2", "-o", str(out_png), str(svg_path)],
        ["inkscape", str(svg_path), f"--export-filename={out_png}", "--export-dpi=300"],
    ):
        if shutil.which(cmd_tpl[0]):
            res = subprocess.run(cmd_tpl, capture_output=True, check=False)
            if res.returncode == 0 and out_png.exists():
                return True
    return False


def _clamp_emu_size(cx: int, cy: int, max_w: int = MAX_FIG_WIDTH_EMU, max_h: int = MAX_FIG_HEIGHT_EMU) -> tuple[int, int]:
    if cx <= 0 or cy <= 0:
        return max_w, int(max_w * 0.5625)
    scale = min(1.0, max_w / float(cx), max_h / float(cy))
    return round(cx * scale), round(cy * scale)


def _preprocess_markdown_for_compile(md_path: Path, work_dir: Path) -> Path:
    """Convert `<!-- landscape -->` comments to structural markers and rasterize `.svg` figures if possible."""
    raw = md_path.read_text(encoding="utf-8")
    lines = raw.splitlines()
    out_lines: list[str] = []
    svg_counter = 0

    img_pattern = re.compile(r"^(!\[[^\]]*\]\()([^)]+)(\)\s*)$")

    for line in lines:
        s = line.strip()
        if re.match(r"^<!--\s*landscape\s*-->$", s, re.IGNORECASE):
            out_lines.append("")
            out_lines.append(LANDSCAPE_MARKER_TOKEN)
            out_lines.append("")
            continue

        m_img = img_pattern.match(s)
        if m_img:
            prefix, raw_target, suffix = m_img.groups()
            target_clean = raw_target.strip().strip("\"'")
            resolved = (md_path.parent / target_clean).resolve() if not Path(target_clean).is_absolute() else Path(target_clean)
            if resolved.exists():
                if resolved.suffix.lower() == ".svg":
                    svg_counter += 1
                    png_out = work_dir / f"converted_svg_{svg_counter}.png"
                    if _try_rasterize_svg_to_png(resolved, png_out):
                        out_lines.append(f"{prefix}{png_out.as_posix()}{suffix}")
                        continue
                out_lines.append(f"{prefix}{resolved.as_posix()}{suffix}")
                continue

        out_lines.append(line)

    staged_md = work_dir / "staged_input.md"
    staged_md.write_text("\n".join(out_lines) + "\n", encoding="utf-8")
    return staged_md


def _set_rfonts_and_black(
    rpr: etree._Element,
    size_half_pt: int | None = None,
    bold: bool | None = None,
    is_math: bool = False,
    strip_underline: bool = False,
) -> None:
    """Enforce Times New Roman + SimSun (`宋体`) and pure black `#000000` on `<w:rPr>`."""
    for tag in ("rStyle", "shd", "highlight", "bdr"):
        el = rpr.find(f"{W}{tag}")
        if el is not None and (tag != "rStyle" or el.get(f"{W}val") in ("48", "VerbatimChar", "Hyperlink", "20")):
            rpr.remove(el)
    if strip_underline:
        u_el = rpr.find(f"{W}u")
        if u_el is not None:
            rpr.remove(u_el)

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


def _logical_col_count(tr: etree._Element) -> int:
    total = 0
    for tc in tr.findall(f"{W}tc"):
        gs = tc.find(f"{W}tcPr/{W}gridSpan")
        total += int(gs.get(f"{W}val", "1")) if gs is not None else 1
    return total


def _detect_and_merge_two_row_header(rows: list[etree._Element]) -> bool:
    """Detect 2-row grouped headers (`>` continuation markers or repeated group labels) and apply `<w:gridSpan>`."""
    if len(rows) < 3:
        return False

    # Already merged `<w:gridSpan>` in Row 0 from an existing `.docx` in `--restyle` mode
    r0_cells = rows[0].findall(f"{W}tc")
    r1_cells = rows[1].findall(f"{W}tc")
    if any(tc.find(f"{W}tcPr/{W}gridSpan") is not None for tc in r0_cells):
        r1_t0 = "".join(r1_cells[0].itertext()).strip() if r1_cells else ""
        if r1_t0 == "" or rows[1].find(f"{W}trPr/{W}tblHeader") is not None:
            return True

    if len(r0_cells) != len(r1_cells) or len(r0_cells) < 3:
        return False

    r0_texts = ["".join(tc.itertext()).strip() for tc in r0_cells]
    r1_texts = ["".join(tc.itertext()).strip() for tc in r1_cells]

    has_marker = any(t in (">", "^") for t in r0_texts[1:])
    has_repeated_group = r1_texts[0] == "" and any(
        r0_texts[j] != "" and r0_texts[j] == r0_texts[j - 1] for j in range(2, len(r0_texts))
    )
    if not (has_marker or has_repeated_group):
        return False

    # Merge contiguous group spans across Row 0 (`>`/`^` or identical adjacent group label)
    leader_tc: etree._Element | None = None
    leader_text = ""
    span = 1

    def _commit_span(tc_el: etree._Element | None, count: int) -> None:
        if tc_el is None or count <= 1:
            return
        tc_pr = tc_el.find(f"{W}tcPr")
        if tc_pr is None:
            tc_pr = etree.Element(f"{W}tcPr")
            tc_el.insert(0, tc_pr)
        gs = tc_pr.find(f"{W}gridSpan")
        if gs is None:
            gs = etree.SubElement(tc_pr, f"{W}gridSpan")
        gs.set(f"{W}val", str(count))

    for c_idx, tc in enumerate(r0_cells):
        txt = r0_texts[c_idx]
        if c_idx == 0:
            leader_tc = tc
            leader_text = txt
            span = 1
            continue
        is_cont = txt in (">", "^") or (has_repeated_group and txt != "" and txt == leader_text)
        if is_cont and leader_tc is not None:
            span += 1
            rows[0].remove(tc)
        else:
            _commit_span(leader_tc, span)
            leader_tc = tc
            leader_text = txt
            span = 1
    _commit_span(leader_tc, span)
    return True


def _is_stacked_se_row(cells: list[etree._Element]) -> bool:
    """Identify regression standard-error rows (`cell[0] == ''` and `(0.291)` style cells)."""
    if len(cells) < 2:
        return False
    texts = ["".join(tc.itertext()).strip() for tc in cells]
    if texts[0] != "":
        return False
    non_empty = [t for t in texts[1:] if t not in ("", "—", "-", "--")]
    return bool(non_empty) and all(bool(re.match(r"^[（(\[]\s*-?\d+\.\d+\s*[)）\]]$", t)) for t in non_empty)


def _format_three_line_table(tbl: etree._Element) -> int:
    """Format `<w:tbl>` into a GB/T 7713.1 Academic Three-Line Table (`三段式三线表`) and return `n_cols`."""
    rows = tbl.findall(f"{W}tr")
    if not rows:
        return 0

    has_two_row_header = _detect_and_merge_two_row_header(rows)
    header_rows = 2 if has_two_row_header else 1
    n_cols = max(_logical_col_count(r) for r in rows)
    cell_font_half_pt = 21 if n_cols <= 4 else 18  # 10.5pt (`五号`) or 9pt (`小五`) for >=5 cols

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
    h_pad = "60" if n_cols >= 7 else "80"
    for side, dxa in (("top", "65"), ("left", h_pad), ("bottom", "65"), ("right", h_pad)):
        m = etree.SubElement(cmar, f"{W}{side}")
        m.set(f"{W}w", dxa)
        m.set(f"{W}type", "dxa")

    n_rows = len(rows)
    se_flags = [_is_stacked_se_row(r.findall(f"{W}tc")) for r in rows]

    for r_idx, tr in enumerate(rows):
        tr_pr = tr.find(f"{W}trPr")
        if tr_pr is None:
            tr_pr = etree.Element(f"{W}trPr")
            tr.insert(0, tr_pr)
        for child in list(tr_pr):
            tr_pr.remove(child)
        etree.SubElement(tr_pr, f"{W}cantSplit")
        if r_idx < header_rows:
            etree.SubElement(tr_pr, f"{W}tblHeader")

        cells = tr.findall(f"{W}tc")
        cell_texts = ["".join(tc.itertext()).strip() for tc in cells]
        is_panel_subhead = (
            r_idx >= header_rows
            and len(cell_texts) > 1
            and bool(re.match(r"^(Panel\s+[A-Z]|[A-Z]\.)\s", cell_texts[0]))
            and all(t == "" for t in cell_texts[1:])
        )
        is_se_row = r_idx >= header_rows and se_flags[r_idx]
        next_is_se_row = (r_idx + 1 < n_rows) and (r_idx >= header_rows) and se_flags[r_idx + 1]

        for c_idx, tc in enumerate(cells):
            tc_pr = tc.find(f"{W}tcPr")
            if tc_pr is None:
                tc_pr = etree.Element(f"{W}tcPr")
                tc.insert(0, tc_pr)
            for tag in ("tcBorders", "shd", "vAlign"):
                old = tc_pr.find(f"{W}{tag}")
                if old is not None:
                    tc_pr.remove(old)

            gs_el = tc_pr.find(f"{W}gridSpan")
            span_val = int(gs_el.get(f"{W}val", "1")) if gs_el is not None else 1

            tc_bdrs = etree.SubElement(tc_pr, f"{W}tcBorders")
            top_spec = ("single", "12") if r_idx == 0 else ("none", "0")

            if has_two_row_header and r_idx == 0:
                # Spanned group headers get a partial 0.5pt (`sz="4"`) cmidrule; stub col 0 stays clean
                if span_val > 1 or (c_idx > 0 and cell_texts[c_idx] != ""):
                    bot_spec = ("single", "4")
                else:
                    bot_spec = ("none", "0")
            elif r_idx == header_rows - 1:
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

            if r_idx < header_rows:
                col_jc = "left" if (c_idx == 0 and span_val == 1) else "center"
            elif c_idx == 0 or (n_cols == 4 and c_idx == 3 and not cell_texts[c_idx].startswith("-")):
                col_jc = "left"
            else:
                col_jc = "center"

            p_before = 0 if is_se_row else 20
            p_after = 0 if next_is_se_row else (28 if is_se_row else 20)
            p_line = 220 if (is_se_row or next_is_se_row) else 260

            for p in tc.findall(f"{W}p"):
                ppr = p.find(f"{W}pPr")
                if ppr is None:
                    ppr = etree.Element(f"{W}pPr")
                    p.insert(0, ppr)
                _set_para_layout(
                    ppr,
                    jc=col_jc,
                    before=p_before,
                    after=p_after,
                    line=p_line,
                    first_line_chars=0,
                    keep_next=next_is_se_row,
                )
                for r in p.findall(f".//{W}r"):
                    if any(anc.tag == f"{M}oMath" for anc in r.iterancestors()):
                        continue
                    rpr = r.find(f"{W}rPr")
                    if rpr is None:
                        rpr = etree.Element(f"{W}rPr")
                        r.insert(0, rpr)
                    force_bold = True if (r_idx < header_rows or is_panel_subhead) else None
                    _set_rfonts_and_black(rpr, size_half_pt=cell_font_half_pt, bold=force_bold, strip_underline=True)
                for mr in p.findall(f".//{M}r"):
                    rpr = mr.find(f"{W}rPr")
                    if rpr is None:
                        rpr = etree.Element(f"{W}rPr")
                        mr.insert(0, rpr)
                    _set_rfonts_and_black(rpr, size_half_pt=cell_font_half_pt, bold=None, is_math=True)

    return n_cols


def _clamp_paragraph_drawings(p: etree._Element, max_w_emu: int = MAX_FIG_WIDTH_EMU) -> bool:
    """Clamp all `<wp:inline>` / `<wp:anchor>` figure extents inside `p` proportionally and return True if found."""
    found = False
    for drawing in p.findall(f".//{W}drawing"):
        found = True
        for extent in drawing.findall(f".//{WP}extent"):
            cx = int(extent.get("cx", "0"))
            cy = int(extent.get("cy", "0"))
            new_cx, new_cy = _clamp_emu_size(cx, cy, max_w=max_w_emu)
            extent.set("cx", str(new_cx))
            extent.set("cy", str(new_cy))
            for a_ext in drawing.findall(f".//{PIC}spPr/{A}xfrm/{A}ext"):
                a_ext.set("cx", str(new_cx))
                a_ext.set("cy", str(new_cy))
    if p.find(f".//{W}pict") is not None:
        found = True
    return found


def _build_sect_pr(ftr_rid: str, landscape: bool = False) -> etree._Element:
    """Create a standard A4 Portrait or Landscape `<w:sectPr>` linked to centered page numbers."""
    sect_pr = etree.Element(f"{W}sectPr")
    ftr_ref = etree.SubElement(sect_pr, f"{W}footerReference")
    ftr_ref.set(f"{W}type", "default")
    ftr_ref.set(f"{R}id", ftr_rid)

    sect_type = etree.SubElement(sect_pr, f"{W}type")
    sect_type.set(f"{W}val", "nextPage")

    pg_sz = etree.SubElement(sect_pr, f"{W}pgSz")
    if landscape:
        pg_sz.set(f"{W}w", "16838")
        pg_sz.set(f"{W}h", "11906")
        pg_sz.set(f"{W}orient", "landscape")
    else:
        pg_sz.set(f"{W}w", "11906")
        pg_sz.set(f"{W}h", "16838")

    pg_mar = etree.SubElement(sect_pr, f"{W}pgMar")
    pg_mar.set(f"{W}top", "1440")
    pg_mar.set(f"{W}bottom", "1440")
    lr = "1440" if landscape else "1588"
    pg_mar.set(f"{W}left", lr)
    pg_mar.set(f"{W}right", lr)
    pg_mar.set(f"{W}header", "851")
    pg_mar.set(f"{W}footer", "992")
    pg_mar.set(f"{W}gutter", "0")
    return sect_pr


def _attach_sect_pr_to_para(p: etree._Element, sect_pr: etree._Element) -> None:
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    for old in ppr.findall(f"{W}sectPr"):
        ppr.remove(old)
    ppr.append(sect_pr)


def apply_academic_word_formatting(docx_path: Path, landscape_cols: int = 7) -> None:
    """Post-process OpenXML inside `docx_path` to enforce fonts, black color, 3-line tables, figures, & landscape."""
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
            _set_rfonts_and_black(rpr_def, size_half_pt=24, bold=False, strip_underline=True)

        for st in styles_root.findall(f"{W}style"):
            nm = sid_to_name.get(st.get(f"{W}styleId", ""), "")
            rpr = st.find(f"{W}rPr")
            if rpr is None:
                rpr = etree.SubElement(st, f"{W}rPr")
            if nm == "title":
                _set_rfonts_and_black(rpr, size_half_pt=32, bold=True, strip_underline=True)
            elif nm == "subtitle":
                _set_rfonts_and_black(rpr, size_half_pt=22, bold=False, strip_underline=True)
            elif nm == "heading 1":
                _set_rfonts_and_black(rpr, size_half_pt=28, bold=True, strip_underline=True)
            elif nm in ("heading 2", "heading 3"):
                _set_rfonts_and_black(rpr, size_half_pt=24, bold=True, strip_underline=True)
            else:
                _set_rfonts_and_black(rpr, size_half_pt=24, bold=None, strip_underline=True)

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

        # Strip explicit `<!-- landscape -->` marker paragraphs and record forced-landscape targets
        forced_landscape_elements: set[etree._Element] = set()
        raw_children = list(body)
        idx_scan = 0
        while idx_scan < len(raw_children):
            el = raw_children[idx_scan]
            if el.tag == f"{W}p" and "".join(el.itertext()).strip() == LANDSCAPE_MARKER_TOKEN:
                body.remove(el)
                # Tag the next table encountered in `raw_children`
                for look_ahead in range(idx_scan + 1, len(raw_children)):
                    cand = raw_children[look_ahead]
                    if cand.tag == f"{W}tbl":
                        forced_landscape_elements.add(cand)
                        break
            idx_scan += 1

        # Clear any pre-existing paragraph-level `<w:sectPr>` when restyling messy documents so page geometry is clean
        for old_p_sect in body.findall(f"{W}p/{W}pPr/{W}sectPr"):
            parent_ppr = old_p_sect.getparent()
            if parent_ppr is not None:
                parent_ppr.remove(old_p_sect)

        children = [c for c in body if c.tag != f"{W}sectPr"]
        tbl_col_counts: dict[etree._Element, int] = {}
        fig_para_indices: set[int] = set()

        for idx, child in enumerate(children):
            if child.tag == f"{W}tbl":
                tbl_col_counts[child] = _format_three_line_table(child)
            elif child.tag == f"{W}p" and _clamp_paragraph_drawings(child):
                fig_para_indices.add(idx)

        for idx, child in enumerate(children):
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

            has_figure = idx in fig_para_indices
            has_display_math = child.find(f"{M}oMathPara") is not None
            next_is_tbl = idx + 1 < len(children) and children[idx + 1].tag == f"{W}tbl"
            prev_is_tbl = idx > 0 and children[idx - 1].tag == f"{W}tbl"
            prev_is_fig = (idx - 1) in fig_para_indices
            prev_prev_is_fig = (idx - 2) in fig_para_indices and bool(
                re.match(r"^图\s*\d+", "".join(children[idx - 1].itertext()).strip())
            )

            is_tbl_caption = next_is_tbl and bool(re.match(r"^表\s*\d+", p_text))
            is_tbl_note = prev_is_tbl and p_text.startswith(("注：", "资料来源："))
            is_fig_caption = prev_is_fig and bool(re.match(r"^图\s*\d+", p_text))
            is_fig_note = prev_prev_is_fig and p_text.startswith(("注：", "资料来源："))
            next_is_fig_note = (
                is_fig_caption
                and idx + 1 < len(children)
                and children[idx + 1].tag == f"{W}p"
                and "".join(children[idx + 1].itertext()).strip().startswith(("注：", "资料来源："))
            )

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
            elif has_figure:
                _set_para_layout(ppr, jc="center", before=140, after=60, line=240, first_line_chars=0, keep_next=True)
                target_sz, target_bold = 21, None
            elif is_fig_caption:
                _set_para_layout(
                    ppr,
                    jc="center",
                    before=40,
                    after=40 if next_is_fig_note else 160,
                    line=300,
                    first_line_chars=0,
                    keep_next=next_is_fig_note,
                )
                target_sz, target_bold = 21, True
            elif is_fig_note:
                _set_para_layout(ppr, jc="both", before=20, after=160, line=280, first_line_chars=0, keep_next=False)
                target_sz, target_bold = 18, False
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
                _set_rfonts_and_black(rpr, size_half_pt=target_sz, bold=target_bold, strip_underline=True)

            for mr in child.findall(f".//{M}r"):
                rpr = mr.find(f"{W}rPr")
                if rpr is None:
                    rpr = etree.Element(f"{W}rPr")
                    mr.insert(0, rpr)
                _set_rfonts_and_black(rpr, size_half_pt=target_sz, bold=None, is_math=True)

        # Configure centered `10pt` footer1.xml + relationships + content types
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

        # Auto-Landscape Section Wrapping (`<!-- landscape -->` or `n_cols >= landscape_cols`)
        landscape_flags = [False] * len(children)
        for idx, child in enumerate(children):
            if child.tag != f"{W}tbl":
                continue
            n_c = tbl_col_counts.get(child, 0)
            needs_land = (child in forced_landscape_elements) or (landscape_cols > 0 and n_c >= landscape_cols)
            if not needs_land:
                continue
            landscape_flags[idx] = True
            if idx > 0 and children[idx - 1].tag == f"{W}p":
                cap_t = "".join(children[idx - 1].itertext()).strip()
                if re.match(r"^表\s*\d+", cap_t):
                    landscape_flags[idx - 1] = True
            if idx + 1 < len(children) and children[idx + 1].tag == f"{W}p":
                note_t = "".join(children[idx + 1].itertext()).strip()
                if note_t.startswith(("注：", "资料来源：")):
                    landscape_flags[idx + 1] = True

        # Group contiguous landscape runs `[seg_start .. seg_end]` and inject Portrait/Landscape `<w:sectPr>` breaks
        final_body_landscape = False
        seg_i = 0
        while seg_i < len(children):
            if not landscape_flags[seg_i]:
                seg_i += 1
                continue
            seg_start = seg_i
            while seg_i + 1 < len(children) and landscape_flags[seg_i + 1]:
                seg_i += 1
            seg_end = seg_i

            # Close Portrait section immediately BEFORE `children[seg_start]` (if preceded by Portrait content)
            if seg_start > 0:
                prev_el = children[seg_start - 1]
                if prev_el.tag == f"{W}p":
                    _attach_sect_pr_to_para(prev_el, _build_sect_pr(ftr_rid, landscape=False))
                else:
                    carrier = etree.Element(f"{W}p")
                    _attach_sect_pr_to_para(carrier, _build_sect_pr(ftr_rid, landscape=False))
                    body.insert(body.index(children[seg_start]), carrier)

            # Close Landscape section immediately AFTER `children[seg_end]` (unless it reaches document end)
            if seg_end < len(children) - 1:
                end_el = children[seg_end]
                if end_el.tag == f"{W}p":
                    _attach_sect_pr_to_para(end_el, _build_sect_pr(ftr_rid, landscape=True))
                else:
                    carrier = etree.Element(f"{W}p")
                    _attach_sect_pr_to_para(carrier, _build_sect_pr(ftr_rid, landscape=True))
                    body.insert(body.index(end_el) + 1, carrier)
            else:
                final_body_landscape = True

            seg_i += 1

        old_body_sect = body.find(f"{W}sectPr")
        if old_body_sect is not None:
            body.remove(old_body_sect)
        body.append(_build_sect_pr(ftr_rid, landscape=final_body_landscape))

        doc_tree.write(str(doc_path), xml_declaration=True, encoding="UTF-8")

        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for fpath in sorted(tmp.rglob("*")):
                if fpath.is_file():
                    zout.write(fpath, fpath.relative_to(tmp).as_posix())


def _make_inline_drawing_element(r_id: str, doc_pr_id: int, cx_emu: int, cy_emu: int, alt_text: str) -> etree._Element:
    """Construct standard OpenXML `<w:drawing><wp:inline>...</wp:inline></w:drawing>` for Pure-Python image embedding."""
    drawing = etree.Element(f"{W}drawing")
    inline = etree.SubElement(
        drawing,
        f"{WP}inline",
        distT="0",
        distB="0",
        distL="0",
        distR="0",
    )
    etree.SubElement(inline, f"{WP}extent", cx=str(cx_emu), cy=str(cy_emu))
    etree.SubElement(inline, f"{WP}effectExtent", l="0", t="0", r="0", b="0")
    etree.SubElement(inline, f"{WP}docPr", id=str(doc_pr_id), name=f"Picture {doc_pr_id}", descr=alt_text)
    c_nv = etree.SubElement(inline, f"{WP}cNvGraphicFramePr")
    etree.SubElement(c_nv, f"{A}graphicFrameLocks", noChangeAspect="1")

    graphic = etree.SubElement(inline, f"{A}graphic")
    g_data = etree.SubElement(graphic, f"{A}graphicData", uri=PIC_NS)
    pic = etree.SubElement(g_data, f"{PIC}pic")

    nv_pic = etree.SubElement(pic, f"{PIC}nvPicPr")
    etree.SubElement(nv_pic, f"{PIC}cNvPr", id=str(doc_pr_id), name=f"image{doc_pr_id}")
    etree.SubElement(nv_pic, f"{PIC}cNvPicPr")

    blip_fill = etree.SubElement(pic, f"{PIC}blipFill")
    blip = etree.SubElement(blip_fill, f"{A}blip")
    blip.set(f"{R}embed", r_id)
    stretch = etree.SubElement(blip_fill, f"{A}stretch")
    etree.SubElement(stretch, f"{A}fillRect")

    sp_pr = etree.SubElement(pic, f"{PIC}spPr")
    xfrm = etree.SubElement(sp_pr, f"{A}xfrm")
    etree.SubElement(xfrm, f"{A}off", x="0", y="0")
    etree.SubElement(xfrm, f"{A}ext", cx=str(cx_emu), cy=str(cy_emu))
    prst = etree.SubElement(sp_pr, f"{A}prstGeom", prst="rect")
    etree.SubElement(prst, f"{A}avLst")
    return drawing


def _build_pure_python_base_docx(md_path: Path, docx_path: Path) -> None:
    """Zero-Pandoc pure-Python Markdown -> OpenXML `.docx` builder with native OMML math & image embedding."""
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from latex_to_omml import latex_to_omml_element

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

    nsmap = {"w": W_NS, "m": M_NS, "r": R_NS, "wp": WP_NS, "a": A_NS, "pic": PIC_NS}
    doc_el = etree.Element(f"{W}document", nsmap=nsmap)
    body = etree.SubElement(doc_el, f"{W}body")

    media_entries: list[tuple[str, str, bytes, str]] = []  # (rId, zip_path, bytes, ext)
    inline_token_re = re.compile(r"(\$[^\$\n]+\$|\*\*[^\*]+\*\*|\*[^\*]+\*)")
    md_img_re = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$")

    def _append_rich_runs(container_p: etree._Element, text: str) -> None:
        for part in inline_token_re.split(text):
            if not part:
                continue
            if part.startswith("$") and part.endswith("$") and len(part) > 2:
                container_p.append(latex_to_omml_element(part[1:-1], display=False))
                continue
            r = etree.SubElement(container_p, f"{W}r")
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                rpr = etree.SubElement(r, f"{W}rPr")
                etree.SubElement(rpr, f"{W}b")
                part = part[2:-2]
            elif part.startswith("*") and part.endswith("*") and len(part) > 2:
                rpr = etree.SubElement(r, f"{W}rPr")
                etree.SubElement(rpr, f"{W}i")
                part = part[1:-1]
            t = etree.SubElement(r, f"{W}t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = part

    def add_para(text: str, style_id: str | None = None) -> etree._Element:
        p = etree.SubElement(body, f"{W}p")
        if style_id:
            ppr = etree.SubElement(p, f"{W}pPr")
            ps = etree.SubElement(ppr, f"{W}pStyle")
            ps.set(f"{W}val", style_id)
        _append_rich_runs(p, text)
        return p

    if title:
        add_para(title, "Title")
    if subtitle:
        add_para(subtitle, "Subtitle")

    while idx < len(lines):
        s = lines[idx].strip()
        if not s:
            idx += 1
            continue

        # Multi-line or single-line display math `$$ ... $$`
        if s.startswith("$$"):
            if s.endswith("$$") and len(s) > 4:
                latex_expr = s[2:-2].strip()
                idx += 1
            else:
                math_lines = [s[2:]]
                idx += 1
                while idx < len(lines) and not lines[idx].strip().endswith("$$"):
                    math_lines.append(lines[idx])
                    idx += 1
                if idx < len(lines):
                    math_lines.append(lines[idx].strip()[:-2])
                    idx += 1
                latex_expr = "\n".join(math_lines).strip()
            p_math = etree.SubElement(body, f"{W}p")
            p_math.append(latex_to_omml_element(latex_expr, display=True))
            continue

        # Markdown image `![alt](path)`
        m_img = md_img_re.match(s)
        if m_img:
            alt_txt, raw_img = m_img.groups()
            img_path = Path(raw_img.strip().strip("\"'"))
            if not img_path.is_absolute():
                img_path = (md_path.parent / img_path).resolve()
            if img_path.exists():
                w_px, h_px = _read_image_dimensions_px(img_path)
                cx_emu, cy_emu = _clamp_emu_size(w_px * EMU_PER_PX_96DPI, h_px * EMU_PER_PX_96DPI)
                ext = img_path.suffix.lower().lstrip(".")
                if ext == "svg":
                    img_bytes = _create_minimal_white_png(w_px, h_px)
                    ext = "png"
                else:
                    img_bytes = img_path.read_bytes()
                    if ext == "jpg":
                        ext = "jpeg"
                pic_num = len(media_entries) + 1
                r_id = f"rIdMediaImg{pic_num}"
                zip_target = f"media/image{pic_num}.{ext}"
                media_entries.append((r_id, zip_target, img_bytes, ext))
                p_img = etree.SubElement(body, f"{W}p")
                r_img = etree.SubElement(p_img, f"{W}r")
                r_img.append(_make_inline_drawing_element(r_id, pic_num, cx_emu, cy_emu, alt_txt))
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
        elif s.startswith(">"):
            add_para(s.lstrip(">").strip())
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
                        if cell_txt:
                            _append_rich_runs(p, cell_txt)
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

    media_rel_lines = [
        f'  <Relationship Id="{r_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{zip_target}"/>'
        for r_id, zip_target, _, _ in media_entries
    ]
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        + ("\n".join(media_rel_lines) + "\n" if media_rel_lines else "")
        + "</Relationships>"
    )

    used_exts = {ext for _, _, _, ext in media_entries}
    mime_map = {"png": "image/png", "jpeg": "image/jpeg", "gif": "image/gif", "bmp": "image/bmp"}
    ext_defaults = "\n".join(
        f'  <Default Extension="{ext}" ContentType="{mime_map.get(ext, "application/octet-stream")}"/>'
        for ext in sorted(used_exts)
    )
    ct_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        + (ext_defaults + "\n" if ext_defaults else "")
        + '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
        + '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
        + "</Types>"
    )

    docx_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct_xml)
        zf.writestr("_rels/.rels", rels_root)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/styles.xml", styles_xml)
        for _, zip_target, img_bytes, _ in media_entries:
            zf.writestr(f"word/{zip_target}", img_bytes)
        zf.writestr("word/document.xml", etree.tostring(doc_el, xml_declaration=True, encoding="UTF-8"))


def inspect_academic_docx(docx_path: Path, require_captions: bool = True) -> dict[str, int]:
    """Run strict OpenXML assertions on fonts, pure black color, 3-line tables, figures, & page sections."""
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
    children = [c for c in body if c.tag != f"{W}sectPr"]
    tbls = [c for c in children if c.tag == f"{W}tbl"]

    for idx, c in enumerate(children):
        if c.tag == f"{W}p" and c.find(f".//{W}drawing") is not None:
            for ext_el in c.findall(f".//{WP}extent"):
                cx_val = int(ext_el.get("cx", "0"))
                if cx_val > MAX_FIG_WIDTH_EMU:
                    raise AssertionError(f"Figure width {cx_val} EMU exceeds max {MAX_FIG_WIDTH_EMU} EMU (14.5 cm)")
            if require_captions:
                fig_cap = "".join(children[idx + 1].itertext()).strip() if idx + 1 < len(children) else ""
                if not re.match(r"^图\s*\d+", fig_cap):
                    raise AssertionError(
                        f"Figure missing centered '图 X' caption immediately below (found: '{fig_cap[:30]}')"
                    )
            continue

        if c.tag != f"{W}tbl":
            continue

        if require_captions:
            cap_txt = "".join(children[idx - 1].itertext()).strip() if idx > 0 else ""
            note_txt = "".join(children[idx + 1].itertext()).strip() if idx + 1 < len(children) else ""
            if not re.match(r"^表\s*\d+", cap_txt):
                raise AssertionError(f"Table missing centered '表 X' caption immediately above (found: '{cap_txt[:30]}')")
            if not note_txt.startswith(("注：", "资料来源：")):
                raise AssertionError(f"Table missing '注：' note immediately below (found: '{note_txt[:30]}')")

        bdrs = c.find(f"{W}tblPr/{W}tblBorders")
        assert bdrs is not None
        assert bdrs.find(f"{W}top").get(f"{W}sz") == "12"
        assert bdrs.find(f"{W}bottom").get(f"{W}sz") == "12"
        for side in ("left", "right", "insideH", "insideV"):
            assert bdrs.find(f"{W}{side}").get(f"{W}val") == "none"

        rows = c.findall(f"{W}tr")
        header_rows = sum(1 for r in rows if r.find(f"{W}trPr/{W}tblHeader") is not None)
        assert header_rows in (1, 2), f"Expected 1 or 2 tblHeader rows, found {header_rows}"

        last_hdr_bot = rows[header_rows - 1].find(f"{W}tc/{W}tcPr/{W}tcBorders/{W}bottom")
        assert last_hdr_bot is not None and last_hdr_bot.get(f"{W}val") == "single" and last_hdr_bot.get(f"{W}sz") == "6"

        if header_rows == 2:
            cmid_bots = [
                tc.find(f"{W}tcPr/{W}tcBorders/{W}bottom")
                for tc in rows[0].findall(f"{W}tc")
                if tc.find(f"{W}tcPr/{W}gridSpan") is not None
            ]
            assert cmid_bots and all(b is not None and b.get(f"{W}sz") == "4" for b in cmid_bots)

        if len(rows) > header_rows + 1:
            body_r0_bot = rows[header_rows].find(f"{W}tc/{W}tcPr/{W}tcBorders/{W}bottom")
            assert body_r0_bot is not None and body_r0_bot.get(f"{W}val") == "none"

    all_sects = doc.findall(f".//{W}sectPr")
    n_landscape = sum(
        1 for s in all_sects if (s.find(f"{W}pgSz") is not None and s.find(f"{W}pgSz").get(f"{W}orient") == "landscape")
    )
    n_figs = len(doc.findall(f".//{W}drawing"))
    n_omath = len(doc.findall(f".//{M}oMath"))
    full_text = "".join(doc.itertext())
    return {
        "tables": len(tbls),
        "landscape_sections": n_landscape,
        "figures": n_figs,
        "equations": n_omath,
        "chars": len(full_text),
    }


def verify_academic_docx(docx_path: Path, require_captions: bool = True) -> tuple[int, int]:
    """Backwards-compatible verification returning `(n_tables, n_chars)`."""
    stats = inspect_academic_docx(docx_path, require_captions=require_captions)
    return stats["tables"], stats["chars"]


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(
        description="Compile Markdown (or restyle existing .docx) into strictly formatted Academic Word (.docx, v1.1.0)."
    )
    parser.add_argument("input_path", type=Path, help="Input .md or .docx file path.")
    parser.add_argument("-o", "--output", type=Path, default=None, help="Output .docx file path.")
    parser.add_argument(
        "--restyle",
        action="store_true",
        help="Directly restyle an existing .docx file into SimSun + Times New Roman + #000000 + Three-Line Tables.",
    )
    parser.add_argument("--skip-prose-audit", action="store_true", help="Skip prose style audit before compiling.")
    parser.add_argument("--profile", type=str, default=None, help="Prose rule profile(s), e.g. 'core' or 'core,historical'.")
    parser.add_argument("--rules", type=Path, default=None, help="Custom prose_rules.json path.")
    parser.add_argument(
        "--data-source",
        type=Path,
        default=None,
        help="Optional .json/.csv/.tsv ground-truth file for empirical number cross-verification.",
    )
    parser.add_argument(
        "--skip-number-check",
        action="store_true",
        help="Skip internal table-vs-prose empirical number consistency gate.",
    )
    parser.add_argument(
        "--landscape-cols",
        type=int,
        default=7,
        help="Auto-switch tables with >= N columns onto Landscape A4 pages (default: 7; 0 to disable).",
    )
    parser.add_argument("--verify", action="store_true", help="Run OpenXML verification assertions on output .docx.")
    args = parser.parse_args()

    if not args.input_path.exists():
        print(f"ERROR: Input file not found: {args.input_path}", file=sys.stderr)
        return 2

    is_restyle = args.restyle or args.input_path.suffix.lower() == ".docx"
    sys.path.insert(0, str(Path(__file__).resolve().parent))

    if not is_restyle:
        md_text = args.input_path.read_text(encoding="utf-8")

        if not args.skip_prose_audit:
            from audit_report_prose import audit_markdown

            profiles = [p.strip() for p in args.profile.split(",") if p.strip()] if args.profile else None
            issues = audit_markdown(md_text, rules_path=args.rules, profiles=profiles)
            if issues:
                print(
                    f"Prose audit failed with {len(issues)} issue(s) (pass --skip-prose-audit to bypass):",
                    file=sys.stderr,
                )
                for item in issues:
                    print(f"  - {item}", file=sys.stderr)
                return 1

        if not args.skip_number_check or args.data_source is not None:
            from verify_numbers import verify_markdown_numbers

            num_issues = verify_markdown_numbers(md_text, data_source_path=args.data_source)
            if num_issues:
                print(
                    f"Empirical number verification failed with {len(num_issues)} issue(s) "
                    f"(pass --skip-number-check to bypass):",
                    file=sys.stderr,
                )
                for item in num_issues:
                    print(f"  - {item}", file=sys.stderr)
                return 1

    desired_out = args.output or (
        args.input_path.with_name(f"{args.input_path.stem}_academic.docx")
        if is_restyle and args.output is None
        else args.input_path.with_suffix(".docx")
    )
    out_docx, used_fallback = safe_resolve_output_path(desired_out)
    if used_fallback:
        print(
            f"[WARN] Target '{desired_out.name}' is currently locked by Word/WPS; writing safely to '{out_docx.name}'.",
            file=sys.stderr,
        )

    if is_restyle:
        if args.input_path.resolve() != out_docx.resolve():
            shutil.copy2(args.input_path, out_docx)
        engine = "Direct OpenXML Restyler (--restyle)"
    else:
        with tempfile.TemporaryDirectory() as stage_dir:
            staged_md = _preprocess_markdown_for_compile(args.input_path, Path(stage_dir))
            pandoc_bin = None if os.environ.get("FORCE_PURE_PYTHON_DOCX") == "1" else find_pandoc()
            if pandoc_bin:
                res_path = os.pathsep.join([str(args.input_path.parent.resolve()), str(Path(stage_dir).resolve()), "."])
                subprocess.run(
                    [
                        pandoc_bin,
                        str(staged_md),
                        "-f",
                        "markdown+tex_math_dollars-implicit_figures",
                        f"--resource-path={res_path}",
                        "-o",
                        str(out_docx),
                    ],
                    check=True,
                )
                engine = f"Pandoc ({Path(pandoc_bin).name}) + OpenXML Post-Processor"
            else:
                _build_pure_python_base_docx(staged_md, out_docx)
                engine = "Pure-Python Native OMML & OpenXML Builder"

    apply_academic_word_formatting(out_docx, landscape_cols=args.landscape_cols)
    print(f"[OK] Compiled Academic Word DOCX via {engine}: {out_docx}")

    if args.verify:
        stats = inspect_academic_docx(out_docx, require_captions=not is_restyle)
        print(
            f"[PASS] OpenXML Verification Gate passed: {stats['tables']} Three-Line Table(s), "
            f"{stats['landscape_sections']} Landscape Section(s), {stats['figures']} Figure(s), "
            f"{stats['equations']} OMML Equation(s), {stats['chars']:,} chars, "
            f"100% 宋体 + Times New Roman, pure black #000000."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
