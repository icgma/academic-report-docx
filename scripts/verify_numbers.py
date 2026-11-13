#!/usr/bin/env python3
"""Empirical Number & Table-Prose Cross-Verification Gate (`verify_numbers.py`, v1.1.0).

Prevents hand-typed empirical hallucinations, digit transpositions, and sign flips in academic reports:
1. Internal Table-vs-Prose Consistency Gate (always active, zero external files required):
   - Extracts all signed coefficients, standard errors, R^2, percentages, and sample sizes (`N`) per table (`表 k`).
   - Flags sign contradictions (e.g., `-1.842` in `表 2` vs `+1.842` or unsigned `系数为 1.842` in prose).
   - Flags digit transpositions / ungrounded 3+ decimal statistics in paragraphs citing `表 k`.
2. External Ground-Truth Data Source Gate (`--data-source results.json | results.csv`):
   - Flattens all numeric leaves from JSON/CSV ground truth (including rounded 1-4 decimal forms,
     percentage `x100` forms, and `100% - pct` complements) and verifies every empirical decimal
     and table statistic traces back to computed output.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path
from typing import Any

CONVENTIONAL_THRESHOLDS: set[str] = {
    "0.1",
    "0.10",
    "0.05",
    "0.01",
    "0.005",
    "0.001",
    "0.0001",
    "90%",
    "95%",
    "99%",
    "100%",
    "100.0%",
    "100.00%",
}

SE_PAREN_RE = re.compile(r"[（(]\s*(\d+\.\d+)\s*[)）]")
SIGNED_NUM_RE = re.compile(r"(?<![A-Za-z0-9_])([+\-−]?)\s*(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?(%?)")
TABLE_REF_RE = re.compile(r"表\s*(\d+)")
COEF_UNSIGNED_CLAIM_RE = re.compile(
    r"(?:系数|估计值|边际效应)(?:保持在|仅为|约为|达到|等于|为)\s*\$?([+]?\d+\.\d{2,})"
)
EXPLICIT_N_RE = re.compile(r"\bN\s*=\s*(\d{1,3}(?:,\d{3})+|\d{3,})\b")


def _norm_mag(int_part: str, dec_part: str, pct: str) -> str:
    clean_int = int_part.replace(",", "")
    return f"{clean_int}{dec_part}{pct}"


def _extract_ground_truth_tokens(data: Any) -> set[str]:
    """Flatten all numbers from JSON/CSV structures into canonical & rounded string representations."""
    raw_floats: list[float] = []

    def _walk(obj: Any) -> None:
        if isinstance(obj, bool) or obj is None:
            return
        if isinstance(obj, (int, float)):
            raw_floats.append(float(obj))
        elif isinstance(obj, str):
            s = obj.strip().rstrip("%").replace(",", "").replace("*", "")
            try:
                raw_floats.append(float(s))
            except ValueError:
                pass
        elif isinstance(obj, dict):
            for v in obj.values():
                _walk(v)
        elif isinstance(obj, (list, tuple)):
            for item in obj:
                _walk(item)

    _walk(data)

    allowed: set[str] = set(CONVENTIONAL_THRESHOLDS)
    for val in raw_floats:
        mag = abs(val)
        if mag.is_integer():
            iv = round(mag)
            allowed.add(str(iv))
            allowed.add(f"{iv:,}")
        for d in (1, 2, 3, 4):
            s = f"{mag:.{d}f}"
            allowed.add(s)
            allowed.add(f"{s}%")
            # Complement for 0-100 percentages (e.g. 96.42% valid -> 3.58% missing)
            if 0.0 <= mag <= 100.0:
                comp = 100.0 - mag
                allowed.add(f"{comp:.{d}f}%")
                allowed.add(f"{comp:.{d}f}")
            # Proportion 0..1 expressed as percentage (e.g. 0.9642 -> 96.42%, and 1-p -> 3.58%)
            if 0.0 <= mag <= 1.0:
                pct_val = mag * 100.0
                allowed.add(f"{pct_val:.{d}f}%")
                allowed.add(f"{pct_val:.{d}f}")
                comp_pct = 100.0 - pct_val
                allowed.add(f"{comp_pct:.{d}f}%")
    return allowed


def load_data_source_tokens(path: Path) -> set[str]:
    suffix = path.suffix.lower()
    if suffix == ".json":
        payload = json.loads(path.read_text(encoding="utf-8"))
        return _extract_ground_truth_tokens(payload)
    if suffix in (".csv", ".tsv"):
        delim = "\t" if suffix == ".tsv" else ","
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f, delimiter=delim))
        return _extract_ground_truth_tokens(rows)
    raise ValueError(f"Unsupported --data-source format '{suffix}' (use .json, .csv, or .tsv)")


def verify_markdown_numbers(md_text: str, data_source_path: Path | None = None) -> list[str]:
    """Cross-verify numeric consistency between Markdown tables, prose claims, and optional ground truth."""
    issues: list[str] = []
    gt_tokens = load_data_source_tokens(data_source_path) if data_source_path else None

    lines = md_text.splitlines()
    # Skip YAML frontmatter
    start_idx = 0
    if lines and lines[0].strip() == "---":
        for idx in range(1, len(lines)):
            if lines[idx].strip() == "---":
                start_idx = idx + 1
                break

    # Pass 1: Parse all tables (`表 k`) along with their captions and notes
    table_magnitudes: dict[str, set[str]] = {}
    table_signs: dict[str, dict[str, set[str]]] = {}
    all_doc_sample_Ns: set[str] = set()
    table_line_indices: set[int] = set()

    i = start_idx
    in_code = False
    while i < len(lines):
        s = lines[i].strip()
        if s.startswith("```"):
            in_code = not in_code
            i += 1
            continue
        if in_code:
            i += 1
            continue

        if s.startswith("|"):
            tbl_start = i
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_line_indices.add(i)
                i += 1
            tbl_end = i - 1

            # Locate caption above & note below
            prev_i = tbl_start - 1
            while prev_i >= start_idx and (not lines[prev_i].strip() or lines[prev_i].strip().startswith("<!--")):
                prev_i -= 1
            cap_line = lines[prev_i].strip() if prev_i >= start_idx else ""
            if prev_i >= start_idx and re.match(r"^表\s*\d+", cap_line):
                table_line_indices.add(prev_i)

            next_i = tbl_end + 1
            while next_i < len(lines) and (not lines[next_i].strip() or lines[next_i].strip().startswith("<!--")):
                next_i += 1
            note_line = lines[next_i].strip() if next_i < len(lines) else ""
            if next_i < len(lines) and note_line.startswith("注："):
                table_line_indices.add(next_i)

            m_id = re.match(r"^表\s*(\d+)", cap_line)
            tbl_id = m_id.group(1) if m_id else f"line_{tbl_start + 1}"
            mags = table_magnitudes.setdefault(tbl_id, set())
            signs = table_signs.setdefault(tbl_id, {})

            # Record explicit N in caption
            for n_match in EXPLICIT_N_RE.findall(cap_line):
                clean_n = n_match.replace(",", "")
                mags.add(clean_n)
                all_doc_sample_Ns.add(clean_n)

            for r_line_idx in range(tbl_start, tbl_end + 1):
                r_txt = lines[r_line_idx].strip()
                cells = [c.strip() for c in r_txt.strip("|").split("|")]
                if all(re.match(r"^:?-{2,}:?$", c) for c in cells):
                    continue

                row_is_N = bool(cells and ("观测" in cells[0] or re.search(r"\bN\b", cells[0])))
                for cell in cells:
                    # Standard errors inside parentheses are unsigned magnitudes
                    se_vals = set(SE_PAREN_RE.findall(cell))
                    for se in se_vals:
                        mags.add(se)
                        if gt_tokens is not None and se not in gt_tokens:
                            issues.append(
                                f"[数据源对账失败] 第 {r_line_idx + 1} 行表 {tbl_id} 标准误 ({se}) 未在数据源中找到匹配值。"
                            )

                    cell_without_se = SE_PAREN_RE.sub(" ", cell)
                    for sign_ch, int_p, dec_p, pct_p in SIGNED_NUM_RE.findall(cell_without_se):
                        mag = _norm_mag(int_p, dec_p, pct_p)
                        mags.add(mag)
                        # Also add rounded 2-decimal and 3-decimal equivalents so prose summaries like "1.84" match "-1.842"
                        if dec_p:
                            fval = float(f"{int_p.replace(',', '')}{dec_p}")
                            for rd in (1, 2, 3):
                                mags.add(f"{fval:.{rd}f}{pct_p}")

                        if row_is_N and not dec_p and int(int_p.replace(",", "")) >= 50:
                            all_doc_sample_Ns.add(int_p.replace(",", ""))

                        if dec_p:
                            norm_sign = "-" if sign_ch in ("-", "−") else "+"
                            signs.setdefault(mag, set()).add(norm_sign)

                        # External ground-truth check on empirical decimals (`>= 2` decimal places) & explicit N rows
                        should_check_gt = (dec_p and len(dec_p) >= 3 and mag not in CONVENTIONAL_THRESHOLDS) or (
                            row_is_N and not dec_p and int(int_p.replace(",", "")) >= 50
                        )
                        if gt_tokens is not None and should_check_gt and mag not in gt_tokens:
                            issues.append(
                                f"[数据源对账失败] 第 {r_line_idx + 1} 行表 {tbl_id} 数值 '{sign_ch}{mag}' 未在数据源中找到。"
                            )
            continue

        i += 1

    # Pass 2: Verify narrative prose against cited tables and external ground truth
    in_code = False
    for idx in range(start_idx, len(lines)):
        raw = lines[idx]
        s = raw.strip()
        line_no = idx + 1
        if s.startswith("```"):
            in_code = not in_code
            continue
        if in_code or not s or idx in table_line_indices or s.startswith(("#", "$$")):
            continue

        # External ground-truth check for prose empirical decimals (`>= 2` decimal places)
        if gt_tokens is not None:
            for sign_ch, int_p, dec_p, pct_p in SIGNED_NUM_RE.findall(s):
                mag = _norm_mag(int_p, dec_p, pct_p)
                if dec_p and len(dec_p) >= 3 and mag not in CONVENTIONAL_THRESHOLDS and mag not in gt_tokens:
                    issues.append(
                        f"[数据源对账失败] 第 {line_no} 行正文实证数值 '{sign_ch}{mag}' 未在数据源中找到匹配项。"
                    )

        cited_tables = TABLE_REF_RE.findall(s)
        if not cited_tables:
            continue

        # Merge magnitudes & sign maps for all tables cited in this paragraph
        ref_mags: set[str] = set()
        ref_signs: dict[str, set[str]] = {}
        for tid in cited_tables:
            ref_mags.update(table_magnitudes.get(tid, set()))
            for m_key, s_set in table_signs.get(tid, {}).items():
                ref_signs.setdefault(m_key, set()).update(s_set)

        if not ref_mags:
            continue

        # 2a. Unsigned positive coefficient phrasing when table value is strictly negative
        for m_unsigned in COEF_UNSIGNED_CLAIM_RE.finditer(s):
            token = m_unsigned.group(1).lstrip("+")
            if token in ref_signs and ref_signs[token] == {"-"}:
                issues.append(
                    f"[正负号不一致] 第 {line_no} 行引用表 {','.join(cited_tables)} 时将负向系数写作正值 '{m_unsigned.group(0)}'（表中为 -{token}）。"
                )

        # 2b. Explicit sign contradiction & 3+ decimal exact match check
        for sign_ch, int_p, dec_p, pct_p in SIGNED_NUM_RE.findall(s):
            mag = _norm_mag(int_p, dec_p, pct_p)
            if mag in CONVENTIONAL_THRESHOLDS:
                continue

            if sign_ch in ("+", "-", "−") and dec_p:
                prose_sign = "-" if sign_ch in ("-", "−") else "+"
                if mag in ref_signs and prose_sign not in ref_signs[mag]:
                    expected_sign = next(iter(ref_signs[mag]))
                    issues.append(
                        f"[正负号不一致] 第 {line_no} 行数值 '{sign_ch}{mag}' 与表 {','.join(cited_tables)} 中的符号（{expected_sign}{mag}）相反。"
                    )

            # 3+ decimal place empirical numbers in a table-citing paragraph must exist in the cited table
            if dec_p and len(dec_p) >= 4 and mag not in ref_mags:
                issues.append(
                    f"[表文数值失配] 第 {line_no} 行引用表 {','.join(cited_tables)} 时出现高精度数值 '{sign_ch}{mag}'，但该表内无对应数值（请排查数字抄错或移位）。"
                )

        # 2c. Explicit N = ... sample size check in table-citing paragraph
        for n_str in EXPLICIT_N_RE.findall(s):
            clean_n = n_str.replace(",", "")
            if clean_n not in ref_mags and all_doc_sample_Ns and clean_n not in all_doc_sample_Ns:
                issues.append(
                    f"[样本量失配] 第 {line_no} 行样本量 'N = {n_str}' 与表 {','.join(cited_tables)} 记录的样本量不一致。"
                )

    return issues


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Cross-verify empirical numbers between Markdown prose, tables, and optional JSON/CSV data source."
    )
    parser.add_argument("markdown_path", type=Path, help="Path to Markdown report.")
    parser.add_argument(
        "--data-source",
        type=Path,
        default=None,
        help="Optional ground-truth .json / .csv / .tsv file exported by analysis scripts.",
    )
    args = parser.parse_args()

    if not args.markdown_path.exists():
        print(f"ERROR: File not found: {args.markdown_path}", file=sys.stderr)
        return 2
    if args.data_source is not None and not args.data_source.exists():
        print(f"ERROR: Data source not found: {args.data_source}", file=sys.stderr)
        return 2

    md_text = args.markdown_path.read_text(encoding="utf-8")
    issues = verify_markdown_numbers(md_text, args.data_source)

    src_label = f" + data-source={args.data_source.name}" if args.data_source else " (internal table-prose gate)"
    print(f"=== Empirical Number Verification: {args.markdown_path.name}{src_label} ===")
    if not issues:
        print("Result: PASS (0 sign flips, digit transpositions, or ungrounded empirical numbers!)")
        return 0

    print(f"Result: FAIL ({len(issues)} numerical discrepancy issue(s) detected):")
    for item in issues:
        print(f"  - {item}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
