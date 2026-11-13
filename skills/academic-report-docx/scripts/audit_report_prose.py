#!/usr/bin/env python3
"""Automated Academic Reporting Prose Linter (`audit_report_prose.py`).

Enforces the objective, plain-spoken Chinese academic reporting style:
1. Zero 1st/2nd-person pronouns (`我`, `我们`, `咱们`, `你`, `您`) in prose.
2. Zero obscure classical bureaucratic archaisms, colloquial metaphors, or AI buzzwords.
3. Zero unbroken CJK narrative clauses exceeding 38 Chinese characters.
4. Zero inline parenthetical English roadblocks (`（MAR）`, `（IHP）`, `（H1）`) in narrative sentences.
5. Every Markdown pipe table must be preceded by a `表 X ...` caption and followed by a `注：...` note.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

PRONOUN_BANS: tuple[str, ...] = ("我", "咱们", "你", "您")

LEXICAL_BANS: dict[str, str] = {
    "奏疏": "请改用直白的‘奏折与题本’或‘公文’",
    "藩臬": "请改用具体直白的‘布政使、按察使等省内下属官员’",
    "铨选": "请改用直白的‘日常人事任免’或‘官员选任’",
    "奏销": "请改用直白的‘地方财务报销’或‘财政核销’",
    "审拟": "请改用直白的‘刑案判决’或‘案件审理’",
    "抑或": "请改用平实的‘还是’或‘或者’",
    "确系": "请改用平实的‘确实是’或‘确实’",
    "毛坯": "严禁口语化装修比喻，请直接描述原始数据状态",
    "精装修": "严禁口语化装修比喻，请直接描述清洗后状态",
    "对账": "请改用‘核对’、‘校验’或‘复核’",
    "跨库实体链接": "请拆解AI工程黑话，直接写明按编号或字段跨库匹配",
    "定稿宽表": "请改用自然的‘标准数据表’或‘合并总表’",
    "合并流水线": "请改用自然的‘数据合并环节’或‘处理流程’",
}

CLAUSE_SPLIT_RE = re.compile(r"[，。；：？！、\n—…（）()《》“”\"']")
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
INLINE_PAREN_EN_RE = re.compile(r"[\u4e00-\u9fff]\s*[（(][A-Za-z][A-Za-z0-9_\-\s/]{0,24}[)）]")
MATH_SPAN_RE = re.compile(r"\$\$.*?\$\$|\$[^\$\n]+\$", re.DOTALL)


def strip_frontmatter(text: str) -> tuple[str, int]:
    """Strip YAML frontmatter and return `(body_text, line_offset)`."""
    lines = text.splitlines()
    if lines and lines[0].strip() == "---":
        for idx in range(1, len(lines)):
            if lines[idx].strip() == "---":
                return "\n".join(lines[idx + 1 :]), idx + 1
    return text, 0


def audit_markdown(text: str, max_clause_cjk: int = 38) -> list[str]:
    issues: list[str] = []
    body, offset = strip_frontmatter(text)
    lines = body.splitlines()

    in_code_block = False
    i = 0
    while i < len(lines):
        raw_line = lines[i]
        line_no = offset + i + 1
        stripped = raw_line.strip()

        if stripped.startswith("```"):
            in_code_block = not in_code_block
            i += 1
            continue
        if in_code_block or not stripped:
            i += 1
            continue

        # Table block validation (Three-part structure: Caption -> Pipe Table -> Note)
        if stripped.startswith("|"):
            tbl_start = i
            while i < len(lines) and lines[i].strip().startswith("|"):
                i += 1
            tbl_end = i - 1

            # Find nearest non-empty line before tbl_start
            prev_idx = tbl_start - 1
            while prev_idx >= 0 and not lines[prev_idx].strip():
                prev_idx -= 1
            prev_line = lines[prev_idx].strip() if prev_idx >= 0 else ""
            if not re.match(r"^表\s*\d+", prev_line):
                issues.append(
                    f"[三段式表格缺失表题] 第 {offset + tbl_start + 1} 行表格上方缺少以 '表 X' 开头的独立表题段落。"
                )

            # Find nearest non-empty line after tbl_end
            next_idx = tbl_end + 1
            while next_idx < len(lines) and not lines[next_idx].strip():
                next_idx += 1
            next_line = lines[next_idx].strip() if next_idx < len(lines) else ""
            if not next_line.startswith("注："):
                issues.append(
                    f"[三段式表格缺失表注] 第 {offset + tbl_end + 1} 行表格下方缺少以 '注：' 开头的独立表注段落。"
                )
            continue

        # Skip pure display math lines for narrative checks
        if stripped.startswith("$$") and stripped.endswith("$$"):
            i += 1
            continue

        clean_no_math = MATH_SPAN_RE.sub("公式", stripped)

        # 1. Pronoun check (outside quoted historical archive strings like “臣等”)
        for p in PRONOUN_BANS:
            if p in clean_no_math:
                issues.append(
                    f"[人称代词违规] 第 {line_no} 行出现第一/第二人称 '{p}'，请改为客观汇报句式。"
                )

        # 2. Banned lexical items
        for word, hint in LEXICAL_BANS.items():
            if word in clean_no_math:
                issues.append(f"[文风词表告警] 第 {line_no} 行包含 '{word}'：{hint}。")

        # 3. Formatting artifacts forbidden in Word prose
        if "`" in clean_no_math:
            issues.append(f"[格式残留告警] 第 {line_no} 行包含反引号 '`'，Word 汇报正文中请直接书写或改用斜体。")
        if "【" in clean_no_math or "】" in clean_no_math:
            issues.append(f"[格式残留告警] 第 {line_no} 行包含方头括号 '【】'，请改用标准中文层级标题或圆括号。")

        # Narrative-only checks (skip headings, table captions, and table notes)
        is_heading = stripped.startswith("#")
        is_caption = bool(re.match(r"^表\s*\d+", stripped))
        is_note = stripped.startswith("注：")
        if not (is_heading or is_caption or is_note):
            # 4. Inline parenthetical English check
            en_matches = INLINE_PAREN_EN_RE.findall(clean_no_math)
            if en_matches:
                issues.append(
                    f"[括号英文路障] 第 {line_no} 行正文夹带括号英文 {en_matches}，请用纯中文顺畅表述并将英文变量归入三线表。"
                )

            # 5. Long unbroken CJK clause check
            for clause in CLAUSE_SPLIT_RE.split(clean_no_math):
                cjk_count = len(CJK_RE.findall(clause))
                if cjk_count > max_clause_cjk:
                    preview = clause.strip()[:30] + "..."
                    issues.append(
                        f"[超长分句告警] 第 {line_no} 行出现 {cjk_count} 字无标点分句（上限 {max_clause_cjk} 字）：'{preview}'"
                    )

        i += 1

    return issues


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit Chinese academic report Markdown prose style.")
    parser.add_argument("markdown_path", type=Path, help="Path to the Markdown source file.")
    parser.add_argument("--max-clause", type=int, default=38, help="Max CJK chars per unbroken clause (default: 38).")
    args = parser.parse_args()

    if not args.markdown_path.exists():
        print(f"ERROR: File not found: {args.markdown_path}", file=sys.stderr)
        return 2

    text = args.markdown_path.read_text(encoding="utf-8")
    issues = audit_markdown(text, max_clause_cjk=args.max_clause)

    print(f"=== Academic Report Prose Audit: {args.markdown_path.name} ===")
    if not issues:
        print("Result: PASS (0 prose, pronoun, clause-length, or 3-part table issues detected!)")
        return 0

    print(f"Result: FAIL ({len(issues)} issue(s) flagged):")
    for item in issues:
        print(f"  - {item}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
