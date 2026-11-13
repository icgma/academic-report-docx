#!/usr/bin/env python3
"""Context-Aware Academic Reporting Prose Linter (`audit_report_prose.py`, v1.1.0).

Upgrades in v1.1.0:
1. Externalized pluggable rule profiles (`configs/prose_rules.json` + `--profile` / `--rules` CLI flags
   + YAML frontmatter `prose_profile` override).
2. Quoted archival/interview speech exemption: inline quoted speech (`“……”` / `"..."`) and blockquotes (`> ...`)
   are exempted from 1st/2nd-person pronoun bans and clause-length checks to eliminate false positives on
   interview transcripts, survey items, and historical citations.
3. Enforces BOTH academic table ordering (`表题在表上 + 表体 + 表下表注`) AND academic figure ordering
   (`图片在上 + 图题在图下`).
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_RULES_PATH = SKILL_ROOT / "configs" / "prose_rules.json"

CLAUSE_SPLIT_RE = re.compile(r"[，。；：？！、\n—…（）()《》“”\"']")
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
INLINE_PAREN_EN_RE = re.compile(r"[\u4e00-\u9fff]\s*[（(][A-Za-z][A-Za-z0-9_\-\s/]{0,24}[)）]")
MATH_SPAN_RE = re.compile(r"\$\$.*?\$\$|\$[^\$\n]+\$", re.DOTALL)
QUOTED_SPEECH_RE = re.compile(r"“[^”\n]*”|‘[^’\n]*’|\"[^\"\n]*\"")
MD_IMAGE_RE = re.compile(r"^!\[[^\]]*\]\([^)]+\)\s*$")
HTML_COMMENT_RE = re.compile(r"^<!--.*-->$")


def load_rules_config(
    rules_path: Path | None = None,
    active_profiles: list[str] | None = None,
) -> tuple[list[str], dict[str, str], int]:
    path = rules_path or DEFAULT_RULES_PATH
    if path.exists():
        cfg = json.loads(path.read_text(encoding="utf-8"))
    else:
        cfg = {
            "default_profiles": ["core"],
            "max_clause_cjk": 38,
            "pronoun_bans": ["我", "咱们", "你", "您"],
            "profiles": {"core": {"lexical_bans": {}}},
        }

    chosen = active_profiles or cfg.get("default_profiles", ["core"])
    lexical_bans: dict[str, str] = {}
    profiles_map = cfg.get("profiles", {})
    for prof in chosen:
        if prof in profiles_map:
            lexical_bans.update(profiles_map[prof].get("lexical_bans", {}))
    if "lexical_bans" in cfg and isinstance(cfg["lexical_bans"], dict):
        lexical_bans.update(cfg["lexical_bans"])

    pronouns = list(cfg.get("pronoun_bans", ["我", "咱们", "你", "您"]))
    max_clause = int(cfg.get("max_clause_cjk", 38))
    return pronouns, lexical_bans, max_clause


def parse_frontmatter(text: str) -> tuple[dict[str, str], str, int]:
    """Parse simple YAML frontmatter keys and return `(meta, body_text, line_offset)`."""
    lines = text.splitlines()
    meta: dict[str, str] = {}
    if lines and lines[0].strip() == "---":
        for idx in range(1, len(lines)):
            if lines[idx].strip() == "---":
                return meta, "\n".join(lines[idx + 1 :]), idx + 1
            if ":" in lines[idx]:
                k, v = lines[idx].split(":", 1)
                meta[k.strip()] = v.strip().strip("\"'[]")
    return meta, text, 0


def audit_markdown(
    text: str,
    max_clause_cjk: int | None = None,
    rules_path: Path | None = None,
    profiles: list[str] | None = None,
) -> list[str]:
    meta, body, offset = parse_frontmatter(text)
    if profiles is None and "prose_profile" in meta:
        profiles = [p.strip() for p in meta["prose_profile"].split(",") if p.strip()]

    pronoun_bans, lexical_bans, cfg_max_clause = load_rules_config(rules_path, profiles)
    clause_limit = max_clause_cjk if max_clause_cjk is not None else cfg_max_clause

    issues: list[str] = []
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
        if in_code_block or not stripped or HTML_COMMENT_RE.match(stripped):
            i += 1
            continue

        # Blockquote exemption (`> ...` archival / interview citation block)
        if stripped.startswith(">"):
            i += 1
            continue

        # 1. Table block validation: Caption (`表 X`) above + `注：` below
        if stripped.startswith("|"):
            tbl_start = i
            while i < len(lines) and lines[i].strip().startswith("|"):
                i += 1
            tbl_end = i - 1

            prev_idx = tbl_start - 1
            while prev_idx >= 0 and (
                not lines[prev_idx].strip() or HTML_COMMENT_RE.match(lines[prev_idx].strip())
            ):
                prev_idx -= 1
            prev_line = lines[prev_idx].strip() if prev_idx >= 0 else ""
            if not re.match(r"^表\s*\d+", prev_line):
                issues.append(
                    f"[三段式表格缺失表题] 第 {offset + tbl_start + 1} 行表格上方缺少以 '表 X' 开头的独立表题段落。"
                )

            next_idx = tbl_end + 1
            while next_idx < len(lines) and (
                not lines[next_idx].strip() or HTML_COMMENT_RE.match(lines[next_idx].strip())
            ):
                next_idx += 1
            next_line = lines[next_idx].strip() if next_idx < len(lines) else ""
            if not next_line.startswith("注："):
                issues.append(
                    f"[三段式表格缺失表注] 第 {offset + tbl_end + 1} 行表格下方缺少以 '注：' 开头的独立表注段落。"
                )
            continue

        # 2. Figure block validation: Image (`![...](...)`) -> `图 X ...` MUST be below, NOT above!
        if MD_IMAGE_RE.match(stripped):
            prev_idx = i - 1
            while prev_idx >= 0 and not lines[prev_idx].strip():
                prev_idx -= 1
            prev_line = lines[prev_idx].strip() if prev_idx >= 0 else ""
            if re.match(r"^图\s*\d+", prev_line):
                issues.append(
                    f"[图题位置违规] 第 {line_no} 行图片上方出现了图题 '{prev_line[:20]}'；学术规范要求‘表题在表上、图题在图下’。"
                )

            next_idx = i + 1
            while next_idx < len(lines) and not lines[next_idx].strip():
                next_idx += 1
            next_line = lines[next_idx].strip() if next_idx < len(lines) else ""
            if not re.match(r"^图\s*\d+", next_line):
                issues.append(
                    f"[插图缺失图题] 第 {line_no} 行图片正下方缺少以 '图 X' 开头的居中图题段落。"
                )
            i += 1
            continue

        # Skip pure display math lines
        if stripped.startswith("$$") and stripped.endswith("$$"):
            i += 1
            continue

        clean_no_math = MATH_SPAN_RE.sub("公式", stripped)
        # Mask quoted archival/interview speech before pronoun & lexical scans!
        unquoted_text = QUOTED_SPEECH_RE.sub("“引文”", clean_no_math)

        # 3. Pronoun check (outside direct quotes!)
        for p in pronoun_bans:
            if p in unquoted_text:
                issues.append(
                    f"[人称代词违规] 第 {line_no} 行出现第一/第二人称 '{p}'，请改为客观汇报句式（直接引语请置于中文双引号内）。"
                )

        # 4. Banned lexical items (outside direct quotes!)
        for word, hint in lexical_bans.items():
            if word in unquoted_text:
                issues.append(f"[文风词表告警] 第 {line_no} 行包含 '{word}'：{hint}。")

        # 5. Formatting artifacts forbidden in Word prose
        if "`" in unquoted_text:
            issues.append(f"[格式残留告警] 第 {line_no} 行包含反引号 '`'，Word 汇报正文中请直接书写或改用斜体。")
        if "【" in unquoted_text or "】" in unquoted_text:
            issues.append(f"[格式残留告警] 第 {line_no} 行包含方头括号 '【】'，请改用标准中文层级标题或圆括号。")

        is_heading = stripped.startswith("#")
        is_caption = bool(re.match(r"^(表|图)\s*\d+", stripped))
        is_note = stripped.startswith("注：")
        if not (is_heading or is_caption or is_note):
            # 6. Inline parenthetical English check (outside quotes)
            en_matches = INLINE_PAREN_EN_RE.findall(unquoted_text)
            if en_matches:
                issues.append(
                    f"[括号英文路障] 第 {line_no} 行正文夹带括号英文 {en_matches}，请用纯中文顺畅表述并将英文变量归入三线表。"
                )

            # 7. Long unbroken CJK clause check (outside quotes)
            for clause in CLAUSE_SPLIT_RE.split(unquoted_text):
                cjk_count = len(CJK_RE.findall(clause))
                if cjk_count > clause_limit:
                    preview = clause.strip()[:30] + "..."
                    issues.append(
                        f"[超长分句告警] 第 {line_no} 行出现 {cjk_count} 字无标点分句（上限 {clause_limit} 字）：'{preview}'"
                    )

        i += 1

    return issues


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="Audit Chinese academic report Markdown prose style (v1.1.0).")
    parser.add_argument("markdown_path", type=Path, help="Path to the Markdown source file.")
    parser.add_argument("--max-clause", type=int, default=None, help="Override max CJK chars per unbroken clause.")
    parser.add_argument("--rules", type=Path, default=None, help="Path to custom prose_rules.json.")
    parser.add_argument(
        "--profile",
        type=str,
        default=None,
        help="Comma-separated rule profiles from prose_rules.json (e.g. 'core' or 'core,historical').",
    )
    args = parser.parse_args()

    if not args.markdown_path.exists():
        print(f"ERROR: File not found: {args.markdown_path}", file=sys.stderr)
        return 2

    profiles = [p.strip() for p in args.profile.split(",") if p.strip()] if args.profile else None
    text = args.markdown_path.read_text(encoding="utf-8")
    issues = audit_markdown(text, max_clause_cjk=args.max_clause, rules_path=args.rules, profiles=profiles)

    print(f"=== Academic Report Prose Audit (v1.1.0): {args.markdown_path.name} ===")
    if not issues:
        print("Result: PASS (0 prose, pronoun, clause-length, figure, or 3-part table issues detected!)")
        return 0

    print(f"Result: FAIL ({len(issues)} issue(s) flagged):")
    for item in issues:
        print(f"  - {item}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
