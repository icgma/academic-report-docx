#!/usr/bin/env python3
"""Mirror root skill files into `skills/academic-report-docx/` for GitHub catalog installers."""

from __future__ import annotations

import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MIRROR = ROOT / "skills" / "academic-report-docx"


def main() -> None:
    if MIRROR.exists():
        shutil.rmtree(MIRROR)
    MIRROR.mkdir(parents=True, exist_ok=True)

    shutil.copy2(ROOT / "SKILL.md", MIRROR / "SKILL.md")
    for folder in ("agents", "configs", "references", "examples"):
        shutil.copytree(ROOT / folder, MIRROR / folder)

    (MIRROR / "scripts").mkdir(parents=True, exist_ok=True)
    for script_name in (
        "audit_report_prose.py",
        "latex_to_omml.py",
        "md_to_academic_docx.py",
        "verify_numbers.py",
    ):
        shutil.copy2(ROOT / "scripts" / script_name, MIRROR / "scripts" / script_name)

    (MIRROR / "assets").mkdir(parents=True, exist_ok=True)
    for icon_name in ("icon.svg", "icon.png"):
        shutil.copy2(ROOT / "assets" / icon_name, MIRROR / "assets" / icon_name)

    print(f"[OK] Synced catalog skill bundle -> {MIRROR}")


if __name__ == "__main__":
    main()
