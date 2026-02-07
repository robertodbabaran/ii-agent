"""
Compile all completed sections into a single interview preparation brief.
"""

from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List

from .config import SECTIONS
from .section_writer import read_section


# Order for compiled brief
_BRIEF_ORDER = ["T01", "T03", "T04", "T02", "T05", "T06", "T09", "T07"]


def compile_brief(case_folder: str, context_dict: Dict = None) -> Dict:
    """Merge all completed sections into a single brief document.

    Args:
        case_folder: Absolute path to case workspace.
        context_dict: PrepContext as dict for title page info.

    Returns:
        Dict with brief_path, total_words, sections_included.
    """
    sections_included = []
    body_parts = []

    # Title page
    title_lines = _build_title_page(context_dict)
    body_parts.append("\n".join(title_lines))

    # Table of contents
    toc_lines = ["## Table of Contents", ""]
    toc_idx = 1
    for sid in _BRIEF_ORDER:
        content = read_section(case_folder, sid)
        if content:
            name = SECTIONS[sid]["name"]
            toc_lines.append(f"{toc_idx}. {name}")
            toc_idx += 1

    # Add flashcards and appendix entries if they exist
    flashcard_md = Path(case_folder) / "05_flashcards" / "interview_prep_cards.md"
    if flashcard_md.exists():
        toc_lines.append(f"{toc_idx}. Flashcards")
        toc_idx += 1

    toc_lines.append("")
    body_parts.append("\n".join(toc_lines))

    # Main sections
    total_words = 0
    for sid in _BRIEF_ORDER:
        content = read_section(case_folder, sid)
        if content:
            name = SECTIONS[sid]["name"]
            section_text = f"---\n\n# {name}\n\n{content}"
            body_parts.append(section_text)
            word_count = len(content.split())
            total_words += word_count
            sections_included.append({
                "section_id": sid,
                "name": name,
                "word_count": word_count,
            })

    # Append flashcards if they exist
    if flashcard_md.exists():
        fc_content = flashcard_md.read_text(encoding="utf-8")
        body_parts.append(f"---\n\n# Flashcards\n\n{fc_content}")
        total_words += len(fc_content.split())

    # Write compiled brief
    brief_dir = Path(case_folder) / "06_brief"
    brief_dir.mkdir(parents=True, exist_ok=True)
    brief_path = brief_dir / "full_brief.md"

    full_text = "\n\n".join(body_parts)
    brief_path.write_text(full_text, encoding="utf-8")

    return {
        "brief_path": str(brief_path),
        "total_words": total_words,
        "sections_included": sections_included,
        "compiled_at": datetime.now(timezone.utc).isoformat(),
    }


def _build_title_page(context_dict: Dict = None) -> List[str]:
    """Build title page lines from context."""
    if not context_dict:
        return [
            "# Interview Preparation Brief",
            "",
            f"*Compiled: {datetime.now().strftime('%B %d, %Y')}*",
            "",
        ]

    company = context_dict.get("company_name", "Unknown Company")
    role = context_dict.get("role_title", "Unknown Role")
    function = context_dict.get("target_function", "")
    date = context_dict.get("interview_date", "")

    lines = [
        f"# Interview Preparation Brief",
        f"## {company} — {role}",
        "",
    ]
    if function:
        lines.append(f"**Function:** {function}")
    if date:
        lines.append(f"**Interview Date:** {date}")
    lines.append(f"**Compiled:** {datetime.now().strftime('%B %d, %Y')}")
    lines.append("")

    return lines
