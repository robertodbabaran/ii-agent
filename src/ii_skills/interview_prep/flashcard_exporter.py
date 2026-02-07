"""
Export interview prep flashcards to Anki TSV and human-readable Markdown.

Card format matches spaced_repetition's exporter for Anki compatibility.
"""

from pathlib import Path
from typing import Dict, List, Optional

from .config import TSV_SEPARATOR


def export_tsv(
    cards: List[Dict[str, str]],
    output_path: str,
) -> str:
    """Export cards to TSV for Anki import.

    Format: question<TAB>answer<TAB>tags
    No header row — Anki expects raw card data.

    Args:
        cards: List of dicts with keys: question, answer, tags (optional).
        output_path: Absolute path to write .tsv file.

    Returns:
        Path to written file.
    """
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    lines = []
    for card in cards:
        question = card.get("question", "").replace("\t", " ").replace("\n", " ")
        answer = card.get("answer", "").replace("\t", " ").replace("\n", " ")
        tags = card.get("tags", "")
        if isinstance(tags, list):
            tags = " ".join(tags)
        line = TSV_SEPARATOR.join([question, answer, tags])
        lines.append(line)

    path.write_text("\n".join(lines), encoding="utf-8")
    return str(path)


def export_markdown(
    cards: List[Dict[str, str]],
    output_path: str,
    title: str = "Interview Prep Flashcards",
) -> str:
    """Export cards as human-readable Markdown Q&A.

    Args:
        cards: List of dicts with keys: question, answer, tags (optional).
        output_path: Absolute path to write .md file.
        title: Document title.

    Returns:
        Path to written file.
    """
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    lines = [f"# {title}", ""]

    # Group by tags if available
    tagged: Dict[str, List[Dict]] = {}
    for card in cards:
        tags = card.get("tags", "General")
        if isinstance(tags, list):
            tag_key = tags[0] if tags else "General"
        else:
            tag_key = tags or "General"
        tagged.setdefault(tag_key, []).append(card)

    for tag, tag_cards in tagged.items():
        lines.append(f"## {tag}")
        lines.append("")
        for i, card in enumerate(tag_cards, 1):
            lines.append(f"### Q{i}: {card.get('question', '')}")
            lines.append("")
            lines.append(f"**A:** {card.get('answer', '')}")
            lines.append("")
        lines.append("---")
        lines.append("")

    lines.append(f"*Total cards: {len(cards)}*")

    path.write_text("\n".join(lines), encoding="utf-8")
    return str(path)
