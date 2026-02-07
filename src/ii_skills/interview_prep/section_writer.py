"""
Write and read interview prep sections to/from case workspace.
"""

from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Optional

from .config import SECTIONS


# Section-to-subfolder mapping
_SECTION_SUBFOLDER = {sid: meta["subfolder"] for sid, meta in SECTIONS.items()}
_SECTION_FILENAME = {sid: meta["filename"] for sid, meta in SECTIONS.items()}


def write_section(case_folder: str, section_id: str, content: str) -> Dict:
    """Write a completed section to the correct subfolder.

    Args:
        case_folder: Absolute path to case workspace.
        section_id: Section identifier (T01-T09).
        content: Markdown content for the section.

    Returns:
        Dict with file_path, word_count, updated_at.
    """
    if section_id not in SECTIONS:
        raise ValueError(f"Unknown section_id '{section_id}'. Available: {list(SECTIONS)}")

    subfolder = _SECTION_SUBFOLDER[section_id]
    filename = _SECTION_FILENAME[section_id]

    target_dir = Path(case_folder) / subfolder
    target_dir.mkdir(parents=True, exist_ok=True)

    file_path = target_dir / filename
    file_path.write_text(content, encoding="utf-8")

    word_count = len(content.split())

    return {
        "file_path": str(file_path),
        "word_count": word_count,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }


def read_section(case_folder: str, section_id: str) -> Optional[str]:
    """Read a completed section from the case workspace.

    Returns:
        Section content as string, or None if not written yet.
    """
    if section_id not in SECTIONS:
        return None

    subfolder = _SECTION_SUBFOLDER[section_id]
    filename = _SECTION_FILENAME[section_id]
    file_path = Path(case_folder) / subfolder / filename

    if file_path.exists():
        return file_path.read_text(encoding="utf-8")
    return None


def get_all_sections(case_folder: str) -> Dict[str, Dict]:
    """Return all completed sections with metadata.

    Returns:
        Dict of section_id -> {name, content, word_count, file_path} for written sections.
    """
    result = {}
    for sid, meta in SECTIONS.items():
        content = read_section(case_folder, sid)
        if content:
            result[sid] = {
                "name": meta["name"],
                "content": content,
                "word_count": len(content.split()),
                "file_path": str(Path(case_folder) / meta["subfolder"] / meta["filename"]),
            }
    return result
