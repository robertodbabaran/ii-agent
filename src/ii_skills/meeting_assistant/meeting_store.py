"""Meeting Store — JSON file persistence for meeting records.

Each meeting is stored as a separate JSON file in the meetings/ directory.
File naming: {meeting_id}.json
"""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

MEETINGS_DIR = Path(__file__).parent / "meetings"


def _ensure_dir() -> None:
    """Create meetings directory if it doesn't exist."""
    MEETINGS_DIR.mkdir(parents=True, exist_ok=True)


def save(meeting: Dict) -> Path:
    """Save a meeting record to JSON.

    Args:
        meeting: Meeting dict (must contain 'id' key).

    Returns:
        Path to the saved JSON file.
    """
    _ensure_dir()
    meeting_id = meeting["id"]
    path = MEETINGS_DIR / f"{meeting_id}.json"
    path.write_text(json.dumps(meeting, indent=2, default=str), encoding="utf-8")
    logger.info("Saved meeting %s to %s", meeting_id, path)
    return path


def load(meeting_id: str) -> Dict:
    """Load a meeting record by ID.

    Args:
        meeting_id: The meeting identifier.

    Returns:
        Meeting dict.

    Raises:
        FileNotFoundError: If the meeting file doesn't exist.
    """
    path = MEETINGS_DIR / f"{meeting_id}.json"
    if not path.exists():
        raise FileNotFoundError(f"Meeting not found: {meeting_id}")
    return json.loads(path.read_text(encoding="utf-8"))


def update(meeting_id: str, meeting: Dict) -> Path:
    """Update an existing meeting record.

    Args:
        meeting_id: The meeting identifier.
        meeting: Updated meeting dict.

    Returns:
        Path to the saved JSON file.
    """
    meeting["updated_at"] = datetime.now(timezone.utc).isoformat()
    return save(meeting)


def list_all(limit: int = 20) -> List[Dict]:
    """List all meeting records, most recent first.

    Args:
        limit: Maximum number of records to return.

    Returns:
        List of meeting summary dicts.
    """
    _ensure_dir()
    files = sorted(MEETINGS_DIR.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)

    summaries = []
    for f in files[:limit]:
        try:
            meeting = json.loads(f.read_text(encoding="utf-8"))
            summaries.append({
                "id": meeting.get("id"),
                "title": meeting.get("title"),
                "date": meeting.get("date"),
                "action_item_count": len(meeting.get("action_items", [])),
                "email_sent": meeting.get("email_sent", False),
            })
        except (json.JSONDecodeError, KeyError) as e:
            logger.warning("Skipping corrupt meeting file %s: %s", f.name, e)

    return summaries
