"""Todoist Bridge — Todoist API v1 client for task management.

Creates and closes tasks linked to meeting action items.
API docs: https://developer.todoist.com/rest/v1/
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Dict, List, Optional

import requests

logger = logging.getLogger(__name__)

# Token location: ii-agent/Config/todoist-token.txt
_TOKEN_PATH = Path(__file__).resolve().parents[3] / "Config" / "todoist-token.txt"
_API_BASE = "https://api.todoist.com/api/v1"


def _load_token() -> str:
    """Load Todoist API token from Config/todoist-token.txt.

    Returns:
        API token string.

    Raises:
        FileNotFoundError: If token file doesn't exist.
    """
    if not _TOKEN_PATH.exists():
        raise FileNotFoundError(
            f"Todoist token not found at {_TOKEN_PATH}. "
            "Ensure Config/todoist-token.txt exists."
        )
    return _TOKEN_PATH.read_text(encoding="utf-8").strip()


def _get_headers() -> Dict[str, str]:
    """Build authorization headers."""
    return {
        "Authorization": f"Bearer {_load_token()}",
        "Content-Type": "application/json",
    }


def _slugify(text: str) -> str:
    """Convert text to a URL-safe slug for labels."""
    slug = re.sub(r"[^\w\s-]", "", text.lower())
    return re.sub(r"[\s_]+", "-", slug).strip("-")[:50]


def create_task(
    content: str,
    due_date: Optional[str] = None,
    description: Optional[str] = None,
    labels: Optional[List[str]] = None,
    project_id: Optional[str] = None,
) -> Dict:
    """Create a Todoist task.

    Args:
        content: Task title.
        due_date: Due date in YYYY-MM-DD format.
        description: Task description/notes.
        labels: List of label strings.
        project_id: Project ID to assign task to.

    Returns:
        Dict with task id and url on success, or error dict on failure.
    """
    try:
        body: Dict = {"content": content}
        if due_date:
            body["due_date"] = due_date
        if description:
            body["description"] = description
        if labels:
            body["labels"] = labels
        if project_id:
            body["project_id"] = project_id

        resp = requests.post(
            f"{_API_BASE}/tasks",
            headers=_get_headers(),
            json=body,
            timeout=15,
        )
        resp.raise_for_status()
        task = resp.json()

        logger.info("Created Todoist task: %s (id=%s)", content, task.get("id"))
        return {
            "success": True,
            "task_id": task.get("id"),
            "url": task.get("url"),
            "content": task.get("content"),
        }

    except requests.RequestException as e:
        logger.exception("Failed to create Todoist task")
        return {"success": False, "error": str(e)}
    except FileNotFoundError as e:
        return {"success": False, "error": str(e)}


def close_task(task_id: str) -> Dict:
    """Close (complete) a Todoist task.

    Args:
        task_id: The Todoist task ID.

    Returns:
        Dict with success status.
    """
    try:
        resp = requests.post(
            f"{_API_BASE}/tasks/{task_id}/close",
            headers=_get_headers(),
            timeout=15,
        )
        resp.raise_for_status()
        logger.info("Closed Todoist task: %s", task_id)
        return {"success": True, "task_id": task_id}

    except requests.RequestException as e:
        logger.exception("Failed to close Todoist task")
        return {"success": False, "error": str(e)}
    except FileNotFoundError as e:
        return {"success": False, "error": str(e)}


def list_projects() -> List[Dict]:
    """List all Todoist projects.

    Returns:
        List of project dicts, or empty list on failure.
    """
    try:
        resp = requests.get(
            f"{_API_BASE}/projects",
            headers=_get_headers(),
            timeout=15,
        )
        resp.raise_for_status()
        return resp.json()

    except (requests.RequestException, FileNotFoundError) as e:
        logger.exception("Failed to list Todoist projects")
        return []


def get_project_id_by_name(name: str) -> Optional[str]:
    """Find a project ID by name (case-insensitive).

    Args:
        name: Project name to search for.

    Returns:
        Project ID string, or None if not found.
    """
    projects = list_projects()
    name_lower = name.lower()
    for p in projects:
        if p.get("name", "").lower() == name_lower:
            return p.get("id")
    return None
