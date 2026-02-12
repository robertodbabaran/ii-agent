"""Todoist REST API v2 client for ii-agent.

Handles authentication, task CRUD, and project/label operations against
https://api.todoist.com/rest/v2/.
"""

import re
import logging
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional

import requests

logger = logging.getLogger(__name__)

_TOKEN_PATH = Path(__file__).resolve().parents[3] / "Config" / "todoist-token.txt"
_API_BASE = "https://api.todoist.com/rest/v2"
_TIMEOUT = 15


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def _load_token() -> str:
    """Read the Todoist API token from disk."""
    if not _TOKEN_PATH.exists():
        raise FileNotFoundError(
            f"Todoist token not found at {_TOKEN_PATH}. "
            "Place your API token in Config/todoist-token.txt."
        )
    return _TOKEN_PATH.read_text(encoding="utf-8").strip()


def _headers() -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {_load_token()}",
        "Content-Type": "application/json",
    }


# ---------------------------------------------------------------------------
# Task operations
# ---------------------------------------------------------------------------

def create_task(
    content: str,
    *,
    description: str = "",
    due_string: Optional[str] = None,
    due_date: Optional[str] = None,
    priority: int = 1,
    labels: Optional[List[str]] = None,
    project_id: Optional[str] = None,
) -> Dict:
    """Create a new Todoist task.

    Args:
        content: Task title (required).
        description: Longer description / notes.
        due_string: Natural-language due date ("tomorrow", "next Monday").
        due_date: ISO date string (YYYY-MM-DD). Overridden by due_string if both set.
        priority: 1 (normal) to 4 (urgent).
        labels: List of label names.
        project_id: Target project ID.

    Returns:
        Dict with success flag + task details.
    """
    body: Dict = {"content": content}
    if description:
        body["description"] = description
    if due_string:
        body["due_string"] = due_string
    elif due_date:
        body["due_date"] = due_date
    if priority and priority != 1:
        body["priority"] = priority
    if labels:
        body["labels"] = labels
    if project_id:
        body["project_id"] = project_id

    try:
        resp = requests.post(
            f"{_API_BASE}/tasks", json=body, headers=_headers(), timeout=_TIMEOUT,
        )
        resp.raise_for_status()
        data = resp.json()
        return {
            "success": True,
            "task_id": data["id"],
            "url": data.get("url", ""),
            "content": data["content"],
        }
    except requests.RequestException as exc:
        logger.exception("Failed to create Todoist task")
        return {"success": False, "error": str(exc)}
    except FileNotFoundError as exc:
        return {"success": False, "error": str(exc)}


def list_tasks(
    *,
    filter_str: Optional[str] = None,
    project_id: Optional[str] = None,
    label: Optional[str] = None,
) -> List[Dict]:
    """List active tasks, optionally filtered.

    Args:
        filter_str: Todoist filter string (e.g. "today", "overdue", "p1").
        project_id: Limit to a specific project.
        label: Limit to a specific label.

    Returns:
        List of task dicts.
    """
    params: Dict = {}
    if filter_str:
        params["filter"] = filter_str
    if project_id:
        params["project_id"] = project_id
    if label:
        params["label"] = label

    try:
        resp = requests.get(
            f"{_API_BASE}/tasks", params=params, headers=_headers(), timeout=_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()
    except requests.RequestException as exc:
        logger.exception("Failed to list Todoist tasks")
        return []


def get_task(task_id: str) -> Optional[Dict]:
    """Fetch a single task by ID."""
    try:
        resp = requests.get(
            f"{_API_BASE}/tasks/{task_id}", headers=_headers(), timeout=_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()
    except requests.RequestException:
        return None


def close_task(task_id: str) -> Dict:
    """Mark a task as complete."""
    try:
        resp = requests.post(
            f"{_API_BASE}/tasks/{task_id}/close",
            headers=_headers(),
            timeout=_TIMEOUT,
        )
        resp.raise_for_status()
        return {"success": True, "task_id": task_id}
    except requests.RequestException as exc:
        logger.exception("Failed to close Todoist task %s", task_id)
        return {"success": False, "error": str(exc)}


def reopen_task(task_id: str) -> Dict:
    """Re-open a completed task."""
    try:
        resp = requests.post(
            f"{_API_BASE}/tasks/{task_id}/reopen",
            headers=_headers(),
            timeout=_TIMEOUT,
        )
        resp.raise_for_status()
        return {"success": True, "task_id": task_id}
    except requests.RequestException as exc:
        return {"success": False, "error": str(exc)}


# ---------------------------------------------------------------------------
# Project helpers
# ---------------------------------------------------------------------------

def list_projects() -> List[Dict]:
    """Return all projects."""
    try:
        resp = requests.get(
            f"{_API_BASE}/projects", headers=_headers(), timeout=_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()
    except requests.RequestException:
        return []


def get_project_id_by_name(name: str) -> Optional[str]:
    """Case-insensitive project lookup by name."""
    for proj in list_projects():
        if proj.get("name", "").lower() == name.lower():
            return proj["id"]
    return None


# ---------------------------------------------------------------------------
# Label helpers
# ---------------------------------------------------------------------------

def slugify(text: str) -> str:
    """Convert text to a Todoist-safe label slug."""
    slug = re.sub(r"[^\w\s-]", "", text.lower())
    return re.sub(r"[\s_]+", "-", slug).strip("-")[:50]


# ---------------------------------------------------------------------------
# Convenience: list today's tasks
# ---------------------------------------------------------------------------

def list_today() -> List[Dict]:
    """Return tasks due today (or overdue)."""
    return list_tasks(filter_str="today | overdue")
