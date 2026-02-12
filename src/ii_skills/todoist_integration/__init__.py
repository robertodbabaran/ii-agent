"""Todoist Integration Skill for ii-agent.

A standalone skill for managing Todoist tasks via natural language.
Supports creating tasks, listing today's agenda, completing tasks,
and bulk-creating tasks from meeting notes or email content.

Actions:
    create_task         - Create a single task from natural language
    list_today          - List today's and overdue tasks
    complete_task       - Mark a task as complete by ID or keyword match
    create_from_notes   - Parse meeting notes into multiple tasks
    create_from_email   - Parse email content into tasks
"""

import logging
import re
from datetime import date, datetime
from typing import Dict, List, Optional

from ii_skills import BaseSkill, register_skill

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# NLP helpers for extracting tasks from unstructured text
# ---------------------------------------------------------------------------

_ACTION_PATTERNS = [
    # "- [ ] do something by Friday"
    re.compile(r"[-*]\s*\[[ ]\]\s*(.+)", re.IGNORECASE),
    # "Action item: do something"
    re.compile(r"(?:action\s*item|todo|task|follow[- ]?up)\s*[:—–-]\s*(.+)", re.IGNORECASE),
    # "- @owner: do something (due: date)"
    re.compile(r"[-*]\s*@?\w+\s*[:—–-]\s*(.+)", re.IGNORECASE),
]

_DUE_PATTERN = re.compile(
    r"\b(?:due|by|deadline)\s*[:—–-]?\s*(.+?)(?:\)|$)", re.IGNORECASE,
)

_OWNER_PATTERN = re.compile(r"@(\w+)", re.IGNORECASE)


def _extract_tasks_from_text(text: str) -> List[Dict]:
    """Parse unstructured text into a list of task dicts.

    Returns list of {"content": str, "due_string": str|None, "owner": str|None}.
    """
    tasks: List[Dict] = []
    seen: set = set()

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        content: Optional[str] = None
        for pattern in _ACTION_PATTERNS:
            m = pattern.match(line)
            if m:
                content = m.group(1).strip()
                break

        if not content:
            continue

        # Deduplicate
        key = content.lower()
        if key in seen:
            continue
        seen.add(key)

        # Extract optional due date
        due_string: Optional[str] = None
        due_match = _DUE_PATTERN.search(content)
        if due_match:
            due_string = due_match.group(1).strip().rstrip(")")
            content = _DUE_PATTERN.sub("", content).strip().rstrip("(").strip()

        # Extract optional owner
        owner: Optional[str] = None
        owner_match = _OWNER_PATTERN.search(content)
        if owner_match:
            owner = owner_match.group(1)

        if content:
            tasks.append({
                "content": content,
                "due_string": due_string,
                "owner": owner,
            })

    return tasks


# ---------------------------------------------------------------------------
# Skill class
# ---------------------------------------------------------------------------

@register_skill
class TodoistIntegrationSkill(BaseSkill):
    """Todoist task management via REST API v2."""

    name = "todoist_integration"
    version = "1.0.0"
    description = (
        "Create, list, and complete Todoist tasks from natural language, "
        "meeting notes, or email content"
    )

    def __init__(self, config: Optional[Dict] = None):
        super().__init__(config)

    def validate_config(self) -> List[str]:
        from pathlib import Path
        token_path = Path(__file__).resolve().parents[3] / "Config" / "todoist-token.txt"
        if not token_path.exists():
            return [f"Missing Todoist token at {token_path}"]
        return []

    def get_capabilities(self) -> List[str]:
        return [
            "create_task",
            "list_today",
            "complete_task",
            "create_from_notes",
            "create_from_email",
        ]

    def execute(self, action: str, **kwargs) -> Dict:
        self._action_count += 1
        self._last_action = action
        self._last_action_at = datetime.now().isoformat()

        actions = {
            "create_task": self._create_task,
            "list_today": self._list_today,
            "complete_task": self._complete_task,
            "create_from_notes": self._create_from_notes,
            "create_from_email": self._create_from_email,
        }

        handler = actions.get(action)
        if not handler:
            return {
                "success": False,
                "error": f"Unknown action: '{action}'. Available: {list(actions)}",
            }

        try:
            return handler(**kwargs)
        except Exception as exc:
            logger.exception("Action '%s' failed", action)
            return {"success": False, "action": action, "error": str(exc)}

    # ------------------------------------------------------------------
    # Action handlers
    # ------------------------------------------------------------------

    def _create_task(
        self,
        content: str,
        *,
        description: str = "",
        due_string: Optional[str] = None,
        due_date: Optional[str] = None,
        priority: int = 1,
        labels: Optional[List[str]] = None,
        project: Optional[str] = None,
        **kwargs,
    ) -> Dict:
        """Create a single Todoist task.

        Args:
            content: Task title.
            description: Optional longer description.
            due_string: Natural-language due ("tomorrow", "next Friday").
            due_date: ISO date (YYYY-MM-DD).
            priority: 1-4 (4 = urgent).
            labels: Label names.
            project: Project name (resolved to ID).
        """
        from . import todoist_client as tc

        project_id = None
        if project:
            project_id = tc.get_project_id_by_name(project)
            if not project_id:
                return {
                    "success": False,
                    "error": f"Project '{project}' not found in Todoist",
                }

        result = tc.create_task(
            content,
            description=description,
            due_string=due_string,
            due_date=due_date,
            priority=priority,
            labels=labels,
            project_id=project_id,
        )
        return result

    def _list_today(self, **kwargs) -> Dict:
        """List today's and overdue tasks."""
        from . import todoist_client as tc

        tasks = tc.list_today()
        formatted = []
        for t in tasks:
            due_info = t.get("due", {}) or {}
            formatted.append({
                "id": t["id"],
                "content": t["content"],
                "description": t.get("description", ""),
                "due": due_info.get("date", ""),
                "due_string": due_info.get("string", ""),
                "priority": t.get("priority", 1),
                "labels": t.get("labels", []),
                "url": t.get("url", ""),
            })

        return {
            "success": True,
            "count": len(formatted),
            "tasks": formatted,
            "as_of": date.today().isoformat(),
        }

    def _complete_task(
        self,
        *,
        task_id: Optional[str] = None,
        keyword: Optional[str] = None,
        **kwargs,
    ) -> Dict:
        """Complete a task by ID or keyword match.

        Provide either task_id (exact) or keyword (searches today's tasks
        for a case-insensitive substring match and completes the first hit).
        """
        from . import todoist_client as tc

        if task_id:
            return tc.close_task(task_id)

        if not keyword:
            return {
                "success": False,
                "error": "Provide either task_id or keyword to identify the task",
            }

        # Search today's tasks for a keyword match
        tasks = tc.list_today()
        keyword_lower = keyword.lower()
        for t in tasks:
            if keyword_lower in t.get("content", "").lower():
                result = tc.close_task(t["id"])
                if result["success"]:
                    result["matched_content"] = t["content"]
                return result

        return {
            "success": False,
            "error": f"No task matching '{keyword}' found in today's list",
        }

    def _create_from_notes(
        self,
        notes: str,
        *,
        meeting_title: Optional[str] = None,
        project: Optional[str] = None,
        labels: Optional[List[str]] = None,
        **kwargs,
    ) -> Dict:
        """Parse meeting notes and create Todoist tasks for each action item.

        Recognized patterns:
            - [ ] Task description
            Action item: Task description
            Follow-up: Task description
            - @owner: Task description (due: date)

        Args:
            notes: Raw meeting notes text.
            meeting_title: Optional meeting name (added as label).
            project: Optional Todoist project name.
            labels: Additional labels to apply.
        """
        from . import todoist_client as tc

        parsed = _extract_tasks_from_text(notes)
        if not parsed:
            return {
                "success": True,
                "tasks_created": 0,
                "message": "No action items found in the notes. "
                           "Use patterns like '- [ ] task' or 'Action item: task'.",
            }

        project_id = None
        if project:
            project_id = tc.get_project_id_by_name(project)

        base_labels = list(labels or [])
        if meeting_title:
            base_labels.append(tc.slugify(meeting_title))

        results = []
        for item in parsed:
            task_labels = base_labels.copy()
            if item.get("owner"):
                task_labels.append(item["owner"])

            desc = ""
            if meeting_title:
                desc = f"From meeting: {meeting_title}"
            if item.get("owner"):
                desc += f"\nOwner: {item['owner']}"

            r = tc.create_task(
                item["content"],
                description=desc.strip(),
                due_string=item.get("due_string"),
                labels=task_labels if task_labels else None,
                project_id=project_id,
            )
            results.append(r)

        created = [r for r in results if r.get("success")]
        failed = [r for r in results if not r.get("success")]

        return {
            "success": True,
            "tasks_created": len(created),
            "tasks_failed": len(failed),
            "created": created,
            "errors": [r.get("error") for r in failed] if failed else [],
        }

    def _create_from_email(
        self,
        email_body: str,
        *,
        subject: Optional[str] = None,
        sender: Optional[str] = None,
        project: Optional[str] = None,
        labels: Optional[List[str]] = None,
        **kwargs,
    ) -> Dict:
        """Parse email content and create Todoist tasks.

        Extracts action items from the email body. If no explicit action
        items are found, creates a single follow-up task from the subject.

        Args:
            email_body: Raw email text.
            subject: Email subject line.
            sender: Sender name/address (added to description).
            project: Optional Todoist project name.
            labels: Additional labels.
        """
        from . import todoist_client as tc

        parsed = _extract_tasks_from_text(email_body)

        base_labels = list(labels or [])
        base_labels.append("email")
        if subject:
            base_labels.append(tc.slugify(subject))

        project_id = None
        if project:
            project_id = tc.get_project_id_by_name(project)

        # If no structured items found, create a single follow-up
        if not parsed and subject:
            desc = f"Follow-up from email: {subject}"
            if sender:
                desc += f"\nFrom: {sender}"

            r = tc.create_task(
                f"Follow up: {subject}",
                description=desc,
                labels=base_labels,
                project_id=project_id,
            )
            return {
                "success": r.get("success", False),
                "tasks_created": 1 if r.get("success") else 0,
                "created": [r] if r.get("success") else [],
                "errors": [r.get("error")] if not r.get("success") else [],
                "note": "No explicit action items found; created follow-up task from subject.",
            }

        if not parsed:
            return {
                "success": True,
                "tasks_created": 0,
                "message": "No action items or subject line found in the email.",
            }

        results = []
        for item in parsed:
            desc = ""
            if subject:
                desc = f"From email: {subject}"
            if sender:
                desc += f"\nFrom: {sender}"
            if item.get("owner"):
                desc += f"\nOwner: {item['owner']}"

            r = tc.create_task(
                item["content"],
                description=desc.strip(),
                due_string=item.get("due_string"),
                labels=base_labels,
                project_id=project_id,
            )
            results.append(r)

        created = [r for r in results if r.get("success")]
        failed = [r for r in results if not r.get("success")]

        return {
            "success": True,
            "tasks_created": len(created),
            "tasks_failed": len(failed),
            "created": created,
            "errors": [r.get("error") for r in failed] if failed else [],
        }


# ---------------------------------------------------------------------------
# Factory
# ---------------------------------------------------------------------------

def get_todoist(config: Optional[Dict] = None) -> TodoistIntegrationSkill:
    """Convenience factory."""
    skill = TodoistIntegrationSkill(config)
    skill.initialize()
    return skill
