"""
Meeting Assistant — Process meeting notes into action items, calendar follow-ups, and summary emails.

Actions (9 total):
  - process_notes:         Parse raw notes into structured meeting record with action items
  - create_followups:      Create Google Calendar events for action items with due dates
  - create_todoist_tasks:  Create Todoist tasks for each action item
  - send_summary:          Email formatted meeting summary + action items to attendees
  - list_meetings:         List all stored meeting records (most recent first)
  - get_meeting:           Retrieve a specific meeting by ID
  - full_pipeline:         Run process → calendar → todoist → email in sequence
  - update_action_item:    Mark an action item as complete or update its status
  - complete_todoist_task: Close a Todoist task linked to an action item
"""

from __future__ import annotations

import importlib
import logging
import os
import smtplib
import sys
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Dict, List, Optional

from ii_skills import BaseSkill, register_skill

logger = logging.getLogger(__name__)

# Gmail credential cascade — borrow from existing skills (relative to ii_skills/)
_SKILLS_DIR = Path(__file__).resolve().parents[1]
_CREDENTIAL_SOURCES = [
    _SKILLS_DIR / "networth_newsletter",
    _SKILLS_DIR / "market_newsletter",
    _SKILLS_DIR / "health_dashboard",
]

DEFAULT_RECIPIENT = os.environ.get("DEFAULT_RECIPIENT", "")


def _load_gmail_credentials() -> tuple:
    """Load Gmail credentials from existing skill configs.

    Returns:
        Tuple of (gmail_address, gmail_app_password).

    Raises:
        ImportError: If no credentials can be found.
    """
    for source in _CREDENTIAL_SOURCES:
        config_path = source / "config.py"
        if config_path.exists():
            sys.path.insert(0, str(source))
            try:
                config_module = importlib.import_module("config")
                importlib.reload(config_module)
                address = getattr(config_module, "GMAIL_ADDRESS", None)
                password = getattr(config_module, "GMAIL_APP_PASSWORD", None)
                if address and password:
                    return (address, password)
            except Exception as e:
                logger.debug("Failed to load credentials from %s: %s", source, e)
            finally:
                if str(source) in sys.path:
                    sys.path.remove(str(source))

    raise ImportError(
        "Gmail credentials not found. Ensure at least one skill "
        "(networth_newsletter, market_newsletter, or health_dashboard) "
        "has a config.py with GMAIL_ADDRESS and GMAIL_APP_PASSWORD."
    )


@register_skill
class MeetingAssistantSkill(BaseSkill):
    """Process meeting notes into action items, calendar follow-ups, and summary emails."""

    name = "meeting_assistant"
    version = "1.0.0"
    description = (
        "Meeting notes pipeline: parse raw notes into structured records, "
        "create Google Calendar follow-ups, Todoist tasks, and send formatted "
        "summary emails to attendees."
    )

    SKILL_DIR = Path(__file__).parent

    def __init__(self, config: Optional[Dict] = None):
        super().__init__(config)

    def validate_config(self) -> List[str]:
        issues = []
        try:
            _load_gmail_credentials()
        except ImportError as e:
            issues.append(str(e))
        return issues

    def get_capabilities(self) -> List[str]:
        return [
            "process_notes",
            "create_followups",
            "create_todoist_tasks",
            "send_summary",
            "list_meetings",
            "get_meeting",
            "full_pipeline",
            "update_action_item",
            "complete_todoist_task",
        ]

    def execute(self, action: str, **kwargs) -> Dict:
        """Route action to handler."""
        self._action_count += 1
        self._last_action = action
        self._last_action_at = datetime.now(timezone.utc).isoformat()

        actions = {
            "process_notes": self._process_notes,
            "create_followups": self._create_followups,
            "create_todoist_tasks": self._create_todoist_tasks,
            "send_summary": self._send_summary,
            "list_meetings": self._list_meetings,
            "get_meeting": self._get_meeting,
            "full_pipeline": self._full_pipeline,
            "update_action_item": self._update_action_item,
            "complete_todoist_task": self._complete_todoist_task,
        }

        handler = actions.get(action)
        if not handler:
            return {
                "success": False,
                "error": f"Unknown action: '{action}'. Available: {list(actions.keys())}",
            }

        try:
            return handler(**kwargs)
        except Exception as e:
            logger.exception("Action '%s' failed", action)
            return {"success": False, "action": action, "error": str(e)}

    # ─────────────────────────────────────────────────────────────
    # Action 1: process_notes
    # ─────────────────────────────────────────────────────────────

    def _process_notes(self, **kwargs) -> Dict:
        """Parse raw meeting notes into a structured meeting record.

        Required kwargs:
            title (str): Meeting name.
            raw_notes (str): Bullet-point meeting notes.
            action_items (list[dict]): Each with description, owner, due_date, status.

        Optional kwargs:
            date (str): YYYY-MM-DD, defaults to today.
            time (str): HH:MM, defaults to now.
            attendees (list[dict]): Each with name and email.
            summary (str): Executive summary.
        """
        from . import meeting_store

        title = kwargs.get("title")
        raw_notes = kwargs.get("raw_notes")
        action_items = kwargs.get("action_items")

        if not title:
            return {"success": False, "error": "Missing required parameter: title"}
        if not raw_notes:
            return {"success": False, "error": "Missing required parameter: raw_notes"}
        if not action_items:
            return {"success": False, "error": "Missing required parameter: action_items"}

        now = datetime.now(timezone.utc)
        meeting_date = kwargs.get("date", now.strftime("%Y-%m-%d"))
        meeting_time = kwargs.get("time", now.strftime("%H:%M"))
        attendees = kwargs.get("attendees", [])
        summary = kwargs.get("summary", "")

        # Generate meeting ID
        ts = now.strftime("%Y%m%d_%H%M%S")
        meeting_id = f"mtg_{ts}"

        # Assign IDs to action items
        for i, item in enumerate(action_items, 1):
            item.setdefault("id", f"ai_{i:03d}")
            item.setdefault("status", "pending")
            item.setdefault("calendar_event_id", None)
            item.setdefault("todoist_task_id", None)

        meeting = {
            "id": meeting_id,
            "title": title,
            "date": meeting_date,
            "time": meeting_time,
            "attendees": attendees,
            "raw_notes": raw_notes,
            "summary": summary,
            "action_items": action_items,
            "calendar_events_created": [],
            "email_sent": False,
            "email_sent_at": None,
            "created_at": now.isoformat(),
            "updated_at": now.isoformat(),
        }

        meeting_store.save(meeting)

        return {
            "success": True,
            "meeting_id": meeting_id,
            "action_item_count": len(action_items),
        }

    # ─────────────────────────────────────────────────────────────
    # Action 2: create_followups
    # ─────────────────────────────────────────────────────────────

    def _create_followups(self, **kwargs) -> Dict:
        """Create Google Calendar events for action items with due dates.

        Required kwargs:
            meeting_id (str): The meeting identifier.
        """
        from . import calendar_bridge, meeting_store

        meeting_id = kwargs.get("meeting_id")
        if not meeting_id:
            return {"success": False, "error": "Missing required parameter: meeting_id"}

        try:
            meeting = meeting_store.load(meeting_id)
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}

        result = calendar_bridge.create_followup_events(
            action_items=meeting["action_items"],
            meeting_title=meeting["title"],
            attendees=meeting.get("attendees"),
        )

        # Persist updated action items (with calendar_event_ids)
        meeting["calendar_events_created"] = result.get("event_links", [])
        meeting_store.update(meeting_id, meeting)

        return result

    # ─────────────────────────────────────────────────────────────
    # Action 3: create_todoist_tasks
    # ─────────────────────────────────────────────────────────────

    def _create_todoist_tasks(self, **kwargs) -> Dict:
        """Create Todoist tasks for each action item.

        Required kwargs:
            meeting_id (str): The meeting identifier.

        Optional kwargs:
            project_name (str): Todoist project name to assign tasks to.
            labels (list[str]): Additional labels for all tasks.
        """
        from . import meeting_store, todoist_bridge

        meeting_id = kwargs.get("meeting_id")
        if not meeting_id:
            return {"success": False, "error": "Missing required parameter: meeting_id"}

        try:
            meeting = meeting_store.load(meeting_id)
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}

        project_name = kwargs.get("project_name")
        extra_labels = kwargs.get("labels", [])

        # Resolve project ID if name provided
        project_id = None
        if project_name:
            project_id = todoist_bridge.get_project_id_by_name(project_name)

        title_slug = todoist_bridge._slugify(meeting["title"])
        base_labels = ["meeting", title_slug] + extra_labels

        results = []
        for item in meeting["action_items"]:
            description = (
                f"Owner: {item.get('owner', 'Unassigned')}\n"
                f"Meeting: {meeting['title']} ({meeting['date']})"
            )
            result = todoist_bridge.create_task(
                content=item["description"],
                due_date=item.get("due_date"),
                description=description,
                labels=base_labels,
                project_id=project_id,
            )
            results.append(result)

            if result.get("success") and result.get("task_id"):
                item["todoist_task_id"] = result["task_id"]

        meeting_store.update(meeting_id, meeting)

        created = [r for r in results if r.get("success")]
        return {
            "success": True,
            "tasks_created": len(created),
            "task_urls": [r.get("url") for r in created if r.get("url")],
            "errors": [r.get("error") for r in results if not r.get("success")],
        }

    # ─────────────────────────────────────────────────────────────
    # Action 4: send_summary
    # ─────────────────────────────────────────────────────────────

    def _send_summary(self, **kwargs) -> Dict:
        """Email formatted meeting summary to attendees.

        Required kwargs:
            meeting_id (str): The meeting identifier.

        Optional kwargs:
            to_emails (list[str]): Override recipient list. Defaults to all attendees.
        """
        from . import email_builder, meeting_store

        meeting_id = kwargs.get("meeting_id")
        if not meeting_id:
            return {"success": False, "error": "Missing required parameter: meeting_id"}

        try:
            meeting = meeting_store.load(meeting_id)
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}

        # Load credentials early so we can use sender as fallback recipient
        try:
            gmail_address, gmail_password = _load_gmail_credentials()
        except ImportError as e:
            return {"success": False, "error": str(e)}

        fallback_recipient = DEFAULT_RECIPIENT or gmail_address

        # Build recipient list
        to_emails = kwargs.get("to_emails")
        if not to_emails:
            to_emails = [
                a["email"]
                for a in meeting.get("attendees", [])
                if a.get("email")
            ]
        if not to_emails:
            to_emails = [fallback_recipient]

        # Always CC the user
        cc_set = {fallback_recipient} - set(to_emails)

        html = email_builder.build_summary_html(meeting)
        subject = f"Meeting Summary: {meeting['title']} ({meeting['date']})"

        try:
            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"] = gmail_address
            msg["To"] = ", ".join(to_emails)
            if cc_set:
                msg["Cc"] = ", ".join(cc_set)
            msg.attach(MIMEText(html, "html"))

            all_recipients = list(set(to_emails) | cc_set)

            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(gmail_address, gmail_password)
                server.send_message(msg, to_addrs=all_recipients)

            # Update meeting record
            meeting["email_sent"] = True
            meeting["email_sent_at"] = datetime.now(timezone.utc).isoformat()
            meeting_store.update(meeting_id, meeting)

            logger.info("Sent meeting summary to %s", all_recipients)
            return {
                "success": True,
                "sent_to": all_recipients,
                "subject": subject,
            }

        except Exception as e:
            logger.exception("Failed to send meeting summary email")
            return {"success": False, "error": str(e)}

    # ─────────────────────────────────────────────────────────────
    # Action 5: list_meetings
    # ─────────────────────────────────────────────────────────────

    def _list_meetings(self, **kwargs) -> Dict:
        """List all stored meeting records.

        Optional kwargs:
            limit (int): Max records to return (default 20).
        """
        from . import meeting_store

        limit = kwargs.get("limit", 20)
        meetings = meeting_store.list_all(limit=limit)
        return {
            "success": True,
            "meetings": meetings,
            "count": len(meetings),
        }

    # ─────────────────────────────────────────────────────────────
    # Action 6: get_meeting
    # ─────────────────────────────────────────────────────────────

    def _get_meeting(self, **kwargs) -> Dict:
        """Retrieve a specific meeting by ID.

        Required kwargs:
            meeting_id (str): The meeting identifier.
        """
        from . import meeting_store

        meeting_id = kwargs.get("meeting_id")
        if not meeting_id:
            return {"success": False, "error": "Missing required parameter: meeting_id"}

        try:
            meeting = meeting_store.load(meeting_id)
            return {"success": True, "meeting": meeting}
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}

    # ─────────────────────────────────────────────────────────────
    # Action 7: full_pipeline
    # ─────────────────────────────────────────────────────────────

    def _full_pipeline(self, **kwargs) -> Dict:
        """Run the full meeting pipeline: process → calendar → todoist → email.

        Accepts all process_notes kwargs plus:
            todoist (bool): Whether to create Todoist tasks (default True).
            project_name (str): Todoist project name.
            to_emails (list[str]): Override email recipients.
        """
        todoist_enabled = kwargs.pop("todoist", True)
        project_name = kwargs.pop("project_name", None)
        to_emails = kwargs.pop("to_emails", None)

        results = {}

        # Step 1: Process notes
        process_result = self._process_notes(**kwargs)
        results["process_notes"] = process_result
        if not process_result.get("success"):
            return {"success": False, "step_failed": "process_notes", "results": results}

        meeting_id = process_result["meeting_id"]

        # Step 2: Create calendar follow-ups
        cal_result = self._create_followups(meeting_id=meeting_id)
        results["create_followups"] = cal_result

        # Step 3: Create Todoist tasks (optional)
        if todoist_enabled:
            todoist_result = self._create_todoist_tasks(
                meeting_id=meeting_id,
                project_name=project_name,
            )
            results["create_todoist_tasks"] = todoist_result

        # Step 4: Send summary email
        send_kwargs = {"meeting_id": meeting_id}
        if to_emails:
            send_kwargs["to_emails"] = to_emails
        email_result = self._send_summary(**send_kwargs)
        results["send_summary"] = email_result

        return {
            "success": True,
            "meeting_id": meeting_id,
            "results": results,
        }

    # ─────────────────────────────────────────────────────────────
    # Action 8: update_action_item
    # ─────────────────────────────────────────────────────────────

    def _update_action_item(self, **kwargs) -> Dict:
        """Update an action item's status.

        Required kwargs:
            meeting_id (str): The meeting identifier.
            action_item_id (str): The action item ID (e.g. 'ai_001').
            status (str): New status ('pending', 'complete').
        """
        from . import meeting_store, todoist_bridge

        meeting_id = kwargs.get("meeting_id")
        action_item_id = kwargs.get("action_item_id")
        new_status = kwargs.get("status")

        if not meeting_id:
            return {"success": False, "error": "Missing required parameter: meeting_id"}
        if not action_item_id:
            return {"success": False, "error": "Missing required parameter: action_item_id"}
        if not new_status:
            return {"success": False, "error": "Missing required parameter: status"}

        try:
            meeting = meeting_store.load(meeting_id)
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}

        # Find the action item
        target = None
        for item in meeting["action_items"]:
            if item.get("id") == action_item_id:
                target = item
                break

        if not target:
            return {"success": False, "error": f"Action item not found: {action_item_id}"}

        target["status"] = new_status

        # If completing and has a Todoist task, close it
        todoist_result = None
        if new_status == "complete" and target.get("todoist_task_id"):
            todoist_result = todoist_bridge.close_task(target["todoist_task_id"])

        meeting_store.update(meeting_id, meeting)

        result = {
            "success": True,
            "action_item_id": action_item_id,
            "new_status": new_status,
        }
        if todoist_result:
            result["todoist_closed"] = todoist_result.get("success", False)

        return result

    # ─────────────────────────────────────────────────────────────
    # Action 9: complete_todoist_task
    # ─────────────────────────────────────────────────────────────

    def _complete_todoist_task(self, **kwargs) -> Dict:
        """Close a Todoist task linked to an action item.

        Required kwargs:
            meeting_id (str): The meeting identifier.
            action_item_id (str): The action item ID.
        """
        from . import meeting_store, todoist_bridge

        meeting_id = kwargs.get("meeting_id")
        action_item_id = kwargs.get("action_item_id")

        if not meeting_id:
            return {"success": False, "error": "Missing required parameter: meeting_id"}
        if not action_item_id:
            return {"success": False, "error": "Missing required parameter: action_item_id"}

        try:
            meeting = meeting_store.load(meeting_id)
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}

        target = None
        for item in meeting["action_items"]:
            if item.get("id") == action_item_id:
                target = item
                break

        if not target:
            return {"success": False, "error": f"Action item not found: {action_item_id}"}

        task_id = target.get("todoist_task_id")
        if not task_id:
            return {"success": False, "error": "No Todoist task linked to this action item"}

        result = todoist_bridge.close_task(task_id)

        if result.get("success"):
            target["status"] = "complete"
            meeting_store.update(meeting_id, meeting)

        return {
            "success": result.get("success", False),
            "action_item_id": action_item_id,
            "todoist_task_id": task_id,
            "error": result.get("error"),
        }


# ── Convenience accessor ──────────────────────────────────────

def get_skill(config: Optional[Dict] = None) -> MeetingAssistantSkill:
    """Get an initialized MeetingAssistantSkill instance."""
    skill = MeetingAssistantSkill(config)
    skill.initialize()
    return skill
