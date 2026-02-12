"""Calendar Bridge — Thin wrapper over google_calendar client.

Creates follow-up events for meeting action items with attendee support.
"""

from __future__ import annotations

import logging
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


def create_followup_event(
    action_item: Dict,
    meeting_title: str,
    owner_email: Optional[str] = None,
) -> Dict:
    """Create a Google Calendar follow-up event for an action item.

    Args:
        action_item: Action item dict with 'description', 'owner', 'due_date'.
        meeting_title: Title of the source meeting.
        owner_email: Optional email of the action item owner (added as attendee).

    Returns:
        Dict with event details on success, or error dict on failure.
    """
    try:
        from ii_skills.google_calendar import calendar_client
    except ImportError as e:
        return {"success": False, "error": f"Google Calendar not available: {e}"}

    description = action_item.get("description", "Follow-up")
    owner = action_item.get("owner", "Unassigned")
    due_date = action_item.get("due_date")

    if not due_date:
        return {"success": False, "error": "Action item has no due_date"}

    try:
        service = calendar_client.get_service()

        body = {
            "summary": f"[Follow-up] {description}",
            "start": {"date": due_date},
            "end": {"date": due_date},
            "description": (
                f"Owner: {owner}\n"
                f"Meeting: {meeting_title}\n\n"
                f"{description}"
            ),
        }

        if owner_email:
            body["attendees"] = [{"email": owner_email}]

        event = service.events().insert(calendarId="primary", body=body).execute()

        logger.info("Created follow-up event: %s", event.get("htmlLink"))
        return {
            "success": True,
            "event_id": event.get("id"),
            "link": event.get("htmlLink"),
            "summary": event.get("summary"),
        }

    except Exception as e:
        logger.exception("Failed to create follow-up event")
        return {"success": False, "error": str(e)}


def create_followup_events(
    action_items: List[Dict],
    meeting_title: str,
    attendees: Optional[List[Dict]] = None,
) -> Dict:
    """Create follow-up events for all action items with due dates.

    Args:
        action_items: List of action item dicts.
        meeting_title: Title of the source meeting.
        attendees: Optional list of attendee dicts with 'name' and 'email'.

    Returns:
        Summary dict with events_created count and event_links.
    """
    attendee_map = {}
    if attendees:
        for a in attendees:
            if a.get("name") and a.get("email"):
                attendee_map[a["name"]] = a["email"]

    results = []
    for item in action_items:
        if not item.get("due_date"):
            continue

        owner_email = attendee_map.get(item.get("owner"))
        result = create_followup_event(item, meeting_title, owner_email)
        results.append(result)

        if result.get("success") and result.get("event_id"):
            item["calendar_event_id"] = result["event_id"]

    created = [r for r in results if r.get("success")]
    return {
        "success": True,
        "events_created": len(created),
        "event_links": [r.get("link") for r in created if r.get("link")],
        "errors": [r.get("error") for r in results if not r.get("success")],
    }
