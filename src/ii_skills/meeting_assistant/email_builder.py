"""Email Builder — HTML email constructor for meeting summaries.

Uses inline CSS only for maximum email client compatibility.
Color scheme: #1F4E79 header, #E8F0FE callouts, #2E75B6 accents.
"""

from __future__ import annotations

from datetime import datetime, date
from typing import Dict, List


def _status_badge(status: str, due_date: str | None = None) -> str:
    """Render an inline status badge.

    Returns HTML span with colored background:
      - pending: orange
      - complete: green
      - overdue: red (pending + past due date)
    """
    today = date.today().isoformat()
    if status == "complete":
        color = "#28a745"
        label = "Complete"
    elif due_date and due_date < today and status != "complete":
        color = "#dc3545"
        label = "Overdue"
    else:
        color = "#fd7e14"
        label = "Pending"

    return (
        f'<span style="display:inline-block;padding:2px 8px;border-radius:3px;'
        f'font-size:11px;font-weight:600;color:#fff;background-color:{color};">'
        f"{label}</span>"
    )


def _action_items_table(items: List[Dict]) -> str:
    """Build the action items HTML table."""
    if not items:
        return '<p style="color:#666;font-style:italic;">No action items.</p>'

    rows = ""
    for item in items:
        badge = _status_badge(item.get("status", "pending"), item.get("due_date"))
        rows += (
            "<tr>"
            f'<td style="padding:8px 12px;border-bottom:1px solid #e0e0e0;">'
            f'{item.get("owner", "Unassigned")}</td>'
            f'<td style="padding:8px 12px;border-bottom:1px solid #e0e0e0;">'
            f'{item.get("description", "")}</td>'
            f'<td style="padding:8px 12px;border-bottom:1px solid #e0e0e0;white-space:nowrap;">'
            f'{item.get("due_date", "N/A")}</td>'
            f'<td style="padding:8px 12px;border-bottom:1px solid #e0e0e0;text-align:center;">'
            f"{badge}</td>"
            "</tr>"
        )

    return (
        '<table style="width:100%;border-collapse:collapse;margin:16px 0;">'
        "<thead>"
        "<tr>"
        '<th style="padding:10px 12px;text-align:left;background-color:#1F4E79;color:#fff;'
        'font-size:12px;text-transform:uppercase;letter-spacing:0.5px;">Owner</th>'
        '<th style="padding:10px 12px;text-align:left;background-color:#1F4E79;color:#fff;'
        'font-size:12px;text-transform:uppercase;letter-spacing:0.5px;">Description</th>'
        '<th style="padding:10px 12px;text-align:left;background-color:#1F4E79;color:#fff;'
        'font-size:12px;text-transform:uppercase;letter-spacing:0.5px;">Due Date</th>'
        '<th style="padding:10px 12px;text-align:center;background-color:#1F4E79;color:#fff;'
        'font-size:12px;text-transform:uppercase;letter-spacing:0.5px;">Status</th>'
        "</tr>"
        "</thead>"
        f"<tbody>{rows}</tbody>"
        "</table>"
    )


def _attendees_list(attendees: List[Dict]) -> str:
    """Build the attendees section."""
    if not attendees:
        return ""

    items = "".join(
        f'<li style="padding:2px 0;color:#333;">'
        f'{a.get("name", "Unknown")}'
        f'{" (" + a["email"] + ")" if a.get("email") else ""}'
        f"</li>"
        for a in attendees
    )

    return (
        f'<div style="margin:16px 0;">'
        f'<h3 style="font-size:14px;color:#1F4E79;margin:0 0 8px 0;">Attendees</h3>'
        f'<ul style="margin:0;padding-left:20px;">{items}</ul>'
        f"</div>"
    )


def build_summary_html(meeting: Dict) -> str:
    """Build a complete HTML email document for a meeting summary.

    Args:
        meeting: Full meeting dict.

    Returns:
        Complete HTML string ready for email body.
    """
    title = meeting.get("title", "Meeting Summary")
    meeting_date = meeting.get("date", datetime.now().strftime("%Y-%m-%d"))
    meeting_time = meeting.get("time", "")
    summary = meeting.get("summary", "")
    action_items = meeting.get("action_items", [])
    attendees = meeting.get("attendees", [])

    time_str = f" at {meeting_time}" if meeting_time else ""

    summary_section = ""
    if summary:
        summary_section = (
            f'<div style="background-color:#E8F0FE;border-left:4px solid #2E75B6;'
            f'padding:16px;margin:16px 0;border-radius:0 4px 4px 0;">'
            f'<h3 style="font-size:14px;color:#1F4E79;margin:0 0 8px 0;">Summary</h3>'
            f'<p style="margin:0;color:#333;line-height:1.6;">{summary}</p>'
            f"</div>"
        )

    action_table = _action_items_table(action_items)
    attendee_section = _attendees_list(attendees)

    pending_count = sum(1 for ai in action_items if ai.get("status") != "complete")
    complete_count = sum(1 for ai in action_items if ai.get("status") == "complete")
    stats_line = (
        f'<p style="font-size:12px;color:#666;margin:4px 0;">'
        f"{len(action_items)} action items "
        f"({complete_count} complete, {pending_count} pending)</p>"
    )

    return f"""\
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;font-family:Calibri,Arial,sans-serif;background-color:#f5f5f5;">
  <div style="max-width:680px;margin:0 auto;background-color:#fff;">

    <!-- Header -->
    <div style="background:linear-gradient(135deg,#1F4E79,#2E75B6);padding:28px 32px;">
      <h1 style="margin:0;font-size:22px;color:#fff;font-weight:600;">{title}</h1>
      <p style="margin:6px 0 0 0;font-size:13px;color:#B8D4F0;">{meeting_date}{time_str}</p>
    </div>

    <!-- Body -->
    <div style="padding:24px 32px;">
      {summary_section}

      <h2 style="font-size:16px;color:#1F4E79;margin:24px 0 8px 0;border-bottom:2px solid #2E75B6;\
padding-bottom:6px;">Action Items</h2>
      {stats_line}
      {action_table}

      {attendee_section}
    </div>

    <!-- Footer -->
    <div style="padding:16px 32px;background-color:#f8f9fa;border-top:1px solid #e0e0e0;">
      <p style="margin:0;font-size:11px;color:#999;text-align:center;">
        Generated by ii-agent Meeting Assistant
      </p>
    </div>

  </div>
</body>
</html>"""
