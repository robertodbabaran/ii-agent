# Meeting Assistant Skill

Process raw meeting notes into structured action items, Google Calendar follow-ups, Todoist tasks, and formatted summary emails.

## Actions (9)

| # | Action | Purpose |
|---|--------|---------|
| 1 | `process_notes` | Parse raw notes into structured meeting record with action items |
| 2 | `create_followups` | Create Google Calendar events for action items with due dates |
| 3 | `create_todoist_tasks` | Create Todoist tasks for each action item |
| 4 | `send_summary` | Email formatted meeting summary to attendees |
| 5 | `list_meetings` | List all stored meeting records (most recent first) |
| 6 | `get_meeting` | Retrieve a specific meeting by ID |
| 7 | `full_pipeline` | Run process + calendar + todoist + email in sequence |
| 8 | `update_action_item` | Mark an action item as complete or update status |
| 9 | `complete_todoist_task` | Close a Todoist task linked to an action item |

## Setup Requirements

1. **Google Calendar OAuth** — Must be authenticated via `google_calendar` skill (`authenticate` action)
2. **Todoist API token** — Place at `Config/todoist-token.txt`
3. **Gmail credentials** — Borrows from existing skills (networth_newsletter, market_newsletter, or health_dashboard)

## Example Commands

```
"Process these meeting notes: [paste notes]"
"Create calendar follow-ups for meeting mtg_20260212_143000"
"Send the meeting summary for today's portfolio review"
"Run full meeting pipeline with these notes"
"List my recent meetings"
"Mark action item ai_001 as complete in meeting mtg_20260212_143000"
```

## Data Model

Meetings are stored as JSON files in `meetings/`. Each contains:
- Meeting metadata (title, date, time, attendees)
- Raw notes (original bullet points)
- Executive summary
- Action items (with owner, due date, status, linked calendar/todoist IDs)
- Email delivery status

## How It Works

1. **You** (via Claude Code conversation) parse the raw meeting notes and identify action items
2. Call `process_notes` with the structured data — skill stores the meeting record
3. Call `create_followups` — creates all-day Google Calendar events for each due date
4. Call `create_todoist_tasks` — creates tasks with labels and optional project assignment
5. Call `send_summary` — emails a styled HTML summary to all attendees

Or use `full_pipeline` to run all four steps in sequence.

## File Structure

```
meeting_assistant/
├── __init__.py          # Skill class (9 actions)
├── meeting_store.py     # JSON persistence
├── calendar_bridge.py   # Google Calendar wrapper
├── todoist_bridge.py    # Todoist API client
├── email_builder.py     # HTML email builder
├── meetings/            # Meeting JSON archives
├── SKILL.md             # This file
└── requirements.txt     # Dependencies
```

---

*Version: 1.0.0*
