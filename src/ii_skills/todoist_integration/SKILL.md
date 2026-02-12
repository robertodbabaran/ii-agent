# Todoist Integration Skill

Standalone Todoist task management via the REST API v2.

## Actions (5)

| Action | Description |
|--------|-------------|
| `create_task` | Create a single task from natural language |
| `list_today` | List today's and overdue tasks |
| `complete_task` | Mark a task complete by ID or keyword match |
| `create_from_notes` | Parse meeting notes into multiple Todoist tasks |
| `create_from_email` | Parse email content into tasks (or create follow-up) |

## Configuration

**Required:** `Config/todoist-token.txt` containing your Todoist API token.

Get your token at: https://app.todoist.com/app/settings/integrations/developer

## API

Uses **Todoist REST API v2** at `https://api.todoist.com/rest/v2/`.

## Commands

```
"Create a Todoist task: Review Q4 earnings by Friday"
"Show my Todoist tasks for today"
"Mark the budget review task complete"
"Create tasks from these meeting notes: [paste notes]"
"Create tasks from this email: [paste email]"
```

## Task Extraction Patterns

When parsing notes or emails, the following patterns are recognized:

- `- [ ] Task description` (checkbox items)
- `Action item: Task description`
- `TODO: Task description`
- `Follow-up: Task description`
- `- @owner: Task description (due: date)`

Due dates are extracted from `(due: Friday)` or `by: next Monday` patterns.

## Integration with Meeting Assistant

This skill is complementary to the Meeting Assistant's built-in Todoist bridge.
Use **Meeting Assistant** for structured meeting workflows (notes -> calendar -> Todoist -> email).
Use **Todoist Integration** for ad-hoc task management outside of meetings.

## Examples

### Create a task with natural language due date
```python
skill.execute("create_task",
    content="Prepare board deck",
    due_string="next Wednesday",
    priority=3,
    project="Work")
```

### List today's agenda
```python
result = skill.execute("list_today")
for task in result["tasks"]:
    print(f"[{'!' * task['priority']}] {task['content']} (due: {task['due']})")
```

### Complete by keyword
```python
skill.execute("complete_task", keyword="board deck")
```

### Bulk create from meeting notes
```python
skill.execute("create_from_notes",
    notes="""
    Meeting: Q4 Planning
    - [ ] @Alice: Finalize budget model (due: Friday)
    - [ ] @Bob: Send updated comps (due: next Monday)
    Action item: Schedule follow-up with LP team
    """,
    meeting_title="Q4 Planning",
    project="Deals")
```
