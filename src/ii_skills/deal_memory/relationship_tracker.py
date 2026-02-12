"""Relationship Tracker — persistent contact management with weekly email reminders.

Stores contacts in a local JSON file and generates weekly digest emails
listing stale contacts (30+ days) and upcoming follow-ups.

Storage: src/ii_skills/deal_memory/relationships.json

Data model per contact:
    {
        "id": str (UUID),
        "name": str,
        "relationship_type": "professional" | "personal" | "family",
        "last_contact_date": "YYYY-MM-DD" | null,
        "key_notes": str,
        "next_follow_up": "YYYY-MM-DD" | null,
        "email": str | null,
        "phone": str | null,
        "company": str | null,
        "created_at": "YYYY-MM-DD",
        "updated_at": "YYYY-MM-DD"
    }
"""

import json
import logging
import smtplib
import sys
import uuid
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

_DATA_PATH = Path(__file__).parent / "relationships.json"

_RELATIONSHIP_TYPES = {"professional", "personal", "family"}


# ---------------------------------------------------------------------------
# Persistence helpers
# ---------------------------------------------------------------------------

def _load() -> Dict:
    """Load the relationships JSON file."""
    if not _DATA_PATH.exists():
        return {"contacts": [], "metadata": {"created_at": date.today().isoformat(), "version": "1.0.0"}}
    with open(_DATA_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save(data: Dict) -> None:
    """Write the relationships JSON file."""
    with open(_DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# CRUD operations
# ---------------------------------------------------------------------------

def add_contact(
    name: str,
    relationship_type: str = "professional",
    *,
    last_contact_date: Optional[str] = None,
    key_notes: str = "",
    next_follow_up: Optional[str] = None,
    email: Optional[str] = None,
    phone: Optional[str] = None,
    company: Optional[str] = None,
) -> Dict:
    """Add a new contact.

    Args:
        name: Full name (required).
        relationship_type: One of professional, personal, family.
        last_contact_date: ISO date YYYY-MM-DD.
        key_notes: Free-text notes.
        next_follow_up: ISO date YYYY-MM-DD.
        email: Email address.
        phone: Phone number.
        company: Company / organization.

    Returns:
        Dict with success flag and the new contact record.
    """
    if relationship_type not in _RELATIONSHIP_TYPES:
        return {
            "success": False,
            "error": f"Invalid relationship_type '{relationship_type}'. "
                     f"Must be one of: {sorted(_RELATIONSHIP_TYPES)}",
        }

    today = date.today().isoformat()
    contact = {
        "id": str(uuid.uuid4()),
        "name": name,
        "relationship_type": relationship_type,
        "last_contact_date": last_contact_date,
        "key_notes": key_notes,
        "next_follow_up": next_follow_up,
        "email": email,
        "phone": phone,
        "company": company,
        "created_at": today,
        "updated_at": today,
    }

    data = _load()
    data["contacts"].append(contact)
    _save(data)

    return {"success": True, "contact": contact}


def update_contact(contact_id: str, **fields) -> Dict:
    """Update fields on an existing contact.

    Pass any combination of: name, relationship_type, last_contact_date,
    key_notes, next_follow_up, email, phone, company.
    """
    data = _load()
    for c in data["contacts"]:
        if c["id"] == contact_id:
            if "relationship_type" in fields and fields["relationship_type"] not in _RELATIONSHIP_TYPES:
                return {"success": False, "error": f"Invalid relationship_type"}
            for key, val in fields.items():
                if key in c:
                    c[key] = val
            c["updated_at"] = date.today().isoformat()
            _save(data)
            return {"success": True, "contact": c}
    return {"success": False, "error": f"Contact {contact_id} not found"}


def delete_contact(contact_id: str) -> Dict:
    """Remove a contact by ID."""
    data = _load()
    before = len(data["contacts"])
    data["contacts"] = [c for c in data["contacts"] if c["id"] != contact_id]
    if len(data["contacts"]) == before:
        return {"success": False, "error": f"Contact {contact_id} not found"}
    _save(data)
    return {"success": True, "deleted_id": contact_id}


def get_contact(contact_id: str) -> Optional[Dict]:
    """Fetch a single contact by ID."""
    data = _load()
    for c in data["contacts"]:
        if c["id"] == contact_id:
            return c
    return None


def search_contacts(
    query: Optional[str] = None,
    relationship_type: Optional[str] = None,
) -> List[Dict]:
    """Search contacts by name/notes substring and/or type."""
    data = _load()
    results = data["contacts"]

    if relationship_type:
        results = [c for c in results if c.get("relationship_type") == relationship_type]

    if query:
        q = query.lower()
        results = [
            c for c in results
            if q in c.get("name", "").lower()
            or q in c.get("key_notes", "").lower()
            or q in (c.get("company") or "").lower()
        ]

    return results


def list_all_contacts() -> List[Dict]:
    """Return all contacts."""
    return _load()["contacts"]


def log_contact(contact_id: str, notes: str = "") -> Dict:
    """Record that you contacted someone today.

    Updates last_contact_date to today and appends to key_notes.
    """
    data = _load()
    for c in data["contacts"]:
        if c["id"] == contact_id:
            c["last_contact_date"] = date.today().isoformat()
            if notes:
                existing = c.get("key_notes", "")
                stamp = date.today().isoformat()
                c["key_notes"] = f"{existing}\n[{stamp}] {notes}".strip()
            c["updated_at"] = date.today().isoformat()
            _save(data)
            return {"success": True, "contact": c}
    return {"success": False, "error": f"Contact {contact_id} not found"}


# ---------------------------------------------------------------------------
# Reminder logic
# ---------------------------------------------------------------------------

def get_stale_contacts(days: int = 30) -> List[Dict]:
    """Contacts not reached out to in `days`+ days."""
    cutoff = date.today() - timedelta(days=days)
    results = []
    for c in list_all_contacts():
        lcd = c.get("last_contact_date")
        if not lcd:
            # Never contacted — always stale
            results.append({**c, "_days_since_contact": None})
            continue
        try:
            last = date.fromisoformat(lcd)
            if last <= cutoff:
                results.append({**c, "_days_since_contact": (date.today() - last).days})
        except (ValueError, TypeError):
            continue
    # Sort: never-contacted first, then by staleness descending
    results.sort(key=lambda x: -(x["_days_since_contact"] or 9999))
    return results


def get_upcoming_followups(days: int = 7) -> List[Dict]:
    """Contacts with follow-up dates within the next `days` days."""
    today = date.today()
    horizon = today + timedelta(days=days)
    results = []
    for c in list_all_contacts():
        fu = c.get("next_follow_up")
        if not fu:
            continue
        try:
            fu_date = date.fromisoformat(fu)
            if today <= fu_date <= horizon:
                results.append({**c, "_days_until_followup": (fu_date - today).days})
            elif fu_date < today:
                results.append({**c, "_days_until_followup": (fu_date - today).days})  # overdue = negative
        except (ValueError, TypeError):
            continue
    results.sort(key=lambda x: x["_days_until_followup"])
    return results


# ---------------------------------------------------------------------------
# Weekly email digest
# ---------------------------------------------------------------------------

def _load_gmail_credentials():
    """Borrow Gmail credentials from existing skills (cascade pattern)."""
    skills_dir = Path(__file__).resolve().parents[1]
    sources = [
        skills_dir / "networth_newsletter",
        skills_dir / "market_newsletter",
        skills_dir / "health_dashboard",
    ]
    for src in sources:
        config_path = src / "config.py"
        if not config_path.exists():
            continue
        try:
            src_str = str(src)
            if src_str not in sys.path:
                sys.path.insert(0, src_str)
            import importlib
            mod = importlib.import_module("config")
            importlib.reload(mod)
            addr = getattr(mod, "GMAIL_ADDRESS", None)
            pwd = getattr(mod, "GMAIL_APP_PASSWORD", None)
            if addr and pwd:
                return addr, pwd
        except Exception:
            continue
        finally:
            src_str = str(src)
            if src_str in sys.path:
                sys.path.remove(src_str)
    raise ImportError("No Gmail credentials found in any skill config")


def build_reminder_html() -> str:
    """Build the weekly relationship reminder HTML email."""
    stale = get_stale_contacts(days=30)
    followups = get_upcoming_followups(days=7)

    today_str = date.today().strftime("%B %d, %Y")

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family: Calibri, Arial, sans-serif; max-width: 700px; margin: 0 auto; color: #333;">

<div style="background: linear-gradient(135deg, #1F4E79, #2E75B6); padding: 24px; border-radius: 8px 8px 0 0;">
  <h1 style="color: #fff; margin: 0; font-size: 22px;">Weekly Relationship Reminder</h1>
  <p style="color: #D6E8F7; margin: 4px 0 0; font-size: 14px;">{today_str}</p>
</div>

<div style="padding: 20px; border: 1px solid #ddd; border-top: none; border-radius: 0 0 8px 8px;">
"""

    # --- Stale Contacts ---
    html += '<h2 style="color: #1F4E79; border-bottom: 2px solid #2E75B6; padding-bottom: 6px;">Haven\'t Reached Out in 30+ Days</h2>'
    if stale:
        html += '<table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">'
        html += '<tr style="background: #1F4E79; color: #fff;"><th style="padding: 8px; text-align: left;">Name</th><th style="padding: 8px;">Last Contact</th><th style="padding: 8px;">Days Ago</th><th style="padding: 8px;">Type</th></tr>'
        for i, c in enumerate(stale[:20]):  # Cap at 20
            bg = "#f9f9f9" if i % 2 == 0 else "#fff"
            lcd = c.get("last_contact_date", "Never")
            days_ago = c["_days_since_contact"]
            days_label = "Never" if days_ago is None else str(days_ago)
            html += f'<tr style="background: {bg};"><td style="padding: 8px;">{c["name"]}</td><td style="padding: 8px; text-align: center;">{lcd}</td><td style="padding: 8px; text-align: center;">{days_label}</td><td style="padding: 8px; text-align: center;">{c.get("relationship_type", "")}</td></tr>'
        html += '</table>'
        if len(stale) > 20:
            html += f'<p style="color: #777; font-style: italic;">...and {len(stale) - 20} more.</p>'
    else:
        html += '<p style="color: #777;">All contacts are up to date!</p>'

    # --- Upcoming Follow-ups ---
    html += '<h2 style="color: #1F4E79; border-bottom: 2px solid #2E75B6; padding-bottom: 6px;">Upcoming Follow-ups</h2>'
    if followups:
        html += '<table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">'
        html += '<tr style="background: #1F4E79; color: #fff;"><th style="padding: 8px; text-align: left;">Name</th><th style="padding: 8px;">Follow-up Date</th><th style="padding: 8px;">Status</th><th style="padding: 8px;">Notes</th></tr>'
        for i, c in enumerate(followups):
            bg = "#f9f9f9" if i % 2 == 0 else "#fff"
            fu_date = c.get("next_follow_up", "")
            days_until = c["_days_until_followup"]
            if days_until < 0:
                status = f'<span style="color: #c0392b; font-weight: bold;">Overdue ({abs(days_until)}d)</span>'
            elif days_until == 0:
                status = '<span style="color: #e67e22; font-weight: bold;">Today</span>'
            else:
                status = f'In {days_until} day(s)'
            # Truncate notes
            notes_preview = (c.get("key_notes", "") or "")[:80]
            if len(c.get("key_notes", "") or "") > 80:
                notes_preview += "..."
            html += f'<tr style="background: {bg};"><td style="padding: 8px;">{c["name"]}</td><td style="padding: 8px; text-align: center;">{fu_date}</td><td style="padding: 8px; text-align: center;">{status}</td><td style="padding: 8px; font-size: 12px;">{notes_preview}</td></tr>'
        html += '</table>'
    else:
        html += '<p style="color: #777;">No follow-ups scheduled this week.</p>'

    # --- Stats ---
    total = len(list_all_contacts())
    html += f"""
<div style="margin-top: 20px; padding: 12px; background: #E8F0FE; border-left: 4px solid #2E75B6; border-radius: 4px; font-size: 13px;">
  <strong>Summary:</strong> {total} total contacts &middot; {len(stale)} stale &middot; {len(followups)} follow-up(s)
</div>

<p style="color: #999; font-size: 11px; margin-top: 16px; text-align: center;">
  Generated by ii-agent Relationship Tracker
</p>

</div>
</body>
</html>"""

    return html


def send_weekly_reminder() -> Dict:
    """Build and send the weekly relationship reminder email."""
    try:
        gmail_address, gmail_password = _load_gmail_credentials()
    except ImportError as exc:
        return {"success": False, "error": str(exc)}

    html = build_reminder_html()

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Weekly Relationship Reminder - {date.today().strftime('%B %d, %Y')}"
    msg["From"] = gmail_address
    msg["To"] = gmail_address
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(gmail_address, gmail_password)
            server.send_message(msg)
        return {"success": True, "sent_to": gmail_address}
    except Exception as exc:
        logger.exception("Failed to send relationship reminder email")
        return {"success": False, "error": str(exc)}
