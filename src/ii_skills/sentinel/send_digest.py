#!/usr/bin/env python3
"""
SENTINEL — Weekly AI Agent Ecosystem Digest
Sends to babaranrob@gmail.com with Gmail label 'Claude'.
"""

import smtplib
import os
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path

# ── Credentials ──────────────────────────────────────────────────────────────
GMAIL_ADDRESS = (
    os.environ.get("GMAIL_ADDRESS")
    or os.environ.get("SMTP_EMAIL")
    or os.environ.get("SENDER_EMAIL")
    or ""
)
GMAIL_PASSWORD = (
    os.environ.get("GMAIL_APP_PASSWORD")
    or os.environ.get("SMTP_PASSWORD")
    or os.environ.get("SENDER_PASSWORD")
    or ""
)
RECIPIENT = "babaranrob@gmail.com"

# ── Digest Content ────────────────────────────────────────────────────────────
SUBJECT = "CLAUDE: Sentinel Weekly Digest — 2026-07-11 [IMPLEMENTED]"
DIGEST_HTML_PATH = Path(__file__).parent / "digest_2026-07-11.html"

with open(DIGEST_HTML_PATH) as f:
    HTML_BODY = f.read()


def send_via_smtp(html: str) -> bool:
    if not GMAIL_ADDRESS or not GMAIL_PASSWORD:
        print("SMTP credentials not found in environment (GMAIL_ADDRESS / GMAIL_APP_PASSWORD).")
        print("Set those env vars and re-run to deliver the email.")
        return False

    msg = MIMEMultipart("alternative")
    msg["Subject"] = SUBJECT
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = RECIPIENT
    msg["X-Gmail-Labels"] = "Claude"
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(GMAIL_ADDRESS, GMAIL_PASSWORD)
            server.send_message(msg)
        print(f"Email sent successfully to {RECIPIENT}")
        return True
    except Exception as e:
        print(f"SMTP error: {e}")
        return False


if __name__ == "__main__":
    print(f"Subject: {SUBJECT}")
    print(f"To:      {RECIPIENT}")
    print()
    sent = send_via_smtp(HTML_BODY)
    if not sent:
        print()
        print("=== EMAIL NOT SENT — credentials missing ===")
        print("Digest content is compiled and ready.")
        print("To send, set GMAIL_ADDRESS and GMAIL_APP_PASSWORD and re-run:")
        print("  python3 src/ii_skills/sentinel/send_digest.py")
        sys.exit(1)
    sys.exit(0)
