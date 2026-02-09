#!/usr/bin/env python3
"""Generate a daily email newsletter from Substack + X sources and metals prices."""

from __future__ import annotations

import argparse
import json
import smtplib
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Iterable
import xml.etree.ElementTree as ET

import requests

try:
    import yfinance as yf
except ImportError:  # pragma: no cover
    yf = None

import sys

CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from settings import DailyMacroMetalsSettings


@dataclass
class Post:
    source: str
    title: str
    url: str
    published: datetime
    summary: str


@dataclass
class MetalSnapshot:
    ticker: str
    name: str
    price: float
    daily_change_pct: float
    weekly_change_pct: float
    monthly_change_pct: float
    yearly_change_pct: float


def _parse_rfc822(value: str | None) -> datetime | None:
    if not value:
        return None
    for fmt in ("%a, %d %b %Y %H:%M:%S %z", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%SZ"):
        try:
            dt = datetime.strptime(value, fmt)
            return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def fetch_rss_posts(feed_url: str, source_label: str, days_back: int = 7) -> list[Post]:
    try:
        resp = requests.get(feed_url, timeout=20)
        resp.raise_for_status()
        root = ET.fromstring(resp.content)
    except Exception as exc:
        print(f"Warning: failed to fetch {source_label} ({feed_url}): {exc}")
        return []
    items = root.findall(".//item")
    cutoff = datetime.now(timezone.utc) - timedelta(days=days_back)
    posts: list[Post] = []

    for item in items:
        title = (item.findtext("title") or "Untitled").strip()
        link = (item.findtext("link") or "").strip()
        summary = (item.findtext("description") or "").strip()
        pub_raw = item.findtext("pubDate")
        pub = _parse_rfc822(pub_raw)
        if not pub:
            continue
        if pub < cutoff:
            continue
        posts.append(Post(source=source_label, title=title, url=link, published=pub, summary=summary))
    return posts


def fetch_x_posts_via_nitter(handle: str, settings: DailyMacroMetalsSettings, days_back: int = 7) -> list[Post]:
    last_error: Exception | None = None
    for instance in settings.nitter_instances:
        try:
            feed_url = f"{instance.rstrip('/')}/{handle}/rss"
            return fetch_rss_posts(feed_url, f"X @{handle}", days_back=days_back)
        except Exception as exc:  # pragma: no cover
            last_error = exc
    if last_error:
        print(f"Warning: could not fetch X posts for @{handle}: {last_error}")
    return []


def fetch_metal_snapshot(ticker: str, name: str) -> MetalSnapshot | None:
    if yf is None:
        print("Warning: yfinance not installed; cannot fetch metals data.")
        return None

    try:
        hist = yf.Ticker(ticker).history(period="1y")
    except Exception as exc:
        print(f"Warning: failed to fetch market data for {ticker}: {exc}")
        return None

    if hist.empty:
        return None

    close = hist["Close"]
    current = float(close.iloc[-1])

    def pct_from_lookback(days: int) -> float:
        idx = max(0, len(close) - days)
        ref = float(close.iloc[idx])
        return ((current - ref) / ref) * 100 if ref else 0.0

    daily = pct_from_lookback(2)
    weekly = pct_from_lookback(6)
    monthly = pct_from_lookback(22)
    yearly = pct_from_lookback(len(close))

    return MetalSnapshot(
        ticker=ticker,
        name=name,
        price=current,
        daily_change_pct=daily,
        weekly_change_pct=weekly,
        monthly_change_pct=monthly,
        yearly_change_pct=yearly,
    )


def summarize_posts(source_name: str, posts: Iterable[Post]) -> str:
    posts = sorted(posts, key=lambda p: p.published, reverse=True)
    if not posts:
        return f"### {source_name}\n- No new updates in the lookback window.\n"

    lines = [f"### {source_name}"]
    for p in posts[:5]:
        lines.append(f"- **{p.title}** ({p.published.strftime('%Y-%m-%d')})")
        if p.url:
            lines.append(f"  - {p.url}")
    return "\n".join(lines) + "\n"


def summarize_metals(snapshots: Iterable[MetalSnapshot | None]) -> str:
    lines = []
    for snap in snapshots:
        if not snap:
            continue
        lines.append(
            f"- **{snap.name}** ({snap.ticker}): ${snap.price:,.2f} | "
            f"1D {snap.daily_change_pct:+.2f}% | 1W {snap.weekly_change_pct:+.2f}% | "
            f"1M {snap.monthly_change_pct:+.2f}% | 1Y {snap.yearly_change_pct:+.2f}%"
        )
    return "\n".join(lines) if lines else "- Metals data unavailable."


def render_markdown(settings: DailyMacroMetalsSettings, source_summary: str, metals_summary: str) -> str:
    template_path = Path(__file__).parent / "templates" / "newsletter-template.md"
    template = template_path.read_text(encoding="utf-8")
    return (
        template.replace("{{DATE}}", datetime.now().strftime("%Y-%m-%d"))
        .replace("{{SOURCE_SUMMARY}}", source_summary)
        .replace("{{METALS_SUMMARY}}", metals_summary)
    )


def send_email(subject: str, markdown_body: str, settings: DailyMacroMetalsSettings, to_email: str) -> None:
    if not settings.gmail_address or not settings.gmail_app_password:
        raise ValueError("GMAIL_ADDRESS and GMAIL_APP_PASSWORD must be set to send email.")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = settings.gmail_address
    msg["To"] = to_email
    msg.attach(MIMEText(markdown_body, "plain"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(settings.gmail_address, settings.gmail_app_password)
        server.sendmail(settings.gmail_address, to_email, msg.as_string())


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate daily macro + metals newsletter")
    parser.add_argument("--days-back", type=int, default=7, help="How far back to check sources")
    parser.add_argument("--send", action="store_true", help="Send via configured Gmail SMTP")
    parser.add_argument("--to", type=str, default="", help="Override recipient email")
    args = parser.parse_args()

    settings = DailyMacroMetalsSettings()
    settings.ensure_dirs()

    substack_posts = fetch_rss_posts(settings.campbell_substack_feed, "Substack @campbellramble", args.days_back)
    crossborder_posts = fetch_x_posts_via_nitter(settings.x_crossbordercap, settings, args.days_back)
    abcampbell_posts = fetch_x_posts_via_nitter(settings.x_abcampbell, settings, args.days_back)

    source_sections = [
        summarize_posts("Substack @campbellramble", substack_posts),
        summarize_posts("X @crossbordercap", crossborder_posts),
        summarize_posts("X @abcampbell", abcampbell_posts),
    ]
    source_summary = "\n".join(source_sections)

    silver = fetch_metal_snapshot("SI=F", "Silver")
    gold = fetch_metal_snapshot("GC=F", "Gold")
    metals_summary = summarize_metals([silver, gold])

    markdown = render_markdown(settings, source_summary, metals_summary)

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_md = settings.base_dir / "archive" / f"newsletter-{stamp}.md"
    out_json = settings.base_dir / "archive" / f"newsletter-{stamp}.json"
    out_md.write_text(markdown, encoding="utf-8")

    structured = {
        "created_at": datetime.now(timezone.utc).isoformat(),
        "substack_posts": [p.__dict__ | {"published": p.published.isoformat()} for p in substack_posts],
        "crossborder_posts": [p.__dict__ | {"published": p.published.isoformat()} for p in crossborder_posts],
        "abcampbell_posts": [p.__dict__ | {"published": p.published.isoformat()} for p in abcampbell_posts],
        "silver": silver.__dict__ if silver else None,
        "gold": gold.__dict__ if gold else None,
    }
    out_json.write_text(json.dumps(structured, indent=2), encoding="utf-8")

    if args.send:
        recipient = args.to or settings.default_recipient
        if not recipient:
            raise ValueError("Recipient email is required via --to or RECIPIENT_EMAIL.")
        send_email("Daily Macro & Metals Newsletter", markdown, settings, recipient)
        print(f"Sent newsletter to {recipient}")

    print(f"Saved markdown: {out_md}")
    print(f"Saved structured data: {out_json}")


if __name__ == "__main__":
    main()
