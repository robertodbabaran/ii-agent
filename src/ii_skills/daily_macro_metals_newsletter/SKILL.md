---
name: daily-macro-metals-newsletter
description: "Generate a daily email newsletter that summarizes updates from Campbell Ramble Substack, X accounts @crossbordercap and @abcampbell, and includes silver/gold spot price with daily/weekly/monthly/yearly changes."
---

# Daily Macro Metals Newsletter Skill

## Overview
This skill generates a daily newsletter with:
- New posts (if any) from:
  - `https://substack.com/@campbellramble` (via RSS feed)
  - `https://x.com/crossbordercap`
  - `https://x.com/abcampbell`
- Precious metals snapshot:
  - Current **silver** and **gold** prices
  - **1D / 1W / 1M / 1Y** percentage changes

## Why this design
- Substack offers a reliable RSS feed.
- X does not provide native public RSS, so this skill attempts Nitter RSS mirrors.
- If X sources are unavailable on a given day, the newsletter marks them as "No new updates" rather than failing.

## Usage
Generate and archive newsletter only:

```bash
python src/ii_skills/daily_macro_metals_newsletter/generate_newsletter.py
```

Generate and send email:

```bash
python src/ii_skills/daily_macro_metals_newsletter/generate_newsletter.py --send --to you@example.com
```

## Configuration
Environment variables for email sending:
- `GMAIL_ADDRESS`
- `GMAIL_APP_PASSWORD`
- `RECIPIENT_EMAIL` (optional override; defaults to `user@example.com`)

## Output
Artifacts are written to:
- `Memory/daily_macro_metals_newsletter/archive/newsletter-<timestamp>.md`
- `Memory/daily_macro_metals_newsletter/archive/newsletter-<timestamp>.json`

## Scheduling
Run daily with cron (example at 6:30 AM):

```cron
30 6 * * * cd /workspace/ii-agent && /usr/bin/python src/ii_skills/daily_macro_metals_newsletter/generate_newsletter.py --send
```
