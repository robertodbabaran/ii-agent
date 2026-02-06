---
name: youtube-transcript-notes
description: Transform a YouTube podcast/interview transcript into high-signal, memorable conversational quotes with context cues. Use when a user provides a video URL or transcript and asks for distilled insights, worldview extraction, quote curation, or socially portable talking points.
---

# YouTube Transcript Notes

Process long-form video transcripts into concise, worldview-level quotes that are easy to recall and use in sophisticated conversation.

## Workflow

1. Ingest source material
2. Map worldview and conceptual territory
3. Extract candidate quotes
4. Filter with strict quality constraints
5. Output final markdown table

## 1) Ingest source material

- Accept one of:
  - YouTube URL
  - Full transcript text
  - Notes or transcript fragments
- If only a URL is provided, obtain transcript text before synthesis.
- Treat transcript accuracy as first-order. Avoid adding ideas not grounded in the source.

## 2) Map worldview before selecting quotes

Identify the speaker's implicit model of:

- Reality
- Human nature
- Power and incentives
- Time horizons
- Systems behavior under pressure
- Meaning and tradeoffs

Write a short internal map first, then choose quotes that represent the map.

## 3) Generate candidate quote pool

- Pull 20-30 candidate lines from the transcript.
- Keep spoken cadence; avoid formal prose.
- Allow light editing only for clarity and compression.
- Do not change core meaning, voice, or intent.

## 4) Filter candidates with hard constraints

Every final quote must satisfy all of these:

- 30 words maximum
- Standalone without external context
- Worldview-revealing (not generic advice)
- Memorable and verbally natural
- No self-help clichés, corporate language, or therapeutic framing
- Suitable for both professional and intimate social contexts

Reject quotes that are:

- Platitudes
- Anecdotes without transferable insight
- Mostly rhetorical or comedic
- Context-dependent to the point of ambiguity
- Impressionistic but low-information

## 5) Build context cues

For each selected quote, add one context cue that is 8-15 words and states:

- Conceptual territory
- Situation or question addressed
- Angle/lens for interpretation

Do not mention the speaker name in cues.

## 6) Output format

Return a markdown table with exactly two columns:

| Context Cue | Quote |
|---|---|
| ... | ... |

Default output count: 10 quotes.

If user instructions conflict on quote count, prioritize the most explicit final formatting instruction in the active request.

## Tone constraints

- No emojis
- No hype language
- No moralizing
- No "wisdom-signaling"
- Minimal jargon unless required by the source
- Write for an already-intelligent audience

## Reusable prompt template

Use the template in `references/prompt-template.md` when you need a ready-to-run instruction block for quote extraction.
