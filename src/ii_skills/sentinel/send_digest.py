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
from datetime import datetime

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
SUBJECT = "CLAUDE: Sentinel Weekly Digest — 2026-05-23 [IMPLEMENTED]"

HTML_BODY = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Sentinel Weekly Digest</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    background: #0f0f13;
    color: #e4e4ef;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif;
    font-size: 15px;
    line-height: 1.65;
  }
  .wrapper { max-width: 680px; margin: 0 auto; padding: 32px 20px; }

  /* Header */
  .header {
    border-bottom: 1px solid #2a2a38;
    padding-bottom: 24px;
    margin-bottom: 32px;
  }
  .header-label {
    font-size: 11px;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #6b6b80;
    margin-bottom: 8px;
  }
  .header h1 {
    font-size: 26px;
    font-weight: 700;
    color: #ffffff;
    letter-spacing: -0.5px;
  }
  .header .date {
    font-size: 13px;
    color: #6b6b80;
    margin-top: 6px;
  }

  /* Section */
  .section { margin-bottom: 36px; }
  .section-title {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #7c6af5;
    border-left: 3px solid #7c6af5;
    padding-left: 10px;
    margin-bottom: 16px;
  }

  /* Headline cards */
  .headline {
    background: #16161f;
    border: 1px solid #22222e;
    border-radius: 8px;
    padding: 16px 18px;
    margin-bottom: 10px;
  }
  .headline-num {
    font-size: 10px;
    color: #4a4a5a;
    font-weight: 600;
    letter-spacing: 1px;
    margin-bottom: 4px;
  }
  .headline-title {
    font-size: 15px;
    font-weight: 600;
    color: #f0f0ff;
    margin-bottom: 6px;
  }
  .headline-body {
    font-size: 14px;
    color: #9090a8;
    line-height: 1.6;
  }
  .tag {
    display: inline-block;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    padding: 2px 7px;
    border-radius: 4px;
    margin-right: 6px;
    vertical-align: middle;
  }
  .tag-anthropic { background: #1e1a38; color: #a98bfa; }
  .tag-google    { background: #0d2010; color: #5fbf7a; }
  .tag-aws       { background: #1a1000; color: #f5a623; }
  .tag-github    { background: #0d1117; color: #8b949e; }
  .tag-memory    { background: #1a0d1a; color: #d88aee; }

  /* List items */
  .item {
    padding: 12px 0;
    border-bottom: 1px solid #1c1c28;
  }
  .item:last-child { border-bottom: none; }
  .item-title {
    font-size: 14px;
    font-weight: 600;
    color: #e0e0f0;
    margin-bottom: 4px;
  }
  .item-body {
    font-size: 13.5px;
    color: #8888a0;
    line-height: 1.55;
  }
  .ii-pill {
    display: inline-block;
    font-size: 10px;
    background: #7c6af5;
    color: #fff;
    padding: 1px 6px;
    border-radius: 3px;
    margin-left: 6px;
    font-weight: 600;
    letter-spacing: 0.5px;
    vertical-align: middle;
  }

  /* Links */
  a { color: #7c6af5; text-decoration: none; }
  a:hover { text-decoration: underline; }
  .read-item { padding: 10px 0; border-bottom: 1px solid #1c1c28; }
  .read-item:last-child { border-bottom: none; }
  .read-link { font-size: 14px; font-weight: 500; color: #7c6af5; }
  .read-why { font-size: 13px; color: #6b6b80; margin-top: 2px; }

  /* Recommendation box */
  .rec-box {
    background: linear-gradient(135deg, #18103a 0%, #10181a 100%);
    border: 1px solid #3a2a6a;
    border-radius: 10px;
    padding: 20px 22px;
  }
  .rec-box .rec-label {
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #a98bfa;
    margin-bottom: 8px;
  }
  .rec-box .rec-title {
    font-size: 16px;
    font-weight: 700;
    color: #ffffff;
    margin-bottom: 10px;
  }
  .rec-box .rec-body {
    font-size: 14px;
    color: #b0b0c8;
    line-height: 1.65;
  }

  /* Footer */
  .footer {
    margin-top: 40px;
    padding-top: 24px;
    border-top: 1px solid #1c1c28;
    font-size: 12px;
    color: #4a4a5a;
    text-align: center;
    line-height: 1.6;
  }
  .sig { font-size: 14px; color: #6b6b80; margin-top: 28px; font-style: italic; }
</style>
</head>
<body>
<div class="wrapper">

  <!-- Header -->
  <div class="header">
    <div class="header-label">SENTINEL &nbsp;/&nbsp; AI Agent Ecosystem Monitor</div>
    <h1>Weekly Digest</h1>
    <div class="date">Week of May 23, 2026 &nbsp;·&nbsp; ii-agent edition</div>
  </div>

  <!-- Top 5 Headlines -->
  <div class="section">
    <div class="section-title">Top 5 Headlines</div>

    <div class="headline">
      <div class="headline-num">01 &nbsp;<span class="tag tag-anthropic">ANTHROPIC</span></div>
      <div class="headline-title">Anthropic ships 5 new Managed Agents features at "Code with Claude 2026"</div>
      <div class="headline-body">
        In two drops (May 7 &amp; May 19), Anthropic shipped: <strong>Dreaming</strong> — a scheduled background process that reviews past sessions, extracts patterns, and rewrites agent memory for self-improvement; <strong>multiagent orchestration</strong> — a lead agent breaks jobs into pieces and delegates to specialist sub-agents running in parallel on a shared filesystem; <strong>MCP tunnels</strong> — route private MCP servers through an outbound-only channel so agents can reach internal services without public exposure; plus <strong>secure sandboxes</strong> and <strong>self-hosted options</strong> for enterprise compliance. The most consequential Anthropic agent infrastructure week to date.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">02 &nbsp;<span class="tag tag-google">GOOGLE</span></div>
      <div class="headline-title">Google ADK 2.0 goes GA — graph-based Workflow Runtime replaces hierarchical executor</div>
      <div class="headline-body">
        Released May 19, ADK 2.0 is a breaking change that transitions the framework from a hierarchical agent executor to a <strong>graph-based Workflow Runtime</strong> — supporting routing, fan-out/fan-in, loops, retry, state management, dynamic nodes, and human-in-the-loop as first-class primitives. The new <strong>Task API</strong> provides structured agent-to-agent delegation. Sessions from ADK 2.0 are readable by 1.28+ but incompatible with older versions.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">03 &nbsp;<span class="tag tag-aws">AWS</span></div>
      <div class="headline-title">Amazon Bedrock AgentCore Payments preview — agents now carry wallets</div>
      <div class="headline-body">
        Built with Coinbase and Stripe, AgentCore Payments lets autonomous agents pay for API calls, MCP servers, web content, and other agents mid-execution. Developers connect a CDP or Stripe Privy wallet, set session-level spending limits, and the agent transacts without interruption. The first managed payment primitive purpose-built for autonomous agents — early but architecturally significant.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">04 &nbsp;<span class="tag tag-github">GITHUB</span></div>
      <div class="headline-title">OpenHuman tops GitHub Trending — "reads you first, then answers"</div>
      <div class="headline-body">
        Launched May 12 by tinyhumansai, <strong>OpenHuman</strong> reached #1 on GitHub Trending for 7 consecutive days, crossing 20k+ stars. Its core innovation: a local-first memory that continuously ingests 118+ connected services (emails, repos, calendars, chats) so every conversation starts with full context — no onboarding prompts needed. Rivals OpenClaw (210k stars, skill self-writing) and Hermes Agent (153k stars, Nous Research, closed-loop skill rewriting) on a different axis: <em>depth of personal context</em> vs. breadth of capability.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">05 &nbsp;<span class="tag tag-memory">MEMORY</span></div>
      <div class="headline-title">Agent memory matures from experiment to infrastructure — but accuracy is still only 27.9%</div>
      <div class="headline-body">
        Cloudflare launched a <strong>managed agent memory service</strong> (April); Mem0 released a <strong>token-efficient memory algorithm</strong> using single-pass hierarchical extraction and multi-signal retrieval; and the new <strong>MINTEval benchmark</strong> (May 2026) stress-tested 7 memory systems across contexts averaging 138.8k tokens — average accuracy: 27.9%. The gap between memory-as-feature and memory-as-reliable-infrastructure remains wide. Meanwhile, memory costs 10-20× less than context-only approaches at $0.05–0.15/call vs. $0.50+ for million-token windows.
      </div>
    </div>
  </div>

  <!-- Architecture Watch -->
  <div class="section">
    <div class="section-title">Architecture Watch</div>

    <div class="item">
      <div class="item-title">Context Engineering replaces Prompt Engineering as the primary discipline</div>
      <div class="item-body">
        The framing shift of 2026: the craft is no longer <em>how you write a prompt</em> but <em>what you put in the context window and when</em>. Memory retrieval, tool pre-loading, session summarization, and selective injection are the new levers. Simon Willison is documenting these as "Agentic Engineering Patterns" — worth tracking.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Graph-based execution is winning over hierarchical trees</div>
      <div class="item-body">
        Google ADK 2.0, LangGraph v0.4 (40% of production deployments), and Anthropic's own multiagent orchestration all converge on directed-graph execution. The benefits over tree-based delegation: cleaner fan-out/fan-in, explicit retry semantics, and auditable state at every node — critical for finance workflows.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Context inconsistency — not pattern choice — is why multi-agent pilots fail</div>
      <div class="item-body">
        57% of organizations run multi-step agent workflows in production; 40% of pilots fail within 6 months. The primary failure mode identified in April 2026 research: sub-agents diverge because they don't share a coherent view of state. Implication: shared filesystem / shared memory is not optional — it's the foundation.
      </div>
    </div>

    <div class="item">
      <div class="item-title">OpenHuman's "pre-load before prompt" pattern <span class="ii-pill">ii-agent relevant</span></div>
      <div class="item-body">
        Rather than asking the user for context, OpenHuman continuously indexes the user's digital life locally and injects relevant context automatically. This is the logical evolution of the <code>deal_memory</code> skill — instead of explicitly calling "remember this," the system could passively track deal-relevant signals across sessions.
      </div>
    </div>
  </div>

  <!-- Capability Spotlight -->
  <div class="section">
    <div class="section-title">Capability Spotlight</div>

    <div class="item">
      <div class="item-title">Claude Code desktop redesign: parallel sessions + SSH on Mac <span class="ii-pill">use now</span></div>
      <div class="item-body">
        New sidebar filters sessions by status/project/environment; integrated terminal + file editor; faster diff viewer; SSH remote sessions now supported on Mac. Direct upgrade for ii-agent development workflows — run the IB Toolkit and IR Toolkit simultaneously in separate sessions.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Cache Diagnostics (public beta) — understand prompt cache misses <span class="ii-pill">evaluate</span></div>
      <div class="item-body">
        Pass a <code>diagnostics</code> param on any Messages API request to get a structured explanation of why a cache miss occurred. Given ii-agent's repeated template-heavy calls (LBO models, slide generation), this could reveal concrete savings. Enable in dev, instrument 20 calls, measure.
      </div>
    </div>

    <div class="item">
      <div class="item-title">MCP tunnels for private-network MCP servers</div>
      <div class="item-body">
        Anthropic's MCP tunnel feature lets Claude Managed Agents reach MCP servers behind a firewall via outbound-only encrypted channels — no public exposure. Relevant if ii-agent skills ever need to hit internal databases or private APIs without cloud exposure.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Claude Connectors: Gmail, Slack, Google Workspace, Microsoft 365</div>
      <div class="item-body">
        Claude can now natively connect to Gmail, Slack, meeting notes, HubSpot, DocuSign, and Microsoft 365 tools (Outlook in public beta; Excel/PowerPoint/Word GA). The market_newsletter, networth_newsletter, and health_dashboard skills could potentially be simplified by delegating delivery to Claude Connectors rather than maintaining SMTP infrastructure.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Mem0 token-efficient memory algorithm — 10–20× cheaper than context stuffing</div>
      <div class="item-body">
        Single-pass hierarchical extraction + multi-signal retrieval brings per-call cost to $0.05–0.15 vs. $0.50+ for million-token context windows. The <code>deal_memory</code> skill's current approach could be benchmarked against Mem0's algorithm — worth a spike.
      </div>
    </div>
  </div>

  <!-- Workflow Ideas -->
  <div class="section">
    <div class="section-title">Workflow Ideas</div>

    <div class="item">
      <div class="item-title">"Dreaming" nightly for deal_memory <span class="ii-pill">build this</span></div>
      <div class="item-body">
        Schedule a nightly job that reads the last 7 days of deal_memory sessions, extracts recurring patterns (modelling preferences, deal assumptions, analyst flags), condenses stale entries, and promotes load-bearing ones. Mirrors Anthropic's Dreaming implementation — small prompt + structured output, ~$0.10/run.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Auto-inject deal context before every IB Toolkit invocation</div>
      <div class="item-body">
        Before any LBO or IR module runs, pull the active deal's memory snapshot and prepend it silently. The user never has to say "remember this" — the toolkit already knows. Requires a 3-line hook in the IB Toolkit dispatcher.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Name and document your own agentic patterns (Willison-style)</div>
      <div class="item-body">
        ii-agent has accumulated real patterns across 84+ modules: the "Tier 3 workspace → Tier 2 template read" pattern, the "Excel formula → PPTX mirror" pair, the "case intake → module routing" flow. Giving these names and writing them up pays compound interest every time a new skill is designed.
      </div>
    </div>
  </div>

  <!-- Worth Reading -->
  <div class="section">
    <div class="section-title">Worth Reading</div>

    <div class="read-item">
      <div class="read-link"><a href="https://simonw.substack.com/p/agentic-engineering-patterns">Agentic Engineering Patterns — Simon Willison (Substack)</a></div>
      <div class="read-why">Willison is crowd-sourcing a taxonomy of patterns for building with coding agents like Claude Code and Codex — directly useful for documenting ii-agent's own patterns.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://ruben.substack.com/p/prompt-47">Prompt 4.7 — Ruben Hassid, How to AI</a></div>
      <div class="read-why">Specific guidance on how prompting Claude Opus 4.7 differs from previous models — "drastically different," per Hassid. Relevant before next IB Toolkit prompt update.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://ruben.substack.com/p/claude-connectors">Connect. — Ruben Hassid, How to AI</a></div>
      <div class="read-why">Walkthrough of Claude Connectors for Gmail, Slack, and meeting notes — the simplest path to integrating ii-agent with live data sources without custom code.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://mem0.ai/blog/state-of-ai-agent-memory-2026">State of AI Agent Memory 2026 — mem0.ai</a></div>
      <div class="read-why">Comprehensive benchmarks, architecture patterns, and cost analysis. The 27.9% MINTEval accuracy finding alone is worth reading — it resets expectations about what memory systems can actually guarantee today.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://www.latent.space/podcast">Agent-Native Infrastructure — Latent Space Podcast</a></div>
      <div class="read-why">Episode covering why agents need real OS machines (not just API calls), why CLI matters more than MCP in some contexts, and how the AI infrastructure layer is evolving toward payment/auth primitives (Stripe-like) rather than compute (AWS-like).</div>
    </div>
  </div>

  <!-- Recommendation -->
  <div class="section">
    <div class="rec-box">
      <div class="rec-label">⚡ This Week's Recommendation</div>
      <div class="rec-title">Enable Cache Diagnostics on ii-agent's Claude API calls</div>
      <div class="rec-body">
        Anthropic's Cache Diagnostics is now in public beta. Pass <code>"diagnostics": true</code> on any Messages API request and you'll receive a structured JSON explanation of each prompt cache miss — whether it was a content change, a model mismatch, or a TTL expiry.<br /><br />
        ii-agent's IB Toolkit and IR Toolkit make many repeated API calls with large, near-identical system prompts (module templates, sector playbooks, Excel conventions). Even a 20% cache hit improvement across 100 module invocations/week translates to ~$15–30/month in savings and meaningfully lower latency. The diagnostic cost is zero — it's a parameter flag.<br /><br />
        <strong>Action:</strong> Add <code>diagnostics=True</code> to one batch of LBO module calls in dev, review the miss breakdown, then restructure the static portions of system prompts to be cache-stable.
      </div>
    </div>
  </div>

  <!-- Sign-off -->
  <div class="sig">— Jordan</div>

  <!-- Footer -->
  <div class="footer">
    SENTINEL monitors Anthropic/Claude, agent frameworks, GitHub trending, and notable content weekly.<br/>
    ii-agent · robertodbabaran/ii-agent · Digest generated 2026-05-23
  </div>

</div>
</body>
</html>
"""


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
