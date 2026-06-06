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
SUBJECT = "CLAUDE: Sentinel Weekly Digest — 2026-06-06 [IMPLEMENTED]"

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
  .tag-mcp       { background: #0d1a20; color: #5bc4d8; }
  .tag-billing   { background: #1a0d0d; color: #ef8f8f; }

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
  .warning-pill {
    display: inline-block;
    font-size: 10px;
    background: #6b1a1a;
    color: #ffaaaa;
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
    <div class="date">Week of June 6, 2026 &nbsp;·&nbsp; ii-agent edition</div>
  </div>

  <!-- Top 5 Headlines -->
  <div class="section">
    <div class="section-title">Top 5 Headlines</div>

    <div class="headline">
      <div class="headline-num">01 &nbsp;<span class="tag tag-anthropic">ANTHROPIC</span></div>
      <div class="headline-title">Claude Opus 4.8 + Dynamic Workflows land — up to 1,000 parallel subagents in Claude Code</div>
      <div class="headline-body">
        Released May 28, <strong>Opus 4.8</strong> scores 69.2% on SWE-Bench Pro (beating GPT-5.5 and Gemini 3.1 Pro) and is 4× less likely to silently pass code flaws. The headline feature is <strong>Dynamic Workflows</strong> in Claude Code: turn on the new <code>ultracode</code> setting (effort = xhigh, auto-workflow) and Claude spins up a harness, breaks tasks into parallel subtasks, dispatches up to 1,000 subagents, validates results, and returns a synthesized answer — no manual orchestration needed. Early testers have used it for codebase-wide bug hunts, security audits across thousands of files, and large-scale migrations. Fast mode on Opus 4.8 is now 3× cheaper than the previous generation.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">02 &nbsp;<span class="tag tag-billing">BILLING</span></div>
      <div class="headline-title">Anthropic billing split goes live June 15 — Agent SDK gets its own $20–$200/mo credit pool</div>
      <div class="headline-body">
        Starting June 15, automated Claude Agent SDK usage (<code>claude -p</code>, Claude Code GitHub Actions, third-party apps authenticating via Agent SDK) will draw from a <strong>separate monthly credit</strong> — $20 for Pro, $100 for Max 5×, $200 for Max 20× — billed at full API list rates. Interactive use (chat, Claude Code in terminal, Cowork) is unaffected. When the credit is exhausted, automated requests stop with no fallback unless you manually enable overflow billing. This is Anthropic's third attempt to fix the structural economics problem created by flat-rate subscriptions and compute-hungry coding agents. If you run scheduled newsletter jobs or CI pipelines against Claude, audit now.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">03 &nbsp;<span class="tag tag-mcp">MCP</span></div>
      <div class="headline-title">MCP 2026-07-28 Release Candidate locks — stateless protocol, sessions eliminated</div>
      <div class="headline-body">
        The largest revision of MCP since launch is locked as of May 21. The RC removes the <code>initialize</code> handshake and <code>Mcp-Session-Id</code> header entirely — every request is now self-contained, so servers run behind plain round-robin load balancers with no sticky sessions or shared session stores. New routable headers (<code>Mcp-Method</code>, <code>Mcp-Name</code>) enable gateway routing without body inspection. The Extensions framework adds <strong>Tasks</strong> (long-running work via explicit handles) and <strong>MCP Apps</strong> (server-rendered HTML UIs). OAuth hardening aligns with RFC 9207 for production deployments. Final spec ships July 28 — breaking changes require migration if your MCP servers use sessions, initialize, or the old Tasks API.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">04 &nbsp;<span class="tag tag-aws">AWS</span></div>
      <div class="headline-title">Amazon Bedrock AgentCore Payments — agents now carry Coinbase/Stripe wallets mid-execution</div>
      <div class="headline-body">
        Amazon Bedrock AgentCore's April preview of agent payments went deeper in May. Developers connect a <strong>Coinbase CDP or Stripe Privy wallet</strong>, set session-level spending limits, and agents transact autonomously via the <strong>x402 protocol</strong>: when the agent hits an HTTP 402 response from a paid resource (API, MCP server, paywalled content), AgentCore negotiates the payment, authenticates, executes a stablecoin micropayment, and continues without interrupting the reasoning loop. First version covers micropayments for data feeds and APIs; roadmap includes larger transactions (hotel bookings, travel). Architecturally significant — this is the first managed payment primitive purpose-built for autonomous agents at production scale.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">05 &nbsp;<span class="tag tag-github">GITHUB</span></div>
      <div class="headline-title">OpenClaw reaches 346k stars — the "self-writing skills" pattern goes mainstream</div>
      <div class="headline-body">
        OpenClaw (the open-source personal AI agent by Peter Steinberger) has surpassed <strong>346,000 GitHub stars</strong> and 3.2M active users. Its defining capability: agents that <em>write their own new skills</em> — when it can't do something, it generates and installs the code to do it. The community skills registry now has 100+ integrations (Gmail, GitHub, Notion, Obsidian, Spotify, Home Assistant). The self-extending pattern — autonomous capability expansion without manual configuration — is the most watched architectural idea in the open-source agent space right now.
      </div>
    </div>
  </div>

  <!-- Architecture Watch -->
  <div class="section">
    <div class="section-title">Architecture Watch</div>

    <div class="item">
      <div class="item-title">MCP going stateless — infrastructure implications are large</div>
      <div class="item-body">
        The 2026-07-28 RC means remote MCP servers no longer need sticky sessions, shared session stores, or deep packet inspection at the gateway. Application state moves to explicit handles (a <code>basket_id</code>, <code>browser_id</code>) that the model passes back as arguments — visible, composable, and auditable. If ii-agent's skills ever expose MCP servers publicly, this is the architecture to target. Migration window is 10 weeks from May 21.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Graph-based execution is the consensus pattern for 2026</div>
      <div class="item-body">
        Claude Code Dynamic Workflows, LangGraph v0.4, Google ADK 1.0, and AutoGen 1.0 GA all converge on directed-graph execution. Benefits over hierarchical tree delegation: cleaner fan-out/fan-in, explicit retry semantics, auditable state at every node, and human-in-the-loop gates that surface automatically rather than requiring custom plumbing. LangGraph v0.4 (April 2026) ships these as first-class primitives — <code>Interrupt</code> objects appear in <code>.invoke()</code> return values automatically, and every action is persisted to a checkpointer (Postgres-compatible) for compliance.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Context inconsistency — not pattern choice — is why multi-agent pilots fail <span class="ii-pill">ii-agent relevant</span></div>
      <div class="item-body">
        40% of multi-agent production pilots fail within 6 months. The primary failure mode: sub-agents diverge because they don't share a coherent view of state. Claude Code's Dynamic Workflows solution is a shared filesystem; LangGraph's is a shared checkpointer. For ii-agent: the <code>deal_memory</code> skill is the right primitive — auto-injecting the active deal's memory snapshot before any IB Toolkit dispatch (a 3-line hook in the dispatcher) directly addresses this failure mode.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Context Engineering replaces Prompt Engineering as the primary discipline</div>
      <div class="item-body">
        The framing shift of 2026: the craft is no longer <em>how you write a prompt</em> but <em>what you put in the context window and when</em>. Memory retrieval, tool pre-loading, selective injection, and session summarization are the new levers. Simon Willison is documenting these under the name "Agentic Engineering Patterns" — a crowd-sourced taxonomy worth tracking for ii-agent's own pattern library.
      </div>
    </div>
  </div>

  <!-- Capability Spotlight -->
  <div class="section">
    <div class="section-title">Capability Spotlight</div>

    <div class="item">
      <div class="item-title"><code>ultracode</code> — one toggle for 1,000-subagent orchestration <span class="ii-pill">use now</span></div>
      <div class="item-body">
        Switch on <code>ultracode</code> in Claude Code's effort menu to combine xhigh reasoning with automatic workflow orchestration. Claude plans a harness for each substantive task, breaks it into parallel subtasks, runs up to 1,000 subagents, and validates before returning. No orchestration code required. Ideal for ii-agent's most complex prompts: comprehensive LBO builds, full IC packages, multi-sector rollup models.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Security-guidance plugin — real-time vulnerability review on every edit <span class="ii-pill">evaluate</span></div>
      <div class="item-body">
        Anthropic's new security-guidance plugin runs a fast pattern check on each edit, a model review at the end of each turn, and a deeper agentic review on commit or push — all within the same Claude Code session. Directly relevant if ii-agent's Python skill code ever handles external API data, CIM PDFs, or user-provided inputs.
      </div>
    </div>

    <div class="item">
      <div class="item-title">fastmcp-remote (June 2) — stdio-only MCP hosts now reach HTTP servers</div>
      <div class="item-body">
        Released June 2, <code>fastmcp-remote</code> bridges stdio-only MCP hosts (like Claude Code) to servers hosted over HTTP, with OAuth enabled automatically for HTTPS. A single URL in, local stdio proxy out. Relevant if you want to host ii-agent skills as MCP servers behind a private HTTP endpoint without client-side transport rewrites.
      </div>
    </div>

    <div class="item">
      <div class="item-title">GitHub MCP: <code>search_commits</code> tool added</div>
      <div class="item-body">
        The GitHub MCP server added <code>search_commits</code> — search commits across repos using GitHub's full search syntax. Also added <code>set_issue_fields</code> for custom org-level issue field values. If ii-agent workflows reference GitHub (deal tracking, version history), these are now available as native MCP tools without custom integration code.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Claude Cowork: legal MCP connectors + enterprise embedding <span class="ii-pill">watch</span></div>
      <div class="item-body">
        Cowork added 20+ legal MCP connectors and 12 practice-area plugins (research, contracts, discovery, matter management). KPMG embedded Managed Agents inside Cowork for their 276,000-person workforce. The broader pattern: Claude as a runtime, not just a chatbot — Managed Agents handle orchestration, Cowork handles the UX layer.
      </div>
    </div>
  </div>

  <!-- Workflow Ideas -->
  <div class="section">
    <div class="section-title">Workflow Ideas</div>

    <div class="item">
      <div class="item-title">Audit your automated Claude calls before June 15 <span class="warning-pill">action required</span></div>
      <div class="item-body">
        The networth_newsletter, market_newsletter, and health_dashboard runners that call Claude via the API will now draw from the separate Agent SDK credit pool starting June 15. Check each script: if it authenticates via the Agent SDK (not interactive session), it counts against the $20–$200/month credit. Either enable overflow billing, restructure as interactive sessions, or budget the new pool. The cutoff is June 15 — no automatic fallback.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Turn on <code>ultracode</code> for the next comprehensive LBO build <span class="ii-pill">build this</span></div>
      <div class="item-body">
        The next time someone asks for a "7-day model" or "full IC package," add the ultracode toggle. Claude will automatically dispatch parallel subagents for each module (Sources &amp; Uses, Debt Schedule, Returns, Sensitivity) rather than running them sequentially. Compare quality and time to the standard approach — one run is enough to calibrate.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Self-writing skills pattern — apply to ii-agent <span class="ii-pill">design</span></div>
      <div class="item-body">
        OpenClaw's viral growth is built on one idea: when the agent can't do something, it writes the skill to do it and installs it. ii-agent has a formal skill registry and bridge layer already. The missing piece: a prompt that detects "no skill covers this request" and drafts a new skill module, scaffolded to the existing Tier 3 workspace pattern, for human review before installation.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Name your own agentic patterns — Willison-style documentation</div>
      <div class="item-body">
        ii-agent has accumulated real patterns: the "Tier 3 workspace → Tier 2 template read" pattern, the "Excel formula → PPTX mirror" pair, the "case intake → module routing" flow. Writing these up with names and rationale (a 2-page <code>PATTERNS.md</code>) pays compound interest every time a new skill is designed, and positions ii-agent's architecture as a reference system.
      </div>
    </div>
  </div>

  <!-- Worth Reading -->
  <div class="section">
    <div class="section-title">Worth Reading</div>

    <div class="read-item">
      <div class="read-link"><a href="https://blog.modelcontextprotocol.io/posts/2026-07-28-release-candidate/">MCP 2026-07-28 Release Candidate — Official MCP Blog</a></div>
      <div class="read-why">The authoritative post on the stateless spec. Read the "Breaking Changes" section before July if you have any MCP server dependencies — the initialize handshake removal is a real migration effort.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://simonw.substack.com/p/agentic-engineering-patterns">Agentic Engineering Patterns — Simon Willison (Substack)</a></div>
      <div class="read-why">Willison is crowd-sourcing a taxonomy of patterns for building with coding agents — directly useful for documenting ii-agent's own architecture. He's actively collecting submissions.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://codersera.com/blog/anthropic-june-2026-billing-change-claude-code/">Anthropic's June 15 Billing Change — Codersera</a></div>
      <div class="read-why">The clearest breakdown of what changes, what doesn't, and what to do before June 15. Includes a math section showing 12×–175× effective price increase by workload type — essential reading before the cutoff.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://techcrunch.com/2026/05/28/anthropic-releases-opus-4-8-with-new-dynamic-workflow-tool/">Anthropic releases Opus 4.8 with Dynamic Workflows — TechCrunch</a></div>
      <div class="read-why">Good overview of the Opus 4.8 release covering benchmark improvements, the dynamic workflows feature, and the $65B funding round context. Useful for calibrating whether to upgrade default model in ii-agent prompts.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://lilianweng.github.io/posts/2025-05-01-thinking/">Why We Think — Lilian Weng (Lil'Log)</a></div>
      <div class="read-why">Weng's deep dive on test-time compute, chain-of-thought reasoning, and latent thoughts — explains why effort-level settings like ultracode actually work at the model level. Dense but worth the read for architectural intuitions.</div>
    </div>
  </div>

  <!-- Recommendation -->
  <div class="section">
    <div class="rec-box">
      <div class="rec-label">⚡ This Week's Recommendation</div>
      <div class="rec-title">Audit Agent SDK usage before the June 15 billing split</div>
      <div class="rec-body">
        Anthropic's June 15 billing change is the most operationally urgent item this week. Any script that calls the Claude API programmatically (newsletter runners, CI pipelines, automated skill invocations) may now count against a separate $20–$200/month credit pool with no automatic fallback.<br /><br />
        <strong>Concrete steps:</strong><br />
        1. Open each newsletter/dashboard runner (<code>networth.py</code>, <code>newsletter.py</code>, <code>whoop_newsletter.py</code>) and check whether they authenticate via the Agent SDK or a direct API key call.<br />
        2. Log into your Anthropic account and check the new credit pool allocation for your plan tier.<br />
        3. Enable overflow billing if you want uninterrupted automated runs, or set a usage alert at 80% of the credit pool.<br />
        4. If any runner is close to the credit ceiling, consider restructuring it to call a lower-cost model (Haiku 4.5) for data aggregation and only use Opus 4.8 for the final synthesis step.<br /><br />
        The cost of missing this is automated jobs silently stopping on June 15 with no fallback — a low-visibility, high-impact failure.
      </div>
    </div>
  </div>

  <!-- Sign-off -->
  <div class="sig">— Jordan</div>

  <!-- Footer -->
  <div class="footer">
    SENTINEL monitors Anthropic/Claude, agent frameworks, GitHub trending, and notable content weekly.<br/>
    ii-agent · robertodbabaran/ii-agent · Digest generated 2026-06-06
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
