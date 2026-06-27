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
SUBJECT = "CLAUDE: Sentinel Weekly Digest — 2026-06-27 [IMPLEMENTED]"

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
  .tag-openai    { background: #0d1a14; color: #10a37f; }
  .tag-github    { background: #0d1117; color: #8b949e; }
  .tag-mcp       { background: #0d1a20; color: #5bc4d8; }
  .tag-urgent    { background: #1a0d0d; color: #ef8f8f; }
  .tag-microsoft { background: #0d1020; color: #7b9ff5; }

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
  .new-pill {
    display: inline-block;
    font-size: 10px;
    background: #0e2a1a;
    color: #5fbf7a;
    padding: 1px 6px;
    border-radius: 3px;
    margin-left: 6px;
    font-weight: 600;
    letter-spacing: 0.5px;
    vertical-align: middle;
  }
  .action-pill {
    display: inline-block;
    font-size: 10px;
    background: #1a1620;
    color: #c4a0f5;
    padding: 1px 6px;
    border-radius: 3px;
    margin-left: 6px;
    font-weight: 600;
    letter-spacing: 0.5px;
    vertical-align: middle;
  }

  /* Alert banner */
  .alert-banner {
    background: #1a0d10;
    border: 1px solid #5a1a20;
    border-left: 4px solid #ef4444;
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 24px;
  }
  .alert-banner .alert-label {
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: #ef4444;
    margin-bottom: 4px;
  }
  .alert-banner .alert-text {
    font-size: 14px;
    color: #e0b0b0;
    line-height: 1.5;
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
    <div class="date">Week of June 20–27, 2026 &nbsp;·&nbsp; ii-agent edition</div>
  </div>

  <!-- Top 5 Headlines -->
  <div class="section">
    <div class="section-title">Top 5 Headlines</div>

    <div class="headline">
      <div class="headline-num">01 &nbsp;<span class="tag tag-anthropic">ANTHROPIC</span></div>
      <div class="headline-title">Claude Tag lands in Slack (June 23) — one persistent AI teammate per channel, ambient mode included</div>
      <div class="headline-body">
        Anthropic launched <strong>Claude Tag</strong> in public beta for Enterprise and Team customers on June 23. Unlike per-user Claude integrations, Tag is a single Claude identity per channel: every team member sees the same context, can hand off half-finished tasks, and can reference prior conversations the AI had with colleagues. When assigned a task, Tag breaks it into stages and works through them autonomously, posting updates in a thread. <strong>Ambient mode</strong> is the sleeper feature: Claude proactively jumps into conversations unprompted — flagging forgotten threads, surfacing relevant context from across the organisation, and nudging on stalled tasks. Anthropic's framing: "making AI multiplayer." For ii-agent, the pattern is worth stealing: scheduled sentinel digests could be delivered as a Claude Tag channel message rather than email, enabling team threads and follow-up in one place.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">02 &nbsp;<span class="tag tag-anthropic">CLAUDE CODE</span></div>
      <div class="headline-title">Claude Code week 26 ships <code>mcp login/logout</code>, <code>/cd</code>, <code>/rewind</code> and 37% CPU savings</div>
      <div class="headline-body">
        The week 26 Claude Code drop (June 22–26) landed four notable additions. <strong><code>claude mcp login &lt;server&gt;</code></strong> authenticates a configured MCP server from the shell without opening the interactive /mcp menu — critical for scripted and cron-driven sessions. <strong><code>/cd &lt;path&gt;</code></strong> moves a session to a new working directory without breaking the prompt cache mid-session. <strong><code>/rewind</code></strong> now restores conversation context from before a <code>/clear</code> was run, not just generic checkpoints. And a coalescing fix in the streaming layer cut CPU usage by ~37% during long sessions. Also shipping this cycle: <strong><code>--safe-mode</code></strong> disables all customisations (CLAUDE.md, plugins, skills, hooks, MCP servers) for clean-room troubleshooting. <strong>ii-agent impact</strong>: <code>claude mcp login</code> unblocks scripted MCP authentication in the three newsletter runners without interactive prompts.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">03 &nbsp;<span class="tag tag-mcp">MCP</span></div>
      <div class="headline-title">MCP 2026-07-28 RC: stateless-by-default, handles-as-arguments, MCP Apps — final spec in 31 days</div>
      <div class="headline-body">
        The release candidate for the July 28 MCP specification dropped May 21 and is in its migration window. The biggest structural change: <strong>sessions are eliminated</strong>. The <code>initialize</code> handshake and <code>Mcp-Session-Id</code> header are gone. Every request is self-contained; servers can run behind any round-robin load balancer. State (browser handles, spreadsheet IDs, search contexts) becomes an explicit argument in every tool call — visible, auditable, cacheable. New additions: <strong>Tasks extension</strong> for long-running work via opaque handles; <strong>MCP Apps</strong> for server-rendered HTML UIs surfaced inside the client; routable headers (<code>Mcp-Method</code>, <code>Mcp-Name</code>) for gateway routing without body inspection. The MCP ecosystem now spans 15,930+ indexed servers. Migration window closes July 28 — any ii-agent skill you plan to expose as an MCP server should be built stateless-first now.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">04 &nbsp;<span class="tag tag-anthropic">ANTHROPIC ✕ AWS</span></div>
      <div class="headline-title">Claude Platform on AWS reaches GA — Anthropic-managed API with AWS billing, IAM, and full feature set</div>
      <div class="headline-body">
        Anthropic brought Claude Platform to AWS as a generally available service in May/June, completing the AWS partnership announced in April. Developers use AWS IAM credentials and AWS billing instead of Anthropic-native accounts, but the service is Anthropic-operated (data processed outside AWS infrastructure — different from Bedrock). All Claude Platform capabilities are available: managed agents, code execution, web search, prompt caching, citations, batch processing, Skills, and MCP connectors. Claude Agent SDK credit is now metered separately starting June 15 — if ii-agent's scheduled runs use the Agent SDK, check billing to confirm the credit allocation is sufficient for weekly digest + three newsletter runners combined.
      </div>
    </div>

    <div class="headline">
      <div class="headline-num">05 &nbsp;<span class="tag tag-github">GITHUB</span></div>
      <div class="headline-title">OpenClaw hits 210k stars — self-extending AI assistant that writes its own new skills is the breakout repo of 2026</div>
      <div class="headline-body">
        <strong>OpenClaw</strong> (by PSPDFKit founder Peter Steinberger) surged to 210,000+ GitHub stars in June, making it the fastest-growing AI agent repo in 2026. Its distinguishing feature: it can <em>write its own new skills</em> to extend its capabilities at runtime — a skill-composition loop that ii-agent's architecture partially implements manually. Other June trending repos include <code>garrytan/gstack</code> (a Claude Code harness with 23 opinionated tools), <code>stablyai/orca</code> (fleet-scale parallel coding agent runner), and <code>ai-boost/awesome-harness-engineering</code> (curated list for harness engineering: evals, memory, MCP, observability). The broader GitHub theme this week: tooling for running <em>more</em> agents faster with better context, not just single-agent productivity.
      </div>
    </div>
  </div>

  <!-- Architecture Watch -->
  <div class="section">
    <div class="section-title">Architecture Watch</div>

    <div class="item">
      <div class="item-title">Stateless MCP = the "handle as argument" pattern formalised at protocol level <span class="ii-pill">architecture</span></div>
      <div class="item-body">
        The 2026-07-28 RC enforces a clean design principle that ii-agent's skill architecture already follows: every resource (Excel workbook, PDF, deal context) is a file-path argument, not a shared session object. The MCP spec is now mandating this at protocol level. Any future ii-agent skill exposed as an MCP server should be stateless-first: receive all inputs as explicit arguments, execute, return a result, hold no session state. This is already true of pdf_extractor, excel_modules.py, and the slide generators — the pattern is correct. The RC makes it the only legal pattern for remote servers. Start planning the migration for any stateful skill now; the window closes July 28.
      </div>
    </div>

    <div class="item">
      <div class="item-title">P2 prompt pattern — structured contract between orchestrator and subagent is the 2026 production standard <span class="ii-pill">adopt</span></div>
      <div class="item-body">
        Multi-agent research across 2026 deployments shows that "multi-agent in production" converges on one prompting pattern: the <strong>P2 contract</strong>. Every subagent dispatch must include four explicit elements: (1) objective, (2) output format, (3) tool and source guidance, (4) clear task boundary. Teams that skip any element see hallucinated tool calls and unresolvable sub-agent failures. For ii-agent's case orchestrator, the implication is concrete: every time a module agent is dispatched (Revenue Build, Debt Schedule, etc.), the dispatch prompt should explicitly state what the sub-agent should <em>not</em> do as well as what it should do. The boundary condition is as important as the objective. LBO orchestration already partially does this via the module keyword system — making the boundary explicit is the upgrade.
      </div>
    </div>

    <div class="item">
      <div class="item-title">40% of multi-agent pilots fail in production — single-agent with the same tools wins 64% of benchmarked tasks <span class="ii-pill">strategic</span></div>
      <div class="item-body">
        A Microsoft ISE blog post on coordinator patterns this week includes a sobering finding: in controlled benchmarks, a single agent given the same tools and context as a multi-agent system matched or outperformed the multi-agent setup on 64% of tasks. 40% of multi-agent pilots fail within six months of production. The failure mode is almost always pattern mismatch — teams pick the wrong orchestration pattern, or pick correctly but don't understand how it breaks. Practical implication for ii-agent: for single-case LBO work (one analyst, one model), the sequential skill dispatch pattern may be over-engineered. The multi-agent architecture is justified for parallel sheet generation (Debt Schedule + Working Capital simultaneously) but not for sequential workflows. Audit which ii-agent cases actually benefit from parallelism before expanding the orchestration surface.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Context Engineering formalised — write / select / compress / isolate are the four production primitives</div>
      <div class="item-body">
        LangChain published a formal taxonomy of context engineering strategies this week, reducing the discipline to four primitives: <em>write</em> (author instructions and memory), <em>select</em> (retrieve only the relevant context slice), <em>compress</em> (reduce token waste via summarisation or truncation), and <em>isolate</em> (keep unrelated context physically separate across sub-agents). The companion arXiv paper (2604.04258) shows the "isolate" strategy cutting sub-agent error rates by 31% on complex multi-step tasks. For ii-agent: apply <em>isolate</em> between LBO modules — each sub-agent gets its own context window with only the inputs it needs; no module should see another module's working notes. Apply <em>select</em> to deal_memory: fetch only the tagged slice relevant to the current module at dispatch time, not the full deal context at session start.
      </div>
    </div>
  </div>

  <!-- Capability Spotlight -->
  <div class="section">
    <div class="section-title">Capability Spotlight</div>

    <div class="item">
      <div class="item-title">Claude Tag ambient mode — proactive AI teammate that monitors channels and flags forgotten threads <span class="new-pill">live June 23</span></div>
      <div class="item-body">
        Claude Tag's ambient mode is the capability most worth evaluating for personal workflows. When enabled, Tag reads the channel's message history and proactively surfaces relevant context, flags tasks that have gone quiet, and follows up on threads. For a solo ii-agent operator, this is essentially a Slack-native version of the scheduled sentinel check-in — except it responds to the actual channel activity rather than a fixed timer. Use case: a deal workspace channel where Tag tracks open questions from IC prep and nudges when a sub-task has been unaddressed for 48 hours. The ambient mode requires explicit trust delegation (which channels Tag can read proactively) — configure with care on channels that contain sensitive deal data.
      </div>
    </div>

    <div class="item">
      <div class="item-title"><code>claude mcp login</code> — scripted MCP authentication without interactive prompts <span class="ii-pill">wire this</span></div>
      <div class="item-body">
        This is the unblocking capability for fully automated sentinel runs. The new <code>claude mcp login &lt;server-name&gt;</code> command authenticates a configured MCP server from the shell, storing credentials for subsequent non-interactive sessions. Paired with <code>claude mcp logout &lt;server-name&gt;</code> for credential rotation. For ii-agent's newsletter runners (which run on schedule as background sessions), this means MCP servers (Gmail, market data, health APIs) can be pre-authenticated once and then used without human approval in every subsequent run. Wire the login step into the session-start hook so credentials are refreshed at each scheduled run without accumulating stale sessions.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Lilian Weng returns to Lil'Log — scaling laws deep-dive published June 24 <span class="new-pill">published</span></div>
      <div class="item-body">
        Lilian Weng published her first Lil'Log post in months on June 24 — a comprehensive guide to scaling laws covering Kaplan vs. Chinchilla disagreements, compute-optimal allocation, and why simple power-law extrapolation breaks down once high-quality unique tokens run low. Less directly relevant to ii-agent's finance work, but foundational for understanding why Opus 4.8 vs. Haiku 4.5 is the right skill-routing question: the performance gap between models narrows sharply on structured, tool-use tasks (where the constraint is tool schema compliance, not reasoning depth) but widens on open-ended analysis (where Opus 4.8's larger compute budget matters). The practical routing heuristic: Haiku for module dispatch and formatting; Opus for IC-level investment thesis generation and scenario analysis.
      </div>
    </div>

    <div class="item">
      <div class="item-title">AgentCore CLI v0.4.0 GA — multi-framework agent deployment in minutes, SOC 1/2/3 compliant</div>
      <div class="item-body">
        Amazon Bedrock AgentCore CLI reached GA at v0.4.0 this week. Key additions: support for Google ADK, OpenAI Agents, and Claude Agent SDK in a single deploy target; interactive shell sessions with up to 10 concurrent shells per runtime (state persists across commands); and SOC 1/2/3 compliance certification. The per-account rate limit for <code>InvokeAgentRuntime</code> increased from 25 to 200 TPS. Relevant if ii-agent's newsletter runners ever move off a local cron to a cloud execution environment — AgentCore is now the production-grade option with full compliance documentation and multi-framework support.
      </div>
    </div>
  </div>

  <!-- Workflow Ideas -->
  <div class="section">
    <div class="section-title">Workflow Ideas</div>

    <div class="item">
      <div class="item-title">Wire <code>claude mcp login</code> into the session-start hook for all three newsletter runners <span class="warning-pill">do this now</span></div>
      <div class="item-body">
        The new <code>claude mcp login &lt;server&gt;</code> command enables non-interactive MCP authentication. Add it to the <code>SessionStart</code> hook in <code>.claude/settings.json</code> for each newsletter runner:<br /><br />
        <code>"hooks": { "session_start": "claude mcp login gmail-mcp &amp;&amp; claude mcp login market-data-mcp" }</code><br /><br />
        This ensures MCP credentials are fresh at the start of every scheduled run without requiring human approval. If a credential expires mid-run, the hook fires at the next session start and silently re-authenticates. Pair with <code>claude mcp logout</code> on session-end to avoid accumulating stale credential entries.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Apply the P2 dispatch pattern to all ii-agent module calls — add explicit task boundary to every sub-agent prompt <span class="action-pill">build this</span></div>
      <div class="item-body">
        The P2 contract requires four elements per sub-agent dispatch: objective, output format, tool/source guidance, and <em>task boundary</em> (what the sub-agent should NOT do). The LBO case orchestrator currently sends objective + output format but often omits explicit boundaries. Example boundary to add for Revenue Build: "Do not model the debt schedule — that is handled by a separate agent. Do not make assumptions about entry multiple or exit timing." Adding explicit NOT-do conditions reduces cross-module contamination and hallucinated downstream assumptions. Takes 5 minutes to add to the dispatch template in the orchestration layer.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Route haiku to module dispatch, opus to IC-level analysis — model-tier routing for ii-agent skills <span class="ii-pill">optimise</span></div>
      <div class="item-body">
        Lilian Weng's scaling laws post reinforces a practical heuristic: model performance difference shrinks on structured tool-use tasks (schema compliance, data formatting) and widens on open-ended reasoning (thesis generation, scenario analysis). Map this to ii-agent: <strong>Haiku 4.5</strong> for module dispatch, formula generation, PDF extraction, and Excel formatting; <strong>Opus 4.8</strong> for investment thesis, IC oral prep, scenario analysis, and deal screening. In Claude Code, use the <code>model</code> parameter in <code>claude()</code> calls or route via <code>--model</code> flag in subagent spawns. The cost saving on the mechanical modules can fund more Opus calls for the analysis modules that actually need it.
      </div>
    </div>

    <div class="item">
      <div class="item-title">Set up a deal workspace Slack channel with Claude Tag ambient mode for IC prep tracking</div>
      <div class="item-body">
        Claude Tag's ambient mode makes sense for deal workflow tracking. Create a Slack channel per active deal, add Claude Tag with read access, and configure it to flag any message thread that goes quiet for 48+ hours. During IC prep, analysts post questions and outstanding items to the channel; Tag tracks them and proactively surfaces unresolved items before the IC date. This replaces the informal "did we follow up on that?" check with a persistent AI-driven thread monitor. Configure the trust scope carefully — give Tag read access to the deal channel only, not to your general Slack workspace, to contain exposure on sensitive deal data.
      </div>
    </div>
  </div>

  <!-- Worth Reading -->
  <div class="section">
    <div class="section-title">Worth Reading</div>

    <div class="read-item">
      <div class="read-link"><a href="https://techcrunch.com/2026/06/23/anthropics-claude-tag-is-learning-your-company-one-slack-message-at-a-time/">Anthropic's Claude Tag is learning your company, one Slack message at a time — TechCrunch</a></div>
      <div class="read-why">Best overview of how ambient mode works in practice, including the trust delegation model and the "one Claude per channel" architectural choice. Read before configuring Tag on any channel with sensitive deal data.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://devblogs.microsoft.com/ise/coordinator-patterns-multi-agent-systems/">Orchestration Patterns for Multi-Agent Systems: Performance and Trade-offs — Microsoft ISE</a></div>
      <div class="read-why">The source for the "single agent wins 64% of benchmarked tasks" finding. The post walks through five coordinator patterns (router, supervisor-worker, peer-to-peer, pipeline, ensemble) with production failure modes for each. Essential reading before expanding ii-agent's multi-agent surface.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://simonw.substack.com/p/agentic-engineering-patterns">Agentic Engineering Patterns — Simon Willison's Newsletter</a></div>
      <div class="read-why">Willison's take on the patterns that actually survive production: explicit tool schemas, safe execution layers, and why "the agent decides everything" fails at the boundary between reasoning and side effects. Short, opinionated, directly applicable to how ii-agent skills are wired.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://lilianweng.github.io/">Scaling Laws: What They Predict and Where They Break — Lil'Log (June 24)</a></div>
      <div class="read-why">Foundational for model selection decisions. The section on compute-optimal allocation explains why the Haiku/Opus routing heuristic above is principled rather than just a cost shortcut — the performance gap on structured tasks is smaller than intuition suggests.</div>
    </div>

    <div class="read-item">
      <div class="read-link"><a href="https://niteagent.com/blog/multi-agent-production-2026/">Multi-Agent in Production 2026: 3 Patterns That Survived — NiteAgent</a></div>
      <div class="read-why">The P2 prompt contract details and the "40% fail in 6 months" statistic come from here. Section 3 (why teams pick the wrong pattern) is the most useful — includes a decision tree for choosing between supervisor-worker, pipeline, and ensemble based on task characteristics.</div>
    </div>
  </div>

  <!-- Recommendation -->
  <div class="section">
    <div class="rec-box">
      <div class="rec-label">&#9889; This Week's Recommendation</div>
      <div class="rec-title">Wire <code>claude mcp login</code> into the session-start hook — unblocks fully automated scheduled runs</div>
      <div class="rec-body">
        The three newsletter runners (networth, market, health) all depend on MCP servers that require OAuth credentials. Until now, those credentials required interactive approval at least once per session, making fully headless scheduled runs fragile. The new <code>claude mcp login</code> command solves this.<br /><br />

        <strong>Step 1 — Pre-authenticate each MCP server (one-time, ~5 minutes)</strong><br />
        Run once from an interactive session for each MCP server the newsletters use:<br />
        <code>claude mcp login gmail-mcp</code><br />
        <code>claude mcp login market-data-mcp</code><br />
        Stored credentials persist across sessions.<br /><br />

        <strong>Step 2 — Wire to session-start hook (2 minutes)</strong><br />
        In <code>.claude/settings.json</code>, add:<br />
        <code>"hooks": { "session_start": "claude mcp login gmail-mcp 2>/dev/null; claude mcp login market-data-mcp 2>/dev/null" }</code><br />
        The <code>2>/dev/null</code> swallows re-auth noise when credentials are still valid.<br /><br />

        This is the single change with the highest leverage-to-effort ratio this week: it makes every future scheduled ii-agent run independent of human presence. The newsletters stop being "usually automated" and become reliably automated.
      </div>
    </div>
  </div>

  <!-- Sign-off -->
  <div class="sig">— Jordan</div>

  <!-- Footer -->
  <div class="footer">
    SENTINEL monitors Anthropic/Claude, agent frameworks, GitHub trending, and notable content weekly.<br/>
    ii-agent · robertodbabaran/ii-agent · Digest generated 2026-06-27
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
