# regen trigger 2026-07-15T17:36:25
# build: DFT roadmap regenerate trigger 2026-06-19T12:30
#!/usr/bin/env python3
"""
Fuel Finance Roadmap Generator
================================
Reads data directly from Google Sheets and generates fuel_roadmap_2026.html.

SETUP (one time only):
  1. Go to https://console.cloud.google.com/
  2. Create a project (or use existing)
  3. Enable "Google Sheets API" and "Google Drive API"
  4. Create a Service Account: IAM & Admin → Service Accounts → Create
  5. Download the JSON key: click the account → Keys → Add Key → JSON
  6. Save the key file next to this script as: service_account.json
  7. Open your Google Sheet → Share → paste the service account email → Viewer

USAGE:
  python generate_roadmap.py
  python generate_roadmap.py --output my_roadmap.html
  python generate_roadmap.py --sheet-id YOUR_SHEET_ID --key path/to/key.json

The script reads from the Google Sheet and overwrites fuel_roadmap_2026.html
(or whatever --output specifies), placed in the same folder as this script.
"""

import sys
import re
import argparse
from pathlib import Path
from datetime import datetime

# ── Check dependencies ────────────────────────────────────────────────────
try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    print("ERROR: Missing dependencies. Run:")
    print("  pip install gspread google-auth")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("ERROR: Missing pandas. Run:  pip install pandas")
    sys.exit(1)

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG — edit these if needed
# ═══════════════════════════════════════════════════════════════════════════

SHEET_ID    = "1qyZK4YNsZqXwAAfpfNZj1Ajh8qk1YOvYRo7eEeUCvK4"
KEY_FILE    = "service_account.json"   # path to your Service Account JSON key

SHEET_TABS = {
    "delivery":   "In delivery",             # in engineering
    "shipped":    "Shipped",                 # fully shipped features
    "evaluation": "In evaluation",           # taken by delivery team, being evaluated
    "ready":      "Ready for evaluation",    # fully described, waiting to be taken
    "discovery":  "Features in discovery",
    "new":        "New features for discovery",
}

# ═══════════════════════════════════════════════════════════════════════════
# FEATURE DESCRIPTIONS
# Key = display title   Value = (gtm, product, client)
# Edit when a feature's description changes. Add new entries for new features.
# ═══════════════════════════════════════════════════════════════════════════

DESCRIPTIONS = {
    "Customize Dashboard Charts": (
        "Every prospect asks \"can I customize dashboards?\" — now the answer is yes. Each chart is configurable per metric, date range and type.",
        "Widget config panel: metric selector from metrics registry, date range presets + custom, chart type toggle. Config stored per widget per dashboard.",
        "Choose what each chart shows, which time period, and how it looks. Your dashboard, your way.",
    ),
    "MCP for Metrics": (
        "Enables AI agents built on Claude, GPT or any MCP-compatible platform to query Fuel data directly. Strong signal for technical buyers.",
        "MCP server exposing metrics, dimensions and forecast data as typed tools. Supports: query_metric, list_metrics, get_forecast, run_scenario. Auth via API key.",
        "Fuel's financial data is accessible to external AI tools — so your team can build custom workflows and automations on top of your real numbers.",
    ),
    "Copilot Remembers Your Conversations": (
        "No more \"as I mentioned before\" — Copilot picks up exactly where the user left off. Reduces friction, increases daily engagement.",
        "Session message store with role/content/timestamp/referenced metrics. Context window manager with priority scoring for compression.",
        "Copilot remembers everything you discussed — no need to repeat context or re-explain what you were looking at. Picks up right where you left off.",
    ),
    "Formula Engine": (
        "Powers every custom calculation in the platform — enables clients to define their own metrics without engineering involvement.",
        "Backend service for evaluating custom metric formulas without manual per-formula coding. Deterministic calculation engine all formula features depend on.",
        "The engine behind all custom calculations in Fuel — ensures every formula you write is evaluated correctly and consistently.",
    ),
    "Forecast Period Management": (
        "Clients can extend or adjust their forecast horizon without losing existing data — removes a common friction point.",
        "Allows users to define, adjust, and extend the forecast time horizon in a controlled, non-destructive way.",
        "Extend or adjust your planning horizon at any time — add more months, shift periods — without losing any data you've already entered.",
    ),
    "Undo Function": (
        "Removes the fear of making mistakes in the forecast. Clients edit more freely when they know they can instantly revert.",
        "Reverts the most recent table action (edit, delete, formula change) without full page reload.",
        "Made a mistake in your forecast? Hit undo — the last change is reverted instantly. Edit freely without worrying about breaking anything.",
    ),
    "Value Hover Tooltip": (
        "Clients can understand why a number is what it is without leaving the grid. Reduces support questions — strong transparency signal in demos.",
        "On hover over a monthly cell: shows the inputs that produced the value, formula or method used, and delta vs. prior month.",
        "Hover over any number in the forecast to instantly see why it's that value — what inputs drove it and how it changed from the previous month.",
    ),
    "Metric Hover Tooltip": (
        "Makes the forecast self-explanatory — clients understand how each row is calculated without asking the team.",
        "On hover over a metric name (row header): shows a static row-level summary of how this metric is forecasted.",
        "Hover over any metric name to see how it's calculated — what drives it, which method is used, and what assumptions are behind it.",
    ),
    "Formula Builder": (
        "Clients can write their own calculation logic — like \"COGS = 30% of revenue\" — without asking finance. Removes a key blocker in complex deals.",
        "Formula editor with operators, functions (SUM, IF, MAX, MIN), syntax highlighting and circular reference detection. Stored as a versioned expression.",
        "Write formulas directly in your forecast — just like Excel, but connected to your real data. No spreadsheets, no workarounds.",
    ),
    "Custom Metrics from Google Sheets": (
        "Clients can bring in any KPI tracked outside accounting — NPS, pipeline coverage, custom ratios — alongside their financials.",
        "Google Sheets OAuth2 connection, standardized template with column mapping UI, metric upload + sync to metrics DB, sync history with error handling.",
        "Connect Google Sheets, map columns to Fuel fields, sync custom KPIs alongside native metrics.",
    ),
    "Compare Metrics Over Time (MoM, YoY, YTD)": (
        "One of the most common requests from prospects: \"I want to see how this month compares to last year.\" Now it's one click.",
        "Period comparison toggles: MoM, QoQ, YoY, YTD, Custom. Shows current value, prior period value, variance in $ and %.",
        "Compare any number against last month, last quarter, or last year — with the difference shown right next to the current value. One click, no setup.",
    ),
    "Drag to Extend Forecast Values": (
        "A familiar gesture — drag a cell across months to fill it. Removes friction for clients building their first forecast.",
        "Horizontal drag on forecast cells to extend values to future periods. Behavior adapts to the cell's logic capsule type (fixed value, formula, growth rate).",
        "Fill your forecast forward by dragging — like in Excel. Drag a number across months and it copies forward automatically.",
    ),
    "AI Forecast Assistant": (
        "The #1 objection to FP&A tools is setup time. A client describes what they want and the forecast updates live — great for demos.",
        "NL intent parsing → structured forecast operations. Covers: add/remove rows, set values, apply formulas, create scenarios.",
        "Tell the forecast what to do in plain English — \"add 3 hires in Q3\" or \"grow revenue 10% from April\" — and it updates instantly.",
    ),
    "Forecast Sidebar": (
        "Turns a complex grid into a conversation. Show a CFO typing a question and watching the model update — closes deals.",
        "Split view: forecast grid left, sidebar right. Forecast management moves to the sidebar — you can immediately see the impact of changes on the full forecast.",
        "Click any number to open a sidebar — change values, see their instant impact on the full forecast, or ask questions in plain English.",
    ),
    "Context Awareness": (
        "Copilot answers become dramatically more relevant when it knows the company — business model, challenges, growth stage, goals.",
        "Structured context layer: company profile (business model, revenue streams, stage, goals, pain points) + user profile. Persisted across sessions. Injected into every Copilot request.",
        "Copilot understands your business — your model, your goals, your current challenges — so its answers are relevant to your specific situation.",
    ),
    "Auto-Save Key Takeaways from AI Chats": (
        "Shows Fuel builds institutional memory, not just answers. Great story for CFOs who manage knowledge across their team.",
        "Post-session job: extract topics, decisions, metrics referenced, open items. Saved to conversation_summaries with session + user + company linkage.",
        "After each Copilot session, key decisions and insights are automatically saved — nothing gets lost, and context carries into your next conversation.",
    ),
    "Build Custom Reports": (
        "Every CFO has a specific monthly report they send to the CEO or board. Build it once, reuse it every month with fresh data.",
        "Report builder: metric selector, date range config, filter builder, display options, save as template with name + visibility scope.",
        "Build the exact report you need — pick your metrics, set the time period, add filters — then save it as a template that refreshes automatically.",
    ),
    "Filter Reports by Legal Entity": (
        "Multi-entity is a hard requirement for our ICP. This unlocks deals with companies that have subsidiaries or international operations.",
        "Entity dimension on all financial reports. Consolidated view applies intercompany elimination rules. Filter persists across report types in session.",
        "If you have multiple entities or subsidiaries, switch between them in one click — see each separately or consolidated together.",
    ),
    "Hiring Plan & Payroll Forecast": (
        "People costs are 60–70% of expenses for our ICP. This is the #1 forecast feature prospects ask about in sales calls.",
        "Structured headcount table: role × dept × start date × salary × benefits %. Engine computes fully-loaded monthly cost. Auto-feeds into Personnel section.",
        "Add each planned hire with role, salary and start date — Fuel calculates the full cost including benefits and adds it to your forecast automatically.",
    ),
    "Rolling Forecast": (
        "Prospects are tired of forecasts that go stale. This one updates itself — actuals lock in, future periods adjust. Strong retention driver.",
        "On actuals import: lock closed periods, compute variance, re-project open periods using latest assumptions, append new period at horizon end.",
        "Your forecast stays current automatically — when real numbers come in they lock in, and the rest of the forecast adjusts forward. Always a live view.",
    ),
    "Financial Thinking Mode": (
        "Finance leaders need to trust AI output. This shows exactly how every number was derived — kills the \"black box\" objection in sales.",
        "Extended reasoning panel: calculation breakdown (formula + inputs + result), source traceability, assumptions log, methodology note.",
        "See exactly how Copilot arrived at any answer — which data it used, what formula it applied, and what it assumed. Full transparency, no black box.",
    ),
    "Build Your Own Financial Metrics": (
        "Every company tracks something custom that standard tools don't cover. Removes the \"but we measure it differently\" objection.",
        "Formula builder for derived metrics: operators + functions over existing metric registry. Sharing: private / team / org.",
        "Create metrics that match how your business thinks — like \"Adjusted EBITDA\" or \"Revenue per Employee\" — and use them anywhere in reports and dashboards.",
    ),
    "Page Context Capture": (
        "Makes every Copilot answer feel like it was written for exactly what the user is looking at.",
        "Detects current page type, active entity, report, and period. Captures all visible metric IDs and values. Injected into Copilot context on every message.",
        "Copilot reads what's on your screen — the exact metrics, the date range, the entity — and uses those numbers directly in its answers.",
    ),
    "AI Summary in Forecast": (
        "CFOs spend hours writing board commentary. This drafts it in seconds. One of the highest-wow features in demos.",
        "AI-generated narrative: key assumptions, growth trajectory, cost drivers, risks. 2–3 paragraphs, labeled as AI-generated. User can edit or regenerate.",
        "Get a plain-English summary of your forecast — key growth assumptions, biggest cost drivers, cash runway, main risks — generated in seconds.",
    ),
    "Department Budget Planning": (
        "CFOs at $10M+ companies need to involve department heads in budgeting without giving everyone access to sensitive data.",
        "Role-scoped forecast workspace: section assignment per user/role, completion status tracker, approval workflow (submit → review → approve).",
        "Each team lead fills in only their own budget. Marketing sees marketing, engineering sees engineering. Finance sees everything consolidated.",
    ),
    "AI-Guided Forecast Setup": (
        "Eliminates the blank-page problem. A new client answers 5 questions and gets a working forecast. Dramatically reduces time-to-value.",
        "Conversational setup: business model → revenue drivers → cost structure → growth stage → headcount. AI maps responses to forecast template.",
        "New to forecasting in Fuel? Answer a few questions about your business and the AI builds your starting forecast — the right structure, ready to fill in.",
    ),
    "AI Explains Anything in Fuel": (
        "Reduces support burden and increases confidence for non-finance users. Great for companies where the CEO uses Fuel directly.",
        "Knowledge base: finance concepts by category + difficulty. User proficiency tracking from interaction patterns. In-context retrieval for Copilot responses.",
        "Not sure what a metric means or how a report works? Just ask Copilot — it explains anything in plain English, using your own numbers as examples.",
    ),
    "MCP Server for Fuel": (
        "Enables any MCP-compatible AI platform to read and interact with Fuel data — new distribution channel for technical buyers.",
        "Full MCP server implementation: typed tools for metrics, reports, forecast data. Auth via API key. Supports tool discovery.",
        "Fuel's data is accessible to any AI tool your team uses — connect it once and query your financials from anywhere.",
    ),
    "Reports Page Design": (
        "Updated reports interface that makes financial data easier to navigate and act on.",
        "Redesigned reports page layout: improved navigation, better data hierarchy, faster filtering, cleaner visual structure.",
        "A cleaner, faster reports experience — everything you need is easier to find and use.",
    ),
    "Company Selector": (
        "Multi-entity switching without leaving the current page. Critical for CFOs managing multiple companies or subsidiaries.",
        "Fast entity switcher: keyboard shortcut, recent entities list, search. Context preserved on switch.",
        "Jump between companies or entities instantly — no full page reload, no lost context. One keyboard shortcut away.",
    ),
    "Page Tutorials": (
        "Reduces onboarding time and support burden. New users understand Fuel faster — improves activation rates.",
        "Contextual in-app tutorials triggered by page type and user activity. Step-by-step guided flows with skip/resume.",
        "Short guided tours that show you exactly how to use each part of Fuel — available any time, skippable, resumable.",
    ),
    "Customer Retention Analysis (Cohorts)": (
        "Subscription and recurring revenue businesses need this to understand churn.",
        "Cohort matrix: rows = cohorts by signup month, columns = months since start. Heatmap overlay. Filters: segment, entity.",
        "See how well you keep customers over time, grouped by when they started. Spot which groups stay longest.",
    ),
    "Single Demo Account": (
        "Internal tool for sales demos — consistent demo environment for the team.",
        "Shared demo company with pre-populated data, reset functionality, and role-based demo scenarios.",
        "A dedicated demo environment for showing Fuel to prospects.",
    ),
    "Editing Forecast with AI Agent": (
        "Removes the manual editing bottleneck — clients can make complex forecast changes just by describing them.",
        "AI agent interprets natural language edit commands and applies structured changes to the forecast model: adjust values, add rows, apply assumptions across periods.",
        "Just tell the AI what to change in your forecast — \"increase marketing spend by 15% from Q2\" — and it updates the numbers for you.",
    ),
    "AI Data Quality Checker": (
        "Finance teams spend hours fixing bad data. This catches issues automatically.",
        "Rules engine: required fields, anomaly detection (z-score + ML), issue queue with severity levels, auto-fix suggestions.",
        "Fuel automatically checks your data for problems — missing categories, unusual amounts, duplicates.",
    ),
    "AI Bank & Accounting Reconciliation": (
        "Reconciliation is one of the most painful manual tasks for our ICP.",
        "Two-pass matching: exact then fuzzy (±$0.50, ±3 days). Match review UI with confidence scores.",
        "Fuel automatically matches your bank transactions against your accounting records and flags discrepancies.",
    ),
    "Sensitivity Analysis": (
        "Board meetings always involve \"what if revenue is 20% lower?\" — now CFOs answer that in seconds.",
        "Variable selection (1–5), range definition (low/base/high), tornado chart showing impact rank on selected output.",
        "Test how sensitive your forecast is to key assumptions — change revenue growth and instantly see the impact.",
    ),
    "Smart Alerts": (
        "Keeps Fuel top-of-mind between logins. Clients get notified when something needs attention.",
        "Alert rule builder: metric + condition + threshold + notification targets. Email + Slack delivery.",
        "Set thresholds that matter — \"alert me if cash runway drops below 3 months\" — and Fuel notifies you automatically.",
    ),
    "NetSuite · Xero · Salesforce": (
        "Unlocks the segment of our ICP that has outgrown QuickBooks.",
        "NetSuite: subsidiary sync. Xero: OAuth2 multi-tenant. Salesforce: OAuth2, configurable object + field sync.",
        "If your company uses NetSuite, Xero, or Salesforce, Fuel connects directly — your data syncs automatically.",
    ),
    "Weekly Cash Flow Planning": (
        "Cash gaps are a real pain for our ICP. Weekly visibility is a strong hook, especially for companies with uneven payment cycles.",
        "Weekly period grid (ISO weeks), line item structure, opening/closing balance, actuals auto-population from bank feeds, rolling horizon.",
        "Track cash week by week. Bank transactions fill in automatically — so you always know your real cash position and when gaps might appear.",
    ),
    "AI Insights on Every Metric and Chart": (
        "Turns passive dashboards into active alerts. Clients don't have to hunt for problems — Fuel surfaces them automatically.",
        "Post-sync background job: per-metric analysis (MoM trend, YoY, forecast deviation, anomalies). Insight stored with severity. Badge rendered on widget.",
        "When something unusual happens in your numbers — a spike, a drop, an anomaly — a small badge appears on that metric. Click to see what happened, in plain English.",
    ),
    "Customer Retention Analysis (Cohorts)": (
        "Subscription and recurring revenue businesses need this to understand churn. Strong hook for SaaS and services companies.",
        "Cohort matrix: rows = cohorts by signup month, columns = months since start. Heatmap overlay. Filters: segment, entity, plan, region.",
        "See how well you keep customers over time, grouped by when they started. Spot which groups stay longest and which leave early — at a glance.",
    ),
    "Sage Integration": (
        "Opens the door to companies on Sage who currently can't use Fuel. Expands our addressable market significantly.",
        "Supports Sage 50 (local connector), Sage Intacct (web + company ID), Business Cloud (OAuth2). Syncs GL, AP, AR. Daily incremental sync.",
        "If your company uses Sage for accounting, Fuel connects directly — your financial data syncs automatically, no manual exports.",
    ),
    "HubSpot Integration": (
        "Pipeline data from CRM + financials in Fuel = full revenue picture. Unlocks deals with sales-led companies.",
        "OAuth2, configurable object sync (Deals, Contacts, Companies), property mapping UI, pipeline stage mapping, daily sync.",
        "Connect your HubSpot CRM to Fuel — your sales pipeline flows in automatically, so you can see revenue projections alongside your actual financials.",
    ),
    "AI Data Quality Checker": (
        "Finance teams spend hours fixing bad data. This catches issues automatically — strong ROI story for companies migrating from Excel.",
        "Rules engine: required fields, anomaly detection (z-score + ML), issue queue with severity levels, auto-fix suggestions.",
        "Fuel automatically checks your data for problems — missing categories, unusual amounts, duplicates — and suggests fixes before they affect your reports.",
    ),
    "AI Bank & Accounting Reconciliation": (
        "Reconciliation is one of the most painful manual tasks for our ICP. Automating it is a direct hours-saved-per-month story.",
        "Two-pass matching: exact (amount + date + ref) then fuzzy (±$0.50, ±3 days). Match review UI with confidence scores. Adjustment entry generator.",
        "Fuel automatically matches your bank transactions against your accounting records, flags anything that doesn't line up, and suggests how to fix it.",
    ),
    "Sensitivity Analysis": (
        "Board meetings always involve \"what if revenue is 20% lower?\" — now CFOs answer that in seconds, not days.",
        "Variable selection (1–5), range definition (low/base/high), independent variation, tornado chart output showing impact rank on selected output metric.",
        "Test how sensitive your forecast is to key assumptions — change revenue growth, churn, or headcount and instantly see the impact on profit, cash and runway.",
    ),
    "Smart Alerts": (
        "Keeps Fuel top-of-mind between logins. Clients get notified when something needs attention — before it becomes a problem.",
        "Alert rule builder: metric + condition + threshold + notification targets. Evaluates on every forecast save + actuals import. Email + Slack delivery.",
        "Set thresholds that matter — \"alert me if cash runway drops below 3 months\" — and Fuel notifies you automatically before it becomes a problem.",
    ),
    "NetSuite · Xero · Salesforce": (
        "Unlocks the segment of our ICP that has outgrown QuickBooks. NetSuite alone opens a large portion of $10M+ companies we can't serve today.",
        "NetSuite: TBA auth, subsidiary sync, custom fields. Xero: OAuth2 multi-tenant. Salesforce: OAuth2, configurable object + field sync.",
        "If your company uses NetSuite, Xero, or Salesforce, Fuel connects directly — your data syncs automatically so reports and forecasts are always up to date.",
    ),
    "Dashboard Builder": (
        "Custom dashboards close deals with CEOs who want a specific view. This is the \"build your own\" answer to every custom dashboard request.",
        "Widget library: KPI cards, line/bar/pie charts, data tables, comparison widgets, text blocks. 12-column grid with snap positioning. Per-role sharing.",
        "Build exactly the dashboard you need — drag and drop charts, KPI cards and tables, arrange them how you like, and share with your team.",
    ),
    "Control Who Sees What": (
        "Enterprise deals often stall on access control. This unblocks them — granular permissions per user, per entity, per report.",
        "Role definitions: Admin, Finance Manager, Dept Head, Viewer + custom. Permission scoping: entity, dashboard, report, forecast section. Audit log.",
        "Decide exactly who sees what — board members see summaries, department heads see their own budgets, finance sees everything.",
    ),
    "Plan-Actual, Plan-Plan Comparison in Forecast": (
        "Finance teams preparing for board meetings need instant variance analysis. Eliminates manual export to Excel for plan vs. actual review.",
        "Side-by-side columns in the forecast grid: Budget vs. Actuals, Scenario A vs. Scenario B. Delta columns ($, %) with color-coded variance. Switchable comparison pairs.",
        "Compare your forecast against actuals or another scenario directly in the grid — see exactly where you're over or under plan, without leaving Fuel.",
    ),
    "Add Custom Roles with Specific Permissions": (
        "Enterprise deals require role-based access control. Custom roles let us match any org structure — removes the #1 security objection in mid-market deals.",
        "Role builder: name + description, permission matrix (read/write/admin per module), role assignment to users. Inherits from base role templates.",
        "Create roles that match your team structure — give your CFO full access, department heads their own section, and investors read-only dashboards.",
    ),
    "Access to Reports/Dashboards Management": (
        "Multi-team companies need report-level access control. Closes data governance gaps that block enterprise adoption.",
        "Per-report and per-dashboard permission layer: view / edit / share rights assignable per role. Inherits entity-level scope from role definition.",
        "Control which teams or individuals can view, edit, or share each report and dashboard — keep sensitive financials private while sharing what matters.",
    ),
    "Access to Data Management": (
        "Data governance is a blocker for finance teams handling sensitive GL data. Scoped data access unlocks regulated industries and multi-entity setups.",
        "Data-level permission scoping: restrict by integration source, entity, account group, or date range. Applied at query level — not just UI visibility.",
        "Decide which users can see raw data from each integration or entity — your US entity data stays separate from EU, payroll stays private.",
    ),
    "Access to Settings Management": (
        "Prevents accidental mis-configuration by non-admin users. Required for enterprise deployments with multiple admins and strict change control.",
        "Granular settings permissions: integration management, user management, currency/period config, and billing — each assignable independently per role.",
        "Control who can change integrations, add users, or modify company settings — so only the right people can make structural changes.",
    ),
    "Access to Forecast Management": (
        "Forecast integrity is critical for board reporting. Scoped forecast access prevents unauthorized edits while enabling collaborative planning.",
        "Forecast permission layer: view / edit / lock rights per section, scenario, and entity. Department heads edit their rows only; finance team locks and reviews.",
        "Choose who can edit, view, or lock each part of the forecast — department heads update their own sections, finance controls everything else.",
    ),
    "New Dashboard Design": (
        "Dashboard UX directly affects demo conversion. A modern, polished design signals product maturity and helps close design-sensitive buyers.",
        "Redesigned dashboard shell: updated layout grid, refreshed widget styles, improved information hierarchy, better empty states, and responsive behavior.",
        "A cleaner, more modern dashboard layout — easier to read at a glance, better organized, and more visually polished for sharing with stakeholders.",
    ),
    "Cohort Analysis (Projections)": (
        "Subscription companies need to project revenue from new cohorts. Closes a major gap vs. spreadsheet models.",
        "Retention curve inputs, new cohort planning table, projected matrix combining actuals + forecasts, P&L integration, scenario comparison.",
        "Project how much revenue you'll get from future customers — based on how your existing ones behave. See expected revenue from new cohorts month by month.",
    ),
}

DEFAULT_DESC = (
    "Coming soon — details will be added as discovery progresses.",
    "Coming soon — details will be added as discovery progresses.",
    "Coming soon — details will be added as discovery progresses.",
)

# ═══════════════════════════════════════════════════════════════════════════
# PROGRESS BARS — edit when stage changes
# (progress%, stage_label, hex_color)
# ═══════════════════════════════════════════════════════════════════════════

PROGRESS = {
    # ── March — In Delivery ───────────────────────────────────────────────
    "Customize Dashboard Charts":                (90,  "Partially shipped",  "#22c55e"),
    "MCP for Metrics":                           (100, "Fully shipped",      "#22c55e"),
    "Copilot Remembers Your Conversations":      (100, "Fully shipped",      "#22c55e"),
    "Formula Engine":                            (100, "Fully shipped",      "#22c55e"),
    "Forecast Period Management":                (100, "Fully shipped",      "#22c55e"),
    "Undo Function":                             (100, "Fully shipped",      "#22c55e"),
    "Value Hover Tooltip":                       (100, "Fully shipped",      "#22c55e"),
    "Metric Hover Tooltip":                      (100, "Fully shipped",      "#22c55e"),
    # ── April — In Delivery / Ready ──────────────────────────────────────
    "AI Forecast Assistant":                     (70,  "In development",     "#22c55e"),
    "Formula Builder":                           (50,  "Ready to deliver",   "#f59e0b"),
    "Custom Metrics from Google Sheets":         (50,  "Ready to deliver",   "#f59e0b"),
    "Drag to Extend Forecast Values":            (50,  "Ready to deliver",   "#f59e0b"),
    "Reports Page Design":                       (15,  "Research",           "#6366f1"),
    "Build Custom Reports":                      (30,  "Preparing PRD",      "#8b5cf6"),
    # ── May — Ready / Discovery ──────────────────────────────────────────
    "MCP Server for Fuel":                       (50,  "Ready to deliver",   "#f59e0b"),
    "Context Awareness":                         (50,  "Ready to deliver",   "#f59e0b"),
    "Forecast Sidebar":                          (45,  "In design",          "#f59e0b"),
    "Compare Metrics Over Time (MoM, YoY, YTD)":(45,  "In design",          "#f59e0b"),
    "Company Selector":                          (45,  "In design",          "#f59e0b"),
    "Build Your Own Financial Metrics":          (15,  "Research",           "#6366f1"),
    "Page Tutorials":                            (15,  "Research",           "#6366f1"),
    # Planned features have no progress bar — add entry here to show one
}

# ═══════════════════════════════════════════════════════════════════════════
# PREVIEW THUMBNAILS (prototype screenshot URLs)
# ═══════════════════════════════════════════════════════════════════════════

PREVIEWS = {}  # cleared for DFT — keyed by Fuel titles, no match

# ═══════════════════════════════════════════════════════════════════════════
# ONE-PAGERS (PDF files)
# Key = display title   Value = relative path to PDF in one-pagers/ folder
# Add a row here when a one-pager PDF is ready for a feature.
# If a feature has no entry → falls back to the "Description" link as before.
# ═══════════════════════════════════════════════════════════════════════════

ONE_PAGERS = {}  # cleared for DFT — keyed by Fuel titles, no match

# ═══════════════════════════════════════════════════════════════════════════
# TITLE NORMALIZATION
# Maps Excel/Sheets "Epic" names → display names in the roadmap
# Add a row here whenever a title in the sheet differs from the roadmap title
# ═══════════════════════════════════════════════════════════════════════════

TITLE_MAP = {
    # Only __SKIP__ entries — to suppress sub-task duplicates
    "Context Window Manager":          "__SKIP__",
    "Conversation Summary Generator":  "__SKIP__",
}

AREA_CONFIG = {
    "Platform foundations":      ("--platform",     "Platform foundations"),
    "Data engine":               ("--data",         "Data engine"),
    "Shared context (16 layers)":("--context",      "Shared context"),
    "Agents (10 in MVP)":        ("--agents",       "Agents"),
    "Workflows":                 ("--workflows",    "Workflows"),
    "Cabinet \u2014 pages":       ("--cabinet",      "Cabinet pages"),
    "Agent settings UI":         ("--settings",     "Agent settings"),
    "Custom artifacts":          ("--artifacts",    "Custom artifacts"),
    "Slack-first interface":     ("--slack",        "Slack-first"),
    "Cross-cutting concerns":    ("--crosscut",     "Cross-cutting"),
    "AI Dev Acceleration":       ("--aidev",        "AI Dev Acceleration"),
    "Documentation Tool":        ("--docs",         "Documentation Tool"),
}

# Normalize sheet area values that differ from AREA_CONFIG keys.
AREA_ALIASES = {
    "Data integrations": "Data engine",
    "Agents": "Agents (10 in MVP)",
    "Context": "Shared context (16 layers)",
}

# Short labels for the filter buttons (fallback = the area string itself)
AREA_FILTER_LABEL = {
    "Platform foundations": "Platform",
    "Data engine": "Data engine",
    "Shared context (16 layers)": "Context",
    "Agents (10 in MVP)": "Agents",
    "Workflows": "Workflows",
    "Cabinet \u2014 pages": "Cabinet",
    "Agent settings UI": "Agent settings",
    "Custom artifacts": "Artifacts",
    "Slack-first interface": "Slack",
    "Cross-cutting concerns": "Cross-cutting",
    "AI Dev Acceleration": "AI Dev",
    "Documentation Tool": "Docs Tool",
}

def build_filter_buttons(features, include_marketing=True):
    """Render filter buttons only for areas actually present among features."""
    present = [a for a in AREA_CONFIG if any(f.get("area") == a for f in features)]
    extras = [a for a in dict.fromkeys(f.get("area") for f in features)
              if a and a not in AREA_CONFIG]
    btns = ['<button class="filter-btn active" data-filter="all">All</button>']
    for a in present + extras:
        lbl = AREA_FILTER_LABEL.get(a, a)
        btns.append(f'<button class="filter-btn" data-filter="{esc(a)}">{esc(lbl)}</button>')
    if include_marketing and any(f.get("marketing") for f in features):
        btns.append('<button class="filter-btn" data-filter="marketing">✦ Marketing</button>')
    return "\n    ".join(btns)

def inject_filters(base, features, include_marketing=True):
    return base.replace("<!--FILTER_BTNS-->", build_filter_buttons(features, include_marketing))


# ═══════════════════════════════════════════════════════════════════════════
# GOOGLE SHEETS READER
# ═══════════════════════════════════════════════════════════════════════════

def connect_sheets(key_file, sheet_id):
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_file(key_file, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(sheet_id)

def sheet_to_df(spreadsheet, tab_name):
    """Read a tab and return a pandas DataFrame.

    Uses get_all_values() instead of get_all_records() so that blank rows in
    the middle of the sheet don't silently truncate the data.
    get_all_records() stops at the first fully-empty row — a common pitfall
    when the sheet has visual separator rows between feature groups.
    """
    try:
        ws = spreadsheet.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        print(f"  ⚠ Tab not found: '{tab_name}' — skipping")
        return pd.DataFrame()

    all_values = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
    if not all_values:
        return pd.DataFrame()

    headers = all_values[0]
    rows = []
    for row in all_values[1:]:
        # Pad row to header length if shorter
        padded = list(row) + [""] * (len(headers) - len(row))
        # Skip rows where EVERY cell is empty (pure separator rows)
        if all(str(v).strip() == "" for v in padded):
            continue
        rows.append(dict(zip(headers, padded)))

    return pd.DataFrame(rows)

def fmt_date(val):
    """Return 'Apr 2026' or None.
    gspread with UNFORMATTED_VALUE returns dates as Excel serial numbers (float).
    E.g. 2026-04-30 → 46941.0
    """
    import datetime as _dt
    if val is None or val == "" or val == 0:
        return None
    # Excel serial number (gspread UNFORMATTED_VALUE)
    if isinstance(val, (int, float)):
        try:
            serial = int(val)
            if serial < 40000 or serial > 60000:   # not a plausible date serial
                return None
            # Excel epoch: Dec 30, 1899 (with Lotus 1-2-3 leap year bug)
            dt = _dt.date(1899, 12, 30) + _dt.timedelta(days=serial)
            return dt.strftime("%b %Y")
        except Exception:
            return None
    # String fallbacks
    s = str(val).strip()
    if not s or s.lower() in ("nan", "nat", "none", "0"):
        return None
    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%d.%m.%Y",
        "%d/%m/%Y",
        "%m/%d/%Y %H:%M:%S",
    ]
    for fmt in formats:
        try:
            dt = _dt.datetime.strptime(s, fmt)
            return dt.strftime("%b %Y")
        except ValueError:
            continue
    try:
        dt = pd.to_datetime(s, dayfirst=True)
        if pd.isna(dt):
            return None
        return dt.strftime("%b %Y")
    except Exception:
        return None

def iso_date(val):
    """Return 'YYYY-MM-DD' for a delivery cell (serial number or string), else ''."""
    import datetime as _dt
    if val is None or val == "" or val == 0:
        return ""
    if isinstance(val, (int, float)):
        try:
            serial = int(val)
            if serial < 40000 or serial > 60000:
                return ""
            return (_dt.date(1899, 12, 30) + _dt.timedelta(days=serial)).isoformat()
        except Exception:
            return ""
    sv = str(val).strip()
    if not sv or sv.lower() in ("nan", "nat", "none", "0"):
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y %H:%M:%S"):
        try:
            return _dt.datetime.strptime(sv, fmt).date().isoformat()
        except ValueError:
            continue
    try:
        import pandas as _pd
        return _pd.to_datetime(sv).date().isoformat()
    except Exception:
        return ""

def get_col(deliver_str):
    """Map delivery date string → column key."""
    if not deliver_str:
        return "march"  # no date -> first active delivery month (July)
    try:
        dt = pd.to_datetime(deliver_str)
        if dt.year < 2026 or (dt.year == 2026 and dt.month < 7):
            return "february"  # before July -> history
        elif dt.year == 2026 and dt.month == 7:
            return "march"   # July 2026
        elif dt.year == 2026 and dt.month == 8:
            return "april"   # August 2026
        elif dt.year == 2026 and dt.month == 9:
            return "may"     # September 2026
        elif dt.year == 2026 and dt.month == 10:
            return "june"    # October 2026
        else:
            return "july"    # Later
    except Exception:
        return "july"

def get_links(row):
    """Extract Notion/Figma/prototype links from a row dict."""
    links = []
    mapping = [
        ("Description", "ln",   ["Description"]),
        ("PRD",         "lprd", ["PRD"]),
        ("Design",      "ld",   ["Brief for design", "Design"]),
        ("Research",    "lr",   ["Market research", "Market Research"]),
        ("Prototype",   "lpr",  ["Prototype"]),
        ("Release Brief","lbr",  ["Release Brief link"]),
    ]
    for label, css, col_names in mapping:
        for col in col_names:
            val = str(row.get(col, "")).strip()
            if val and val.lower() not in ("", "nan", "not required") and "http" in val:
                links.append((label, css, val))
                break
    return links

def normalize_status(stage_str):
    s = str(stage_str).strip().lower()
    if not s or s in ("nan", ""):
        return "Planned"
    if any(x in s for x in ["ready for delivery", "ready to deliver"]):
        return "Ready to Deliver"
    if any(x in s for x in ["development", "ship", "delivery"]):
        return "In Delivery"
    if any(x in s for x in ["design", "prd", "preparing", "research", "collecting", "discovery"]):
        return "Discovery"
    return "Planned"

def clean_desc(val):
    """Strip whitespace, return empty string if blank/nan."""
    s = str(val).strip() if val is not None else ""
    return "" if s.lower() in ("nan", "none", "") else s

def clean_bool(val):
    """Return True if value is truthy (True/true/1/yes). Safe against any type."""
    try:
        if val is None or val == "" or val is False: return False
        return str(val).strip().lower() in ("true", "1", "yes")
    except Exception:
        return False

def read_features(spreadsheet):
    features = []
    seen = set()

    def add(title_raw, area, status, disc_date, deliver_date, internal, links,
             sheet_status="", desc_gtm="", desc_product="", desc_client="", marketing=False,
             pilot_date="", rollout_date="", stage="", responsible="", connected="", deliver_iso="", disc_iso=""):
        title = TITLE_MAP.get(title_raw, title_raw)
        if title == "__SKIP__" or not title or title == "nan":
            return
        if title in seen:
            return
        seen.add(title)
        area = AREA_ALIASES.get((area or "").strip(), area)
        features.append({
            "title": title,
            "area": area or "Platform foundations",
            "status": status,
            "sheet_status": sheet_status,
            "disc_date": disc_date,
            "deliver_date": deliver_date,
            "internal": internal,
            "links": links,
            "desc_gtm": desc_gtm,
            "desc_product": desc_product,
            "desc_client": desc_client,
            "marketing": marketing,
            "pilot_date": pilot_date,
            "rollout_date": rollout_date,
            "stage": clean_desc(stage),
            "responsible": clean_desc(responsible),
            "connected": clean_desc(connected),
            "deliver_iso": deliver_iso,
            "disc_iso": disc_iso,
        })

    # ── Shipped (fully delivered features) — read FIRST so past dates win ─
    df = sheet_to_df(spreadsheet, SHEET_TABS["shipped"])
    for _, row in df.iterrows():
        epic = str(row.get("Feature") or row.get("Epic", "")).strip()
        if not epic or epic == "nan":
            continue
        internal = str(row.get("Client-facing?", "Yes")).strip().lower() in ("no", "false", "0")
        deliver = fmt_date(
            row.get("Delivered date")
            or row.get("Date of delivery")
            or row.get("Moved to delivery stage date")
            or row.get("Expected month of delivery")
            or row.get("Expected end date of the delivery", "")
        )
        sheet_status = str(row.get("status", "")).strip() or "Fully shipped"
        add(epic,
            area=str(row.get("Area", "Platform foundations")).strip(),
            connected=row.get("Connected features", "") or row.get("Blocking features", ""),
            deliver_iso=iso_date(row.get("Date of delivery", "")),
            disc_iso=iso_date(row.get("Date of discovery end", "")),
            status="In Delivery",
            disc_date=fmt_date(row.get("Date of discovery end", "")),
            deliver_date=deliver,
            internal=internal,
            links=get_links(row),
            sheet_status=sheet_status,
            desc_gtm=clean_desc(row.get("Desc GTM", "")),
            desc_product=clean_desc(row.get("Desc Product", "")),
            desc_client=clean_desc(row.get("Desc Client", "")),
            marketing=clean_bool(row.get("Is significant for separate marketing release?", False)),
            pilot_date=fmt_date(row.get("Pilot date", "") or ""),
            rollout_date=fmt_date(row.get("Full rollout date", "") or ""))

    # ── In delivery (features actively in engineering) ───────────────────
    df = sheet_to_df(spreadsheet, SHEET_TABS["delivery"])
    for _, row in df.iterrows():
        epic = str(row.get("Feature") or row.get("Epic", "")).strip()
        if not epic or epic == "nan":
            continue
        internal = str(row.get("Client-facing?", "Yes")).strip().lower() in ("no", "false", "0")
        deliver = fmt_date(
            row.get("Date of delivery")
            or row.get("Expected month for the delivery")
            or row.get("Expected end date of the delivery", "")
            or row.get("Expected month of delivery", "")
            or row.get("Expected month of delivery (Oleksandr's gestimation)", "")
        )
        # Use actual status from sheet for progress bar lookup
        sheet_status = str(row.get("status", "")).strip()
        add(epic,
            area=str(row.get("Area", "Platform foundations")).strip(),
            connected=row.get("Connected features", "") or row.get("Blocking features", ""),
            deliver_iso=iso_date(row.get("Date of delivery", "")),
            disc_iso=iso_date(row.get("Date of discovery end", "")),
            status="In Delivery",
            disc_date=fmt_date(row.get("Date of discovery end", "")),
            deliver_date=deliver,
            internal=internal,
            links=get_links(row),
            sheet_status=sheet_status,
            desc_gtm=clean_desc(row.get("Desc GTM", "")),
            desc_product=clean_desc(row.get("Desc Product", "")),
            desc_client=clean_desc(row.get("Desc Client", "")),
            marketing=clean_bool(row.get("Is significant for separate marketing release?", False)),
            pilot_date=fmt_date(row.get("Pilot date", "") or ""),
            rollout_date=fmt_date(row.get("Full rollout date", "") or ""))

    # ── In evaluation (taken by delivery team, being evaluated) ─────────
    df = sheet_to_df(spreadsheet, SHEET_TABS["evaluation"])
    for _, row in df.iterrows():
        epic = str(row.get("Feature") or row.get("Epic", "")).strip()
        if not epic or epic == "nan":
            continue
        internal = str(row.get("Client-facing?", "Yes")).strip().lower() in ("no", "false", "0")
        deliver = fmt_date(
            row.get("Date of delivery") or row.get("Expected month of delivery") or row.get("Expected month of delivery (Oleksandr's gestimation)", "")
            or row.get("Expected end date of the delivery", "")
        )
        add(epic,
            area=str(row.get("Area", "Platform foundations")).strip(),
            connected=row.get("Connected features", "") or row.get("Blocking features", ""),
            deliver_iso=iso_date(row.get("Date of delivery", "")),
            disc_iso=iso_date(row.get("Date of discovery end", "")),
            status="Ready to Deliver",
            disc_date=fmt_date(row.get("Date of discovery end", "")),
            deliver_date=deliver,
            internal=internal,
            links=get_links(row),
            desc_gtm=clean_desc(row.get("Desc GTM", "")),
            desc_product=clean_desc(row.get("Desc Product", "")),
            desc_client=clean_desc(row.get("Desc Client", "")),
            marketing=clean_bool(row.get("Is significant for separate marketing release?", False)),
            pilot_date=fmt_date(row.get("Pilot date", "") or ""),
            rollout_date=fmt_date(row.get("Full rollout date", "") or ""))

    # ── Ready for evaluation (fully described, not yet taken) ────────────
    df = sheet_to_df(spreadsheet, SHEET_TABS["ready"])
    for _, row in df.iterrows():
        epic = str(row.get("Feature") or row.get("Epic", "")).strip()
        if not epic or epic == "nan":
            continue
        internal = str(row.get("Client-facing?", "Yes")).strip().lower() in ("no", "false", "0")
        disc  = fmt_date(row.get("Date of discovery end") or row.get("Expected date of ready of discovery", ""))
        deliv = fmt_date(row.get("Date of delivery") or row.get("Expected month of delivery") or row.get("Expected month of delivery (Oleksandr's gestimation)", ""))
        add(epic,
            area=str(row.get("Area", "Platform foundations")).strip(),
            connected=row.get("Connected features", "") or row.get("Blocking features", ""),
            deliver_iso=iso_date(row.get("Date of delivery", "")),
            disc_iso=iso_date(row.get("Date of discovery end", "")),
            status="Ready to Deliver",
            disc_date=disc,
            deliver_date=deliv,
            internal=internal,
            links=get_links(row),
            desc_gtm=clean_desc(row.get("Desc GTM", "")),
            desc_product=clean_desc(row.get("Desc Product", "")),
            desc_client=clean_desc(row.get("Desc Client", "")),
            marketing=clean_bool(row.get("Is significant for separate marketing release?", False)),
            pilot_date=fmt_date(row.get("Pilot date", "") or ""),
            rollout_date=fmt_date(row.get("Full rollout date", "") or ""),
            stage=str(row.get("Current stage", "")).strip(),
            responsible=row.get("Responsible for discovery", ""))

    # ── Features in discovery ────────────────────────────────────────────
    df = sheet_to_df(spreadsheet, SHEET_TABS["discovery"])
    for _, row in df.iterrows():
        epic = str(row.get("Feature") or row.get("Epic", "")).strip()
        if not epic or epic == "nan":
            continue
        internal = str(row.get("Client-facing?", "Yes")).strip().lower() in ("no", "false", "0")
        stage = str(row.get("Current stage", "")).strip()
        status = normalize_status(stage)

        disc  = fmt_date(row.get("Date of discovery end") or row.get("Expected date of ready of discovery", ""))
        deliv = fmt_date(row.get("Date of delivery", ""))

        add(epic,
            area=str(row.get("Area", "Platform foundations")).strip(),
            connected=row.get("Connected features", "") or row.get("Blocking features", ""),
            deliver_iso=iso_date(row.get("Date of delivery", "")),
            disc_iso=iso_date(row.get("Date of discovery end", "")),
            status=status,
            disc_date=disc,
            deliver_date=deliv,
            internal=internal,
            links=get_links(row),
            desc_gtm=clean_desc(row.get("Desc GTM", "")),
            desc_product=clean_desc(row.get("Desc Product", "")),
            desc_client=clean_desc(row.get("Desc Client", "")),
            marketing=clean_bool(row.get("Is significant for separate marketing release?", False)),
            pilot_date=fmt_date(row.get("Pilot date", "") or ""),
            rollout_date=fmt_date(row.get("Full rollout date", "") or ""),
            stage=stage,
            responsible=row.get("Responsible for discovery", ""))

    # ── New features for discovery ───────────────────────────────────────
    df = sheet_to_df(spreadsheet, SHEET_TABS["new"])
    for _, row in df.iterrows():
        feat = str(row.get("Feature", "")).strip()
        if not feat or feat == "nan":
            continue
        internal = str(row.get("Client-facing?", "Yes")).strip().lower() in ("no", "false", "0")
        disc  = fmt_date(row.get("Date of discovery end") or row.get("Date of discovery") or row.get("Expected month for the discovery", ""))
        deliv = fmt_date(row.get("Date of delivery") or row.get("Expected month for the delivery", ""))
        add(feat,
            area=str(row.get("Area", "Platform foundations")).strip(),
            connected=row.get("Connected features", "") or row.get("Blocking features", ""),
            deliver_iso=iso_date(row.get("Date of delivery", "")),
            disc_iso=iso_date(row.get("Date of discovery end", "")),
            status="Planned",
            disc_date=disc,
            deliver_date=deliv,
            internal=internal,
            links=[],
            desc_gtm=clean_desc(row.get("Desc GTM", "")),
            desc_product=clean_desc(row.get("Desc Product", "")),
            desc_client=clean_desc(row.get("Desc Client", "")),
            marketing=clean_bool(row.get("Is significant for separate marketing release?", False)),
            pilot_date="",  # planned features: no separate pilot milestone
            rollout_date=fmt_date(row.get("Pilot date") or row.get("Full rollout date", "")))  # Pilot date = release date

    return features


# ═══════════════════════════════════════════════════════════════════════════
# HTML BUILDERS
# ═══════════════════════════════════════════════════════════════════════════

def esc(s):
    return str(s).replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")

def build_card(f):
    title = f["title"]
    area  = f["area"]
    color_var, area_label = AREA_CONFIG.get(area, ("--platform", area))
    # Priority: Sheets columns > DESCRIPTIONS dict > DEFAULT_DESC
    _desc = DESCRIPTIONS.get(title, DEFAULT_DESC)
    gtm     = f.get("desc_gtm")     or _desc[0]
    product = f.get("desc_product") or _desc[1]
    client  = f.get("desc_client")  or _desc[2]
    # sheet_status from Sheets overrides PROGRESS dict for In Delivery features
    sheet_status = f.get("sheet_status", "")
    SHEET_STATUS_PROGRESS = {
        "Fully shipped":    (100, "Fully shipped",    "#22c55e"),
        "Partially shipped":(90,  "Partially shipped", "#22c55e"),
        "In development":   (70,  "In development",    "#22c55e"),
        "Ready for delivery":(50, "Ready to deliver",  "#f59e0b"),
    }
    if sheet_status in SHEET_STATUS_PROGRESS:
        prog_data = SHEET_STATUS_PROGRESS[sheet_status]
    else:
        prog_data = PROGRESS.get(title)
    prog, stage, pcol = prog_data if prog_data else (None, None, None)
    preview = PREVIEWS.get(title)
    purl, popen = preview if preview else (None, None)
    internal = f.get("internal", False)

    badge = {
        "In Delivery":     "badge-delivery",
        "Ready to Deliver":"badge-ready",
        "Shipping":        "badge-delivery",
        "Discovery":       "badge-discovery",
        "Planned":         "badge-planned",
    }.get(f["status"], "badge-planned")
    badge_label = {
        "In Delivery":     "In Delivery",
        "Ready to Deliver":"Ready to Deliver",
        "Shipping":        "In Delivery",
        "Discovery":       "Discovery",
        "Planned":         "Planned",
    }.get(f["status"], f["status"])
    if f.get("_past"):
        badge_label = "Shipped"
    marketing = f.get("marketing", False)
    int_attr  = ' data-internal="true"' if internal else ""
    mkt_attr  = ' data-marketing="true"' if marketing else ""
    prev_attr = f'\n           data-preview-url="{purl}"\n           data-preview-open="{popen}"' if purl else ""

    preview_html = ""
    if purl:
        preview_html = f"""        <div class="card-preview">
          <img alt="{esc(title)} preview" />
          <div class="card-preview-overlay"></div>
          <a class="card-preview-open" href="{popen}" target="_blank">Open prototype ↗</a>
        </div>\n"""

    prog_html = ""
    if prog is not None:
        prog_html = f"""<div class="card-progress">
          <div class="progress-header">
            <span class="progress-stage">{stage}</span>
            <span class="progress-pct">{prog}%</span>
          </div>
          <div class="progress-track"><div class="progress-fill" style="width:{prog}%;background:{pcol}"></div></div>
        </div>"""

    d_parts = []
    if f.get("disc_date"):
        d_parts.append(f'<span class="date-badge date-disc">🔍 {f["disc_date"]}</span>')
    if f.get("deliver_date"):
        d_parts.append(f'<span class="date-badge date-deliver">🚀 {f["deliver_date"]}</span>')
    if f.get("pilot_date") and str(f["pilot_date"]).lower() not in ("", "nan", "none"):
        d_parts.append(f'<span class="date-badge date-pilot">🧪 Pilot: {f["pilot_date"]}</span>')
    if f.get("rollout_date") and str(f["rollout_date"]).lower() not in ("", "nan", "none"):
        d_parts.append(f'<span class="date-badge date-rollout">🌐 Rollout: {f["rollout_date"]}</span>')
    dates_html = f'<div class="card-dates">{"".join(d_parts)}</div>' if d_parts else ""

    meta_html = f'\n        <div class="card-meta">{prog_html}{dates_html}</div>' if (prog_html or dates_html) else ""

    lnk_html = ""
    one_pager = ONE_PAGERS.get(title, "")
    link_items = []

    # One-pager button always comes first if available
    if one_pager:
        link_items.append(
            f'<a class="lp ln" href="#" '
            f'onclick="openOnePager(\'{one_pager}\', \'{esc(title)}\'); return false;">'
            f'One-pager</a>'
        )

    if f.get("links"):
        for lbl, cls, url in f["links"]:
            if lbl == "Description" and one_pager:
                continue  # skip Description if one-pager already added
            link_items.append(f'<a class="lp {cls}" href="{url}" target="_blank">{lbl}</a>')

    if link_items:
        items = "\n            ".join(link_items)
        lnk_html = f'\n          <div class="card-links">\n            {items}\n          </div>'

    mkt_badge_html = '<span class="mkt-badge" title="Significant for marketing release">✦</span>' if marketing else ""

    return f"""      <div class="card" data-area="{area}"{prev_attr}{int_attr}{mkt_attr}>
<button class="card-custom-toggle" title="Hide from custom view"></button>
{preview_html}        <div class="card-content">
          <div class="card-top"><div class="area-dot" style="background:var({color_var})"></div><span class="area-label" style="color:var({color_var})">{area_label.upper()}</span><span class="status-badge {badge}">{badge_label}</span>{mkt_badge_html}</div>
          <div class="card-title">{esc(title)}</div>
          <div class="card-desc desc-gtm">{esc(gtm)}</div>
          <div class="card-desc desc-product">{esc(product)}</div>
          <div class="card-desc desc-client">{esc(client)}</div>{meta_html}{lnk_html}
        </div>
      </div>"""

def build_col(css_cls, month, sub, cards, past=False):
    body = "\n\n".join(cards)
    past_cls = " col-past" if past else ""
    past_sub = "Shipped" if past else sub
    return f"""  <div class="{css_cls}{past_cls}">
    <div class="col-header"><div class="month">{month}</div><div class="subtitle">{past_sub}</div></div>
    <div class="col-body">

{body}

    </div>
  </div>"""

def build_sum_chip(f):
    color_var, _ = AREA_CONFIG.get(f["area"], ("--platform", f["area"]))
    int_attr = ' data-internal="true"' if f.get("internal") else ""
    mkt_attr = ' data-marketing="true"' if f.get("marketing") else ""
    return f'        <div class="sum-chip" data-area="{f["area"]}"{int_attr}{mkt_attr}><div class="sum-dot" style="background:var({color_var})"></div><span class="sum-name">{esc(f["title"])}</span></div>'

def build_sum_col(css_cls, month, sub, chips):
    body = "\n".join(chips)
    return f"""    <div class="{css_cls}">
      <div class="sum-col-header"><div class="month">{month}</div><div class="subtitle">{sub}</div></div>
      <div class="sum-col-body">
{body}
      </div>
    </div>"""


# ═══════════════════════════════════════════════════════════════════════════
# GENERATE
# ═══════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════
# WORK BOARD (Kanban) — board.html
# ═══════════════════════════════════════════════════════════════════════════

def board_col_of(f):
    """Map a feature to a board column."""
    st = f.get("status")
    ss = str(f.get("sheet_status", "")).strip().lower()
    if st == "In Delivery":
        return "shipped" if "shipped" in ss else "delivery"
    if st == "Ready to Deliver":
        return "ready"
    if st == "Discovery":
        return "discovery"
    return "planned"

def _board_date_key(f, field):
    import datetime as _dt
    v = f.get(field)
    try:
        return _dt.datetime.strptime(v, "%b %Y") if v else _dt.datetime.max
    except Exception:
        return _dt.datetime.max

def _is_overdue(date_str):
    """True if the 'Mon YYYY' month is strictly before the current month."""
    import datetime as _dt
    if not date_str:
        return False
    try:
        d = _dt.datetime.strptime(str(date_str), "%b %Y").date().replace(day=1)
    except Exception:
        return False
    today_m = _dt.date.today().replace(day=1)
    return d < today_m

def _date_state(date_str):
    """Classify a 'Mon YYYY' milestone: done / now / planned / tbd."""
    import datetime as _dt
    if not date_str or str(date_str).strip().lower() in ("", "nan", "none"):
        return "tbd"
    try:
        d = _dt.datetime.strptime(str(date_str), "%b %Y").date().replace(day=1)
    except Exception:
        return "tbd"
    cur = _dt.date.today().replace(day=1)
    if d < cur: return "done"
    if d == cur: return "now"
    return "planned"

def feature_detail_inner(f):
    """Inner HTML for the feature detail modal: timeline + descriptions + document links.
    Shared by the Board and Weekly pages."""
    desc_rows = []
    for lbl, key in (("Product", "desc_product"), ("GTM", "desc_gtm"), ("Client", "desc_client")):
        v = clean_desc(f.get(key, ""))
        if v:
            desc_rows.append(f'<div class="md-desc"><h4>{lbl}</h4><p>{esc(v)}</p></div>')
    desc_html = "".join(desc_rows) or '<div class="md-empty">No description in the sheet yet.</div>'

    phases = [("Discovery", f.get("disc_date")), ("Delivery", f.get("deliver_date"))]
    if f.get("pilot_date") and str(f.get("pilot_date")).strip().lower() not in ("", "nan", "none"):
        phases.append(("Pilot", f.get("pilot_date")))
    phases.append(("Release", f.get("rollout_date")))
    cells = []
    for _lbl, _dv in phases:
        _st = _date_state(_dv)
        _dt_txt = esc(_dv) if (_dv and str(_dv).strip().lower() not in ("", "nan", "none")) else "TBD"
        cells.append(f'<div class="md-phase {_st}"><div class="md-phase-dot"></div>'
                     f'<div class="md-phase-label">{_lbl}</div>'
                     f'<div class="md-phase-date">{_dt_txt}</div></div>')
    timeline_html = f'<div class="md-timeline-label">Planned timeline</div><div class="md-timeline">{"".join(cells)}</div>'

    link_items = [f'<a class="md-link" href="{esc(url)}" target="_blank" rel="noopener">{esc(label)} ↗</a>'
                  for (label, css, url) in (f.get("links") or [])]
    links_html = (f'<div class="md-links-label">Documents</div><div class="md-links">{"".join(link_items)}</div>'
                  if link_items else '<div class="md-links-label">Documents</div><div class="md-empty">No linked documents yet.</div>')
    return timeline_html + desc_html + links_html

def build_board_card(f):
    area = f["area"]
    color_var, area_label = AREA_CONFIG.get(area, ("--platform", area))
    col = board_col_of(f)
    can_overdue = col not in ("shipped",)
    badges = []
    if f.get("disc_date"):
        disc_over = can_overdue and col == "discovery" and _is_overdue(f["disc_date"])
        cls = "bbadge disc overdue" if disc_over else "bbadge disc"
        pre = "⚠ " if disc_over else ""
        badges.append(f'<span class="{cls}">{pre}🔍 disc {esc(f["disc_date"])}</span>')
    if f.get("deliver_date"):
        deliv_over = can_overdue and _is_overdue(f["deliver_date"])
        cls = "bbadge deadline overdue" if deliv_over else "bbadge deadline"
        pre = "⚠ " if deliv_over else ""
        badges.append(f'<span class="{cls}">{pre}🚀 {esc(f["deliver_date"])}</span>')
    if f.get("pilot_date") and str(f["pilot_date"]).lower() not in ("", "nan", "none"):
        badges.append(f'<span class="bbadge pilot">🧪 {esc(f["pilot_date"])}</span>')
    if f.get("rollout_date") and str(f["rollout_date"]).lower() not in ("", "nan", "none"):
        badges.append(f'<span class="bbadge rollout">🌐 {esc(f["rollout_date"])}</span>')
    if f.get("marketing"):
        badges.append('<span class="bbadge mkt">✦ Marketing</span>')
    badges_html = f'<div class="bmeta">{"".join(badges)}</div>' if badges else ""
    stage_txt = f.get("stage") or f.get("sheet_status") or ""
    stage_html = f'<div class="bstage">{esc(stage_txt)}</div>' if stage_txt and str(stage_txt).lower() not in ("nan", "none") else ""
    resp = f.get("responsible")
    resp_html = f'<div class="bresp">👤 {esc(resp)}</div>' if resp and str(resp).lower() not in ("nan", "none", "") else ""
    mkt_attr = ' data-mkt="1"' if f.get("marketing") else ""
    card_over = can_overdue and _is_overdue(f.get("deliver_date"))
    over_attr = ' data-overdue="1"' if card_over else ""
    over_cls = " overdue-card" if card_over else ""

    detail_html = f'<div class="bcard-detail" hidden>{feature_detail_inner(f)}</div>'

    return (
        f'      <div class="bcard{over_cls}" data-area="{esc(area)}"{mkt_attr}{over_attr} style="border-left-color:var({color_var})">\n'
        f'        <div class="bcard-area"><span class="dot" style="background:var({color_var})"></span><span style="color:var({color_var})">{esc(area_label)}</span></div>\n'
        f'        <div class="bcard-title">{esc(f["title"])}</div>\n'
        f'        {stage_html}{badges_html}{resp_html}\n'
        f'        {detail_html}\n'
        f'      </div>'
    )

def build_board_col(col_id, title, dot_color, cards):
    inner = "\n".join(cards) if cards else '      <div class="bempty">No features here yet</div>'
    return (
        f'    <div class="bcol" data-col="{col_id}">\n'
        f'      <div class="bcol-head">\n'
        f'        <div class="bcol-title"><span class="bcol-dot" style="background:{dot_color}"></span>{title}</div>\n'
        f'        <span class="bcol-count">{len(cards)}</span>\n'
        f'      </div>\n'
        f'{inner}\n'
        f'    </div>'
    )

def generate_board(features):
    from pathlib import Path as _P
    cols = {"planned": [], "discovery": [], "ready": [], "delivery": [], "shipped": []}
    for f in features:
        cols[board_col_of(f)].append(f)
    cols["discovery"].sort(key=lambda f: _board_date_key(f, "disc_date"))
    cols["ready"].sort(key=lambda f: _board_date_key(f, "deliver_date"))
    cols["delivery"].sort(key=lambda f: _board_date_key(f, "deliver_date"))
    cols["planned"].sort(key=lambda f: _board_date_key(f, "deliver_date"))

    COLDEF = [
        ("planned",   "Planned",            "#6366f1"),
        ("discovery", "Discovery",          "#f59e0b"),
        ("ready",     "Ready for delivery", "#eab308"),
        ("delivery",  "In delivery",        "#22c55e"),
        ("shipped",   "Shipped",            "#94a3b8"),
    ]
    cols_html = "\n".join(
        build_board_col(cid, title, dot, [build_board_card(f) for f in cols[cid]])
        for cid, title, dot in COLDEF
    )
    board_html = f'<div class="board" id="boardView">\n{cols_html}\n</div><!-- /board -->'

    base_path = _P(__file__).parent / "board.html"
    if not base_path.exists():
        print("  ⚠ board.html template not found — skipping board")
        return
    base = open(base_path, encoding="utf-8").read()
    s = base.index('<div class="board" id="boardView">')
    e = base.index('</div><!-- /board -->') + len('</div><!-- /board -->')
    base = base[:s] + board_html + base[e:]
    base = re.sub(r'Updated \w+ \d{4}', f'Updated {datetime.now().strftime("%B %Y")}', base)
    base = inject_filters(base, features, include_marketing=True)
    open(base_path, "w", encoding="utf-8").write(base)
    print(f"✓  Saved board → {base_path}  " + str({k: len(v) for k, v in cols.items()}))


# ═══════════════════════════════════════════════════════════════════════════
# PRODUCT MAP (graph of features + connections) — map.html
# ═══════════════════════════════════════════════════════════════════════════

def _norm_title(s):
    return re.sub(r"\s+", " ", str(s or "").strip().lower())

def build_graph(features):
    """Build {nodes, links} from features + their 'Connected features' column."""
    nodes = []
    title_index = {}
    norm_titles = []
    for i, f in enumerate(features):
        color_var, _ = AREA_CONFIG.get(f["area"], ("--platform", f["area"]))
        nodes.append({"id": i, "title": f["title"], "area": f["area"], "color": color_var})
        nt = _norm_title(f["title"])
        title_index[nt] = i
        norm_titles.append((nt, i))

    def resolve(token):
        t = _norm_title(token)
        if not t or len(t) < 3:
            return None
        if t in title_index:
            return title_index[t]
        cands = [(len(nt), i) for nt, i in norm_titles if t in nt]   # token is substring of a title
        if cands:
            cands.sort()        # shortest (most specific) title wins
            return cands[0][1]
        return None

    links = set()
    for i, f in enumerate(features):
        raw = f.get("connected", "")
        if not raw or str(raw).strip().lower() in ("", "nan", "none"):
            continue
        for tok in re.split(r"[,;\n]", str(raw)):
            tok = tok.strip()
            if not tok:
                continue
            j = resolve(tok)
            if j is not None and j != i:
                links.add(tuple(sorted((i, j))))
    links = [{"source": a, "target": b} for a, b in sorted(links)]
    return {"nodes": nodes, "links": links}

def generate_map(features):
    from pathlib import Path as _P
    import json
    graph = build_graph(features)
    base_path = _P(__file__).parent / "map.html"
    if not base_path.exists():
        print("  ⚠ map.html template not found — skipping map")
        return
    base = open(base_path, encoding="utf-8").read()
    start_tag = '<script id="graphData" type="application/json">'
    if start_tag not in base:
        print("  ⚠ graphData marker not found in map.html — skipping map")
        return
    s = base.index(start_tag) + len(start_tag)
    e = base.index('</script>', s)
    payload = json.dumps(graph, ensure_ascii=False)
    base = base[:s] + "\n" + payload + "\n" + base[e:]
    base = re.sub(r'Updated \w+ \d{4}', f'Updated {datetime.now().strftime("%B %Y")}', base)
    base = inject_filters(base, features, include_marketing=False)
    open(base_path, "w", encoding="utf-8").write(base)
    print(f"✓  Saved map → {base_path}  nodes={len(graph['nodes'])} links={len(graph['links'])}")


# ═══════════════════════════════════════════════════════════════════════════
# WEEKLY PLANNING — Discovery / Delivery Gantt
# A feature occupies a lane for the 30 days before its planned completion date
# (Discovery → Date of discovery end; Delivery → Date of delivery).
# ═══════════════════════════════════════════════════════════════════════════

WEEKLY_WINDOW_DAYS = {"discovery": 7, "delivery": 30, "finteam": 30}  # days before completion, per lane
FIN_TEAM_AREAS = {"Agents (10 in MVP)", "Shared context (16 layers)", "Workflows"}  # shown only in the Fin team work lane

def generate_weekly(features):
    import datetime as _dt
    from pathlib import Path as _P

    def monday(d):
        return d - _dt.timedelta(days=d.weekday())

    lanes = {"discovery": [], "delivery": [], "finteam": []}
    def add_bar(lane, f, end_iso):
        try:
            end = _dt.date.fromisoformat(end_iso)
        except Exception:
            return
        span = max(1, round(WEEKLY_WINDOW_DAYS[lane] / 7))
        endMon = monday(end)
        lanes[lane].append({"f": f, "startMon": endMon - _dt.timedelta(weeks=span - 1), "endMon": endMon})
    for f in features:
        if f.get("area") in FIN_TEAM_AREAS:
            # Fin team work: only this lane, positioned by delivery date
            if f.get("deliver_iso"): add_bar("finteam", f, f["deliver_iso"])
            continue
        if f.get("disc_iso"):    add_bar("discovery", f, f["disc_iso"])
        if f.get("deliver_iso"): add_bar("delivery",  f, f["deliver_iso"])

    all_bars = lanes["discovery"] + lanes["delivery"] + lanes["finteam"]
    if not all_bars:
        weekly_html = ('<div class="weekly" id="weeklyView">'
                       '<div class="gantt-empty">No dated features yet — add discovery/delivery dates in the sheet.</div>'
                       '</div><!-- /weekly -->')
    else:
        minMon = min(b["startMon"] for b in all_bars)
        maxMon = max(b["endMon"] for b in all_bars)
        weeks = []
        cur = minMon
        while cur <= maxMon:
            weeks.append(cur); cur += _dt.timedelta(days=7)
        idx_of = {m: i for i, m in enumerate(weeks)}
        N = len(weeks)
        def widx(d): return idx_of[monday(d)]

        def pack(bars):
            bars = sorted(bars, key=lambda b: (idx_of[b["startMon"]], idx_of[b["endMon"]]))
            row_end = []
            placed = []
            for b in bars:
                s, e = idx_of[b["startMon"]], idx_of[b["endMon"]]
                r = None
                for ri in range(len(row_end)):
                    if row_end[ri] < s:
                        r = ri; row_end[ri] = e; break
                if r is None:
                    row_end.append(e); r = len(row_end) - 1
                placed.append((b, s, e, r))
            return placed, max(1, len(row_end))

        def bar_html(b, s, e, r):
            f = b["f"]; cv, alabel = AREA_CONFIG.get(f["area"], ("--platform", f["area"]))
            return (f'<div class="bar" data-area="{esc(f["area"])}" data-title="{esc(f["title"])}" '
                    f'data-color="{cv}" data-arealabel="{esc(alabel)}" title="{esc(f["title"])}" '
                    f'style="grid-column:{s+1}/{e+2};grid-row:{r+1};border-left-color:var({cv})">'
                    f'<span class="bar-dot" style="background:var({cv})"></span>{esc(f["title"])}'
                    f'<div class="bcard-detail" hidden>{feature_detail_inner(f)}</div></div>')

        def grid_html(bars):
            placed, nrows = pack(bars)
            cells = "".join(bar_html(b, s, e, r) for (b, s, e, r) in placed)
            return (f'<div class="gantt-grid" style="grid-template-columns:repeat({N},var(--gw));'
                    f'grid-template-rows:repeat({nrows},auto)">{cells}</div>')

        heads = []
        for m in weeks:
            su = m + _dt.timedelta(days=6)
            lbl = (f'{m.strftime("%-d")}–{su.strftime("%-d %b")}' if m.month == su.month
                   else f'{m.strftime("%-d %b")}–{su.strftime("%-d %b")}')
            heads.append(f'<div class="gw">{lbl}<small>W{m.isocalendar()[1]}</small></div>')
        head_html = f'<div class="gantt-weeks" style="grid-template-columns:repeat({N},var(--gw))">{"".join(heads)}</div>'

        weekly_html = (
            '<div class="weekly" id="weeklyView"><div class="gantt">'
            f'<div class="gantt-row gantt-head"><div class="gantt-lane-label"></div>{head_html}</div>'
            f'<div class="gantt-row gantt-disc"><div class="gantt-lane-label disc">DISCOVERY</div>{grid_html(lanes["discovery"])}</div>'
            f'<div class="gantt-row gantt-deliv"><div class="gantt-lane-label deliv">DELIVERY</div>{grid_html(lanes["delivery"])}</div>'
            f'<div class="gantt-row gantt-fin"><div class="gantt-lane-label fin">FIN TEAM WORK</div>{grid_html(lanes["finteam"])}</div>'
            '</div></div><!-- /weekly -->'
        )

    base_path = _P(__file__).parent / "weekly.html"
    if not base_path.exists():
        print("  ⚠ weekly.html template not found — skipping weekly")
        return
    base = open(base_path, encoding="utf-8").read()
    sidx = base.index('<div class="weekly" id="weeklyView">')
    eidx = base.index('</div><!-- /weekly -->') + len('</div><!-- /weekly -->')
    base = base[:sidx] + weekly_html + base[eidx:]
    base = re.sub(r'Updated \w+ \d{4}', f'Updated {datetime.now().strftime("%B %Y")}', base)
    base = inject_filters(base, features, include_marketing=False)
    open(base_path, "w", encoding="utf-8").write(base)
    print(f"✓  Saved weekly → {base_path}  discovery={len(lanes['discovery'])} delivery={len(lanes['delivery'])}")

def generate(features, output_path):
    import datetime as _dt
    _today = _dt.date.today()
    _this_month = _today.replace(day=1)
    COL_MONTHS = {
        "february": _dt.date(2026, 5, 1),  # before July -> history
        "march":    _dt.date(2026, 7, 1),  # July 2026
        "april":    _dt.date(2026, 8, 1),  # August 2026
        "may":      _dt.date(2026, 9, 1),  # September 2026
        "june":     _dt.date(2026, 10, 1), # October 2026
        "july":     _dt.date(2026, 11, 1), # Later
    }
    COL_PAST = {col: month < _this_month for col, month in COL_MONTHS.items()}

    # Assign to columns
    cols = {"february": [], "march": [], "april": [], "may": [], "june": [], "july": []}
    for f in features:
        col = get_col(f.get("deliver_date"))
        if f["status"] == "Shipping" and not f.get("deliver_date"):
            col = "march"
        if COL_PAST[col]:
            f = dict(f, _past=True)
        cols[col].append(f)

    detailed_html = f"""<div class="timeline" id="detailedView">

{build_col('col-feb','Before July 2026','Shipped / earlier',            [build_card(f) for f in cols['february']], past=COL_PAST['february'])}

{build_col('col-now','July 2026','MVP delivery',               [build_card(f) for f in cols['march']], past=COL_PAST['march'])}

{build_col('col-apr','August 2026','MVP delivery',           [build_card(f) for f in cols['april']], past=COL_PAST['april'])}

{build_col('col-may','September 2026','MVP delivery',             [build_card(f) for f in cols['may']], past=COL_PAST['may'])}

{build_col('col-jun','October 2026','Backlog · subject to change',            [build_card(f) for f in cols['june']], past=COL_PAST['june'])}

{build_col('col-q3','Later','Backlog · subject to change', [build_card(f) for f in cols['july']], past=COL_PAST['july'])}

</div><!-- /detailed -->"""

    summary_html = f"""<div class="summary-view" id="summaryView">
  <div class="summary-grid">

{build_sum_col('sum-col-feb','Before July 2026','Shipped / earlier',            [build_sum_chip(f) for f in cols['february']])}

{build_sum_col('sum-col-now','July 2026','MVP delivery',               [build_sum_chip(f) for f in cols['march']])}

{build_sum_col('sum-col-apr','August 2026','MVP delivery',           [build_sum_chip(f) for f in cols['april']])}

{build_sum_col('sum-col-may','September 2026','MVP delivery',                [build_sum_chip(f) for f in cols['may']])}

{build_sum_col('sum-col-jun','October 2026','Backlog',               [build_sum_chip(f) for f in cols['june']])}

{build_sum_col('sum-col-q3','Later','Backlog · subject to change', [build_sum_chip(f) for f in cols['july']])}

  </div>
</div>"""

    # Read existing HTML and replace timeline sections
    base_path = Path(__file__).parent / "index.html"
    if not base_path.exists():
        print(f"ERROR: Template not found: {base_path}")
        print("Place index.html in the same folder as this script.")
        sys.exit(1)

    # Debug: show column distribution
    from collections import Counter
    dist = Counter()
    for feat in features:
        col = get_col(feat.get("deliver_date"))
        if feat["status"] in ("In Delivery", "Shipping") and not feat.get("deliver_date"):
            col = "march"  # no-date features default to first active month
        dist[col] += 1
    print("Column distribution:", dict(dist))
    sample = [(f["title"][:30], f.get("deliver_date"), get_col(f.get("deliver_date")))
              for f in features[:5]]
    print("Sample features:", sample)

    with open(base_path, encoding="utf-8") as fh:
        base = fh.read()

    # Replace detailed view
    s = base.index('<div class="timeline" id="detailedView">')
    e = base.index('</div><!-- /detailed -->') + len('</div><!-- /detailed -->')
    base = base[:s] + detailed_html + base[e:]

    # Replace summary view
    s = base.index('<div class="summary-view" id="summaryView">')
    # Find end: closing </div> before <script>
    e = base.index('\n\n\n<script>', s)
    base = base[:s] + summary_html + base[e:]

    # Update date in header
    today = datetime.now().strftime("%B %Y")
    base = re.sub(r'Updated \w+ \d{4}', f'Updated {today}', base)

    base = inject_filters(base, features, include_marketing=True)
    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write(base)

    print(f"\n✓  Saved → {output_path}")
    total = sum(len(v) for v in cols.values())
    for col_name, feats in cols.items():
        print(f"   {col_name.capitalize():8}  {len(feats):2} features")
    print(f"   {'Total':8}  {total:2} features")


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="Generate Fuel roadmap from Google Sheets.")
    parser.add_argument("--sheet-id", default=SHEET_ID,   help="Google Sheet ID")
    parser.add_argument("--key",      default=KEY_FILE,   help="Path to service_account.json")
    parser.add_argument("--output",   default="index.html", help="Output HTML file")
    args = parser.parse_args()

    key_path = Path(args.key)
    if not key_path.exists():
        print(f"ERROR: Service account key not found: {key_path}")
        print()
        print("One-time setup:")
        print("  1. Go to https://console.cloud.google.com/")
        print("  2. Enable 'Google Sheets API' and 'Google Drive API'")
        print("  3. Create a Service Account → download JSON key")
        print(f"  4. Save the key as: {key_path.resolve()}")
        print(f"  5. Share your Google Sheet with the service account email (Viewer)")
        sys.exit(1)

    print(f"Connecting to Google Sheets ({args.sheet_id[:20]}...)...")
    try:
        spreadsheet = connect_sheets(str(key_path), args.sheet_id)
        print(f"✓ Connected: {spreadsheet.title}")
    except Exception as e:
        print(f"ERROR connecting to Google Sheets: {e}")
        sys.exit(1)

    print("Reading features...")
    features = read_features(spreadsheet)
    print(f"✓ Found {len(features)} unique features")

    output = Path(args.output)
    generate(features, output)
    generate_board(features)
    generate_map(features)
    generate_weekly(features)

if __name__ == "__main__":
    main()
