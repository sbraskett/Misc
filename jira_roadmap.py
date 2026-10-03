#!/usr/bin/env python3
"""
jira_roadmap.py
Generate a single-file HTML product roadmap from Jira Server/Data Center.

Authentication:
    Environment variable JIRA_PAT is required.
    The PAT is sent as: Authorization: Bearer <token>

Install:
    pip install requests

Run:
    Windows PowerShell:
        $env:JIRA_PAT = "your-token"
        python jira_roadmap.py

    macOS/Linux:
        export JIRA_PAT="your-token"
        python jira_roadmap.py

IMPORTANT:
1. Set JIRA_BASE_URL and JQL below.
2. Review FIELD_NAMES / FIELD_IDS. The script auto-discovers custom fields by name.
3. Configure parent linkage if Features are children of Epics in your Jira hierarchy.
"""

from __future__ import annotations

import html
import json
import os
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any

import requests
import urllib3

# =============================================================================
# CONFIGURATION
# =============================================================================

JIRA_BASE_URL = "https://jira.yourcompany.com"
JQL = 'issuetype in (Epic, Feature) ORDER BY project, created'

OUTPUT_FILE = "jira_product_roadmap.html"

# Jira Server/Data Center commonly exposes the v2 endpoint.
API_VERSION = "2"
PAGE_SIZE = 100

# Set False only if your on-prem Jira uses a private/self-signed certificate and
# you cannot supply a CA bundle. Prefer a CA bundle in production.
VERIFY_SSL: bool | str = True
# Example:
# VERIFY_SSL = r"C:\certs\company-ca.pem"

EPIC_ISSUE_TYPE = "Epic"
FEATURE_ISSUE_TYPE = "Feature"

# Fiscal/calendar display year. None = choose from issue dates/current year.
ROADMAP_YEAR: int | None = None

# Auto-discovery uses these exact Jira field names, case-insensitively.
# Change the values if your Jira names differ.
FIELD_NAMES = {
    "health": "Health",
    "value_statement": "Value Statement",
    "additional_notes": "Additional Notes",
    "start_date": "Start Date",
    "target_date": "Target Date",
    "percent_complete": "Percent Complete",
    # Configure ONE of these for Epic -> Feature relationship if applicable.
    "parent_link": "Parent Link",
    "epic_link": "Epic Link",
}

# Optional explicit IDs. These override name discovery.
# Example: "health": "customfield_12345"
FIELD_IDS = {
    "health": None,
    "value_statement": None,
    "additional_notes": None,
    "start_date": None,
    "target_date": None,
    "percent_complete": None,
    "parent_link": None,
    "epic_link": None,
}

# Health values are normalized case-insensitively.
HEALTH_MAP = {
    "green": "green",
    "g": "green",
    "on track": "green",
    "yellow": "yellow",
    "amber": "yellow",
    "y": "yellow",
    "at risk": "yellow",
    "red": "red",
    "r": "red",
    "off track": "red",
}

# If no Percent Complete field exists, calculate completion from child issue
# status categories if issue children are present in the result.
CALCULATE_PERCENT_IF_MISSING = True

# Rich text from Jira Server/DC may be wiki markup or HTML depending on field.
# For safety this script converts text to escaped HTML and recognizes simple
# bullets. Set True only if these custom fields are trusted HTML produced by Jira.
TRUST_RICH_TEXT_HTML = False

# =============================================================================
# JIRA CLIENT
# =============================================================================

def jira_session() -> requests.Session:
    token = os.environ.get("JIRA_PAT")
    if not token:
        sys.exit("ERROR: environment variable JIRA_PAT is not set.")

    s = requests.Session()
    s.headers.update({
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    })
    if VERIFY_SSL is False:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    return s


def api_url(path: str) -> str:
    return f"{JIRA_BASE_URL.rstrip('/')}{path}"


def get_fields(session: requests.Session) -> list[dict[str, Any]]:
    """GET /rest/api/2/field works across many Jira Server/DC versions."""
    r = session.get(
        api_url(f"/rest/api/{API_VERSION}/field"),
        timeout=60,
        verify=VERIFY_SSL,
    )
    r.raise_for_status()
    return r.json()


def discover_field_ids(session: requests.Session) -> dict[str, str | None]:
    fields = get_fields(session)
    by_name = {str(f.get("name", "")).strip().lower(): f.get("id") for f in fields}

    resolved: dict[str, str | None] = {}
    print("\nField mapping:")
    for logical_name, jira_name in FIELD_NAMES.items():
        explicit = FIELD_IDS.get(logical_name)
        field_id = explicit or by_name.get(jira_name.lower())
        resolved[logical_name] = field_id
        print(f"  {logical_name:18} {jira_name:22} -> {field_id or 'NOT FOUND'}")
    return resolved


def search_issues(
    session: requests.Session,
    field_ids: dict[str, str | None],
) -> list[dict[str, Any]]:
    fields = [
        "summary", "issuetype", "project", "status", "assignee",
        "created", "updated", "resolutiondate", "issuelinks",
        "parent",
    ]
    fields += [x for x in field_ids.values() if x]
    fields = list(dict.fromkeys(fields))

    all_issues: list[dict[str, Any]] = []
    start_at = 0

    while True:
        payload = {
            "jql": JQL,
            "startAt": start_at,
            "maxResults": PAGE_SIZE,
            "fields": fields,
        }
        r = session.post(
            api_url(f"/rest/api/{API_VERSION}/search"),
            json=payload,
            timeout=120,
            verify=VERIFY_SSL,
        )
        if not r.ok:
            print(f"\nJira search failed: HTTP {r.status_code}", file=sys.stderr)
            print(r.text[:2000], file=sys.stderr)
            r.raise_for_status()

        data = r.json()
        batch = data.get("issues", [])
        all_issues.extend(batch)
        print(f"\rRetrieved {len(all_issues)} issue(s)...", end="", flush=True)

        if not batch:
            break

        start_at += len(batch)
        total = data.get("total")
        if total is not None and start_at >= total:
            break
        if len(batch) < data.get("maxResults", PAGE_SIZE):
            break

    print()
    return all_issues


# =============================================================================
# NORMALIZATION
# =============================================================================

def field_value(issue: dict, field_id: str | None, default=None):
    if not field_id:
        return default
    return issue.get("fields", {}).get(field_id, default)


def scalar(value: Any) -> Any:
    """Turn common Jira select/user/custom-field JSON shapes into a scalar."""
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, list):
        return ", ".join(str(scalar(v)) for v in value if scalar(v) is not None)
    if isinstance(value, dict):
        for key in ("value", "name", "displayName", "key"):
            if value.get(key) is not None:
                return value[key]
    return str(value)


def normalize_health(value: Any) -> str:
    v = str(scalar(value) or "").strip().lower()
    return HEALTH_MAP.get(v, "unknown")


def parse_pct(value: Any) -> float | None:
    value = scalar(value)
    if value is None or value == "":
        return None
    try:
        n = float(str(value).replace("%", "").strip())
        if 0 <= n <= 1 and "." in str(value):
            n *= 100
        return max(0, min(100, n))
    except (TypeError, ValueError):
        return None


def parse_date(value: Any) -> date | None:
    value = scalar(value)
    if not value:
        return None
    text = str(value)[:10]
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None


def user_name(value: Any) -> str:
    if not value:
        return ""
    if isinstance(value, dict):
        return value.get("displayName") or value.get("name") or ""
    return str(value)


def parent_key(issue: dict, field_ids: dict[str, str | None]) -> str | None:
    """Try standard parent, Parent Link, then Epic Link."""
    f = issue.get("fields", {})

    parent = f.get("parent")
    if isinstance(parent, dict) and parent.get("key"):
        return parent["key"]

    for logical in ("parent_link", "epic_link"):
        v = field_value(issue, field_ids.get(logical))
        if isinstance(v, dict):
            if v.get("key"):
                return v["key"]
            if v.get("value"):
                return str(v["value"])
        if isinstance(v, str) and v.strip():
            # Parent Link sometimes returns a key or a display string containing it.
            m = re.search(r"\b[A-Z][A-Z0-9_]+-\d+\b", v)
            return m.group(0) if m else v.strip()
    return None


def rich_to_html(value: Any) -> str:
    """Safe, lightweight renderer for Jira text/wiki-ish content."""
    if value is None:
        return "<span class='empty'>No content provided.</span>"

    if isinstance(value, (dict, list)):
        # Some apps/custom fields return structured JSON.
        text = json.dumps(value, ensure_ascii=False, indent=2)
    else:
        text = str(value)

    if TRUST_RICH_TEXT_HTML:
        return text

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    out: list[str] = []
    bullets: list[str] = []

    def flush_bullets():
        nonlocal bullets
        if bullets:
            out.append("<ul>" + "".join(f"<li>{b}</li>" for b in bullets) + "</ul>")
            bullets = []

    for raw in lines:
        line = raw.strip()
        if not line:
            flush_bullets()
            continue
        # Common Jira wiki bullets: *, -, #
        if re.match(r"^[*#-]\s+", line):
            cleaned = re.sub(r"^[*#-]\s+", "", line)
            bullets.append(html.escape(cleaned))
        else:
            flush_bullets()
            safe = html.escape(line)
            # Simple Jira wiki bold *text*
            safe = re.sub(r"\*([^*\n]+)\*", r"<strong>\1</strong>", safe)
            out.append(f"<p>{safe}</p>")
    flush_bullets()
    return "".join(out) or "<span class='empty'>No content provided.</span>"


def normalize_issue(issue: dict, ids: dict[str, str | None]) -> dict[str, Any]:
    f = issue.get("fields", {})
    project = f.get("project") or {}
    issue_type = f.get("issuetype") or {}
    status = f.get("status") or {}
    status_category = status.get("statusCategory") or {}

    return {
        "key": issue.get("key", ""),
        "summary": f.get("summary") or "(Untitled)",
        "issue_type": issue_type.get("name") or "",
        "project": project.get("key") or project.get("name") or "",
        "project_name": project.get("name") or project.get("key") or "",
        "health": normalize_health(field_value(issue, ids["health"])),
        "value_statement": rich_to_html(field_value(issue, ids["value_statement"])),
        "additional_notes": rich_to_html(field_value(issue, ids["additional_notes"])),
        "start": parse_date(field_value(issue, ids["start_date"])),
        "target": parse_date(field_value(issue, ids["target_date"])),
        "percent": parse_pct(field_value(issue, ids["percent_complete"])),
        "assignee": user_name(f.get("assignee")),
        "status": status.get("name") or "",
        "status_category": status_category.get("key") or "",
        "parent_key": parent_key(issue, ids),
        "updated": str(f.get("updated") or "")[:10],
    }


def build_hierarchy(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    epics = [x for x in items if x["issue_type"].lower() == EPIC_ISSUE_TYPE.lower()]
    features = [x for x in items if x["issue_type"].lower() == FEATURE_ISSUE_TYPE.lower()]
    by_parent: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for feature in features:
        if feature["parent_key"]:
            by_parent[feature["parent_key"]].append(feature)

    for epic in epics:
        epic["features"] = sorted(
            by_parent.get(epic["key"], []),
            key=lambda x: (x["start"] or date.max, x["key"]),
        )

        if epic["percent"] is None and CALCULATE_PERCENT_IF_MISSING and epic["features"]:
            done = sum(1 for x in epic["features"] if x["status_category"] == "done")
            epic["percent"] = round(100 * done / len(epic["features"]))

    return sorted(epics, key=lambda x: (x["start"] or date.max, x["project"], x["key"]))


# =============================================================================
# HTML GENERATION
# =============================================================================

def choose_year(epics: list[dict[str, Any]]) -> int:
    if ROADMAP_YEAR:
        return ROADMAP_YEAR
    years = []
    for e in epics:
        for d in (e["start"], e["target"]):
            if d:
                years.append(d.year)
    return min(years) if years else date.today().year


def pct_position(d: date | None, year: int, default: float) -> float:
    if not d:
        return default
    start = date(year, 1, 1)
    end = date(year, 12, 31)
    if d <= start:
        return 0
    if d >= end:
        return 100
    return ((d - start).days / max(1, (end - start).days)) * 100


def health_class(h: str) -> str:
    return {"green": "ok", "yellow": "warn", "red": "bad"}.get(h, "unknown")


def esc(x: Any) -> str:
    return html.escape(str(x or ""))


def render_epic(epic: dict[str, Any], year: int) -> str:
    left = pct_position(epic["start"], year, 2)
    right = pct_position(epic["target"], year, min(left + 20, 98))
    width = max(5, right - left)
    pct = int(round(epic["percent"] or 0))
    hc = health_class(epic["health"])
    feature_rows = "".join(render_feature(f) for f in epic["features"]) or \
        "<div class='no-features'>No linked Features returned by the configured JQL/parent mapping.</div>"

    data = {
        "key": epic["key"],
        "summary": epic["summary"],
        "project": epic["project"],
        "owner": epic["assignee"],
        "health": epic["health"],
        "status": epic["status"],
        "percent": pct,
        "start": epic["start"].isoformat() if epic["start"] else "",
        "target": epic["target"].isoformat() if epic["target"] else "",
        "updated": epic["updated"],
        "value": epic["value_statement"],
        "notes": epic["additional_notes"],
    }
    encoded = html.escape(json.dumps(data, ensure_ascii=False), quote=True)

    return f"""
    <section class="epic-row" data-project="{esc(epic['project'])}" data-health="{esc(epic['health'])}">
      <div class="epic-info">
        <div class="headline">
          <button class="expand" onclick="toggleDetails(this)" aria-label="Expand">+</button>
          <div>
            <div class="summary">{esc(epic['summary'])}</div>
            <div class="key">{esc(epic['key'])} · {esc(epic['project'])}</div>
          </div>
        </div>
        <div class="signals">
          <span class="health {hc}">HEALTH: {esc(epic['health']).upper()}</span>
          <span class="pct">{pct}%</span>
        </div>
        <div class="mini-progress"><span style="width:{pct}%"></span></div>
        <div class="meta">{esc(epic['assignee']) or 'Unassigned'} · {len(epic['features'])} feature(s)</div>
      </div>

      <div class="track">
        <div class="month-grid"></div>
        <div class="bar {hc}" style="left:{left:.2f}%;width:{width:.2f}%"
             data-epic="{encoded}" onclick="openDrawer(this)">
          <span>{esc(epic['key'])} · {esc(epic['summary'])}</span>
          <b>{pct}%</b>
        </div>
      </div>

      <div class="expanded">
        <div class="rich-grid">
          <div class="rich-card">
            <h4>Value Statement</h4>
            <div class="rich-body">{epic['value_statement']}</div>
          </div>
          <div class="rich-card">
            <h4>Additional Notes / Talking Points</h4>
            <div class="rich-body">{epic['additional_notes']}</div>
          </div>
        </div>
        <div class="features">
          <h4>Features</h4>
          {feature_rows}
        </div>
      </div>
    </section>
    """


def render_feature(f: dict[str, Any]) -> str:
    pct = int(round(f["percent"] or 0))
    hc = health_class(f["health"])
    return f"""
      <div class="feature">
        <div class="feature-main">
          <span class="feature-key">{esc(f['key'])}</span>
          <span class="feature-summary">{esc(f['summary'])}</span>
        </div>
        <div class="feature-signals">
          <span class="dot {hc}" title="Health: {esc(f['health'])}"></span>
          <span>{pct}%</span>
          <span class="feature-project">{esc(f['project'])}</span>
        </div>
        <div class="feature-progress"><span style="width:{pct}%"></span></div>
      </div>
    """


def build_html(epics: list[dict[str, Any]]) -> str:
    year = choose_year(epics)
    projects = sorted({e["project"] for e in epics if e["project"]})
    health_counts = {h: sum(1 for e in epics if e["health"] == h) for h in ("green", "yellow", "red")}
    avg_pct = round(sum((e["percent"] or 0) for e in epics) / len(epics)) if epics else 0

    rows = "\n".join(render_epic(e, year) for e in epics)
    project_options = "".join(f"<option>{esc(p)}</option>" for p in projects)

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Jira Product Roadmap</title>
<style>
:root{{--ink:#172b4d;--muted:#626f86;--line:#dcdfe4;--bg:#f6f7f9;--nav:#0b1f3a;--blue:#0c66e4;
--green:#1f845a;--yellow:#b65c02;--red:#c9372c;--label:310px}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font:14px Inter,Segoe UI,Arial,sans-serif}}
header{{background:var(--nav);color:white;padding:21px 28px;position:sticky;top:0;z-index:20}}
.top{{display:flex;justify-content:space-between;align-items:flex-end;gap:18px;flex-wrap:wrap}}h1{{margin:0;font-size:26px}}
.subtitle{{color:#b6c2cf;margin-top:4px;font-size:12px}}.filters{{display:flex;gap:7px;flex-wrap:wrap}}
input,select{{border:1px solid #8590a2;border-radius:6px;padding:8px 10px;background:white;color:var(--ink)}}
main{{max-width:1700px;margin:auto;padding:20px 28px 45px}}
.kpis{{display:grid;grid-template-columns:repeat(5,1fr);gap:11px;margin-bottom:15px}}
.kpi{{background:white;border:1px solid var(--line);border-radius:10px;padding:14px 16px}}
.kpi label{{font-size:10px;color:var(--muted);font-weight:800;text-transform:uppercase}}.kpi strong{{display:block;font-size:25px;margin-top:4px}}
.timeline-head{{display:grid;grid-template-columns:var(--label) 1fr;background:#f1f2f4;border:1px solid var(--line);border-radius:10px 10px 0 0;overflow:hidden}}
.label-head{{padding:12px 15px;font-size:11px;font-weight:800}}
.quarters{{display:grid;grid-template-columns:repeat(4,1fr)}}.quarters div{{padding:10px;text-align:center;font-weight:800;border-left:1px solid var(--line)}}
.epic-row{{display:grid;grid-template-columns:var(--label) 1fr;background:white;border-left:1px solid var(--line);border-right:1px solid var(--line);border-bottom:1px solid var(--line);position:relative}}
.epic-info{{padding:12px 14px;border-right:1px solid var(--line);min-height:105px}}.headline{{display:flex;gap:8px;align-items:flex-start}}
.expand{{width:23px;height:23px;border:0;border-radius:5px;background:#f1f2f4;color:#44546f;font-weight:900;cursor:pointer}}
.summary{{font-weight:800;font-size:13px}}.key,.meta{{font-size:10px;color:var(--muted);margin-top:3px}}
.signals{{display:flex;align-items:center;gap:8px;margin-top:8px}}.health{{font-size:9px;font-weight:850;border-radius:20px;padding:3px 7px}}
.health.ok{{background:#dcfff1;color:var(--green)}}.health.warn{{background:#fff3d6;color:var(--yellow)}}.health.bad{{background:#ffeceb;color:var(--red)}}.health.unknown{{background:#f1f2f4;color:#626f86}}
.pct{{font-size:10px;font-weight:850}}.mini-progress,.feature-progress{{height:5px;background:#ebecf0;border-radius:10px;overflow:hidden;margin-top:7px}}
.mini-progress span,.feature-progress span{{display:block;height:100%;background:var(--blue)}}
.track{{position:relative;min-height:105px;overflow:hidden}}.month-grid{{position:absolute;inset:0;background:repeating-linear-gradient(to right,transparent 0,transparent calc(8.333% - 1px),#eef0f2 calc(8.333% - 1px),#eef0f2 8.333%)}}
.bar{{position:absolute;top:36px;height:34px;border-radius:7px;color:white;padding:8px 9px;font-size:10px;font-weight:750;display:flex;justify-content:space-between;gap:8px;cursor:pointer;overflow:hidden;white-space:nowrap}}
.bar.ok{{background:#1f845a}}.bar.warn{{background:#b65c02}}.bar.bad{{background:#c9372c}}.bar.unknown{{background:#626f86}}.bar span{{overflow:hidden;text-overflow:ellipsis}}
.expanded{{display:none;grid-column:1/3;padding:14px 16px 18px 45px;background:#fafbfc;border-top:1px solid #eef0f2}}.epic-row.open .expanded{{display:block}}
.rich-grid{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}.rich-card{{background:white;border:1px solid var(--line);border-radius:8px;padding:12px 14px}}
.rich-card h4,.features h4{{font-size:10px;text-transform:uppercase;color:var(--muted);letter-spacing:.05em;margin:0 0 7px}}
.rich-body p{{font-size:11px;line-height:1.5;margin:0 0 7px}}.rich-body ul{{margin:4px 0;padding-left:18px}}.rich-body li{{font-size:11px;line-height:1.5;margin:3px 0}}.empty{{font-size:11px;color:#8993a4;font-style:italic}}
.features{{margin-top:13px}}.feature{{display:grid;grid-template-columns:1fr auto;gap:5px 15px;background:white;border:1px solid var(--line);border-radius:7px;padding:9px 11px;margin-top:6px}}
.feature-main{{display:flex;gap:8px}}.feature-key{{color:var(--blue);font-size:10px;font-weight:800}}.feature-summary{{font-size:11px;font-weight:700}}
.feature-signals{{display:flex;gap:8px;align-items:center;font-size:10px;color:var(--muted)}}.dot{{width:9px;height:9px;border-radius:50%}}.dot.ok{{background:var(--green)}}.dot.warn{{background:#e2b203}}.dot.bad{{background:var(--red)}}.dot.unknown{{background:#8993a4}}.feature-progress{{grid-column:1/3}}
.drawer{{position:fixed;right:-470px;top:0;width:min(455px,96vw);height:100vh;background:white;z-index:50;box-shadow:-12px 0 35px rgba(9,30,66,.2);transition:right .2s;padding:24px;overflow:auto}}.drawer.open{{right:0}}
.close{{float:right;border:0;background:#f1f2f4;border-radius:6px;padding:7px 10px;cursor:pointer}}.drawer h2{{margin:5px 0}}.drawer .dkey{{color:var(--blue);font-weight:800;font-size:11px;margin-top:28px}}
.dgrid{{display:grid;grid-template-columns:105px 1fr;gap:9px;font-size:11px;margin-top:18px}}.dlab{{color:var(--muted)}}.drawer h3{{font-size:11px;text-transform:uppercase;color:var(--muted);margin-top:22px}}
.drawer .rich-body{{font-size:12px}}.footer{{font-size:10px;color:var(--muted);margin-top:10px}}
@media(max-width:950px){{.kpis{{grid-template-columns:repeat(2,1fr)}}.timeline-wrap{{overflow:auto}}.timeline-head,.epic-row{{min-width:1100px}}.rich-grid{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<header><div class="top"><div><h1>Executive Product Roadmap</h1><div class="subtitle">Jira Server / Data Center · {year}</div></div>
<div class="filters"><input id="search" placeholder="Search Summary / key…" oninput="filterRows()">
<select id="projectFilter" onchange="filterRows()"><option value="">All Teams / Projects</option>{project_options}</select>
<select id="healthFilter" onchange="filterRows()"><option value="">All Health</option><option value="green">Green</option><option value="yellow">Yellow</option><option value="red">Red</option><option value="unknown">Unknown</option></select></div></div></header>
<main>
<div class="kpis">
<div class="kpi"><label>Epics</label><strong>{len(epics)}</strong></div>
<div class="kpi"><label>Average completion</label><strong>{avg_pct}%</strong></div>
<div class="kpi"><label>Green health</label><strong>{health_counts['green']}</strong></div>
<div class="kpi"><label>Yellow health</label><strong>{health_counts['yellow']}</strong></div>
<div class="kpi"><label>Red health</label><strong>{health_counts['red']}</strong></div>
</div>
<div class="timeline-wrap">
<div class="timeline-head"><div class="label-head">EPIC / TEAM</div><div class="quarters"><div>Q1</div><div>Q2</div><div>Q3</div><div>Q4</div></div></div>
{rows}
</div>
<div class="footer">Generated {datetime.now().strftime("%Y-%m-%d %H:%M")} from Jira. PAT is used only during generation and is not written into this HTML file.</div>
</main>

<aside class="drawer" id="drawer"><button class="close" onclick="drawer.classList.remove('open')">✕</button>
<div class="dkey" id="dkey"></div><h2 id="dsummary"></h2><span class="health" id="dhealth"></span>
<div class="dgrid">
<div class="dlab">Project / Team</div><div id="dproject"></div>
<div class="dlab">Owner</div><div id="downer"></div>
<div class="dlab">Status</div><div id="dstatus"></div>
<div class="dlab">Progress</div><div id="dpct"></div>
<div class="dlab">Start</div><div id="dstart"></div>
<div class="dlab">Target</div><div id="dtarget"></div>
<div class="dlab">Updated</div><div id="dupdated"></div>
</div>
<h3>Value Statement</h3><div class="rich-body" id="dvalue"></div>
<h3>Additional Notes / Talking Points</h3><div class="rich-body" id="dnotes"></div>
</aside>

<script>
const drawer=document.getElementById("drawer");
function toggleDetails(btn){{
 const row=btn.closest(".epic-row"); row.classList.toggle("open");
 btn.textContent=row.classList.contains("open")?"−":"+";
}}
function openDrawer(el){{
 const d=JSON.parse(el.dataset.epic);
 dkey.textContent=d.key; dsummary.textContent=d.summary; dproject.textContent=d.project;
 downer.textContent=d.owner||"Unassigned"; dstatus.textContent=d.status; dpct.textContent=d.percent+"%";
 dstart.textContent=d.start||"—"; dtarget.textContent=d.target||"—"; dupdated.textContent=d.updated||"—";
 dhealth.textContent="HEALTH: "+(d.health||"unknown").toUpperCase();
 dhealth.className="health "+(d.health==="green"?"ok":d.health==="yellow"?"warn":d.health==="red"?"bad":"unknown");
 dvalue.innerHTML=d.value; dnotes.innerHTML=d.notes; drawer.classList.add("open");
}}
function filterRows(){{
 const q=document.getElementById("search").value.toLowerCase();
 const p=document.getElementById("projectFilter").value;
 const h=document.getElementById("healthFilter").value;
 document.querySelectorAll(".epic-row").forEach(r=>{{
   const ok=(!q||r.textContent.toLowerCase().includes(q))&&(!p||r.dataset.project===p)&&(!h||r.dataset.health===h);
   r.style.display=ok?"grid":"none";
 }});
}}
</script>
</body></html>"""


def main() -> None:
    print(f"Jira: {JIRA_BASE_URL}")
    print(f"JQL : {JQL}")
    session = jira_session()

    try:
        ids = discover_field_ids(session)
        issues = search_issues(session, ids)
    except requests.exceptions.SSLError as e:
        sys.exit(f"\nSSL ERROR: {e}\nConfigure VERIFY_SSL with your corporate CA bundle; avoid disabling verification if possible.")
    except requests.exceptions.HTTPError as e:
        if e.response is not None and e.response.status_code in (401, 403):
            sys.exit("\nAUTH/PERMISSION ERROR: Check JIRA_PAT and the token user's Jira permissions.")
        raise

    items = [normalize_issue(i, ids) for i in issues]
    epics = build_hierarchy(items)

    print(f"Epics   : {len(epics)}")
    print(f"Features: {sum(len(e['features']) for e in epics)}")

    output = Path(OUTPUT_FILE).resolve()
    output.write_text(build_html(epics), encoding="utf-8")
    print(f"\nCreated: {output}")


if __name__ == "__main__":
    main()
