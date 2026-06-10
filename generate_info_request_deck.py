#!/usr/bin/env python3
"""
Travellers Autobarn — Information-Request Trends Deck Generator (v3)
====================================================================
Input is the aggregated array from Make:
    [ {"branch": "...", "issue": "Category :: specific question"}, ... ]
The `issue` field may be just "Category" or "Category :: detail".
Splits category (for trends) from detail (the actual question), so the deck
shows what customers really asked, not just the bucket.
"""

import sys
import json
import argparse
import os
import subprocess
import tempfile
from datetime import datetime
from collections import Counter, defaultdict
import anthropic

# ─── Branch colour config ───────────────────────────────────────────────────

BRANCH_COLORS = {
    "brisbane": "E97132", "bne": "E97132",
    "sydney": "2563EB", "syd": "2563EB",
    "melbourne": "0891B2", "mel": "0891B2",
    "perth": "16A34A", "per": "16A34A",
    "darwin": "CA8A04", "drw": "CA8A04",
    "cairns": "DC2626", "cns": "DC2626",
    "auckland": "6D28D9", "auk": "6D28D9",
    "christchurch": "DB2777", "chc": "DB2777",
    "los angeles": "0F766E", "lax": "0F766E",
    "las vegas": "9333EA", "las": "9333EA",
    "san francisco": "BE185D", "sfo": "BE185D",
}
FALLBACK_PALETTE = ["E97132", "2563EB", "0891B2", "16A34A", "CA8A04",
                    "DC2626", "6D28D9", "DB2777", "0F766E", "9333EA", "BE185D"]


def branch_color(branch, idx):
    key = str(branch).strip().lower()
    return BRANCH_COLORS.get(key, FALLBACK_PALETTE[idx % len(FALLBACK_PALETTE)])


import re

# Canonical branch names — unifies tags ("auckland"), codes ("AUK"), case.
BRANCH_CANON = {
    "auckland": "Auckland", "auk": "Auckland",
    "christchurch": "Christchurch", "chc": "Christchurch",
    "brisbane": "Brisbane", "bne": "Brisbane",
    "sydney": "Sydney", "syd": "Sydney",
    "melbourne": "Melbourne", "mel": "Melbourne",
    "perth": "Perth", "per": "Perth",
    "darwin": "Darwin", "dar": "Darwin", "drw": "Darwin",
    "cairns": "Cairns", "cns": "Cairns",
    "los angeles": "Los Angeles", "lax": "Los Angeles",
    "las vegas": "Las Vegas", "las": "Las Vegas",
    "san francisco": "San Francisco", "sfo": "San Francisco",
}


def canon_branch(value):
    key = str(value or "").strip().lower()
    if not key:
        return ""
    return BRANCH_CANON.get(key, str(value).strip().title())


def resolve_branch(branch_field, subject):
    """Use the Pickup Location field if it's filled; otherwise pull the
    branch code (e.g. '(AUK)') out of the ticket subject."""
    b = canon_branch(branch_field)
    if b and b.lower() != "unknown":
        return b
    codes = re.findall(r"\(([A-Za-z]{2,4})\)", str(subject or ""))
    for code in reversed(codes):  # branch code is usually last in the subject
        hit = BRANCH_CANON.get(code.strip().lower())
        if hit:
            return hit
    return "Unknown"


def split_issue(raw):
    """Turn 'Category :: specific question' into (category, detail).
    Tolerates plain 'Category' (no detail) and ':' as a fallback separator."""
    raw = str(raw or "").strip()
    if not raw:
        return ("Unknown", "")
    for sep in ("::", " - ", ":"):
        if sep in raw:
            cat, detail = raw.split(sep, 1)
            cat, detail = cat.strip(), detail.strip()
            if cat:
                return (cat, detail)
    return (raw, "")


# ─── Data loading ─────────────────────────────────────────────────────────────

def load_data(args):
    if args.data:
        with open(args.data) as f:
            return json.load(f)
    return json.loads(sys.stdin.read().strip())


# ─── Calculations ─────────────────────────────────────────────────────────────

def calculate_stats(rows):
    total = len(rows)
    by_issue = Counter()
    by_branch = Counter()
    by_branch_issue = defaultdict(Counter)
    details_by_issue = defaultdict(list)

    for r in rows:
        if isinstance(r, dict):
            branch = resolve_branch(r.get("branch"), r.get("subject"))
            raw_issue = r.get("issue")
        else:
            branch = resolve_branch(r[0] if len(r) > 0 else "", r[2] if len(r) > 2 else "")
            raw_issue = r[1] if len(r) > 1 else None
        category, detail = split_issue(raw_issue)
        by_issue[category] += 1
        by_branch[branch] += 1
        by_branch_issue[branch][category] += 1
        if detail:
            details_by_issue[category].append(detail)

    issue_ranking = by_issue.most_common()
    branch_ranking = by_branch.most_common()

    top_issue = issue_ranking[0] if issue_ranking else ("None", 0)
    busiest = branch_ranking[0] if branch_ranking else ("None", 0)

    branches = []
    for idx, (branch, count) in enumerate(branch_ranking):
        bi = by_branch_issue[branch].most_common()
        top = bi[0] if bi else ("None", 0)
        branches.append({
            "name": branch,
            "color": branch_color(branch, idx),
            "count": count,
            "pct": round(count / total * 100, 1) if total else 0,
            "top_issue": top[0],
            "top_issue_count": top[1],
        })

    # Specific questions per category (deduped, capped), ordered by category size
    details_by_category = []
    for name, count in issue_ranking:
        seen = []
        for d in details_by_issue.get(name, []):
            if d not in seen:
                seen.append(d)
        details_by_category.append({"category": name, "count": count, "details": seen})

    return {
        "total": total,
        "num_categories": len(issue_ranking),
        "top_issue_overall": {"name": top_issue[0], "count": top_issue[1]},
        "busiest_branch": {"name": busiest[0], "count": busiest[1]},
        "issue_ranking": [{"name": n, "count": c} for n, c in issue_ranking],
        "branches": branches,
        "details_by_category": details_by_category,
    }


# ─── Claude API ───────────────────────────────────────────────────────────────

def get_claude_narratives(stats, date_label, api_key=None):
    client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    summary = {
        "period": date_label,
        "total_requests": stats["total"],
        "most_asked_overall": stats["top_issue_overall"],
        "busiest_branch": stats["busiest_branch"],
        "request_types_ranked": stats["issue_ranking"],
        "specific_questions_by_category": [
            {"category": d["category"], "count": d["count"], "examples": d["details"][:6]}
            for d in stats["details_by_category"] if d["details"]
        ],
        "branches": [
            {"name": b["name"], "count": b["count"], "pct": b["pct"],
             "top_request": b["top_issue"]}
            for b in stats["branches"]
        ],
    }

    prompt = f"""You are writing concise insights for an internal monthly presentation for Travellers Autobarn, a campervan rental company. The deck shows what information customers ask for, so the operations manager can target staff training and reduce repeat questions.
All numbers below are pre-calculated — DO NOT recalculate or change any figures.
Base your TRAINING RECOMMENDATIONS on the SPECIFIC questions customers actually asked (specific_questions_by_category), not just the category labels. Each recommendation should name a concrete fix (e.g. "Add Auckland depot directions to the booking confirmation"), not a vague action.
Frame branch differences as training opportunities, never as blame on branch staff.
Return ONLY valid JSON with the exact keys listed. No markdown, no code fences.

STATS:
{json.dumps(summary, indent=2)}

Return JSON with these exact keys:
{{
  "cover_insight": "One sharp sentence (max 20 words) on the overall picture this period.",
  "by_category_insight": "One sentence (max 25 words) on what the category ranking tells us.",
  "by_branch_insight": "One sentence (max 25 words) on how request volume varies across branches.",
  "asked_insight": "One sentence (max 25 words) on the most common specific questions and what they reveal.",
  "training_priority": "One sentence (max 25 words) naming the single biggest, most concrete training/process fix for next month.",
  "training_recommendations": ["3 to 4 specific, concrete actions tied to the actual questions, each max 16 words"]
}}"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1100,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ─── PptxGenJS script builder ─────────────────────────────────────────────────

def build_pptx_script(stats, narratives, date_label, output_path):
    script = f"""
'use strict';
const pptxgen = require('pptxgenjs');

const DATE_LABEL   = {json.dumps(date_label)};
const TOTAL        = {stats["total"]};
const NUM_CATS     = {stats["num_categories"]};
const TOP_ISSUE    = {json.dumps(stats["top_issue_overall"])};
const BUSIEST      = {json.dumps(stats["busiest_branch"])};
const ISSUE_RANK   = {json.dumps(stats["issue_ranking"])};
const BRANCHES     = {json.dumps(stats["branches"])};
const DETAILS      = {json.dumps(stats["details_by_category"])};
const N            = {json.dumps(narratives)};

const NAVY='0F172A', WHITE='FFFFFF', ORANGE='F97316', LIGHT_GREY='F8FAFC',
      MID_GREY='64748B', DARK_GREY='1E293B';
const shadow = () => ({{ type:'outer', blur:8, offset:2, angle:135, color:'000000', opacity:0.12 }});
const clip = (s, n) => (s && s.length > n) ? s.substring(0, n-1) + '…' : (s || '');
const DSUF = DATE_LABEL ? ' — ' + DATE_LABEL : '';

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'Travellers Autobarn — Information Requests' + DSUF;

// ── SLIDE 1 — Cover ─────────────────────────────────────────────────────────
(function() {{
  const s = pres.addSlide();
  s.background = {{ color: NAVY }};
  s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:ORANGE}}, line:{{color:ORANGE}} }});
  s.addText('TRAVELLERS AUTOBARN', {{ x:0.5, y:0.2, w:7, h:0.4, fontSize:11, bold:true, color:'94A3B8', charSpacing:3 }});
  s.addText('INFORMATION REQUESTS', {{ x:0.5, y:0.55, w:9, h:0.9, fontSize:46, bold:true, color:WHITE, charSpacing:1 }});
  s.addText('What customers asked us for most', {{ x:0.5, y:1.4, w:9, h:0.45, fontSize:19, color:'94A3B8', italic:true }});
  if (DATE_LABEL) s.addText(DATE_LABEL, {{ x:0.5, y:1.8, w:9, h:0.35, fontSize:13, color:'64748B' }});
  const cards = [
    {{ label:'Total Requests', value: TOTAL.toString() }},
    {{ label:'Most Asked',     value: clip(TOP_ISSUE.name, 18) }},
    {{ label:'Busiest Branch', value: clip(BUSIEST.name, 16) }},
  ];
  cards.forEach((c, i) => {{
    const x = 0.5 + i * 3.2;
    s.addShape(pres.shapes.RECTANGLE, {{ x, y:2.4, w:2.9, h:1.05, fill:{{color:'1E293B'}}, line:{{color:'334155'}}, shadow:shadow() }});
    s.addText(c.value, {{ x:x+0.1, y:2.5, w:2.7, h:0.55, fontSize:22, bold:true, color:WHITE, align:'center', margin:0, valign:'middle' }});
    s.addText(c.label, {{ x, y:3.05, w:2.9, h:0.3, fontSize:10, color:'64748B', align:'center', margin:0 }});
  }});
  s.addShape(pres.shapes.RECTANGLE, {{ x:0.5, y:3.75, w:9.0, h:1.4, fill:{{color:'7C2D12'}}, line:{{color:ORANGE, pt:1.5}}, shadow:shadow() }});
  s.addText('KEY INSIGHT', {{ x:0.7, y:3.85, w:8.6, h:0.3, fontSize:9, bold:true, color:ORANGE, charSpacing:2 }});
  s.addText(N.cover_insight, {{ x:0.7, y:4.15, w:8.6, h:0.9, fontSize:15, color:WHITE, wrap:true }});
  s.addText('CONFIDENTIAL — INTERNAL USE ONLY', {{ x:0.5, y:5.35, w:9, h:0.2, fontSize:8, color:'334155', align:'center' }});
}})();

// ── SLIDE 2 — Requests by category ──────────────────────────────────────────
(function() {{
  const s = pres.addSlide();
  s.background = {{ color: LIGHT_GREY }};
  s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:NAVY}}, line:{{color:NAVY}} }});
  s.addText('MOST-ASKED INFORMATION REQUESTS', {{ x:0.4, y:0.15, w:9.2, h:0.4, fontSize:20, bold:true, color:NAVY }});
  s.addText('Overall ranking' + DSUF, {{ x:0.4, y:0.52, w:9, h:0.3, fontSize:11, color:MID_GREY, italic:true }});

  const top = ISSUE_RANK.slice(0, 8);
  const maxV = Math.max(...top.map(t => t.count), 1);
  const labelX = 0.4, labelW = 2.4, barX = 2.9, barMaxW = 3.7;
  const areaY = 1.05, rowH = (3.0) / Math.max(top.length, 1);
  top.forEach((t, i) => {{
    const y = areaY + i * rowH;
    s.addText(clip(t.name, 26), {{ x:labelX, y, w:labelW, h:rowH-0.06, fontSize:10, color:DARK_GREY, align:'right', valign:'middle', margin:0 }});
    const bw = Math.max((t.count / maxV) * barMaxW, 0.05);
    s.addShape(pres.shapes.RECTANGLE, {{ x:barX, y:y+0.04, w:barMaxW, h:rowH-0.18, fill:{{color:'E2E8F0'}}, line:{{color:'E2E8F0'}} }});
    s.addShape(pres.shapes.RECTANGLE, {{ x:barX, y:y+0.04, w:bw, h:rowH-0.18, fill:{{color:i===0?ORANGE:'94A3B8'}}, line:{{color:i===0?ORANGE:'94A3B8'}} }});
    s.addText(t.count.toString(), {{ x:barX+barMaxW+0.1, y, w:0.6, h:rowH-0.06, fontSize:11, bold:true, color:DARK_GREY, valign:'middle', margin:0 }});
  }});

  const callouts = [
    {{ label:'Total Requests', value: TOTAL.toString() }},
    {{ label:'Request Types',  value: NUM_CATS.toString() }},
    {{ label:'Top Request',    value: clip(TOP_ISSUE.name, 16) }},
  ];
  callouts.forEach((c, i) => {{
    const y = 1.05 + i * 1.0;
    s.addShape(pres.shapes.RECTANGLE, {{ x:7.0, y, w:2.6, h:0.85, fill:{{color:NAVY}}, line:{{color:'334155'}}, shadow:shadow() }});
    s.addText(c.value, {{ x:7.05, y:y+0.06, w:2.5, h:0.45, fontSize:16, bold:true, color:WHITE, align:'center', margin:0, valign:'middle' }});
    s.addText(c.label, {{ x:7.05, y:y+0.52, w:2.5, h:0.26, fontSize:9, color:'94A3B8', align:'center', margin:0 }});
  }});

  s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:4.5, w:9.2, h:0.72, fill:{{color:'FFF7ED'}}, line:{{color:ORANGE, pt:1}} }});
  s.addText(N.by_category_insight, {{ x:0.55, y:4.55, w:9.0, h:0.62, fontSize:11, color:DARK_GREY, italic:true, wrap:true }});
}})();

// ── SLIDE 3 — Requests by branch ────────────────────────────────────────────
(function() {{
  const s = pres.addSlide();
  s.background = {{ color: WHITE }};
  s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:NAVY}}, line:{{color:NAVY}} }});
  s.addText('WHERE THE QUESTIONS COME FROM', {{ x:0.4, y:0.15, w:9.2, h:0.4, fontSize:20, bold:true, color:NAVY }});
  s.addText('Information requests by branch' + DSUF, {{ x:0.4, y:0.52, w:9, h:0.3, fontSize:11, color:MID_GREY, italic:true }});

  const top = BRANCHES.slice(0, 10);
  const maxV = Math.max(...top.map(b => b.count), 1);
  const labelX = 0.4, labelW = 2.2, barX = 2.7, barMaxW = 3.3;
  const areaY = 1.05, rowH = (3.1) / Math.max(top.length, 1);
  top.forEach((b, i) => {{
    const y = areaY + i * rowH;
    s.addText(clip(b.name, 22), {{ x:labelX, y, w:labelW, h:rowH-0.04, fontSize:10, bold:true, color:DARK_GREY, align:'right', valign:'middle', margin:0 }});
    const bw = Math.max((b.count / maxV) * barMaxW, 0.05);
    s.addShape(pres.shapes.RECTANGLE, {{ x:barX, y:y+0.05, w:bw, h:rowH-0.2, fill:{{color:b.color}}, line:{{color:b.color}} }});
    s.addText(`${{b.count}}  (${{b.pct}}%)`, {{ x:barX+bw+0.1, y, w:1.4, h:rowH-0.04, fontSize:10, bold:true, color:DARK_GREY, valign:'middle', margin:0 }});
  }});

  s.addShape(pres.shapes.RECTANGLE, {{ x:7.7, y:1.05, w:2.0, h:0.95, fill:{{color:NAVY}}, line:{{color:'334155'}}, shadow:shadow() }});
  s.addText(clip(BUSIEST.name, 12), {{ x:7.72, y:1.12, w:1.96, h:0.5, fontSize:16, bold:true, color:WHITE, align:'center', margin:0, valign:'middle' }});
  s.addText('BUSIEST BRANCH', {{ x:7.72, y:1.64, w:1.96, h:0.3, fontSize:9, color:'94A3B8', align:'center', margin:0 }});

  s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:4.5, w:9.2, h:0.72, fill:{{color:'F0F9FF'}}, line:{{color:'BAE6FD', pt:1}} }});
  s.addText(N.by_branch_insight, {{ x:0.55, y:4.55, w:9.0, h:0.62, fontSize:11, color:DARK_GREY, italic:true, wrap:true }});
}})();

// ── SLIDE 4 — What customers actually asked ─────────────────────────────────
(function() {{
  const s = pres.addSlide();
  s.background = {{ color: WHITE }};
  s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:NAVY}}, line:{{color:NAVY}} }});
  s.addText('WHAT CUSTOMERS ACTUALLY ASKED', {{ x:0.4, y:0.15, w:9.2, h:0.4, fontSize:20, bold:true, color:NAVY }});
  s.addText('The specific questions behind each category' + DSUF, {{ x:0.4, y:0.52, w:9.2, h:0.3, fontSize:11, color:MID_GREY, italic:true }});

  const blocks = DETAILS.filter(d => d.details && d.details.length > 0).slice(0, 4);

  if (blocks.length === 0) {{
    s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:1.3, w:9.2, h:1.0, fill:{{color:LIGHT_GREY}}, line:{{color:'E2E8F0', pt:1}} }});
    s.addText('No specific question detail captured yet — once tickets carry a one-line summary, the actual questions will appear here.', {{
      x:0.7, y:1.5, w:8.6, h:0.6, fontSize:13, color:MID_GREY, italic:true, wrap:true, valign:'middle' }});
    return;
  }}

  // two columns of category blocks
  const colX = [0.4, 5.05];
  const colW = 4.55;
  const perCol = Math.ceil(blocks.length / 2);
  blocks.forEach((b, i) => {{
    const col = i < perCol ? 0 : 1;
    const idxInCol = i < perCol ? i : i - perCol;
    const x = colX[col];
    const blockH = 1.85;
    const y = 1.0 + idxInCol * (blockH + 0.15);

    s.addShape(pres.shapes.RECTANGLE, {{ x, y, w:colW, h:blockH, fill:{{color:LIGHT_GREY}}, line:{{color:'E2E8F0', pt:1}}, shadow:shadow() }});
    // count chip
    s.addShape(pres.shapes.RECTANGLE, {{ x:x+0.15, y:y+0.15, w:0.55, h:0.34, fill:{{color:ORANGE}}, line:{{color:ORANGE}} }});
    s.addText(b.count.toString(), {{ x:x+0.15, y:y+0.15, w:0.55, h:0.34, fontSize:14, bold:true, color:WHITE, align:'center', valign:'middle', margin:0 }});
    s.addText(clip(b.category, 34), {{ x:x+0.8, y:y+0.15, w:colW-0.95, h:0.34, fontSize:13, bold:true, color:NAVY, valign:'middle', margin:0 }});

    const qs = b.details.slice(0, 4);
    qs.forEach((q, qi) => {{
      const qy = y + 0.58 + qi * 0.3;
      s.addText([
        {{ text:'•  ', options:{{ color:ORANGE, bold:true }} }},
        {{ text: clip(q, 60), options:{{ color:DARK_GREY }} }},
      ], {{ x:x+0.2, y:qy, w:colW-0.35, h:0.28, fontSize:9.5, valign:'middle', margin:0 }});
    }});
  }});

  s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:5.0, w:9.2, h:0.5, fill:{{color:'FFF7ED'}}, line:{{color:ORANGE, pt:1}} }});
  s.addText(N.asked_insight || '', {{ x:0.55, y:5.03, w:9.0, h:0.44, fontSize:10.5, color:DARK_GREY, italic:true, wrap:true, valign:'middle' }});
}})();

// ── SLIDE 5 — Training focus ────────────────────────────────────────────────
(function() {{
  const s = pres.addSlide();
  s.background = {{ color: LIGHT_GREY }};
  s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:NAVY}}, line:{{color:NAVY}} }});
  s.addText('TRAINING FOCUS FOR NEXT MONTH', {{ x:0.4, y:0.15, w:9.2, h:0.4, fontSize:20, bold:true, color:NAVY }});
  s.addText('Turning customer questions into staff training' + DSUF, {{ x:0.4, y:0.52, w:9.2, h:0.3, fontSize:11, color:MID_GREY, italic:true }});

  s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:0.95, w:9.2, h:0.85, fill:{{color:'7C2D12'}}, line:{{color:ORANGE, pt:1.5}}, shadow:shadow() }});
  s.addText('PRIORITY', {{ x:0.55, y:1.0, w:2, h:0.3, fontSize:9, bold:true, color:ORANGE, charSpacing:2 }});
  s.addText(N.training_priority, {{ x:0.55, y:1.26, w:9.0, h:0.5, fontSize:13, color:WHITE, wrap:true, margin:0 }});

  s.addText('RECOMMENDED ACTIONS', {{ x:0.4, y:2.0, w:4.6, h:0.3, fontSize:12, bold:true, color:NAVY, charSpacing:1 }});
  const recs = (N.training_recommendations || []).slice(0, 4);
  recs.forEach((r, i) => {{
    const y = 2.4 + i * 0.62;
    s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y, w:4.6, h:0.54, fill:{{color:WHITE}}, line:{{color:'E2E8F0', pt:1}}, shadow:shadow() }});
    s.addShape(pres.shapes.OVAL, {{ x:0.52, y:y+0.13, w:0.28, h:0.28, fill:{{color:ORANGE}}, line:{{color:ORANGE}} }});
    s.addText((i+1).toString(), {{ x:0.52, y:y+0.13, w:0.28, h:0.28, fontSize:12, bold:true, color:WHITE, align:'center', valign:'middle', margin:0 }});
    s.addText(r, {{ x:0.95, y, w:3.95, h:0.54, fontSize:10.5, color:DARK_GREY, valign:'middle', wrap:true, margin:0 }});
  }});

  s.addText('BY-BRANCH FOCUS', {{ x:5.2, y:2.0, w:4.4, h:0.3, fontSize:12, bold:true, color:NAVY, charSpacing:1 }});
  const bf = BRANCHES.slice(0, 6);
  bf.forEach((b, i) => {{
    const y = 2.4 + i * 0.44;
    s.addShape(pres.shapes.OVAL, {{ x:5.2, y:y+0.06, w:0.14, h:0.14, fill:{{color:b.color}}, line:{{color:b.color}} }});
    s.addText([
      {{ text: clip(b.name, 14) + ':  ', options:{{ bold:true, color:DARK_GREY }} }},
      {{ text: 'train on ' + clip(b.top_issue, 26), options:{{ color:MID_GREY }} }},
    ], {{ x:5.42, y, w:4.2, h:0.3, fontSize:10, valign:'middle', margin:0 }});
  }});
}})();

const outputPath = {json.dumps(output_path)};
pres.writeFile({{ fileName: outputPath }})
  .then(() => {{ console.log('SUCCESS:' + outputPath); }})
  .catch(err => {{ console.error('ERROR:', err.message); process.exit(1); }});
"""
    return script


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate Information-Request Trends PPTX")
    parser.add_argument("--data",       help="Path to JSON data file (else reads stdin)")
    parser.add_argument("--output",     default="info_request_report.pptx")
    parser.add_argument("--date-label", help="Human-readable period e.g. 'May 2026'")
    parser.add_argument("--api-key",    help="Anthropic API key")
    parser.add_argument("--skip-claude", action="store_true")
    args = parser.parse_args()

    print("Loading data...", file=sys.stderr)
    rows = load_data(args)
    print(f"  Loaded {len(rows)} rows", file=sys.stderr)

    print("Calculating statistics...", file=sys.stderr)
    stats = calculate_stats(rows)
    print(f"  {stats['total']} requests across {len(stats['branches'])} branches", file=sys.stderr)

    date_label = args.date_label if args.date_label is not None else datetime.now().strftime("%B %Y")

    if args.skip_claude:
        print("Skipping Claude API...", file=sys.stderr)
        narratives = {
            "cover_insight": f"{stats['total']} information requests this period, led by {stats['top_issue_overall']['name']}.",
            "by_category_insight": "Category ranking highlights the questions customers ask most often.",
            "by_branch_insight": "Request volume varies across branches, pointing to where support is most needed.",
            "asked_insight": "The specific questions reveal where wording and instructions could be clearer.",
            "training_priority": f"Focus next month's training on {stats['top_issue_overall']['name']}.",
            "training_recommendations": [
                "Add a quick-reference card for the top request type.",
                "Brief front-desk staff on the most common questions.",
                "Review knowledge-base coverage for the busiest branch.",
            ],
        }
    else:
        print("Calling Claude API for narrative text...", file=sys.stderr)
        narratives = get_claude_narratives(stats, date_label, api_key=args.api_key)
        print("  Narratives received.", file=sys.stderr)

    print("Building PptxGenJS script...", file=sys.stderr)
    script = build_pptx_script(stats, narratives, date_label, args.output)

    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False) as f:
        f.write(script)
        script_path = f.name

    print(f"Generating PPTX at {args.output}...", file=sys.stderr)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    result = subprocess.run(["node", script_path], capture_output=True, text=True, cwd=script_dir)
    os.unlink(script_path)

    if result.returncode != 0:
        print("Node.js error:", result.stderr, file=sys.stderr)
        sys.exit(1)

    print(f"\nDone! Output: {args.output}", file=sys.stderr)
    print(args.output)


if __name__ == "__main__":
    main()
