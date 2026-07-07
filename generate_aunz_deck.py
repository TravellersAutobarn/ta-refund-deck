#!/usr/bin/env python3
"""
Travellers Autobarn — AU/NZ Refund Awareness Deck Generator
============================================================
Receives raw Google Sheets data (as JSON via stdin or file path arg),
calculates all stats in Python, calls Claude API for narrative text only,
then shells out to Node/PptxGenJS to build the final .pptx file.

v2 changes:
  - Branches sorted worst offender (highest refund total) to least
  - Speaker notes (talking points) added to every slide via Claude
  - Month-over-month comparison: expects rows for BOTH the report month
    and the prior month; splits them using the Month column (K).
    If prior-month rows are absent, comparison is skipped gracefully.

Usage (from Make):
    echo '<json_data>' | python generate_aunz_deck.py
    python generate_aunz_deck.py --data data.json --output /path/to/output.pptx
"""

import sys
import json
import argparse
import os
import subprocess
import tempfile
from datetime import datetime, date
from collections import defaultdict
import anthropic

# ─── Branch config ───────────────────────────────────────────────────────────

AU_BRANCHES = {
    "BNE": {"name": "Brisbane",  "color": "7C3AED", "region": "AU"},  # purple
    "CNS": {"name": "Cairns",    "color": "065F46", "region": "AU"},  # dark green
    "MEL": {"name": "Melbourne", "color": "D97706", "region": "AU"},  # amber/gold
    "SYD": {"name": "Sydney",    "color": "16A34A", "region": "AU"},  # green
    "PER": {"name": "Perth",     "color": "2563EB", "region": "AU"},  # blue
    "DAR": {"name": "Darwin",    "color": "DC2626", "region": "AU"},  # red
}

NZ_BRANCHES = {
    "AUK": {"name": "Auckland",      "color": "6D28D9", "region": "NZ"},  # indigo/purple
    "CHC": {"name": "Christchurch",  "color": "DB2777", "region": "NZ"},  # magenta/pink
}

ALL_BRANCHES = {**AU_BRANCHES, **NZ_BRANCHES}

REFUND_CATEGORIES = ["DOR - GW", "Camp Gear - Prep", "Mechanical", "Kitchen - Internal"]

# Sheet column mapping (0-indexed, matches the column order described)
COL = {
    "no":               0,
    "res_num":          1,
    "category":         2,
    "rego":             3,
    "fleet":            4,
    "pickup":           5,
    "last_name":        6,
    "date_entered":     7,
    "entered_by":       8,
    "follow_up":        9,
    "month":            10,
    "pickup_branch":    11,
    "refund_category":  12,
    "amount":           13,
    "details":          15,
    "notes":            16,
}


# ─── Data loading ────────────────────────────────────────────────────────────

def load_data(args):
    """Load raw rows from JSON file, stdin, or inline arg."""
    if args.data:
        with open(args.data) as f:
            rows = json.load(f)
    else:
        raw = sys.stdin.read().strip()
        rows = json.loads(raw)
    return rows


def normalise_rows(rows):
    """Convert dict rows (from Make) into plain indexed lists, once."""
    if rows and isinstance(rows[0], dict):
        rows = [list(r.values()) for r in rows]
    return rows


def parse_amount(val):
    """Safely parse dollar amount to float."""
    if val is None or val == "":
        return 0.0
    s = str(val).replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date(val):
    """Parse date string to date object (handles multiple formats)."""
    if not val:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%d/%b/%Y"):
        try:
            return datetime.strptime(str(val).strip(), fmt).date()
        except ValueError:
            continue
    return None


# ─── Month splitting (for month-over-month comparison) ──────────────────────

def split_rows_by_month(rows):
    """
    Group rows by the Month column (K). The most recent month present is
    the REPORT month; the calendar month immediately before it (if present)
    is the PRIOR month used for comparison.

    Returns: (report_rows, prior_rows, report_label, prior_label)
    prior_rows may be an empty list and prior_label None if no prior data.
    """
    grouped = defaultdict(list)
    for row in rows:
        d = parse_date(row[COL["month"]])
        if d is None:
            continue
        grouped[(d.year, d.month)].append(row)

    if not grouped:
        return rows, [], None, None

    report_key = max(grouped.keys())
    ry, rm = report_key
    # Calendar month immediately before the report month
    if rm == 1:
        prior_key = (ry - 1, 12)
    else:
        prior_key = (ry, rm - 1)

    report_rows = grouped[report_key]
    prior_rows  = grouped.get(prior_key, [])

    report_label = date(report_key[0], report_key[1], 1).strftime("%B %Y")
    prior_label  = date(prior_key[0], prior_key[1], 1).strftime("%B %Y") if prior_rows else None

    return report_rows, prior_rows, report_label, prior_label


# ─── Calculations ────────────────────────────────────────────────────────────

def classify_region(branch_code):
    """Return 'AU', 'NZ', or None."""
    bc = str(branch_code).strip().upper()
    if bc in AU_BRANCHES:
        return "AU"
    if bc in NZ_BRANCHES:
        return "NZ"
    return None


def calculate_stats(rows):
    """
    Master stats calculation. Returns a dict with all numbers needed
    for every slide. Python does ALL the maths here — Claude never sees raw data.
    Branch lists are sorted WORST OFFENDER FIRST (highest total to lowest).
    """
    # Separate AU and NZ
    au_rows = [r for r in rows if classify_region(r[COL["pickup_branch"]]) == "AU"]
    nz_rows = [r for r in rows if classify_region(r[COL["pickup_branch"]]) == "NZ"]

    def branch_stats(branch_rows, branch_map):
        """Compute per-branch breakdown."""
        totals   = defaultdict(float)
        counts   = defaultdict(int)
        by_cat   = defaultdict(lambda: defaultdict(float))
        claims   = defaultdict(list)

        for row in branch_rows:
            bc  = str(row[COL["pickup_branch"]]).strip().upper()
            amt = parse_amount(row[COL["amount"]])
            cat = str(row[COL["refund_category"]]).strip()
            totals[bc]  += amt
            counts[bc]  += 1
            by_cat[bc][cat] += amt
            claims[bc].append({
                "res_num":  row[COL["res_num"]],
                "last_name": row[COL["last_name"]],
                "rego":     row[COL["rego"]],
                "fleet":    row[COL["fleet"]],
                "amount":   amt,
                "category": cat,
                "details":  row[COL["details"]],
                "notes":    row[COL["notes"]],
                "date":     row[COL["date_entered"]],
            })

        grand_total = sum(totals.values())
        grand_count = sum(counts.values())

        branches = []
        for bc, cfg in branch_map.items():
            branch_total = totals.get(bc, 0.0)
            branch_count = counts.get(bc, 0)
            pct = (branch_total / grand_total * 100) if grand_total else 0
            cat_breakdown = {c: by_cat[bc].get(c, 0.0) for c in REFUND_CATEGORIES}
            # Largest single claim
            branch_claims = sorted(claims.get(bc, []), key=lambda x: x["amount"], reverse=True)
            biggest_claim = branch_claims[0] if branch_claims else None
            branches.append({
                "code":          bc,
                "name":          cfg["name"],
                "color":         cfg["color"],
                "total":         round(branch_total, 2),
                "count":         branch_count,
                "pct_of_total":  round(pct, 1),
                "by_category":   {k: round(v, 2) for k, v in cat_breakdown.items()},
                "claims":        branch_claims,
                "biggest_claim": biggest_claim,
            })

        # ── SORT: worst offender (highest total) first ──
        branches.sort(key=lambda b: b["total"], reverse=True)

        # Category totals across all branches
        category_totals = defaultdict(float)
        for row in branch_rows:
            cat = str(row[COL["refund_category"]]).strip()
            category_totals[cat] += parse_amount(row[COL["amount"]])

        return {
            "grand_total":      round(grand_total, 2),
            "grand_count":      grand_count,
            "branches":         branches,
            "category_totals":  {k: round(v, 2) for k, v in category_totals.items()},
            "avg_claim":        round(grand_total / grand_count, 2) if grand_count else 0,
            "largest_claim":    max((parse_amount(r[COL["amount"]]) for r in branch_rows), default=0),
        }

    au_stats = branch_stats(au_rows, AU_BRANCHES)
    nz_stats = branch_stats(nz_rows, NZ_BRANCHES)

    # Equipment gap: biggest single "Camp Gear - Prep" claim per branch
    def equipment_gap(branch_list):
        gaps = []
        for b in branch_list:
            gear_claims = [c for c in b["claims"] if "GEAR" in c["category"].upper()]
            if gear_claims:
                top = max(gear_claims, key=lambda x: x["amount"])
                gaps.append({**top, "branch_code": b["code"], "branch_name": b["name"], "color": b["color"]})
        return gaps

    au_stats["equipment_gaps"] = equipment_gap(au_stats["branches"])
    nz_stats["equipment_gaps"] = equipment_gap(nz_stats["branches"])

    return {"au": au_stats, "nz": nz_stats}


def attach_prior_month(stats, prior_stats, prior_label):
    """
    Attach prior-month comparison numbers onto the report stats.
    If prior_stats is None, marks everything as having no comparison.
    """
    for region in ("au", "nz"):
        cur = stats[region]
        if prior_stats is None:
            cur["prev"] = None
            for b in cur["branches"]:
                b["prev_total"] = None
                b["prev_count"] = None
                b["prev_by_category"] = None
            continue

        prev = prior_stats[region]
        cur["prev"] = {
            "label":     prior_label,
            "total":     prev["grand_total"],
            "count":     prev["grand_count"],
            "avg_claim": prev["avg_claim"],
        }
        prev_map = {b["code"]: b for b in prev["branches"]}
        for b in cur["branches"]:
            pb = prev_map.get(b["code"])
            b["prev_total"] = pb["total"] if pb else None
            b["prev_count"] = pb["count"] if pb else None
            b["prev_by_category"] = pb["by_category"] if pb else None
    return stats


# ─── Claude API — narrative + talking points ────────────────────────────────

def get_claude_narratives(stats, date_label, prior_label, api_key=None):
    """
    Call Claude API with PRE-CALCULATED numbers only.
    Claude writes short insight sentences AND presenter talking points —
    never calculates anything.
    """
    client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    au = stats["au"]
    nz = stats["nz"]

    def region_summary(r):
        s = {
            "total": r["grand_total"],
            "count": r["grand_count"],
            "avg_claim": r["avg_claim"],
            "category_totals": r["category_totals"],
            "branches": [
                {
                    "code": b["code"],
                    "name": b["name"],
                    "total": b["total"],
                    "count": b["count"],
                    "pct": b["pct_of_total"],
                    "by_category": b["by_category"],
                    "prev_month_total": b["prev_total"],
                    "prev_month_count": b["prev_count"],
                }
                for b in r["branches"]
            ],
        }
        if r.get("prev"):
            s["prev_month"] = r["prev"]
        return s

    stats_summary = {
        "period": date_label,
        "prior_period": prior_label,
        "au": region_summary(au),
        "nz": region_summary(nz),
    }

    comparison_rule = (
        "Prior-month numbers are included — weave the month-over-month change "
        "(up, down, or flat, and by roughly how much) into insights and talking points. "
        "Refunds going DOWN is good news; going UP needs attention."
        if prior_label else
        "No prior-month data is available — do not mention any month-over-month comparison."
    )

    prompt = f"""You are writing concise insights and presenter talking points for a refund awareness presentation for Travellers Autobarn (a campervan rental company). The goal of the deck is to make branches aware of refunds and drive refund costs DOWN.
All numbers below are pre-calculated — DO NOT recalculate or change any figures.
Branches are listed worst offender first (highest refund total to lowest).
{comparison_rule}
Talking points are spoken presenter notes: plain sentences, no markdown, no bullets, written as if said aloud to branch managers. Each branch talking point should end with one direct question or action to put to that branch.
Return ONLY valid JSON with the exact keys listed. No markdown, no code fences.

STATS:
{json.dumps(stats_summary, indent=2)}

Return JSON with these exact keys:
{{
  "au_cover_insight": "One sharp sentence (max 20 words) about the AU refund picture this period.",
  "au_branch_cards": {{
    "BNE": "One italic insight sentence for Brisbane card (max 15 words)",
    "CNS": "...",
    "MEL": "...",
    "SYD": "...",
    "PER": "...",
    "DAR": "..."
  }},
  "au_where_money_went": "One sentence (max 25 words) summarising the category breakdown across AU branches.",
  "au_equipment_gap": "One sentence (max 20 words) on the equipment/gear refund theme.",
  "nz_cover_insight": "One sharp sentence (max 20 words) about the NZ refund picture.",
  "nz_branch_cards": {{
    "AUK": "One italic insight sentence for Auckland card (max 15 words)",
    "CHC": "..."
  }},
  "au_branch_detail": {{
    "BNE": "Two sentences (max 30 words total) for the Brisbane detail slide narrative.",
    "CNS": "...",
    "MEL": "...",
    "SYD": "...",
    "PER": "...",
    "DAR": "..."
  }},
  "nz_branch_detail": {{
    "AUK": "Two sentences (max 30 words total) for Auckland detail slide narrative.",
    "CHC": "..."
  }},
  "au_cover_talking_points": "Three to four spoken sentences (max 80 words) introducing the AU refund picture, including the month-over-month story if prior data exists.",
  "au_branch_overview_talking_points": "Three to four spoken sentences (max 80 words) walking through the branch comparison, naming the worst branch and any branch that improved or worsened most versus the prior month.",
  "au_where_money_went_talking_points": "Three to four spoken sentences (max 80 words) telling the category story and what it means operationally.",
  "au_equipment_gap_talking_points": "Two to three spoken sentences (max 60 words) on the equipment/gear refunds and what branches should check.",
  "au_branch_talking_points": {{
    "BNE": "Two to three spoken sentences (max 60 words) for the Brisbane detail slide, including month-over-month direction if available, ending with one question or action for the branch.",
    "CNS": "...",
    "MEL": "...",
    "SYD": "...",
    "PER": "...",
    "DAR": "..."
  }}
}}"""

    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=2500,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = message.content[0].text.strip()
    # Strip any accidental markdown fences
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ─── PptxGenJS script builder ─────────────────────────────────────────────────

def build_pptx_script(stats, narratives, date_label, output_path):
    """
    Generate the Node.js / PptxGenJS script as a string.
    All data is embedded as a const — the script is self-contained.
    """
    au = stats["au"]
    nz = stats["nz"]

    au_branches_js = json.dumps(au["branches"])
    nz_branches_js = json.dumps(nz["branches"])
    narratives_js  = json.dumps(narratives)
    au_prev_js     = json.dumps(au.get("prev"))
    prior_label    = au["prev"]["label"] if au.get("prev") else None
    prior_short    = prior_label.split(" ")[0] if prior_label else ""

    script = f"""
'use strict';
const pptxgen = require('pptxgenjs');

// ─── Embedded data (all pre-calculated by Python) ──────────────────────────
const DATE_LABEL    = {json.dumps(date_label)};
const AU_TOTAL      = {au["grand_total"]};
const AU_COUNT      = {au["grand_count"]};
const AU_AVG        = {au["avg_claim"]};
const AU_CAT_TOTALS = {json.dumps(au["category_totals"])};
const AU_BRANCHES   = {au_branches_js};
const NZ_TOTAL      = {nz["grand_total"]};
const NZ_COUNT      = {nz["grand_count"]};
const NZ_AVG        = {nz["avg_claim"]};
const NZ_CAT_TOTALS = {json.dumps(nz["category_totals"])};
const NZ_BRANCHES   = {nz_branches_js};
const AU_EQUIP_GAPS = {json.dumps(au["equipment_gaps"])};
const NZ_EQUIP_GAPS = {json.dumps(nz["equipment_gaps"])};
const AU_PREV       = {au_prev_js};
const PRIOR_SHORT   = {json.dumps(prior_short)};
const N             = {narratives_js};

// ─── Helpers ──────────────────────────────────────────────────────────────
const fmtAUD = v => 'A$' + Number(v).toLocaleString('en-AU', {{minimumFractionDigits:0, maximumFractionDigits:0}});
const fmtNZD = v => 'NZ$' + Number(v).toLocaleString('en-NZ', {{minimumFractionDigits:0, maximumFractionDigits:0}});
const fmtPct = v => v.toFixed(1) + '%';

const NAVY        = '0F172A';
const WHITE       = 'FFFFFF';
const ORANGE      = 'F97316';
const LIGHT_GREY  = 'F8FAFC';
const MID_GREY    = '64748B';
const DARK_GREY   = '1E293B';
const GOOD_GREEN  = '16A34A';
const BAD_RED     = 'DC2626';

const makeShadow = () => ({{ type: 'outer', blur: 8, offset: 2, angle: 135, color: '000000', opacity: 0.12 }});

// Month-over-month delta line. Refunds DOWN = green (good), UP = red (bad).
// kind: 'money' or 'count'. Returns null when no prior data.
const deltaLine = (cur, prev, kind) => {{
  if (prev === null || prev === undefined || !PRIOR_SHORT) return null;
  const diff = cur - prev;
  if (Math.abs(diff) < 0.005) {{
    return {{ text: 'level with ' + PRIOR_SHORT, color: MID_GREY }};
  }}
  const up = diff > 0;
  const arrow = up ? '\\u25B2' : '\\u25BC';
  const color = up ? BAD_RED : GOOD_GREEN;
  const size  = kind === 'count'
    ? Math.abs(Math.round(diff)) + (Math.abs(Math.round(diff)) === 1 ? ' claim' : ' claims')
    : fmtAUD(Math.abs(diff));
  const pct = prev > 0 ? ' (' + (up ? '+' : '-') + Math.abs(diff / prev * 100).toFixed(0) + '%)' : '';
  return {{ text: arrow + ' ' + size + pct + ' vs ' + PRIOR_SHORT, color }};
}};

// ─── Presentation setup ───────────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title  = `Travellers Autobarn — AU/NZ Refund Awareness ${{DATE_LABEL}}`;

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — AU Cover
// ════════════════════════════════════════════════════════════════════════════
(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: NAVY }};

  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: ORANGE }}, line:{{ color: ORANGE }} }});

  // Logo area / wordmark
  slide.addText('TRAVELLERS AUTOBARN', {{
    x:0.5, y:0.2, w:6, h:0.4,
    fontSize:11, bold:true, color:'94A3B8', charSpacing:3, align:'left'
  }});

  slide.addText('AUSTRALIA', {{
    x:0.5, y:0.55, w:9, h:0.9,
    fontSize:52, bold:true, color:WHITE, align:'left', charSpacing:2
  }});
  slide.addText('Refund Awareness Report', {{
    x:0.5, y:1.35, w:9, h:0.45,
    fontSize:20, color:'94A3B8', align:'left', italic:true
  }});
  slide.addText(DATE_LABEL, {{
    x:0.5, y:1.75, w:9, h:0.35,
    fontSize:13, color:'64748B', align:'left'
  }});

  // ── Big stat callouts (with month-over-month deltas when available) ──
  const stats = [
    {{ label:'Total Refunds', value: fmtAUD(AU_TOTAL),
       delta: AU_PREV ? deltaLine(AU_TOTAL, AU_PREV.total, 'money') : null }},
    {{ label:'Total Claims',  value: AU_COUNT.toString(),
       delta: AU_PREV ? deltaLine(AU_COUNT, AU_PREV.count, 'count') : null }},
    {{ label:'Average Claim', value: fmtAUD(AU_AVG),
       delta: AU_PREV ? deltaLine(AU_AVG, AU_PREV.avg_claim, 'money') : null }},
  ];
  stats.forEach((s, i) => {{
    const x = 0.5 + i * 3.2;
    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y:2.35, w:2.9, h:1.05,
      fill:{{ color:'1E293B' }}, line:{{ color:'334155' }},
      shadow: makeShadow()
    }});
    slide.addText(s.value, {{
      x, y:2.40, w:2.9, h:0.45,
      fontSize:26, bold:true, color:WHITE, align:'center', margin:0
    }});
    slide.addText(s.label, {{
      x, y:2.82, w:2.9, h:0.24,
      fontSize:9, color:'64748B', align:'center', margin:0
    }});
    if (s.delta) {{
      slide.addText(s.delta.text, {{
        x, y:3.06, w:2.9, h:0.28,
        fontSize:9, bold:true, color: s.delta.color, align:'center', margin:0
      }});
    }}
  }});

  // ── Branch dots list ──
  slide.addText('Branches covered:', {{
    x:0.5, y:3.62, w:5, h:0.28,
    fontSize:10, color:'64748B', align:'left'
  }});

  AU_BRANCHES.forEach((b, i) => {{
    const col = i < 3 ? 0 : 1;
    const row = i % 3;
    const x   = 0.5 + col * 3.0;
    const y   = 3.9 + row * 0.32;
    // Coloured dot
    slide.addShape(pres.shapes.OVAL, {{
      x: x, y: y + 0.04, w: 0.14, h: 0.14,
      fill: {{ color: b.color }}, line: {{ color: b.color }}
    }});
    slide.addText(`${{b.code}} — ${{b.name}}`, {{
      x: x + 0.2, y, w: 2.5, h: 0.28,
      fontSize:11, color:WHITE, align:'left', bold:false
    }});
  }});

  // ── Key insight box ──
  slide.addShape(pres.shapes.RECTANGLE, {{
    x:6.3, y:3.5, w:3.4, h:1.75,
    fill:{{ color:'7C2D12' }}, line:{{ color: ORANGE, pt:1.5 }},
    shadow: makeShadow()
  }});
  slide.addText('KEY INSIGHT', {{
    x:6.4, y:3.55, w:3.2, h:0.3,
    fontSize:9, bold:true, color: ORANGE, charSpacing:2
  }});
  slide.addText(N.au_cover_insight, {{
    x:6.4, y:3.85, w:3.2, h:1.3,
    fontSize:12, color:WHITE, align:'left', wrap:true
  }});

  // Bottom rule
  slide.addShape(pres.shapes.LINE, {{
    x:0.5, y:5.4, w:9, h:0, line:{{ color:'1E293B', pt:1 }}
  }});
  slide.addText('CONFIDENTIAL — INTERNAL USE ONLY', {{
    x:0.5, y:5.43, w:9, h:0.2,
    fontSize:8, color:'334155', align:'center'
  }});

  // Presenter talking points (hidden speaker notes)
  if (N.au_cover_talking_points) slide.addNotes(N.au_cover_talking_points);
}})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — Refunds by AU Branch (6 cards, worst offender first)
// ════════════════════════════════════════════════════════════════════════════
(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: LIGHT_GREY }};

  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: NAVY }}, line:{{ color: NAVY }} }});
  slide.addText('REFUNDS BY BRANCH', {{
    x:0.4, y:0.15, w:9, h:0.4,
    fontSize:20, bold:true, color: NAVY, align:'left'
  }});
  slide.addText(`Australia — ${{DATE_LABEL}} — worst to best`, {{
    x:0.4, y:0.52, w:9, h:0.3,
    fontSize:11, color: MID_GREY, italic:true
  }});

  // 6 branch cards in 3×2 grid
  const cols = 3;
  const cW = 3.0, cH = 2.0;
  const startX = 0.4, startY = 0.95, gapX = 0.15, gapY = 0.15;

  AU_BRANCHES.forEach((b, i) => {{
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x   = startX + col * (cW + gapX);
    const y   = startY + row * (cH + gapY);
    const ins = N.au_branch_cards[b.code] || '';
    const d   = deltaLine(b.total, b.prev_total, 'money');

    // Card bg
    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y, w:cW, h:cH,
      fill:{{ color: WHITE }},
      line:{{ color:'E2E8F0', pt:1 }},
      shadow: makeShadow()
    }});
    // Coloured left accent
    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y, w:0.07, h:cH,
      fill:{{ color: b.color }}, line:{{ color: b.color }}
    }});
    // Branch code header
    slide.addText(b.code, {{
      x: x+0.15, y: y+0.1, w:1.2, h:0.38,
      fontSize:22, bold:true, color: b.color, margin:0
    }});
    slide.addText(b.name, {{
      x: x+0.15, y: y+0.45, w:2.7, h:0.22,
      fontSize:10, color: MID_GREY, italic:true, margin:0
    }});
    // Amount
    slide.addText(fmtAUD(b.total), {{
      x: x+0.15, y: y+0.72, w:2.7, h:0.42,
      fontSize:22, bold:true, color: DARK_GREY, margin:0
    }});
    // Claim count
    slide.addText(`${{b.count}} claim${{b.count !== 1 ? 's' : ''}}`, {{
      x: x+0.15, y: y+1.1, w:1.5, h:0.24,
      fontSize:11, color: MID_GREY, margin:0
    }});
    // % of total badge
    slide.addShape(pres.shapes.RECTANGLE, {{
      x: x+1.75, y: y+1.08, w:1.1, h:0.28,
      fill:{{ color: b.color }}, line:{{ color: b.color }}
    }});
    slide.addText(fmtPct(b.pct_of_total), {{
      x: x+1.75, y: y+1.08, w:1.1, h:0.28,
      fontSize:10, bold:true, color: WHITE, align:'center', margin:0
    }});
    // Month-over-month delta line
    if (d) {{
      slide.addText(d.text, {{
        x: x+0.15, y: y+1.38, w:2.75, h:0.18,
        fontSize:8, bold:true, color: d.color, margin:0
      }});
    }}
    // Italic insight
    slide.addText(ins, {{
      x: x+0.12, y: y+1.56, w:2.75, h:0.4,
      fontSize:8.5, color: MID_GREY, italic:true, wrap:true, margin:0
    }});
  }});

  if (N.au_branch_overview_talking_points) slide.addNotes(N.au_branch_overview_talking_points);
}})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — Where the Money Went (stacked bar + 4 stat callouts)
// ════════════════════════════════════════════════════════════════════════════
(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: WHITE }};

  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: NAVY }}, line:{{ color: NAVY }} }});
  slide.addText('WHERE THE MONEY WENT', {{
    x:0.4, y:0.15, w:9, h:0.4,
    fontSize:20, bold:true, color: NAVY
  }});
  slide.addText(`Refund category breakdown by branch — ${{DATE_LABEL}}`, {{
    x:0.4, y:0.52, w:9, h:0.28,
    fontSize:11, color: MID_GREY, italic:true
  }});

  const catColors = ['E63946', 'F4A261', '2A9D8F', '457B9D'];
  const ALL_CATS = ['DOR - GW', 'Camp Gear - Prep', 'Mechanical', 'Kitchen - Internal'];

  const barAreaX = 0.4;
  const barAreaY = 0.9;
  const barAreaH = 2.7;
  const barW = 0.8;
  const barGap = 0.28;
  const maxVal = Math.max(...AU_BRANCHES.map(b =>
    ALL_CATS.reduce((sum, cat) => sum + (b.by_category[cat] || 0), 0)
  ), 1);

  AU_BRANCHES.forEach((b, bi) => {{
    const barX = barAreaX + bi * (barW + barGap);
    const branchTotal = ALL_CATS.reduce((sum, cat) => sum + (b.by_category[cat] || 0), 0);
    let stackY = barAreaY + barAreaH;

    ALL_CATS.forEach((cat, ci) => {{
      const val = b.by_category[cat] || 0;
      if (val === 0) return;
      const segH = (val / maxVal) * barAreaH;
      stackY -= segH;
      slide.addShape(pres.shapes.RECTANGLE, {{
        x: barX, y: stackY, w: barW, h: segH,
        fill: {{ color: catColors[ci] }}, line: {{ color: catColors[ci] }}
      }});
      if (segH > 0.25) {{
        slide.addText(`A$${{Math.round(val)}}`, {{
          x: barX, y: stackY + segH/2 - 0.1, w: barW, h: 0.2,
          fontSize: 7, color: WHITE, align: 'center', bold: true, margin: 0
        }});
      }}
    }});

    // Branch label below bar
    slide.addText(b.code, {{
      x: barX, y: barAreaY + barAreaH + 0.08, w: barW, h: 0.22,
      fontSize: 9, color: DARK_GREY, align: 'center', bold: true
    }});
    slide.addText(fmtAUD(branchTotal), {{
      x: barX - 0.1, y: barAreaY + barAreaH + 0.3, w: barW + 0.2, h: 0.2,
      fontSize: 7.5, color: MID_GREY, align: 'center'
    }});
  }});

  // Legend
  ALL_CATS.forEach((cat, ci) => {{
    const lx = barAreaX + ci * 1.55;
    slide.addShape(pres.shapes.RECTANGLE, {{
      x: lx, y: 4.18, w: 0.15, h: 0.15,
      fill: {{ color: catColors[ci] }}, line: {{ color: catColors[ci] }}
    }});
    slide.addText(cat, {{
      x: lx + 0.2, y: 4.16, w: 1.3, h: 0.2,
      fontSize: 8, color: MID_GREY
    }});
  }});

  // 4 callout stats on the right
  const callouts = [
    {{ label: 'Total AU Refunds', value: fmtAUD(AU_TOTAL) }},
    {{ label: 'Total Claims',     value: AU_COUNT.toString() }},
    {{ label: 'Avg Claim Value',  value: fmtAUD(AU_AVG) }},
    {{ label: 'Top Category',
       value: (function() {{
         let top = '', topV = 0;
         for (const [k,v] of Object.entries(AU_CAT_TOTALS)) {{ if(v>topV){{ top=k; topV=v; }} }}
         return top.split(' - ')[0] || top;
       }})()
    }},
  ];
  callouts.forEach((c, i) => {{
    const y = 0.9 + i * 0.88;
    slide.addShape(pres.shapes.RECTANGLE, {{
      x:6.9, y, w:2.8, h:0.8,
      fill:{{ color: NAVY }}, line:{{ color:'334155' }},
      shadow: makeShadow()
    }});
    slide.addText(c.value, {{
      x:6.95, y: y+0.04, w:2.7, h:0.44,
      fontSize:18, bold:true, color:WHITE, align:'center', margin:0
    }});
    slide.addText(c.label, {{
      x:6.95, y: y+0.5, w:2.7, h:0.26,
      fontSize:9, color:'94A3B8', align:'center', margin:0
    }});
  }});

  // Narrative banner
  slide.addShape(pres.shapes.RECTANGLE, {{
    x:0.4, y:4.5, w:9.2, h:0.75,
    fill:{{ color:'FFF7ED' }}, line:{{ color: ORANGE, pt:1 }}
  }});
  slide.addText(N.au_where_money_went, {{
    x:0.55, y:4.55, w:9.0, h:0.65,
    fontSize:11, color:DARK_GREY, italic:true, wrap:true
  }});

  if (N.au_where_money_went_talking_points) slide.addNotes(N.au_where_money_went_talking_points);
}})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — Equipment Gap (6 individual claim cards)
// ════════════════════════════════════════════════════════════════════════════
(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: LIGHT_GREY }};

  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: NAVY }}, line:{{ color: NAVY }} }});
  slide.addText('EQUIPMENT GAP', {{
    x:0.4, y:0.15, w:9, h:0.4,
    fontSize:20, bold:true, color: NAVY
  }});
  slide.addText('Largest equipment & gear refund per branch', {{
    x:0.4, y:0.52, w:9, h:0.28,
    fontSize:11, color: MID_GREY, italic:true
  }});

  const gaps = AU_EQUIP_GAPS.slice(0, 6);
  // pad to 6 if fewer
  while (gaps.length < 6) gaps.push(null);

  const cW = 2.9, cH = 1.85;
  const startX = 0.35, startY = 0.93, gapX = 0.16, gapY = 0.2;

  gaps.forEach((g, i) => {{
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x   = startX + col * (cW + gapX);
    const y   = startY + row * (cH + gapY);

    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y, w:cW, h:cH,
      fill:{{ color: WHITE }}, line:{{ color:'E2E8F0' }},
      shadow: makeShadow()
    }});

    if (!g) {{
      slide.addText('No equipment claims', {{
        x: x+0.1, y: y+0.7, w: cW-0.2, h:0.5,
        fontSize:10, color: MID_GREY, italic:true, align:'center'
      }});
      return;
    }}

    // Branch colour top strip
    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y, w:cW, h:0.08, fill:{{ color: g.color }}, line:{{ color: g.color }}
    }});

    slide.addText(g.branch_code, {{
      x: x+0.12, y: y+0.12, w:0.8, h:0.28,
      fontSize:16, bold:true, color: g.color, margin:0
    }});
    slide.addText(g.category, {{
      x: x+0.95, y: y+0.14, w: cW-1.1, h:0.22,
      fontSize:9, color: MID_GREY, italic:true, align:'right', margin:0
    }});
    slide.addText(fmtAUD(g.amount), {{
      x: x+0.12, y: y+0.42, w: cW-0.2, h:0.38,
      fontSize:22, bold:true, color: DARK_GREY, margin:0
    }});
    slide.addText(g.last_name || '', {{
      x: x+0.12, y: y+0.82, w: cW-0.2, h:0.22,
      fontSize:10, color: MID_GREY, margin:0
    }});
    const detail = (g.details || '').substring(0, 90);
    slide.addText(detail, {{
      x: x+0.12, y: y+1.05, w: cW-0.22, h:0.72,
      fontSize:8.5, color: MID_GREY, wrap:true, margin:0
    }});
  }});

  // Insight banner
  slide.addShape(pres.shapes.RECTANGLE, {{
    x:0.35, y:4.57, w:9.3, h:0.72,
    fill:{{ color:'1E293B' }}, line:{{ color: NAVY }}
  }});
  slide.addText(N.au_equipment_gap, {{
    x:0.5, y:4.62, w:9.1, h:0.62,
    fontSize:11, color:WHITE, italic:true, wrap:true
  }});

  if (N.au_equipment_gap_talking_points) slide.addNotes(N.au_equipment_gap_talking_points);
}})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDES 5+ — AU Branch Detail Slides (one per branch, worst offender first)
// ════════════════════════════════════════════════════════════════════════════
AU_BRANCHES.forEach(branch => {{
  const slide = pres.addSlide();
  slide.background = {{ color: WHITE }};

  // ── Left sidebar panel ──
  const sideW = 2.6;
  slide.addShape(pres.shapes.RECTANGLE, {{
    x:0, y:0, w:sideW, h:5.625,
    fill:{{ color: branch.color }}, line:{{ color: branch.color }}
  }});

  slide.addText(branch.code, {{
    x:0.1, y:0.25, w:sideW-0.15, h:0.7,
    fontSize:38, bold:true, color:WHITE, align:'left', charSpacing:1
  }});
  slide.addText(branch.name, {{
    x:0.1, y:0.9, w:sideW-0.15, h:0.32,
    fontSize:14, color:'FFFFFF', align:'left', italic:true, opacity:0.85
  }});
  slide.addShape(pres.shapes.LINE, {{
    x:0.15, y:1.3, w:sideW-0.3, h:0,
    line:{{ color:'FFFFFF', pt:1, transparency: 50 }}
  }});

  slide.addText('TOTAL REFUNDS', {{
    x:0.1, y:1.45, w:sideW-0.15, h:0.22,
    fontSize:8, color:'FFFFFF', charSpacing:2, opacity:0.7
  }});
  slide.addText(fmtAUD(branch.total), {{
    x:0.1, y:1.65, w:sideW-0.15, h:0.48,
    fontSize:26, bold:true, color:WHITE
  }});

  slide.addText(`${{branch.count}} Claim${{branch.count !== 1 ? 's' : ''}}`, {{
    x:0.1, y:2.15, w:sideW-0.15, h:0.3,
    fontSize:13, color:'FFFFFF', opacity:0.8
  }});

  // Month-over-month delta (white text on branch colour)
  const bd = deltaLine(branch.total, branch.prev_total, 'money');
  if (bd) {{
    slide.addText(bd.text, {{
      x:0.1, y:2.46, w:sideW-0.15, h:0.16,
      fontSize:8, bold:true, color:'FFFFFF'
    }});
  }}

  slide.addShape(pres.shapes.LINE, {{
    x:0.15, y:2.62, w:sideW-0.3, h:0,
    line:{{ color:'FFFFFF', pt:1, transparency: 50 }}
  }});

  // Sidebar mini bar chart (by category)
  const catLabels = Object.keys(branch.by_category);
  const catValues = catLabels.map(k => branch.by_category[k] || 0);
  const maxCatVal = Math.max(...catValues, 1);
  const barMaxW   = sideW - 0.35;

  catLabels.forEach((cat, ci) => {{
    const barY  = 2.78 + ci * 0.58;
    const barPct = catValues[ci] / maxCatVal;
    slide.addText(cat.length > 14 ? cat.substring(0,14)+'…' : cat, {{
      x:0.12, y: barY, w: sideW-0.2, h:0.2,
      fontSize:8, color:'FFFFFF', opacity:0.75
    }});
    // Bar bg
    slide.addShape(pres.shapes.RECTANGLE, {{
      x:0.12, y: barY+0.2, w: barMaxW, h:0.13,
      fill:{{ color:'FFFFFF', transparency:70 }}, line:{{ color:'FFFFFF', transparency:70 }}
    }});
    // Bar fill
    if (barPct > 0) {{
      slide.addShape(pres.shapes.RECTANGLE, {{
        x:0.12, y: barY+0.2, w: barMaxW * barPct, h:0.13,
        fill:{{ color:WHITE }}, line:{{ color:WHITE }}
      }});
    }}
    slide.addText(
      fmtAUD(catValues[ci])
        + ((branch.prev_by_category && PRIOR_SHORT)
            ? '  ·  ' + PRIOR_SHORT + ': ' + fmtAUD(branch.prev_by_category[cat] || 0)
            : ''),
      {{
        x:0.12, y: barY+0.35, w: sideW-0.2, h:0.17,
        fontSize:8.5, bold:true, color:WHITE
      }});
  }});

  // Sidebar narrative
  const sideNarrative = N.au_branch_detail[branch.code] || '';
  slide.addShape(pres.shapes.RECTANGLE, {{
    x:0.08, y:5.1, w:sideW-0.16, h:0.42,
    fill:{{ color:'FFFFFF', transparency:80 }}, line:{{ color:'FFFFFF', transparency:80 }}
  }});
  slide.addText(sideNarrative, {{
    x:0.12, y:5.1, w:sideW-0.22, h:0.42,
    fontSize:8, color:WHITE, italic:true, wrap:true
  }});

  // ── Right panel — individual claims ──
  const rightX = sideW + 0.25;
  const rightW = 10 - rightX - 0.25;

  slide.addText('INDIVIDUAL CLAIMS', {{
    x: rightX, y:0.18, w: rightW, h:0.3,
    fontSize:13, bold:true, color: NAVY, charSpacing:1
  }});
  slide.addShape(pres.shapes.LINE, {{
    x: rightX, y:0.52, w: rightW, h:0,
    line:{{ color:'E2E8F0', pt:1 }}
  }});

  // Show up to 5 claims
  const displayClaims = branch.claims.slice(0, 5);
  const claimH = 0.92;
  const claimStartY = 0.6;

  displayClaims.forEach((claim, ci) => {{
    const cy = claimStartY + ci * (claimH + 0.08);

    slide.addShape(pres.shapes.RECTANGLE, {{
      x: rightX, y: cy, w: rightW, h: claimH,
      fill:{{ color: LIGHT_GREY }}, line:{{ color:'E2E8F0', pt:0.75 }}
    }});
    // Category tag
    slide.addShape(pres.shapes.RECTANGLE, {{
      x: rightX, y: cy, w:0.07, h: claimH,
      fill:{{ color: branch.color }}, line:{{ color: branch.color }}
    }});

    // Row 1: Category tag + amount
    const catShort = (claim.category || 'Other').substring(0, 18);
    slide.addShape(pres.shapes.RECTANGLE, {{
      x: rightX + 0.12, y: cy+0.1, w:1.5, h:0.22,
      fill:{{ color: branch.color }}, line:{{ color: branch.color }}
    }});
    slide.addText(catShort, {{
      x: rightX + 0.12, y: cy+0.1, w:1.5, h:0.22,
      fontSize:8, bold:true, color:WHITE, align:'center', margin:0
    }});
    slide.addText(fmtAUD(claim.amount), {{
      x: rightX + rightW - 1.3, y: cy+0.08, w:1.25, h:0.28,
      fontSize:16, bold:true, color: DARK_GREY, align:'right', margin:0
    }});

    // Row 2: Name + rego
    const nameLine = `${{claim.last_name || ''}}${{claim.rego ? ' · ' + claim.rego : ''}}${{claim.res_num ? ' · Res #' + claim.res_num : ''}}`;
    slide.addText(nameLine, {{
      x: rightX + 0.12, y: cy+0.36, w: rightW-0.2, h:0.22,
      fontSize:9, color: MID_GREY, margin:0
    }});

    // Row 3: Details
    const details = (claim.details || '').substring(0, 120);
    slide.addText(details, {{
      x: rightX + 0.12, y: cy+0.58, w: rightW-0.2, h:0.28,
      fontSize:8.5, color: DARK_GREY, wrap:true, margin:0
    }});
  }});

  // If more claims exist, show count
  if (branch.claims.length > 5) {{
    slide.addText(`+ ${{branch.claims.length - 5}} more claim${{branch.claims.length - 5 !== 1 ? 's' : ''}} not shown`, {{
      x: rightX, y: claimStartY + 5 * (claimH + 0.08), w: rightW, h:0.25,
      fontSize:9, color: MID_GREY, italic:true
    }});
  }}

  // Presenter talking points for this branch
  const tp = (N.au_branch_talking_points || {{}})[branch.code];
  if (tp) slide.addNotes(tp);
}});


// ────────────────────────────────────────────────────────────────────────────
// Write file
// ────────────────────────────────────────────────────────────────────────────
const outputPath = {json.dumps(output_path)};
pres.writeFile({{ fileName: outputPath }})
  .then(() => {{
    console.log('SUCCESS:' + outputPath);
  }})
  .catch(err => {{
    console.error('ERROR:', err.message);
    process.exit(1);
  }});
"""
    return script


# ─── Main orchestrator ────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate AU/NZ Refund Awareness PPTX")
    parser.add_argument("--data",       help="Path to JSON data file (else reads stdin)")
    parser.add_argument("--output",     default="aunz_refund_report.pptx",
                        help="Output .pptx path")
    parser.add_argument("--start-date", help="(unused, kept for compatibility)")
    parser.add_argument("--end-date",   help="(unused, kept for compatibility)")
    parser.add_argument("--date-label", help="Human-readable date label e.g. 'June 2026' (auto-detected from Month column if omitted)")
    parser.add_argument("--api-key",    help="Anthropic API key (or set ANTHROPIC_API_KEY env var)")
    parser.add_argument("--skip-claude", action="store_true",
                        help="Skip Claude API calls (use placeholder narratives)")
    args = parser.parse_args()

    # 1. Load raw data
    print("Loading data...", file=sys.stderr)
    rows = load_data(args)
    rows = normalise_rows(rows)
    print(f"  Loaded {len(rows)} rows", file=sys.stderr)

    # 2. Split into report month and prior month using the Month column (K)
    report_rows, prior_rows, report_label, prior_label = split_rows_by_month(rows)
    print(f"  Report month: {report_label} ({len(report_rows)} rows)", file=sys.stderr)
    if prior_label:
        print(f"  Prior month:  {prior_label} ({len(prior_rows)} rows)", file=sys.stderr)
    else:
        print("  Prior month:  no data — comparison will be skipped", file=sys.stderr)

    # 3. Calculate all stats in Python
    print("Calculating statistics...", file=sys.stderr)
    stats = calculate_stats(report_rows)
    prior_stats = calculate_stats(prior_rows) if prior_rows else None
    stats = attach_prior_month(stats, prior_stats, prior_label)
    print(f"  AU: {stats['au']['grand_count']} claims = {stats['au']['grand_total']}", file=sys.stderr)
    print(f"  NZ: {stats['nz']['grand_count']} claims = {stats['nz']['grand_total']}", file=sys.stderr)

    # 4. Date label — prefer explicit arg, else auto-detected report month
    date_label = args.date_label or report_label or datetime.now().strftime("%B %Y")

    # 5. Get Claude narratives (or use placeholders)
    if args.skip_claude:
        print("Skipping Claude API (--skip-claude flag set)", file=sys.stderr)
        narratives = {
            "au_cover_insight":   "Australia refund data reviewed for this period.",
            "au_branch_cards":    {bc: f"{bc} branch summary." for bc in AU_BRANCHES},
            "au_where_money_went":"Category breakdown shows key refund drivers across branches.",
            "au_equipment_gap":   "Equipment-related refunds flagged for review.",
            "nz_cover_insight":   "New Zealand refund data reviewed for this period.",
            "nz_branch_cards":    {bc: f"{bc} branch summary." for bc in NZ_BRANCHES},
            "au_branch_detail":   {bc: f"Detail view for {AU_BRANCHES[bc]['name']}." for bc in AU_BRANCHES},
            "nz_branch_detail":   {bc: f"Detail view for {NZ_BRANCHES[bc]['name']}." for bc in NZ_BRANCHES},
            "au_cover_talking_points": "Placeholder talking points for the cover slide.",
            "au_branch_overview_talking_points": "Placeholder talking points for the branch overview slide.",
            "au_where_money_went_talking_points": "Placeholder talking points for the category slide.",
            "au_equipment_gap_talking_points": "Placeholder talking points for the equipment slide.",
            "au_branch_talking_points": {bc: f"Placeholder talking points for {AU_BRANCHES[bc]['name']}." for bc in AU_BRANCHES},
        }
    else:
        print("Calling Claude API for narrative text...", file=sys.stderr)
        narratives = get_claude_narratives(stats, date_label, prior_label, api_key=args.api_key)
        print("  Narratives received.", file=sys.stderr)

    # 6. Build Node.js script
    print("Building PptxGenJS script...", file=sys.stderr)
    script = build_pptx_script(stats, narratives, date_label, args.output)

    # 7. Write script to temp file and execute
    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False) as f:
        f.write(script)
        script_path = f.name

    print(f"Generating PPTX at {args.output}...", file=sys.stderr)
    # Run from the directory that has node_modules
    script_dir = os.path.dirname(os.path.abspath(__file__))
    result = subprocess.run(
        ["node", script_path],
        capture_output=True, text=True,
        cwd=script_dir
    )
    os.unlink(script_path)

    if result.returncode != 0:
        print("Node.js error:", result.stderr, file=sys.stderr)
        sys.exit(1)

    output_line = [l for l in result.stdout.splitlines() if l.startswith("SUCCESS:")]
    if output_line:
        print(output_line[0], file=sys.stderr)
    else:
        print(result.stdout, file=sys.stderr)

    print(f"\nDone! Output: {args.output}", file=sys.stderr)
    # Print output path to stdout for Make to capture
    print(args.output)


if __name__ == "__main__":
    main()
