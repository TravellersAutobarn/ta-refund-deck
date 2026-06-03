#!/usr/bin/env python3
"""
Travellers Autobarn — NZ Refund Awareness Deck Generator
=========================================================
Standalone script for New Zealand branches only (AUK, CHC).
Receives raw Google Sheets data (as JSON via stdin or file path arg),
calculates all stats in Python, calls Claude API for narrative text only,
then shells out to Node/PptxGenJS to build the final .pptx file.

Usage:
    echo '<json_data>' | python generate_nz_deck.py
    python generate_nz_deck.py --data data.json --output /path/to/output.pptx
    python generate_nz_deck.py --start-date 2026-05-01 --end-date 2026-05-15
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

# ─── Branch config ────────────────────────────────────────────────────────────

NZ_BRANCHES = {
    "AUK": {"name": "Auckland",     "color": "6D28D9", "region": "NZ"},  # indigo/purple
    "CHC": {"name": "Christchurch", "color": "DB2777", "region": "NZ"},  # magenta/pink
}

REFUND_CATEGORIES = ["DOR - GW", "Camp Gear - Prep", "Mechanical", "Kitchen - Internal"]

# Sheet column mapping (0-indexed)
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


# ─── Data loading ─────────────────────────────────────────────────────────────

def load_data(args):
    if args.data:
        with open(args.data) as f:
            rows = json.load(f)
    else:
        raw = sys.stdin.read().strip()
        rows = json.loads(raw)
    return rows


def parse_amount(val):
    if val is None or val == "":
        return 0.0
    s = str(val).replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date(val):
    if not val:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%d/%b/%Y"):
        try:
            return datetime.strptime(str(val).strip(), fmt).date()
        except ValueError:
            continue
    return None


# ─── Calculations ─────────────────────────────────────────────────────────────

def filter_by_date(rows, start_date, end_date):
    if not start_date and not end_date:
        return rows
    filtered = []
    for row in rows:
        d = parse_date(row[COL["date_entered"]])
        if d is None:
            continue
        if start_date and d < start_date:
            continue
        if end_date and d > end_date:
            continue
        filtered.append(row)
    return filtered


def is_nz_branch(branch_code):
    return str(branch_code).strip().upper() in NZ_BRANCHES


def calculate_stats(rows):
    if rows and isinstance(rows[0], dict):
        rows = [list(r.values()) for r in rows]

    nz_rows = [r for r in rows if is_nz_branch(r[COL["pickup_branch"]])]

    totals   = defaultdict(float)
    counts   = defaultdict(int)
    by_cat   = defaultdict(lambda: defaultdict(float))
    claims   = defaultdict(list)

    for row in nz_rows:
        bc  = str(row[COL["pickup_branch"]]).strip().upper()
        amt = parse_amount(row[COL["amount"]])
        cat = str(row[COL["refund_category"]]).strip()
        totals[bc]  += amt
        counts[bc]  += 1
        by_cat[bc][cat] += amt
        claims[bc].append({
            "res_num":   row[COL["res_num"]],
            "last_name": row[COL["last_name"]],
            "rego":      row[COL["rego"]],
            "fleet":     row[COL["fleet"]],
            "amount":    amt,
            "category":  cat,
            "details":   row[COL["details"]],
            "notes":     row[COL["notes"]],
            "date":      row[COL["date_entered"]],
        })

    grand_total = sum(totals.values())
    grand_count = sum(counts.values())

    branches = []
    for bc, cfg in NZ_BRANCHES.items():
        branch_total = totals.get(bc, 0.0)
        branch_count = counts.get(bc, 0)
        pct = (branch_total / grand_total * 100) if grand_total else 0
        cat_breakdown = {c: by_cat[bc].get(c, 0.0) for c in REFUND_CATEGORIES}
        branch_claims = sorted(claims.get(bc, []), key=lambda x: x["amount"], reverse=True)
        biggest_claim = branch_claims[0] if branch_claims else None
        branches.append({
            "code":         bc,
            "name":         cfg["name"],
            "color":        cfg["color"],
            "total":        round(branch_total, 2),
            "count":        branch_count,
            "pct_of_total": round(pct, 1),
            "by_category":  {k: round(v, 2) for k, v in cat_breakdown.items()},
            "claims":       branch_claims,
            "biggest_claim": biggest_claim,
        })

    category_totals = defaultdict(float)
    for row in nz_rows:
        cat = str(row[COL["refund_category"]]).strip()
        category_totals[cat] += parse_amount(row[COL["amount"]])

    # Equipment gap: biggest DOR/Gear claim per branch
    equipment_gaps = []
    for b in branches:
        dor_claims = [c for c in b["claims"] if "DOR" in c["category"].upper() or "GEAR" in c["category"].upper()]
        if dor_claims:
            top = max(dor_claims, key=lambda x: x["amount"])
            equipment_gaps.append({**top, "branch_code": b["code"], "branch_name": b["name"], "color": b["color"]})

    return {
        "grand_total":     round(grand_total, 2),
        "grand_count":     grand_count,
        "avg_claim":       round(grand_total / grand_count, 2) if grand_count else 0,
        "branches":        branches,
        "category_totals": {k: round(v, 2) for k, v in category_totals.items()},
        "largest_claim":   max((parse_amount(r[COL["amount"]]) for r in nz_rows), default=0),
        "equipment_gaps":  equipment_gaps,
    }


# ─── Claude API ───────────────────────────────────────────────────────────────

def get_claude_narratives(stats, date_label, api_key=None):
    client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    stats_summary = {
        "period": date_label,
        "total":  stats["grand_total"],
        "count":  stats["grand_count"],
        "avg_claim": stats["avg_claim"],
        "category_totals": stats["category_totals"],
        "branches": [
            {
                "code":  b["code"],
                "name":  b["name"],
                "total": b["total"],
                "count": b["count"],
                "pct":   b["pct_of_total"],
                "by_category": b["by_category"],
            }
            for b in stats["branches"]
        ],
    }

    prompt = f"""You are writing concise insights for a refund awareness presentation for Travellers Autobarn New Zealand.
All numbers below are pre-calculated — DO NOT recalculate or change any figures.
Return ONLY valid JSON with the exact keys listed. No markdown, no code fences.

STATS:
{json.dumps(stats_summary, indent=2)}

Return JSON with these exact keys:
{{
  "nz_cover_insight": "One sharp sentence (max 20 words) about the NZ refund picture this period.",
  "nz_branch_cards": {{
    "AUK": "One italic insight sentence for Auckland card (max 15 words)",
    "CHC": "One italic insight sentence for Christchurch card (max 15 words)"
  }},
  "nz_where_money_went": "One sentence (max 25 words) summarising the category breakdown across NZ branches.",
  "nz_equipment_gap": "One sentence (max 20 words) on the equipment/gear refund theme.",
  "nz_branch_detail": {{
    "AUK": "Two sentences (max 30 words total) for the Auckland detail slide narrative.",
    "CHC": "Two sentences (max 30 words total) for the Christchurch detail slide narrative."
  }}
}}"""

    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=800,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ─── PptxGenJS script builder ─────────────────────────────────────────────────

def build_pptx_script(stats, narratives, date_label, output_path):
    branches_js   = json.dumps(stats["branches"])
    narratives_js = json.dumps(narratives)

    script = f"""
'use strict';
const pptxgen = require('pptxgenjs');

// ─── Embedded data ─────────────────────────────────────────────────────────
const DATE_LABEL    = {json.dumps(date_label)};
const NZ_TOTAL      = {stats["grand_total"]};
const NZ_COUNT      = {stats["grand_count"]};
const NZ_AVG        = {stats["avg_claim"]};
const NZ_CAT_TOTALS = {json.dumps(stats["category_totals"])};
const NZ_BRANCHES   = {branches_js};
const NZ_EQUIP_GAPS = {json.dumps(stats["equipment_gaps"])};
const N             = {narratives_js};

// ─── Helpers ──────────────────────────────────────────────────────────────
const fmtNZD = v => 'NZ$' + Number(v).toLocaleString('en-NZ', {{minimumFractionDigits:0, maximumFractionDigits:0}});
const fmtPct = v => v.toFixed(1) + '%';

const NAVY       = '0F172A';
const WHITE      = 'FFFFFF';
const ORANGE     = 'F97316';
const LIGHT_GREY = 'F8FAFC';
const MID_GREY   = '64748B';
const DARK_GREY  = '1E293B';

const makeShadow = () => ({{ type: 'outer', blur: 8, offset: 2, angle: 135, color: '000000', opacity: 0.12 }});

// ─── Presentation setup ───────────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title  = `Travellers Autobarn — NZ Refund Awareness ${{DATE_LABEL}}`;

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — NZ Cover
// ════════════════════════════════════════════════════════════════════════════
(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: NAVY }};

  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: ORANGE }}, line:{{ color: ORANGE }} }});

  slide.addText('TRAVELLERS AUTOBARN', {{
    x:0.5, y:0.2, w:6, h:0.4,
    fontSize:11, bold:true, color:'94A3B8', charSpacing:3, align:'left'
  }});
  slide.addText('NEW ZEALAND', {{
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

  // Big stat callouts
  const stats = [
    {{ label:'Total Refunds', value: fmtNZD(NZ_TOTAL) }},
    {{ label:'Total Claims',  value: NZ_COUNT.toString() }},
    {{ label:'Average Claim', value: fmtNZD(NZ_AVG) }},
  ];
  stats.forEach((s, i) => {{
    const x = 0.5 + i * 3.2;
    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y:2.35, w:2.9, h:1.05,
      fill:{{ color:'1E293B' }}, line:{{ color:'334155' }},
      shadow: makeShadow()
    }});
    slide.addText(s.value, {{
      x, y:2.38, w:2.9, h:0.55,
      fontSize:28, bold:true, color:WHITE, align:'center', margin:0
    }});
    slide.addText(s.label, {{
      x, y:2.88, w:2.9, h:0.35,
      fontSize:10, color:'64748B', align:'center', margin:0
    }});
  }});

  // Branch dots
  slide.addText('Branches covered:', {{
    x:0.5, y:3.62, w:5, h:0.28,
    fontSize:10, color:'64748B', align:'left'
  }});
  NZ_BRANCHES.forEach((b, i) => {{
    const x = 0.5 + i * 3.0;
    const y = 3.9;
    slide.addShape(pres.shapes.OVAL, {{
      x, y: y+0.04, w:0.14, h:0.14,
      fill:{{ color: b.color }}, line:{{ color: b.color }}
    }});
    slide.addText(`${{b.code}} — ${{b.name}}`, {{
      x: x+0.2, y, w:2.5, h:0.28,
      fontSize:11, color:WHITE, align:'left'
    }});
  }});

  // Key insight box
  slide.addShape(pres.shapes.RECTANGLE, {{
    x:6.3, y:3.5, w:3.4, h:1.75,
    fill:{{ color:'7C2D12' }}, line:{{ color: ORANGE, pt:1.5 }},
    shadow: makeShadow()
  }});
  slide.addText('KEY INSIGHT', {{
    x:6.4, y:3.55, w:3.2, h:0.3,
    fontSize:9, bold:true, color: ORANGE, charSpacing:2
  }});
  slide.addText(N.nz_cover_insight, {{
    x:6.4, y:3.85, w:3.2, h:1.3,
    fontSize:12, color:WHITE, align:'left', wrap:true
  }});

  slide.addShape(pres.shapes.LINE, {{
    x:0.5, y:5.4, w:9, h:0, line:{{ color:'1E293B', pt:1 }}
  }});
  slide.addText('CONFIDENTIAL — INTERNAL USE ONLY', {{
    x:0.5, y:5.43, w:9, h:0.2,
    fontSize:8, color:'334155', align:'center'
  }});
}})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — Refunds by NZ Branch (2 cards — full width)
// ════════════════════════════════════════════════════════════════════════════
(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: LIGHT_GREY }};

  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: NAVY }}, line:{{ color: NAVY }} }});
  slide.addText('REFUNDS BY BRANCH', {{
    x:0.4, y:0.15, w:9, h:0.4,
    fontSize:20, bold:true, color: NAVY, align:'left'
  }});
  slide.addText(`New Zealand — ${{DATE_LABEL}}`, {{
    x:0.4, y:0.52, w:9, h:0.3,
    fontSize:11, color: MID_GREY, italic:true
  }});

  // 2 wide cards side by side
  const cW = 4.55, cH = 2.8;
  const startX = 0.4, startY = 1.0, gap = 0.1;

  NZ_BRANCHES.forEach((b, i) => {{
    const x   = startX + i * (cW + gap);
    const ins = N.nz_branch_cards[b.code] || '';

    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y:startY, w:cW, h:cH,
      fill:{{ color: WHITE }}, line:{{ color:'E2E8F0', pt:1 }},
      shadow: makeShadow()
    }});
    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y:startY, w:0.07, h:cH,
      fill:{{ color: b.color }}, line:{{ color: b.color }}
    }});

    slide.addText(b.code, {{
      x: x+0.18, y: startY+0.12, w:1.4, h:0.48,
      fontSize:30, bold:true, color: b.color, margin:0
    }});
    slide.addText(b.name, {{
      x: x+0.18, y: startY+0.58, w: cW-0.3, h:0.26,
      fontSize:13, color: MID_GREY, italic:true, margin:0
    }});

    slide.addText(fmtNZD(b.total), {{
      x: x+0.18, y: startY+0.92, w: cW-0.3, h:0.52,
      fontSize:32, bold:true, color: DARK_GREY, margin:0
    }});
    slide.addText(`${{b.count}} claim${{b.count !== 1 ? 's' : ''}}`, {{
      x: x+0.18, y: startY+1.46, w:2.0, h:0.3,
      fontSize:14, color: MID_GREY, margin:0
    }});

    // % badge
    slide.addShape(pres.shapes.RECTANGLE, {{
      x: x+cW-1.35, y: startY+1.44, w:1.2, h:0.32,
      fill:{{ color: b.color }}, line:{{ color: b.color }}
    }});
    slide.addText(fmtPct(b.pct_of_total), {{
      x: x+cW-1.35, y: startY+1.44, w:1.2, h:0.32,
      fontSize:13, bold:true, color: WHITE, align:'center', margin:0
    }});

    // Category mini breakdown
    const cats = Object.entries(b.by_category).filter(([,v]) => v > 0);
    cats.forEach(([cat, val], ci) => {{
      const cy = startY + 1.9 + ci * 0.22;
      const catLabel = cat.length > 18 ? cat.substring(0,18)+'…' : cat;
      slide.addText(`${{catLabel}}: ${{fmtNZD(val)}}`, {{
        x: x+0.18, y: cy, w: cW-0.3, h:0.2,
        fontSize:9, color: MID_GREY, margin:0
      }});
    }});

    // Italic insight
    slide.addText(ins, {{
      x: x+0.15, y: startY+cH-0.45, w: cW-0.3, h:0.38,
      fontSize:9.5, color: MID_GREY, italic:true, wrap:true, margin:0
    }});
  }});

  // NZ vs AU comparison note
  slide.addShape(pres.shapes.RECTANGLE, {{
    x:0.4, y:4.0, w:9.2, h:0.55,
    fill:{{ color:'F0F9FF' }}, line:{{ color:'BAE6FD', pt:1 }}
  }});
  slide.addText(`NZ Total this period: ${{fmtNZD(NZ_TOTAL)}} across ${{NZ_COUNT}} claims — Average: ${{fmtNZD(NZ_AVG)}} per claim`, {{
    x:0.55, y:4.05, w:9.0, h:0.45,
    fontSize:11, color:DARK_GREY, align:'center', italic:true
  }});
}})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — Where the Money Went (category breakdown + callouts)
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
  const chartData = Object.keys(NZ_CAT_TOTALS).map((cat, ci) => ({{
    name:   cat,
    labels: NZ_BRANCHES.map(b => b.code),
    values: NZ_BRANCHES.map(b => b.by_category[cat] || 0)
  }}));

  slide.addChart(pres.charts.BAR, chartData, {{
    x:0.4, y:0.9, w:6.2, h:3.5,
    barDir: 'col',
    barGrouping: 'stacked',
    catAxisLabelFontSize: 11,
    chartColors: catColors,
    showLegend: true,
    legendPos: 'b',
    legendFontSize: 9,
    chartArea: {{ fill: {{ color: WHITE }} }},
    valAxisLabelColor: MID_GREY,
    catAxisLabelColor: MID_GREY,
    valGridLine: {{ color: 'E2E8F0', size: 0.5 }},
    catGridLine: {{ style: 'none' }},
    showValue: true,
    dataLabelFontSize: 9,
  }});

  // Callout stats
  const callouts = [
    {{ label:'Total NZ Refunds', value: fmtNZD(NZ_TOTAL) }},
    {{ label:'Total Claims',     value: NZ_COUNT.toString() }},
    {{ label:'Avg Claim Value',  value: fmtNZD(NZ_AVG) }},
    {{ label:'Top Category',
       value: (function() {{
         let top = '', topV = 0;
         for (const [k,v] of Object.entries(NZ_CAT_TOTALS)) {{ if(v>topV){{ top=k; topV=v; }} }}
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
    x:0.4, y:4.55, w:9.2, h:0.75,
    fill:{{ color:'FFF7ED' }}, line:{{ color: ORANGE, pt:1 }}
  }});
  slide.addText(N.nz_where_money_went, {{
    x:0.55, y:4.6, w:9.0, h:0.65,
    fontSize:11, color:DARK_GREY, italic:true, wrap:true
  }});
}})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — Equipment Gap (NZ)
// ════════════════════════════════════════════════════════════════════════════
(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: LIGHT_GREY }};

  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: NAVY }}, line:{{ color: NAVY }} }});
  slide.addText('EQUIPMENT GAP', {{
    x:0.4, y:0.15, w:9, h:0.4,
    fontSize:20, bold:true, color: NAVY
  }});
  slide.addText('Largest equipment & gear refund per NZ branch', {{
    x:0.4, y:0.52, w:9, h:0.28,
    fontSize:11, color: MID_GREY, italic:true
  }});

  const gaps = NZ_EQUIP_GAPS.slice(0, 2);
  while (gaps.length < 2) gaps.push(null);

  // 2 wide cards
  const cW = 4.55, cH = 3.2;
  const startX = 0.4, startY = 0.93;

  gaps.forEach((g, i) => {{
    const x = startX + i * (cW + 0.1);

    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y:startY, w:cW, h:cH,
      fill:{{ color: WHITE }}, line:{{ color:'E2E8F0' }},
      shadow: makeShadow()
    }});

    if (!g) {{
      const bc = NZ_EQUIP_GAPS.length === 0 ? NZ_BRANCHES[i] : null;
      slide.addText('No equipment claims this period', {{
        x: x+0.1, y: startY+1.3, w: cW-0.2, h:0.5,
        fontSize:11, color: MID_GREY, italic:true, align:'center'
      }});
      return;
    }}

    slide.addShape(pres.shapes.RECTANGLE, {{
      x, y:startY, w:cW, h:0.08,
      fill:{{ color: g.color }}, line:{{ color: g.color }}
    }});
    slide.addText(g.branch_code, {{
      x: x+0.15, y: startY+0.15, w:1.0, h:0.35,
      fontSize:20, bold:true, color: g.color, margin:0
    }});
    slide.addText(g.branch_name, {{
      x: x+0.15, y: startY+0.48, w: cW-0.3, h:0.24,
      fontSize:12, color: MID_GREY, italic:true, margin:0
    }});
    slide.addText(g.category, {{
      x: x+cW-1.8, y: startY+0.17, w:1.6, h:0.22,
      fontSize:9, color: MID_GREY, italic:true, align:'right', margin:0
    }});
    slide.addText(fmtNZD(g.amount), {{
      x: x+0.15, y: startY+0.8, w: cW-0.3, h:0.55,
      fontSize:32, bold:true, color: DARK_GREY, margin:0
    }});
    slide.addText(g.last_name || '', {{
      x: x+0.15, y: startY+1.4, w: cW-0.3, h:0.26,
      fontSize:12, color: MID_GREY, margin:0
    }});
    const detail = (g.details || '').substring(0, 140);
    slide.addText(detail, {{
      x: x+0.15, y: startY+1.7, w: cW-0.25, h:1.3,
      fontSize:9.5, color: MID_GREY, wrap:true, margin:0
    }});
  }});

  // Insight banner
  slide.addShape(pres.shapes.RECTANGLE, {{
    x:0.4, y:4.3, w:9.2, h:0.72,
    fill:{{ color:'1E293B' }}, line:{{ color: NAVY }}
  }});
  slide.addText(N.nz_equipment_gap, {{
    x:0.55, y:4.35, w:9.0, h:0.62,
    fontSize:11, color:WHITE, italic:true, wrap:true
  }});
}})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDES 5–6 — NZ Branch Detail (one per branch: AUK, CHC)
// ════════════════════════════════════════════════════════════════════════════
NZ_BRANCHES.forEach(branch => {{
  const slide = pres.addSlide();
  slide.background = {{ color: WHITE }};

  // Left sidebar
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
    fontSize:14, color:'FFFFFF', align:'left', italic:true
  }});
  slide.addShape(pres.shapes.LINE, {{
    x:0.15, y:1.3, w:sideW-0.3, h:0,
    line:{{ color:'FFFFFF', pt:1, transparency: 50 }}
  }});

  slide.addText('TOTAL REFUNDS', {{
    x:0.1, y:1.45, w:sideW-0.15, h:0.22,
    fontSize:8, color:'FFFFFF', charSpacing:2
  }});
  slide.addText(fmtNZD(branch.total), {{
    x:0.1, y:1.65, w:sideW-0.15, h:0.48,
    fontSize:26, bold:true, color:WHITE
  }});
  slide.addText(`${{branch.count}} Claim${{branch.count !== 1 ? 's' : ''}}`, {{
    x:0.1, y:2.15, w:sideW-0.15, h:0.3,
    fontSize:13, color:'FFFFFF'
  }});
  slide.addShape(pres.shapes.LINE, {{
    x:0.15, y:2.55, w:sideW-0.3, h:0,
    line:{{ color:'FFFFFF', pt:1, transparency: 50 }}
  }});

  // Sidebar category bars
  const catLabels = Object.keys(branch.by_category);
  const catValues = catLabels.map(k => branch.by_category[k] || 0);
  const maxCatVal = Math.max(...catValues, 1);
  const barMaxW   = sideW - 0.35;

  catLabels.forEach((cat, ci) => {{
    const barY   = 2.75 + ci * 0.6;
    const barPct = catValues[ci] / maxCatVal;
    slide.addText(cat.length > 14 ? cat.substring(0,14)+'…' : cat, {{
      x:0.12, y: barY, w: sideW-0.2, h:0.2,
      fontSize:8, color:'FFFFFF'
    }});
    slide.addShape(pres.shapes.RECTANGLE, {{
      x:0.12, y: barY+0.22, w: barMaxW, h:0.14,
      fill:{{ color:'FFFFFF', transparency:70 }}, line:{{ color:'FFFFFF', transparency:70 }}
    }});
    if (barPct > 0) {{
      slide.addShape(pres.shapes.RECTANGLE, {{
        x:0.12, y: barY+0.22, w: barMaxW * barPct, h:0.14,
        fill:{{ color:WHITE }}, line:{{ color:WHITE }}
      }});
    }}
    slide.addText(fmtNZD(catValues[ci]), {{
      x:0.12, y: barY+0.38, w: sideW-0.2, h:0.18,
      fontSize:8.5, bold:true, color:WHITE
    }});
  }});

  // Sidebar narrative
  const sideNarrative = N.nz_branch_detail[branch.code] || '';
  slide.addShape(pres.shapes.RECTANGLE, {{
    x:0.08, y:5.1, w:sideW-0.16, h:0.42,
    fill:{{ color:'FFFFFF', transparency:80 }}, line:{{ color:'FFFFFF', transparency:80 }}
  }});
  slide.addText(sideNarrative, {{
    x:0.12, y:5.1, w:sideW-0.22, h:0.42,
    fontSize:8, color:WHITE, italic:true, wrap:true
  }});

  // Right panel — individual claims
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

  const displayClaims = branch.claims.slice(0, 5);
  const claimH = 0.92;
  const claimStartY = 0.6;

  displayClaims.forEach((claim, ci) => {{
    const cy = claimStartY + ci * (claimH + 0.08);

    slide.addShape(pres.shapes.RECTANGLE, {{
      x: rightX, y: cy, w: rightW, h: claimH,
      fill:{{ color: LIGHT_GREY }}, line:{{ color:'E2E8F0', pt:0.75 }}
    }});
    slide.addShape(pres.shapes.RECTANGLE, {{
      x: rightX, y: cy, w:0.07, h: claimH,
      fill:{{ color: branch.color }}, line:{{ color: branch.color }}
    }});

    const catShort = (claim.category || 'Other').substring(0, 18);
    slide.addShape(pres.shapes.RECTANGLE, {{
      x: rightX + 0.12, y: cy+0.1, w:1.5, h:0.22,
      fill:{{ color: branch.color }}, line:{{ color: branch.color }}
    }});
    slide.addText(catShort, {{
      x: rightX + 0.12, y: cy+0.1, w:1.5, h:0.22,
      fontSize:8, bold:true, color:WHITE, align:'center', margin:0
    }});
    slide.addText(fmtNZD(claim.amount), {{
      x: rightX + rightW - 1.3, y: cy+0.08, w:1.25, h:0.28,
      fontSize:16, bold:true, color: DARK_GREY, align:'right', margin:0
    }});

    const nameLine = `${{claim.last_name || ''}}${{claim.rego ? ' · ' + claim.rego : ''}}${{claim.res_num ? ' · Res #' + claim.res_num : ''}}`;
    slide.addText(nameLine, {{
      x: rightX + 0.12, y: cy+0.36, w: rightW-0.2, h:0.22,
      fontSize:9, color: MID_GREY, margin:0
    }});

    const details = (claim.details || '').substring(0, 120);
    slide.addText(details, {{
      x: rightX + 0.12, y: cy+0.58, w: rightW-0.2, h:0.28,
      fontSize:8.5, color: DARK_GREY, wrap:true, margin:0
    }});
  }});

  if (branch.claims.length > 5) {{
    slide.addText(`+ ${{branch.claims.length - 5}} more claim${{branch.claims.length - 5 !== 1 ? 's' : ''}} not shown`, {{
      x: rightX, y: claimStartY + 5 * (claimH + 0.08), w: rightW, h:0.25,
      fontSize:9, color: MID_GREY, italic:true
    }});
  }}
}});

// ─── Write file ──────────────────────────────────────────────────────────────
const outputPath = {json.dumps(output_path)};
pres.writeFile({{ fileName: outputPath }})
  .then(() => {{ console.log('SUCCESS:' + outputPath); }})
  .catch(err => {{ console.error('ERROR:', err.message); process.exit(1); }});
"""
    return script


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate NZ Refund Awareness PPTX")
    parser.add_argument("--data",        help="Path to JSON data file (else reads stdin)")
    parser.add_argument("--output",      default="nz_refund_report.pptx")
    parser.add_argument("--start-date",  help="Filter start date YYYY-MM-DD")
    parser.add_argument("--end-date",    help="Filter end date YYYY-MM-DD")
    parser.add_argument("--date-label",  help="Human-readable date range e.g. '1–15 June 2026'")
    parser.add_argument("--api-key",     help="Anthropic API key")
    parser.add_argument("--skip-claude", action="store_true")
    args = parser.parse_args()

    print("Loading data...", file=sys.stderr)
    rows = load_data(args)
    print(f"  Loaded {len(rows)} rows", file=sys.stderr)

    start_date = parse_date(args.start_date) if args.start_date else None
    end_date   = parse_date(args.end_date)   if args.end_date   else None
    if start_date or end_date:
        rows = filter_by_date(rows, start_date, end_date)
        print(f"  After date filter: {len(rows)} rows", file=sys.stderr)

    print("Calculating statistics...", file=sys.stderr)
    stats = calculate_stats(rows)
    print(f"  NZ: {stats['grand_count']} claims = NZ${stats['grand_total']}", file=sys.stderr)

    date_label = args.date_label
    if not date_label:
        if start_date and end_date:
            date_label = f"{start_date.strftime('%d %b')}–{end_date.strftime('%d %b %Y')}"
        else:
            date_label = datetime.now().strftime("%B %Y")

    if args.skip_claude:
        print("Skipping Claude API...", file=sys.stderr)
        narratives = {
            "nz_cover_insight":   "New Zealand refund data reviewed for this period.",
            "nz_branch_cards":    {bc: f"{bc} branch summary." for bc in NZ_BRANCHES},
            "nz_where_money_went":"Category breakdown shows key refund drivers across NZ branches.",
            "nz_equipment_gap":   "Equipment-related refunds flagged for review.",
            "nz_branch_detail":   {bc: f"Detail view for {NZ_BRANCHES[bc]['name']}." for bc in NZ_BRANCHES},
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
    result = subprocess.run(
        ["node", script_path],
        capture_output=True, text=True,
        cwd=script_dir
    )
    os.unlink(script_path)

    if result.returncode != 0:
        print("Node.js error:", result.stderr, file=sys.stderr)
        sys.exit(1)

    print(f"\nDone! Output: {args.output}", file=sys.stderr)
    print(args.output)


if __name__ == "__main__":
    main()
