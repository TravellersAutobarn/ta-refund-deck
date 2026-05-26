#!/usr/bin/env python3
"""
Travellers Autobarn — AU/NZ Refund Awareness Deck Generator
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

AU_BRANCHES = {
    "BNE": {"name": "Brisbane",  "color": "7C3AED", "region": "AU"},
    "CNS": {"name": "Cairns",    "color": "065F46", "region": "AU"},
    "MEL": {"name": "Melbourne", "color": "D97706", "region": "AU"},
    "SYD": {"name": "Sydney",    "color": "16A34A", "region": "AU"},
    "PER": {"name": "Perth",     "color": "2563EB", "region": "AU"},
    "DAR": {"name": "Darwin",    "color": "DC2626", "region": "AU"},
}

NZ_BRANCHES = {
    "AUK": {"name": "Auckland",     "color": "6D28D9", "region": "NZ"},
    "CHC": {"name": "Christchurch", "color": "DB2777", "region": "NZ"},
}

ALL_BRANCHES = {**AU_BRANCHES, **NZ_BRANCHES}
REFUND_CATEGORIES = ["DOR-GW", "Camp Gear-Prep", "Mechanical", "Kitchen-Internal"]

COL = {
    "no": 0, "res_num": 1, "category": 2, "rego": 3, "fleet": 4,
    "pickup": 5, "last_name": 6, "date_entered": 7, "entered_by": 8,
    "follow_up": 9, "month": 10, "pickup_branch": 11,
    "refund_category": 12, "amount": 13, "details": 14, "notes": 15,
}

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

def classify_region(branch_code):
    bc = str(branch_code).strip().upper()
    if bc in AU_BRANCHES:
        return "AU"
    if bc in NZ_BRANCHES:
        return "NZ"
    return None

def calculate_stats(rows):
    au_rows = [r for r in rows if classify_region(r[COL["pickup_branch"]]) == "AU"]
    nz_rows = [r for r in rows if classify_region(r[COL["pickup_branch"]]) == "NZ"]

    def branch_stats(branch_rows, branch_map):
        totals = defaultdict(float)
        counts = defaultdict(int)
        by_cat = defaultdict(lambda: defaultdict(float))
        claims = defaultdict(list)

        for row in branch_rows:
            bc  = str(row[COL["pickup_branch"]]).strip().upper()
            amt = parse_amount(row[COL["amount"]])
            cat = str(row[COL["refund_category"]]).strip()
            totals[bc] += amt
            counts[bc] += 1
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
        for bc, cfg in branch_map.items():
            branch_total = totals.get(bc, 0.0)
            branch_count = counts.get(bc, 0)
            pct = (branch_total / grand_total * 100) if grand_total else 0
            cat_breakdown = {c: by_cat[bc].get(c, 0.0) for c in REFUND_CATEGORIES}
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

        category_totals = defaultdict(float)
        for row in branch_rows:
            cat = str(row[COL["refund_category"]]).strip()
            category_totals[cat] += parse_amount(row[COL["amount"]])

        return {
            "grand_total":     round(grand_total, 2),
            "grand_count":     grand_count,
            "branches":        branches,
            "category_totals": {k: round(v, 2) for k, v in category_totals.items()},
            "avg_claim":       round(grand_total / grand_count, 2) if grand_count else 0,
            "largest_claim":   max((parse_amount(r[COL["amount"]]) for r in branch_rows), default=0),
        }

    au_stats = branch_stats(au_rows, AU_BRANCHES)
    nz_stats = branch_stats(nz_rows, NZ_BRANCHES)

    def equipment_gap(branch_list):
        gaps = []
        for b in branch_list:
            dor_claims = [c for c in b["claims"] if "DOR" in c["category"].upper() or "GEAR" in c["category"].upper()]
            if dor_claims:
                top = max(dor_claims, key=lambda x: x["amount"])
                gaps.append({**top, "branch_code": b["code"], "branch_name": b["name"], "color": b["color"]})
        return gaps

    au_stats["equipment_gaps"] = equipment_gap(au_stats["branches"])
    nz_stats["equipment_gaps"] = equipment_gap(nz_stats["branches"])

    return {"au": au_stats, "nz": nz_stats}

def get_claude_narratives(stats, date_label, api_key=None):
    client = anthropic.Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))
    au = stats["au"]
    nz = stats["nz"]

    stats_summary = {
        "period": date_label,
        "au": {
            "total": au["grand_total"],
            "count": au["grand_count"],
            "avg_claim": au["avg_claim"],
            "category_totals": au["category_totals"],
            "branches": [{"code": b["code"], "name": b["name"], "total": b["total"],
                          "count": b["count"], "pct": b["pct_of_total"], "by_category": b["by_category"]}
                         for b in au["branches"]],
        },
        "nz": {
            "total": nz["grand_total"],
            "count": nz["grand_count"],
            "avg_claim": nz["avg_claim"],
            "category_totals": nz["category_totals"],
            "branches": [{"code": b["code"], "name": b["name"], "total": b["total"],
                          "count": b["count"], "pct": b["pct_of_total"], "by_category": b["by_category"]}
                         for b in nz["branches"]],
        },
    }

    prompt = f"""You are writing concise insights for a refund awareness presentation for Travellers Autobarn.
All numbers below are pre-calculated — DO NOT recalculate or change any figures.
Return ONLY valid JSON with the exact keys listed. No markdown, no code fences.

STATS:
{json.dumps(stats_summary, indent=2)}

Return JSON with these exact keys:
{{
  "au_cover_insight": "One sharp sentence (max 20 words) about the AU refund picture this period.",
  "au_branch_cards": {{"BNE": "max 15 words", "CNS": "...", "MEL": "...", "SYD": "...", "PER": "...", "DAR": "..."}},
  "au_where_money_went": "One sentence (max 25 words) summarising the category breakdown across AU branches.",
  "au_equipment_gap": "One sentence (max 20 words) on the equipment/gear refund theme.",
  "nz_cover_insight": "One sharp sentence (max 20 words) about the NZ refund picture.",
  "nz_branch_cards": {{"AUK": "max 15 words", "CHC": "..."}},
  "au_branch_detail": {{"BNE": "max 30 words", "CNS": "...", "MEL": "...", "SYD": "...", "PER": "...", "DAR": "..."}},
  "nz_branch_detail": {{"AUK": "max 30 words", "CHC": "..."}}
}}"""

    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())

def build_pptx_script(stats, narratives, date_label, output_path):
    au = stats["au"]
    nz = stats["nz"]
    au_branches_js = json.dumps(au["branches"])
    nz_branches_js = json.dumps(nz["branches"])
    narratives_js  = json.dumps(narratives)

    script = f"""
'use strict';
const pptxgen = require('pptxgenjs');

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
const N             = {narratives_js};

const fmtAUD = v => 'A$' + Number(v).toLocaleString('en-AU', {{minimumFractionDigits:0, maximumFractionDigits:0}});
const fmtNZD = v => 'NZ$' + Number(v).toLocaleString('en-NZ', {{minimumFractionDigits:0, maximumFractionDigits:0}});
const fmtPct = v => v.toFixed(1) + '%';

const NAVY       = '0F172A';
const WHITE      = 'FFFFFF';
const ORANGE     = 'F97316';
const LIGHT_GREY = 'F8FAFC';
const MID_GREY   = '64748B';
const DARK_GREY  = '1E293B';
const makeShadow = () => ({{ type: 'outer', blur: 8, offset: 2, angle: 135, color: '000000', opacity: 0.12 }});

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title  = `Travellers Autobarn — AU/NZ Refund Awareness ${{DATE_LABEL}}`;

(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: NAVY }};
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: ORANGE }}, line:{{ color: ORANGE }} }});
  slide.addText('TRAVELLERS AUTOBARN', {{ x:0.5, y:0.2, w:6, h:0.4, fontSize:11, bold:true, color:'94A3B8', charSpacing:3 }});
  slide.addText('AUSTRALIA', {{ x:0.5, y:0.55, w:9, h:0.9, fontSize:52, bold:true, color:WHITE, charSpacing:2 }});
  slide.addText('Refund Awareness Report', {{ x:0.5, y:1.35, w:9, h:0.45, fontSize:20, color:'94A3B8', italic:true }});
  slide.addText(DATE_LABEL, {{ x:0.5, y:1.75, w:9, h:0.35, fontSize:13, color:'64748B' }});
  const stats = [
    {{ label:'Total Refunds', value: fmtAUD(AU_TOTAL) }},
    {{ label:'Total Claims',  value: AU_COUNT.toString() }},
    {{ label:'Average Claim', value: fmtAUD(AU_AVG) }},
  ];
  stats.forEach((s, i) => {{
    const x = 0.5 + i * 3.2;
    slide.addShape(pres.shapes.RECTANGLE, {{ x, y:2.35, w:2.9, h:1.05, fill:{{ color:'1E293B' }}, line:{{ color:'334155' }}, shadow: makeShadow() }});
    slide.addText(s.value, {{ x, y:2.38, w:2.9, h:0.55, fontSize:28, bold:true, color:WHITE, align:'center', margin:0 }});
    slide.addText(s.label, {{ x, y:2.88, w:2.9, h:0.35, fontSize:10, color:'64748B', align:'center', margin:0 }});
  }});
  slide.addText('Branches covered:', {{ x:0.5, y:3.62, w:5, h:0.28, fontSize:10, color:'64748B' }});
  AU_BRANCHES.forEach((b, i) => {{
    const col = i < 3 ? 0 : 1;
    const row = i % 3;
    const x = 0.5 + col * 3.0;
    const y = 3.9 + row * 0.32;
    slide.addShape(pres.shapes.OVAL, {{ x, y: y+0.04, w:0.14, h:0.14, fill:{{ color: b.color }}, line:{{ color: b.color }} }});
    slide.addText(`${{b.code}} — ${{b.name}}`, {{ x: x+0.2, y, w:2.5, h:0.28, fontSize:11, color:WHITE }});
  }});
  slide.addShape(pres.shapes.RECTANGLE, {{ x:6.3, y:3.5, w:3.4, h:1.75, fill:{{ color:'7C2D12' }}, line:{{ color: ORANGE, pt:1.5 }}, shadow: makeShadow() }});
  slide.addText('KEY INSIGHT', {{ x:6.4, y:3.55, w:3.2, h:0.3, fontSize:9, bold:true, color: ORANGE, charSpacing:2 }});
  slide.addText(N.au_cover_insight, {{ x:6.4, y:3.85, w:3.2, h:1.3, fontSize:12, color:WHITE, wrap:true }});
  slide.addShape(pres.shapes.LINE, {{ x:0.5, y:5.4, w:9, h:0, line:{{ color:'1E293B', pt:1 }} }});
  slide.addText('CONFIDENTIAL — INTERNAL USE ONLY', {{ x:0.5, y:5.43, w:9, h:0.2, fontSize:8, color:'334155', align:'center' }});
}})();

(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: LIGHT_GREY }};
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: NAVY }}, line:{{ color: NAVY }} }});
  slide.addText('REFUNDS BY BRANCH', {{ x:0.4, y:0.15, w:9, h:0.4, fontSize:20, bold:true, color: NAVY }});
  slide.addText(`Australia — ${{DATE_LABEL}}`, {{ x:0.4, y:0.52, w:9, h:0.3, fontSize:11, color: MID_GREY, italic:true }});
  const cW = 3.0, cH = 2.0;
  const startX = 0.4, startY = 0.95, gapX = 0.15, gapY = 0.15;
  AU_BRANCHES.forEach((b, i) => {{
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = startX + col * (cW + gapX);
    const y = startY + row * (cH + gapY);
    const ins = N.au_branch_cards[b.code] || '';
    slide.addShape(pres.shapes.RECTANGLE, {{ x, y, w:cW, h:cH, fill:{{ color: WHITE }}, line:{{ color:'E2E8F0', pt:1 }}, shadow: makeShadow() }});
    slide.addShape(pres.shapes.RECTANGLE, {{ x, y, w:0.07, h:cH, fill:{{ color: b.color }}, line:{{ color: b.color }} }});
    slide.addText(b.code, {{ x: x+0.15, y: y+0.1, w:1.2, h:0.38, fontSize:22, bold:true, color: b.color, margin:0 }});
    slide.addText(b.name, {{ x: x+0.15, y: y+0.45, w:2.7, h:0.22, fontSize:10, color: MID_GREY, italic:true, margin:0 }});
    slide.addText(fmtAUD(b.total), {{ x: x+0.15, y: y+0.72, w:2.7, h:0.42, fontSize:22, bold:true, color: DARK_GREY, margin:0 }});
    slide.addText(`${{b.count}} claim${{b.count !== 1 ? 's' : ''}}`, {{ x: x+0.15, y: y+1.1, w:1.5, h:0.24, fontSize:11, color: MID_GREY, margin:0 }});
    slide.addShape(pres.shapes.RECTANGLE, {{ x: x+1.75, y: y+1.08, w:1.1, h:0.28, fill:{{ color: b.color }}, line:{{ color: b.color }} }});
    slide.addText(fmtPct(b.pct_of_total), {{ x: x+1.75, y: y+1.08, w:1.1, h:0.28, fontSize:10, bold:true, color: WHITE, align:'center', margin:0 }});
    slide.addText(ins, {{ x: x+0.12, y: y+1.44, w:2.75, h:0.46, fontSize:9, color: MID_GREY, italic:true, wrap:true, margin:0 }});
  }});
}})();

(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: WHITE }};
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: NAVY }}, line:{{ color: NAVY }} }});
  slide.addText('WHERE THE MONEY WENT', {{ x:0.4, y:0.15, w:9, h:0.4, fontSize:20, bold:true, color: NAVY }});
  slide.addText(`Refund category breakdown by branch — ${{DATE_LABEL}}`, {{ x:0.4, y:0.52, w:9, h:0.28, fontSize:11, color: MID_GREY, italic:true }});
  const catColors = ['E63946', 'F4A261', '2A9D8F', '457B9D'];
  const categories = Object.keys(AU_CAT_TOTALS);
  const branchLabels = AU_BRANCHES.map(b => b.code);
  const chartData = categories.map((cat, ci) => ({{ name: cat, labels: branchLabels, values: AU_BRANCHES.map(b => b.by_category[cat] || 0) }}));
  slide.addChart(pres.charts.BAR, chartData, {{ x:0.4, y:0.9, w:6.2, h:3.5, barDir:'col', barGrouping:'stacked', chartColors: catColors, showLegend:true, legendPos:'b', legendFontSize:9, chartArea:{{ fill:{{ color: WHITE }} }}, catAxisLabelColor: MID_GREY, valAxisLabelColor: MID_GREY, valGridLine:{{ color:'E2E8F0', size:0.5 }}, catGridLine:{{ style:'none' }} }});
  const callouts = [
    {{ label:'Total AU Refunds', value: fmtAUD(AU_TOTAL) }},
    {{ label:'Total Claims',     value: AU_COUNT.toString() }},
    {{ label:'Avg Claim Value',  value: fmtAUD(AU_AVG) }},
    {{ label:'Top Category',     value: (function() {{ let top='', topV=0; for(const [k,v] of Object.entries(AU_CAT_TOTALS)) {{ if(v>topV){{top=k;topV=v;}} }} return top; }})() }},
  ];
  callouts.forEach((c, i) => {{
    const y = 0.9 + i * 0.88;
    slide.addShape(pres.shapes.RECTANGLE, {{ x:6.9, y, w:2.8, h:0.8, fill:{{ color: NAVY }}, line:{{ color:'334155' }}, shadow: makeShadow() }});
    slide.addText(c.value, {{ x:6.95, y: y+0.04, w:2.7, h:0.44, fontSize:18, bold:true, color:WHITE, align:'center', margin:0 }});
    slide.addText(c.label, {{ x:6.95, y: y+0.5, w:2.7, h:0.26, fontSize:9, color:'94A3B8', align:'center', margin:0 }});
  }});
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:4.55, w:9.2, h:0.75, fill:{{ color:'FFF7ED' }}, line:{{ color: ORANGE, pt:1 }} }});
  slide.addText(N.au_where_money_went, {{ x:0.55, y:4.6, w:9.0, h:0.65, fontSize:11, color:DARK_GREY, italic:true, wrap:true }});
}})();

(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: LIGHT_GREY }};
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: NAVY }}, line:{{ color: NAVY }} }});
  slide.addText('EQUIPMENT GAP', {{ x:0.4, y:0.15, w:9, h:0.4, fontSize:20, bold:true, color: NAVY }});
  slide.addText('Largest equipment & gear refund per branch', {{ x:0.4, y:0.52, w:9, h:0.28, fontSize:11, color: MID_GREY, italic:true }});
  const gaps = AU_EQUIP_GAPS.slice(0, 6);
  while (gaps.length < 6) gaps.push(null);
  const cW = 2.9, cH = 1.85;
  gaps.forEach((g, i) => {{
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.35 + col * (cW + 0.16);
    const y = 0.93 + row * (cH + 0.2);
    slide.addShape(pres.shapes.RECTANGLE, {{ x, y, w:cW, h:cH, fill:{{ color: WHITE }}, line:{{ color:'E2E8F0' }}, shadow: makeShadow() }});
    if (!g) {{ slide.addText('No equipment claims', {{ x: x+0.1, y: y+0.7, w:cW-0.2, h:0.5, fontSize:10, color: MID_GREY, italic:true, align:'center' }}); return; }}
    slide.addShape(pres.shapes.RECTANGLE, {{ x, y, w:cW, h:0.08, fill:{{ color: g.color }}, line:{{ color: g.color }} }});
    slide.addText(g.branch_code, {{ x: x+0.12, y: y+0.12, w:0.8, h:0.28, fontSize:16, bold:true, color: g.color, margin:0 }});
    slide.addText(g.category, {{ x: x+0.95, y: y+0.14, w:cW-1.1, h:0.22, fontSize:9, color: MID_GREY, italic:true, align:'right', margin:0 }});
    slide.addText(fmtAUD(g.amount), {{ x: x+0.12, y: y+0.42, w:cW-0.2, h:0.38, fontSize:22, bold:true, color: DARK_GREY, margin:0 }});
    slide.addText(g.last_name || '', {{ x: x+0.12, y: y+0.82, w:cW-0.2, h:0.22, fontSize:10, color: MID_GREY, margin:0 }});
    slide.addText((g.details || '').substring(0, 90), {{ x: x+0.12, y: y+1.05, w:cW-0.22, h:0.72, fontSize:8.5, color: MID_GREY, wrap:true, margin:0 }});
  }});
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0.35, y:4.57, w:9.3, h:0.72, fill:{{ color:'1E293B' }}, line:{{ color: NAVY }} }});
  slide.addText(N.au_equipment_gap, {{ x:0.5, y:4.62, w:9.1, h:0.62, fontSize:11, color:WHITE, italic:true, wrap:true }});
}})();

AU_BRANCHES.forEach(branch => {{
  const slide = pres.addSlide();
  slide.background = {{ color: WHITE }};
  const sideW = 2.6;
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:sideW, h:5.625, fill:{{ color: branch.color }}, line:{{ color: branch.color }} }});
  slide.addText(branch.code, {{ x:0.1, y:0.25, w:sideW-0.15, h:0.7, fontSize:38, bold:true, color:WHITE }});
  slide.addText(branch.name, {{ x:0.1, y:0.9, w:sideW-0.15, h:0.32, fontSize:14, color:WHITE, italic:true }});
  slide.addShape(pres.shapes.LINE, {{ x:0.15, y:1.3, w:sideW-0.3, h:0, line:{{ color:WHITE, pt:1, transparency:50 }} }});
  slide.addText('TOTAL REFUNDS', {{ x:0.1, y:1.45, w:sideW-0.15, h:0.22, fontSize:8, color:WHITE, charSpacing:2 }});
  slide.addText(fmtAUD(branch.total), {{ x:0.1, y:1.65, w:sideW-0.15, h:0.48, fontSize:26, bold:true, color:WHITE }});
  slide.addText(`${{branch.count}} Claim${{branch.count !== 1 ? 's' : ''}}`, {{ x:0.1, y:2.15, w:sideW-0.15, h:0.3, fontSize:13, color:WHITE }});
  slide.addShape(pres.shapes.LINE, {{ x:0.15, y:2.55, w:sideW-0.3, h:0, line:{{ color:WHITE, pt:1, transparency:50 }} }});
  const catLabels = Object.keys(branch.by_category);
  const catValues = catLabels.map(k => branch.by_category[k] || 0);
  const maxCatVal = Math.max(...catValues, 1);
  const barMaxW = sideW - 0.35;
  catLabels.forEach((cat, ci) => {{
    const barY = 2.75 + ci * 0.6;
    const barPct = catValues[ci] / maxCatVal;
    slide.addText(cat.length > 14 ? cat.substring(0,14)+'…' : cat, {{ x:0.12, y: barY, w:sideW-0.2, h:0.2, fontSize:8, color:WHITE }});
    slide.addShape(pres.shapes.RECTANGLE, {{ x:0.12, y: barY+0.22, w:barMaxW, h:0.14, fill:{{ color:WHITE, transparency:70 }}, line:{{ color:WHITE, transparency:70 }} }});
    if (barPct > 0) {{ slide.addShape(pres.shapes.RECTANGLE, {{ x:0.12, y: barY+0.22, w:barMaxW*barPct, h:0.14, fill:{{ color:WHITE }}, line:{{ color:WHITE }} }}); }}
    slide.addText(fmtAUD(catValues[ci]), {{ x:0.12, y: barY+0.38, w:sideW-0.2, h:0.18, fontSize:8.5, bold:true, color:WHITE }});
  }});
  slide.addText(N.au_branch_detail[branch.code] || '', {{ x:0.12, y:5.1, w:sideW-0.22, h:0.42, fontSize:8, color:WHITE, italic:true, wrap:true }});
  const rightX = sideW + 0.25;
  const rightW = 10 - rightX - 0.25;
  slide.addText('INDIVIDUAL CLAIMS', {{ x:rightX, y:0.18, w:rightW, h:0.3, fontSize:13, bold:true, color: NAVY, charSpacing:1 }});
  slide.addShape(pres.shapes.LINE, {{ x:rightX, y:0.52, w:rightW, h:0, line:{{ color:'E2E8F0', pt:1 }} }});
  branch.claims.slice(0, 5).forEach((claim, ci) => {{
    const cy = 0.6 + ci * 1.0;
    slide.addShape(pres.shapes.RECTANGLE, {{ x:rightX, y:cy, w:rightW, h:0.92, fill:{{ color: LIGHT_GREY }}, line:{{ color:'E2E8F0', pt:0.75 }} }});
    slide.addShape(pres.shapes.RECTANGLE, {{ x:rightX, y:cy, w:0.07, h:0.92, fill:{{ color: branch.color }}, line:{{ color: branch.color }} }});
    slide.addShape(pres.shapes.RECTANGLE, {{ x:rightX+0.12, y:cy+0.1, w:1.5, h:0.22, fill:{{ color: branch.color }}, line:{{ color: branch.color }} }});
    slide.addText((claim.category||'').substring(0,18), {{ x:rightX+0.12, y:cy+0.1, w:1.5, h:0.22, fontSize:8, bold:true, color:WHITE, align:'center', margin:0 }});
    slide.addText(fmtAUD(claim.amount), {{ x:rightX+rightW-1.3, y:cy+0.08, w:1.25, h:0.28, fontSize:16, bold:true, color: DARK_GREY, align:'right', margin:0 }});
    slide.addText(`${{claim.last_name||''}}${{claim.rego?' · '+claim.rego:''}}${{claim.res_num?' · Res #'+claim.res_num:''}}`, {{ x:rightX+0.12, y:cy+0.36, w:rightW-0.2, h:0.22, fontSize:9, color: MID_GREY, margin:0 }});
    slide.addText((claim.details||'').substring(0,120), {{ x:rightX+0.12, y:cy+0.58, w:rightW-0.2, h:0.28, fontSize:8.5, color: DARK_GREY, wrap:true, margin:0 }});
  }});
  if (branch.claims.length > 5) {{
    slide.addText(`+ ${{branch.claims.length - 5}} more claims not shown`, {{ x:rightX, y:0.6+5*1.0, w:rightW, h:0.25, fontSize:9, color: MID_GREY, italic:true }});
  }}
}});

(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: NAVY }};
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color:'6D28D9' }}, line:{{ color:'6D28D9' }} }});
  slide.addText('TRAVELLERS AUTOBARN', {{ x:0.5, y:0.2, w:6, h:0.4, fontSize:11, bold:true, color:'94A3B8', charSpacing:3 }});
  slide.addText('NEW ZEALAND', {{ x:0.5, y:0.55, w:9, h:0.9, fontSize:52, bold:true, color:WHITE, charSpacing:2 }});
  slide.addText('Refund Awareness Report', {{ x:0.5, y:1.35, w:9, h:0.45, fontSize:20, color:'94A3B8', italic:true }});
  slide.addText(DATE_LABEL, {{ x:0.5, y:1.75, w:9, h:0.35, fontSize:13, color:'64748B' }});
  [{{ label:'Total Refunds', value: fmtNZD(NZ_TOTAL) }}, {{ label:'Total Claims', value: NZ_COUNT.toString() }}, {{ label:'Average Claim', value: fmtNZD(NZ_AVG) }}].forEach((s, i) => {{
    const x = 0.5 + i * 3.2;
    slide.addShape(pres.shapes.RECTANGLE, {{ x, y:2.35, w:2.9, h:1.05, fill:{{ color:'1E293B' }}, line:{{ color:'334155' }}, shadow: makeShadow() }});
    slide.addText(s.value, {{ x, y:2.38, w:2.9, h:0.55, fontSize:28, bold:true, color:WHITE, align:'center', margin:0 }});
    slide.addText(s.label, {{ x, y:2.88, w:2.9, h:0.35, fontSize:10, color:'64748B', align:'center', margin:0 }});
  }});
  slide.addText('Branches covered:', {{ x:0.5, y:3.62, w:5, h:0.28, fontSize:10, color:'64748B' }});
  NZ_BRANCHES.forEach((b, i) => {{
    const y = 3.9 + i * 0.32;
    slide.addShape(pres.shapes.OVAL, {{ x:0.5, y:y+0.04, w:0.14, h:0.14, fill:{{ color:b.color }}, line:{{ color:b.color }} }});
    slide.addText(`${{b.code}} — ${{b.name}}`, {{ x:0.7, y, w:3, h:0.28, fontSize:11, color:WHITE }});
  }});
  slide.addShape(pres.shapes.RECTANGLE, {{ x:6.3, y:3.5, w:3.4, h:1.75, fill:{{ color:'2E1065' }}, line:{{ color:'6D28D9', pt:1.5 }}, shadow: makeShadow() }});
  slide.addText('KEY INSIGHT', {{ x:6.4, y:3.55, w:3.2, h:0.3, fontSize:9, bold:true, color:'A78BFA', charSpacing:2 }});
  slide.addText(N.nz_cover_insight, {{ x:6.4, y:3.85, w:3.2, h:1.3, fontSize:12, color:WHITE, wrap:true }});
  slide.addShape(pres.shapes.LINE, {{ x:0.5, y:5.4, w:9, h:0, line:{{ color:'1E293B', pt:1 }} }});
  slide.addText('CONFIDENTIAL — INTERNAL USE ONLY', {{ x:0.5, y:5.43, w:9, h:0.2, fontSize:8, color:'334155', align:'center' }});
}})();

(function() {{
  const slide = pres.addSlide();
  slide.background = {{ color: LIGHT_GREY }};
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color: NAVY }}, line:{{ color: NAVY }} }});
  slide.addText('REFUNDS BY BRANCH', {{ x:0.4, y:0.15, w:9, h:0.4, fontSize:20, bold:true, color: NAVY }});
  slide.addText(`New Zealand — ${{DATE_LABEL}}`, {{ x:0.4, y:0.52, w:9, h:0.3, fontSize:11, color: MID_GREY, italic:true }});
  NZ_BRANCHES.forEach((b, i) => {{
    const x = 0.4 + i * 4.65;
    const cW = 4.4, cH = 4.35;
    slide.addShape(pres.shapes.RECTANGLE, {{ x, y:0.95, w:cW, h:cH, fill:{{ color: WHITE }}, line:{{ color:'E2E8F0' }}, shadow: makeShadow() }});
    slide.addShape(pres.shapes.RECTANGLE, {{ x, y:0.95, w:0.08, h:cH, fill:{{ color: b.color }}, line:{{ color: b.color }} }});
    slide.addText(b.code, {{ x:x+0.18, y:1.1, w:1.5, h:0.55, fontSize:30, bold:true, color:b.color, margin:0 }});
    slide.addText(b.name, {{ x:x+0.18, y:1.63, w:cW-0.3, h:0.28, fontSize:13, color:MID_GREY, italic:true }});
    slide.addText(fmtNZD(b.total), {{ x:x+0.18, y:2.0, w:cW-0.3, h:0.55, fontSize:30, bold:true, color:DARK_GREY }});
    slide.addText(`${{b.count}} claim${{b.count!==1?'s':''}}`, {{ x:x+0.18, y:2.57, w:1.8, h:0.3, fontSize:13, color:MID_GREY }});
    slide.addShape(pres.shapes.RECTANGLE, {{ x:x+2.5, y:2.55, w:1.65, h:0.32, fill:{{ color:b.color }}, line:{{ color:b.color }} }});
    slide.addText(fmtPct(b.pct_of_total), {{ x:x+2.5, y:2.55, w:1.65, h:0.32, fontSize:12, bold:true, color:WHITE, align:'center', margin:0 }});
    slide.addText(N.nz_branch_cards[b.code]||'', {{ x:x+0.18, y:4.9, w:cW-0.3, h:0.3, fontSize:9, color:MID_GREY, italic:true, wrap:true }});
  }});
}})();

NZ_BRANCHES.forEach(branch => {{
  const slide = pres.addSlide();
  slide.background = {{ color: WHITE }};
  const sideW = 2.6;
  slide.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:sideW, h:5.625, fill:{{ color:branch.color }}, line:{{ color:branch.color }} }});
  slide.addText(branch.code, {{ x:0.1, y:0.25, w:sideW-0.15, h:0.7, fontSize:38, bold:true, color:WHITE }});
  slide.addText(branch.name, {{ x:0.1, y:0.9, w:sideW-0.15, h:0.32, fontSize:14, color:WHITE, italic:true }});
  slide.addText('TOTAL REFUNDS', {{ x:0.1, y:1.45, w:sideW-0.15, h:0.22, fontSize:8, color:WHITE, charSpacing:2 }});
  slide.addText(fmtNZD(branch.total), {{ x:0.1, y:1.65, w:sideW-0.15, h:0.48, fontSize:26, bold:true, color:WHITE }});
  slide.addText(`${{branch.count}} Claim${{branch.count!==1?'s':''}}`, {{ x:0.1, y:2.15, w:sideW-0.15, h:0.3, fontSize:13, color:WHITE }});
  slide.addText(N.nz_branch_detail[branch.code]||'', {{ x:0.12, y:5.1, w:sideW-0.22, h:0.42, fontSize:8, color:WHITE, italic:true, wrap:true }});
  const rightX = sideW + 0.25;
  const rightW = 10 - rightX - 0.25;
  slide.addText('INDIVIDUAL CLAIMS', {{ x:rightX, y:0.18, w:rightW, h:0.3, fontSize:13, bold:true, color:NAVY, charSpacing:1 }});
  slide.addShape(pres.shapes.LINE, {{ x:rightX, y:0.52, w:rightW, h:0, line:{{ color:'E2E8F0', pt:1 }} }});
  branch.claims.slice(0, 5).forEach((claim, ci) => {{
    const cy = 0.6 + ci * 1.0;
    slide.addShape(pres.shapes.RECTANGLE, {{ x:rightX, y:cy, w:rightW, h:0.92, fill:{{ color:LIGHT_GREY }}, line:{{ color:'E2E8F0', pt:0.75 }} }});
    slide.addShape(pres.shapes.RECTANGLE, {{ x:rightX, y:cy, w:0.07, h:0.92, fill:{{ color:branch.color }}, line:{{ color:branch.color }} }});
    slide.addShape(pres.shapes.RECTANGLE, {{ x:rightX+0.12, y:cy+0.1, w:1.5, h:0.22, fill:{{ color:branch.color }}, line:{{ color:branch.color }} }});
    slide.addText((claim.category||'').substring(0,18), {{ x:rightX+0.12, y:cy+0.1, w:1.5, h:0.22, fontSize:8, bold:true, color:WHITE, align:'center', margin:0 }});
    slide.addText(fmtNZD(claim.amount), {{ x:rightX+rightW-1.3, y:cy+0.08, w:1.25, h:0.28, fontSize:16, bold:true, color:DARK_GREY, align:'right', margin:0 }});
    slide.addText(`${{claim.last_name||''}}${{claim.rego?' · '+claim.rego:''}}`, {{ x:rightX+0.12, y:cy+0.36, w:rightW-0.2, h:0.22, fontSize:9, color:MID_GREY, margin:0 }});
    slide.addText((claim.details||'').substring(0,120), {{ x:rightX+0.12, y:cy+0.58, w:rightW-0.2, h:0.28, fontSize:8.5, color:DARK_GREY, wrap:true, margin:0 }});
  }});
}});

const outputPath = {json.dumps(output_path)};
pres.writeFile({{ fileName: outputPath }})
  .then(() => {{ console.log('SUCCESS:' + outputPath); }})
  .catch(err => {{ console.error('ERROR:', err.message); process.exit(1); }});
"""
    return script

def main():
    parser = argparse.ArgumentParser(description="Generate AU/NZ Refund Awareness PPTX")
    parser.add_argument("--data",        help="Path to JSON data file")
    parser.add_argument("--output",      default="aunz_refund_report.pptx")
    parser.add_argument("--start-date",  help="Filter start date YYYY-MM-DD")
    parser.add_argument("--end-date",    help="Filter end date YYYY-MM-DD")
    parser.add_argument("--date-label",  help="Human-readable date range label")
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

    date_label = args.date_label
    if not date_label:
        if start_date and end_date:
            date_label = f"{start_date.strftime('%d %b')}–{end_date.strftime('%d %b %Y')}"
        else:
            date_label = datetime.now().strftime("%B %Y")

    if args.skip_claude:
        narratives = {
            "au_cover_insight":   "Australia refund data reviewed for this period.",
            "au_branch_cards":    {bc: f"{bc} branch summary." for bc in AU_BRANCHES},
            "au_where_money_went":"Category breakdown shows key refund drivers across branches.",
            "au_equipment_gap":   "Equipment-related refunds flagged for review.",
            "nz_cover_insight":   "New Zealand refund data reviewed for this period.",
            "nz_branch_cards":    {bc: f"{bc} branch summary." for bc in NZ_BRANCHES},
            "au_branch_detail":   {bc: f"Detail view for {AU_BRANCHES[bc]['name']}." for bc in AU_BRANCHES},
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
