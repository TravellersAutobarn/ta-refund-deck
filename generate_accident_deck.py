#!/usr/bin/env python3
"""
Travellers Autobarn — Accident Report Deck Generator
====================================================
Builds one accident-report PowerPoint per region (America, or Australia & NZ).

Input from Make (after Claude has summarised cause + decided country):
{
  "region": "America",          # or "Australia & New Zealand"
  "date_label": "10 May - 10 Jun 2026",
  "rows": [
    {
      "license_plate": "ABC123",
      "cause_summary": "Reversed into a low bollard in a car park.",
      "location": "Las Vegas, NV",
      "date": "2026-05-22",
      "photo_url": "https://www.jotform.com/uploads/.../damage.jpg"
    }
  ]
}

Each accident gets a card: photo of the damage, registration, cause, location/date.
Missing photos render as a labelled placeholder box.
A summary slide shows a bar chart of accident categories.
"""

import sys
import json
import argparse
import os
import subprocess
import tempfile
import base64
import mimetypes
from datetime import datetime, timedelta
from urllib.request import Request, urlopen

# ─── Colours ──────────────────────────────────────────────────────────────────
NAVY = "0F172A"
ORANGE = "F97316"
TA_BRAND = "E97132"


# ─── Accident categorisation ──────────────────────────────────────────────────
CATEGORIES = [
    ("Reversing / Backing", ["revers", "backed", "backing"]),
    ("Low Clearance / Roof Strike", ["clearance", "roof", "solar panel", "parking garage", "parking structure", "beam", "pole"]),
    ("Kangaroo Strike", ["kangaroo"]),
    ("Animal Strike", ["bird", "animal", "struck animal", "wildlife"]),
    ("Collision – Other Vehicle", ["collision", "collided", "contact with vehicle", "struck vehicle", "head-on", "rear-end", "lane change", "struck by"]),
    ("Campsite Damage", ["campsite", "campground", "tree branch", "tree", "camp"]),
    ("Parking Lot Incident", ["parking lot", "parking space", "parking", "parked"]),
    ("Weather / Road Hazard", ["gravel", "construction", "road hazard", "pothole", "weather"]),
    ("Unknown / Other", []),  # catch-all — always last
]


def categorise(cause_summary):
    """Return a category label for a cause summary string."""
    text = (cause_summary or "").lower()
    for label, keywords in CATEGORIES:
        if label == "Unknown / Other":
            return label
        for kw in keywords:
            if kw in text:
                return label
    return "Unknown / Other"


def build_category_counts(accidents):
    """Return ordered list of (label, count) for categories with count > 0."""
    counts = {}
    for a in accidents:
        cat = categorise(a.get("cause_summary", ""))
        counts[cat] = counts.get(cat, 0) + 1
    # Return in CATEGORIES order, dropping zero-count ones
    result = []
    for label, _ in CATEGORIES:
        if counts.get(label, 0) > 0:
            result.append((label, counts[label]))
    return result


# ─── Data loading ──────────────────────────────────────────────────────────────
def load_data(args):
    if args.data:
        with open(args.data) as f:
            return json.load(f)
    return json.loads(sys.stdin.read().strip())


# ─── Photo handling ──────────────────────────────────────────────────────────────
def resolve_photo(photo, jotform_key=None):
    """Turn a photo reference into a base64 data URI for PptxGenJS.
    Accepts an http(s) URL, a local file path, or an existing data URI.
    Returns None on anything missing or unfetchable (caller draws a placeholder)."""
    if not photo:
        return None
    photo = str(photo).strip()
    if not photo:
        return None
    try:
        if photo.startswith("data:"):
            return photo
        if photo.startswith("http://") or photo.startswith("https://"):
            url = photo
            # JotForm file URLs need the API key to serve the image
            if jotform_key and "jotform" in url.lower() and "apiKey=" not in url:
                sep = "&" if "?" in url else "?"
                url = url + sep + "apiKey=" + jotform_key
            req = Request(url, headers={
                "User-Agent": "Mozilla/5.0 (compatible; TravellersAutobarn/1.0)",
                "Accept": "image/*,*/*;q=0.8",
            })
            with urlopen(req, timeout=20) as resp:
                raw = resp.read()
                ctype = resp.headers.get("Content-Type", "") or ""
            mime = ctype.split(";")[0].strip() if ctype.startswith("image/") else "image/jpeg"
        else:
            # local file path
            with open(photo, "rb") as f:
                raw = f.read()
            mime = mimetypes.guess_type(photo)[0] or "image/jpeg"
        b64 = base64.b64encode(raw).decode("ascii")
        return f"data:{mime};base64,{b64}"
    except Exception as e:
        print(f"  (photo unavailable: {e})", file=sys.stderr)
        return None


# ─── PptxGenJS script builder ──────────────────────────────────────────────────
def build_pptx_script(accidents, category_counts, region, date_label, output_path):

    # Build bar chart data for the summary slide
    # Each bar: label (possibly wrapped) + count + bar width as % of max
    max_count = max((c for _, c in category_counts), default=1)
    bars_json = json.dumps([
        {"label": label, "count": count, "pct": round(count / max_count * 100)}
        for label, count in category_counts
    ])

    script = f"""
'use strict';
const pptxgen = require('pptxgenjs');

const REGION        = {json.dumps(region)};
const DATE_LABEL    = {json.dumps(date_label)};
const TOTAL         = {len(accidents)};
const ITEMS         = {json.dumps(accidents)};
const CATEGORY_BARS = {bars_json};

const NAVY='0F172A', WHITE='FFFFFF', ORANGE='F97316', LIGHT_GREY='F8FAFC',
      MID_GREY='64748B', DARK_GREY='1E293B', BORDER='E2E8F0';
const shadow = () => ({{ type:'outer', blur:8, offset:2, angle:135, color:'000000', opacity:0.12 }});
const clip = (s, n) => (s && s.length > n) ? s.substring(0, n-1) + '\\u2026' : (s || '');
const DSUF = DATE_LABEL ? ' \\u2014 ' + DATE_LABEL : '';

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'Travellers Autobarn \\u2014 Accident Report (' + REGION + ')';

// ── Cover ───────────────────────────────────────────────────────────────────
(function() {{
  const s = pres.addSlide();
  s.background = {{ color: NAVY }};
  s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:ORANGE}}, line:{{color:ORANGE}} }});
  s.addText('TRAVELLERS AUTOBARN', {{ x:0.5, y:0.5, w:9, h:0.4, fontSize:11, bold:true, color:'94A3B8', charSpacing:3 }});
  s.addText('ACCIDENT REPORT', {{ x:0.5, y:0.9, w:9, h:0.9, fontSize:46, bold:true, color:WHITE }});
  s.addText(REGION, {{ x:0.5, y:1.85, w:9, h:0.5, fontSize:20, color:ORANGE, bold:true }});
  if (DATE_LABEL) s.addText('Accidents reported: ' + DATE_LABEL, {{ x:0.5, y:2.35, w:9, h:0.4, fontSize:14, color:'94A3B8', italic:true }});

  s.addShape(pres.shapes.RECTANGLE, {{ x:0.5, y:3.1, w:3.0, h:1.3, fill:{{color:'1E293B'}}, line:{{color:'334155'}}, shadow:shadow() }});
  s.addText(TOTAL.toString(), {{ x:0.5, y:3.25, w:3.0, h:0.7, fontSize:40, bold:true, color:WHITE, align:'center', margin:0, valign:'middle' }});
  s.addText('ACCIDENTS REPORTED', {{ x:0.5, y:3.95, w:3.0, h:0.3, fontSize:10, color:'94A3B8', align:'center', margin:0 }});

  s.addText('CONFIDENTIAL \\u2014 INTERNAL USE ONLY', {{ x:0.5, y:5.35, w:9, h:0.2, fontSize:8, color:'334155', align:'center' }});
}})();

// ── Summary / breakdown slide ────────────────────────────────────────────────
(function() {{
  const s = pres.addSlide();
  s.background = {{ color: LIGHT_GREY }};
  s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:NAVY}}, line:{{color:NAVY}} }});
  s.addText('ACCIDENT BREAKDOWN', {{ x:0.4, y:0.15, w:9.2, h:0.4, fontSize:18, bold:true, color:NAVY }});
  s.addText(REGION + DSUF, {{ x:0.4, y:0.52, w:9.2, h:0.25, fontSize:10, color:MID_GREY, italic:true }});

  // Total callout box
  s.addShape(pres.shapes.RECTANGLE, {{ x:7.4, y:0.82, w:2.2, h:1.1, fill:{{color:NAVY}}, line:{{color:NAVY}}, shadow:shadow() }});
  s.addText(TOTAL.toString(), {{ x:7.4, y:0.88, w:2.2, h:0.6, fontSize:36, bold:true, color:WHITE, align:'center', margin:0, valign:'middle' }});
  s.addText('TOTAL ACCIDENTS', {{ x:7.4, y:1.5, w:2.2, h:0.28, fontSize:8, color:'94A3B8', align:'center', margin:0 }});

  // Bar chart — manual horizontal bars
  const barAreaX = 0.4;
  const barAreaY = 0.82;
  const barAreaW = 6.8;   // total width available for label + bar + count
  const labelW   = 2.5;
  const countW   = 0.5;
  const barMaxW  = barAreaW - labelW - countW - 0.15;
  const barH     = 0.32;
  const rowGap   = 0.10;
  const rowH     = barH + rowGap;

  CATEGORY_BARS.forEach((item, i) => {{
    const y = barAreaY + i * rowH;
    const barW = Math.max(0.05, barMaxW * item.pct / 100);

    // label
    s.addText(item.label, {{
      x: barAreaX, y: y, w: labelW, h: barH,
      fontSize: 9, color: DARK_GREY, valign: 'middle', align: 'right', margin: 0
    }});

    // bar background (track)
    s.addShape(pres.shapes.RECTANGLE, {{
      x: barAreaX + labelW + 0.08, y: y + 0.04, w: barMaxW, h: barH - 0.08,
      fill:{{color:'E2E8F0'}}, line:{{color:'E2E8F0'}}
    }});

    // filled bar — orange for highest, navy for rest
    const barColor = (item.pct === 100) ? ORANGE : '334155';
    s.addShape(pres.shapes.RECTANGLE, {{
      x: barAreaX + labelW + 0.08, y: y + 0.04, w: barW, h: barH - 0.08,
      fill:{{color:barColor}}, line:{{color:barColor}}
    }});

    // count label
    s.addText(item.count.toString(), {{
      x: barAreaX + labelW + 0.08 + barMaxW + 0.06, y: y, w: countW, h: barH,
      fontSize: 10, bold: true, color: DARK_GREY, valign: 'middle', margin: 0
    }});
  }});
}})();

// ── One slide per accident ─────────────────────────────────────────────────
if (ITEMS.length === 0) {{
  const s = pres.addSlide();
  s.background = {{ color: LIGHT_GREY }};
  s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:NAVY}}, line:{{color:NAVY}} }});
  s.addText('ACCIDENT DETAIL', {{ x:0.4, y:0.12, w:6, h:0.4, fontSize:18, bold:true, color:NAVY }});
  s.addText(REGION + DSUF, {{ x:0.4, y:0.49, w:9.2, h:0.28, fontSize:10, color:MID_GREY, italic:true }});
  s.addText('No accidents reported for this region in the period.', {{ x:0.4, y:2.5, w:9.2, h:0.6, fontSize:14, color:MID_GREY, italic:true, align:'center' }});
}}

ITEMS.forEach((a, idx) => {{
  const s = pres.addSlide();
  s.background = {{ color: LIGHT_GREY }};

  // ── header bar ──
  s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:NAVY}}, line:{{color:NAVY}} }});
  s.addText('ACCIDENT REPORT', {{ x:0.4, y:0.12, w:5, h:0.35, fontSize:16, bold:true, color:NAVY }});
  s.addText('Incident ' + (idx+1) + ' of ' + ITEMS.length + '  \\u2022  ' + REGION + DSUF, {{
    x:0.4, y:0.44, w:9.2, h:0.25, fontSize:9.5, color:MID_GREY, italic:true }});

  // ── rego chip ──
  s.addShape(pres.shapes.RECTANGLE, {{ x:7.0, y:0.1, w:2.6, h:0.55, fill:{{color:NAVY}}, line:{{color:'334155'}}, shadow:shadow() }});
  s.addText([
    {{ text:'REGO  ', options:{{ fontSize:8, color:'94A3B8', bold:true }} }},
    {{ text: clip(a.license_plate || '\\u2014', 12), options:{{ fontSize:16, color:WHITE, bold:true }} }},
  ], {{ x:7.0, y:0.1, w:2.6, h:0.55, align:'center', valign:'middle', margin:0 }});

  // ── cause / description box ──
  s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:0.82, w:9.2, h:1.4, fill:{{color:WHITE}}, line:{{color:BORDER, pt:1}}, shadow:shadow() }});
  s.addText('WHAT HAPPENED', {{ x:0.6, y:0.88, w:8.8, h:0.28, fontSize:9, bold:true, color:ORANGE, charSpacing:1.5 }});
  s.addText(clip(a.cause_summary || 'No cause recorded.', 300), {{
    x:0.6, y:1.16, w:8.7, h:0.85, fontSize:13, color:DARK_GREY, wrap:true, valign:'top', margin:0 }});

  // ── location / date ──
  const meta = [a.location, a.date].filter(Boolean).join('   \\u2022   ');
  if (meta) s.addText(meta, {{ x:0.6, y:2.12, w:8.7, h:0.22, fontSize:9.5, color:MID_GREY, italic:true }});

  // ── photo area ──
  const photoY = 2.45, photoH = 3.0;
  s.addText('DAMAGE PHOTO', {{ x:0.4, y:photoY-0.28, w:4, h:0.25, fontSize:9, bold:true, color:MID_GREY, charSpacing:1.5 }});
  if (a.photo) {{
    s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:photoY, w:9.2, h:photoH, fill:{{color:'EEF2F7'}}, line:{{color:BORDER, pt:1}} }});
    s.addImage({{ data:a.photo, x:0.4, y:photoY, w:9.2, h:photoH, sizing:{{ type:'contain', w:9.2, h:photoH }} }});
  }} else {{
    s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:photoY, w:9.2, h:photoH, fill:{{color:'F1F5F9'}}, line:{{color:'CBD5E1', pt:1, dashType:'dash'}} }});
    s.addText('Photo of damage not supplied', {{ x:0.4, y:photoY, w:9.2, h:photoH, fontSize:13, color:'94A3B8', align:'center', valign:'middle', italic:true }});
  }}
}});

const outputPath = {json.dumps(output_path)};
pres.writeFile({{ fileName: outputPath }})
  .then(() => {{ console.log('SUCCESS:' + outputPath); }})
  .catch(err => {{ console.error('ERROR:', err.message); process.exit(1); }});
"""
    return script


# ─── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Generate Accident Report PPTX")
    parser.add_argument("--data", help="Path to JSON data file (else reads stdin)")
    parser.add_argument("--output", default="accident_report.pptx")
    parser.add_argument("--region", help="Region label e.g. 'America'")
    parser.add_argument("--date-label", help="Date range e.g. '10 May - 10 Jun 2026'")
    parser.add_argument("--jotform-key", help="JotForm API key for fetching photos")
    parser.add_argument("--api-key", help="(unused; accepted for route compatibility)")
    args = parser.parse_args()

    data = load_data(args)
    region = args.region
    date_label = args.date_label
    if isinstance(data, dict):
        rows = data.get("rows", [])
        if region is None:
            region = data.get("region")
        if date_label is None:
            date_label = data.get("date_label")
    else:
        rows = data
    rows = [r for r in rows if isinstance(r, dict)]

    region = region or "All Regions"
    if date_label is None:
        end = datetime.now()
        start = end - timedelta(days=31)
        date_label = f"{start.strftime('%d %b')} – {end.strftime('%d %b %Y')}"

    jotform_key = args.jotform_key or os.environ.get("JOTFORM_API_KEY")

    print(f"Building accident deck for {region}: {len(rows)} accidents", file=sys.stderr)

    accidents = []
    for r in rows:
        accidents.append({
            "license_plate": r.get("license_plate") or r.get("rego") or r.get("plate") or "",
            "cause_summary": r.get("cause_summary") or r.get("cause") or "",
            "location": r.get("location") or "",
            "date": r.get("date") or "",
            "photo": resolve_photo(r.get("photo_url") or r.get("photo"), jotform_key),
        })

    category_counts = build_category_counts(accidents)

    script = build_pptx_script(accidents, category_counts, region, date_label, args.output)
    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False) as f:
        f.write(script)
        script_path = f.name

    script_dir = os.path.dirname(os.path.abspath(__file__))
    result = subprocess.run(["node", script_path], capture_output=True, text=True, cwd=script_dir)
    os.unlink(script_path)

    if result.returncode != 0:
        print("Node.js error:", result.stderr, file=sys.stderr)
        sys.exit(1)

    print(f"Done! Output: {args.output}", file=sys.stderr)
    print(args.output)


if __name__ == "__main__":
    main()
