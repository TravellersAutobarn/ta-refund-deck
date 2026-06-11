#!/usr/bin/env python3
"""
Travellers Autobarn — Accident Report Deck Generator
====================================================
Photos are fetched by Node.js in small batches to avoid timeouts.
"""

import sys
import json
import argparse
import os
import subprocess
import tempfile
from datetime import datetime, timedelta

CATEGORIES = [
    ("Reversing / Backing", ["revers", "backed", "backing"]),
    ("Low Clearance / Roof Strike", ["clearance", "roof", "solar panel", "parking garage", "parking structure", "beam", "pole"]),
    ("Kangaroo Strike", ["kangaroo"]),
    ("Animal Strike", ["bird", "animal", "struck animal", "wildlife"]),
    ("Collision – Other Vehicle", ["collision", "collided", "contact with vehicle", "struck vehicle", "head-on", "rear-end", "lane change", "struck by"]),
    ("Campsite Damage", ["campsite", "campground", "tree branch", "tree", "camp"]),
    ("Parking Lot Incident", ["parking lot", "parking space", "parking", "parked"]),
    ("Weather / Road Hazard", ["gravel", "construction", "road hazard", "pothole", "weather"]),
    ("Unknown / Other", []),
]


def categorise(cause_summary):
    text = (cause_summary or "").lower()
    for label, keywords in CATEGORIES:
        if label == "Unknown / Other":
            return label
        for kw in keywords:
            if kw in text:
                return label
    return "Unknown / Other"


def build_category_counts(accidents):
    counts = {}
    for a in accidents:
        cat = categorise(a.get("cause_summary", ""))
        counts[cat] = counts.get(cat, 0) + 1
    result = []
    for label, _ in CATEGORIES:
        if counts.get(label, 0) > 0:
            result.append((label, counts[label]))
    return result


def load_data(args):
    if args.data:
        with open(args.data) as f:
            return json.load(f)
    return json.loads(sys.stdin.read().strip())


def make_photo_url(photo_url, jotform_key):
    if not photo_url:
        return None
    url = str(photo_url).strip()
    if not url:
        return None
    if jotform_key and "jotform" in url.lower() and "apiKey=" not in url:
        sep = "&" if "?" in url else "?"
        url = url + sep + "apiKey=" + jotform_key
    return url


def build_pptx_script(accidents, category_counts, region, date_label, output_path):

    max_count = max((c for _, c in category_counts), default=1)
    bars_json = json.dumps([
        {"label": label, "count": count, "pct": round(count / max_count * 100)}
        for label, count in category_counts
    ])

    script = f"""
'use strict';
const pptxgen = require('pptxgenjs');
const https = require('https');
const http = require('http');

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

// Fetch a single image URL, following redirects
function fetchImage(rawUrl, redirectsLeft) {{
  redirectsLeft = (redirectsLeft === undefined) ? 5 : redirectsLeft;
  return new Promise((resolve) => {{
    if (!rawUrl) {{ resolve(null); return; }}
    let parsed;
    try {{ parsed = new URL(rawUrl); }} catch(e) {{ resolve(null); return; }}
    const lib = parsed.protocol === 'https:' ? https : http;
    const req = lib.get(rawUrl, {{
      headers: {{
        'User-Agent': 'Mozilla/5.0 (compatible; TravellersAutobarn/1.0)',
        'Accept': 'image/*,*/*;q=0.8'
      }},
      timeout: 15000
    }}, (res) => {{
      if ([301,302,303,307,308].includes(res.statusCode) && res.headers.location && redirectsLeft > 0) {{
        res.resume();
        resolve(fetchImage(res.headers.location, redirectsLeft - 1));
        return;
      }}
      if (res.statusCode !== 200) {{
        console.error('  Photo HTTP ' + res.statusCode + ': ' + (rawUrl||'').substring(0,60));
        res.resume();
        resolve(null);
        return;
      }}
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {{
        const buf = Buffer.concat(chunks);
        const mime = (res.headers['content-type'] || 'image/jpeg').split(';')[0].trim();
        resolve('data:' + mime + ';base64,' + buf.toString('base64'));
      }});
    }});
    req.on('error', (e) => {{ console.error('  Photo error: ' + e.message); resolve(null); }});
    req.on('timeout', () => {{ console.error('  Photo timeout'); req.destroy(); resolve(null); }});
  }});
}}

// Fetch in small batches of 5 to avoid memory spikes and timeouts
async function fetchInBatches(urls, batchSize) {{
  batchSize = batchSize || 5;
  const results = [];
  for (let i = 0; i < urls.length; i += batchSize) {{
    const batch = urls.slice(i, i + batchSize);
    console.error('  Fetching photos ' + (i+1) + '-' + Math.min(i+batchSize, urls.length) + ' of ' + urls.length);
    const batchResults = await Promise.all(batch.map(u => fetchImage(u)));
    results.push(...batchResults);
  }}
  return results;
}}

async function buildDeck() {{
  const photoUrls = ITEMS.map(a => a.photo_url || null);
  console.error('Fetching photos in batches...');
  const photos = await fetchInBatches(photoUrls, 5);
  const photoCount = photos.filter(Boolean).length;
  console.error('Photos: ' + photoCount + '/' + ITEMS.length + ' loaded');

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = 'Travellers Autobarn \\u2014 Accident Report (' + REGION + ')';

  // ── Cover ─────────────────────────────────────────────────────────────────
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

  // ── Breakdown slide ────────────────────────────────────────────────────────
  (function() {{
    const s = pres.addSlide();
    s.background = {{ color: NAVY }};
    s.addText('TRAVELLERS AUTOBARN', {{ x:0.5, y:0.18, w:7, h:0.25, fontSize:9, bold:true, color:'94A3B8', charSpacing:3 }});
    s.addText('Accident Breakdown', {{ x:0.5, y:0.42, w:7, h:0.55, fontSize:26, bold:true, color:WHITE }});
    s.addText(REGION + DSUF, {{ x:0.5, y:0.94, w:7, h:0.28, fontSize:10, color:'64748B', italic:true }});

    const cardY = 1.38, cardH = 1.05;
    s.addShape(pres.shapes.RECTANGLE, {{ x:7.8, y:cardY, w:1.8, h:cardH, fill:{{color:'1E293B'}}, line:{{color:'334155'}} }});
    s.addText(TOTAL.toString(), {{ x:7.8, y:cardY+0.08, w:1.8, h:0.6, fontSize:34, bold:true, color:ORANGE, align:'center', margin:0, valign:'middle' }});
    s.addText('TOTAL', {{ x:7.8, y:cardY+0.72, w:1.8, h:0.24, fontSize:8, color:'94A3B8', align:'center', charSpacing:2, margin:0 }});

    const topCats = CATEGORY_BARS.slice(0, 4);
    const tailCats = CATEGORY_BARS.slice(4);
    const cardW = 2.15, cardGap = 0.12, cardStartX = 0.4;

    topCats.forEach((item, i) => {{
      const cx = cardStartX + i * (cardW + cardGap);
      const isTop = (item.pct === 100);
      const bgColor = isTop ? ORANGE : '1E293B';
      const lblColor = isTop ? 'FED7AA' : '94A3B8';
      s.addShape(pres.shapes.RECTANGLE, {{ x:cx, y:cardY, w:cardW, h:cardH, fill:{{color:bgColor}}, line:{{color:isTop ? ORANGE : '334155'}} }});
      s.addText(item.count.toString(), {{ x:cx, y:cardY+0.08, w:cardW, h:0.55, fontSize:30, bold:true, color:WHITE, align:'center', margin:0, valign:'middle' }});
      s.addText(item.label, {{ x:cx+0.1, y:cardY+0.65, w:cardW-0.2, h:0.32, fontSize:9, color:lblColor, align:'center', margin:0, wrap:true }});
    }});

    if (tailCats.length > 0) {{
      const tailY = cardY + cardH + 0.22;
      const tailLabelW = 1.9, tailBarMaxW = 4.9, tailRowH = 0.38;
      s.addText('OTHER CATEGORIES', {{ x:0.4, y:tailY-0.28, w:5, h:0.22, fontSize:8, color:'475569', charSpacing:2, bold:true }});
      tailCats.forEach((item, i) => {{
        const ty = tailY + i * tailRowH;
        const barW = Math.max(0.05, tailBarMaxW * item.pct / 100);
        s.addText(item.label, {{ x:0.4, y:ty, w:tailLabelW, h:0.28, fontSize:9, color:'94A3B8', valign:'middle', align:'right', margin:0 }});
        s.addShape(pres.shapes.RECTANGLE, {{ x:0.4+tailLabelW+0.1, y:ty+0.06, w:tailBarMaxW, h:0.16, fill:{{color:'1E293B'}}, line:{{color:'334155'}} }});
        s.addShape(pres.shapes.RECTANGLE, {{ x:0.4+tailLabelW+0.1, y:ty+0.06, w:barW, h:0.16, fill:{{color:ORANGE}}, line:{{color:ORANGE}} }});
        s.addText(item.count.toString(), {{ x:0.4+tailLabelW+0.1+tailBarMaxW+0.08, y:ty, w:0.4, h:0.28, fontSize:10, bold:true, color:WHITE, valign:'middle', margin:0 }});
      }});
    }}
  }})();

  // ── One slide per accident ─────────────────────────────────────────────────
  if (ITEMS.length === 0) {{
    const s = pres.addSlide();
    s.background = {{ color: LIGHT_GREY }};
    s.addText('No accidents reported for this region in the period.', {{ x:0.4, y:2.5, w:9.2, h:0.6, fontSize:14, color:MID_GREY, italic:true, align:'center' }});
  }}

  ITEMS.forEach((a, idx) => {{
    const s = pres.addSlide();
    s.background = {{ color: LIGHT_GREY }};
    s.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{color:NAVY}}, line:{{color:NAVY}} }});
    s.addText('ACCIDENT REPORT', {{ x:0.4, y:0.12, w:5, h:0.35, fontSize:16, bold:true, color:NAVY }});
    s.addText('Incident ' + (idx+1) + ' of ' + ITEMS.length + '  \\u2022  ' + REGION + DSUF, {{
      x:0.4, y:0.44, w:9.2, h:0.25, fontSize:9.5, color:MID_GREY, italic:true }});
    s.addShape(pres.shapes.RECTANGLE, {{ x:7.0, y:0.1, w:2.6, h:0.55, fill:{{color:NAVY}}, line:{{color:'334155'}}, shadow:shadow() }});
    s.addText([
      {{ text:'REGO  ', options:{{ fontSize:8, color:'94A3B8', bold:true }} }},
      {{ text: clip(a.license_plate || '\\u2014', 12), options:{{ fontSize:16, color:WHITE, bold:true }} }},
    ], {{ x:7.0, y:0.1, w:2.6, h:0.55, align:'center', valign:'middle', margin:0 }});
    s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:0.82, w:9.2, h:1.4, fill:{{color:WHITE}}, line:{{color:BORDER, pt:1}}, shadow:shadow() }});
    s.addText('WHAT HAPPENED', {{ x:0.6, y:0.88, w:8.8, h:0.28, fontSize:9, bold:true, color:ORANGE, charSpacing:1.5 }});
    s.addText(clip(a.cause_summary || 'No cause recorded.', 300), {{
      x:0.6, y:1.16, w:8.7, h:0.85, fontSize:13, color:DARK_GREY, wrap:true, valign:'top', margin:0 }});
    const meta = [a.location, a.date].filter(Boolean).join('   \\u2022   ');
    if (meta) s.addText(meta, {{ x:0.6, y:2.12, w:8.7, h:0.22, fontSize:9.5, color:MID_GREY, italic:true }});

    const photoY = 2.45, photoH = 3.0;
    s.addText('DAMAGE PHOTO', {{ x:0.4, y:photoY-0.28, w:4, h:0.25, fontSize:9, bold:true, color:MID_GREY, charSpacing:1.5 }});
    const photo = photos[idx];
    if (photo) {{
      s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:photoY, w:9.2, h:photoH, fill:{{color:'EEF2F7'}}, line:{{color:BORDER, pt:1}} }});
      s.addImage({{ data:photo, x:0.4, y:photoY, w:9.2, h:photoH, sizing:{{ type:'contain', w:9.2, h:photoH }} }});
    }} else {{
      s.addShape(pres.shapes.RECTANGLE, {{ x:0.4, y:photoY, w:9.2, h:photoH, fill:{{color:'F1F5F9'}}, line:{{color:'CBD5E1', pt:1, dashType:'dash'}} }});
      s.addText('Photo of damage not supplied', {{ x:0.4, y:photoY, w:9.2, h:photoH, fontSize:13, color:'94A3B8', align:'center', valign:'middle', italic:true }});
    }}
  }});

  const outputPath = {json.dumps(output_path)};
  await pres.writeFile({{ fileName: outputPath }});
  console.log('SUCCESS:' + outputPath);
}}

buildDeck().catch(err => {{ console.error('ERROR:', err.message); process.exit(1); }});
"""
    return script


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
    print(f"JOTFORM_API_KEY present: {bool(jotform_key)}", file=sys.stderr)
    print(f"Building deck for {region}: {len(rows)} accidents", file=sys.stderr)

    accidents = []
    for r in rows:
        accidents.append({
            "license_plate": r.get("license_plate") or r.get("rego") or r.get("plate") or "",
            "cause_summary": r.get("cause_summary") or r.get("cause") or "",
            "location": r.get("location") or "",
            "date": r.get("date") or "",
            "photo_url": make_photo_url(r.get("photo_url") or r.get("photo"), jotform_key),
        })

    category_counts = build_category_counts(accidents)
    script = build_pptx_script(accidents, category_counts, region, date_label, args.output)

    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False) as f:
        f.write(script)
        script_path = f.name

    script_dir = os.path.dirname(os.path.abspath(__file__))
    result = subprocess.run(["node", script_path], capture_output=True, text=True, cwd=script_dir)
    os.unlink(script_path)

    if result.stderr:
        print(result.stderr, file=sys.stderr)

    if result.returncode != 0:
        print("Node.js error:", result.stderr, file=sys.stderr)
        sys.exit(1)

    print(f"Done: {args.output}", file=sys.stderr)
    print(args.output)


def make_photo_url(photo_url, jotform_key):
    if not photo_url:
        return None
    url = str(photo_url).strip()
    if not url:
        return None
    if jotform_key and "jotform" in url.lower() and "apiKey=" not in url:
        sep = "&" if "?" in url else "?"
        url = url + sep + "apiKey=" + jotform_key
    return url


if __name__ == "__main__":
    main()
