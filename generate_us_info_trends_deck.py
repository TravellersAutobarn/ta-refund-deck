#!/usr/bin/env python3
"""
Travellers Autobarn — US More Information Trends Deck Generator
================================================================

Creates a PowerPoint deck from slide-ready JSON produced by Make/Claude.

Expected input JSON:

{
  "date_label": "Previous Month",
  "slides": [
    {
      "slide_number": 1,
      "title": "Las Vegas Top 3 Trends",
      "bullets": [
        "Trend 1: ...",
        "Trend 2: ...",
        "Trend 3: ..."
      ]
    }
  ]
}

Usage:
    python generate_us_info_trends_deck.py --data data.json --output output.pptx
"""

import argparse
import json
import os
import subprocess
import sys
import tempfile
from datetime import datetime


EXPECTED_SLIDES = [
    "Las Vegas Top 3 Trends",
    "Las Vegas Suggestions for Top 3 Trends",
    "Los Angeles Top 3 Trends",
    "Los Angeles Suggestions for Top 3 Trends",
    "San Francisco Top 3 Trends",
    "San Francisco Suggestions for Top 3 Trends",
    "United States Overall Top 3 Trends",
    "United States Overall Suggestions for Top 3 Trends",
]


def load_data(args):
    if args.data:
        with open(args.data, "r", encoding="utf-8") as f:
            return json.load(f)

    raw = sys.stdin.read().strip()
    if not raw:
        raise ValueError("No JSON input received.")

    return json.loads(raw)


def normalize_slides(data):
    """
    Accepts either:
      {"date_label": "...", "slides": [...]}
    or:
      [{"slide_number": 1, "title": "...", "bullets": [...]}]
    """
    if isinstance(data, list):
        slides = data
    elif isinstance(data, dict):
        slides = data.get("slides", [])
    else:
        slides = []

    if not isinstance(slides, list):
        raise ValueError("Input must contain a slides array.")

    cleaned = []

    for idx, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            continue

        raw_slide_number = slide.get("slide_number", idx)

        try:
            slide_number = int(raw_slide_number)
        except (TypeError, ValueError):
            slide_number = idx

        default_title = EXPECTED_SLIDES[idx - 1] if idx <= len(EXPECTED_SLIDES) else f"Slide {idx}"
        title = str(slide.get("title", default_title)).strip() or default_title

        bullets = slide.get("bullets", [])

        if isinstance(bullets, str):
            bullets = [bullets]

        if not isinstance(bullets, list):
            bullets = []

        bullet_text = []

        for bullet in bullets:
            text = str(bullet).replace("\n", " ").strip()
            if text:
                bullet_text.append(text)

        cleaned.append({
            "slide_number": slide_number,
            "title": title,
            "bullets": bullet_text[:6],
        })

    cleaned.sort(key=lambda item: item["slide_number"])

    return cleaned


def build_pptx_script(slides, date_label, output_path):
    slides_js = json.dumps(slides)
    date_label_js = json.dumps(date_label)
    output_path_js = json.dumps(output_path)

    script = f"""
'use strict';

const pptxgen = require('pptxgenjs');

const SLIDES = {slides_js};
const DATE_LABEL = {date_label_js};
const OUTPUT_PATH = {output_path_js};

const pres = new pptxgen();
pres.defineLayout({ name: 'LAYOUT_CUSTOM_WIDE', width: 10, height: 5.625 });
pres.layout = 'LAYOUT_CUSTOM_WIDE';
pres.author = 'Travellers Autobarn';
pres.company = 'Travellers Autobarn';
pres.subject = 'US More Information Ticket Trends';
pres.title = 'US More Information Ticket Trends';
pres.lang = 'en-US';
pres.theme = {{
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'en-US'
}};

const NAVY = '0F172A';
const WHITE = 'FFFFFF';
const DARK = '1E293B';
const GREY = '64748B';
const LIGHT = 'F8FAFC';
const LINE = 'E2E8F0';
const ORANGE = 'F97316';

const LOCATION_CONFIG = {{
  'las vegas': {{
    label: 'Las Vegas',
    short: 'LAS',
    color: 'F97316',
    accent: 'FFF7ED'
  }},
  'los angeles': {{
    label: 'Los Angeles',
    short: 'LA',
    color: '2563EB',
    accent: 'EFF6FF'
  }},
  'san francisco': {{
    label: 'San Francisco',
    short: 'SF',
    color: '16A34A',
    accent: 'F0FDF4'
  }},
  'united states': {{
    label: 'United States',
    short: 'US',
    color: '7C3AED',
    accent: 'F5F3FF'
  }}
}};

function inferLocation(title) {{
  const t = String(title || '').toLowerCase();

  if (t.includes('las vegas')) return LOCATION_CONFIG['las vegas'];
  if (t.includes('los angeles')) return LOCATION_CONFIG['los angeles'];
  if (t.includes('san francisco')) return LOCATION_CONFIG['san francisco'];
  if (t.includes('united states') || t.includes('overall') || t.includes('country')) return LOCATION_CONFIG['united states'];

  return {{
    label: 'Travellers Autobarn',
    short: 'TA',
    color: NAVY,
    accent: LIGHT
  }};
}}

function isSuggestionSlide(title) {{
  const t = String(title || '').toLowerCase();

  return (
    t.includes('suggestion') ||
    t.includes('action') ||
    t.includes('recommendation')
  );
}}

function cleanBullet(text) {{
  return String(text || '')
    .replace(/\\s+/g, ' ')
    .replace(/^[-•*]\\s*/, '')
    .trim();
}}

function makeShadow() {{
  return {{
    type: 'outer',
    blur: 8,
    offset: 2,
    angle: 135,
    color: '000000',
    opacity: 0.12
  }};
}}

function addFooter(slide, slideNumber) {{
  slide.addShape(pres.ShapeType.line, {{
    x: 0.5,
    y: 5.28,
    w: 9,
    h: 0,
    line: {{ color: LINE, pt: 1 }}
  }});

  slide.addText('Travellers Autobarn — More Information Ticket Trends', {{
    x: 0.5,
    y: 5.35,
    w: 6.9,
    h: 0.18,
    fontSize: 7,
    color: GREY,
    margin: 0
  }});

  slide.addText(String(slideNumber), {{
    x: 8.8,
    y: 5.35,
    w: 0.7,
    h: 0.18,
    fontSize: 7,
    color: GREY,
    align: 'right',
    margin: 0
  }});
}}

function addTopBar(slide, cfg) {{
  slide.addShape(pres.ShapeType.rect, {{
    x: 0,
    y: 0,
    w: 10,
    h: 0.11,
    fill: {{ color: cfg.color }},
    line: {{ color: cfg.color }}
  }});
}}

function addHeader(slide, cfg, slideTitle, slideNumber) {{
  addTopBar(slide, cfg);

  slide.addText('TRAVELLERS AUTOBARN', {{
    x: 0.45,
    y: 0.22,
    w: 4.5,
    h: 0.22,
    fontSize: 8,
    bold: true,
    color: GREY,
    charSpacing: 2,
    margin: 0
  }});

  slide.addText(DATE_LABEL, {{
    x: 7.0,
    y: 0.22,
    w: 2.55,
    h: 0.22,
    fontSize: 8,
    color: GREY,
    align: 'right',
    margin: 0
  }});

  slide.addText(slideTitle, {{
    x: 0.45,
    y: 0.55,
    w: 7.6,
    h: 0.55,
    fontSize: 24,
    bold: true,
    color: NAVY,
    margin: 0,
    fit: 'shrink'
  }});

  slide.addShape(pres.ShapeType.roundRect, {{
    x: 8.35,
    y: 0.54,
    w: 1.2,
    h: 0.45,
    rectRadius: 0.08,
    fill: {{ color: cfg.color }},
    line: {{ color: cfg.color }}
  }});

  slide.addText(cfg.short, {{
    x: 8.35,
    y: 0.63,
    w: 1.2,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: WHITE,
    align: 'center',
    margin: 0
  }});

  addFooter(slide, slideNumber);
}}

function addCoverSlide() {{
  const slide = pres.addSlide();

  slide.background = {{ color: NAVY }};

  slide.addShape(pres.ShapeType.rect, {{
    x: 0,
    y: 0,
    w: 10,
    h: 0.12,
    fill: {{ color: ORANGE }},
    line: {{ color: ORANGE }}
  }});

  slide.addText('TRAVELLERS AUTOBARN', {{
    x: 0.55,
    y: 0.45,
    w: 6.2,
    h: 0.3,
    fontSize: 10,
    bold: true,
    color: '94A3B8',
    charSpacing: 3,
    margin: 0
  }});

  slide.addText('MORE INFORMATION', {{
    x: 0.55,
    y: 1.08,
    w: 9,
    h: 0.75,
    fontSize: 44,
    bold: true,
    color: WHITE,
    margin: 0,
    fit: 'shrink'
  }});

  slide.addText('Ticket Trends Report', {{
    x: 0.55,
    y: 1.85,
    w: 9,
    h: 0.45,
    fontSize: 22,
    color: 'CBD5E1',
    italic: true,
    margin: 0
  }});

  slide.addText(DATE_LABEL, {{
    x: 0.55,
    y: 2.32,
    w: 9,
    h: 0.3,
    fontSize: 13,
    color: '94A3B8',
    margin: 0
  }});

  const branches = [
    LOCATION_CONFIG['las vegas'],
    LOCATION_CONFIG['los angeles'],
    LOCATION_CONFIG['san francisco'],
    LOCATION_CONFIG['united states']
  ];

  branches.forEach((cfg, i) => {{
    const x = 0.65 + i * 2.25;

    slide.addShape(pres.ShapeType.roundRect, {{
      x,
      y: 3.35,
      w: 1.85,
      h: 0.75,
      rectRadius: 0.08,
      fill: {{ color: '1E293B' }},
      line: {{ color: '334155' }},
      shadow: makeShadow()
    }});

    slide.addShape(pres.ShapeType.ellipse, {{
      x: x + 0.15,
      y: 3.57,
      w: 0.24,
      h: 0.24,
      fill: {{ color: cfg.color }},
      line: {{ color: cfg.color }}
    }});

    slide.addText(cfg.label, {{
      x: x + 0.48,
      y: 3.54,
      w: 1.2,
      h: 0.25,
      fontSize: 9,
      color: WHITE,
      bold: true,
      margin: 0,
      fit: 'shrink'
    }});
  }});

  slide.addText('CONFIDENTIAL — INTERNAL USE ONLY', {{
    x: 0.5,
    y: 5.35,
    w: 9,
    h: 0.2,
    fontSize: 8,
    color: '475569',
    align: 'center',
    margin: 0
  }});
}}

function addTrendSlide(slide, cfg, slideData, slideNumber) {{
  slide.background = {{ color: WHITE }};

  addHeader(slide, cfg, slideData.title, slideNumber);

  const bullets = (slideData.bullets || [])
    .map(cleanBullet)
    .filter(Boolean)
    .slice(0, 3);

  const cardY = 1.35;
  const cardH = 1.08;
  const gap = 0.2;

  bullets.forEach((bullet, idx) => {{
    const y = cardY + idx * (cardH + gap);

    slide.addShape(pres.ShapeType.roundRect, {{
      x: 0.65,
      y,
      w: 8.7,
      h: cardH,
      rectRadius: 0.08,
      fill: {{ color: cfg.accent }},
      line: {{ color: LINE, pt: 1 }},
      shadow: makeShadow()
    }});

    slide.addShape(pres.ShapeType.ellipse, {{
      x: 0.88,
      y: y + 0.27,
      w: 0.52,
      h: 0.52,
      fill: {{ color: cfg.color }},
      line: {{ color: cfg.color }}
    }});

    slide.addText(String(idx + 1), {{
      x: 0.88,
      y: y + 0.39,
      w: 0.52,
      h: 0.18,
      fontSize: 12,
      bold: true,
      color: WHITE,
      align: 'center',
      margin: 0
    }});

    slide.addText(bullet, {{
      x: 1.58,
      y: y + 0.16,
      w: 7.45,
      h: 0.75,
      fontSize: 12.5,
      color: DARK,
      bold: idx === 0,
      margin: 0.03,
      fit: 'shrink',
      valign: 'mid'
    }});
  }});

  if (bullets.length === 0) {{
    slide.addText('No supported trends were available for this location.', {{
      x: 0.8,
      y: 2.35,
      w: 8.4,
      h: 0.4,
      fontSize: 16,
      color: GREY,
      italic: true,
      align: 'center'
    }});
  }}
}}

function addSuggestionSlide(slide, cfg, slideData, slideNumber) {{
  slide.background = {{ color: LIGHT }};

  addHeader(slide, cfg, slideData.title, slideNumber);

  const bullets = (slideData.bullets || [])
    .map(cleanBullet)
    .filter(Boolean)
    .slice(0, 4);

  slide.addShape(pres.ShapeType.rect, {{
    x: 0.45,
    y: 1.25,
    w: 2.1,
    h: 3.75,
    fill: {{ color: cfg.color }},
    line: {{ color: cfg.color }},
    shadow: makeShadow()
  }});

  slide.addText('ACTION\\nPLAN', {{
    x: 0.65,
    y: 2.18,
    w: 1.7,
    h: 0.8,
    fontSize: 25,
    bold: true,
    color: WHITE,
    align: 'center',
    margin: 0,
    fit: 'shrink'
  }});

  slide.addText(cfg.label, {{
    x: 0.65,
    y: 3.08,
    w: 1.7,
    h: 0.35,
    fontSize: 10,
    color: WHITE,
    align: 'center',
    italic: true,
    margin: 0
  }});

  const startX = 2.85;
  const startY = 1.25;
  const rowH = 0.86;

  bullets.forEach((bullet, idx) => {{
    const y = startY + idx * (rowH + 0.12);

    slide.addShape(pres.ShapeType.roundRect, {{
      x: startX,
      y,
      w: 6.65,
      h: rowH,
      rectRadius: 0.08,
      fill: {{ color: WHITE }},
      line: {{ color: LINE, pt: 1 }},
      shadow: makeShadow()
    }});

    slide.addText('✓', {{
      x: startX + 0.18,
      y: y + 0.22,
      w: 0.35,
      h: 0.3,
      fontSize: 15,
      bold: true,
      color: cfg.color,
      margin: 0
    }});

    slide.addText(bullet, {{
      x: startX + 0.65,
      y: y + 0.13,
      w: 5.75,
      h: 0.6,
      fontSize: 11,
      color: DARK,
      margin: 0.03,
      fit: 'shrink',
      valign: 'mid'
    }});
  }});

  if (bullets.length === 0) {{
    slide.addText('No supported suggestions were available for this location.', {{
      x: 3.0,
      y: 2.4,
      w: 6.2,
      h: 0.4,
      fontSize: 15,
      color: GREY,
      italic: true,
      align: 'center'
    }});
  }}
}}

addCoverSlide();

SLIDES.forEach((slideData, idx) => {{
  const slide = pres.addSlide();
  const cfg = inferLocation(slideData.title);
  const slideNumber = slideData.slide_number || idx + 1;

  if (isSuggestionSlide(slideData.title)) {{
    addSuggestionSlide(slide, cfg, slideData, slideNumber);
  }} else {{
    addTrendSlide(slide, cfg, slideData, slideNumber);
  }}
}});

pres.writeFile({{ fileName: OUTPUT_PATH }})
  .then(() => {{
    console.log('SUCCESS:' + OUTPUT_PATH);
  }})
  .catch(err => {{
    console.error('ERROR:', err.message);
    process.exit(1);
  }});
"""

    return script


def main():
    parser = argparse.ArgumentParser(description="Generate US More Information Trends PPTX")
    parser.add_argument("--data", help="Path to JSON data file. If omitted, reads stdin.")
    parser.add_argument("--output", default="us_more_information_trends_deck.pptx")
    parser.add_argument("--date-label", default=None)

    args = parser.parse_args()

    print("Loading slide data...", file=sys.stderr)

    data = load_data(args)

    date_label = args.date_label

    if isinstance(data, dict) and data.get("date_label"):
        date_label = str(data.get("date_label"))

    if not date_label:
        date_label = datetime.now().strftime("%B %Y")

    slides = normalize_slides(data)

    if not slides:
        raise ValueError("No slides found. Expected JSON with a slides array.")

    print(f"Loaded {len(slides)} slides", file=sys.stderr)
    print(f"Date label: {date_label}", file=sys.stderr)

    script = build_pptx_script(slides, date_label, args.output)

    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False) as f:
        f.write(script)
        script_path = f.name

    print(f"Generating PPTX at {args.output}...", file=sys.stderr)

    script_dir = os.path.dirname(os.path.abspath(__file__))

    result = subprocess.run(
        ["node", script_path],
        capture_output=True,
        text=True,
        cwd=script_dir,
        env={**os.environ, "NODE_PATH": os.path.join(script_dir, "node_modules")},
    )

    os.unlink(script_path)

    if result.returncode != 0:
        print("Node.js error:", result.stderr, file=sys.stderr)
        print(result.stdout, file=sys.stderr)
        sys.exit(1)

    print(f"Done! Output: {args.output}", file=sys.stderr)
    print(args.output)


if __name__ == "__main__":
    main()
