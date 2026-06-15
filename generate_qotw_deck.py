#!/usr/bin/env python3
"""
generate_qotw_deck.py

Builds "Question of the Week" (KB Spotlight) decks for Travellers Autobarn.
Claude (upstream in Make) reads the week's tickets, finds the single most
common customer QUESTION per region, and writes the correct answer/process
plus a one-line talking point staff can use. It doubles as a weekly QA pass
on the Retell KB: the manager compares the answer to the official KB.

Split by region/country:
  - America deck   -> region "america" / country "united_states"
  - AUS/NZ deck    -> region "aunz"    / country "australia" or "new_zealand"

Payload:
    {
      "date_label": "June 2026",
      "questions": [
        {"region": "america", "question": "...", "answer": "...",
         "talking_point": "...", "kb_topic": "...", "times_asked": 7,
         "kb_check": "..."}
      ]
    }

Usage (CLI):
    python generate_qotw_deck.py --input qotw.json --outdir ./out

Usage (HTTP, via app.py routes):
    POST /generate-qotw-deck/america
    POST /generate-qotw-deck/aunz
"""

import argparse
import json
import os
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# Teal / green / blue "knowledge" palettes — distinct from the other decks.
PALETTES = [
    {"bg": "0E7C7B", "band": "08504F", "ink": "FFFFFF", "mark": "9FD8D6"},  # teal
    {"bg": "2A6F4E", "band": "184631", "ink": "FFFFFF", "mark": "A7D3BC"},  # green
    {"bg": "1B6CA8", "band": "0E4368", "ink": "FFFFFF", "mark": "AFD4EC"},  # blue
    {"bg": "8A5A2B", "band": "5A3A1A", "ink": "FFFFFF", "mark": "E0C3A0"},  # amber
]

REGION_TITLES = {
    "america": "Question of the Week — USA",
    "aunz": "Question of the Week — Australia & New Zealand",
}


def _rgb(hex_str):
    return RGBColor.from_string(hex_str)


def _set_fill(shape, hex_str):
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(hex_str)
    shape.line.fill.background()


def _normalise(value):
    if not value:
        return ""
    return value.strip().lower().replace("_", " ").replace("-", " ")


def _region_for_country(country):
    c = _normalise(country)
    if c in ("united states", "usa", "us", "united states of america", "america"):
        return "america"
    if c in ("australia", "new zealand", "aus", "nz", "aotearoa"):
        return "aunz"
    return None


def _region_of(item):
    r = _normalise(item.get("region"))
    if r in ("america", "usa", "us"):
        return "america"
    if r in ("aunz", "au nz", "aus nz", "anz"):
        return "aunz"
    return _region_for_country(item.get("country"))


def _coerce_list(value):
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except (ValueError, TypeError):
            return []
    if isinstance(value, dict):
        value = [value]
    if not isinstance(value, (list, tuple)):
        return []
    out = []
    for item in value:
        if isinstance(item, str):
            try:
                item = json.loads(item)
            except (ValueError, TypeError):
                continue
        if isinstance(item, dict):
            out.append(item)
    return out


def _coerce_steps(answer):
    """Answer may be a paragraph, a list of steps, or a newline/semicolon string."""
    if not answer:
        return []
    if isinstance(answer, (list, tuple)):
        return [str(a).strip(" -•\t") for a in answer if str(a).strip(" -•\t")]
    if isinstance(answer, str):
        # If it reads like multiple steps, split; otherwise return as one block
        if "\n" in answer or ";" in answer:
            raw = answer.replace(";", "\n").splitlines()
            return [a.strip(" -•\t") for a in raw if a.strip(" -•\t")]
        return [answer.strip()]
    return [str(answer)]


def _add_lightbulb(slide, color):
    """Simple lightbulb motif from basic shapes, top-left."""
    bulb = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.55), Inches(0.35), Inches(1.3), Inches(1.3))
    _set_fill(bulb, color)
    bulb.shadow.inherit = False
    base = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(1.55), Inches(0.5), Inches(0.35)
    )
    _set_fill(base, color)
    base.shadow.inherit = False


def _add_slide(prs, item, palette, region):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    _set_fill(bg, palette["bg"])
    bg.shadow.inherit = False

    _add_lightbulb(slide, palette["mark"])

    # Header banner, top-right
    band = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0), Inches(0.55), Inches(7.9), Inches(0.95)
    )
    _set_fill(band, palette["band"])
    band.shadow.inherit = False
    btf = band.text_frame
    btf.word_wrap = True
    btf.vertical_anchor = MSO_ANCHOR.MIDDLE
    btf.margin_left = Inches(0.2)
    btf.margin_right = Inches(0.2)
    bp = btf.paragraphs[0]
    bp.alignment = PP_ALIGN.CENTER
    br = bp.add_run()
    br.text = REGION_TITLES.get(region, "Question of the Week")
    br.font.size = Pt(19)
    br.font.bold = True
    br.font.name = "Calibri"
    br.font.color.rgb = _rgb(palette["ink"])

    # White card: question, answer, talking point
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(1.95), Inches(11.13), Inches(4.5)
    )
    _set_fill(card, "FFFFFF")
    card.shadow.inherit = False
    ctf = card.text_frame
    ctf.word_wrap = True
    ctf.vertical_anchor = MSO_ANCHOR.TOP
    ctf.margin_left = Inches(0.5)
    ctf.margin_right = Inches(0.5)
    ctf.margin_top = Inches(0.4)
    ctf.margin_bottom = Inches(0.3)

    question = str(item.get("question", "") or "").strip()
    steps = _coerce_steps(item.get("answer"))
    talking = str(item.get("talking_point", "") or "").strip()

    # The question (prominent)
    qp = ctf.paragraphs[0]
    qp.alignment = PP_ALIGN.LEFT
    qr = qp.add_run()
    qr.text = "Q:  " + question
    qr.font.size = Pt(19)
    qr.font.bold = True
    qr.font.name = "Calibri"
    qr.font.color.rgb = _rgb(palette["bg"])
    qp.space_after = Pt(12)

    # The answer label
    ap = ctf.add_paragraph()
    ap.alignment = PP_ALIGN.LEFT
    ar = ap.add_run()
    ar.text = "The answer / process"
    ar.font.size = Pt(13)
    ar.font.bold = True
    ar.font.name = "Calibri"
    ar.font.color.rgb = _rgb("6B6B7B")
    ap.space_after = Pt(4)

    # The answer body (paragraph or steps)
    multi = len(steps) > 1
    for s in steps:
        bp2 = ctf.add_paragraph()
        bp2.alignment = PP_ALIGN.LEFT
        br2 = bp2.add_run()
        br2.text = ("•  " + s) if multi else s
        br2.font.size = Pt(15)
        br2.font.name = "Calibri"
        br2.font.color.rgb = _rgb("2B2D42")
        bp2.space_after = Pt(4)

    # Talking point
    if talking:
        tp = ctf.add_paragraph()
        tp.alignment = PP_ALIGN.LEFT
        tp.space_before = Pt(10)
        tl = tp.add_run()
        tl.text = "Say it simply:  "
        tl.font.size = Pt(14)
        tl.font.bold = True
        tl.font.name = "Calibri"
        tl.font.color.rgb = _rgb(palette["bg"])
        tv = tp.add_run()
        tv.text = talking
        tv.font.size = Pt(14)
        tv.font.italic = True
        tv.font.name = "Calibri"
        tv.font.color.rgb = _rgb("3D3D4E")

    # Footer: KB topic and how often it came up
    kb_topic = str(item.get("kb_topic", "") or "").strip()
    times = item.get("times_asked")
    foot_bits = []
    if kb_topic:
        foot_bits.append(("KB topic", kb_topic))
    if times not in (None, "", 0, "0"):
        foot_bits.append(("Asked this week", str(times) + " times"))

    if foot_bits:
        foot = slide.shapes.add_textbox(Inches(1.1), Inches(6.6), Inches(11.13), Inches(0.7))
        ftf = foot.text_frame
        ftf.word_wrap = True
        fp = ftf.paragraphs[0]
        fp.alignment = PP_ALIGN.CENTER
        for i, (label, value) in enumerate(foot_bits):
            rl = fp.add_run()
            rl.text = f"{label}: "
            rl.font.size = Pt(14)
            rl.font.bold = True
            rl.font.name = "Calibri"
            rl.font.color.rgb = _rgb(palette["ink"])
            rv = fp.add_run()
            rv.text = value
            rv.font.size = Pt(14)
            rv.font.name = "Calibri"
            rv.font.color.rgb = _rgb(palette["ink"])
            if i < len(foot_bits) - 1:
                rs = fp.add_run()
                rs.text = "      •      "
                rs.font.size = Pt(14)
                rs.font.bold = True
                rs.font.name = "Calibri"
                rs.font.color.rgb = _rgb(palette["ink"])

    # Manager QA reminder + any KB-check note -> speaker notes
    kb_check = str(item.get("kb_check", "") or "").strip()
    notes_parts = ["KB QA CHECK (not shown on slide)",
                   "\nVerify this answer matches the official KB / Retell article before Friday.",
                   "If they differ, update whichever is wrong — that's the weekly QA win."]
    if kb_check:
        notes_parts.append("\nClaude's note:\n" + kb_check)
    slide.notes_slide.notes_text_frame.text = "\n".join(notes_parts)

    return slide


def _build_deck(items, region, date_label):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    cover = prs.slides.add_slide(prs.slide_layouts[6])
    cbg = cover.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    _set_fill(cbg, "08504F")
    cbg.shadow.inherit = False

    title_box = cover.shapes.add_textbox(Inches(1.0), Inches(2.6), Inches(11.33), Inches(2.0))
    ttf = title_box.text_frame
    ttf.word_wrap = True
    tp = ttf.paragraphs[0]
    tp.alignment = PP_ALIGN.CENTER
    tr = tp.add_run()
    tr.text = REGION_TITLES.get(region, "Question of the Week")
    tr.font.size = Pt(42)
    tr.font.bold = True
    tr.font.name = "Calibri"
    tr.font.color.rgb = _rgb("FFFFFF")

    if date_label:
        sub = cover.shapes.add_textbox(Inches(1.0), Inches(4.4), Inches(11.33), Inches(0.8))
        stf = sub.text_frame
        sp = stf.paragraphs[0]
        sp.alignment = PP_ALIGN.CENTER
        sr = sp.add_run()
        sr.text = date_label
        sr.font.size = Pt(24)
        sr.font.name = "Calibri"
        sr.font.color.rgb = _rgb("9FD8D6")

    for idx, item in enumerate(items):
        palette = PALETTES[idx % len(PALETTES)]
        _add_slide(prs, item, palette, region)

    return prs


def generate(payload, outdir):
    return [p for region in ("america", "aunz")
            for p in [generate_region(payload, outdir, region)] if p]


def generate_region(payload, outdir, region):
    os.makedirs(outdir, exist_ok=True)

    if isinstance(payload, str):
        try:
            payload = json.loads(payload)
        except (ValueError, TypeError):
            payload = {}

    date_label = payload.get("date_label") or datetime.now().strftime("%B %Y")
    raw = payload.get("questions")
    if raw is None:
        raw = payload.get("items", [])
    questions = _coerce_list(raw or [])

    items = [q for q in questions if _region_of(q) == region]
    if not items:
        return None

    prs = _build_deck(items, region, date_label)
    filenames = {"america": "qotw_america.pptx", "aunz": "qotw_aunz.pptx"}
    path = os.path.join(outdir, filenames[region])
    prs.save(path)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--outdir", default="./out")
    args = ap.parse_args()
    with open(args.input) as f:
        payload = json.load(f)
    written = generate(payload, args.outdir)
    if not written:
        print("No questions matched a known region. Nothing generated.")
    for p in written:
        print(f"Wrote {p}")


if __name__ == "__main__":
    main()
