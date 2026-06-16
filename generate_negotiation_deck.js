/**
 * generate_negotiation_deck.js
 * "Difficult Conversations" — a weekly de-escalation curriculum deck.
 * 24-lesson ladder across 4 tiers. State (which lesson) lives in Supabase;
 * Make injects the lesson number, Claude writes the teaching content and
 * picks the week's real difficult ticket as the worked example.
 *
 * Usage:
 *   node generate_negotiation_deck.js --json payload.json --out neg_aunz.pptx
 *   (or pipe JSON via stdin)
 *
 * Expected JSON shape (Claude returns this):
 * {
 *   "date_label": "June 2025",
 *   "region": "aunz",
 *   "tier": 4,
 *   "tier_name": "The Hard Cases",
 *   "lesson_number": 17,
 *   "total_lessons": 24,
 *   "pass_number": 1,
 *   "skill_name": "Don't Take the Bait",
 *   "catch_line": "You don't need to swing at every ball pitched.",
 *   "skill_what": "One short paragraph: what the technique is, in plain words.",
 *   "skill_why": "One short paragraph: the psychology — why it works on an upset person.",
 *   "real_call": {
 *     "what_happened": "2-3 sentence anonymised summary of this week's real difficult ticket.",
 *     "before": [ {"who":"Customer","line":"..."}, {"who":"Staff","line":"..."} ],
 *     "after":  [ {"who":"Customer","line":"..."}, {"who":"Staff","line":"..."} ]
 *   },
 *   "cheat_phrases": ["Exact phrase 1", "Exact phrase 2", "Exact phrase 3"],
 *   "facilitator_guide": "Goes in speaker notes — how to run the 5-min discussion.",
 *   "black_swan": "Goes in speaker notes — the deeper/advanced insight for this skill."
 * }
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");

// ── Palette: deep slate + warm copper. Serious craft, not corporate. ──
const C = {
  dark:    "1E2A32",  // deep slate (covers, dark panels)
  panel:   "27353E",  // lifted slate (cards on dark)
  copper:  "D98E3F",  // warm copper accent
  copperD: "B5722D",  // darker copper
  cream:   "F4F1EC",  // warm off-white (light bg + text on dark)
  card:    "FFFFFF",  // white cards on light
  body:    "2A2A2A",  // near-black body text
  muted:   "8A95A0",  // muted slate-grey
  steel:   "6B7B85",  // steel-blue (the "before" / usual way)
  sage:    "6B9B7A",  // calm green (the "after" / better way)
};

const TIER_COUNT = 4;

function makeShadow() {
  return { type: "outer", color: "000000", blur: 9, offset: 3, angle: 90, opacity: 0.16 };
}

// ── 4-segment tier progress bar (the recurring motif) ──
function progressBar(slide, x, y, w, currentTier, lessonNumber, totalLessons, onDark) {
  const gap = 0.08;
  const segW = (w - gap * (TIER_COUNT - 1)) / TIER_COUNT;
  for (let i = 0; i < TIER_COUNT; i++) {
    const tierNo = i + 1;
    let fill;
    if (tierNo < currentTier) fill = C.copperD;       // completed
    else if (tierNo === currentTier) fill = C.copper;  // current
    else fill = onDark ? C.panel : "E2DDD4";           // upcoming
    slide.addShape("roundRect", {
      x: x + i * (segW + gap), y, w: segW, h: 0.14,
      fill: { color: fill }, rectRadius: 0.05,
      line: { color: fill, width: 0 },
    });
  }
  slide.addText(`Lesson ${lessonNumber} of ${totalLessons}`, {
    x, y: y + 0.2, w, h: 0.22,
    fontSize: 9, color: onDark ? C.muted : C.muted, align: "left", margin: 0, charSpacing: 1,
  });
}

// ── Speaker bubble for the before/after dialogue ──
function bubble(slide, x, y, w, who, line, accentColor) {
  const h = 0.92;
  // label
  slide.addText(who.toUpperCase(), {
    x, y, w, h: 0.2,
    fontSize: 8, bold: true, color: accentColor, charSpacing: 1.5, margin: 0,
  });
  slide.addShape("roundRect", {
    x, y: y + 0.22, w, h: h - 0.22,
    fill: { color: C.card }, rectRadius: 0.08,
    line: { color: "E5E2DC", width: 1 },
    shadow: makeShadow(),
  });
  slide.addText(line, {
    x: x + 0.12, y: y + 0.3, w: w - 0.24, h: h - 0.38,
    fontSize: 10.5, color: C.body, align: "left", valign: "top", margin: 0,
  });
  return y + h + 0.12;
}

// ── Slide 1: Cover ──
function addCoverSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  // Series eyebrow
  s.addText("DIFFICULT CONVERSATIONS", {
    x: 0.5, y: 0.42, w: 7, h: 0.25,
    fontSize: 10, bold: true, color: C.copper, charSpacing: 3.5, margin: 0,
  });

  // Tier label top-right
  s.addText(`TIER ${d.tier} OF ${TIER_COUNT} · ${d.tier_name.toUpperCase()}`, {
    x: 4.5, y: 0.42, w: 5, h: 0.25,
    fontSize: 9, color: C.muted, align: "right", charSpacing: 2, margin: 0,
  });

  // Big skill name
  s.addText(d.skill_name, {
    x: 0.5, y: 1.35, w: 9, h: 1.5,
    fontSize: 52, bold: true, color: C.cream, align: "left", valign: "top", margin: 0,
  });

  // THE CATCH LINE — the sticky hero element
  s.addText(`\u201C${d.catch_line}\u201D`, {
    x: 0.5, y: 3.05, w: 8.6, h: 1.0,
    fontSize: 22, italic: true, color: C.copper, align: "left", valign: "top", margin: 0,
  });

  // Progress bar bottom-left
  progressBar(s, 0.5, 4.7, 5.0, d.tier, d.lesson_number, d.total_lessons, true);

  // Region + date bottom-right
  const regionLabel = d.region === "america" ? "Americas" : "AU / NZ";
  s.addText(`${regionLabel}  ·  ${d.date_label}`, {
    x: 5.5, y: 4.92, w: 4.0, h: 0.25,
    fontSize: 10, color: C.muted, align: "right", margin: 0,
  });
}

// ── Slide 2: The Skill ──
function addSkillSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  // Header
  s.addText("THE SKILL", {
    x: 0.5, y: 0.4, w: 6, h: 0.25,
    fontSize: 10, bold: true, color: C.copperD, charSpacing: 3, margin: 0,
  });
  s.addText(d.skill_name, {
    x: 0.5, y: 0.66, w: 9, h: 0.7,
    fontSize: 30, bold: true, color: C.dark, margin: 0,
  });

  // "What it is" card
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 1.65, w: 4.4, h: 3.2,
    fill: { color: C.card }, rectRadius: 0.1,
    line: { color: "E5E2DC", width: 1 }, shadow: makeShadow(),
  });
  s.addText("WHAT IT IS", {
    x: 0.75, y: 1.9, w: 3.9, h: 0.25,
    fontSize: 11, bold: true, color: C.copper, charSpacing: 1.5, margin: 0,
  });
  s.addText(d.skill_what, {
    x: 0.75, y: 2.25, w: 3.9, h: 2.45,
    fontSize: 13.5, color: C.body, align: "left", valign: "top", margin: 0,
  });

  // "Why it works" card (dark, for contrast)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.1, y: 1.65, w: 4.4, h: 3.2,
    fill: { color: C.dark }, rectRadius: 0.1,
    line: { color: C.dark, width: 0 }, shadow: makeShadow(),
  });
  s.addText("WHY IT WORKS", {
    x: 5.35, y: 1.9, w: 3.9, h: 0.25,
    fontSize: 11, bold: true, color: C.copper, charSpacing: 1.5, margin: 0,
  });
  s.addText(d.skill_why, {
    x: 5.35, y: 2.25, w: 3.9, h: 2.45,
    fontSize: 13.5, color: C.cream, align: "left", valign: "top", margin: 0,
  });

  // Slim progress strip at the very bottom
  progressBar(s, 0.5, 5.15, 4.0, d.tier, d.lesson_number, d.total_lessons, false);
}

// ── Slide 3: This Week's Real Call (before / after) ──
function addCallSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  const rc = d.real_call || {};

  s.addText("THIS WEEK'S REAL CALL", {
    x: 0.5, y: 0.4, w: 9, h: 0.25,
    fontSize: 10, bold: true, color: C.copperD, charSpacing: 3, margin: 0,
  });
  s.addText(rc.what_happened || "", {
    x: 0.5, y: 0.68, w: 9, h: 0.85,
    fontSize: 13, color: C.body, italic: true, align: "left", valign: "top", margin: 0,
  });

  // Column headers
  s.addText("HOW IT USUALLY GOES", {
    x: 0.5, y: 1.7, w: 4.4, h: 0.25,
    fontSize: 10, bold: true, color: C.steel, charSpacing: 1.5, margin: 0,
  });
  s.addText("HOW IT LANDS BETTER", {
    x: 5.1, y: 1.7, w: 4.4, h: 0.25,
    fontSize: 10, bold: true, color: C.sage, charSpacing: 1.5, margin: 0,
  });

  // Before column
  let by = 2.05;
  (rc.before || []).slice(0, 3).forEach(turn => {
    by = bubble(s, 0.5, by, 4.4, turn.who, turn.line, C.steel);
  });

  // After column
  let ay = 2.05;
  (rc.after || []).slice(0, 3).forEach(turn => {
    ay = bubble(s, 5.1, ay, 4.4, turn.who, turn.line, C.sage);
  });
}

// ── Slide 4: Your Desk Cheat-Card ──
function addCheatSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addText("AT THE DESK THIS WEEK", {
    x: 0.5, y: 0.45, w: 9, h: 0.25,
    fontSize: 10, bold: true, color: C.copper, charSpacing: 3, margin: 0,
  });
  s.addText("Say it like this", {
    x: 0.5, y: 0.72, w: 9, h: 0.6,
    fontSize: 28, bold: true, color: C.cream, margin: 0,
  });

  // Phrase cards
  const phrases = (d.cheat_phrases || []).slice(0, 3);
  const cardW = (9.0 - 0.4 * (phrases.length - 1)) / phrases.length;
  phrases.forEach((p, i) => {
    const x = 0.5 + i * (cardW + 0.4);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.6, w: cardW, h: 2.0,
      fill: { color: C.panel }, rectRadius: 0.1,
      line: { color: C.copper, width: 1.2 }, shadow: makeShadow(),
    });
    s.addText(`${i + 1}`, {
      x: x + 0.15, y: 1.72, w: 0.5, h: 0.45,
      fontSize: 22, bold: true, color: C.copper, margin: 0,
    });
    s.addText(`\u201C${p}\u201D`, {
      x: x + 0.18, y: 2.2, w: cardW - 0.36, h: 1.3,
      fontSize: 13, color: C.cream, italic: true, align: "left", valign: "top", margin: 0,
    });
  });

  // Catch line repeated as the takeaway (stickiness through repetition)
  s.addShape("roundRect", {
    x: 0.5, y: 3.95, w: 9.0, h: 0.85,
    fill: { color: C.copper }, rectRadius: 0.08,
    line: { color: C.copper, width: 0 },
  });
  s.addText("REMEMBER", {
    x: 0.7, y: 4.07, w: 9, h: 0.2,
    fontSize: 8, bold: true, color: C.dark, charSpacing: 2, margin: 0,
  });
  s.addText(`\u201C${d.catch_line}\u201D`, {
    x: 0.7, y: 4.28, w: 8.6, h: 0.45,
    fontSize: 17, bold: true, italic: true, color: C.dark, align: "left", valign: "middle", margin: 0,
  });

  // Footer progress
  progressBar(s, 0.5, 5.0, 4.0, d.tier, d.lesson_number, d.total_lessons, true);

  // Speaker notes: facilitator guide + black swan
  const notes =
    `FACILITATOR GUIDE (5 min):\n${d.facilitator_guide || ""}\n\n` +
    `BLACK SWAN — go deeper:\n${d.black_swan || ""}`;
  s.addNotes(notes);
}

async function main() {
  let raw;
  const ji = process.argv.indexOf("--json");
  if (ji !== -1) raw = fs.readFileSync(process.argv[ji + 1], "utf8");
  else raw = fs.readFileSync("/dev/stdin", "utf8");

  const d = JSON.parse(raw);

  const oi = process.argv.indexOf("--out");
  const out = oi !== -1 ? process.argv[oi + 1] : `negotiation_${d.region || "deck"}.pptx`;

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = `Difficult Conversations — Lesson ${d.lesson_number}: ${d.skill_name}`;
  pres.author = "Travellers Autobarn";

  addCoverSlide(pres, d);
  addSkillSlide(pres, d);
  addCallSlide(pres, d);
  addCheatSlide(pres, d);

  await pres.writeFile({ fileName: out });
  process.stdout.write(out);
}

main().catch(e => { process.stderr.write(e.stack + "\n"); process.exit(1); });
