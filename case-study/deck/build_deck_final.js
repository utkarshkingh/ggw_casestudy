const pptxgen = require("pptxgenjs");
// Every time and money figure comes from economics.py via figures.json.
// Nothing below is typed by hand, so slide 1 and slide 6 cannot disagree.
const F = require("./figures.json");

const START_FILL = "E6F8F7", START_STROKE = "007A82", START_TEXT = "00373A";
const DECISION_FILL = "FFFFFF", DECISION_STROKE = "5B7D80", DECISION_TEXT = "243038";
const EXC_FILL = "FCEFD6", EXC_STROKE = "EF9F27", EXC_TEXT = "6B4A10";
const EXIT_FILL = "EDF7DB", EXIT_STROKE = "6FA01F", EXIT_TEXT = "33480A";
const ARROW = "5B7D80";
const AI_FILL = "E3EDF6", AI_STROKE = "3B6FA0", AI_TEXT = "1B2430";
const DET_FILL = "ECEFF3", DET_STROKE = "6B7787", DET_TEXT = "2A3340";   // rules only, no model call
const HUM_FILL = "F1E6ED", HUM_STROKE = "7A4A66", HUM_TEXT = "4A2C3B";   // human decision point
const INK = "1B2430", MUTED = "5B6472", LINE = "D7DCE2", PAPER = "F7F7F5", WHITE = "FFFFFF";
const FONT_HEAD = "Cambria", FONT_BODY = "Arial";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "GGW Case Study";
pres.title = "FNOL — Today vs. AI-Native, in three parts";

// ---- fixed column contract: every diamond/box on this deck obeys this ----
const COL_W = 6.0;
const LEFT_X = 0.4, LC = LEFT_X + COL_W / 2;      // 3.4
const RIGHT_X = 6.9, RC = RIGHT_X + COL_W / 2;    // 9.9
const DIA_W = 1.9, DIA_HALF = DIA_W / 2;
const BOX_W = 1.6;              // guaranteed to fit: COL_W/2 - DIA_HALF - gap(0.3) = 1.75 > 1.6
const GAP = 0.3;

function fitH(text, w, fontSize, minH) {
  const usableW = w * 0.90;
  const capChars = usableW / (fontSize / 72 * 0.52);
  const lineH = fontSize / 72 * 1.15;
  let lines = 0;
  for (const para of text.split("\n")) lines += Math.max(1, Math.ceil((para.length || 1) / capChars));
  return Math.max(minH || 0, lines * lineH + 0.11);
}

function rectNode(slide, label, x, y, w, fill, stroke, textColor, opts = {}) {
  const fs = opts.fontSize || 9;
  const h = fitH(label, w, fs, opts.minH);
  slide.addText(label, {
    shape: pres.ShapeType.roundRect, rectRadius: 0.04,
    x, y, w, h, fill: { color: fill }, line: { color: stroke, width: 1.25, dashType: opts.dashed ? "dash" : "solid" },
    fontFace: FONT_BODY, fontSize: fs, bold: !!opts.bold, color: textColor,
    align: "center", valign: "middle", isTextBox: true, margin: 3, lineSpacingMultiple: 1.08,
  });
  return h;
}
function diamondNode(slide, label, cx, y, h, opts = {}) {
  slide.addText(label, {
    shape: pres.ShapeType.diamond,
    x: cx - DIA_HALF, y, w: DIA_W, h, fill: { color: DECISION_FILL }, line: { color: DECISION_STROKE, width: 1.25 },
    fontFace: FONT_BODY, fontSize: opts.fontSize || 9.5, bold: true, color: DECISION_TEXT,
    align: "center", valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 1.0,
  });
}
function vArrow(slide, x, y1, y2) {
  slide.addShape("line", { x, y: y1, w: 0, h: y2 - y1, line: { color: ARROW, width: 1, endArrowType: "triangle" } });
}
function diagArrow(slide, x1, y1, x2, y2, dash) {
  slide.addShape("line", {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: ARROW, width: 1, endArrowType: "triangle", dashType: dash ? "dash" : "solid" },
  });
}
function elbowDown(slide, x, y1, y2) { vArrow(slide, x, y1, y2); }
function elbowAcross(slide, x1, x2, y) { slide.addShape("line", { x: Math.min(x1, x2), y, w: Math.abs(x2 - x1), h: 0, flipH: x2 < x1, line: { color: ARROW, width: 1, endArrowType: "triangle" } }); }
function edgeLabel(slide, label, x, y, w, opts = {}) {
  slide.addText(label, { x, y, w, h: opts.h || 0.2, fontFace: FONT_BODY, fontSize: opts.fontSize || 8, italic: !!opts.italic, color: opts.color || MUTED, align: opts.align || "center", isTextBox: true, margin: 0 });
}
function legendChip(slide, x, y, fill, stroke, label, dashed) {
  slide.addShape(pres.ShapeType.roundRect, { x, y, w: 0.15, h: 0.15, rectRadius: 0.02, fill: { color: fill }, line: { color: stroke, width: 1, dashType: dashed ? "dash" : "solid" } });
  slide.addText(label, { x: x + 0.22, y: y - 0.05, w: 1.45, h: 0.25, fontFace: FONT_BODY, fontSize: 8, color: MUTED, isTextBox: true, margin: 0, valign: "middle" });
}
// Left group describes the original diagram's vocabulary; right group describes
// the redesign's node types. Grouping them per column is the point.
function legendRow(slide, y) {
  legendChip(slide, 0.4, y, START_FILL, START_STROKE, "Trigger / start");
  legendChip(slide, 1.85, y, DECISION_FILL, DECISION_STROKE, "Step / decision");
  legendChip(slide, 3.35, y, EXC_FILL, EXC_STROKE, "Exception / rework");
  legendChip(slide, 5.35, y, EXIT_FILL, EXIT_STROKE, "Exit");

  legendChip(slide, 6.9, y, AI_FILL, AI_STROKE, "AI agent (LLM call)");
  legendChip(slide, 8.85, y, DET_FILL, DET_STROKE, "Rules only, no model");
  legendChip(slide, 10.85, y, HUM_FILL, HUM_STROKE, "Human checkpoint");
}
function slideHeader(slide, kicker, subtitle, partLabel) {
  slide.addText(kicker, { x: 0.4, y: 0.13, w: 8.9, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  slide.addText(subtitle, { x: 0.4, y: 0.44, w: 11.6, h: 0.26, fontFace: FONT_BODY, fontSize: 9.5, color: MUTED, isTextBox: true, margin: 0 });
  slide.addText(partLabel, { x: 9.4, y: 0.16, w: 3.53, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });
  legendRow(slide, 0.78);
}
function columnHeaders(slide) {
  slide.addShape("line", { x: 6.63, y: 1.12, w: 0, h: 5.95, line: { color: LINE, width: 0.75 } });
  slide.addText("TODAY", { x: LEFT_X, y: 1.1, w: COL_W, h: 0.26, fontFace: FONT_BODY, fontSize: 11.5, bold: true, color: MUTED, isTextBox: true, margin: 0 });
  slide.addText([
    { text: "AI-NATIVE  ", options: { bold: true, color: AI_STROKE } },
    { text: "— one LangGraph state graph · shared claim state carries the audit trail", options: { bold: false, italic: true, color: MUTED, fontSize: 8.5 } },
  ], { x: RIGHT_X, y: 1.1, w: COL_W, h: 0.26, fontFace: FONT_BODY, fontSize: 11.5, isTextBox: true, margin: 0 });
}
function footer(slide, pageNum) {
  slide.addText("GGW Claims Workflow Redesign · FNOL", { x: 0.4, y: 7.22, w: 6, h: 0.24, fontFace: FONT_BODY, fontSize: 8, color: MUTED, isTextBox: true, margin: 0 });
  slide.addText(String(pageNum) + " / 7", { x: 12.3, y: 7.22, w: 0.65, h: 0.24, fontFace: FONT_BODY, fontSize: 8, color: MUTED, align: "right", isTextBox: true, margin: 0 });
}
function continues(slide, cx, y, text) {
  slide.addText(text, { x: cx - 1.6, y, w: 3.2, h: 0.22, fontFace: FONT_BODY, fontSize: 8.3, italic: true, color: MUTED, align: "center", isTextBox: true, margin: 0 });
}

// Draws a branch box beside a diamond (never wider than BOX_W, so it is
// guaranteed to stay inside this diamond's own column) and the short arrow
// connecting them. Returns the box's placed geometry for a later rejoin call.
function branch(slide, cx, diaCy, side, edgeText, boxText, kind, opts = {}) {
  const [fill, stroke, textColor] = kind === "exception" ? [EXC_FILL, EXC_STROKE, EXC_TEXT]
    : kind === "exit" ? [EXIT_FILL, EXIT_STROKE, EXIT_TEXT] : [AI_FILL, AI_STROKE, AI_TEXT];
  const tipX = side === "left" ? cx - DIA_HALF : cx + DIA_HALF;
  const w = opts.boxW || BOX_W;
  const fs = opts.fontSize || 8;
  const h = fitH(boxText, w, fs);
  const boxX = side === "left" ? tipX - GAP - w : tipX + GAP;
  const boxY = diaCy - h / 2;
  diagArrow(slide, tipX, diaCy, side === "left" ? boxX + w : boxX, diaCy, opts.dashed);
  edgeLabel(slide, edgeText, side === "left" ? boxX - 0.1 : boxX, diaCy - h / 2 - 0.24, w + 0.2, { align: side === "left" ? "right" : "left", italic: true });
  rectNode(slide, boxText, boxX, boxY, w, fill, stroke, textColor, { fontSize: fs, dashed: opts.dashed, minH: h });
  return { x: boxX, y: boxY, w, h, side };
}
// Simple rejoin: box sits immediately before the very next diamond -- a
// short diagonal is safe because nothing sits between them.
function rejoinSimple(slide, box, targetX, targetY, dashed) {
  const fromX = box.side === "left" ? box.x + box.w : box.x;
  diagArrow(slide, fromX, box.y + box.h, targetX, targetY, dashed);
}
// Elbow rejoin: box must skip PAST an intermediate diamond to reach a
// later point in the spine. Runs straight down in the box's own outer
// lane (never crossing the diamond, which sits at cx +/- DIA_HALF) then
// cuts inward only once level with the target row.
function rejoinElbow(slide, box, targetX, targetY) {
  const runX = box.side === "left" ? box.x : box.x + box.w;
  elbowDown(slide, runX, box.y + box.h, targetY);
  elbowAcross(slide, runX, targetX, targetY);
}

// =========================================================================
// SLIDE 1 — Executive summary (answer first)
// =========================================================================
const DEEP = "0B3B3C", AMBER = "D97706", TEAL = "0891B2", MAG = "BE185D", INDIGO = "4338CA";
const AMBER_BG = "FDF3E3", TEAL_BG = "E4F4F8", MAG_BG = "FBE9F1", INDIGO_BG = "EDECFB";

{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 1.16, fill: { color: DEEP }, line: { type: "none" } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 1.16, w: 13.333, h: 0.05, fill: { color: TEAL }, line: { type: "none" } });
  s.addText("EXECUTIVE SUMMARY", {
    x: 9.6, y: 0.17, w: 3.23, h: 0.26, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: "7FB3B6",
    charSpacing: 1.6, align: "right", isTextBox: true, margin: 0,
  });
  s.addText(`Redesigning intake releases ${F.hoursDay} handler-hours a day — the capacity of ${F.fte} full-time people`, {
    x: 0.5, y: 0.15, w: 9.0, h: 0.44, fontFace: FONT_HEAD, fontSize: 19, bold: true, color: "FFFFFF", isTextBox: true, margin: 0,
  });
  s.addText("Two thirds of that work goes to claims that were always going to close quietly · a person still decides anything that can go against a claimant",
    { x: 0.5, y: 0.63, w: 12.3, h: 0.36, fontFace: FONT_BODY, fontSize: 10.5, color: "9FC9CB", isTextBox: true, margin: 0 });

  const py = 1.45, ph = 2.72, pw = 3.75, pwC = 4.23;
  const pA = 0.5, pB = 4.55, pC = 8.6;

  function panel(x, w, tint, accent, eyebrow) {
    s.addShape(pres.ShapeType.roundRect, { x, y: py, w, h: ph, rectRadius: 0.06, fill: { color: tint }, line: { color: accent, width: 1.1 } });
    s.addText(eyebrow.toUpperCase(), { x: x + 0.22, y: py + 0.15, w: w - 0.44, h: 0.24, fontFace: FONT_BODY, fontSize: 8.8, bold: true, color: accent, charSpacing: 1, isTextBox: true, margin: 0 });
  }

  panel(pA, pw, AMBER_BG, AMBER, "Today");
  s.addText("What one claim costs one handler today, start to finish", { x: pA + 0.22, y: py + 0.44, w: pw - 0.44, h: 0.4, fontFace: FONT_BODY, fontSize: 9, color: "44403A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  s.addText([{ text: F.stdTodayInt, options: { fontSize: 42, bold: true, color: AMBER } }, { text: "  minutes per claim", options: { fontSize: 12, color: "6B6257" } }],
    { x: pA + 0.22, y: py + 0.86, w: pw - 0.44, h: 0.6, fontFace: FONT_HEAD, isTextBox: true, margin: 0 });
  s.addText("Reading and re-reading the report, opening attachments, chasing a missing CPR, switching between EASY, IDB, e-mail and paper. Plus 30% for breaks and interruptions.",
    { x: pA + 0.22, y: py + 1.5, w: pw - 0.44, h: 1.08, fontFace: FONT_BODY, fontSize: 8.4, color: "44403A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.18 });

  panel(pB, pw, TEAL_BG, TEAL, "Redesigned");
  s.addText("The same checks and the same exits — reordered, some reassigned", { x: pB + 0.22, y: py + 0.44, w: pw - 0.44, h: 0.4, fontFace: FONT_BODY, fontSize: 9, color: "34474A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  s.addText([{ text: F.bankedAfterInt, options: { fontSize: 42, bold: true, color: TEAL } }, { text: "  minutes per claim", options: { fontSize: 12, color: "5A6E70" } }],
    { x: pB + 0.22, y: py + 0.86, w: pw - 0.44, h: 0.6, fontFace: FONT_HEAD, isTextBox: true, margin: 0 });
  s.addText(`A ${F.cutPct}% cut, not the ${F.designCutPct}% the design allows: handlers will re-check drafts before they trust them. Four AI agents, one rules-only step, two human checkpoints.`,
    { x: pB + 0.22, y: py + 1.5, w: pw - 0.44, h: 1.08, fontFace: FONT_BODY, fontSize: 8.4, color: "34474A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.18 });

  panel(pC, pwC, INDIGO_BG, INDIGO, "Where the effort actually goes");
  const split = [
    ["5%", "Models and machines", `€${F.perClaimEurR} a claim · €${F.runYrR} a year`],
    ["25%", "Plumbing and audit trail", "EASY/IDB wiring, logging every decision"],
    ["70%", "Process, training, trust", "adoption is the real schedule risk"],
  ];
  let sy = py + 0.44;
  split.forEach(([pct, label, sub]) => {
    s.addShape(pres.ShapeType.rect, { x: pC + 0.22, y: sy + 0.02, w: 0.055, h: 0.58, fill: { color: INDIGO }, line: { type: "none" } });
    s.addText(pct, { x: pC + 0.36, y: sy - 0.02, w: 0.72, h: 0.3, fontFace: FONT_HEAD, fontSize: 17, bold: true, color: INDIGO, isTextBox: true, margin: 0 });
    s.addText(label, { x: pC + 1.12, y: sy - 0.03, w: pwC - 1.34, h: 0.24, fontFace: FONT_BODY, fontSize: 9.3, bold: true, color: "2A2A3E", isTextBox: true, margin: 0 });
    s.addText(sub, { x: pC + 1.12, y: sy + 0.2, w: pwC - 1.34, h: 0.4, fontFace: FONT_BODY, fontSize: 8.2, color: "4A4A5E", isTextBox: true, margin: 0, lineSpacingMultiple: 1.1 });
    sy += 0.66;
  });
  s.addText("Follows BCG's 10-20-70 rule, adjusted for this process", { x: pC + 0.22, y: py + ph - 0.28, w: pwC - 0.44, h: 0.24, fontFace: FONT_BODY, fontSize: 7.4, italic: true, color: "5A5A6E", isTextBox: true, margin: 0 });

  // KPI strip
  const kpis = [
    [`${F.hoursDay} h`, "handler-hours freed every working day, across the claims team", AMBER],
    [`${F.fte} FTE`, "the same capacity, expressed as full-time people", TEAL],
    [`€${F.valueYrK}k`, "value of that capacity, per year", MAG],
    [`${F.payback} mths`, `to pay back the one-off build of €${F.buildK}k`, INDIGO],
  ];
  let kx = 0.5; const kw = 2.873, kgap = 0.28;
  kpis.forEach(([big, small, col]) => {
    s.addShape(pres.ShapeType.roundRect, { x: kx, y: 4.34, w: kw, h: 1.0, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: col, width: 1.4 } });
    s.addText(big, { x: kx + 0.18, y: 4.41, w: kw - 0.36, h: 0.42, fontFace: FONT_HEAD, fontSize: 20, bold: true, color: col, isTextBox: true, margin: 0 });
    s.addText(small, { x: kx + 0.18, y: 4.83, w: kw - 0.36, h: 0.44, fontFace: FONT_BODY, fontSize: 8.4, color: "4A4A4A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    kx += kw + kgap;
  });

  // definition + so-what
  s.addShape(pres.ShapeType.roundRect, { x: 0.5, y: 5.52, w: 6.15, h: 1.0, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
  s.addText([
    { text: `What “€${F.valueYrK}k” means:  `, options: { bold: true, color: MAG } },
    { text: `${F.hoursYr} handler-hours a year × €${F.hourly}, the real cost of one handler-hour. This is capacity, not cash. It becomes money only if the hours go to the backlog or are absorbed by turnover — and only if intake sits with a small team, not spread thinly across all 30 (slide 6).`, options: { color: "3A3A3A" } },
  ], { x: 0.68, y: 5.52, w: 5.8, h: 1.0, fontFace: FONT_BODY, fontSize: 8.5, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.18 });

  s.addShape(pres.ShapeType.roundRect, { x: 6.85, y: 5.52, w: 5.98, h: 1.0, rectRadius: 0.06, fill: { color: MAG_BG }, line: { color: MAG, width: 1.2 } });
  s.addText([
    { text: "So what:  ", options: { bold: true, color: MAG } },
    { text: `cost is not what decides this. The AI costs €${F.runYrR} a year to run; keeping it healthy takes a quarter of an engineer, €${F.maintYrR} — so €${F.allInR} a year all in, against €${F.valueYrR} of capacity released. What decides it is governance: proving a person reviewed anything that can go against a claimant, and whether handlers trust the drafts.`, options: { color: "3A2A32" } },
  ], { x: 7.03, y: 5.52, w: 5.62, h: 1.0, fontFace: FONT_BODY, fontSize: 8.5, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.18 });

  s.addText(`Conservative throughout: the ${F.pertInt}-minute-per-claim saving is a three-point estimate (best ${F.pertO} / likely ${F.pertM} / worst ${F.pertP}), not the design maximum. Build-up on slide 6, sources on slide 7.`,
    { x: 0.5, y: 6.62, w: 12.33, h: 0.28, fontFace: FONT_BODY, fontSize: 7.6, italic: true, color: "6A6A6A", isTextBox: true, margin: 0 });
  footer(s, 1);
}

// =========================================================================
// SLIDE 1 of 3 — Entry & Classification
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "Three ways in and a separate side-track diagram become one classified intake", "FNOL as it runs today (left) against the redesign (right) · first report to an opened, policy-linked claim", "FNOL REDESIGN · 1 OF 3");
  columnHeaders(s);
  footer(s, 2);

  // ---- LEFT: today ----
  rectNode(s, "First report filed\n(employer almost always; sometimes union or injured)", LC - 2.3, 1.55, 4.6, START_FILL, START_STROKE, START_TEXT, { fontSize: 9.5 });
  vArrow(s, LC, 1.95, 2.35);

  diamondNode(s, "Channel?", LC, 2.35, 0.7);
  const b1r = branch(s, LC, 2.70, "right", "Direct to AES, secure mail", "AES already gathered the info — insurer only re-checks coverage", "exit", { fontSize: 7.6 });
  branch(s, LC, 2.70, "left", "Teeth/glasses: 'tooth' · email + invoices", "Side track — not PD/LOEC, see 1.1", "exception", { dashed: true, fontSize: 7.8 });
  edgeLabel(s, "EASY / IDB auto-import", LC + 0.12, 3.06, 2.4, { align: "left" });
  vArrow(s, LC, 3.05, 3.55);

  diamondNode(s, "Auto-import\nworked?", LC, 3.55, 0.7);
  const b2l = branch(s, LC, 3.90, "left", "No", "Manual claim creation", "exception", { fontSize: 8.5 });
  rejoinSimple(s, b2l, LC, 4.55);
  edgeLabel(s, "Yes", LC + 0.12, 4.28, 1.0, { align: "left" });

  vArrow(s, LC, 4.25, 4.7);
  rejoinElbow(s, b1r, LC, 4.68);
  continues(s, LC, 4.80, "▼ continues, part 2 · Data / Policy / Duplicate checks");

  // ---- RIGHT: AI-native ----
  const boxW = COL_W - 0.6, boxX = RC - boxW / 2;
  rectNode(s, "Claim arrives — any channel, any format (EASY, AES secure mail, email)", boxX, 1.55, boxW, START_FILL, START_STROKE, START_TEXT, { fontSize: 9.5 });
  vArrow(s, RC, 1.95, 2.35);

  const h1 = rectNode(s, "Extraction Agent\nreads any channel/format · extracts structured fields · matches CPR / CVR",
    boxX, 2.35, boxW, AI_FILL, AI_STROKE, AI_TEXT, { fontSize: 10, bold: true });
  vArrow(s, RC, 2.35 + h1, 2.35 + h1 + 0.2);
  const y2 = 2.35 + h1 + 0.2;
  const h2 = rectNode(s, "Classification Agent\nlabels claim type (standard, teeth, glasses, violence-cover) with a confidence score",
    boxX, y2, boxW, AI_FILL, AI_STROKE, AI_TEXT, { fontSize: 10, bold: true });
  vArrow(s, RC, y2 + h2, y2 + h2 + 0.35);
  const dY = y2 + h2 + 0.35;

  diamondNode(s, "Confidence ≥\nthreshold?", RC, dY, 0.7);
  edgeLabel(s, "threshold starts strict, loosens as override-rate data accumulates",
    RC + DIA_HALF + 0.15, dY + 0.12, 1.85, { align: "left", italic: true, fontSize: 7.2, h: 0.46 });
  branch(s, RC, dY + 0.35, "left", "No", "Falls back to today's manual queue — never silently wrong", "exception", { fontSize: 7.8 });
  edgeLabel(s, "Yes", RC + 0.12, dY + 0.73, 1.0, { align: "left" });
  vArrow(s, RC, dY + 0.70, dY + 1.15);
  continues(s, RC, dY + 1.23, "▼ continues, part 2 · Parallel Validation");

  const cardY = dY + 1.55;
  s.addShape(pres.ShapeType.roundRect, { x: boxX - 0.2, y: cardY, w: boxW + 0.4, h: 0.62, rectRadius: 0.05, fill: { color: PAPER }, line: { color: LINE, width: 1 } });
  s.addText("These two agents together replace 3 entry channels, the auto-import failure path, and the entire 1.1 side-track diagram.", {
    x: boxX - 0.05, y: cardY, w: boxW + 0.1, h: 0.62, fontFace: FONT_BODY, fontSize: 8.6, color: "3A4250", isTextBox: true, margin: 4, valign: "middle", align: "center", lineSpacingMultiple: 1.15,
  });
}

// =========================================================================
// SLIDE 2 of 3 — Validation Gates
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "Three checks done one after another become one check done all at once", "Is it complete, linked to the right policy, and not already in the system? None of the three depends on the others", "FNOL REDESIGN · 2 OF 3");
  columnHeaders(s);
  footer(s, 3);

  continues(s, LC, 1.40, "▲ continued from part 1");
  vArrow(s, LC, 1.62, 1.80);

  diamondNode(s, "Data correct\nand complete?", LC, 1.80, 0.7);
  const c1 = branch(s, LC, 2.15, "left", "Missing CPR", "Ask / fill fields (via questionnaire)", "exception", { fontSize: 7.8 });
  rejoinSimple(s, c1, LC, 2.90);
  edgeLabel(s, "OK", LC + 0.12, 2.45, 1.0, { align: "left" });
  vArrow(s, LC, 2.42, 2.90);

  diamondNode(s, "Valid policy\nattached?", LC, 2.90, 0.7);
  const c2l = branch(s, LC, 3.25, "left", "Wrong / old policy", "Reattach correct policy", "exception", { fontSize: 8 });
  rejoinSimple(s, c2l, LC, 4.08);
  branch(s, LC, 3.25, "right", "No valid policy", "Back to employer: correct insurer", "exit", { fontSize: 7.8 });
  edgeLabel(s, "Yes", LC + 0.12, 3.63, 1.0, { align: "left" });
  vArrow(s, LC, 3.60, 4.08);

  diamondNode(s, "Earlier or\nduplicate?", LC, 4.08, 0.7);
  const c3 = branch(s, LC, 4.43, "left", "Duplicate", "Dedup / connect same-CPR claims", "exception", { fontSize: 7.8 });
  rejoinSimple(s, c3, LC, 5.30);
  edgeLabel(s, "New", LC + 0.12, 4.81, 1.0, { align: "left" });
  vArrow(s, LC, 4.78, 5.30);
  continues(s, LC, 5.40, "▼ continues, part 3 · Escalation & Close-Out");

  const boxW = COL_W - 0.6, boxX = RC - boxW / 2;
  continues(s, RC, 1.40, "▲ continued from part 1");
  vArrow(s, RC, 1.62, 1.80);

  const h1 = rectNode(s, "Validation Node — rules only\nCPR dedup · policy in force · completeness — three lookups at once, no model call",
    boxX, 1.80, boxW, DET_FILL, DET_STROKE, DET_TEXT, { fontSize: 10, bold: true });
  vArrow(s, RC, 1.80 + h1, 1.80 + h1 + 0.35);
  const dY = 1.80 + h1 + 0.35;

  diamondNode(s, "Any check\nfailed?", RC, dY, 0.7);
  const exBox = branch(s, RC, dY + 0.35, "right", "Yes", "Exception Research Agent — gathers CVR / fuzzy-match candidates as evidence", "exception", { fontSize: 7.6 });
  vArrow(s, exBox.x + exBox.w / 2, exBox.y + exBox.h, exBox.y + exBox.h + 0.24);
  rectNode(s, "Underwriter decides\n(AI prepares, never decides)", exBox.x, exBox.y + exBox.h + 0.24, exBox.w, HUM_FILL, HUM_STROKE, HUM_TEXT, { fontSize: 7.6 });
  edgeLabel(s, "No", RC + 0.12, dY + 0.73, 1.0, { align: "left" });
  vArrow(s, RC, dY + 0.70, 5.30);
  continues(s, RC, 5.40, "▼ continues, part 3 · Outcome Drafting Agent");

  s.addShape(pres.ShapeType.roundRect, { x: RIGHT_X, y: 6.05, w: 3.3, h: 0.72, rectRadius: 0.05, fill: { color: PAPER }, line: { color: LINE, width: 1 } });
  s.addText("3 sequential gates, each with its own rework loop, become 1 parallel check.", {
    x: RIGHT_X + 0.1, y: 6.05, w: 3.1, h: 0.72, fontFace: FONT_BODY, fontSize: 8.4, color: "3A4250", isTextBox: true, margin: 4, valign: "middle", align: "center", lineSpacingMultiple: 1.2,
  });
}

// =========================================================================
// SLIDE 3 of 3 — Escalation & Close-Out
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "The AI writes the draft; the handler still makes the decision", "From a system warning to a closed claim or a fully opened one — nothing that can go against a claimant is automated", "FNOL REDESIGN · 3 OF 3");
  columnHeaders(s);
  footer(s, 4);

  continues(s, LC, 1.40, "▲ continued from part 2");
  vArrow(s, LC, 1.62, 1.80);

  diamondNode(s, "Red warning\non claim?", LC, 1.80, 0.7);
  const d1 = branch(s, LC, 2.15, "left", "Yes", "On hold: check with UW before it can move on", "exception", { fontSize: 7.6 });
  rejoinSimple(s, d1, LC, 3.05);
  edgeLabel(s, "No", LC + 0.12, 2.45, 1.0, { align: "left" });
  vArrow(s, LC, 2.42, 3.05);

  diamondNode(s, "Injury minor?", LC, 3.05, 0.7);
  branch(s, LC, 3.40, "left", "Yes, ~65% (likely more)", "Close at intake — letter, no permanent injury", "exit", { fontSize: 7.6 });
  branch(s, LC, 3.40, "right", "No", "Claim in IDB: fields checked, task opened, questionnaire", "exit", { fontSize: 7.6 });

  s.addShape(pres.ShapeType.roundRect, { x: LEFT_X, y: 4.55, w: COL_W, h: 0.75, rectRadius: 0.05, fill: { color: PAPER }, line: { color: LINE, width: 1 } });
  s.addText("Today: 7 decision gates, 3 entry channels, 10 exception or exit branches across the full FNOL flow.", {
    x: LEFT_X + 0.15, y: 4.55, w: COL_W - 0.3, h: 0.75, fontFace: FONT_BODY, fontSize: 9, color: "3A4250", isTextBox: true, margin: 4, valign: "middle", align: "center", lineSpacingMultiple: 1.2,
  });

  const boxW = COL_W - 0.6, boxX = RC - boxW / 2;
  continues(s, RC, 1.40, "▲ continued from part 2");
  vArrow(s, RC, 1.62, 1.80);

  const h1 = rectNode(s, "Outcome Drafting Agent\ndrafts the close-out letter, or the full IDB record, with its reasoning attached",
    boxX, 1.80, boxW, AI_FILL, AI_STROKE, AI_TEXT, { fontSize: 10, bold: true });
  vArrow(s, RC, 1.80 + h1, 1.80 + h1 + 0.22);
  const humY = 1.80 + h1 + 0.22;
  const hH = rectNode(s, "Handler approves before anything sends — edit rate is the monitored metric, not model accuracy",
    boxX + 0.5, humY, boxW - 1.0, HUM_FILL, HUM_STROKE, HUM_TEXT, { fontSize: 8.6, bold: true });
  vArrow(s, RC, humY + hH, humY + hH + 0.32);
  const dY = humY + hH + 0.32;

  diamondNode(s, "Injury minor?\n(~65%)", RC, dY, 0.7);
  branch(s, RC, dY + 0.35, "left", "Yes", "Close at intake — AI-drafted, handler approved", "exit", { fontSize: 7.6 });
  branch(s, RC, dY + 0.35, "right", "No", "Claim in IDB — runs parallel with Coverage Check", "exit", { fontSize: 7.6 });

  s.addShape(pres.ShapeType.roundRect, { x: boxX - 0.2, y: dY + 1.15, w: boxW + 0.4, h: 1.0, rectRadius: 0.05, fill: { color: PAPER }, line: { color: LINE, width: 1 } });
  s.addText([
    { text: "Net result across all three parts:\n", options: { bold: true, color: INK, breakLine: true } },
    { text: "4 LLM agents + 1 rules-only node · 2 human checkpoints kept by design · 3 entry channels → 1 · 3 sequential gates → 1 parallel check · shared claim state carries the audit trail", options: { color: "3A4250" } },
  ], { x: boxX - 0.05, y: dY + 1.15, w: boxW + 0.1, h: 1.0, fontFace: FONT_BODY, fontSize: 9, isTextBox: true, margin: 4, valign: "middle", lineSpacingMultiple: 1.2, align: "center" });
}

// =========================================================================
// SLIDE 5 — Sequencing quadrant (decluttered: 2 items max per quadrant)
// palette validated for colour-vision separation, not eyeballed
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Start where a mistake is cheap — that is what earns the trust the bigger wins depend on", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0,
  });
  s.addText("SEQUENCING", { x: 10.4, y: 0.16, w: 2.53, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });
  s.addText("“Hard” means build effort and regulatory exposure together, not just weeks of work", {
    x: 0.4, y: 0.5, w: 12.5, h: 0.26, fontFace: FONT_BODY, fontSize: 9.5, color: MUTED, isTextBox: true, margin: 0 });

  const gx = 1.15, gy = 1.0, gw = 8.6, gh = 5.0;
  const hw = gw / 2, hh = gh / 2;

  s.addText("HOW OFTEN IT HAPPENS  →", { x: 0.02, y: gy + gh / 2 - 0.14, w: 1.75, h: 0.28, fontFace: FONT_BODY, fontSize: 8, bold: true, color: MUTED, align: "center", isTextBox: true, margin: 0, rotate: 270 });
  s.addText("HOW HARD AND HOW RISKY IT IS  →", { x: gx, y: gy + gh + 0.1, w: gw, h: 0.24, fontFace: FONT_BODY, fontSize: 8, bold: true, color: MUTED, align: "center", isTextBox: true, margin: 0 });

  function quad(qx, qy, bg, stroke, label, sub, items) {
    s.addShape(pres.ShapeType.rect, { x: qx, y: qy, w: hw, h: hh, fill: { color: bg }, line: { color: "FFFFFF", width: 2.5 } });
    s.addText(label.toUpperCase(), { x: qx + 0.2, y: qy + 0.16, w: hw - 0.4, h: 0.26, fontFace: FONT_BODY, fontSize: 10, bold: true, color: stroke, charSpacing: 0.8, isTextBox: true, margin: 0 });
    s.addText(sub, { x: qx + 0.2, y: qy + 0.42, w: hw - 0.4, h: 0.24, fontFace: FONT_BODY, fontSize: 8, italic: true, color: MUTED, isTextBox: true, margin: 0 });
    items.forEach(([title, meta], i) => {
      const w = hw - 0.5, h = 0.62, y = qy + 0.78 + i * (h + 0.14);
      s.addShape(pres.ShapeType.roundRect, { x: qx + 0.25, y, w, h, rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: stroke, width: 1.2 } });
      s.addText([
        { text: title + "\n", options: { bold: true, color: INK, fontSize: 9.2, breakLine: true } },
        { text: meta, options: { color: MUTED, fontSize: 8 } },
      ], { x: qx + 0.4, y, w: w - 0.3, h, fontFace: FONT_BODY, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.15 });
    });
  }

  quad(gx, gy, "FDF3E3", "D97706", "Wave 1 — start here", "cheap to get wrong, happens constantly", [
    ["Read the claim and classify it", "ends the one-person dependency on teeth claims"],
    ["Check CPR, policy and duplicates", "plain rules, no AI model"],
  ]);
  quad(gx + hw, gy, "E4F4F8", "0891B2", "Wave 2 — once trusted", "the real prize, but needs the audit trail live", [
    ["Draft the close-out letter", "~65% of claims · a handler still approves"],
    ["Read AES secure mail", "the only channel with no automation today"],
  ]);
  quad(gx, gy + hh, "EDECFB", "4338CA", "Fix when convenient", "worth doing, but it is an IT defect", [
    ["EASY → IDB import retries", "an integration bug, not an AI job"],
  ]);
  quad(gx + hw, gy + hh, "FBE9F1", "BE185D", "Leave manual", "too rare to earn the effort", [
    ["\"Violence cover\" entitlement letter", "~2 claims ever · a checklist beats a model"],
  ]);

  // How the one IT item actually gets fixed -- the detail an engineer would ask for
  s.addShape(pres.ShapeType.roundRect, { x: 10.05, y: 1.0, w: 2.88, h: 5.0, rectRadius: 0.06, fill: { color: "EDECFB" }, line: { color: "4338CA", width: 1.3 } });
  s.addText("FIXING THE IMPORT FAILURES", { x: 10.27, y: 1.16, w: 2.44, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: "4338CA", charSpacing: 0.7, isTextBox: true, margin: 0 });
  s.addText("The one item that is not an AI problem. How it is actually fixed:", { x: 10.27, y: 1.42, w: 2.44, h: 0.36, fontFace: FONT_BODY, fontSize: 8, italic: true, color: "4A4A5E", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  const fixes = [
    ["Sort the failure first", "A timeout is worth retrying. A missing field never is — retrying it just fails again."],
    ["Retry with growing gaps", "1s, 2s, 4s, plus a random offset, so 50 stuck claims do not all retry at once."],
    ["One reference per claim", "The same key on every attempt, so a replay cannot create the claim twice."],
    ["Park failures, never drop them", "They go to a holding queue with the error, so someone can fix and replay."],
    ["Alert on the rate, not the count", "One stuck claim is normal. Twenty in an hour means something broke."],
  ];
  let fy = 1.92;
  fixes.forEach(([t, d]) => {
    s.addText(t, { x: 10.27, y: fy, w: 2.44, h: 0.2, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: "2A2A3E", isTextBox: true, margin: 0 });
    s.addText(d, { x: 10.27, y: fy + 0.19, w: 2.44, h: 0.56, fontFace: FONT_BODY, fontSize: 7.6, color: "4A4A5E", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    fy += 0.79;
  });

  s.addText("AES rulings and underwriting authority are deliberately absent — those decisions are not ours to make · frequencies from the case study", { x: 0.4, y: 6.6, w: 12.5, h: 0.28, fontFace: FONT_BODY, fontSize: 7, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 5);
}

// =========================================================================
// =========================================================================
// SLIDE 6 — How the numbers are built (backup / defensible detail)
// Three accent colours only: AMBER = time, TEAL = cost, MAG = conclusion.
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Every figure is built from observed steps and list prices — none of it is a best case", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("HOW THE NUMBERS ARE BUILT", { x: 8.9, y: 0.16, w: 4.03, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  const L = 0.4, R = 6.85, CW = 6.0, RH = 0.235;
  const blockHead = (t, x, y, col, w) => s.addText(t, { x, y, w: w || CW, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: col, charSpacing: 0.8, isTextBox: true, margin: 0 });
  const note = (t, x, y, h) => s.addText(t, { x, y, w: CW, h, fontFace: FONT_BODY, fontSize: 7.4, italic: true, color: "6A6A6A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  const tbl = (rows, x, y, colW, rh) => s.addTable(rows, {
    x, y, w: CW, colW, fontFace: FONT_BODY, fontSize: 7.8,
    border: { type: "solid", color: "E4E4E0", pt: 0.5 }, color: "3A3A3A",
    align: "left", valign: "middle", rowH: rh || RH, autoPage: false, fill: { color: "FFFFFF" } });
  const B = (t) => ({ text: t, options: { bold: true } });

  // ---------- A · time (left) ----------
  blockHead("A \u00b7 HANDLER TIME \u2014 MINUTES PER CLAIM, PER HANDLER", L, 0.58, AMBER);
  tbl([
    [B("Step"), B("Today"), B("Redesigned")],
    ...F.rows,
    [B("Touch time"), B(F.touchToday), B(F.touchAfter)],
    [`Breaks, fatigue, interruptions, rework (+${F.allowPct}%)`, F.allowToday, F.allowAfter],
    [B("Standard time \u2014 what the design allows"), B(F.stdToday), B(F.stdAfter)],
    [B("What the business case actually banks"), B(F.stdToday), B(F.bankedAfter)],
  ], L, 0.86, [3.9, 1.05, 1.05], 0.205);
  note(`The design allows ${F.stdAfter} min a claim. We do not bank that. The case uses a three-point estimate \u2014 (best ${F.pertO} + 4 \u00d7 likely ${F.pertM} + worst ${F.pertP}) \u00f7 6 = ${F.pert} min saved a claim \u2014 which is the ${F.bankedAfter} min above, and the ${F.bankedAfterInt} minutes on slide 1. In plain terms: we assume handlers capture ${F.capturePct}% of the saving the design allows.`, L, 3.42, 0.36);

  s.addShape(pres.ShapeType.roundRect, { x: L, y: 3.82, w: CW, h: 0.60, rectRadius: 0.05, fill: { color: "FDF3E3" }, line: { color: AMBER, width: 1 } });
  s.addText([
    { text: "Cross-check:  ", options: { bold: true, color: "9A5A0A" } },
    { text: `${F.stdTodayInt} min \u00d7 ${F.claimsYr} claims a year = ${F.todayHoursYr} handler-hours \u2014 ${F.teamSharePct}% of a 30-person team's yearly capacity. Plausible for the front end of a claims process. Had it come out at 60%, the model would be wrong.`, options: { color: "44403A" } },
  ], { x: L + 0.18, y: 3.82, w: CW - 0.36, h: 0.60, fontFace: FONT_BODY, fontSize: 8.2, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.15 });

  // ---------- D · business case (left, bottom) ----------
  // Two sections, because the unit differs. "Cost to run" means the system
  // only; the engineer who maintains it is a separate, clearly named line.
  s.addShape(pres.ShapeType.roundRect, { x: L, y: 4.50, w: CW, h: 2.62, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: MAG, width: 1.2 } });
  blockHead("D \u00b7 THE BUSINESS CASE", L + 0.2, 4.60, MAG, CW - 0.4);

  const eyebrow = (t, y) => s.addText(t, { x: L + 0.2, y, w: CW - 0.4, h: 0.20, fontFace: FONT_BODY, fontSize: 7.6, bold: true, color: MAG, charSpacing: 1.2, isTextBox: true, margin: 0 });
  const money = (k, v, y, hi) => {
    s.addText(k, { x: L + 0.2, y, w: 3.85, h: 0.21, fontFace: FONT_BODY, fontSize: 8.4, color: hi ? INK : "3A3A3A", bold: !!hi, isTextBox: true, margin: 0 });
    s.addText(v, { x: L + 4.05, y, w: 1.75, h: 0.21, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: hi ? MAG : INK, align: "right", isTextBox: true, margin: 0 });
  };

  eyebrow("EVERY YEAR", 4.84);
  [["Capacity released", `${F.hoursYr} handler-h \u00b7 ${F.fte} FTE`, 0],
   [`Worth, at \u20ac${F.hourly} per handler-hour`, `\u20ac${F.valueYrR} a year`, 0],
   ["Cost to run it \u2014 AI models and Azure", `\u2212 \u20ac${F.runYrR}`, 0],
   ["Ongoing support \u2014 0.25 of an engineer", `\u2212 \u20ac${F.maintYrR}`, 0],
   ["Net value, every year", `\u20ac${F.netYrR}`, 1],
  ].forEach(([k, v, hi], i) => money(k, v, 5.04 + i * 0.215, hi));

  s.addShape(pres.ShapeType.rect, { x: L + 0.2, y: 6.14, w: CW - 0.4, h: 0.012, fill: { color: "E4E4E0" }, line: { type: "none" } });

  eyebrow("ONE-OFF", 6.20);
  [["Build \u2014 3 people \u00d7 6 months, fully loaded", `\u20ac${F.buildR}`, 1],
   ["Payback", `${F.payback} months`, 1],
  ].forEach(([k, v, hi], i) => money(k, v, 6.42 + i * 0.215, hi));

  s.addText(`The build is loaded team capacity, not new cash \u2014 the AI team is already funded. New cash needed: \u20ac${F.runYrR} a year. A shorter build costs less: 4 months \u2248 \u20ac98k.`,
    { x: L + 0.2, y: 6.86, w: CW - 0.4, h: 0.24, fontFace: FONT_BODY, fontSize: 7.2, italic: true, color: "6A6A6A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });

  // ---------- B · AI cost (right) ----------
  blockHead("B · AI COST TO PROCESS ONE CLAIM", R, 0.58, TEAL);
  tbl([
    [B("Agent"), B("Tokens in"), B("Out"), B("Runs"), B("US$ / claim")],
    ["Extraction", "11,500", "700", "100%", "0.0314"],
    ["Classification", "900", "120", "100%", "0.0032"],
    ["Validation — rules only, no model", "—", "—", "100%", "0.0000"],
    ["Exception research", "2,500", "400", "20%", "0.0020"],
    ["Outcome drafting", "2,200", "550", "100%", "0.0110"],
    [B("Subtotal"), "", "", "", B("0.0476")],
    ["× 1.75 for retries, guardrails, evals", "", "", "", "0.0833"],
    [B("Per claim"), "", "", "", B("≈ €0.077")],
  ], R, 0.86, [2.65, 1.0, 0.7, 0.65, 1.0]);
  note("One claim ≈ a scanned form, an e-mail and 3–4 attachments. All four agents run on one model tier — gpt-5.6 terra, $2 in / $12 out per million tokens.", R, 3.00, 0.26);

  // ---------- C · infra (right) ----------
  blockHead("C · AZURE, WEST EUROPE — MONTHLY", R, 3.32, TEAL);
  tbl([
    [B("Component"), B("US$/mth")],
    ["4 container apps — 2 always on, 2 burst only", "55"],
    ["PostgreSQL B2s + 64 GiB storage and backup", "61"],
    ["Blob storage for documents + audit logs", "8"],
    [B("Net new, incl. ~8% West Europe premium"), B("134")],
  ], R, 3.60, [4.85, 1.15]);
  note("Registry, networking, Key Vault and DevOps (≈ €670/yr) already exist, so they are not charged here. Azure DevOps gives repos and pipelines, not a database — PostgreSQL is genuinely new.", R, 4.80, 0.28);

  // ---------- E · who gets the time back (right, bottom) ----------
  blockHead("E · WHO ACTUALLY GETS THE TIME BACK", R, 5.14, MAG);
  s.addText("The brief says a support team does intake, not all 30 handlers — so the truth is in the middle rows. Our first question for GGW.",
    { x: R, y: 5.36, w: CW, h: 0.22, fontFace: FONT_BODY, fontSize: 7.8, italic: true, color: "6A1F42", isTextBox: true, margin: 0 });
  tbl([
    [B("Intake handled by…"), B("Claims each, per day"), B("Intake work, per day"), B("Time freed, per day")],
    ["all 30 claims handlers", "1.7", "1.1 h", "35 min"],
    ["a support team of 12", "4.2", "2.8 h", "88 min"],
    ["a support team of 8", "6.2", "4.2 h", "2.2 h"],
    ["a support team of 6", "8.3", "5.6 h", "2.9 h"],
  ], R, 5.62, [2.1, 1.35, 1.3, 1.25]);

  footer(s, 6);
}
// =========================================================================
// SLIDE 7 — Sources (clickable) and what is assumption vs. fact
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("What is sourced, and what is our assumption — kept separate on purpose", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("SOURCES & ASSUMPTIONS", { x: 9.4, y: 0.16, w: 3.53, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  const SRC = [
    ["Wage base", "Sachbearbeiter Schadenregulierung, €45,100 gross",
      "StepStone Gehaltsreport", "https://www.stepstone.de/gehalt/Schadenregulierer-in.html"],
    ["Employer on-cost", "≈ €23 per €100 gross — health, pension, unemployment, care",
      "Lohnnebenkosten 2026, sevdesk", "https://sevdesk.de/ratgeber/buchhaltung-finanzen/lohnbuchhaltung/lohnnebenkosten/"],
    ["Working time, breaks", "38 h tariff week; statutory break rules",
      "Arbeitszeitgesetz §4", "https://www.gesetze-im-internet.de/arbzg/__4.html"],
    ["Model prices", "gpt-5.6 terra — $2 in / $12 out per million tokens",
      "OpenAI API pricing", "https://openai.com/api/pricing/"],
    ["Container hosting", "Consumption plan rates and the monthly free grant",
      "Azure Container Apps pricing", "https://azure.microsoft.com/en-us/pricing/details/container-apps/"],
    ["Database", "PostgreSQL Flexible Server B2s — 2 vCore / 4 GiB",
      "Azure Database for PostgreSQL", "https://azure.microsoft.com/en-us/pricing/details/postgresql/flexible-server/"],
    ["Build-team pay", "Senior ML engineer ≈ €98k; juniors at €70k as briefed",
      "Glassdoor, ML Engineer Germany", "https://www.glassdoor.com/Salaries/germany-machine-learning-engineer-salary-SRCH_IL.0,7_IN96_KO8,33.htm"],
    ["Upkeep after go-live", "15–20% of build effort a year — our 0.25 FTE is 17%",
      "Software maintenance rule of thumb", ""],
    ["Industry baseline", "17 of 20 carriers still manual; none let AI decide alone",
      "Decerto, 2026 Claims Decisioning Pulse", "https://www.decerto.com/us/post/ai-claims-decisioning-framework-2026-what-us-claims-leaders-told-us-about-how-decisions-actually-get-made"],
    ["Where value comes from", "The 10-20-70 split — algorithms, technology, people",
      "BCG, AI in customer service operations", "https://www.bcg.com/publications/2024/transforming-customer-service-operations-with-genai"],
    ["Fixing import failures", "Retry vs. dead-letter, idempotency keys, alert on rate",
      "Dead-letter queues and poison messages", "https://www.glukhov.org/app-architecture/integration-patterns/dead-letter-queues/"],
    ["Volumes and shares", "~50 first reports a day · 65% minor · ~30 handlers",
      "GGW case study brief", ""],
  ];

  let y = 0.58;
  SRC.forEach(([label, what, name, url]) => {
    s.addShape(pres.ShapeType.rect, { x: 0.4, y: y + 0.02, w: 0.05, h: 0.34, fill: { color: TEAL }, line: { type: "none" } });
    s.addText(label, { x: 0.58, y: y, w: 2.15, h: 0.20, fontFace: FONT_BODY, fontSize: 8.6, bold: true, color: INK, isTextBox: true, margin: 0 });
    s.addText(what, { x: 0.58, y: y + 0.17, w: 5.4, h: 0.20, fontFace: FONT_BODY, fontSize: 8, color: "5A5A5A", isTextBox: true, margin: 0 });
    s.addText(name, {
      x: 6.2, y: y + 0.06, w: 6.6, h: 0.24, fontFace: FONT_BODY, fontSize: 8.4,
      color: url ? "0B5F8A" : "5A5A5A", italic: !url, isTextBox: true, margin: 0,
      ...(url ? { hyperlink: { url, tooltip: name } } : {}),
    });
    y += 0.40;
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.4, y: 5.42, w: 12.43, h: 1.20, rectRadius: 0.06, fill: { color: "FDF3E3" }, line: { color: AMBER, width: 1.2 } });
  s.addText("OUR ASSUMPTIONS — NOT SOURCED, AND THE FIRST THING WE WOULD TEST", {
    x: 0.6, y: 5.53, w: 12.0, h: 0.24, fontFace: FONT_BODY, fontSize: 8.6, bold: true, color: "9A5A0A", charSpacing: 0.6, isTextBox: true, margin: 0 });
  s.addText("Every per-step minute on slide 6 · the 30% allowance for breaks and interruptions · how often a CPR is missing (30%) or underwriting is asked (25%) · how often the redesign raises an exception (20%) · the tokens a typical claim needs · the €146k build · the 0.25 FTE upkeep · and how many people staff intake, which sets the per-person saving.\n\nTwo weeks of watching live intake would replace all of these with real measurements. That is step one of the plan, not an afterthought.",
    { x: 0.6, y: 5.78, w: 12.0, h: 0.76, fontFace: FONT_BODY, fontSize: 8.4, color: "44403A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  s.addText("Prices checked September 2026 · US$ converted at €0.92 · every euro figure is fully loaded — gross pay plus employer contributions",
    { x: 0.4, y: 6.70, w: 12.5, h: 0.26, fontFace: FONT_BODY, fontSize: 7, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 7);
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user-ggw-casestudy/2b8e5a67-8934-50d1-9acc-621207f6c407/scratchpad/deck/GGW_FNOL_Redesign.pptx" })
  .then(() => console.log("written"))
  .catch((e) => { console.error(e); process.exit(1); });
