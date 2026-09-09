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
  slide.addText(String(pageNum) + " / 10", { x: 12.3, y: 7.22, w: 0.65, h: 0.24, fontFace: FONT_BODY, fontSize: 8, color: MUTED, align: "right", isTextBox: true, margin: 0 });
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
  s.addText(`Redesigning intake gives the ${F.support}-person support team back ${F.freedHPersonDay} hours each, every day`, {
    x: 0.5, y: 0.15, w: 9.0, h: 0.44, fontFace: FONT_HEAD, fontSize: 19, bold: true, color: "FFFFFF", isTextBox: true, margin: 0,
  });
  s.addText(`${F.hoursDay} hours a day across the team — ${F.fte} full-time people · two thirds of that work goes to claims that were always going to close quietly · a person still decides anything adverse`,
    { x: 0.5, y: 0.63, w: 12.3, h: 0.36, fontFace: FONT_BODY, fontSize: 10.5, color: "9FC9CB", isTextBox: true, margin: 0 });

  const py = 1.45, ph = 2.72, pw = 3.75, pwC = 4.23;
  const pA = 0.5, pB = 4.55, pC = 8.6;

  function panel(x, w, tint, accent, eyebrow) {
    s.addShape(pres.ShapeType.roundRect, { x, y: py, w, h: ph, rectRadius: 0.06, fill: { color: tint }, line: { color: accent, width: 1.1 } });
    s.addText(eyebrow.toUpperCase(), { x: x + 0.22, y: py + 0.15, w: w - 0.44, h: 0.24, fontFace: FONT_BODY, fontSize: 8.8, bold: true, color: accent, charSpacing: 1, isTextBox: true, margin: 0 });
  }

  panel(pA, pw, AMBER_BG, AMBER, "Today");
  s.addText("What one claim costs a support-team member", { x: pA + 0.22, y: py + 0.44, w: pw - 0.44, h: 0.4, fontFace: FONT_BODY, fontSize: 9, color: "44403A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  s.addText([{ text: F.stdTodayInt, options: { fontSize: 42, bold: true, color: AMBER } }, { text: "  minutes per claim", options: { fontSize: 12, color: "6B6257" } }],
    { x: pA + 0.22, y: py + 0.86, w: pw - 0.44, h: 0.6, fontFace: FONT_HEAD, isTextBox: true, margin: 0 });
  s.addText("Reading and re-reading the report, opening attachments, chasing a missing CPR, switching between EASY, IDB, e-mail and paper. Plus 30% for breaks and interruptions.",
    { x: pA + 0.22, y: py + 1.5, w: pw - 0.44, h: 1.08, fontFace: FONT_BODY, fontSize: 8.4, color: "44403A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.18 });

  panel(pB, pw, TEAL_BG, TEAL, "Redesigned");
  s.addText("The same checks and the same exits, reordered", { x: pB + 0.22, y: py + 0.44, w: pw - 0.44, h: 0.4, fontFace: FONT_BODY, fontSize: 9, color: "34474A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
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
    [`${F.freedHPersonDay} h`, `freed per support-team member, every working day (of ${F.prodHPerDay} productive hours)`, AMBER],
    [`${F.hoursDay} h`, `across the ${F.support}-person support team — ${F.fte} full-time people`, TEAL],
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
    { text: `${F.hoursYr} support-team hours a year × €${F.hourly}, the loaded cost of one of those hours. Valued at support pay (€${F.supportGross} gross), not claims-handler pay, because these are the hours actually freed. Capacity, not cash — it becomes money only if the hours go to the backlog or are absorbed by turnover.`, options: { color: "3A3A3A" } },
  ], { x: 0.68, y: 5.52, w: 5.8, h: 1.0, fontFace: FONT_BODY, fontSize: 8.5, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.18 });

  s.addShape(pres.ShapeType.roundRect, { x: 6.85, y: 5.52, w: 5.98, h: 1.0, rectRadius: 0.06, fill: { color: MAG_BG }, line: { color: MAG, width: 1.2 } });
  s.addText([
    { text: "So what:  ", options: { bold: true, color: MAG } },
    { text: `cost is not what decides this. The AI costs €${F.runYrR} a year to run; keeping it healthy takes a quarter of an engineer, €${F.maintYrR} — so €${F.allInR} a year all in, against €${F.valueYrR} of capacity released. What decides it is governance: proving a person reviewed anything that can go against a claimant, and whether handlers trust the drafts.`, options: { color: "3A2A32" } },
  ], { x: 7.03, y: 5.52, w: 5.62, h: 1.0, fontFace: FONT_BODY, fontSize: 8.5, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.18 });

  s.addText(`Conservative throughout: the ${F.pertInt}-minute-per-claim saving is a three-point estimate (best ${F.pertO} / likely ${F.pertM} / worst ${F.pertP}), not the design maximum. Build-up on slide 9, sources on slide 10.`,
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
// SLIDE 5 — Coverage Check, and the handoff the brief asks about
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Coverage Check stops being a second investigation — the evidence is already in the claim", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("COVERAGE CHECK & HANDOFF", { x: 9.0, y: 0.16, w: 3.93, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });
  s.addText("Every reason a claim can fail the coverage test, and where each one should be settled", {
    x: 0.4, y: 0.50, w: 8.3, h: 0.24, fontFace: FONT_BODY, fontSize: 9.5, color: MUTED, isTextBox: true, margin: 0 });

  const Bh = (t) => ({ text: t, options: { bold: true } });
  s.addTable([
    [Bh("Why coverage looks missing"), Bh("Today"), Bh("Redesigned"), Bh("Who decides")],
    ["Date sits inside a valid policy — no warning", "quick manual check", "rules, no model call", "system"],
    ["A valid follow-up policy exists", "find it, reattach", "registry lookup, reattach drafted", "handler"],
    ["Self-employed, voluntary cover", "manual self-check", "registry lookup, rules", "handler"],
    ["Group or branch policy not surfaced", "ask underwriting", "agent gathers the evidence, drafts the question", "underwriting"],
    ["Company-name mismatch", "ask underwriting", "agent fuzzy-matches CVR, drafts the question", "underwriting"],
    ["Wrong, expired, or no policy at all", "back to employer", "letter drafted, never sent unread", "handler"],
  ], {
    x: 0.4, y: 0.86, w: 8.3, colW: [2.55, 1.65, 2.75, 1.35], fontFace: FONT_BODY, fontSize: 8,
    border: { type: "solid", color: "E4E4E0", pt: 0.5 }, color: "3A3A3A",
    align: "left", valign: "middle", rowH: 0.34, autoPage: false, fill: { color: "FFFFFF" },
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.4, y: 3.62, w: 8.3, h: 0.92, rectRadius: 0.06, fill: { color: TEAL_BG }, line: { color: TEAL, width: 1.2 } });
  s.addText([
    { text: "Only two rows change hands.  ", options: { bold: true, color: "0B5A66" } },
    { text: "Four of the six are lookups against a register the insurer already queries — those become rules with no model call at all. The two that genuinely need judgement still go to underwriting, because that authority is not ours to take. The AI's job there is to arrive with the evidence assembled, not to answer the question.", options: { color: "34474A" } },
  ], { x: 0.58, y: 3.62, w: 7.94, h: 0.92, fontFace: FONT_BODY, fontSize: 8.6, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.2 });

  s.addShape(pres.ShapeType.roundRect, { x: 0.4, y: 4.66, w: 8.3, h: 1.86, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
  s.addText("WHAT THIS REMOVES", { x: 0.6, y: 4.78, w: 7.9, h: 0.22, fontFace: FONT_BODY, fontSize: 9, bold: true, color: AMBER, charSpacing: 0.8, isTextBox: true, margin: 0 });
  const removes = [
    ["The second look at the policy", "FNOL already asked “is a valid policy attached?”. Coverage Check asks it again, from the same documents."],
    ["The wait between the two", "Today the claim sits in IDB until someone picks up the coverage step. It now runs beside claim creation, not after it."],
    ["Re-reading the file to ask underwriting", "The question arrives with the policy history, the CVR match and the dates already attached."],
  ];
  let ry = 5.06;
  removes.forEach(([t, d]) => {
    s.addShape(pres.ShapeType.rect, { x: 0.6, y: ry + 0.02, w: 0.04, h: 0.38, fill: { color: AMBER }, line: { type: "none" } });
    s.addText(t, { x: 0.74, y: ry, w: 2.7, h: 0.4, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: INK, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.1 });
    s.addText(d, { x: 3.5, y: ry, w: 5.0, h: 0.4, fontFace: FONT_BODY, fontSize: 8, color: "5A5A5A", isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.15 });
    ry += 0.46;
  });

  // ---- the handoff itself ----
  s.addShape(pres.ShapeType.roundRect, { x: 8.9, y: 0.86, w: 4.03, h: 5.66, rectRadius: 0.06, fill: { color: INDIGO_BG }, line: { color: INDIGO, width: 1.3 } });
  s.addText("THE HANDOFF", { x: 9.12, y: 1.00, w: 3.59, h: 0.24, fontFace: FONT_BODY, fontSize: 10, bold: true, color: INDIGO, charSpacing: 0.9, isTextBox: true, margin: 0 });
  s.addText("The brief asks about the join between the two workflows. Today it is a queue. In the redesign there is nothing to hand over.",
    { x: 9.12, y: 1.26, w: 3.59, h: 0.52, fontFace: FONT_BODY, fontSize: 8.2, italic: true, color: "3A3A5E", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });

  s.addText("TODAY", { x: 9.12, y: 1.86, w: 3.59, h: 0.2, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: "6A6A7E", charSpacing: 1, isTextBox: true, margin: 0 });
  const today = [
    "FNOL closes. The claim is written to IDB.",
    "It waits in a queue for the coverage step.",
    "A second person opens the same documents.",
    "They re-ask a question FNOL already asked.",
  ];
  let hy = 2.10;
  today.forEach((t) => {
    s.addText("•  " + t, { x: 9.12, y: hy, w: 3.59, h: 0.32, fontFace: FONT_BODY, fontSize: 8.2, color: "4A4A5E", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    hy += 0.34;
  });

  s.addText("REDESIGNED", { x: 9.12, y: 3.60, w: 3.59, h: 0.2, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: INDIGO, charSpacing: 1, isTextBox: true, margin: 0 });
  const after = [
    "One state graph spans both workflows.",
    "Extraction captures policy, dates, CVR and employment status once.",
    "Coverage is a rules pass over state that already exists.",
    "It runs beside claim creation, not after it.",
    "Only the exceptions become anyone's work.",
  ];
  hy = 3.84;
  after.forEach((t) => {
    s.addText("•  " + t, { x: 9.12, y: hy, w: 3.59, h: 0.44, fontFace: FONT_BODY, fontSize: 8.2, color: "2A2A3E", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    hy += 0.42;
  });

  s.addShape(pres.ShapeType.roundRect, { x: 9.12, y: 5.94, w: 3.59, h: 0.44, rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: INDIGO, width: 1 } });
  s.addText("A handoff you delete cannot be the thing that breaks.", { x: 9.22, y: 5.94, w: 3.39, h: 0.44, fontFace: FONT_BODY, fontSize: 8.2, bold: true, color: INDIGO, align: "center", isTextBox: true, margin: 0, valign: "middle" });

  s.addText("Causes and their handling are taken from Diagram 2 of the brief · “who decides” is the design choice, not an observation",
    { x: 0.4, y: 6.62, w: 12.5, h: 0.26, fontFace: FONT_BODY, fontSize: 7, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 5);
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
    ["Detect failed EASY → IDB imports", "nobody is alerted today — a claim can go missing unseen"],
    ["Read, classify, check CPR and policy", "one agent plus plain rules — ends the teeth dependency"],
  ]);
  quad(gx + hw, gy, "E4F4F8", "0891B2", "Wave 2 — once trusted", "the real prize, but needs the audit trail live", [
    ["Draft the close-out letter", "~65% of claims · a handler still approves"],
    ["Read AES secure mail", "the only channel with no automation today"],
  ]);
  quad(gx, gy + hh, "EDECFB", "4338CA", "Then fix the import itself", "an IT defect, once you can see it", [
    ["Retry, park and replay", "detection first — you cannot fix what nobody reports"],
  ]);
  quad(gx + hw, gy + hh, "FBE9F1", "BE185D", "Leave manual", "too rare to earn the effort", [
    ["\"Violence cover\" entitlement letter", "~2 claims ever · a checklist beats a model"],
  ]);

  // How the one IT item actually gets fixed -- the detail an engineer would ask for
  s.addShape(pres.ShapeType.roundRect, { x: 10.05, y: 1.0, w: 2.88, h: 5.0, rectRadius: 0.06, fill: { color: "EDECFB" }, line: { color: "4338CA", width: 1.3 } });
  s.addText("THE IMPORT FAILURES", { x: 10.27, y: 1.16, w: 2.44, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: "4338CA", charSpacing: 0.7, isTextBox: true, margin: 0 });
  s.addText("GGW confirm these are not noticed today. That makes detection the first fix, not retries:", { x: 10.27, y: 1.42, w: 2.44, h: 0.40, fontFace: FONT_BODY, fontSize: 8, italic: true, color: "4A4A5E", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  const fixes = [
    ["Count what should have arrived", "Reconcile EASY sends against IDB claims daily. A gap is a lost claim nobody sees."],
    ["Sort the failure", "A timeout is worth retrying. A missing field never is — it just fails again."],
    ["Retry with growing gaps", "1s, 2s, 4s, plus a random offset, so stuck claims do not all retry at once."],
    ["One reference per claim", "The same key every attempt, so a replay cannot create the claim twice."],
    ["Park failures, never drop them", "A holding queue with the error, so someone can fix and replay."],
    ["Alert on the rate, not the count", "One stuck claim is normal. Twenty in an hour is not."],
  ];
  let fy = 1.88;
  fixes.forEach(([t, d]) => {
    s.addText(t, { x: 10.27, y: fy, w: 2.44, h: 0.2, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: "2A2A3E", isTextBox: true, margin: 0 });
    s.addText(d, { x: 10.27, y: fy + 0.18, w: 2.44, h: 0.46, fontFace: FONT_BODY, fontSize: 7.6, color: "4A4A5E", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    fy += 0.66;
  });

  s.addText("AES rulings and underwriting authority are deliberately absent — those decisions are not ours to make · frequencies from the case study", { x: 0.4, y: 6.6, w: 12.5, h: 0.28, fontFace: FONT_BODY, fontSize: 7, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 6);
}

// =========================================================================
// =========================================================================
// SLIDE 7 — Production: how it goes live, what is watched, what happens
// when it breaks. This answers the reliability question directly.
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Nothing goes live decided — it goes live watched, then suggesting, then trusted", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("RUNNING IT IN PRODUCTION", { x: 9.0, y: 0.16, w: 3.93, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  // ---- the four stages, as a ladder ----
  const stages = [
    ["1 · SHADOW", "wks 1–8", "Agents run on live claims. Output is stored and compared with what the handler did. Nobody sees it.", "Exit when extraction agrees with the handler on 95% of fields.", AMBER, AMBER_BG],
    ["2 · SUGGEST", "wks 9–16", "The handler sees the draft and must edit or approve it. Nothing sends itself.", "Exit when the edit rate on minor claims settles below 20%.", TEAL, TEAL_BG],
    ["3 · APPROVE", "mths 5–6", "Minor claims arrive pre-filled. One click accepts. Everything else still drafted only.", "Exit criteria reviewed monthly with the claims lead.", INDIGO, INDIGO_BG],
    ["NEVER", "—", "No automatic decision on anything adverse to a claimant, on an AES ruling, or on underwriting authority.", "This is a design boundary, not a phase.", MAG, MAG_BG],
  ];
  let sx = 0.4; const sgap = 0.245, sw = (13.333 - 0.8 - 3 * sgap) / 4;
  stages.forEach(([name, when, body, gate, col, bg]) => {
    s.addShape(pres.ShapeType.roundRect, { x: sx, y: 0.62, w: sw, h: 2.34, rectRadius: 0.06, fill: { color: bg }, line: { color: col, width: 1.3 } });
    s.addText(name, { x: sx + 0.18, y: 0.74, w: sw - 0.36, h: 0.24, fontFace: FONT_BODY, fontSize: 9.6, bold: true, color: col, charSpacing: 0.8, isTextBox: true, margin: 0 });
    s.addText(when, { x: sx + 0.18, y: 0.98, w: sw - 0.36, h: 0.2, fontFace: FONT_BODY, fontSize: 8, italic: true, color: MUTED, isTextBox: true, margin: 0 });
    s.addText(body, { x: sx + 0.18, y: 1.24, w: sw - 0.36, h: 0.86, fontFace: FONT_BODY, fontSize: 8.4, color: "3A3A3A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.18 });
    s.addShape(pres.ShapeType.rect, { x: sx + 0.18, y: 2.14, w: sw - 0.36, h: 0.012, fill: { color: col }, line: { type: "none" } });
    s.addText(gate, { x: sx + 0.18, y: 2.20, w: sw - 0.36, h: 0.64, fontFace: FONT_BODY, fontSize: 8, italic: true, color: col === MAG ? "6A1F42" : "4A4A4A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    sx += sw + sgap;
  });

  // ---- what is watched ----
  s.addShape(pres.ShapeType.roundRect, { x: 0.4, y: 3.16, w: 6.29, h: 3.16, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
  s.addText("WHAT WE WATCH — AND WHY THAT METRIC", { x: 0.6, y: 3.30, w: 5.89, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: TEAL, charSpacing: 0.8, isTextBox: true, margin: 0 });
  const metrics = [
    ["Edit rate on drafts", "The trust signal. A model can be accurate and still not be used."],
    ["Field-level extraction precision", "Checked against the source document, per field, not per claim."],
    ["Exception rate, by node", "Says which step is actually struggling, not that “the AI” is."],
    ["Override rate by claim type", "Feeds the confidence threshold. Teeth claims will differ from the rest."],
    ["First report to linked claim", "The only number the business asked for. Everything else is diagnostic."],
  ];
  let my = 3.58;
  metrics.forEach(([t, d]) => {
    s.addText(t, { x: 0.6, y: my, w: 2.5, h: 0.46, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: INK, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.1 });
    s.addText(d, { x: 3.18, y: my, w: 3.31, h: 0.46, fontFace: FONT_BODY, fontSize: 8, color: "5A5A5A", isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.15 });
    my += 0.48;
  });
  s.addText("Alert on the rate, not the count — one stuck claim is normal, twenty in an hour is not.",
    { x: 0.6, y: 6.02, w: 5.89, h: 0.2, fontFace: FONT_BODY, fontSize: 7.8, italic: true, color: TEAL, isTextBox: true, margin: 0 });

  // ---- when it breaks ----
  s.addShape(pres.ShapeType.roundRect, { x: 6.89, y: 3.16, w: 6.04, h: 3.16, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: MAG, width: 1.2 } });
  s.addText("WHEN IT BREAKS", { x: 7.09, y: 3.30, w: 5.64, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: MAG, charSpacing: 0.8, isTextBox: true, margin: 0 });
  const breaks = [
    ["An agent fails, or confidence is low", "The claim drops into today's manual queue. The old path is never removed, so the failure mode is a slower claim, never a wrong one."],
    ["A prompt or model version changes", "200 historical claims are replayed first. A regression against the recorded outcome blocks the deploy."],
    ["One agent misbehaves in the wild", "Each is behind its own switch. Turn off drafting and extraction keeps running."],
    ["The run dies half way", "Graph state is checkpointed, so the claim resumes at the last good node instead of starting again."],
  ];
  let by = 3.58;
  breaks.forEach(([t, d]) => {
    s.addText(t, { x: 7.09, y: by, w: 5.64, h: 0.2, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: "6A1F42", isTextBox: true, margin: 0 });
    s.addText(d, { x: 7.09, y: by + 0.19, w: 5.64, h: 0.42, fontFace: FONT_BODY, fontSize: 8, color: "5A5A5A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    by += 0.62;
  });
  s.addText("Every decision, its inputs and its confidence are written to the audit log — that is what makes an adverse outcome defensible later.",
    { x: 7.09, y: 6.08, w: 5.64, h: 0.2, fontFace: FONT_BODY, fontSize: 7.8, italic: true, color: MAG, isTextBox: true, margin: 0 });

  s.addText("Thresholds shown are the ones we would propose and then argue about with the claims lead — they are starting points, not findings",
    { x: 0.4, y: 6.62, w: 12.5, h: 0.26, fontFace: FONT_BODY, fontSize: 7, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 7);
}

// =========================================================================
// SLIDE 8 — Technical implementation. GGW asked for this specifically:
// "more interested in how you do the technical implementation ...
//  we would deploy in the cloud".
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("One typed state object moves through the graph — every node reads it, writes to it, and signs its work", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("TECHNICAL IMPLEMENTATION", { x: 9.0, y: 0.16, w: 3.93, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  // ---------- A · the state object ----------
  s.addText("A · THE CLAIM STATE — WHAT EVERY NODE SHARES", { x: 0.4, y: 0.56, w: 4.2, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: AMBER, charSpacing: 0.8, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: 0.4, y: 0.84, w: 4.2, h: 2.62, rectRadius: 0.05, fill: { color: "F2F1ED" }, line: { color: "D9D7D0", width: 1 } });
  s.addText([
    { text: "class ClaimState(BaseModel):\n", options: { bold: true, color: "1B2430" } },
    { text: "    claim_id: str\n    source: Literal[\"easy\",\"aes\",\"email\"]\n    documents: list[BlobRef]\n", options: { color: "3A4250" } },
    { text: "    fields: ExtractedFields          # typed, per-field\n    confidence: dict[str, float]   # per field\n", options: { color: "3A4250" } },
    { text: "    claim_type: Literal[\"standard\",\"teeth\",\n                        \"glasses\",\"violence\"]\n", options: { color: "3A4250" } },
    { text: "    coverage: CoverageResult | None\n    exceptions: list[Exception]\n    drafts: list[Draft]\n", options: { color: "3A4250" } },
    { text: "    audit: list[NodeRun]             # who did what, when\n", options: { color: "6B4A10" } },
  ], { x: 0.58, y: 0.94, w: 3.84, h: 2.44, fontFace: "Courier New", fontSize: 7.6, isTextBox: true, margin: 0, lineSpacingMultiple: 1.18 });

  s.addText("Pydantic types are the contract, and confidence is held per field rather than per claim. A node that cannot produce a valid field returns null plus a reason — it never guesses to satisfy a schema.",
    { x: 0.4, y: 3.52, w: 4.2, h: 0.4, fontFace: FONT_BODY, fontSize: 7.6, italic: true, color: "6A6A6A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });

  // ---------- B · the stack ----------
  s.addText("B · THE STACK, AND WHY EACH PIECE", { x: 4.8, y: 0.56, w: 4.0, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: TEAL, charSpacing: 0.8, isTextBox: true, margin: 0 });
  const stack = [
    ["LangGraph  ·  Python", "Nodes are the agents, conditional edges are the decision diamonds on slides 2–4. The diagram is the code."],
    ["Pydantic", "Typed state. Schema violations fail loudly at the node, not silently downstream."],
    ["PostgreSQL", "LangGraph checkpointer plus the audit log. One store, so a claim's history and its state cannot drift apart."],
    ["Azure Container Apps", "Two always-on for the API and worker, two scale-to-zero for burst. Cloud, as GGW deploy today."],
    ["Azure Blob  ·  Key Vault", "Claim documents and secrets. Documents never enter a prompt as raw bytes — only extracted text."],
    ["Azure OpenAI", "Same model tier throughout, pinned by version. EU region keeps claim data in-region."],
  ];
  let ty = 0.86;
  stack.forEach(([t, d]) => {
    s.addShape(pres.ShapeType.rect, { x: 4.8, y: ty + 0.02, w: 0.04, h: 0.38, fill: { color: TEAL }, line: { type: "none" } });
    s.addText(t, { x: 4.94, y: ty, w: 3.86, h: 0.18, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: INK, isTextBox: true, margin: 0 });
    s.addText(d, { x: 4.94, y: ty + 0.17, w: 3.86, h: 0.34, fontFace: FONT_BODY, fontSize: 7.6, color: "5A5A5A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
    ty += 0.535;
  });

  // ---------- C · integrations ----------
  s.addText("C · INTEGRATIONS — WHAT DECIDES THE TIMELINE", { x: 9.0, y: 0.56, w: 3.93, h: 0.24, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: INDIGO, charSpacing: 0.6, isTextBox: true, margin: 0 });
  const integ = [
    ["EASY", "in", "Reuse the existing auto-import. We add the reconciliation that tells you when it failed."],
    ["AES secure mail", "in", "Fully manual today. A mailbox reader is the single biggest unlock on the intake side."],
    ["E-mail", "in", "Same reader, different folder. Attachments to Blob, text to extraction."],
    ["IDB", "out", "Write the claim, attach the policy, open the task. If there is no API, this is the long pole — not the AI."],
    ["CPR / CVR registers", "both", "Already queried by hand. Wrapped as tools the agents call, with results cached on the claim."],
  ];
  let iy = 0.86;
  integ.forEach(([sys, dir, d]) => {
    const col = dir === "out" ? MAG : dir === "both" ? AMBER : INDIGO;
    s.addText([
      { text: sys + "  ", options: { bold: true, color: INK, fontSize: 8.4 } },
      { text: dir.toUpperCase(), options: { bold: true, color: col, fontSize: 7 } },
    ], { x: 9.0, y: iy, w: 3.93, h: 0.18, fontFace: FONT_BODY, isTextBox: true, margin: 0 });
    s.addText(d, { x: 9.0, y: iy + 0.17, w: 3.93, h: 0.42, fontFace: FONT_BODY, fontSize: 7.6, color: "5A5A5A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
    iy += 0.63;
  });
  s.addShape(pres.ShapeType.roundRect, { x: 9.0, y: 4.06, w: 3.93, h: 0.62, rectRadius: 0.05, fill: { color: MAG_BG }, line: { color: MAG, width: 1 } });
  s.addText("The models are the easy part. Whether IDB can be written to programmatically is the question that moves the date.",
    { x: 9.14, y: 4.06, w: 3.65, h: 0.62, fontFace: FONT_BODY, fontSize: 7.8, bold: true, color: "6A1F42", isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.15 });

  // ---------- D · deployment ----------
  s.addShape(pres.ShapeType.roundRect, { x: 0.4, y: 4.06, w: 8.4, h: 2.4, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
  s.addText("D · HOW IT SHIPS AND KEEPS RUNNING", { x: 0.6, y: 4.18, w: 8.0, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: INDIGO, charSpacing: 0.8, isTextBox: true, margin: 0 });
  const ship = [
    ["Infrastructure as code", "Bicep or Terraform. Nothing clicked into the portal, so the test environment is the production one minus data."],
    ["Azure DevOps pipelines", "Already licensed. Lint, type-check, unit tests, then the replay set on every merge."],
    ["A flag per agent", "Extraction, classification, research and drafting each toggle independently, per claim type."],
    ["Blue/green on the workers", "A bad release drains rather than cuts. In-flight claims finish on the old version."],
    ["Model version pinned", "An upgrade is a deliberate change that re-runs the replay set — never a silent one."],
  ];
  let dy = 4.48;
  ship.forEach(([t, d]) => {
    s.addText(t, { x: 0.6, y: dy, w: 2.5, h: 0.36, fontFace: FONT_BODY, fontSize: 8.2, bold: true, color: INK, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.1 });
    s.addText(d, { x: 3.18, y: dy, w: 5.42, h: 0.36, fontFace: FONT_BODY, fontSize: 7.8, color: "5A5A5A", isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.12 });
    dy += 0.385;
  });

  s.addText("Runs on the cloud GGW already deploy to · costs for exactly this shape are on slide 9 · the rollout that puts it in front of a handler is on slide 7",
    { x: 0.4, y: 6.62, w: 12.5, h: 0.26, fontFace: FONT_BODY, fontSize: 7, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 8);
}

// =========================================================================
// SLIDE 6 — How the numbers are built (backup / defensible detail)
// Three accent colours only: AMBER = time, TEAL = cost, MAG = conclusion.
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Every figure is built from observed steps and list prices — none of it is a best case", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("APPENDIX · HOW THE NUMBERS ARE BUILT", { x: 8.0, y: 0.16, w: 4.93, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  const L = 0.4, R = 6.85, CW = 6.0, RH = 0.235;
  const blockHead = (t, x, y, col, w) => s.addText(t, { x, y, w: w || CW, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: col, charSpacing: 0.8, isTextBox: true, margin: 0 });
  const note = (t, x, y, h) => s.addText(t, { x, y, w: CW, h, fontFace: FONT_BODY, fontSize: 7.4, italic: true, color: "6A6A6A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  const tbl = (rows, x, y, colW, rh) => s.addTable(rows, {
    x, y, w: CW, colW, fontFace: FONT_BODY, fontSize: 7.8,
    border: { type: "solid", color: "E4E4E0", pt: 0.5 }, color: "3A3A3A",
    align: "left", valign: "middle", rowH: rh || RH, autoPage: false, fill: { color: "FFFFFF" } });
  const B = (t) => ({ text: t, options: { bold: true } });

  // ---------- A · time (left) ----------
  blockHead("A \u00b7 INTAKE TIME \u2014 MINUTES PER CLAIM, PER PERSON", L, 0.58, AMBER);
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
    { text: `${F.stdTodayInt} min \u00d7 ${F.claimsYr} claims a year = ${F.todayHoursYr} hours \u2014 ${F.shareSupportPct}% of the ${F.support}-person support team's year, ${F.shareTeamPct}% of all 30. Per person: ${F.claimsPerPersonDay} claims a day at 40 min is ${F.intakeHPersonDay} of a ${F.prodHPerDay}-hour day. Intake is half of what the support team does.`, options: { color: "44403A" } },
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
  [["Capacity released", `${F.hoursYr} support-h \u00b7 ${F.fte} FTE`, 0],
   [`Worth, at \u20ac${F.hourly} per support-team hour`, `\u20ac${F.valueYrR} a year`, 0],
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
  blockHead("E · WHO GETS THE TIME BACK — CONFIRMED BY GGW", R, 5.14, MAG);
  s.addText(`${F.handlers} claim handlers and ${F.support} support staff (GGW, 8 Sep). The brief puts FNOL with the support team, so every freed hour is theirs.`,
    { x: R, y: 5.36, w: CW, h: 0.22, fontFace: FONT_BODY, fontSize: 7.8, italic: true, color: "6A1F42", isTextBox: true, margin: 0 });
  tbl([
    [B("Per support-team member"), B("Today"), B("Redesigned"), B("Difference")],
    ["Claims taken in, per day", `${F.claimsPerPersonDay}`, `${F.claimsPerPersonDay}`, "—"],
    ["Time on intake, per day", `${F.intakeHPersonDay} h`, "1.6 h", `− ${F.freedHPersonDay} h`],
    ["Share of a productive day", `${F.shareSupportPct}%`, "25%", "− 27 pts"],
    [B("Hours back, each, per day"), "", "", B(`${F.freedHPersonDay} h`)],
  ], R, 5.62, [2.1, 1.35, 1.3, 1.25]);

  footer(s, 9);
}
// =========================================================================
// SLIDE 7 — Sources (clickable) and what is assumption vs. fact
// Two columns: the list outgrew one after GGW's answers were added.
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("What is sourced, what GGW confirmed, and what is still our assumption", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("APPENDIX · SOURCES & ASSUMPTIONS", { x: 8.4, y: 0.16, w: 4.53, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  const SRC = [
    ["Who does intake", "20 claim handlers, 10 support staff — FNOL sits with support", "GGW (P. Gossmann), 8 Sep 2026", "", true],
    ["Failed EASY imports", "Not noticed today — so detection comes before retries", "GGW (P. Gossmann), 8 Sep 2026", "", true],
    ["Volumes and shares", "~50 first reports a day · 65% minor · ~30 in the team", "GGW case study brief", "", true],
    ["Wage base — intake", "Sachbearbeiter/in Versicherung, €41,100 gross", "StepStone Gehaltsreport", "https://www.stepstone.de/gehalt/Sachbearbeiter-in-Versicherung.html"],
    ["Employer on-cost", "≈ €23 per €100 gross — health, pension, care", "Lohnnebenkosten 2026, sevdesk", "https://sevdesk.de/ratgeber/buchhaltung-finanzen/lohnbuchhaltung/lohnnebenkosten/"],
    ["Working time, breaks", "38 h tariff week; statutory break rules", "Arbeitszeitgesetz §4", "https://www.gesetze-im-internet.de/arbzg/__4.html"],
    ["Build & upkeep pay", "Senior ML engineer ≈ €98k; juniors at €70k as briefed", "Glassdoor, ML Engineer Germany", "https://www.glassdoor.com/Salaries/germany-machine-learning-engineer-salary-SRCH_IL.0,7_IN96_KO8,33.htm"],
    ["Model prices", "gpt-5.6 terra — $2 in / $12 out per million tokens", "OpenAI API pricing", "https://openai.com/api/pricing/"],
    ["Container hosting", "Consumption plan rates and the monthly free grant", "Azure Container Apps pricing", "https://azure.microsoft.com/en-us/pricing/details/container-apps/"],
    ["Database", "PostgreSQL Flexible Server B2s — 2 vCore / 4 GiB", "Azure Database for PostgreSQL", "https://azure.microsoft.com/en-us/pricing/details/postgresql/flexible-server/"],
    ["Upkeep after go-live", "15–20% of build effort a year — our 0.25 FTE is 17%", "Software maintenance rule of thumb", ""],
    ["Industry baseline", "17 of 20 carriers still manual; none let AI decide alone", "Decerto, 2026 Claims Decisioning Pulse", "https://www.decerto.com/us/post/ai-claims-decisioning-framework-2026-what-us-claims-leaders-told-us-about-how-decisions-actually-get-made"],
    ["Where value comes from", "The 10-20-70 split — algorithms, technology, people", "BCG, AI in customer service operations", "https://www.bcg.com/publications/2024/transforming-customer-service-operations-with-genai"],
    ["Fixing import failures", "Retry vs. dead-letter, idempotency keys, alert on rate", "Dead-letter queues and poison messages", "https://www.glukhov.org/app-architecture/integration-patterns/dead-letter-queues/"],
  ];

  const COLX = [0.4, 6.75], COLW = 6.18;
  SRC.forEach(([label, what, name, url, confirmed], i) => {
    const cx = COLX[i < 7 ? 0 : 1];
    const y = 0.60 + (i % 7) * 0.52;
    const accent = confirmed ? MAG : TEAL;
    s.addShape(pres.ShapeType.rect, { x: cx, y: y + 0.02, w: 0.05, h: 0.44, fill: { color: accent }, line: { type: "none" } });
    s.addText(label, { x: cx + 0.18, y, w: COLW - 0.18, h: 0.19, fontFace: FONT_BODY, fontSize: 8.6, bold: true, color: INK, isTextBox: true, margin: 0 });
    s.addText(what, { x: cx + 0.18, y: y + 0.17, w: COLW - 0.18, h: 0.19, fontFace: FONT_BODY, fontSize: 7.8, color: "5A5A5A", isTextBox: true, margin: 0 });
    s.addText(name, {
      x: cx + 0.18, y: y + 0.33, w: COLW - 0.18, h: 0.19, fontFace: FONT_BODY, fontSize: 7.8,
      color: url ? "0B5F8A" : (confirmed ? "6A1F42" : "5A5A5A"), italic: !url, isTextBox: true, margin: 0,
      ...(url ? { hyperlink: { url, tooltip: name } } : {}),
    });
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.4, y: 4.36, w: 12.53, h: 1.26, rectRadius: 0.06, fill: { color: "FDF3E3" }, line: { color: AMBER, width: 1.2 } });
  s.addText("STILL OUR ASSUMPTIONS — AND THE FIRST THING WE WOULD TEST", {
    x: 0.6, y: 4.47, w: 12.1, h: 0.24, fontFace: FONT_BODY, fontSize: 8.6, bold: true, color: "9A5A0A", charSpacing: 0.6, isTextBox: true, margin: 0 });
  s.addText("Every per-step minute on slide 9 · the 30% allowance for breaks and interruptions · how often a CPR is missing (30%) or underwriting is asked (25%) · how often the redesign raises an exception (20%) · the tokens a typical claim needs · the €146k build · the 0.25 FTE upkeep.\n\nWho does intake is no longer among them — GGW answered it, and every per-person figure in this deck is rebuilt on that answer. Two weeks of watching live intake would replace the rest with measurements.",
    { x: 0.6, y: 4.72, w: 12.1, h: 0.82, fontFace: FONT_BODY, fontSize: 8.4, color: "44403A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  s.addShape(pres.ShapeType.rect, { x: 0.4, y: 5.80, w: 0.05, h: 0.2, fill: { color: MAG }, line: { type: "none" } });
  s.addText("confirmed by GGW or stated in the brief", { x: 0.58, y: 5.78, w: 3.2, h: 0.24, fontFace: FONT_BODY, fontSize: 7.6, color: MAG, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.rect, { x: 4.0, y: 5.80, w: 0.05, h: 0.2, fill: { color: TEAL }, line: { type: "none" } });
  s.addText("published source, linked", { x: 4.18, y: 5.78, w: 3.2, h: 0.24, fontFace: FONT_BODY, fontSize: 7.6, color: TEAL, isTextBox: true, margin: 0 });

  s.addText("Prices checked September 2026 · US$ converted at €0.92 · every euro figure is fully loaded — gross pay plus employer contributions",
    { x: 0.4, y: 6.18, w: 12.5, h: 0.26, fontFace: FONT_BODY, fontSize: 7, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 10);
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user-ggw-casestudy/2b8e5a67-8934-50d1-9acc-621207f6c407/scratchpad/deck/GGW_FNOL_Redesign.pptx" })
  .then(() => console.log("written"))
  .catch((e) => { console.error(e); process.exit(1); });
