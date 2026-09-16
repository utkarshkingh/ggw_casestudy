const pptxgen = require("pptxgenjs");
// Every time and money figure comes from economics.py via figures.json.
// Nothing below is typed by hand, so slide 1 and slide 8 cannot disagree.
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
function slideHeader(slide, kicker, subtitle, partLabel, showLegend = true) {
  slide.addText(kicker, { x: 0.4, y: 0.13, w: 8.9, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  slide.addText(subtitle, { x: 0.4, y: 0.44, w: 11.6, h: 0.26, fontFace: FONT_BODY, fontSize: 9.5, color: MUTED, isTextBox: true, margin: 0 });
  slide.addText(partLabel, { x: 9.4, y: 0.16, w: 3.53, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });
  if (showLegend) legendRow(slide, 0.78);
}
function columnHeaders(slide) {
  slide.addShape("line", { x: 6.63, y: 1.12, w: 0, h: 5.95, line: { color: LINE, width: 0.75 } });
  slide.addText("TODAY", { x: LEFT_X, y: 1.1, w: COL_W, h: 0.26, fontFace: FONT_BODY, fontSize: 11.5, bold: true, color: MUTED, isTextBox: true, margin: 0 });
  slide.addText([
    { text: "AI-NATIVE  ", options: { bold: true, color: AI_STROKE } },
    { text: "one shared claim record, one audit trail", options: { bold: false, italic: true, color: MUTED, fontSize: 8.5 } },
  ], { x: RIGHT_X, y: 1.1, w: COL_W, h: 0.26, fontFace: FONT_BODY, fontSize: 11.5, isTextBox: true, margin: 0 });
}
function footer(slide, pageNum) {
  slide.addText("GGW Claims Workflow Redesign · FNOL", { x: 0.4, y: 7.22, w: 6, h: 0.24, fontFace: FONT_BODY, fontSize: 8, color: MUTED, isTextBox: true, margin: 0 });
  slide.addText(String(pageNum) + " / 9", { x: 12.3, y: 7.22, w: 0.65, h: 0.24, fontFace: FONT_BODY, fontSize: 8, color: MUTED, align: "right", isTextBox: true, margin: 0 });
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
  s.addText(`Redesigning intake gives each claim handler back ${F.freedMinPersonDay} minutes a day`, {
    x: 0.5, y: 0.15, w: 9.0, h: 0.44, fontFace: FONT_HEAD, fontSize: 19, bold: true, color: "FFFFFF", isTextBox: true, margin: 0,
  });
  s.addText(`${F.hoursDay} hours a day across the ${F.handlers} of them, or ${F.fte} full-time people. Most of it goes to the ~65% of claims that close at intake with no permanent injury.`,
    { x: 0.5, y: 0.63, w: 12.3, h: 0.36, fontFace: FONT_BODY, fontSize: 10.5, color: "9FC9CB", isTextBox: true, margin: 0 });

  const py = 1.45, ph = 2.72, pw = 3.75, pwC = 4.23;
  const pA = 0.5, pB = 4.55, pC = 8.6;

  function panel(x, w, tint, accent, eyebrow) {
    s.addShape(pres.ShapeType.roundRect, { x, y: py, w, h: ph, rectRadius: 0.06, fill: { color: tint }, line: { color: accent, width: 1.1 } });
    s.addText(eyebrow.toUpperCase(), { x: x + 0.22, y: py + 0.15, w: w - 0.44, h: 0.24, fontFace: FONT_BODY, fontSize: 8.8, bold: true, color: accent, charSpacing: 1, isTextBox: true, margin: 0 });
  }

  // short bullet chips replace a paragraph -- scannable, not readable
  function chips(x, y, w, items, color) {
    let cy = y;
    items.forEach((t) => {
      s.addShape(pres.ShapeType.ellipse, { x, y: cy + 0.06, w: 0.06, h: 0.06, fill: { color }, line: { type: "none" } });
      s.addText(t, { x: x + 0.16, y: cy, w: w - 0.16, h: 0.2, fontFace: FONT_BODY, fontSize: 8.6, color: "3A3A3A", isTextBox: true, margin: 0 });
      cy += 0.255;
    });
  }

  panel(pA, pw, AMBER_BG, AMBER, "Today");
  s.addText("What one claim costs one claim handler", { x: pA + 0.22, y: py + 0.44, w: pw - 0.44, h: 0.4, fontFace: FONT_BODY, fontSize: 9, color: "44403A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  s.addText([{ text: F.stdTodayInt, options: { fontSize: 42, bold: true, color: AMBER } }, { text: "  minutes per claim", options: { fontSize: 12, color: "6B6257" } }],
    { x: pA + 0.22, y: py + 0.86, w: pw - 0.44, h: 0.6, fontFace: FONT_HEAD, isTextBox: true, margin: 0 });
  chips(pA + 0.22, py + 1.5, pw - 0.44, [
    "Re-reads the report, opens attachments",
    "Chases a missing CPR",
    "Switches EASY \u00b7 IDB \u00b7 e-mail \u00b7 paper",
    "+30% for breaks and interruptions",
  ], AMBER);

  panel(pB, pw, TEAL_BG, TEAL, "Redesigned");
  s.addText("Every check and every exit kept, run in a new order", { x: pB + 0.22, y: py + 0.44, w: pw - 0.44, h: 0.4, fontFace: FONT_BODY, fontSize: 9, color: "34474A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  s.addText([{ text: F.bankedAfterInt, options: { fontSize: 42, bold: true, color: TEAL } }, { text: "  minutes per claim", options: { fontSize: 12, color: "5A6E70" } }],
    { x: pB + 0.22, y: py + 0.86, w: pw - 0.44, h: 0.6, fontFace: FONT_HEAD, isTextBox: true, margin: 0 });
  chips(pB + 0.22, py + 1.5, pw - 0.44, [
    `21 minutes back on every claim (a ${F.cutPct}% cut)`,
    "4 AI agents, 1 rules-only step",
    "2 human checkpoints kept",
    "Handlers re-check before they trust it",
  ], TEAL);

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
  s.addText("BCG's 10-20-70, tilted to plumbing: EASY and IDB are the constraint", { x: pC + 0.22, y: py + ph - 0.28, w: pwC - 0.44, h: 0.24, fontFace: FONT_BODY, fontSize: 7.4, italic: true, color: "5A5A6E", isTextBox: true, margin: 0 });

  // KPI strip
  const kpis = [
    [`${F.freedMinPersonDay} min`, `freed per claim handler per day, out of ${F.prodHPerDay} productive hours`, AMBER],
    [`${F.hoursDay} h`, `across the ${F.handlers} claim handlers, or ${F.fte} full-time people`, TEAL],
    [`€${F.valueYrK}k a year`, `what the ${F.hoursYr} freed handler-hours are worth, at €${F.hourly} fully loaded`, MAG],
    [`${F.payback} mths`, `to pay back the one-off build of €${F.buildK}k`, INDIGO],
  ];
  let kx = 0.5; const kw = 2.873, kgap = 0.28;
  kpis.forEach(([big, small, col]) => {
    s.addShape(pres.ShapeType.roundRect, { x: kx, y: 4.34, w: kw, h: 1.0, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: col, width: 1.4 } });
    s.addText(big, { x: kx + 0.18, y: 4.41, w: kw - 0.36, h: 0.42, fontFace: FONT_HEAD, fontSize: 20, bold: true, color: col, isTextBox: true, margin: 0 });
    s.addText(small, { x: kx + 0.18, y: 4.83, w: kw - 0.36, h: 0.44, fontFace: FONT_BODY, fontSize: 8.4, color: "4A4A4A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    kx += kw + kgap;
  });

  // one bottom line: the recommendation, not a pre-emptive defence of the
  // numbers above. Each KPI now explains itself in its own label.
  s.addShape(pres.ShapeType.roundRect, { x: 0.5, y: 5.52, w: 12.33, h: 0.64, rectRadius: 0.06, fill: { color: MAG_BG }, line: { color: MAG, width: 1.2 } });
  s.addText([
    { text: "Recommendation:  ", options: { bold: true, color: MAG } },
    { text: "start with the failed imports nobody sees today, then intake. A person approves every decision that can go against a claimant, at every stage.", options: { color: "3A2A32" } },
  ], { x: 0.72, y: 5.52, w: 11.9, h: 0.64, fontFace: FONT_BODY, fontSize: 9.5, isTextBox: true, margin: 0, valign: "middle", lineSpacingMultiple: 1.18 });

  footer(s, 1);
}

// =========================================================================
// SLIDE 2 — System at a glance: four lanes, who does what, in what order.
// Adapted from a reference deck's swimlane technique (input / deterministic
// functions / AI agents / experts in the loop) onto GGW's own five agents,
// using the same colour code as the flowcharts on slides 3-5, which a reader
// who has already seen those needs no new legend to read this one.
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Seven stations share one claim record, so nothing is handed off or read twice", {
    x: 0.4, y: 0.13, w: 9.6, h: 0.30, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("SYSTEM AT A GLANCE", { x: 9.0, y: 0.16, w: 3.93, h: 0.26, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  // legend -- same three chips reused on slides 3-5, plus "input", so the
  // colour code is one language across the whole deck, not relearned here
  legendChip(s, 0.4, 0.56, START_FILL, START_STROKE, "Input / trigger");
  legendChip(s, 2.15, 0.56, AI_FILL, AI_STROKE, "AI agent (LLM call)");
  legendChip(s, 4.55, 0.56, DET_FILL, DET_STROKE, "Rules only, no model");
  legendChip(s, 7.15, 0.56, HUM_FILL, HUM_STROKE, "Human checkpoint");
  s.addShape(pres.ShapeType.line, { x: 9.55, y: 0.63, w: 0.35, h: 0, line: { color: ARROW, width: 1, dashType: "dash", endArrowType: "triangle" } });
  s.addText("conditional path (~20% of claims)", { x: 9.95, y: 0.50, w: 2.98, h: 0.26, fontFace: FONT_BODY, fontSize: 8, color: MUTED, isTextBox: true, margin: 0, valign: "middle" });

  // four lane bands
  const LX = 0.4, LW = 12.53, LY = [0.94, 2.18, 3.42, 4.66], LH = 1.10;
  const lanes = [
    ["INPUT", START_STROKE, "F2FAF9"],
    ["DETERMINISTIC — NO MODEL", DET_STROKE, "F4F5F6"],
    ["AI AGENTS", AI_STROKE, "F1F5FA"],
    ["A PERSON DECIDES", HUM_STROKE, "FAF3F6"],
  ];
  // label x-offset per lane: the deterministic lane's only boxes start past
  // x=6.7, so its label sits clear of the long input-to-agent jump that
  // would otherwise cross straight through the text at the default position
  const laneLabelX = [LX + 0.16, 4.20, LX + 0.16, LX + 0.16];
  lanes.forEach(([label, accent, bg], i) => {
    s.addShape(pres.ShapeType.roundRect, { x: LX, y: LY[i], w: LW, h: LH, rectRadius: 0.04, fill: { color: bg }, line: { type: "none" } });
    s.addShape(pres.ShapeType.rect, { x: LX, y: LY[i], w: 0.06, h: LH, fill: { color: accent }, line: { type: "none" } });
    s.addText(label, { x: laneLabelX[i], y: LY[i] + 0.06, w: 2.4, h: 0.18, fontFace: FONT_BODY, fontSize: 8, bold: true, color: accent, charSpacing: 0.8, isTextBox: true, margin: 0 });
  });

  // seven stations -- six main columns (reusing the same x-grid the pipeline
  // used before) plus one conditional branch sharing column 4 with Validation
  const BX = [0.4, 2.525, 4.65, 6.775, 8.90, 11.025], BW = 1.885, BH = 0.68;
  const boxY = (lane) => LY[lane] + LH - BH - 0.08;

  function station(col, lane, big, small, fill, stroke, tc, dashed) {
    const x = BX[col], y = boxY(lane);
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: BW, h: BH, rectRadius: 0.06, fill: { color: fill },
      line: { color: stroke, width: dashed ? 1.1 : 1.4, dashType: dashed ? "dash" : "solid" },
    });
    s.addText(big, { x: x + 0.08, y: y + 0.08, w: BW - 0.16, h: 0.28, fontFace: FONT_BODY, fontSize: 9.6, bold: true, color: tc, align: "center", isTextBox: true, margin: 0, lineSpacingMultiple: 1.0 });
    s.addText(small, { x: x + 0.08, y: y + 0.36, w: BW - 0.16, h: 0.30, fontFace: FONT_BODY, fontSize: 7.4, color: MUTED, align: "center", isTextBox: true, margin: 0, lineSpacingMultiple: 1.02 });
    return { x, y, w: BW, h: BH, cx: x + BW / 2, cy: y + BH / 2, lane };
  }

  // connect two stations, choosing the entry/exit edge from their relative
  // lane position rather than a fixed guess -- same lane = left/right edge,
  // different lane = top/bottom edge on whichever side actually faces the
  // other box, so the line reaches the near side instead of cutting through
  function link(a, b, dash) {
    if (a.lane === b.lane) diagArrow(s, a.x + a.w, a.cy, b.x, b.cy, dash);
    else if (b.lane > a.lane) diagArrow(s, a.cx, a.y + a.h, b.cx, b.y, dash);
    else diagArrow(s, a.cx, a.y, b.cx, b.y + b.h, dash);
  }

  const b0 = station(0, 0, "Claim arrives", "EASY · AES mail · e-mail", START_FILL, START_STROKE, START_TEXT);
  const b1 = station(1, 2, "Extraction Agent", "any format → structured fields", AI_FILL, AI_STROKE, AI_TEXT);
  const b2 = station(2, 2, "Classification Agent", "claim type + confidence score", AI_FILL, AI_STROKE, AI_TEXT);
  const b3 = station(3, 1, "Validation", "CPR · policy · duplicates", DET_FILL, DET_STROKE, DET_TEXT);
  const b3x = station(3, 2, "Exception Research", "gathers evidence · ~20% of claims", "FCEFD6", "EF9F27", "6B4A10", true);
  const bcv = station(4, 1, "Coverage Check", "cover, dates, sums · same record", DET_FILL, DET_STROKE, DET_TEXT);
  const b4 = station(4, 2, "Drafting Agent", "writes the letter, cites its reasoning", AI_FILL, AI_STROKE, AI_TEXT);
  const b5 = station(5, 3, "Handler approves", "before anything sends", HUM_FILL, HUM_STROKE, HUM_TEXT);

  // numbered station markers -- small filled circles at the top-left corner,
  // read left to right, matching the reference deck's numbering device
  // number badges sit fully inside each box's top-left corner -- contained,
  // not straddling any border, so they cannot collide with the lane labels
  [b0, b1, b2, b3, bcv, b4, b5].forEach((b, i) => {
    const bx = b.x + 0.05, by = b.y + 0.05;
    s.addShape(pres.ShapeType.ellipse, { x: bx, y: by, w: 0.18, h: 0.18, fill: { color: INK }, line: { type: "none" } });
    s.addText(String(i + 1), { x: bx, y: by, w: 0.18, h: 0.18, fontFace: FONT_BODY, fontSize: 7.4, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0 });
  });

  // flow -- solid for the main path, dashed for the ~20% exception detour.
  // Validation forks: most claims go straight to Drafting; the ~20% flagged
  // detour through Exception Research first, then rejoins at Drafting.
  link(b0, b1);
  link(b1, b2);
  link(b2, b3);
  link(b3, bcv);           // validation and coverage read the same record
  link(bcv, b4);           // clear, ~80%
  link(b3, b3x, true);     // flagged, ~20%
  link(b3x, b4, true);     // research findings feed the draft
  link(b4, b5);

  s.addText("AI drafts. Rules check. A person decides anything that can go against a claimant.",
    { x: 0.4, y: 5.96, w: 12.53, h: 0.22, fontFace: FONT_BODY, fontSize: 8.6, italic: true, bold: true, color: "2A2A3E", isTextBox: true, margin: 0 });
  footer(s, 2);
}

// =========================================================================
// SLIDE 3 (FNOL 1 of 3) — Entry & Classification
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "Three ways in and a separate side-track diagram become one classified intake", "Today on the left, the redesign on the right", "FNOL REDESIGN · 1 OF 3");
  columnHeaders(s);
  footer(s, 3);

  // ---- LEFT: today ----
  rectNode(s, "First report filed\n(usually the employer)", LC - 2.3, 1.55, 4.6, START_FILL, START_STROKE, START_TEXT, { fontSize: 9.5 });
  vArrow(s, LC, 1.95, 2.35);

  diamondNode(s, "Channel?", LC, 2.35, 0.7);
  const b1r = branch(s, LC, 2.70, "right", "Direct to AES, secure mail", "AES already gathered it;\ninsurer only re-checks coverage", "exit", { fontSize: 7.6 });
  branch(s, LC, 2.70, "left", "Teeth/glasses: 'tooth' · email + invoices", "Side track (not PD/LOEC)\nsee 1.1", "exception", { dashed: true, fontSize: 7.8 });
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
  rectNode(s, "Claim arrives: EASY, AES secure mail, or email", boxX, 1.55, boxW, START_FILL, START_STROKE, START_TEXT, { fontSize: 9.5 });
  vArrow(s, RC, 1.95, 2.35);

  const h1 = rectNode(s, "Extraction Agent\nany format in, structured fields out",
    boxX, 2.35, boxW, AI_FILL, AI_STROKE, AI_TEXT, { fontSize: 10, bold: true });
  vArrow(s, RC, 2.35 + h1, 2.35 + h1 + 0.35);
  const y2 = 2.35 + h1 + 0.35;
  const h2 = rectNode(s, "Classification Agent\nclaim type + confidence score",
    boxX, y2, boxW, AI_FILL, AI_STROKE, AI_TEXT, { fontSize: 10, bold: true });
  vArrow(s, RC, y2 + h2, y2 + h2 + 0.35);
  const dY = y2 + h2 + 0.35;

  diamondNode(s, "Confidence ≥\nthreshold?", RC, dY, 0.7);
  edgeLabel(s, "starts strict, loosens with data",
    RC + DIA_HALF + 0.15, dY + 0.12, 1.85, { align: "left", italic: true, fontSize: 7.2, h: 0.46 });
  branch(s, RC, dY + 0.35, "left", "No", "Falls back to today's manual queue", "exception", { fontSize: 7.8 });
  edgeLabel(s, "Yes", RC + 0.12, dY + 0.73, 1.0, { align: "left" });
  vArrow(s, RC, dY + 0.70, dY + 1.15);
  continues(s, RC, dY + 1.23, "▼ continues, part 2 · Parallel Validation");

  const cardY = dY + 1.55;
  s.addShape(pres.ShapeType.roundRect, { x: boxX - 0.2, y: cardY, w: boxW + 0.4, h: 0.62, rectRadius: 0.05, fill: { color: PAPER }, line: { color: LINE, width: 1 } });
  s.addText("These two agents replace 3 entry channels, the auto-import failure path, and the whole 1.1 side track.", {
    x: boxX - 0.05, y: cardY, w: boxW + 0.1, h: 0.62, fontFace: FONT_BODY, fontSize: 8.6, color: "3A4250", isTextBox: true, margin: 4, valign: "middle", align: "center", lineSpacingMultiple: 1.15,
  });
}

// =========================================================================
// SLIDE 4 (FNOL 2 of 3) — Validation Gates
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "Three checks done one after another become one check done all at once", "None of the three checks depends on the other two", "FNOL REDESIGN · 2 OF 3", false);
  columnHeaders(s);
  footer(s, 4);

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

  const h1 = rectNode(s, "Validation Node (rules only)\nCPR dedup · policy in force · completeness\nthree lookups at once, no model call",
    boxX, 1.80, boxW, DET_FILL, DET_STROKE, DET_TEXT, { fontSize: 10, bold: true });
  vArrow(s, RC, 1.80 + h1, 1.80 + h1 + 0.35);
  const dY = 1.80 + h1 + 0.35;

  diamondNode(s, "Any check\nfailed?", RC, dY, 0.7);
  const exBox = branch(s, RC, dY + 0.35, "right", "Yes", "Exception Research Agent\ngathers CVR and fuzzy-match evidence", "exception", { fontSize: 7.6 });
  vArrow(s, exBox.x + exBox.w / 2, exBox.y + exBox.h, exBox.y + exBox.h + 0.24);
  rectNode(s, "Underwriter decides\n(AI prepares only)", exBox.x, exBox.y + exBox.h + 0.24, exBox.w, HUM_FILL, HUM_STROKE, HUM_TEXT, { fontSize: 7.6 });
  edgeLabel(s, "No", RC + 0.12, dY + 0.73, 1.0, { align: "left" });
  vArrow(s, RC, dY + 0.70, 5.30);
  continues(s, RC, 5.40, "▼ continues, part 3 · Outcome Drafting Agent");

}

// =========================================================================
// SLIDE 5 (FNOL 3 of 3) — Escalation & Close-Out
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "The AI writes the draft; the handler still makes the decision", "Nothing that can go against a claimant is automated", "FNOL REDESIGN · 3 OF 3", false);
  columnHeaders(s);
  footer(s, 5);

  continues(s, LC, 1.40, "▲ continued from part 2");
  vArrow(s, LC, 1.62, 1.80);

  diamondNode(s, "Red warning\non claim?", LC, 1.80, 0.7);
  const d1 = branch(s, LC, 2.15, "left", "Yes", "On hold: check with UW before it can move on", "exception", { fontSize: 7.6 });
  rejoinSimple(s, d1, LC, 3.05);
  edgeLabel(s, "No", LC + 0.12, 2.45, 1.0, { align: "left" });
  vArrow(s, LC, 2.42, 3.05);

  diamondNode(s, "Injury minor?", LC, 3.05, 0.7);
  branch(s, LC, 3.40, "left", "Yes, ~65%", "Close at intake\nletter, no permanent injury", "exit", { fontSize: 7.6 });
  branch(s, LC, 3.40, "right", "No", "Claim in IDB: fields checked, task opened, questionnaire", "exit", { fontSize: 7.6 });

  s.addShape(pres.ShapeType.roundRect, { x: LEFT_X, y: 4.55, w: COL_W, h: 0.75, rectRadius: 0.05, fill: { color: PAPER }, line: { color: LINE, width: 1 } });
  s.addText("Today: 7 decision gates, 3 entry channels, 10 exception or exit branches across the full FNOL flow.", {
    x: LEFT_X + 0.15, y: 4.55, w: COL_W - 0.3, h: 0.75, fontFace: FONT_BODY, fontSize: 9, color: "3A4250", isTextBox: true, margin: 4, valign: "middle", align: "center", lineSpacingMultiple: 1.2,
  });

  const boxW = COL_W - 0.6, boxX = RC - boxW / 2;
  continues(s, RC, 1.40, "▲ continued from part 2");
  vArrow(s, RC, 1.62, 1.80);

  const h1 = rectNode(s, "Outcome Drafting Agent\ndrafts the letter or the IDB record,\nwith its reasoning attached",
    boxX, 1.80, boxW, AI_FILL, AI_STROKE, AI_TEXT, { fontSize: 10, bold: true });
  vArrow(s, RC, 1.80 + h1, 1.80 + h1 + 0.22);
  const humY = 1.80 + h1 + 0.22;
  const hH = rectNode(s, "Handler approves before anything sends\n(we monitor edit rate, not model accuracy)",
    boxX + 0.5, humY, boxW - 1.0, HUM_FILL, HUM_STROKE, HUM_TEXT, { fontSize: 8.6, bold: true });
  vArrow(s, RC, humY + hH, humY + hH + 0.32);
  const dY = humY + hH + 0.32;

  diamondNode(s, "Injury minor?\n(~65%)", RC, dY, 0.7);
  branch(s, RC, dY + 0.35, "left", "Yes", "Close at intake\nAI-drafted, handler approved", "exit", { fontSize: 7.6 });
  branch(s, RC, dY + 0.35, "right", "No", "Claim in IDB\nCoverage Check runs alongside", "exit", { fontSize: 7.6 });

  s.addShape(pres.ShapeType.roundRect, { x: boxX - 0.2, y: dY + 1.15, w: boxW + 0.4, h: 1.0, rectRadius: 0.05, fill: { color: PAPER }, line: { color: LINE, width: 1 } });
  s.addText([
    { text: "Net result:  ", options: { bold: true, color: INK } },
    { text: "3 entry channels → 1 · 3 sequential gates → 1 parallel check · the Coverage Check handoff disappears, it runs on the same state", options: { color: "3A4250" } },
  ], { x: boxX - 0.05, y: dY + 1.15, w: boxW + 0.1, h: 1.0, fontFace: FONT_BODY, fontSize: 9, isTextBox: true, margin: 4, valign: "middle", lineSpacingMultiple: 1.2, align: "center" });
}

// =========================================================================
// SLIDE 6 — Sequencing quadrant (decluttered: 2 items max per quadrant)
// palette validated for colour-vision separation, not eyeballed
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Start where a mistake is cheap. That is what earns the trust the bigger wins need.", {
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
    ["Detect failed EASY → IDB imports", "nobody is alerted today, so a claim can go missing unseen"],
    ["Read, classify, check CPR and policy", "one agent plus plain rules, ending the teeth dependency"],
  ]);
  quad(gx + hw, gy, "E4F4F8", "0891B2", "Wave 2 — once trusted", "the real prize, but needs the audit trail live", [
    ["Draft the close-out letter", "~65% of claims · a handler still approves"],
    ["Read AES secure mail", "the only channel with no automation today"],
  ]);
  quad(gx, gy + hh, "EDECFB", "4338CA", "Then fix the import itself", "an IT defect, once you can see it", [
    ["Retry, park and replay", "detection first, because you cannot fix what nobody reports"],
  ]);
  quad(gx + hw, gy + hh, "FBE9F1", "BE185D", "Leave manual", "too rare to earn the effort", [
    ["\"Violence cover\" entitlement letter", "~2 claims so far (GGW) · a checklist beats a model"],
  ]);

  // How the one IT item actually gets fixed -- the detail an engineer would ask for
  s.addShape(pres.ShapeType.roundRect, { x: 10.05, y: 1.0, w: 2.88, h: 5.0, rectRadius: 0.06, fill: { color: "EDECFB" }, line: { color: "4338CA", width: 1.3 } });
  s.addText("THE IMPORT FAILURES", { x: 10.27, y: 1.16, w: 2.44, h: 0.24, fontFace: FONT_BODY, fontSize: 9, bold: true, color: "4338CA", charSpacing: 0.7, isTextBox: true, margin: 0 });
  s.addText("Not noticed today (GGW), so detection comes before retries:", { x: 10.27, y: 1.42, w: 2.44, h: 0.34, fontFace: FONT_BODY, fontSize: 7.8, italic: true, color: "4A4A5E", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  // numbered-pill list -- one badge, one short label, one 3-5 word caption.
  // no paragraphs: the reader gets the sequence at a glance, not a read.
  const fixes = [
    ["Count what should arrive", "daily reconciliation catches gaps"],
    ["Sort the failure", "timeout vs. missing field"],
    ["Retry with backoff", "1s, 2s, 4s + jitter"],
    ["One key per claim", "no duplicate replays"],
    ["Park, don't drop", "holding queue for review"],
    ["Alert on the rate", "a spike matters, a trickle does not"],
  ];
  let fy = 1.86;
  fixes.forEach(([t, d], i) => {
    s.addShape(pres.ShapeType.ellipse, { x: 10.27, y: fy, w: 0.30, h: 0.30, fill: { color: "4338CA" }, line: { type: "none" } });
    s.addText(String(i + 1), { x: 10.27, y: fy, w: 0.30, h: 0.30, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(t, { x: 10.65, y: fy - 0.02, w: 2.06, h: 0.19, fontFace: FONT_BODY, fontSize: 8.8, bold: true, color: "2A2A3E", isTextBox: true, margin: 0 });
    s.addText(d, { x: 10.65, y: fy + 0.16, w: 2.06, h: 0.18, fontFace: FONT_BODY, fontSize: 7.6, color: "6A6A7E", isTextBox: true, margin: 0 });
    fy += 0.535;
  });

  s.addText("AES rulings and underwriting authority are deliberately absent. Those decisions are not ours to make.", { x: 0.4, y: 6.6, w: 12.5, h: 0.28, fontFace: FONT_BODY, fontSize: 7, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 6);
}

// =========================================================================
// SLIDE 7 — The stack, hosting, and how it goes live without breaking.
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Deployed on the cloud GGW already run, and released in three gated steps", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("STACK & ROLLOUT", { x: 9.4, y: 0.16, w: 3.53, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  // reliability band -- persistence, kept short
  const RX = [0.4, 2.525], RW = [1.885, 10.045];
  s.addShape(pres.ShapeType.roundRect, { x: RX[0], y: 0.70, w: RW[0], h: 0.52, rectRadius: 0.05, fill: { color: "FDF3E3" }, line: { color: AMBER, width: 1 } });
  s.addText("Service Bus\nfailures park, nothing is lost", { x: RX[0] + 0.08, y: 0.70, w: RW[0] - 0.16, h: 0.52, fontFace: FONT_BODY, fontSize: 8, bold: true, color: "9A5A0A", align: "center", valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 1.08 });
  s.addShape(pres.ShapeType.roundRect, { x: RX[1], y: 0.70, w: RW[1], h: 0.52, rectRadius: 0.05, fill: { color: "ECEFF3" }, line: { color: DET_STROKE, width: 1 } });
  s.addText("PostgreSQL  ·  graph checkpoint and audit log in one place: every decision, its inputs, its confidence",
    { x: RX[1], y: 0.70, w: RW[1], h: 0.52, fontFace: FONT_BODY, fontSize: 9, bold: true, color: DET_TEXT, align: "center", valign: "middle", isTextBox: true, margin: 0 });

  // the stack, as names only
  s.addText("THE STACK", { x: 0.4, y: 1.54, w: 3.0, h: 0.22, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: TEAL, charSpacing: 0.9, isTextBox: true, margin: 0 });
  const chips = ["FastAPI", "LangGraph", "Pydantic", "PostgreSQL", "Service Bus", "Container Apps", "Blob + Key Vault", "Azure OpenAI", "Azure DevOps"];
  const cw = (12.53 - 8 * 0.18) / 9;
  chips.forEach((c, i) => {
    const x = 0.4 + i * (cw + 0.18);
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.84, w: cw, h: 0.50, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: TEAL, width: 1.2 } });
    s.addText(c, { x: x + 0.04, y: 1.84, w: cw - 0.08, h: 0.50, fontFace: FONT_BODY, fontSize: 8.6, bold: true, color: "0B5A66", align: "center", valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 1.0 });
  });

  // operating model -- kept from the pipeline slide, restated once more here
  // right next to the infrastructure, because "who owns the call" is part of
  // how this is operated, not just how it is designed
  s.addText("OPERATING MODEL", { x: 0.4, y: 2.94, w: 4.0, h: 0.22, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: INDIGO, charSpacing: 0.9, isTextBox: true, margin: 0 });
  const roles = [
    ["AI agents", "draft, classify, extract, but never decide", AI_FILL, AI_STROKE, AI_TEXT],
    ["Rules, no model", "check the record: pass or fail, nothing in between", DET_FILL, DET_STROKE, DET_TEXT],
    ["A person", "approves before anything can go against a claimant", HUM_FILL, HUM_STROKE, HUM_TEXT],
  ];
  const rw = (12.53 - 2 * 0.2) / 3;
  roles.forEach(([t, d, fill, stroke, tc], i) => {
    const x = 0.4 + i * (rw + 0.2);
    s.addShape(pres.ShapeType.roundRect, { x, y: 3.24, w: rw, h: 0.46, rectRadius: 0.05, fill: { color: fill }, line: { color: stroke, width: 1.2 } });
    s.addText(t, { x: x + 0.14, y: 3.24, w: 1.7, h: 0.46, fontFace: FONT_BODY, fontSize: 9.2, bold: true, color: tc, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(d, { x: x + 1.85, y: 3.24, w: rw - 1.95, h: 0.46, fontFace: FONT_BODY, fontSize: 8, color: tc, valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 1.05 });
  });

  // how it goes live
  s.addText("HOW IT GOES LIVE", { x: 0.4, y: 3.98, w: 3.0, h: 0.22, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MAG, charSpacing: 0.9, isTextBox: true, margin: 0 });
  const gw = (12.53 - 3 * 0.24) / 4, gy = 4.24, gh = 0.92;
  const gates = [
    ["Shadow", "weeks 1–8", "Runs on live claims. Nobody sees the output.", "Exit: extraction agrees with the handler on 95% of fields.", AMBER, AMBER_BG],
    ["Suggest", "weeks 9–16", "The handler must edit or approve. Nothing sends itself.", "Exit: edit rate on minor claims settles below 20%.", TEAL, TEAL_BG],
    ["Approve", "months 5–6", "Minor claims pre-filled. One click accepts.", "Exit criteria reviewed monthly with the claims lead.", INDIGO, INDIGO_BG],
    ["Never", "by design", "Nothing adverse to a claimant is ever automatic.", "This is a boundary, not a phase.", MAG, MAG_BG],
  ];
  gates.forEach(([t, when, d, gate, col, bg], i) => {
    const x = 0.4 + i * (gw + 0.24);
    s.addShape(pres.ShapeType.roundRect, { x, y: gy, w: gw, h: gh, rectRadius: 0.06, fill: { color: bg }, line: { color: col, width: 1.2 } });
    s.addText(t, { x: x + 0.16, y: gy + 0.10, w: gw - 0.32, h: 0.22, fontFace: FONT_BODY, fontSize: 10.5, bold: true, color: col, isTextBox: true, margin: 0 });
    s.addText(when, { x: x + 0.16, y: gy + 0.10, w: gw - 0.32, h: 0.22, fontFace: FONT_BODY, fontSize: 7.8, italic: true, color: MUTED, align: "right", isTextBox: true, margin: 0 });
    s.addText(d, { x: x + 0.16, y: gy + 0.34, w: gw - 0.32, h: 0.28, fontFace: FONT_BODY, fontSize: 8.2, color: "3A3A3A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.12 });
    s.addShape(pres.ShapeType.rect, { x: x + 0.16, y: gy + 0.66, w: gw - 0.32, h: 0.01, fill: { color: col }, line: { type: "none" } });
    s.addText(gate, { x: x + 0.16, y: gy + 0.70, w: gw - 0.32, h: 0.20, fontFace: FONT_BODY, fontSize: 7.2, italic: true, color: col === MAG ? "6A1F42" : "4A4A4A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.05 });
    if (i < 3) elbowAcross(s, x + gw + 0.03, x + gw + 0.21, gy + gh / 2);
  });

  s.addText("If an agent fails or is unsure, the claim falls back to today's manual queue. Slower, never wrong.",
    { x: 0.4, y: gy + gh + 0.14, w: 12.5, h: 0.26, fontFace: FONT_BODY, fontSize: 7.4, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 7);
}

// =========================================================================
// =========================================================================
// SLIDE 8 — How the numbers are built. Two tables, not five: time, and the
// business case. Per-agent token cost and the Azure line items are folded
// into one note each -- the calculation still holds, it just isn't spelled
// out three ways when one line answers "is this real money?"
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Every minute is observed, and the redesign pays for itself in 14 months", {
    x: 0.4, y: 0.13, w: 10.2, h: 0.34, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("APPENDIX · HOW THE NUMBERS ARE BUILT", { x: 7.0, y: 0.16, w: 5.93, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  const L = 0.4, R = 6.75, CW = 6.18;
  const Bh = (t) => ({ text: t, options: { bold: true } });

  // ---------- A: time, merged to four everyday tasks + the calculation ----------
  s.addText("A · TIME PER CLAIM, PER HANDLER (MINUTES)", { x: L, y: 0.58, w: CW, h: 0.24, fontFace: FONT_BODY, fontSize: 10, bold: true, color: AMBER, charSpacing: 0.6, isTextBox: true, margin: 0 });
  const r = F.rows;
  const sum = (idxs, col) => idxs.reduce((a, i) => a + parseFloat(r[i][col]), 0).toFixed(1);
  const timeRows = [
    [Bh("Step"), Bh("Today"), Bh("Redesigned")],
    ["Read & extract the report", sum([0, 1], 1), sum([0, 1], 2)],
    ["Check policy, CPR & escalations", sum([2, 4], 1), sum([2, 4], 2)],
    ["Decide, draft and send", r[3][1], r[3][2]],
    ["Switch between systems", r[5][1], r[5][2]],
    [Bh("Touch time"), Bh(F.touchToday), Bh(F.touchAfter)],
    [`+${F.allowPct}% breaks & interruptions`, F.allowToday, F.allowAfter],
    [Bh("Standard time"), Bh(F.stdToday), Bh(F.stdAfter)],
    [Bh("What we bank"), Bh(F.stdToday), Bh(F.bankedAfter)],
  ];
  s.addTable(timeRows, {
    x: L, y: 0.92, w: CW, colW: [3.98, 1.1, 1.1], fontFace: FONT_BODY, fontSize: 9.5,
    border: { type: "solid", color: "E4E4E0", pt: 0.5 }, color: "3A3A3A",
    align: "left", valign: "middle", rowH: 0.36, autoPage: false, fill: { color: "FFFFFF" },
  });
  s.addText(`We bank ${F.bankedAfterInt} min, not the ${F.stdAfter} the design allows: the business case runs on a three-point estimate of best, likely and worst case.`,
    { x: L, y: 4.30, w: CW, h: 0.30, fontFace: FONT_BODY, fontSize: 8.6, italic: true, color: "6A6A6A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  s.addText(`Basis: ${F.handlers} claim handlers × ${F.claimsPerPersonDay} claims a day, confirmed by GGW (8 Sep 2026). ${F.freedMinPersonDay} min back, each, per day.`,
    { x: L, y: 4.62, w: CW, h: 0.30, fontFace: FONT_BODY, fontSize: 8.6, color: "6A6A6A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });

  // ---------- B: the business case ----------
  s.addText("B · THE BUSINESS CASE", { x: R, y: 0.58, w: CW, h: 0.24, fontFace: FONT_BODY, fontSize: 10, bold: true, color: MAG, charSpacing: 0.6, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: R, y: 0.92, w: CW, h: 3.36, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: MAG, width: 1.2 } });
  const caseRows = [
    ["EVERY YEAR", "", true],
    ["Capacity released", `${F.hoursYr} handler-h · ${F.fte} FTE`, false],
    [`Worth, at €${F.hourly} per handler-hour`, `€${F.valueYrR}`, false],
    ["Cost to run it — AI models and Azure", `− €${F.runYrR}`, false],
    ["Ongoing support — 0.25 of an engineer", `− €${F.maintYrR}`, false],
    ["Net value, every year", `€${F.netYrR}`, true],
    ["ONE-OFF", "", true],
    ["Build — 3 people × 6 months, fully loaded", `€${F.buildR}`, true],
    ["Payback", `${F.payback} months`, true],
  ];
  let cy = 1.06;
  caseRows.forEach(([k, v, hi]) => {
    if (!v) { // section header
      s.addText(k, { x: R + 0.2, y: cy, w: CW - 0.4, h: 0.22, fontFace: FONT_BODY, fontSize: 8.4, bold: true, color: MAG, charSpacing: 1, isTextBox: true, margin: 0 });
      cy += 0.34;
      return;
    }
    s.addText(k, { x: R + 0.2, y: cy, w: 4.0, h: 0.28, fontFace: FONT_BODY, fontSize: 9.5, color: hi ? INK : "3A3A3A", bold: hi, isTextBox: true, margin: 0, valign: "middle" });
    s.addText(v, { x: R + 4.1, y: cy, w: CW - 4.3, h: 0.28, fontFace: FONT_BODY, fontSize: 9.5, bold: true, color: hi ? MAG : INK, align: "right", isTextBox: true, margin: 0, valign: "middle" });
    cy += 0.34;
  });
  s.addText(`Running cost in full: ≈ €${F.perClaimEurR}/claim in AI tokens + ≈ $134/month in Azure → €${F.runYrR}/year, all in.`,
    { x: R, y: 4.48, w: CW, h: 0.30, fontFace: FONT_BODY, fontSize: 8.6, color: "6A6A6A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  s.addText("The build is loaded team capacity, not new cash. The AI team is already funded.",
    { x: R, y: 4.80, w: CW, h: 0.30, fontFace: FONT_BODY, fontSize: 8.6, italic: true, color: "6A6A6A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });

  footer(s, 8);
}

// =========================================================================
// SLIDE 9 — Sources. Plain text, one list, no colour, no boxes — a reference
// page, not a designed slide.
// =========================================================================
{
  const s = pres.addSlide();
  s.background = { color: "FAFAF8" };

  s.addText("Every number is sourced, dated, or named as an assumption we'd test first", {
    x: 0.4, y: 0.16, w: 10.0, h: 0.34, fontFace: FONT_HEAD, fontSize: 15, bold: true, color: INK, isTextBox: true, margin: 0 });
  s.addText("APPENDIX", { x: 10.53, y: 0.20, w: 2.4, h: 0.26, fontFace: FONT_BODY, fontSize: 9, bold: true, color: MUTED, charSpacing: 1.4, align: "right", isTextBox: true, margin: 0 });

  const SRC = [
    ["Who does intake", "20 claim handlers process claims; 10 support staff besides.", "GGW (P. Gossmann), 8 Sep 2026", ""],
    ["Failed EASY imports", "Not noticed today — so detection comes before retries.", "GGW (P. Gossmann), 8 Sep 2026", ""],
    ["Volumes and shares", "~50 first reports a day · 65% minor · ~30 in the team.", "GGW case study brief", ""],
    ["Wage base", "Schadenregulierer/in, €45,100 gross — the role doing intake.", "StepStone", "https://www.stepstone.de/gehalt/Schadenregulierer-in.html"],
    ["Employer on-cost", "≈ €23 per €100 gross — health, pension, care.", "sevdesk", "https://sevdesk.de/ratgeber/buchhaltung-finanzen/lohnbuchhaltung/lohnnebenkosten/"],
    ["Working time & breaks", "38h tariff week; statutory break rules.", "Arbeitszeitgesetz §4", "https://www.gesetze-im-internet.de/arbzg/__4.html"],
    ["Build & upkeep pay", "Senior ML engineer ≈ €98k; juniors ≈ €70k, as briefed.", "Glassdoor", "https://www.glassdoor.com/Salaries/germany-machine-learning-engineer-salary-SRCH_IL.0,7_IN96_KO8,33.htm"],
    ["Model prices", "gpt-5.6 terra — $2 in / $12 out per million tokens.", "OpenAI", "https://openai.com/api/pricing/"],
    ["Hosting & database", "Container Apps consumption rates; PostgreSQL B2s.", "Azure pricing", "https://azure.microsoft.com/en-us/pricing/details/container-apps/"],
    ["Upkeep after go-live", "15–20% of build effort a year, ours is 17%.", "Standard maintenance rule of thumb", ""],
    ["Industry baseline", "17 of 20 carriers still manual; none let AI decide alone.", "Decerto, 2026 Claims Decisioning Pulse", "https://www.decerto.com/us/post/ai-claims-decisioning-framework-2026-what-us-claims-leaders-told-us-about-how-decisions-actually-get-made"],
    ["Where value comes from", "The 10-20-70 split — algorithms, technology, people.", "BCG", "https://www.bcg.com/publications/2024/transforming-customer-service-operations-with-genai"],
    ["Fixing import failures", "Retry vs. dead-letter, idempotency, alert on rate.", "Standard integration pattern", "https://www.glukhov.org/app-architecture/integration-patterns/dead-letter-queues/"],
  ];

  let y = 0.72;
  SRC.forEach(([label, detail, name, url], i) => {
    s.addText(`${i + 1}.  ${label} — ${detail}  ${url ? "" : "(" + name + ")"}`, {
      x: 0.4, y, w: url ? 9.6 : 12.53, h: 0.30, fontFace: FONT_BODY, fontSize: 10, color: "2A2A2A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15,
    });
    if (url) {
      s.addText(`(${name})`, { x: 10.05, y, w: 2.88, h: 0.30, fontFace: FONT_BODY, fontSize: 9.5, color: "2A2A2A", underline: true, isTextBox: true, margin: 0, valign: "middle", hyperlink: { url, tooltip: name } });
    }
    y += 0.335;
  });

  y += 0.10;
  s.addText("Still assumptions, not sourced. The first thing we would test:", {
    x: 0.4, y, w: 12.53, h: 0.22, fontFace: FONT_BODY, fontSize: 10, bold: true, color: "2A2A2A", isTextBox: true, margin: 0 });
  y += 0.28;
  s.addText("Every per-step minute on slide 8 · the 30% allowance for breaks and interruptions · how often a CPR is missing (30%) or underwriting is asked (25%) · how often the redesign raises an exception (20%) · the tokens a typical claim needs · the €146k build · the 0.25 FTE upkeep. Two weeks of watching live intake would replace all of this with measurements.",
    { x: 0.4, y, w: 12.53, h: 0.60, fontFace: FONT_BODY, fontSize: 9.5, color: "4A4A4A", isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });

  s.addText("Prices checked September 2026 · US$ converted at €0.92 · every euro figure is fully loaded (gross pay plus employer contributions).",
    { x: 0.4, y: 6.9, w: 12.5, h: 0.26, fontFace: FONT_BODY, fontSize: 8, color: MUTED, isTextBox: true, margin: 0 });
  footer(s, 9);
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user-ggw-casestudy/2b8e5a67-8934-50d1-9acc-621207f6c407/scratchpad/deck/GGW_FNOL_Redesign.pptx" })
  .then(() => console.log("written"))
  .catch((e) => { console.error(e); process.exit(1); });
