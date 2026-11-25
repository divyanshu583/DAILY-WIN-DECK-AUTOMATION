const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/128eaa1c-3e57-4a55-afe9-9d5a28ffc7a8_a39d82ad-d47c-43ce-8e2e-7c30c9852d84/pptx/scripts/apply_theme.js");

const OUT = process.argv[2];
const THEME = {
  name: "TMF WIN",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "142D43", lt1: "FFFFFF", dk2: "526575", lt2: "F0F5F6",
    accent1: "087E83", accent2: "CE1734", accent3: "A9D7DB", accent4: "D5DFE3",
    accent5: "0B5E63", accent6: "1E4566", hlink: "087E83", folHlink: "526575",
  },
};
const NAVY = "142D43", TEAL = "087E83", RED = "CE1734", ICE = "F0F5F6", MIST = "D5DFE3",
  SLATE = "526575", AQUA = "A9D7DB", WHITE = "FFFFFF", NAVY2 = "1E4566";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: "Arial", bodyFontFace: "Calibri" };
pres.title = "Vishal Nair - TMF Group - WIN Deck";
pres.author = "Vishal Nair";

const H = "Arial", B = "Calibri";
const FOOT = "VISHAL NAIR  /  TMF GROUP  /  WIN • CANDIDATE PROPOSAL";

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: WHITE },
  objects: [
    { text: { text: FOOT, options: { x: 0.5, y: 7.02, w: 8, h: 0.3, fontFace: B, fontSize: 11, color: SLATE, charSpacing: 1, margin: 0, valign: "middle" } } },
  ],
  slideNumber: { x: 12.33, y: 7.02, w: 0.5, h: 0.3, fontFace: B, fontSize: 11, color: SLATE, align: "right", margin: 0 },
});
pres.defineSlideMaster({ title: "DARK", background: { color: NAVY }, objects: [] });

function tb(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontFace: B, fontSize: 14, color: NAVY, valign: "top" }, o));
}
function header(s, kicker, title, sub) {
  tb(s, kicker.toUpperCase(), { x: 0.5, y: 0.42, w: 12.3, h: 0.28, fontSize: 12, bold: true, color: TEAL, charSpacing: 2 });
  tb(s, title, { x: 0.5, y: 0.72, w: 12.3, h: 0.6, fontFace: H, fontSize: 28, bold: true, valign: "middle" });
  if (sub) tb(s, sub, { x: 0.5, y: 1.36, w: 12.3, h: 0.4, fontSize: 18, color: SLATE, valign: "middle" });
}
function sources(s, text) {
  tb(s, "Sources: " + text, { x: 0.5, y: 6.68, w: 12.3, h: 0.28, fontSize: 11, color: SLATE, italic: true, valign: "middle" });
}
function circleNum(s, n, x, y, d, fill, color) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  tb(s, String(n), { x, y, w: d, h: d, align: "center", valign: "middle", fontFace: H, fontSize: 13, bold: true, color });
}
// labelled block: label + paragraph(s)
function block(s, label, text, x, y, w, h, opts = {}) {
  tb(s, label.toUpperCase(), { x, y, w, h: 0.26, fontSize: 12, bold: true, color: opts.labelColor || TEAL, charSpacing: 1 });
  tb(s, text, { x, y: y + 0.3, w, h: h - 0.3, fontSize: 14, color: opts.color || NAVY, paraSpaceAfter: 3 });
}
function challengePanel(s, x, y, w, h, d) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: ICE }, line: { color: ICE }, rectRadius: 0.08 });
  const ix = x + 0.3, iw = w - 0.6;
  tb(s, "THE CHALLENGE", { x: ix, y: y + 0.25, w: iw, h: 0.3, fontFace: H, fontSize: 14, bold: true, color: RED, charSpacing: 2 });
  let cy = y + 0.7;
  for (const [lab, txt, hh] of d) { block(s, lab, txt, ix, cy, iw, hh); cy += hh + 0.12; }
}
function steps(s, items, x, y, w, rowH) {
  items.forEach((t, i) => {
    circleNum(s, i + 1, x, y + i * rowH, 0.36, TEAL, WHITE);
    tb(s, t, { x: x + 0.52, y: y + i * rowH, w: w - 0.52, h: rowH - 0.05, fontSize: 14, valign: "top" , lineSpacingMultiple: 0.95});
  });
}
function carPanel(s, x, y, w, h, stats, lines) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08 });
  const sw = stats.length > 1 ? 1.75 : 1.45;
  stats.forEach((st, i) => {
    if (stats.length > 1) {
      const yy = y + 0.2 + i * 0.62;
      tb(s, st[0], { x: x + 0.25, y: yy, w: 0.85, h: 0.5, fontFace: H, fontSize: 24, bold: true, color: WHITE, valign: "middle" });
      tb(s, st[1], { x: x + 1.1, y: yy, w: sw - 0.85, h: 0.5, fontSize: 11, color: AQUA, valign: "middle" });
    } else {
      tb(s, st[0], { x: x + 0.25, y: y + 0.18, w: sw, h: 0.6, fontFace: H, fontSize: 32, bold: true, color: WHITE });
      tb(s, st[1], { x: x + 0.25, y: y + 0.78, w: sw, h: h - 0.95, fontSize: 11, color: AQUA });
    }
  });
  const tx = x + 0.25 + sw + 0.15;
  tb(s, "MY TRACK RECORD • CAR", { x: tx, y: y + 0.18, w: x + w - tx - 0.25, h: 0.26, fontSize: 12, bold: true, color: AQUA, charSpacing: 1 });
  tb(s, lines.map((l, i) => ({ text: l[0], options: { bold: true, color: AQUA, breakLine: false } })).flatMap((r, i) => [r, { text: lines[i][1], options: { color: WHITE, breakLine: i < lines.length - 1 } }]),
    { x: tx, y: y + 0.46, w: x + w - tx - 0.2, h: h - 0.55, fontSize: 14, paraSpaceAfter: 0 });
}

// ---------- Slide 1: Title ----------
{
  const s = pres.addSlide({ masterName: "DARK" });
  // motif: concentric rings, right side
  const cx = 10.6, cy = 3.75;
  [[4.6, 85], [3.5, 75], [2.4, 60], [1.3, 40]].forEach(([d, t]) => {
    s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: TEAL, transparency: t }, line: { color: AQUA, width: 0.75, transparency: 50 } });
  });
  [[8.75, 2.2], [12.3, 2.55], [9.1, 5.55], [12.0, 5.2], [10.6, 1.45]].forEach(([x, y]) => {
    s.addShape(pres.shapes.LINE, { x: Math.min(x, cx), y: Math.min(y, cy), w: Math.abs(cx - x), h: Math.abs(cy - y), flipH: (x > cx) !== (y > cy), line: { color: AQUA, width: 0.75, transparency: 55 } });
    s.addShape(pres.shapes.OVAL, { x: x - 0.09, y: y - 0.09, w: 0.18, h: 0.18, fill: { color: AQUA }, line: { color: AQUA } });
  });
  s.addShape(pres.shapes.OVAL, { x: cx - 0.2, y: cy - 0.2, w: 0.4, h: 0.4, fill: { color: RED }, line: { color: RED } });
  // logo zone reserved (empty) at 0.6,0.5 2.2x0.8
  tb(s, "WELL-RESEARCHED INSIGHTS NARRATIVE", { x: 0.6, y: 1.75, w: 7.2, h: 0.3, fontSize: 13, bold: true, color: AQUA, charSpacing: 3 });
  tb(s, "Enabling TMF Group’s digital transformation through quality excellence", { x: 0.6, y: 2.15, w: 7.4, h: 1.9, fontFace: H, fontSize: 38, bold: true, color: WHITE, valign: "top", lineSpacingMultiple: 0.95 });
  const pills = ["Business continuity", "Data integrity", "Global release confidence"];
  let px = 0.6;
  pills.forEach((p) => {
    const w = 0.35 + p.length * 0.105;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: 4.3, w, h: 0.42, fill: { color: NAVY2 }, line: { color: AQUA, width: 0.75 }, rectRadius: 0.21 });
    tb(s, p, { x: px, y: 4.3, w, h: 0.42, align: "center", valign: "middle", fontSize: 14, color: WHITE });
    px += w + 0.15;
  });
  tb(s, [
    { text: "Vishal Nair", options: { fontFace: H, fontSize: 22, bold: true, color: WHITE, breakLine: true } },
    { text: "Test Manager – Digital Client Platform", options: { fontSize: 18, color: AQUA, breakLine: true } },
    { text: "vshal17@gmail.com  •  Pune  •  November 2026", options: { fontSize: 14, color: MIST } },
  ], { x: 0.6, y: 5.45, w: 7.2, h: 1.3, paraSpaceAfter: 4 });
  s.addNotes("Independent candidate proposal. Official TMF Group logo not available — reserved space top-left; insert official logo before sending.");
}

// ---------- Slide 2: Company & alignment ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  header(s, "Company & alignment", "Quality is central to TMF’s client promise", "Accounting & tax  •  Payroll & HR  •  Global entity management  •  Fund services");
  const rows = [
    ["Purpose / mission", "Simplify global operations and be a trusted partner for clients.", "Structured governance and safer banking releases."],
    ["Vision", "Combine technology and local expertise to support global operations.", "Integration testing, SQL validation and cross-team delivery."],
    ["Strategic goal", "Make international investment and operations easier and safer.", "RTM, audit readiness and 18% lower defect leakage."],
  ];
  tb(s, "TMF GROUP", { x: 0.5, y: 2.0, w: 6, h: 0.26, fontSize: 12, bold: true, color: SLATE, charSpacing: 2 });
  tb(s, "MY ALIGNMENT", { x: 7.35, y: 2.0, w: 5, h: 0.26, fontSize: 12, bold: true, color: TEAL, charSpacing: 2 });
  rows.forEach((r, i) => {
    const y = 2.35 + i * 0.86;
    tb(s, r[0], { x: 0.5, y, w: 1.95, h: 0.72, fontFace: H, fontSize: 14, bold: true, valign: "middle" });
    tb(s, r[1], { x: 2.5, y, w: 4.2, h: 0.72, fontSize: 14, color: NAVY, valign: "middle" });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 6.8, y: y + 0.24, w: 0.38, h: 0.26, fill: { color: TEAL }, line: { color: TEAL } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.35, y, w: 5.48, h: 0.72, fill: { color: ICE }, line: { color: ICE }, rectRadius: 0.06 });
    tb(s, r[2], { x: 7.6, y, w: 5.1, h: 0.72, fontSize: 14, color: NAVY, valign: "middle" });
    if (i < 2) s.addShape(pres.shapes.LINE, { x: 0.5, y: y + 0.79, w: 6.2, h: 0, line: { color: MIST, width: 0.75 } });
  });
  tb(s, [
    { text: "Three quality priorities, one execution roadmap", options: { bold: true, fontFace: H, color: NAVY } },
    { text: "  —  a business-first approach to the Digital Client Platform", options: { color: SLATE } },
  ], { x: 0.5, y: 4.98, w: 12.3, h: 0.32, fontSize: 14, valign: "middle" });
  const chips = [["01", "Business continuity", "Migration assurance"], ["02", "Integration & data integrity", "API and reconciliation controls"], ["03", "Global quality governance", "Compliance, automation and metrics"], ["04", "Execution & candidate fit", "180-day roadmap and evidence"]];
  const cw = (12.33 - 3 * 0.2) / 4;
  chips.forEach((c, i) => {
    const x = 0.5 + i * (cw + 0.2);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 5.4, w: cw, h: 1.15, fill: { color: i === 3 ? TEAL : NAVY }, line: { color: i === 3 ? TEAL : NAVY }, rectRadius: 0.08 });
    tb(s, c[0], { x: x + 0.22, y: 5.52, w: 0.7, h: 0.4, fontFace: H, fontSize: 20, bold: true, color: i === 3 ? WHITE : RED === RED ? AQUA : AQUA });
    tb(s, [{ text: c[1], options: { bold: true, fontSize: 14, color: WHITE, breakLine: true } }, { text: c[2], options: { fontSize: 12, color: i === 3 ? WHITE : AQUA } }],
      { x: x + 0.22, y: 5.9, w: cw - 0.4, h: 0.6 });
  });
  sources(s, "TMF purpose & vision • 2025 annual report • mission and strategy");
}

// ---------- Slide 3: Priority 1 (layout A) ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  header(s, "Priority 1 · Business continuity", "Modernize without interrupting client services", "Transformation risk: protect existing operations while changing platforms.");
  challengePanel(s, 0.5, 2.0, 5.5, 4.55, [
    ["Documented business signals", "€74M technology investment in 2025 (vs €59M in 2024). UK/Ireland payroll migration required 1,000+ parallel tests for 400+ clients.", 1.32],
    ["Potential business impact", "Incorrect outputs can delay payroll or reporting; late defects can postpone migration waves and erode client confidence.", 1.1],
    ["What the JD calls for", "End-to-end strategy, estimates, regression, UAT and release validation.", 0.85],
  ]);
  const x = 6.35, w = 6.48;
  tb(s, "THE SOLUTION", { x, y: 2.05, w, h: 0.3, fontFace: H, fontSize: 14, bold: true, color: TEAL, charSpacing: 2 });
  tb(s, [{ text: "Risk-based migration assurance", options: { bold: true, fontFace: H, fontSize: 16, breakLine: true } }, { text: "Protect critical journeys with evidence-led go/no-go decisions.", options: { color: SLATE } }], { x, y: 2.4, w, h: 0.62, fontSize: 14 });
  steps(s, ["Rank journeys by impact; baseline legacy outputs.", "Run parallel comparisons; resolve discrepancies.", "Pilot by wave; agree rollback and UAT sign-off; validate production."], x, 3.1, w, 0.42);
  tb(s, [{ text: "Evidence it can work: ", options: { bold: true, color: TEAL } }, { text: "TMF’s 400+ client migration used phased transition and parallel tests without compromising service (Reward Strategy, 2026). Benefit: earlier detection and safer cutovers." }], { x, y: 4.33, w, h: 0.75, fontSize: 14 });
  carPanel(s, x, 5.15, w, 1.4, [["18%", "lower leakage"]], [["C: ", "CIBC mainframe migration."], ["A: ", "Led six engineers, structured QA and checkout."], ["R: ", "12+ cycles; 18% lower leakage."]]);
  sources(s, "TMF annual report 2025 • Reward Strategy, May 2026 • TMF migration case study");
  s.addNotes("Public signals establish business exposure, not confirmed internal DCP testing deficiencies. Validate the current baseline with the hiring manager. Candidate achievements: updated TMF-specific resume. External case study outcomes are publisher-reported, not guaranteed TMF results.");
}

// ---------- Slide 4: Priority 2 (layout B) ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  header(s, "Priority 2 · Integration & data integrity", "Successful interfaces can still produce wrong data", "Integration risk: verify the business outcome, not just the API response.");
  const cards = [
    ["Documented business signals", "KRAIOS modernizes client interactions; Optix serves 3,600+ entities and 160,000 client employees (2025). JD explicitly calls out data transformations."],
    ["Potential business impact", "Incorrect mapping, duplicate records or missing data can corrupt financial outputs; manual reconciliation increases effort and delays."],
    ["What the JD calls for", "API, integration, data validation, workflows and end-to-end testing."],
  ];
  tb(s, "THE CHALLENGE", { x: 0.5, y: 1.95, w: 6, h: 0.3, fontFace: H, fontSize: 14, bold: true, color: RED, charSpacing: 2 });
  const cw = (12.33 - 2 * 0.2) / 3;
  cards.forEach((c, i) => {
    const x = 0.5 + i * (cw + 0.2);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.3, w: cw, h: 1.55, fill: { color: ICE }, line: { color: ICE }, rectRadius: 0.08 });
    block(s, c[0], c[1], x + 0.22, 2.42, cw - 0.44, 1.4);
  });
  tb(s, "THE SOLUTION", { x: 0.5, y: 4.03, w: 6, h: 0.3, fontFace: H, fontSize: 14, bold: true, color: TEAL, charSpacing: 2 });
  tb(s, [{ text: "API-first testing and automated reconciliation", options: { bold: true, fontFace: H, fontSize: 16, breakLine: true } }, { text: "Validate contracts, mappings and financial outputs at every critical interface.", options: { color: SLATE } }], { x: 0.5, y: 4.36, w: 5.9, h: 0.82, fontSize: 14 });
  steps(s, ["Test schema, permissions, retries and duplicates.", "Use SQL to compare counts, fields and balances.", "Run API/data checks in pipelines; block unexplained critical differences."], 0.5, 5.25, 5.9, 0.43);
  const x = 6.65, w = 6.18;
  tb(s, [{ text: "Evidence it can work: ", options: { bold: true, color: TEAL } }, { text: "Coca-Cola Consolidated: HR data testing 100 estimated → 31 person-days; 500+ checks in Azure DevOps. Expleo: UAE bank validation time 7,000 → 3,500 hours." }], { x, y: 4.05, w, h: 0.95, fontSize: 14 });
  carPanel(s, x, 5.15, w, 1.4, [["22%", "fewer integration defects at JPMorgan"]], [["C: ", "Cross-system payment defects."], ["A: ", "SQL, SIT/UAT, settlement testing."], ["R: ", "22% fewer integration defects at JPMorgan."]]);
  sources(s, "TMF annual report • technology risk • QuerySurge • Coca-Cola • QuerySurge • Expleo • ACM • ETL research");
  s.addNotes("Public signals establish business exposure, not confirmed internal DCP testing deficiencies. Validate the current baseline with the hiring manager. Candidate achievements: updated TMF-specific resume. External case study outcomes are publisher-reported, not guaranteed TMF results.");
}

// ---------- Slide 5: Priority 3 (layout A mirrored) ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  header(s, "Priority 3 · Global quality governance", "Global growth requires consistent quality controls", "Governance risk: balance common standards with local regulatory variation.");
  const x = 0.5, w = 6.48;
  tb(s, "THE SOLUTION", { x, y: 2.05, w, h: 0.3, fontFace: H, fontSize: 14, bold: true, color: TEAL, charSpacing: 2 });
  tb(s, [{ text: "One governance model, adaptable local testing", options: { bold: true, fontFace: H, fontSize: 16, breakLine: true } }, { text: "Make release risks visible and scale reusable quality controls.", options: { color: SLATE } }], { x, y: 2.4, w, h: 0.62, fontSize: 14 });
  steps(s, ["Standardize RTM, severity, estimates and exit criteria in Azure DevOps.", "Global core suite + local SME-approved rules.", "Automate critical regression; use RCA and quality dashboards."], x, 3.1, w, 0.42);
  tb(s, [{ text: "Evidence it can work: ", options: { bold: true, color: TEAL } }, { text: "Capital One: Workday automation reduced weekly manual testing effort by 80% (Kainos, 2022). Microsoft Azure Test Plans supports linked test evidence and automation." }], { x, y: 4.33, w, h: 0.75, fontSize: 14 });
  carPanel(s, x, 5.12, w, 1.46, [["30%", "less rework"], ["35%", "better defect resolution"]], [["C: ", "Rework and defect delays."], ["A: ", "Reusable assets, ALM/JIRA optimization, RTM and coaching."], ["R: ", "30% less rework; 35% better defect resolution."]]);
  challengePanel(s, 7.33, 2.0, 5.5, 4.55, [
    ["Documented business signals", "88 countries in 2025 report; 6 acquisitions in 2025. GBCI 2026: 81 jurisdictions, 292 indicators. 92% revenue retention in 2025.", 1.32],
    ["Potential business impact", "Different local rules and test practices can increase rework and compliance exposure; inconsistent release evidence can delay decisions.", 1.32],
    ["What the JD calls for", "Distributed leadership, Azure DevOps, automation, metrics and continuous improvement.", 0.85],
  ]);
  sources(s, "TMF annual report 2025 • GBCI 2026 research • Kainos • Capital One • Microsoft • Azure Test Plans");
  s.addNotes("Public signals establish business exposure, not confirmed internal DCP testing deficiencies. Validate the current baseline with the hiring manager. Candidate achievements: updated TMF-specific resume. External case study outcomes are publisher-reported, not guaranteed TMF results.");
}

// ---------- Slide 6: Roadmap & why hire me ----------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  header(s, "Execution roadmap & why hire me", "Deliver measurable results in 90–180 days", "Baseline first; pilot improvements; scale only after evidence of value.");
  const phases = [["0–30 days", "Map DCP journeys, interfaces and risks. Baseline metrics and tooling."], ["31–60 days", "Agree governance, RTM, release gates and pilot scope."], ["61–90 days", "Pilot API/data checks, parallel validation and dashboards."], ["91–180 days", "Scale successful pilots, local suites and coaching; review benefits."]];
  const lx = 0.5, lw = 6.55, rowH = 0.6, y0 = 2.05;
  s.addShape(pres.shapes.LINE, { x: lx + 0.14, y: y0 + 0.2, w: 0, h: rowH * 3, line: { color: AQUA, width: 2 } });
  phases.forEach((p, i) => {
    const y = y0 + i * rowH;
    s.addShape(pres.shapes.OVAL, { x: lx, y: y + 0.06, w: 0.28, h: 0.28, fill: { color: i === 3 ? RED : TEAL }, line: { color: WHITE, width: 1.5 } });
    tb(s, p[0], { x: lx + 0.45, y, w: 1.3, h: 0.4, fontFace: H, fontSize: 14, bold: true, color: TEAL, valign: "middle" });
    tb(s, p[1], { x: lx + 1.8, y, w: lw - 1.8, h: 0.5, fontSize: 14, valign: "top" , margin:[2,0,0,0]});
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lx, y: 4.55, w: lw, h: 1.52, fill: { color: ICE }, line: { color: ICE }, rectRadius: 0.08 });
  tb(s, [{ text: "PROPOSED DAY-180 TARGETS ", options: { bold: true, color: TEAL, fontSize: 12, charSpacing: 1 } }, { text: "(subject to baseline and stakeholder approval)", options: { color: SLATE, fontSize: 11 } }], { x: lx + 0.25, y: 4.63, w: lw - 0.5, h: 0.28 });
  tb(s, [
    { text: "15–20% lower leakage", options: { bullet: true, breakLine: true } },
    { text: "15–25% fewer integration defects", options: { bullet: true, breakLine: true } },
    { text: "20–30% less regression effort", options: { bullet: true } },
  ], { x: lx + 0.25, y: 4.95, w: 2.95, h: 1.05, fontSize: 14, paraSpaceAfter: 2 });
  tb(s, [
    { text: "100% critical RTM coverage", options: { bullet: true, breakLine: true } },
    { text: "Zero unexplained critical financial discrepancies at release", options: { bullet: true } },
  ], { x: lx + 3.25, y: 4.95, w: lw - 3.45, h: 1.05, fontSize: 14, paraSpaceAfter: 2 });
  tb(s, "My commitment: protect client trust through reliable, predictable digital delivery.", { x: lx, y: 6.16, w: lw, h: 0.5, fontSize: 14, italic: true, bold: true, color: NAVY, valign: "middle" });

  // right panel
  const px = 7.35, pw = 5.48, py = 1.95, ph = 4.65;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: ph, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08 });
  tb(s, "WHY HIRE ME", { x: px + 0.3, y: py + 0.2, w: pw - 0.6, h: 0.26, fontSize: 12, bold: true, color: AQUA, charSpacing: 2 });
  tb(s, "Proven delivery leadership. Measurable quality outcomes.", { x: px + 0.3, y: py + 0.47, w: pw - 0.6, h: 0.52, fontFace: H, fontSize: 15, bold: true, color: WHITE });
  const st = [["18%", "Lower defect leakage", "CIBC migration governance"], ["22%", "Fewer integration defects", "JPMorgan SIT/UAT"], ["35%", "Improved defect resolution", "ALM/JIRA workflows"], ["30%", "Less testing rework", "Reusable regression assets"]];
  const sw = (pw - 0.6 - 3 * 0.12) / 4;
  st.forEach((v, i) => {
    const x = px + 0.3 + i * (sw + 0.12), y = py + 1.1;
    tb(s, v[0], { x, y, w: sw, h: 0.45, fontFace: H, fontSize: 26, bold: true, color: WHITE });
    tb(s, [{ text: v[1], options: { bold: true, color: AQUA, breakLine: true } }, { text: v[2], options: { color: MIST } }], { x, y: y + 0.48, w: sw, h: 0.8, fontSize: 11 });
  });
  tb(s, [
    { text: "15+ years reported QA experience  |  4 years QA Manager  |  ISTQB Foundation", options: { breakLine: true } },
    { text: "Six engineers led • 12+ migration cycles • Agile, Waterfall and Hybrid", options: { breakLine: true } },
    { text: "SQL and financial validation • Engineering degree + PGDM Finance • Playwright / AI training", options: { breakLine: true } },
    { text: "Transferable ALM/JIRA skills; build Azure DevOps and payroll/HR domain depth with SMEs.", options: { color: AQUA, italic: true } },
  ], { x: px + 0.3, y: py + 2.45, w: pw - 0.6, h: 2.1, fontSize: 14, color: WHITE, paraSpaceAfter: 3 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
