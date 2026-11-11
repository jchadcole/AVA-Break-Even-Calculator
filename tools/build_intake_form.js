// Builds customer-data-gathering.docx: a fill-in form that collects every input the
// AVA break-even calculator needs. Run: node tools/build_intake_form.js
// (needs the "docx" npm package: npm install docx). Keep it in step with index.html inputs.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, BorderStyle, AlignmentType, Header, Footer, PageNumber, LevelFormat,
  TableLayoutType, VerticalAlign,
} = require("docx");

const FONT = "Arial";
const INK = "1F2933", MUTED = "5B6670", ACCENT = "1F5FA8", HEAD_FILL = "E8EEF5", FILL = "FFF9E6", LINE = "C9D2DB";
const W = 10080; // 8.5in page with 1.0in-ish margins (1080 DXA each side)
const COLS = [3200, 2880, 1500, 2500];

const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 18, bold: o.bold, italics: o.italics, color: o.color || INK });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], spacing: { before: o.before || 0, after: o.after ?? 60 }, alignment: o.align, keepNext: o.keepNext, numbering: o.numbering });
const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(children, width, fill) {
  return new TableCell({
    children, width: { size: width, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.TOP,
    shading: fill ? { type: ShadingType.CLEAR, color: "auto", fill } : undefined,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
  });
}

function headerRow(labels) {
  return new TableRow({
    tableHeader: true, cantSplit: true,
    children: labels.map((l, i) => cell([p(t(l, { bold: true, size: 17 }), { after: 0 })], COLS[i], HEAD_FILL)),
  });
}

// item: [name, description, where to find it, example, answer hint]
function itemRow([name, desc, where, example, hint]) {
  return new TableRow({
    cantSplit: true,
    children: [
      cell([p(t(name, { bold: true })), p(t(desc, { size: 16, color: MUTED }), { after: 0 })], COLS[0]),
      cell([p(t(where, { size: 16 }), { after: 0 })], COLS[1]),
      cell([p(t(example, { size: 16 }), { after: 0 })], COLS[2]),
      cell([p(t(hint || "", { size: 15, color: MUTED, italics: true }), { after: 0 })], COLS[3], FILL),
    ],
  });
}

function table(items) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: COLS, layout: TableLayoutType.FIXED,
    rows: [headerRow(["What we need", "Where to find it", "Example", "Your answer"]), ...items.map(itemRow)],
  });
}

function section(num, title, intro, items) {
  return [
    new Paragraph({
      children: [t(`${num}. ${title}`, { size: 24, bold: true, color: ACCENT })],
      spacing: { before: 280, after: 80 }, keepNext: true,
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 2 } },
    }),
    p(t(intro, { size: 18, color: MUTED }), { after: 120, keepNext: true }),
    table(items),
  ];
}

function infoTable(rows) {
  const c = [2800, W - 2800];
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: c, layout: TableLayoutType.FIXED,
    rows: rows.map(r => new TableRow({ cantSplit: true, children: [
      cell([p(t(r, { bold: true }), { after: 0 })], c[0], HEAD_FILL),
      cell([p(t(""), { after: 0 })], c[1], FILL),
    ] })),
  });
}

const bullet = (runs) => p(runs, { numbering: { reference: "dots", level: 0 }, after: 40 });

const children = [
  p(t("AVA Savings Assessment", { size: 40, bold: true, color: ACCENT }), { after: 40 }),
  p(t("Data gathering form", { size: 26, color: MUTED }), { after: 200 }),
  p(t("This form collects the figures we need to estimate what an AI virtual agent (AVA) would save in your voice contact center: the share of calls AVA must handle on its own to break even, and the savings as that share grows. It takes about 20 minutes with your workforce, telephony and finance reports to hand."), { after: 120 }),
  p(t("How to fill it in", { bold: true, size: 20 }), { after: 60, keepNext: true }),
  bullet([t("Type your answers in the shaded "), t("Your answer", { bold: true }), t(" column. The examples show the format only; they are not recommendations.")]),
  bullet([t("Use monthly figures and averages over a recent typical period (3 to 12 months is ideal). Leave out seasonal peaks unless they are the norm.")]),
  bullet([t("Answer for the voice queues in scope only. If several queues are in scope, either combine them or copy this form for each one.")]),
  bullet([t("A best estimate is fine. Mark it "), t("(est.)", { bold: true }), t(" and we will confirm it with you. If you don't know, leave it blank and we will discuss it.")]),
  p(t("About you", { size: 24, bold: true, color: ACCENT }), {
    before: 240, after: 100, keepNext: true,
  }),
  infoTable(["Company", "Contact name and role", "Email and phone", "Date completed", "Queues or lines of business in scope", "Period the figures cover"]),

  ...section(1, "Call volume",
    "Give either your monthly call volume or your current agent headcount. We only need one; the other is calculated from it. Calls are preferred if you have them.",
    [
      ["Inbound calls per month", "Every inbound voice call entering the queues in scope, including calls the IVR resolves.", "Telephony or ACD reports (offered or presented calls); IVR entry counts.", "60,000 calls", "calls per month"],
      ["Or: agents handling these calls", "Full-time equivalent (FTE) agents staffed on these queues today. Two half-time agents count as one FTE.", "Workforce management (WFM) system; HR roster; staffing plan.", "100 FTEs", "FTEs (only if calls are unknown)"],
    ]),

  ...section(2, "How calls flow today",
    "The share of all inbound calls that takes each path. These are percentages of every call coming in.",
    [
      ["Calls that go straight to an agent", "Calls that skip the IVR or opt out at once and reach a person without any self-service.", "ACD routing reports; IVR \"zero-out\" or bypass reports.", "10%", "% of all calls"],
      ["Calls fully resolved in the IVR", "Calls that finish in the IVR with no agent, such as balance checks or payments.", "IVR containment or self-service completion reports.", "20%", "% of all calls"],
      ["Everything else", "The rest of your calls leave the IVR and reach an agent today. These are the calls AVA would answer.", "Calculated: 100% minus the two lines above.", "70%", "No entry needed"],
    ]),

  ...section(3, "Agent cost and time",
    "These figures turn agent time saved into headcount and dollars.",
    [
      ["Fully loaded cost per agent", "Total monthly cost of one FTE agent: wages, benefits, payroll taxes, plus a share of supervision, facilities, technology and outsourcing fees.", "Finance or HR; outsourcer rate card (hourly rate × paid hours).", "$7,000 per month", "$ per FTE per month"],
      ["Paid hours per agent", "Hours one full-time agent is paid in a month.", "HR policy or payroll. 40 hr/week × 52 ÷ 12 = 173.", "173 hours", "hours per month"],
      ["Shrinkage", "Share of paid time when agents can't work at all: breaks, training, meetings, coaching, sick time and vacation. Do not include time on other tasks.", "WFM shrinkage report; workforce planning assumptions.", "30%", "% of paid time"],
      ["Share of working time on these calls", "Of the time agents are available to work, the share spent on these voice calls. The rest is other tasks such as email, chat or back-office work.", "WFM or ACD reports (logged-in time by queue or activity); team leads.", "100% (voice only) or 70%", "% of working time"],
      ["Occupancy on these calls", "Of the time agents are logged in to these calls, the share spent talking, holding or on wrap-up rather than waiting for the next call. If agents fill the gaps between calls with other work, give their overall occupancy.", "WFM or ACD occupancy report for these queues.", "80%", "% of logged-in time"],
      ["Average handle time", "Average agent time per call: talk + hold + after-call work.", "ACD or WFM reports (AHT).", "10 min (600 sec)", "minutes or seconds"],
      ["Count savings in whole agents only?", "Yes if savings only count when a full agent position is removed. No if partial FTEs count (overtime, outsourced hours, attrition).", "Operations or finance leadership.", "Yes", "Yes / No"],
    ]),

  ...section(4, "Per-call automation fees",
    "What each self-service call costs. Each call pays one automation fee: a call the IVR resolves pays the IVR fee; a call that reaches AVA pays the AVA fee instead.",
    [
      ["IVR cost per call", "What you pay for each call handled by your IVR: per-call or per-minute charges and telephony usage.", "IVR or telephony invoices; vendor contract.", "$0.50 per call", "$ per call"],
      ["AVA cost per call", "Proposed price for each call AVA answers. Leave blank if your account team will supply it.", "AVA proposal or quote.", "$1.20 per call", "$ per call"],
      ["Is AVA billed on calls it passes to an agent?", "Yes if every call AVA answers is billed, even when it later transfers to an agent. No if only calls AVA fully resolves are billed.", "AVA proposal or contract terms.", "Yes", "Yes / No"],
    ]),

  ...section(5, "AVA fixed and one-time costs",
    "Enter $0 for any cost that does not apply.",
    [
      ["Monthly platform fee", "Any fixed monthly AVA charge on top of per-call fees: subscription, licenses, hosting, support.", "AVA proposal or quote.", "$0 per month", "$ per month"],
      ["One-time implementation cost", "Setup, integration, design and professional services, plus your own internal project costs.", "AVA proposal; internal project budget.", "$50,000", "$ one-time"],
      ["Months to spread the one-time cost over", "If the one-time cost is financed or budgeted over time, the number of months. Otherwise use 12.", "Finance.", "12 months", "months"],
    ]),

  ...section(6, "Rollout and goals",
    "How quickly AVA goes live and how much it is expected to handle. If you have no view yet, leave these blank and we will propose figures based on your call types.",
    [
      ["Implementation time", "Months from contract signature until AVA is live and answering calls.", "Project plan; AVA proposal.", "3 months", "months (0–12)"],
      ["Target AVA containment, first year", "The share of all inbound calls AVA should resolve without an agent once live.", "Your goals; top call reasons that are simple and repeatable.", "20% of all calls", "% of all calls"],
      ["Yearly increase in containment", "How many percentage points AVA's share should grow each year as more call types are added.", "Your roadmap; AVA proposal.", "+5 points per year", "points per year"],
      ["Years to project", "How many years of savings to show, from 1 to 15.", "Usually matches your planning or contract term.", "5 years", "years"],
    ]),

  new Paragraph({
    children: [t("7. Anything else we should know", { size: 24, bold: true, color: ACCENT })],
    spacing: { before: 280, after: 80 }, keepNext: true,
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 2 } },
  }),
  p(t("For example: planned volume changes, seasonal peaks, outsourcing contracts, top call reasons, or anything that makes your figures unusual.", { size: 18, color: MUTED }), { after: 120, keepNext: true }),
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED,
    rows: [new TableRow({ height: { value: 1300, rule: "atLeast" }, children: [cell([p(t(""))], W, FILL)] })],
  }),
  p(t("Thank you. Please return the completed form to your account team. We will use it to prepare your break-even and savings estimate and review the results with you.", { size: 18, color: MUTED }), { before: 200 }),
];

const doc = new Document({
  creator: "AVA Savings Assessment",
  title: "AVA Savings Assessment: data gathering form",
  styles: { default: { document: { run: { font: FONT, size: 18, color: INK } } } },
  numbering: { config: [{ reference: "dots", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1000, bottom: 1000, left: 1080, right: 1080 } } },
    headers: { default: new Header({ children: [p(t("AVA Savings Assessment · Data gathering form", { size: 15, color: MUTED }), { align: AlignmentType.RIGHT, after: 0 })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [t("Page ", { size: 15, color: MUTED }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 15, color: MUTED }), t(" of ", { size: 15, color: MUTED }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 15, color: MUTED })] })] }) },
    children,
  }],
});

const out = path.join(__dirname, "..", "customer-data-gathering.docx");
Packer.toBuffer(doc).then(b => { fs.writeFileSync(out, b); console.log("wrote", out); });
