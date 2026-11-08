// THROWAWAY PROTOTYPE data. Tickets #13 and #16.
// Everything here is mock data. No Office.js, gateway, or persistence.

export const WORKBOOK = {
  name: "Q3 Forecast.xlsx",
  path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx",
  modified: "Saved to this PC",
};

export const HISTORY_LIMIT_MB = 250;
export const HISTORY_STORE_PATH = "C:\\Users\\mustapha\\AppData\\Local\\ExcelAgent\\history";

export const HISTORY_WORKBOOKS = [
  { key: "q3-forecast", name: "Q3 Forecast.xlsx", path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx", conversations: 5, sizeMb: 62 },
  { key: "fy27-budget", name: "FY27 Budget.xlsx", path: "C:\\Users\\mustapha\\Documents\\FY27 Budget.xlsx", conversations: 3, sizeMb: 37 },
  { key: "pricing-2026", name: "Pricing 2026.xlsx", path: "C:\\Users\\mustapha\\OneDrive\\Finance\\Pricing 2026.xlsx", conversations: 1, sizeMb: 21 },
  { key: "board-pack", name: "Board Pack Q3.xlsx", path: "\\\\fileserver\\finance\\Board Pack Q3.xlsx", conversations: 1, sizeMb: 14 },
  { key: "archive-2024", name: "Archive 2024.xlsx", path: "D:\\Archive\\Archive 2024.xlsx", conversations: 1, sizeMb: 18 },
];

export const HISTORY_WORKBOOKS_RETENTION = [
  { key: "q3-forecast", name: "Q3 Forecast.xlsx", path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx", conversations: 5, sizeMb: 92 },
  { key: "fy27-budget", name: "FY27 Budget.xlsx", path: "C:\\Users\\mustapha\\Documents\\FY27 Budget.xlsx", conversations: 3, sizeMb: 55 },
  { key: "pricing-2026", name: "Pricing 2026.xlsx", path: "C:\\Users\\mustapha\\OneDrive\\Finance\\Pricing 2026.xlsx", conversations: 1, sizeMb: 27 },
  { key: "board-pack", name: "Board Pack Q3.xlsx", path: "\\\\fileserver\\finance\\Board Pack Q3.xlsx", conversations: 1, sizeMb: 16 },
  { key: "archive-2024", name: "Archive 2024.xlsx", path: "D:\\Archive\\Archive 2024.xlsx", conversations: 1, sizeMb: 15 },
];

export const HISTORY_WORKBOOKS_FULL = [
  { key: "q3-forecast", name: "Q3 Forecast.xlsx", path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx", conversations: 5, sizeMb: 120 },
  { key: "fy27-budget", name: "FY27 Budget.xlsx", path: "C:\\Users\\mustapha\\Documents\\FY27 Budget.xlsx", conversations: 3, sizeMb: 70 },
  { key: "pricing-2026", name: "Pricing 2026.xlsx", path: "C:\\Users\\mustapha\\OneDrive\\Finance\\Pricing 2026.xlsx", conversations: 1, sizeMb: 30 },
  { key: "board-pack", name: "Board Pack Q3.xlsx", path: "\\\\fileserver\\finance\\Board Pack Q3.xlsx", conversations: 1, sizeMb: 18 },
  { key: "archive-2024", name: "Archive 2024.xlsx", path: "D:\\Archive\\Archive 2024.xlsx", conversations: 1, sizeMb: 20 },
];

export const INSTRUCTIONS_FROM_FY27 = `Use the FY27 driver table as the source of truth.
Format currency as EUR for the European rows.
Name the scenario and its owner in the first column.`;

export const COLUMNS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];

export const SHEETS = {
  Forecast: {
    label: "Forecast",
    rows: [
      ["Segment", "Index", "Q2 units", "Q3 units", "Q2 ASP", "Q3 ASP", "Q2 revenue", "Q3 revenue"],
      ["Enterprise", "1.08", "12,400", "13,392", "$339", "$304", "$4,203,600", "$4,071,168"],
      ["Mid-market", "1.08", "18,900", "20,412", "$151", "$143", "$2,845,900", "$2,918,916"],
      ["SMB", "1.08", "31,200", "32,760", "$52", "$48", "$1,622,400", "$1,572,480"],
      ["Public sector", "1.08", "7,800", "8,190", "$196", "$191", "$1,528,800", "$1,564,290"],
      [],
      [],
      ["", "", "", "", "", "", "Price index", "1.08"],
      ["", "", "", "", "", "", "Q3 units vs Q2", "+8.2%"],
      ["", "", "", "", "", "", "Q3 revenue vs Q2", "−3.1%"],
      ["", "", "", "", "", "", "Mix effect", "−$126K"],
      [],
      [],
      ["Segment", "Q2 units", "Q3 units", "Q2 ASP", "Q3 ASP"],
      ["Enterprise", "12,400", "13,392", "$339", "$304"],
      ["Mid-market", "18,900", "20,412", "$151", "$143"],
      ["SMB", "31,200", "32,760", "$52", "$48"],
      ["Public sector", "7,800", "8,190", "$196", "$191"],
      [],
      ["Q2 vs Q3 — segment detail"],
      ["", "Q2", "Q3", "Change", "ASP Δ", "Note"],
      ["", "$4.21M", "$4.08M", "−3.1%", "−10.4%", "Price-led"],
      ["", "$2.84M", "$2.91M", "+2.5%", "−5.3%", "Volume-led"],
      ["", "$1.62M", "$1.58M", "−2.5%", "−7.7%", "Price-led"],
      ["", "$1.53M", "$1.56M", "+2.3%", "−2.6%", "Volume-led"],
      [],
      ["", "", "", "", "", "Total", "$8.13M", "$8.13M"],
    ],
  },
  Regions: {
    label: "Regions",
    rows: [
      ["Region", "Q2 revenue", "Q3 revenue", "Change", "Share", "Rank", "Owner", "Note"],
      ["North America", "$3,102,400", "$3,018,900", "−2.7%", "37.1%", "1", "A. Reed", ""],
      ["Europe", "$2,684,100", "$2,746,300", "+2.3%", "33.8%", "2", "M. Weber", "EUR"],
      ["APAC", "$1,418,700", "$1,502,600", "+5.9%", "18.5%", "3", "R. Sato", ""],
      ["LATAM", "$498,300", "$501,900", "+0.7%", "6.2%", "4", "L. Ortiz", ""],
      ["Middle East", "$251,900", "$245,800", "−2.4%", "3.0%", "5", "S. Haddad", ""],
      ["Africa", "$59,800", "$60,200", "+0.7%", "0.7%", "6", "K. Mensah", ""],
      ["Japan", "$0", "$0", "—", "0.0%", "7", "R. Sato", "No activity"],
      ["Other", "$0", "$0", "—", "0.0%", "8", "", ""],
    ],
  },
  Inputs: {
    label: "Inputs",
    rows: [
      ["Driver", "Value", "Unit", "Source", "Owner"],
      ["Price index", "1.08", "index", "Finance", "M. Weber"],
      ["Volume growth", "8.2%", "YoY", "Sales ops", "A. Reed"],
      ["Discount rate", "4.5%", "avg", "Deal desk", "L. Ortiz"],
      ["EUR / USD", "1.087", "FX", "Treasury", "M. Weber"],
      ["GBP / USD", "1.271", "FX", "Treasury", "M. Weber"],
      ["Churn", "6.1%", "annual", "Customer success", "S. Haddad"],
      ["Tax rate", "21.0%", "statutory", "Finance", "A. Reed"],
      [],
      ["This workbook is a prototype. Nothing here is real."],
    ],
  },
};

export const VARIANT_META = {
  hybrid: { code: "A", name: "Hybrid", blurb: "Warm paper · Fluent bones" },
  claude: { code: "B", name: "Claude-close", blurb: "Floating paper · editorial" },
  fluent: { code: "C", name: "Fluent-native", blurb: "Command bar · Excel-native" },
};

export const VARIANTS = ["hybrid", "claude", "fluent"];
export const WIDTHS = [260, 380, 520, 720];
export const THEMES = ["auto", "light", "dark"];

export const SCENARIOS = [
  { key: "empty", label: "Empty / first run" },
  { key: "loading", label: "Loading / boot" },
  { key: "streaming", label: "Streaming answer" },
  { key: "answer", label: "Answer complete" },
  { key: "review", label: "Review proposal" },
  { key: "partial", label: "Partial apply" },
  { key: "applied", label: "Applied" },
  { key: "history", label: "History" },
  { key: "settings", label: "Settings / instructions" },
  { key: "error", label: "Error / offline" },
  { key: "link", label: "Link an old conversation" },
  { key: "unsaved", label: "Unsaved workbook" },
  { key: "compaction", label: "Compaction · 80%" },
  { key: "compaction-progress", label: "Compaction · summarizing" },
  { key: "compaction-stacked", label: "Compaction · second cut" },
  { key: "retention", label: "History at 80%" },
  { key: "retention-full", label: "History over budget" },
  { key: "storefail", label: "History store failure" },
  { key: "instructions", label: "Instructions copy chip" },
  { key: "readonly", label: "Read-only mode" },
  { key: "readonly-forced", label: "Read-only workbook" },
];

export const MODELS = [
  { name: "gpt-5.2", tag: "default" },
  { name: "claude-sonnet-4.5", tag: "" },
  { name: "gemini-3-pro", tag: "" },
  { name: "qwen3-coder", tag: "tools" },
];

const SEGMENTS = [
  ["Enterprise", "$4.21M", "$4.08M", "−3.1%"],
  ["Mid-market", "$2.84M", "$2.91M", "+2.5%"],
  ["SMB", "$1.62M", "$1.58M", "−2.5%"],
  ["Public sector", "$1.53M", "$1.56M", "+2.3%"],
];

const TABLE = [
  "| Segment | Q2 | Q3 | Change |",
  "| --- | ---: | ---: | ---: |",
  ...SEGMENTS.map(([name, q2, q3, change]) => `| ${name} | ${q2} | ${q3} | ${change} |`),
].join("\n");

export const RESPONSES = {
  question: {
    done: "3 steps · 4.8 s",
    steps: [
      "Read Forecast!A1:H400",
      "Read Forecast!B22:F25",
      "Checked 12 formulas",
    ],
    stageProposal: false,
    script: [
      { status: { label: "Reading Forecast!A1:H400", step: "step 1 of 3" }, delay: 420 },
      {
        text: "Unit volume is **up 8.2%**, but the Q3 total is **3.1% lower**. The gap is price, not volume.\n\n",
        delay: 420,
      },
      { status: { label: "Summarising by segment", step: "step 2 of 3" }, delay: 420 },
      { text: `${TABLE}\n\n`, delay: 380 },
      {
        text: "The Enterprise line carries the drop: [Forecast!B22:F25](cite:Forecast!B22:F25) shows average selling price down **10.4%** while units rose **8.0%**. Cell [H8](cite:Forecast!H8) still applies the old **1.08** index, so the Q3 total is understated by about **$126K**.\n\n",
        delay: 520,
      },
      { status: { label: "Checking the index", step: "step 3 of 3" }, delay: 360 },
      {
        text: "> The forecast is sound; the index is stale. Update H8 and refresh the revenue column to reconcile the two totals.",
        delay: 420,
      },
    ],
  },
  summary: {
    done: "3 steps · 5.1 s",
    steps: [
      "Read Forecast!A1:H400",
      "Summarised 4 segments",
      "Compared Q2 and Q3 totals",
    ],
    stageProposal: false,
    script: [
      { status: { label: "Reading Forecast!A1:H400", step: "step 1 of 3" }, delay: 380 },
      {
        text: "The Forecast sheet has **four segments**, **1,240 rows** of quarterly data, and one stale input.\n\n",
        delay: 380,
      },
      { text: `${TABLE}\n\n`, delay: 360 },
      {
        text: "Enterprise and SMB are both price-led declines. Mid-market and Public sector grew on volume. [B22:F25](cite:Forecast!B22:F25) is the comparison table; [H8](cite:Forecast!H8) holds the index that drives Q3 revenue.\n\n",
        delay: 460,
      },
      { status: { label: "Checking the inputs", step: "step 2 of 3" }, delay: 320 },
      {
        text: "- **Units** are healthy: every segment is up year over year.\n- **Price** is the problem: Enterprise ASP fell 10.4%.\n- **Index** is stale: H8 still uses 1.08 instead of 1.124.\n\n",
        delay: 400,
      },
      { status: { label: "Writing the summary", step: "step 3 of 3" }, delay: 340 },
      {
        text: "The headline is simple: demand grew, price fell, and the index made the fall look smaller than it is.",
        delay: 400,
      },
    ],
  },
  formulas: {
    done: "3 steps · 6.2 s",
    steps: [
      "Read Forecast!A1:H400",
      "Scanned 312 formulas",
      "Classified 3 risks",
    ],
    stageProposal: false,
    script: [
      { status: { label: "Reading Forecast!A1:H400", step: "step 1 of 3" }, delay: 360 },
      {
        text: "Three formulas need a second look.\n\n",
        delay: 280,
      },
      { status: { label: "Scanning formulas", step: "step 2 of 3" }, delay: 400 },
      {
        text: "- [H8](cite:Forecast!H8) hard-codes **1.08**; it should point at [Inputs!B2](cite:Inputs!B2).\n- [D4:D40](cite:Forecast!D4:D40) multiplies by that literal instead of `$B$1`.\n- [K2](cite:Forecast!K2) uses `INDIRECT`, which can reach outside this workbook.\n\n",
        delay: 520,
      },
      { status: { label: "Classifying risk", step: "step 3 of 3" }, delay: 360 },
      {
        text: "The first two are ordinary overwrites. The third is a confirm-tier formula: it would need an explicit row tick before it could ever be applied.\n\nWant me to stage the two safe fixes?",
        delay: 440,
      },
    ],
  },
  cleanup: {
    done: "3 steps · 4.2 s",
    steps: [
      "Read Forecast!A1:H400",
      "Read the selected table",
      "Staged 12 changes",
    ],
    stageProposal: true,
    script: [
      { status: { label: "Reading the selected table", step: "step 1 of 3" }, delay: 360 },
      {
        text: "I can normalise the headers, refresh the totals, and clean up the number formats. Review the changes before anything is written.\n\n",
        delay: 380,
      },
      { status: { label: "Preparing the diff", step: "step 2 of 3" }, delay: 420 },
      {
        text: "Two rows are excluded by default: the `INDIRECT` formula and the clear operation. Tick them in the card if you want them included.\n\n",
        delay: 440,
      },
      { status: { label: "Staging 12 changes", step: "step 3 of 3" }, delay: 400 },
      {
        text: "The overwrite acknowledgement is the only card-level gate. Apply is one Ctrl+Z step in Excel.",
        delay: 400,
      },
    ],
  },
  generic: {
    done: "2 steps · 3.4 s",
    steps: ["Read Forecast!A1:H400", "Read the selection"],
    stageProposal: false,
    script: [
      { status: { label: "Reading Forecast!A1:H400", step: "step 1 of 2" }, delay: 360 },
      {
        text: "I read the current selection and the Forecast summary. [B22:F25](cite:Forecast!B22:F25) is the cleanest comparison; [H8](cite:Forecast!H8) is the stale input.\n\n",
        delay: 420,
      },
      { status: { label: "Reading the selection", step: "step 2 of 2" }, delay: 340 },
      {
        text: "Ask me to stage a fix if you want to see the review flow — or try one of the suggestions on a new chat.",
        delay: 420,
      },
    ],
  },
};

export function responseForPrompt(prompt) {
  const p = prompt.toLowerCase();
  const asks = /\b(why|how|compare|comparison|variance|lower|higher|drop|difference|explain|revenue)\b/.test(p);
  const acts = /\b(clean|tidy|fix|update|refresh|sort|format|change|normalise|normalize)\b/.test(p);
  if (asks && !acts) return "question";
  if (acts) return "cleanup";
  if (/formula|second look|risk|check|audit/.test(p)) return "formulas";
  if (/summar|overview|what|describe/.test(p)) return "summary";
  return "generic";
}

export const ANSWER_MESSAGES = [
  { role: "user", text: "Why is the Q3 total lower than Q2 when units are up?" },
  {
    role: "assistant",
    text: "Unit volume is **up 8.2%**, but the Q3 total is **3.1% lower**. The gap is price, not volume.\n\n" +
      TABLE +
      "\n\nThe Enterprise line carries the drop: [Forecast!B22:F25](cite:Forecast!B22:F25) shows average selling price down **10.4%** while units rose **8.0%**. Cell [H8](cite:Forecast!H8) still applies the old **1.08** index, so the Q3 total is understated by about **$126K**.",
    done: "3 steps · 4.8 s",
    steps: ["Read Forecast!A1:H400", "Read Forecast!B22:F25", "Checked 12 formulas"],
  },
];

export const PROPOSAL_ROWS = [
  {
    id: "r1",
    sheet: "Forecast",
    range: "B1",
    op: "Set value",
    before: "1.08",
    after: "1.124",
    risk: "overwrite",
    note: "Overwrites a non-empty input cell.",
    field: "value",
    included: true,
    cells: 1,
  },
  {
    id: "r2",
    sheet: "Forecast",
    range: "D4:D40",
    op: "Set formula",
    before: "=C4*1.08",
    after: "=C4*$B$1",
    risk: "overwrite",
    note: "Updates 37 formulas in the revenue column.",
    field: "formula",
    included: true,
    cells: 37,
  },
  {
    id: "r3",
    sheet: "Forecast",
    range: "F12:F40",
    op: "Format currency",
    before: "General",
    after: "$#,##0",
    risk: "normal",
    note: "",
    field: "numberFormat",
    included: true,
    cells: 29,
  },
  {
    id: "r4",
    sheet: "Forecast",
    range: "H8",
    op: "Set value",
    before: "1.08",
    after: "1.124",
    risk: "overwrite",
    note: "Overwrites the index that drives Q3 revenue.",
    field: "value",
    included: true,
    cells: 1,
  },
  {
    id: "r5",
    sheet: "Forecast",
    range: "K2",
    op: "Set formula",
    before: "blank",
    after: '=INDIRECT("Forecast!B1")',
    risk: "confirm",
    note: "INDIRECT can reach outside this workbook.",
    field: "formula",
    included: false,
    cells: 1,
  },
  {
    id: "r6",
    sheet: "Forecast",
    range: "A14:A18",
    op: "Clear contents",
    before: "5 labels",
    after: "blank",
    risk: "loss",
    note: "Clears 5 labels; no formulas are lost.",
    field: "value",
    included: false,
    cells: 5,
  },
  {
    id: "r7",
    sheet: "Forecast",
    range: "G2:G7",
    op: "Set formula",
    before: "=C2*E2",
    after: "=C2*F2",
    risk: "overwrite",
    note: "Replaces 6 revenue formulas.",
    field: "formula",
    included: true,
    cells: 6,
  },
  {
    id: "r8",
    sheet: "Regions",
    range: "D2:D9",
    op: "Format percent",
    before: "0.031",
    after: "3.1%",
    risk: "normal",
    note: "",
    field: "numberFormat",
    included: true,
    cells: 8,
  },
  {
    id: "r9",
    sheet: "Regions",
    range: "F2:F9",
    op: "Set formula",
    before: "=RANK(C2,$C$2:$C$9)",
    after: "=RANK(C2,$C$2:$C$9,0)",
    risk: "overwrite",
    note: "Replaces 8 rank formulas.",
    field: "formula",
    included: true,
    cells: 8,
  },
  {
    id: "r10",
    sheet: "Regions",
    range: "C2:C9",
    op: "Format currency",
    before: "General",
    after: "€#,##0",
    risk: "normal",
    note: "",
    field: "numberFormat",
    included: true,
    cells: 8,
  },
  {
    id: "r11",
    sheet: "Inputs",
    range: "B2",
    op: "Set value",
    before: "1.08",
    after: "1.124",
    risk: "overwrite",
    note: "Overwrites the price index input.",
    field: "value",
    included: true,
    cells: 1,
  },
  {
    id: "r12",
    sheet: "Inputs",
    range: "D2:D8",
    op: "Set validation",
    before: "none",
    after: "list: Finance, Sales ops, Treasury",
    risk: "normal",
    note: "Native operation. Adds a dropdown to the Source column.",
    field: "validation",
    included: true,
    cells: 7,
  },
];

export function makeProposal(status = "proposed") {
  const rows = PROPOSAL_ROWS.map((row) => ({ ...row }));
  const base = {
    rows,
    ack: false,
    collapsed: false,
    showAllRows: false,
    showApplied: false,
    status,
    appliedCount: 0,
    excludedCount: 0,
    failedCount: 0,
    message: null,
    appliedAt: null,
  };
  if (status === "applied") {
    base.collapsed = true;
    base.ack = true;
    base.status = "applied";
    base.appliedCount = 10;
    base.excludedCount = 2;
    base.appliedAt = "2:14 PM";
    rows.forEach((row) => {
      if (row.risk === "confirm" || row.risk === "loss") {
        row.included = false;
        row.outcome = "excluded";
      } else {
        row.included = true;
        row.outcome = "applied";
      }
    });
  }
  if (status === "partial") {
    base.ack = true;
    base.status = "partial";
    base.appliedCount = 10;
    base.failedCount = 2;
    rows.forEach((row) => {
      row.included = true;
      row.outcome = "applied";
    });
    rows[10].outcome = "failed";
    rows[10].error = "Excel was still in cell-edit mode.";
    rows[11].outcome = "failed";
    rows[11].error = "Excel was busy when the row was attempted.";
  }
  return base;
}

export const INSTRUCTIONS_DEFAULT = `Format numbers with thousands separators.
Bold column headers.
Use EUR for currency.
Always show the formula behind a total.`;

export const COMPACTION_MESSAGES = [
  { role: "user", text: "Why is Q3 revenue down when units are up?" },
  { role: "assistant", text: "Unit volume is **up 8.2%**, but average selling price fell **10.4%**. The gap is price, not demand." },
  { role: "user", text: "Where is the price fall concentrated?" },
  { role: "assistant", text: "Enterprise carries most of it: [B22:F25](cite:Forecast!B22:F25) shows ASP down from **$339** to **$304**." },
  { role: "user", text: "What about the other segments?" },
  { role: "assistant", text: "Mid-market and SMB are down **5.3%** and **7.7%**. Public sector is down **2.6%**." },
  { role: "user", text: "Is the index stale?" },
  { role: "assistant", text: "Yes. [H8](cite:Forecast!H8) still applies **1.08**, while [Inputs!B2](cite:Inputs!B2) says **1.124**." },
  { role: "user", text: "How much does that understate revenue?" },
  { role: "assistant", text: "About **$126K** on the Q3 total, concentrated in Enterprise." },
  { role: "user", text: "Which regions are affected?" },
  { role: "assistant", text: "Europe and APAC are least affected. North America carries the Enterprise exposure." },
  { role: "user", text: "Should we change the forecast?" },
  { role: "assistant", text: "Yes: update [H8](cite:Forecast!H8) and refresh the revenue column, then compare the variance again." },
  { role: "user", text: "What else should I check?" },
  { role: "assistant", text: "Check the two hard-coded formulas in the revenue column and the FX inputs in [Inputs!B2](cite:Inputs!B2)." },
  { role: "user", text: "Can you make those changes?" },
  { role: "assistant", text: "I can stage them as a proposal. Review the diff before anything is written." },
];

export const CONVERSATIONS = [
  {
    id: "c-current",
    title: "Why is Q3 revenue lower?",
    preview: "Unit volume is up 8.2%, but price is down 10.4%…",
    time: "2:14 PM",
    group: "Today",
    date: "8 Nov 2026",
    count: 6,
    sheets: "Forecast",
    workbook: "Q3 Forecast.xlsx",
    path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx",
    current: true,
    pending: false,
  },
  {
    id: "c-cleanup",
    title: "Clean up the forecast table",
    preview: "I’d normalise the headers and refresh the totals…",
    time: "11:02 AM",
    group: "Today",
    date: "8 Nov 2026",
    count: 4,
    sheets: "Forecast",
    workbook: "Q3 Forecast.xlsx",
    path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx",
    pending: true,
    scenario: "partial",
    thread: [
      { role: "user", text: "Clean up the forecast table." },
      {
        role: "assistant",
        text: "I can normalise the headers, refresh the totals, and clean up the number formats. Review the changes before anything is written.",
      },
    ],
  },
  {
    id: "c-pivot",
    title: "Build a regional pivot",
    preview: "Pivot created on a new sheet: Revenue by region.",
    time: "4:38 PM",
    group: "Yesterday",
    date: "7 Nov 2026",
    count: 8,
    sheets: "Forecast, Regions",
    workbook: "Q3 Forecast.xlsx",
    path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx",
    thread: [
      { role: "user", text: "Build a regional pivot by revenue." },
      {
        role: "assistant",
        text: "Pivot created on a new sheet: **Revenue by region**. It uses the Regions table and groups by `Owner`.",
      },
    ],
  },
  {
    id: "c-h8",
    title: "Check the H8 formula",
    preview: "It uses a hard-coded 1.08 index, so Q3 is…",
    time: "9:15 AM",
    group: "Yesterday",
    date: "7 Nov 2026",
    count: 3,
    sheets: "Forecast",
    workbook: "Q3 Forecast.xlsx",
    path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx",
    thread: [
      { role: "user", text: "Check the H8 formula." },
      {
        role: "assistant",
        text: "[H8](cite:Forecast!H8) hard-codes **1.08** instead of pointing at [Inputs!B2](cite:Inputs!B2). That is why Q3 revenue looks understated.",
      },
    ],
  },
  {
    id: "c-variance",
    title: "Q2 vs Q3 variance",
    preview: "The variance is driven by two accounts in Enterprise.",
    time: "Mon",
    group: "Earlier",
    date: "3 Nov 2026",
    count: 5,
    sheets: "Forecast",
    workbook: "Q3 Forecast.xlsx",
    path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx",
    thread: [
      { role: "user", text: "Where is the Q2 vs Q3 variance coming from?" },
      {
        role: "assistant",
        text: "The variance is driven by two accounts in Enterprise. The rest of the book is within **±3%**.",
      },
    ],
  },
  {
    id: "c-budget",
    title: "Draft next year’s budget sheet",
    preview: "Start with the driver table, then add the rollup.",
    time: "Oct 28",
    group: "Earlier",
    date: "28 Oct 2026",
    count: 9,
    sheets: "Inputs",
    workbook: "FY27 Budget.xlsx",
    path: "C:\\Users\\mustapha\\Documents\\FY27 Budget.xlsx",
    thread: [
      { role: "user", text: "Draft next year’s budget sheet." },
      {
        role: "assistant",
        text: "Start with the driver table, then add the rollup. I can stage a first pass when you are ready.",
      },
    ],
  },
  {
    id: "c-cleanup2",
    title: "Normalise the region names",
    preview: "Two spellings of “Middle East” were merged…",
    time: "Oct 24",
    group: "Earlier",
    date: "24 Oct 2026",
    count: 4,
    sheets: "Regions",
    workbook: "FY27 Budget.xlsx",
    path: "C:\\Users\\mustapha\\Documents\\FY27 Budget.xlsx",
    thread: [
      { role: "user", text: "Normalise the region names." },
      {
        role: "assistant",
        text: "Two spellings of “Middle East” were merged, and **APAC** is now consistent across the workbook.",
      },
    ],
  },
  {
    id: "c-headcount",
    title: "Headcount plan for FY27",
    preview: "The plan needs one more approver in each region.",
    time: "Oct 21",
    group: "Earlier",
    date: "21 Oct 2026",
    count: 6,
    sheets: "Inputs",
    workbook: "FY27 Budget.xlsx",
    path: "C:\\Users\\mustapha\\Documents\\FY27 Budget.xlsx",
    thread: [
      { role: "user", text: "Review the FY27 headcount plan." },
      {
        role: "assistant",
        text: "The plan needs one more approver in each region. The **Enterprise** line is the only one over its cap.",
      },
    ],
  },
  {
    id: "c-pricing",
    title: "Pricing model review",
    preview: "Discount bands are still using the 2025 thresholds.",
    time: "Oct 12",
    group: "Earlier",
    date: "12 Oct 2026",
    count: 7,
    sheets: "Pricing",
    workbook: "Pricing 2026.xlsx",
    path: "C:\\Users\\mustapha\\OneDrive\\Finance\\Pricing 2026.xlsx",
    thread: [
      { role: "user", text: "Review the pricing model." },
      {
        role: "assistant",
        text: "Discount bands are still using the 2025 thresholds. The **Enterprise** band is the only one where the floor moved.",
      },
    ],
  },
  {
    id: "c-board",
    title: "Board pack variance notes",
    preview: "Two commentary cells still say “pending”.",
    time: "Oct 2",
    group: "Earlier",
    date: "2 Oct 2026",
    count: 3,
    sheets: "Summary",
    workbook: "Board Pack Q3.xlsx",
    path: "\\\\fileserver\\finance\\Board Pack Q3.xlsx",
    thread: [
      { role: "user", text: "Check the variance notes in the board pack." },
      {
        role: "assistant",
        text: "Two commentary cells still say **pending**. The rest of the pack matches the Q3 close.",
      },
    ],
  },
  {
    id: "c-archive",
    title: "Year-end close notes",
    preview: "The close checklist is complete except for the FX revaluation.",
    time: "Sep 30",
    group: "Earlier",
    date: "30 Sep 2026",
    count: 5,
    sheets: "Close",
    workbook: "Archive 2024.xlsx",
    path: "D:\\Archive\\Archive 2024.xlsx",
    thread: [
      { role: "user", text: "Where did the year-end close stop?" },
      {
        role: "assistant",
        text: "The close checklist is complete except for the **FX revaluation**, which is waiting on Treasury.",
      },
    ],
  },
  {
    id: "c-recovered",
    title: "Q2 close review",
    preview: "Recovered from the last good copy after a store error.",
    time: "2:09 PM",
    group: "Today",
    date: "8 Nov 2026",
    count: 4,
    sheets: "Forecast",
    workbook: "Q3 Forecast.xlsx",
    path: "C:\\Users\\mustapha\\Documents\\Q3 Forecast.xlsx",
    recovered: true,
    thread: [
      { role: "user", text: "Review the Q2 close." },
      {
        role: "assistant",
        text: "The Q2 close is consistent. The only exception is the **FX revaluation** row, which was posted after the lock.",
      },
    ],
  },
];

export const COMMANDS = [
  { label: "Summarise the Forecast sheet", icon: "grid", prompt: "Summarise the Forecast sheet" },
  { label: "Find formulas that need a second look", icon: "alert", prompt: "Find formulas that need a second look" },
  { label: "Clean up the selected table", icon: "pen", prompt: "Clean up the selected table" },
];
