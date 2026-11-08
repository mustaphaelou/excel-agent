// THROWAWAY PROTOTYPE — tickets #13 and #16.
// One route. Variants via ?variant=, scenarios via ?state=, width/theme via query.
// No framework, no dependencies, no persistence. Everything is mocked.
import {
  WORKBOOK,
  HISTORY_LIMIT_MB,
  HISTORY_STORE_PATH,
  HISTORY_WORKBOOKS,
  HISTORY_WORKBOOKS_RETENTION,
  HISTORY_WORKBOOKS_FULL,
  INSTRUCTIONS_FROM_FY27,
  COLUMNS,
  SHEETS,
  VARIANT_META,
  VARIANTS,
  WIDTHS,
  THEMES,
  SCENARIOS,
  MODELS,
  RESPONSES,
  responseForPrompt,
  ANSWER_MESSAGES,
  makeProposal,
  INSTRUCTIONS_DEFAULT,
  COMPACTION_MESSAGES,
  CONVERSATIONS,
  COMMANDS,
} from "./data.js";

const app = document.getElementById("app");
const overlayRoot = document.getElementById("overlay-root");
const params = new URLSearchParams(window.location.search);
const prefersDark = window.matchMedia("(prefers-color-scheme: dark)");
const SHOW_BAR = params.get("switcher") !== "0";

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
const clone = (value) => JSON.parse(JSON.stringify(value));
const escapeHtml = (value) =>
  String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[char]));
const icon = (name, className = "icon") =>
  `<svg class="${className}" aria-hidden="true" focusable="false"><use href="#i-${name}"></use></svg>`;
const isTextInput = (el) =>
  !!el && typeof el.matches === "function" && (el.matches("input, textarea, select") || el.isContentEditable);

function validChoice(value, choices, fallback) {
  return choices.includes(value) ? value : fallback;
}

function initialState() {
  return {
    // route
    variant: validChoice(params.get("variant"), VARIANTS, "hybrid"),
    scenario: validChoice(params.get("state"), SCENARIOS.map((s) => s.key), "empty"),
    width: WIDTHS.includes(Number(params.get("width"))) ? Number(params.get("width")) : 380,
    themeMode: validChoice(params.get("theme"), THEMES, "auto"),
    // surface
    view: "chat",
    messages: [],
    proposal: null,
    loading: false,
    offline: false,
    banner: null,
    snackbar: null,
    announcement: "",
    // composer / model
    composerText: "",
    model: "gpt-5.2",
    models: clone(MODELS),
    modelMenuOpen: false,
    modelsLoading: false,
    settingsModelOpen: false,
    workbookMenuOpen: false,
    rowMenuId: null,
    // history
    conversations: clone(CONVERSATIONS),
    currentId: "c-current",
    historyScope: "workbook",
    // settings
    instructions: INSTRUCTIONS_DEFAULT,
    instructionsState: "saved",
    autoOpen: false,
    // mock workbook
    sheets: clone(SHEETS),
    selection: { sheet: "Forecast", range: "B22:F25", active: "B22" },
    excelStatus: "Ready",
    // interaction plumbing
    dialog: null,
    live: null,
    pendingAutoStream: false,
    focusAfter: null,
    scrollBottom: false,
    scrollToCard: false,
    scrollViewTo: null,
    isScrolled: false,
    saveTimer: null,
    snackbarTimer: null,
    streamTimer: null,
    applyTimer: null,
    lastPrompt: "",
    modelUpdatedAt: "just now",
    // #16 safety and persistence surfaces
    workbook: clone(WORKBOOK),
    unsaved: false,
    historyRows: clone(HISTORY_WORKBOOKS),
    historyLimitMb: HISTORY_LIMIT_MB,
    linkFilter: "",
    linkFrom: "history",
    lastLink: null,
    compaction: null,
    storeFailure: false,
    historyNotice: null,
    recoveredId: null,
    instructionsSource: null,
    readOnly: false,
    workbookReadOnly: false,
    readOnlyJustSaved: false,
  };
}

let state = initialState();
let scrollMemory = null;
let viewScrollMemory = 0;
let lastView = null;

function workbook() {
  return state.workbook;
}

function isReadOnly() {
  return state.readOnly || state.workbookReadOnly;
}

function historyUsedMb() {
  return state.historyRows.reduce((total, row) => total + row.sizeMb, 0);
}

function historyPercent() {
  return Math.round((historyUsedMb() / state.historyLimitMb) * 100);
}

function resolveTheme() {
  if (state.themeMode === "auto") return prefersDark.matches ? "dark" : "light";
  return state.themeMode;
}

function syncUrl() {
  const next = new URLSearchParams();
  next.set("variant", state.variant);
  next.set("state", state.scenario);
  next.set("width", String(state.width));
  next.set("theme", state.themeMode);
  window.history.replaceState(null, "", `?${next.toString()}`);
}

function announce(text) {
  state.announcement = text;
  const region = $("#liveRegion");
  if (region) region.textContent = text;
}

/* -------------------------------------------------------------------------
   Mock workbook
   ------------------------------------------------------------------------- */

function sheetData(name = state.selection.sheet) {
  return state.sheets[name] || state.sheets.Forecast;
}

function cellAddress(colIndex, rowIndex) {
  return `${COLUMNS[colIndex]}${rowIndex + 1}`;
}

function parseRange(range) {
  const value = String(range || "").trim().replace(/\$/g, "").toUpperCase();
  const match = value.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!match) return null;
  const col = (letters) => {
    let n = 0;
    for (const char of letters) n = n * 26 + (char.charCodeAt(0) - 64);
    return n - 1;
  };
  const startCol = col(match[1]);
  const startRow = Number(match[2]) - 1;
  const endCol = match[3] ? col(match[3]) : startCol;
  const endRow = match[4] ? Number(match[4]) - 1 : startRow;
  if ([startCol, startRow, endCol, endRow].some((n) => n < 0)) return null;
  return {
    startCol: Math.min(startCol, endCol),
    endCol: Math.max(startCol, endCol),
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
  };
}

function selectionBounds() {
  return parseRange(state.selection.range);
}

function formulaBarValue() {
  const { sheet, range } = state.selection;
  const data = sheetData(sheet);
  const single = parseRange(range);
  if (!single || single.startRow !== single.endRow || single.startCol !== single.endCol) {
    return `${range} · multiple cells selected`;
  }
  const row = data.rows[single.startRow] || [];
  const value = row[single.startCol] ?? "";
  return value === "" ? "blank" : value;
}

function rangeLabel(sheet, range) {
  if (!range) return sheet;
  if (sheet === state.selection.sheet) return range;
  return `${sheet}!${range}`;
}

function navigateToSelection(sheet, range, { announceIt = true } = {}) {
  if (!state.sheets[sheet]) {
    showSnackbar(`“${sheet}” isn’t in this workbook.`, { tone: "warn" });
    return;
  }
  const parsed = parseRange(range);
  if (!parsed) {
    showSnackbar(`“${range}” isn’t a range I can read.`, { tone: "warn" });
    return;
  }
  state.selection = {
    sheet,
    range,
    active: `${COLUMNS[parsed.startCol]}${parsed.startRow + 1}`,
  };
  state.excelStatus = `Selected ${sheet}!${range}`;
  refreshGridDom();
  if (announceIt) {
    showSnackbar(`Selected ${sheet}!${range} in the grid.`, { tone: "neutral" });
    announce(`Selected ${sheet}, ${range}`);
  }
  requestAnimationFrame(() => {
    const active = $(".cell.is-active");
    if (active) active.scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
  });
}

function refreshGridDom() {
  const mount = $("#gridMount");
  if (mount) mount.innerHTML = renderSheetView();
  const formula = $("#formulaValue");
  if (formula) formula.textContent = formulaBarValue();
  const name = $("#nameBox");
  if (name) name.textContent = state.selection.range;
  const status = $("#excelStatusText");
  if (status) status.textContent = state.excelStatus;
}

function renderSheetView() {
  const data = sheetData();
  const bounds = selectionBounds();
  const selectedSheet = state.selection.sheet;
  const rows = data.rows;
  const header = rows[0] || [];
  const body = rows.map((row, rowIndex) => {
    const cells = COLUMNS.map((col, colIndex) => {
      const value = row[colIndex] ?? "";
      const selected = selectedSheet === state.selection.sheet && bounds &&
        rowIndex >= bounds.startRow && rowIndex <= bounds.endRow &&
        colIndex >= bounds.startCol && colIndex <= bounds.endCol;
      const active = selected &&
        rowIndex === bounds.startRow && colIndex === bounds.startCol;
      const numeric = /^[-+$€£]?[\d,.]+%?$/.test(String(value).trim());
      return `<td class="cell${selected ? " is-selected" : ""}${active ? " is-active" : ""}${numeric ? " is-num" : ""}"
        data-cell="${col}${rowIndex + 1}" data-sheet="${escapeHtml(state.selection.sheet)}"
        tabindex="-1" role="gridcell" aria-selected="${selected ? "true" : "false"}"
      >${escapeHtml(value)}</td>`;
    }).join("");
    return `<tr>${`<th class="rowhead" scope="row">${rowIndex + 1}</th>`}${cells}</tr>`;
  }).join("");
  return `
    <div class="sheet-view">
      <div class="grid-scroll" id="gridScroll">
        <table class="sheet-grid" role="grid" aria-label="${escapeHtml(state.selection.sheet)} sheet">
          <thead>
            <tr>
              <th class="corner" aria-hidden="true"></th>
              ${COLUMNS.map((col) => `<th class="colhead" scope="col">${col}</th>`).join("")}
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>
      <div class="sheet-tabbar" aria-label="Sheets">
        <button class="sheet-nav" type="button" aria-label="First sheet">${icon("chevron-left")}</button>
        ${Object.keys(state.sheets).map((name) => `
          <button class="sheet-tab${name === state.selection.sheet ? " is-active" : ""}"
            type="button" data-action="select-sheet" data-sheet="${escapeHtml(name)}"
            aria-pressed="${name === state.selection.sheet}">${escapeHtml(state.sheets[name].label || name)}</button>
        `).join("")}
        <button class="sheet-add" type="button" aria-label="Add a sheet">${icon("plus")}</button>
        <span class="sheet-tab-spacer"></span>
        <span class="sheet-tab-hint">Mock workbook · not real Excel data</span>
      </div>
    </div>`;
}

function renderExcelChrome() {
  const theme = resolveTheme();
  return `
    <div class="excel-shell" data-theme="${theme}">
      <div class="excel-titlebar">
        <span class="excel-logo" aria-hidden="true"><span></span><span></span><span></span><span></span></span>
        <div class="excel-title">${escapeHtml(workbook().name)} <span>— ${escapeHtml(workbook().modified)}</span></div>
        <div class="excel-search" aria-hidden="true">${icon("search")}<span>Search</span></div>
        <div class="excel-account" aria-hidden="true"><span class="excel-avatar">ME</span><span class="excel-name">Mustapha</span></div>
      </div>
      <div class="excel-ribbon">
        <div class="excel-tabstrip" role="tablist" aria-label="Mock Excel ribbon">
          <span class="excel-filetab">File</span>
          <span role="tab" aria-selected="false">Home</span>
          <span role="tab" aria-selected="false">Insert</span>
          <span role="tab" aria-selected="false">Draw</span>
          <span role="tab" aria-selected="false">Page Layout</span>
          <span role="tab" aria-selected="false">Formulas</span>
          <span role="tab" aria-selected="false">Data</span>
          <span role="tab" aria-selected="false">Review</span>
          <span role="tab" aria-selected="false">View</span>
          <span role="tab" aria-selected="false">Automate</span>
          <span class="excel-agent-tab" role="tab" aria-selected="true">${icon("grid")} Agent</span>
        </div>
        <div class="excel-commandstrip">
          <button class="ribbon-action" type="button" data-action="focus-composer">${icon("grid")}<span>Open Agent</span></button>
          <button class="ribbon-action" type="button" data-action="suggest" data-prompt="Explain the selected cells">${icon("pen")}<span>Ask about selection</span></button>
          <button class="ribbon-action" type="button" data-action="suggest" data-prompt="Summarise the Forecast sheet">${icon("table")}<span>Summarise sheet</span></button>
          <span class="ribbon-divider"></span>
          <span class="ribbon-note">${icon("info")} Host chrome is a mock. The pane is the prototype.</span>
        </div>
      </div>
      <div class="excel-formula">
        <div class="name-box" id="nameBox">${escapeHtml(state.selection.range)}</div>
        <div class="fx-mark" aria-hidden="true">fx</div>
        <div class="formula-input"><span id="formulaValue">${escapeHtml(formulaBarValue())}</span></div>
        <div class="formula-commit" aria-hidden="true">✓</div>
      </div>
      <div class="excel-workspace" role="main" aria-label="Mock Excel window and Excel Agent prototype">
        <section class="sheet-column" aria-label="Mock worksheet">
          <div id="gridMount">${renderSheetView()}</div>
        </section>
        <aside class="pane-shell" id="paneShell" style="width:${state.width}px" aria-label="Excel Agent pane">
          <div class="pane-hostbar">
            <div class="pane-host-title">${icon("grid")}<span>Excel Agent</span><span class="proto-tag">PROTOTYPE</span></div>
            <span class="pane-host-note">docked right · 1 pane per workbook</span>
          </div>
          <div class="pane-app" id="paneApp" aria-label="Excel Agent prototype">${renderPane()}</div>
        </aside>
      </div>
      <div class="excel-statusbar">
        <div class="excel-status-left"><span class="status-ready-dot"></span><span id="excelStatusText">${escapeHtml(state.excelStatus)}</span></div>
        <div class="excel-status-right">
          <span>${escapeHtml(state.selection.sheet)}</span>
          <span>Count: ${sheetData().rows.length}</span>
          <span class="zoom-track"><i></i></span>
          <span>100%</span>
        </div>
      </div>
    </div>`;
}

/* -------------------------------------------------------------------------
   Markdown
   ------------------------------------------------------------------------- */

function renderCitationButton(label, target) {
  const cleanTarget = String(target || "").replace(/^cite:/, "");
  const aria = cleanTarget.replace("!", ", ").replace(":", " to ");
  return `<button class="citation" type="button" data-action="citation" data-cite="${escapeHtml(cleanTarget)}"
    aria-label="Go to ${escapeHtml(aria)}" title="Go to ${escapeHtml(cleanTarget)}">${escapeHtml(label)}</button>`;
}

function inlineMarkdown(text) {
  const slots = [];
  const stash = (html) => `\u0000${slots.push(html) - 1}\u0000`;
  let value = escapeHtml(text);
  value = value.replace(/`([^`]+)`/g, (_, code) => stash(`<code>${code}</code>`));
  value = value.replace(/\[([^\]]+)\]\(cite:([^)]+)\)/g, (_, label, target) =>
    stash(renderCitationButton(label, target)));
  value = value.replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, (_, label, url) =>
    stash(`<a class="md-link" href="${url}" target="_blank" rel="noopener noreferrer">${label}${icon("external", "icon icon-12")}</a>`));
  value = value.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  value = value.replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, "$1<em>$2</em>");
  value = value.replace(/~~([^~]+)~~/g, "<del>$1</del>");
  value = value.replace(/\n/g, "<br>");
  value = value.replace(/\u0000(\d+)\u0000/g, (_, index) => slots[Number(index)]);
  return value;
}

function splitTableRow(line) {
  return line.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((cell) => cell.trim());
}

function markdownToHtml(source, { partial = false } = {}) {
  let text = String(source ?? "").replace(/\r\n/g, "\n");
  if (!text.trim()) return "";
  if (partial) {
    // Keep partial streaming text parseable; chunks are block-complete, but this
    // guards against a half-written emphasis marker at the tail.
    const stars = (text.match(/\*\*/g) || []).length;
    if (stars % 2 === 1) text = text.slice(0, text.lastIndexOf("**"));
    const ticks = (text.match(/`/g) || []).length;
    if (ticks % 2 === 1) text = text.slice(0, text.lastIndexOf("`"));
  }
  const lines = text.split("\n");
  const out = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) { i += 1; continue; }

    if (/^```/.test(line)) {
      const language = line.replace(/^```/, "").trim() || "formula";
      const code = [];
      i += 1;
      while (i < lines.length && !/^```/.test(lines[i])) { code.push(lines[i]); i += 1; }
      i += 1;
      const raw = code.join("\n");
      out.push(`<div class="code-block"><div class="code-head"><span>${escapeHtml(language)}</span><button class="copy-code" type="button" data-action="copy-code" data-copy="${escapeHtml(raw)}">Copy</button></div><pre><code>${escapeHtml(raw)}</code></pre></div>`);
      continue;
    }

    if (/^###\s+/.test(line)) {
      out.push(`<h3>${inlineMarkdown(line.replace(/^###\s+/, ""))}</h3>`);
      i += 1;
      continue;
    }
    if (/^##\s+/.test(line)) {
      out.push(`<h2>${inlineMarkdown(line.replace(/^##\s+/, ""))}</h2>`);
      i += 1;
      continue;
    }

    if (/^>\s?/.test(line)) {
      const quote = [];
      while (i < lines.length && /^>\s?/.test(lines[i])) {
        quote.push(lines[i].replace(/^>\s?/, ""));
        i += 1;
      }
      out.push(`<blockquote>${inlineMarkdown(quote.join("\n"))}</blockquote>`);
      continue;
    }

    if (/^\s*[-*]\s+/.test(line)) {
      const items = [];
      while (i < lines.length && /^\s*[-*]\s+/.test(lines[i])) {
        items.push(`<li>${inlineMarkdown(lines[i].replace(/^\s*[-*]\s+/, ""))}</li>`);
        i += 1;
      }
      out.push(`<ul>${items.join("")}</ul>`);
      continue;
    }

    if (/^\s*\d+\.\s+/.test(line)) {
      const items = [];
      while (i < lines.length && /^\s*\d+\.\s+/.test(lines[i])) {
        items.push(`<li>${inlineMarkdown(lines[i].replace(/^\s*\d+\.\s+/, ""))}</li>`);
        i += 1;
      }
      out.push(`<ol>${items.join("")}</ol>`);
      continue;
    }

    if (
      line.includes("|") &&
      i + 1 < lines.length &&
      /^\s*\|?[\s:|-]+\|?[\s:|-]*$/.test(lines[i + 1]) &&
      lines[i + 1].includes("-")
    ) {
      const headerCells = splitTableRow(line);
      const separator = splitTableRow(lines[i + 1]);
      const align = separator.map((cell) => {
        const trimmed = cell.trim();
        if (/^:.*:$/.test(trimmed)) return "center";
        if (/:$/.test(trimmed)) return "right";
        return "left";
      });
      i += 2;
      const rows = [];
      while (i < lines.length && lines[i].includes("|") && lines[i].trim()) {
        rows.push(splitTableRow(lines[i]));
        i += 1;
      }
      const head = headerCells.map((cell, index) =>
        `<th scope="col" class="align-${align[index] || "left"}">${inlineMarkdown(cell)}</th>`).join("");
      const body = rows.map((cells) => `<tr>${headerCells.map((_, index) =>
        `<td class="align-${align[index] || "left"}">${inlineMarkdown(cells[index] ?? "")}</td>`).join("")}</tr>`).join("");
      out.push(`<div class="tablewrap"><table class="md-table"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`);
      continue;
    }

    if (/^---+$/.test(line.trim())) {
      out.push("<hr>");
      i += 1;
      continue;
    }

    const paragraph = [line];
    i += 1;
    while (
      i < lines.length &&
      lines[i].trim() &&
      !/^```/.test(lines[i]) &&
      !/^#{2,3}\s+/.test(lines[i]) &&
      !/^>\s?/.test(lines[i]) &&
      !/^\s*[-*]\s+/.test(lines[i]) &&
      !/^\s*\d+\.\s+/.test(lines[i]) &&
      !(lines[i].includes("|") && i + 1 < lines.length && /^\s*\|?[\s:|-]+\|?[\s:|-]*$/.test(lines[i + 1]))
    ) {
      paragraph.push(lines[i]);
      i += 1;
    }
    out.push(`<p>${inlineMarkdown(paragraph.join("\n"))}</p>`);
  }
  return out.join("");
}

/* -------------------------------------------------------------------------
   Pane shell
   ------------------------------------------------------------------------- */

function renderPane() {
  if (state.view === "link") return renderLinkView();
  if (state.view === "history") return renderHistoryView();
  if (state.view === "settings") return renderSettingsView();
  if (state.variant === "claude") return renderClaudeChat();
  if (state.variant === "fluent") return renderFluentChat();
  return renderHybridChat();
}

function renderLiveRegion() {
  return `<div id="liveRegion" class="sr-only" role="status" aria-live="polite">${escapeHtml(state.announcement)}</div>`;
}

function renderPrototypeBar() {
  if (!SHOW_BAR) return "";
  const meta = VARIANT_META[state.variant];
  const widthOptions = [...new Set([...WIDTHS, state.width])].sort((a, b) => a - b);
  return `
    <div class="proto-bar" role="region" aria-label="Prototype controls">
      <button class="proto-arrow" type="button" data-action="prev-variant" aria-label="Previous variant">${icon("chevron-left")}</button>
      <div class="proto-variant">
        <span class="proto-kicker">Variant ${meta.code}</span>
        <strong>${escapeHtml(meta.name)}</strong>
        <span class="proto-blurb">${escapeHtml(meta.blurb)}</span>
      </div>
      <button class="proto-arrow" type="button" data-action="next-variant" aria-label="Next variant">${icon("chevron-right")}</button>
      <span class="proto-divider" aria-hidden="true"></span>
      <label class="proto-control">Scenario
        <select id="scenarioSelect">
          ${SCENARIOS.map((item) => `<option value="${item.key}"${item.key === state.scenario ? " selected" : ""}>${escapeHtml(item.label)}</option>`).join("")}
        </select>
      </label>
      <label class="proto-control">Width
        <select id="widthSelect">
          ${widthOptions.map((width) => `<option value="${width}"${width === state.width ? " selected" : ""}>${width} px</option>`).join("")}
        </select>
      </label>
      <div class="proto-theme" role="group" aria-label="Pane theme">
        ${THEMES.map((theme) => `<button type="button" class="${state.themeMode === theme ? "is-active" : ""}" data-action="set-theme" data-theme="${theme}">${theme[0].toUpperCase() + theme.slice(1)}</button>`).join("")}
      </div>
      <button class="proto-reset" type="button" data-action="reset-prototype" title="Reset this scenario">Reset</button>
    </div>`;
}

function renderSnackbarContent() {
  if (!state.snackbar) return "";
  const snack = state.snackbar;
  return `<div class="snackbar snackbar-${snack.tone || "neutral"}" role="status" aria-live="polite">
    <span class="snackbar-text">${escapeHtml(snack.text)}</span>
    ${snack.action ? `<button class="snackbar-action" type="button" data-action="snackbar-action">${escapeHtml(snack.action.label)}</button>` : ""}
    <button class="snackbar-close" type="button" data-action="dismiss-snackbar" aria-label="Dismiss">${icon("x")}</button>
  </div>`;
}

function renderSnackbarSlot() {
  return `<div class="snackbar-slot" id="snackbarSlot">${renderSnackbarContent()}</div>`;
}



/* -------------------------------------------------------------------------
   Chat surfaces
   ------------------------------------------------------------------------- */

function isStreaming() {
  return state.messages.some((message) => message.role === "assistant" && message.streaming);
}

function proposalGatesComposer(proposal = state.proposal) {
  return !!proposal && ["proposed", "applying", "partial"].includes(proposal.status);
}

function composerDisabled() {
  return state.offline || state.loading || proposalGatesComposer();
}

function composerPlaceholder() {
  if (state.loading) return "Starting Excel Agent…";
  if (state.offline) return "Reconnect to continue…";
  if (state.proposal?.status === "applying") return "Applying changes…";
  if (proposalGatesComposer()) return state.proposal.status === "partial"
    ? "Retry or dismiss the remaining changes"
    : "Review the proposed changes above";
  if (isReadOnly()) return "Read-only mode — questions only";
  if (isStreaming()) return "Ask a follow-up…";
  return state.messages.length ? "Ask a follow-up…" : "Ask about this workbook…";
}

function composerHint() {
  if (state.offline) return { icon: "alert", text: "Reconnect to continue. The transcript stays readable." };
  if (state.proposal?.status === "applying") return { icon: "info", text: "Applying changes. Keep the pane open." };
  if (state.proposal?.status === "proposed") return { icon: "info", text: "Apply or discard to keep chatting." };
  if (state.proposal?.status === "partial") return { icon: "info", text: "Retry or dismiss the remaining changes." };
  if (isStreaming()) return { icon: "info", text: "Enter does nothing while a response is streaming. Esc stops it." };
  if (isReadOnly()) return { icon: "lock", text: "Read-only mode. Edits are unavailable; questions still work." };
  return null;
}

function renderNewChatButton() {
  if (!state.messages.length && !state.proposal) return "";
  const disabled = isStreaming() || proposalGatesComposer() || state.proposal?.status === "applying";
  const reason = disabled
    ? isStreaming()
      ? "Stop the current response first"
      : "Apply or discard the pending changes first"
    : "New chat";
  return `<button class="icon-button" type="button" data-action="new-chat" aria-label="${escapeHtml(reason)}" title="${escapeHtml(reason)}"${disabled ? " disabled" : ""}>${icon("plus")}</button>`;
}

function renderReadOnlyChip() {
  if (!isReadOnly()) return "";
  const forced = state.workbookReadOnly;
  const title = forced
    ? "This workbook is open read-only. Questions only."
    : "Read-only mode is on. Questions only. Open Settings to change it.";
  return `<button class="readonly-chip" type="button" data-action="open-readonly-settings" title="${escapeHtml(title)}" aria-label="${escapeHtml(title)}">
    ${icon("lock")}<span>Read-only</span>
  </button>`;
}

function renderHeaderActions() {
  return `
    <div class="header-actions">
      ${renderReadOnlyChip()}
      <button class="icon-button" type="button" data-action="open-history" aria-label="History" title="History">${icon("clock")}</button>
      ${renderNewChatButton()}
      <button class="icon-button" type="button" data-action="open-settings" aria-label="Settings" title="Settings">${icon("gear")}</button>
    </div>`;
}

function renderWorkbookButton() {
  return `<button class="workbook-button" type="button" data-action="toggle-workbook-menu" aria-haspopup="menu" aria-expanded="${state.workbookMenuOpen}">
    ${icon("file", "icon file-icon")}
    <span class="workbook-name" title="${escapeHtml(workbook().path)}">${escapeHtml(workbook().name)}</span>
    ${state.unsaved ? '<span class="unsaved-dot" title="Unsaved changes" aria-label="Unsaved changes"></span>' : ""}
    ${icon("chevron-down", "icon chevron")}
  </button>`;
}

function renderHybridChat() {
  return `
    <div class="pane-view pane-view-hybrid" data-view="chat">
      <header class="pane-header${state.isScrolled ? " is-scrolled" : ""}">
        ${renderWorkbookButton()}
        ${renderHeaderActions()}
      </header>
      ${renderWorkbookMenu()}
      ${renderBudgetBanner()}
      ${renderBanner()}
      <div class="transcript" id="transcript" role="region" aria-label="Conversation" tabindex="0">
        <div class="transcript-inner">${renderTimeline()}</div>
      </div>
      ${renderSnackbarSlot()}
      ${renderCompactionNotice()}
      ${renderComposer()}
    </div>`;
}

function renderClaudeChat() {
  return `
    <div class="pane-view pane-view-claude" data-view="chat">
      <div class="claude-topbar">
        ${renderWorkbookButton()}
        <div class="header-actions">
          ${renderReadOnlyChip()}
          <button class="icon-button" type="button" data-action="open-history" aria-label="History" title="History">${icon("clock")}</button>
          ${renderNewChatButton()}
          <button class="icon-button" type="button" data-action="open-settings" aria-label="Settings" title="Settings">${icon("gear")}</button>
        </div>
      </div>
      ${renderWorkbookMenu()}
      ${renderBudgetBanner()}
      ${renderBanner()}
      <div class="transcript claude-transcript" id="transcript" role="region" aria-label="Conversation" tabindex="0">
        <div class="transcript-inner">${renderTimeline()}</div>
      </div>
      ${renderSnackbarSlot()}
      <div class="claude-composer-wrap">${renderCompactionNotice()}${renderComposer()}</div>
    </div>`;
}

function renderFluentChat() {
  const reviewCount = state.proposal
    ? state.proposal.rows.filter((row) => row.included).length
    : 0;
  return `
    <div class="pane-view pane-view-fluent" data-view="chat">
      <div class="fluent-header">
        <div class="fluent-title-row">
          <span class="fluent-mark">${icon("grid")}</span>
          <div class="fluent-title-copy">
            <button class="fluent-workbook" type="button" data-action="toggle-workbook-menu" aria-haspopup="menu" aria-expanded="${state.workbookMenuOpen}">
              ${escapeHtml(workbook().name)} ${icon("chevron-down", "icon icon-12")}
            </button>
            <div class="fluent-subtitle"><span class="connection-dot"></span> ${state.offline ? "Disconnected" : "Connected"} · local agent</div>
          </div>
          <button class="icon-button fluent-more" type="button" data-action="toggle-workbook-menu" aria-label="Workbook menu">${icon("more")}</button>
        </div>
        <div class="fluent-commandbar" role="toolbar" aria-label="Pane commands">
          <button class="fluent-command" type="button" data-action="new-chat"${isStreaming() || proposalGatesComposer() || (!state.messages.length && !state.proposal) ? " disabled" : ""}>${icon("plus")}<span>New chat</span></button>
          <button class="fluent-command" type="button" data-action="open-history">${icon("clock")}<span>History</span></button>
          ${state.proposal ? `<button class="fluent-command${state.proposal.status === "proposed" ? " is-attention" : ""}" type="button" data-action="review-card">${icon("pen")}<span>Review ${reviewCount}</span></button>` : ""}
          <button class="fluent-command" type="button" data-action="open-settings">${icon("gear")}<span>Settings</span></button>
        </div>
      </div>
      ${renderWorkbookMenu()}
      ${renderBudgetBanner()}
      ${renderBanner()}
      <div class="transcript fluent-transcript" id="transcript" role="region" aria-label="Conversation" tabindex="0">
        ${isStreaming() ? '<div class="fluent-progress" aria-hidden="true"><span></span></div>' : ""}
        <div class="transcript-inner">${renderTimeline()}</div>
      </div>
      ${renderSnackbarSlot()}
      ${renderCompactionNotice()}
      ${renderComposer(true)}
    </div>`;
}

function renderWorkbookMenu() {
  if (!state.workbookMenuOpen) return "";
  return `
    <div class="menu-popover workbook-popover" role="menu">
      <div class="menu-workbook">
        ${icon("file", "icon file-icon")}
        <div class="menu-workbook-copy">
          <strong>${escapeHtml(workbook().name)}</strong>
          <span title="${escapeHtml(workbook().path)}">${escapeHtml(workbook().path || "Not saved yet")}</span>
        </div>
      </div>
      <div class="menu-separator"></div>
      <button class="menu-row" type="button" data-action="copy-path" role="menuitem">${icon("copy", "icon menu-icon")}<span>Copy full path</span></button>
      <button class="menu-row" type="button" data-action="open-instructions" role="menuitem">${icon("pen", "icon menu-icon")}<span>Instructions for this workbook</span></button>
      <button class="menu-row" type="button" data-action="open-link" data-from="menu" role="menuitem" ${state.unsaved ? 'disabled title="Save the workbook first"' : ""}>
        ${icon("link", "icon menu-icon")}<span>Link to an old conversation…</span>
        ${state.unsaved ? `<span class="menu-hint">Save first</span>` : ""}
      </button>
      <button class="menu-row" type="button" data-action="toggle-auto-open" role="menuitemcheckbox" aria-checked="${state.autoOpen}">
        ${icon("check", `icon menu-icon${state.autoOpen ? "" : " is-hidden"}`)}<span>Open with this workbook</span>
        <span class="menu-hint">${state.autoOpen ? "On" : "Off"}</span>
      </button>
      <div class="menu-foot">One pane is bound to this workbook. The pane cannot follow another file.</div>
    </div>`;
}

function renderBanner() {
  if (!state.banner) return "";
  return `
    <div class="global-banner banner-${state.banner.tone || "warn"}" role="status">
      ${icon(state.banner.tone === "danger" ? "alert" : "alert")}
      <div class="banner-copy"><strong>${escapeHtml(state.banner.title)}</strong>${state.banner.text ? ` ${escapeHtml(state.banner.text)}` : ""}</div>
      ${state.banner.action ? `<button class="banner-action" type="button" data-action="${state.banner.action}">${escapeHtml(state.banner.actionLabel)}</button>` : ""}
    </div>`;
}

function renderBudgetBanner() {
  const pct = historyPercent();
  if (pct < 80) return "";
  const over = pct >= 100;
  const title = over
    ? `History is over its ${state.historyLimitMb} MB budget.`
    : `History is ${pct}% of its ${state.historyLimitMb} MB budget.`;
  return `<div class="global-banner banner-${over ? "danger" : "warn"}" role="status">
    ${icon(over ? "alert" : "info")}
    <div class="banner-copy"><strong>${escapeHtml(title)}</strong>${over ? " Nothing is deleted automatically." : ""}</div>
    <button class="banner-action" type="button" data-action="open-settings-history">Manage history</button>
  </div>`;
}

function renderCompactionNotice() {
  const compaction = state.compaction;
  if (!compaction || !compaction.notice || state.view !== "chat") return "";
  if (compaction.summarizing) {
    return `<section class="compaction-notice is-summarizing" aria-label="Summarizing earlier turns" role="status">
      ${icon("spark")}
      <div class="compaction-copy">
        <strong>Summarizing ${compaction.pending?.covered || 0} turns…</strong>
        <span>Earlier turns stay in the transcript; this only changes what the model sees.</span>
        <div class="compaction-progress" role="progressbar" aria-label="Summarizing earlier turns" aria-valuetext="In progress"><span></span></div>
      </div>
    </section>`;
  }
  return `<section class="compaction-notice" aria-label="Context compaction" role="status">
    ${icon("spark")}
    <div class="compaction-copy">
      <strong>This chat is ${compaction.pct}% of the model’s context.</strong>
      <span>Summarize earlier turns or start a new chat.</span>
    </div>
    <div class="compaction-actions">
      <button class="button button-secondary button-small" type="button" data-action="summarize">Summarize earlier turns</button>
      <button class="text-action compaction-new" type="button" data-action="new-chat">Start a new chat</button>
    </div>
  </section>`;
}

function renderCompactionDivider(cut, index) {
  return `<div class="compaction-divider" role="separator" aria-label="Earlier ${cut.covered} turns summarized">
    <span class="compaction-rule" aria-hidden="true"></span>
    <div class="compaction-divider-copy">
      <span>Earlier ${cut.covered} turns summarized</span>
      <button class="text-action" type="button" data-action="toggle-original" data-cut="${index}" aria-expanded="${!!cut.expanded}">${cut.expanded ? "Hide original" : "View original"}</button>
    </div>
    <span class="compaction-rule" aria-hidden="true"></span>
  </div>`;
}

function renderTimeline() {
  if (state.loading) return renderSkeleton();
  const timeline = [];
  const cuts = state.compaction?.cuts || [];
  let index = 0;
  while (index < state.messages.length) {
    const cutIndex = cuts.findIndex((cut) => cut.start === index);
    if (cutIndex !== -1) {
      const cut = cuts[cutIndex];
      timeline.push(renderCompactionDivider(cut, cutIndex));
      if (cut.expanded) {
        for (let original = cut.start; original < cut.end; original += 1) {
          timeline.push(renderMessage(state.messages[original], original, { original: true }));
        }
      }
      index = cut.end;
      continue;
    }
    timeline.push(renderMessage(state.messages[index], index));
    index += 1;
  }
  if (state.proposal) timeline.push(String(renderProposal(state.proposal)));
  if (!timeline.length) timeline.push(renderEmptyState());
  return timeline.join("");
}

function renderSkeleton() {
  return `
    <div class="skeleton-state" aria-label="Loading conversation" role="status">
      <div class="skeleton-line w-40"></div>
      <div class="skeleton-line w-80"></div>
      <div class="skeleton-line w-70"></div>
      <div class="skeleton-gap"></div>
      <div class="skeleton-line w-55"></div>
      <div class="skeleton-line w-90"></div>
      <div class="skeleton-line w-65"></div>
      <div class="skeleton-box"></div>
      <div class="skeleton-line w-45"></div>
    </div>`;
}

function renderEmptyState() {
  return `
    <div class="empty-state empty-${state.variant}">
      <div class="empty-mark">${icon("grid")}</div>
      ${state.variant === "fluent" ? '<div class="empty-eyebrow">Excel Agent</div>' : ""}
      <h1>What are we working on?</h1>
      <p>I can read this workbook, answer questions, and make changes — with your approval.</p>
      <div class="suggestions">
        ${COMMANDS.map((command) => `
          <button class="suggestion" type="button" data-action="suggest" data-prompt="${escapeHtml(command.prompt)}">
            ${icon(command.icon, "icon suggestion-icon")}
            <span>${escapeHtml(command.label)}</span>
            ${icon("chevron-right", "icon suggestion-go")}
          </button>`).join("")}
      </div>
      <div class="trust-line">${icon("shield")}<span>Stored locally · Nothing is written without approval</span></div>
    </div>`;
}

function renderMessage(message, index, options = {}) {
  if (message.role === "user") return renderUserMessage(message);
  if (message.role === "error") return renderErrorCard(message, index);
  return renderAssistantMessage(message, index, options);
}

function renderUserMessage(message) {
  const body = escapeHtml(message.text).replace(/\n/g, "<br>");
  if (state.variant === "fluent") {
    return `<div class="msg-row msg-user-row">
      <span class="message-avatar user-avatar" aria-hidden="true">ME</span>
      <div class="message-body">
        <div class="message-meta"><span>You</span><span class="message-time">now</span></div>
        <div class="user-bubble">${body}</div>
      </div>
    </div>`;
  }
  if (state.variant === "claude") {
    return `<div class="claude-user-row"><div class="user-bubble">${body}</div></div>`;
  }
  return `<div class="user-turn"><div class="user-bubble">${body}</div></div>`;
}

function renderAssistantMessage(message, index, options = {}) {
  const live = !!message.streaming;
  const markdown = markdownToHtml(message.text || "", { partial: live });
  const body = `${markdown}${live ? '<span class="caret" aria-hidden="true"></span>' : ""}`;
  const status = renderStatusLine(message, index, live);
  const steps = message.showSteps ? renderStepList(message) : "";
  const actions = !live && !options.original ? renderMessageActions(message, index) : "";
  const stopped = message.stopped ? `<div class="stopped-line">${icon("square")} Stopped. The partial answer is kept.</div>` : "";
  if (state.variant === "fluent") {
    return `<div class="msg-row msg-assistant-row${options.original ? " is-original" : ""}">
      <span class="message-avatar agent-avatar" aria-hidden="true">${icon("grid")}</span>
      <div class="message-body">
        <div class="message-meta"><span>Excel Agent</span><span class="message-time">${live ? "working" : options.original ? "original" : "now"}</span></div>
        ${status}${steps}
        <div class="markdown"${live ? ' id="liveBody"' : ""}>${body}</div>
        ${stopped}${actions}
      </div>
    </div>`;
  }
  const className = state.variant === "claude" ? "assistant-turn claude-assistant" : "assistant-turn";
  return `<div class="${className}${options.original ? " is-original" : ""}">
    ${status}${steps}
    <div class="markdown"${live ? ' id="liveBody"' : ""}>${body}</div>
    ${stopped}${actions}
  </div>`;
}

function renderStatusLine(message, index, live) {
  if (!message.status) return "";
  const status = message.status;
  const isDone = status.state === "done";
  const isLive = status.state === "live";
  const label = status.label || "Thinking…";
  const step = status.step ? `<span class="status-step">${escapeHtml(status.step)}</span>` : "";
  const dot = `<span class="status-dot${isDone ? " is-done" : isLive ? " is-live" : ""}" aria-hidden="true"></span>`;
  if (isDone && message.steps?.length) {
    return `<button class="status-line status-done" type="button" data-action="toggle-steps" data-index="${index}" aria-expanded="${!!message.showSteps}">
      ${dot}<span class="status-label">${escapeHtml(label)}</span>${icon("chevron-down", "icon status-chevron")}
    </button>`;
  }
  return `<div class="status-line${isLive ? " is-live" : ""}"${live ? ' id="liveStatus"' : ""}>
    ${dot}<span class="status-label">${escapeHtml(label)}</span>${step}
  </div>`;
}

function renderStepList(message) {
  if (!message.steps?.length) return "";
  return `<ol class="step-list" aria-label="Tool steps">
    ${message.steps.map((step, index) => `<li><span class="step-number">${index + 1}</span><span>${escapeHtml(step)}</span></li>`).join("")}
  </ol>`;
}

function renderMessageActions(message, index) {
  return `<div class="message-actions">
    <button class="text-action" type="button" data-action="copy-message" data-index="${index}">${icon("copy")}<span>Copy</span></button>
    <button class="text-action" type="button" data-action="retry-message" data-index="${index}">${icon("retry")}<span>Retry</span></button>
    ${message.steps?.length ? `<button class="text-action" type="button" data-action="toggle-steps" data-index="${index}">${icon("chevron-down", "icon action-chevron")}<span>Steps</span></button>` : ""}
  </div>`;
}

function renderErrorCard(message, index) {
  return `<div class="error-card" role="alert">
    <div class="error-head">
      <span class="error-glyph">${icon("alert")}</span>
      <div class="error-copy">
        <div class="error-title">${escapeHtml(message.title || "The model didn’t answer in time.")}</div>
        <div class="error-text">${escapeHtml(message.text || "Your workbook is unchanged. This is usually a gateway hiccup — try again when the banner clears.")}</div>
      </div>
    </div>
    <div class="error-actions">
      <button class="button button-primary button-small" type="button" data-action="retry-error">Retry</button>
      <button class="button button-ghost button-small" type="button" data-action="copy-details" data-copy="${escapeHtml(message.detail || "")}">${icon("copy")} Copy details</button>
    </div>
    <div class="error-detail">${escapeHtml(message.detail || "Request timed out after 90 s · 22:14:06 · req_8f3a")}</div>
  </div>`;
}

function renderComposer(fluent = false) {
  const disabled = composerDisabled();
  const streaming = isStreaming();
  const empty = !state.composerText.trim();
  const hint = composerHint();
  const modelMenu = renderModelMenu();
  if (state.loading) {
    return `<div class="composer-wrap composer-wrap-loading">
      <div class="composer-skeleton"><span></span><span></span></div>
    </div>`;
  }
  if (fluent) {
    return `<div class="fluent-composer-wrap">
      ${modelMenu}
      <div class="fluent-composer${disabled ? " is-disabled" : ""}">
        <textarea id="composer" class="composer-input" rows="1" placeholder="${escapeHtml(composerPlaceholder())}" ${disabled ? "disabled" : ""} aria-label="Message Excel Agent">${escapeHtml(state.composerText)}</textarea>
        <div class="fluent-composer-footer">
          <button class="model-chip fluent-model" type="button" data-action="toggle-model-menu" aria-haspopup="listbox" aria-expanded="${state.modelMenuOpen}">
            <span class="model-name">${escapeHtml(state.model)}</span>${icon("chevron-down", "icon chevron")}
          </button>
          <span class="composer-spacer"></span>
          ${streaming
            ? `<button class="fluent-send is-stop" type="button" data-action="stop">${icon("square")} Stop</button>`
            : `<button class="fluent-send" type="button" data-action="send" ${disabled || empty ? "disabled" : ""}>${icon("arrow-up")} Send</button>`}
        </div>
      </div>
      ${hint ? `<div class="composer-hint">${icon(hint.icon)} ${escapeHtml(hint.text)}</div>` : ""}
    </div>`;
  }
  return `<div class="composer-wrap">
    ${modelMenu}
    <div class="composer${disabled ? " is-disabled" : ""}${streaming ? " is-streaming" : ""}">
      <textarea id="composer" class="composer-input" rows="1" placeholder="${escapeHtml(composerPlaceholder())}" ${disabled ? "disabled" : ""} aria-label="Message Excel Agent">${escapeHtml(state.composerText)}</textarea>
      <div class="composer-toolbar">
        <button class="model-chip" type="button" data-action="toggle-model-menu" aria-haspopup="listbox" aria-expanded="${state.modelMenuOpen}">
          ${icon("grid", "icon model-icon")}<span class="model-name">${escapeHtml(state.model)}</span>${icon("chevron-down", "icon chevron")}
        </button>
        <span class="composer-spacer"></span>
        ${streaming
          ? `<button class="send-button stop-button" type="button" data-action="stop" aria-label="Stop generating">${icon("square")}</button>`
          : `<button class="send-button" type="button" data-action="send" aria-label="Send message" ${disabled || empty ? "disabled" : ""}>${icon("arrow-up")}</button>`}
      </div>
    </div>
    ${hint ? `<div class="composer-hint">${icon(hint.icon)} ${escapeHtml(hint.text)}</div>` : ""}
  </div>`;
}

function renderModelMenu() {
  if (!state.modelMenuOpen) return "";
  return `<div class="menu-popover model-popover" role="listbox" aria-label="Model">
    <div class="model-menu-head"><span>Model</span><span class="model-source">From your gateway</span></div>
    ${state.models.map((model) => `<button class="menu-row model-row${model.name === state.model ? " is-selected" : ""}" type="button" role="option" aria-selected="${model.name === state.model}" data-action="select-model" data-model="${escapeHtml(model.name)}">
      <span class="model-check">${model.name === state.model ? icon("check") : ""}</span>
      <span>${escapeHtml(model.name)}</span>
      ${model.tag ? `<span class="model-tag">${escapeHtml(model.tag)}</span>` : ""}
    </button>`).join("")}
    <div class="menu-separator"></div>
    <button class="menu-row" type="button" data-action="refresh-models" ${state.modelsLoading ? "disabled" : ""}>
      ${icon("retry", "icon menu-icon")}<span>${state.modelsLoading ? "Refreshing…" : "Refresh list"}</span>
    </button>
    <div class="menu-foot">Used for new messages. The current turn keeps its model.</div>
  </div>`;
}

/* -------------------------------------------------------------------------
   Edit proposal
   ------------------------------------------------------------------------- */

function proposalRiskCounts(proposal = state.proposal) {
  const rows = proposal?.rows || [];
  return {
    overwrite: rows.filter((row) => row.risk === "overwrite").length,
    confirm: rows.filter((row) => row.risk === "confirm").length,
    loss: rows.filter((row) => row.risk === "loss").length,
  };
}

function selectedRows(proposal = state.proposal) {
  return (proposal?.rows || []).filter((row) => row.included);
}

function selectedCount(proposal = state.proposal) {
  return selectedRows(proposal).length;
}

function overwriteSelectedCells(proposal = state.proposal) {
  return selectedRows(proposal)
    .filter((row) => row.risk === "overwrite")
    .reduce((total, row) => total + (row.cells || 1), 0);
}

function sheetGroupCount(proposal = state.proposal) {
  return new Set((proposal?.rows || []).map((row) => row.sheet)).size;
}

function proposalTitle(proposal) {
  switch (proposal.status) {
    case "applying": return "Applying changes";
    case "applied": return `Applied ${proposal.appliedCount} changes`;
    case "partial": return `Applied ${proposal.appliedCount} of ${proposal.appliedCount + proposal.failedCount}`;
    case "dismissed": return `${proposal.failedCount || 2} changes still pending`;
    case "discarded": return "Changes discarded";
    default: return "Proposed changes";
  }
}

function proposalSubtitle(proposal) {
  const groups = sheetGroupCount(proposal);
  switch (proposal.status) {
    case "proposed":
      return `${proposal.rows.length} changes in ${groups} sheets · ${proposal.rows[0]?.sheet || "Forecast"}`;
    case "applying":
      return `${selectedCount(proposal)} of ${proposal.rows.length} selected · writing now`;
    case "applied":
      return `${proposal.excludedCount ? `${proposal.excludedCount} excluded · ` : ""}This apply is one Ctrl+Z step.`;
    case "partial":
      return `${proposal.failedCount} failed · retry or dismiss`;
    case "dismissed":
      return "The failed changes stay pending in History.";
    default:
      return "";
  }
}

function proposalGlyph(proposal) {
  if (proposal.status === "applied") return icon("check");
  if (proposal.status === "partial" || proposal.status === "dismissed") return icon("alert");
  return icon("pen");
}

function riskBadges(proposal) {
  if (proposal.status === "applied") {
    return `<span class="badge badge-success">${icon("check")} applied</span>${proposal.excludedCount ? `<span class="badge badge-muted">${proposal.excludedCount} excluded</span>` : ""}`;
  }
  if (proposal.status === "partial") {
    return `<span class="badge badge-danger">${icon("alert")} ${proposal.failedCount} failed</span><span class="badge badge-success">${proposal.appliedCount} applied</span>`;
  }
  if (proposal.status === "dismissed") {
    return `<span class="badge badge-warn">${icon("alert")} pending</span>`;
  }
  const risks = proposalRiskCounts(proposal);
  const badges = [];
  if (risks.overwrite) badges.push(`<span class="badge badge-warn">${icon("alert")} ${risks.overwrite} overwrite${risks.overwrite === 1 ? "" : "s"}</span>`);
  if (risks.confirm) badges.push(`<span class="badge badge-danger">${icon("alert")} ${risks.confirm} risky formula</span>`);
  if (risks.loss) badges.push(`<span class="badge badge-danger">${risks.loss} loss row</span>`);
  return badges.join("");
}

function renderProposal(proposal) {
  if (!proposal) return "";
  if (state.variant === "fluent") return renderFluentProposal(proposal);
  return renderWarmProposal(proposal);
}

function renderWarmProposal(proposal) {
  const collapsed = proposal.collapsed && !["applying", "partial"].includes(proposal.status);
  return `<section class="edit-card edit-${proposal.status}" id="editCard" aria-label="Edit proposal">
    <div class="edit-head">
      <span class="edit-glyph glyph-${proposal.status}">${proposalGlyph(proposal)}</span>
      <div class="edit-heading">
        <div class="edit-title">${escapeHtml(proposalTitle(proposal))}</div>
        <div class="edit-sub">${escapeHtml(proposalSubtitle(proposal))}</div>
        <div class="edit-badges">${riskBadges(proposal)}</div>
      </div>
      <button class="edit-chevron" type="button" data-action="toggle-card" aria-expanded="${!collapsed}" aria-label="${collapsed ? "Expand changes" : "Collapse changes"}">${icon(collapsed ? "chevron-down" : "chevron-up")}</button>
    </div>
    ${collapsed ? "" : `<div class="edit-body">${renderWarmProposalBody(proposal)}</div>${renderWarmProposalFooter(proposal)}`}
  </section>`;
}

function renderWarmProposalBody(proposal) {
  if (proposal.status === "applying") {
    return `<div class="apply-progress" role="progressbar" aria-label="Applying changes" aria-valuemin="0" aria-valuemax="${selectedCount(proposal)}">
      <span></span>
    </div>
    <div class="apply-caption">Applying ${selectedCount(proposal)} changes in one undo group… <span class="apply-hint">Do not edit the grid while Apply runs.</span></div>
    ${renderProposalGroups(proposal, proposal.rows, { locked: true, applying: true })}`;
  }
  if (proposal.status === "partial" || proposal.status === "dismissed") {
    const failed = proposal.rows.filter((row) => row.outcome === "failed");
    const applied = proposal.rows.filter((row) => row.outcome === "applied" || row.outcome === "excluded");
    return `
      <div class="notice notice-danger edit-alert">
        ${icon("alert")}
        <div><strong>Excel was still in cell-edit mode.</strong> ${proposal.failedCount} changes weren’t applied. Finish the edit, then retry.</div>
      </div>
      ${renderProposalGroups(proposal, failed, { locked: true, failed: true })}
      ${proposal.showApplied ? renderProposalGroups(proposal, applied, { locked: true, applied: true }) : ""}
      <button class="more-changes" type="button" data-action="toggle-applied">${icon(proposal.showApplied ? "chevron-up" : "chevron-down")} ${proposal.showApplied ? "Hide" : `Show ${proposal.appliedCount}`} applied changes</button>`;
  }
  if (proposal.status === "applied") {
    return `
      <div class="applied-note">${icon("check")} All selected ranges are written. ${proposal.excludedCount ? `${proposal.excludedCount} rows were excluded and are not written.` : ""}</div>
      ${renderProposalGroups(proposal, proposal.rows, { locked: true, applied: true })}
      <div class="undo-note">${icon("info")} This apply is one Ctrl+Z step. The grid is unchanged until you press it.</div>`;
  }
  const visibleRows = proposal.showAllRows ? proposal.rows : proposal.rows.slice(0, 8);
  return `
    ${renderProposalGroups(proposal, visibleRows, { locked: false })}
    ${proposal.rows.length > 8 ? `<button class="more-changes" type="button" data-action="toggle-all-rows">${icon(proposal.showAllRows ? "chevron-up" : "chevron-down")} ${proposal.showAllRows ? "Show fewer" : `Show ${proposal.rows.length - 8} more changes`}</button>` : ""}
    ${renderOverwriteAck(proposal)}`;
}

function renderProposalGroups(proposal, rows, options = {}) {
  const groups = [];
  for (const row of rows) {
    let group = groups.find((item) => item.sheet === row.sheet);
    if (!group) {
      group = { sheet: row.sheet, rows: [] };
      groups.push(group);
    }
    group.rows.push(row);
  }
  return groups.map((group) => `
    <div class="change-group">
      <div class="group-head"><span>${escapeHtml(group.sheet)}</span><span class="group-count">${group.rows.length} change${group.rows.length === 1 ? "" : "s"}</span></div>
      ${group.rows.map((row) => renderWarmChangeRow(row, proposal, options)).join("")}
    </div>`).join("");
}

function renderWarmChangeRow(row, proposal, options = {}) {
  const locked = options.locked || proposal.status !== "proposed";
  const outcome = row.outcome;
  const rowClass = [
    "change-row",
    locked ? "is-locked" : "",
    row.included ? "is-included" : "is-excluded",
    outcome ? `is-${outcome}` : "",
    row.risk === "overwrite" ? "is-risk-overwrite" : "",
    row.risk === "confirm" || row.risk === "loss" ? "is-risk-danger" : "",
  ].filter(Boolean).join(" ");
  return `<div class="${rowClass}" data-row="${row.id}">
    <button class="change-check${row.included ? " is-checked" : ""}" type="button" role="checkbox" aria-checked="${row.included}" aria-label="Include change to ${escapeHtml(row.range)}: ${escapeHtml(row.op)}" data-action="toggle-row" data-row-id="${row.id}" ${locked ? "disabled" : ""}>
      ${row.included ? icon("check") : ""}
    </button>
    <div class="change-main">
      <div class="change-line">
        <button class="range-chip" type="button" data-action="citation" data-cite="${escapeHtml(`${row.sheet}!${row.range}`)}">${escapeHtml(row.range)}</button>
        <span class="change-op">${escapeHtml(row.op)}</span>
        ${row.risk === "overwrite" ? '<span class="badge badge-warn">overwrite</span>' : ""}
        ${row.risk === "confirm" ? '<span class="badge badge-danger">risky formula</span>' : ""}
        ${row.risk === "loss" ? '<span class="badge badge-danger">loss</span>' : ""}
      </div>
      <div class="change-diff"><span class="diff-was">${escapeHtml(row.before || "blank")}</span><span class="diff-arrow">→</span><span class="diff-now">${escapeHtml(row.after || "blank")}</span></div>
      ${row.note ? `<div class="change-note${row.risk === "confirm" || row.risk === "loss" ? " is-danger" : ""}">${icon(row.risk === "overwrite" ? "alert" : "info")} ${escapeHtml(row.note)}</div>` : ""}
      ${outcome ? renderRowOutcome(row) : ""}
    </div>
  </div>`;
}

function renderRowOutcome(row) {
  if (row.outcome === "applied") return `<div class="row-status is-success">${icon("check")} Applied</div>`;
  if (row.outcome === "failed") return `<div class="row-status is-danger">${icon("alert")} Failed · ${escapeHtml(row.error || "Excel was busy")}</div>`;
  if (row.outcome === "excluded") return `<div class="row-status is-muted">Excluded from this apply</div>`;
  if (row.outcome === "queued") return `<div class="row-status is-muted">Waiting for Excel…</div>`;
  return "";
}

function renderOverwriteAck(proposal) {
  const cells = overwriteSelectedCells(proposal);
  if (!cells) return "";
  return `<button class="ack-row${proposal.ack ? " is-checked" : ""}" type="button" role="checkbox" aria-checked="${proposal.ack}" data-action="toggle-ack">
    <span class="ack-check">${proposal.ack ? icon("check") : ""}</span>
    <span>Overwrite <strong>${cells} cells</strong> that contain data.</span>
  </button>`;
}

function applyLabel(proposal) {
  const count = selectedCount(proposal);
  if (!count) return "Apply changes";
  return count === proposal.rows.length ? `Apply ${count} changes` : `Apply ${count} of ${proposal.rows.length}`;
}

function renderWarmProposalFooter(proposal) {
  if (proposal.status === "applying") {
    return `<div class="edit-foot"><div class="edit-count"><strong>Applying…</strong> · one undo group</div></div>`;
  }
  if (proposal.status === "partial") {
    return `<div class="edit-foot edit-foot-partial">
      <div class="edit-count"><strong>${proposal.appliedCount} applied</strong> · ${proposal.failedCount} pending</div>
      <div class="edit-actions">
        <button class="button button-ghost" type="button" data-action="dismiss-proposal">Dismiss</button>
        <button class="button button-primary" type="button" data-action="retry-failed">Retry ${proposal.failedCount} failed</button>
      </div>
    </div>`;
  }
  if (proposal.status === "dismissed") {
    return `<div class="edit-foot edit-foot-partial">
      <div class="edit-count"><strong>${proposal.failedCount} pending</strong> · not applied</div>
      <div class="edit-actions"><button class="button button-ghost" type="button" data-action="toggle-card">Review</button></div>
    </div>`;
  }
  if (proposal.status === "applied") {
    return `<div class="edit-foot edit-foot-applied">
      <div class="edit-count"><strong>${proposal.appliedCount} applied</strong>${proposal.excludedCount ? ` · ${proposal.excludedCount} excluded` : ""}</div>
      <div class="edit-actions">
        <button class="button button-ghost" type="button" data-action="reveal-applied">Reveal</button>
        <button class="button button-secondary" type="button" data-action="toggle-card">Review</button>
      </div>
    </div>`;
  }
  const count = selectedCount(proposal);
  const ackNeeded = overwriteSelectedCells(proposal) > 0 && !proposal.ack;
  return `<div class="edit-foot">
    <div class="edit-count"><strong>${count} selected</strong> · ${sheetGroupCount(proposal)} sheets</div>
    <div class="edit-actions">
      <button class="button button-ghost" type="button" data-action="discard-proposal">Discard</button>
      <button class="button button-primary" type="button" data-action="apply" ${!count || ackNeeded ? "disabled" : ""}>${escapeHtml(applyLabel(proposal))}</button>
    </div>
  </div>`;
}

/* ---- Fluent proposal ---------------------------------------------------- */

function renderFluentProposal(proposal) {
  const collapsed = proposal.collapsed && !["applying", "partial"].includes(proposal.status);
  return `<section class="fluent-review-card edit-${proposal.status}" id="editCard" aria-label="Edit proposal">
    <div class="fluent-review-head">
      <span class="fluent-review-glyph glyph-${proposal.status}">${proposalGlyph(proposal)}</span>
      <div class="edit-heading">
        <div class="edit-title">${escapeHtml(proposalTitle(proposal))}</div>
        <div class="edit-sub">${escapeHtml(proposalSubtitle(proposal))}</div>
        <div class="edit-badges">${riskBadges(proposal)}</div>
      </div>
      <button class="edit-chevron" type="button" data-action="toggle-card" aria-expanded="${!collapsed}" aria-label="${collapsed ? "Expand changes" : "Collapse changes"}">${icon(collapsed ? "chevron-down" : "chevron-up")}</button>
    </div>
    ${collapsed ? "" : `<div class="fluent-review-body">${renderFluentProposalBody(proposal)}</div>${renderFluentProposalFooter(proposal)}`}
  </section>`;
}

function renderFluentProposalBody(proposal) {
  if (proposal.status === "applying") {
    return `<div class="fluent-applybar" role="progressbar" aria-label="Applying changes"><span></span></div>
      <div class="fluent-apply-caption">Applying ${selectedCount(proposal)} changes in one undo group…</div>
      ${renderFluentTable(proposal, proposal.rows, true)}`;
  }
  if (proposal.status === "partial" || proposal.status === "dismissed") {
    const failed = proposal.rows.filter((row) => row.outcome === "failed");
    const applied = proposal.rows.filter((row) => row.outcome === "applied" || row.outcome === "excluded");
    return `<div class="fluent-alert">${icon("alert")}<div><strong>Excel was still in cell-edit mode.</strong> ${proposal.failedCount} changes weren’t applied. Finish the edit, then retry.</div></div>
      ${renderFluentTable(proposal, failed, true)}
      ${proposal.showApplied ? renderFluentTable(proposal, applied, true) : ""}
      <button class="more-changes" type="button" data-action="toggle-applied">${icon(proposal.showApplied ? "chevron-up" : "chevron-down")} ${proposal.showApplied ? "Hide" : `Show ${proposal.appliedCount}`} applied changes</button>`;
  }
  if (proposal.status === "applied") {
    return `<div class="fluent-applied-note">${icon("check")} Selected ranges are written.${proposal.excludedCount ? ` ${proposal.excludedCount} rows were excluded.` : ""}</div>
      ${renderFluentTable(proposal, proposal.rows, true)}
      <div class="undo-note">${icon("info")} This apply is one Ctrl+Z step.</div>`;
  }
  const visibleRows = proposal.showAllRows ? proposal.rows : proposal.rows.slice(0, 8);
  return `${renderFluentTable(proposal, visibleRows, false)}
    ${proposal.rows.length > 8 ? `<button class="more-changes" type="button" data-action="toggle-all-rows">${icon(proposal.showAllRows ? "chevron-up" : "chevron-down")} ${proposal.showAllRows ? "Show fewer" : `Show ${proposal.rows.length - 8} more changes`}</button>` : ""}
    ${renderOverwriteAck(proposal)}`;
}

function renderFluentTable(proposal, rows, locked) {
  const groups = [];
  for (const row of rows) {
    let group = groups.find((item) => item.sheet === row.sheet);
    if (!group) {
      group = { sheet: row.sheet, rows: [] };
      groups.push(group);
    }
    group.rows.push(row);
  }
  return `<div class="fluent-review-scroll">
    <table class="fluent-review-table">
      <thead><tr><th class="col-include"><span class="sr-only">Include</span></th><th>Cell</th><th>Change</th><th>Risk</th></tr></thead>
      <tbody>
        ${groups.map((group) => `
          <tr class="review-group"><td colspan="4">${escapeHtml(group.sheet)} <span>${group.rows.length} change${group.rows.length === 1 ? "" : "s"}</span></td></tr>
          ${group.rows.map((row) => `
            <tr class="review-row${row.included ? "" : " is-excluded"}${row.outcome ? ` is-${row.outcome}` : ""}">
              <td class="col-include"><button class="change-check${row.included ? " is-checked" : ""}" type="button" role="checkbox" aria-checked="${row.included}" aria-label="Include change to ${escapeHtml(row.range)}" data-action="toggle-row" data-row-id="${row.id}" ${locked || proposal.status !== "proposed" ? "disabled" : ""}>${row.included ? icon("check") : ""}</button></td>
              <td><button class="range-chip" type="button" data-action="citation" data-cite="${escapeHtml(`${row.sheet}!${row.range}`)}">${escapeHtml(row.range)}</button></td>
              <td>
                <div class="fluent-review-op">${escapeHtml(row.op)}</div>
                <div class="change-diff"><span class="diff-was">${escapeHtml(row.before || "blank")}</span><span class="diff-arrow">→</span><span class="diff-now">${escapeHtml(row.after || "blank")}</span></div>
                ${row.note ? `<div class="change-note${row.risk === "confirm" || row.risk === "loss" ? " is-danger" : ""}">${escapeHtml(row.note)}</div>` : ""}
                <div class="fluent-risk-inline">${row.risk === "overwrite" ? '<span class="badge badge-warn">overwrite</span>' : row.risk === "confirm" ? '<span class="badge badge-danger">risky formula</span>' : row.risk === "loss" ? '<span class="badge badge-danger">loss</span>' : ""}</div>
                ${row.outcome ? renderRowOutcome(row) : ""}
              </td>
              <td>${row.risk === "overwrite" ? '<span class="badge badge-warn">overwrite</span>' : row.risk === "confirm" ? '<span class="badge badge-danger">risky</span>' : row.risk === "loss" ? '<span class="badge badge-danger">loss</span>' : '<span class="risk-normal">—</span>'}</td>
            </tr>`).join("")}`).join("")}
      </tbody>
    </table>
  </div>`;
}

function renderFluentProposalFooter(proposal) {
  if (proposal.status === "applying") {
    return `<div class="fluent-review-foot"><span><strong>Applying…</strong> · one undo group</span></div>`;
  }
  if (proposal.status === "partial") {
    return `<div class="fluent-review-foot">
      <span><strong>${proposal.appliedCount} applied</strong> · ${proposal.failedCount} pending</span>
      <div class="edit-actions"><button class="button button-ghost" type="button" data-action="dismiss-proposal">Dismiss</button><button class="button button-primary" type="button" data-action="retry-failed">Retry ${proposal.failedCount}</button></div>
    </div>`;
  }
  if (proposal.status === "dismissed") {
    return `<div class="fluent-review-foot"><span><strong>${proposal.failedCount} pending</strong></span><button class="button button-ghost" type="button" data-action="toggle-card">Review</button></div>`;
  }
  if (proposal.status === "applied") {
    return `<div class="fluent-review-foot"><span><strong>${proposal.appliedCount} applied</strong>${proposal.excludedCount ? ` · ${proposal.excludedCount} excluded` : ""}</span><div class="edit-actions"><button class="button button-ghost" type="button" data-action="reveal-applied">Reveal</button><button class="button button-secondary" type="button" data-action="toggle-card">Review</button></div></div>`;
  }
  const count = selectedCount(proposal);
  const ackNeeded = overwriteSelectedCells(proposal) > 0 && !proposal.ack;
  return `<div class="fluent-review-foot">
    <span><strong>${count} selected</strong> · ${sheetGroupCount(proposal)} sheets</span>
    <div class="edit-actions">
      <button class="button button-ghost" type="button" data-action="discard-proposal">Discard</button>
      <button class="button button-primary" type="button" data-action="apply" ${!count || ackNeeded ? "disabled" : ""}>${escapeHtml(applyLabel(proposal))}</button>
    </div>
  </div>`;
}

/* -------------------------------------------------------------------------
   History
   ------------------------------------------------------------------------- */

function stripMarkdown(text) {
  return String(text || "")
    .replace(/\[([^\]]+)\]\(cite:[^)]+\)/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .replace(/[#>*_`~|]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function currentConversation() {
  return state.conversations.find((conversation) => conversation.id === state.currentId) || null;
}

function pendingForConversation(conversation) {
  if (conversation.id === state.currentId) {
    return !!state.proposal && ["proposed", "applying", "partial", "dismissed"].includes(state.proposal.status);
  }
  if (conversation._proposal) {
    return ["proposed", "applying", "partial", "dismissed"].includes(conversation._proposal.status);
  }
  return !!conversation.pending;
}

function displayConversation(conversation) {
  if (conversation.id !== state.currentId) return conversation;
  const firstUser = state.messages.find((message) => message.role === "user")?.text;
  const lastAssistant = [...state.messages].reverse().find((message) => message.role === "assistant");
  return {
    ...conversation,
    title: conversation.customTitle || (firstUser ? firstUser.slice(0, 58) : conversation.title),
    preview: lastAssistant ? stripMarkdown(lastAssistant.text).slice(0, 92) : conversation.preview,
    count: state.messages.length + (state.proposal ? 1 : 0),
    pending: pendingForConversation(conversation),
  };
}

function visibleConversations() {
  if (state.unsaved) return [];
  return state.conversations
    .filter((conversation) => state.historyScope === "all" || conversation.workbook === workbook().name)
    .map(displayConversation);
}

function renderHistoryNotice() {
  if (!state.historyNotice) return "";
  return `<div class="history-notice" role="status">
    ${icon("alert")}
    <div class="history-notice-copy">
      <strong>${escapeHtml(state.historyNotice.title)}</strong>
      <span>${escapeHtml(state.historyNotice.text)}</span>
    </div>
    ${state.historyNotice.action ? `<button class="text-action" type="button" data-action="${state.historyNotice.action}">${escapeHtml(state.historyNotice.actionLabel || "Open folder")}</button>` : ""}
  </div>`;
}

function renderHistoryView() {
  const rows = visibleConversations();
  const groups = ["Today", "Yesterday", "Earlier"];
  return `
    <div class="pane-view pane-view-history pane-view-${state.variant} view-${state.variant}" data-view="history">
      ${renderViewHeader("History", { action: "new-chat", actionLabel: "New chat", actionIcon: "plus" })}
      ${renderBudgetBanner()}
      ${renderBanner()}
      <div class="view-body history-body">
        <div class="scope-segmented" role="group" aria-label="History scope">
          <button class="scope-segment${state.historyScope === "workbook" ? " is-active" : ""}" type="button" data-action="history-scope" data-scope="workbook">This workbook</button>
          <button class="scope-segment${state.historyScope === "all" ? " is-active" : ""}" type="button" data-action="history-scope" data-scope="all">All workbooks</button>
        </div>
        ${renderHistoryNotice()}
        ${rows.length ? groups.map((group) => {
          const groupRows = rows.filter((row) => row.group === group);
          if (!groupRows.length) return "";
          return `<div class="history-group">
            <div class="history-group-label">${group}</div>
            ${groupRows.map((row) => renderHistoryRow(row)).join("")}
          </div>`;
        }).join("") : state.unsaved ? `<div class="history-empty">
            <div class="empty-mark">${icon("lock")}</div>
            <h2>History starts when you save this workbook</h2>
            <p>This conversation is memory-only. If the pane closes before you save, the draft is gone.</p>
            <button class="button button-secondary" type="button" disabled title="Save the workbook first">${icon("link")} Link an old conversation…</button>
            <div class="history-empty-hint">Save the workbook first. Then this pane keeps its own history.</div>
          </div>` : `<div class="history-empty">
            <div class="empty-mark">${icon("clock")}</div>
            <h2>No conversations yet</h2>
            <p>Your chats are saved automatically, per workbook.</p>
            <button class="button button-primary" type="button" data-action="new-chat">Start a chat</button>
            <button class="button button-secondary history-link-button" type="button" data-action="open-link" data-from="history">${icon("link")} Link an old conversation…</button>
          </div>`}
      </div>
      ${renderSnackbarSlot()}
      <div class="view-foot history-foot">
        ${icon("lock")}<span>Stored on this machine</span>
        <span class="spacer"></span>
        <button class="text-action danger-text" type="button" data-action="delete-all" ${state.conversations.length ? "" : "disabled"}>Delete all…</button>
      </div>
    </div>`;
}

function renderHistoryRow(row) {
  const active = row.id === state.currentId;
  const menuOpen = state.rowMenuId === row.id;
  const badges = [
    row.pending ? `<span class="badge badge-warn">Pending</span>` : "",
    row.recovered ? `<span class="badge badge-success">Recovered</span>` : "",
    row.movedFrom ? `<span class="badge badge-muted">Moved here</span>` : "",
  ].filter(Boolean).join("");
  return `<div class="history-row-wrap${menuOpen ? " has-menu" : ""}">
    <button class="history-row${active ? " is-active" : ""}" type="button" data-action="open-conversation" data-id="${escapeHtml(row.id)}" ${active ? 'aria-current="true"' : ""}>
      <span class="history-main">
        <span class="history-title-line"><span class="history-title">${escapeHtml(row.title)}</span><span class="history-time">${escapeHtml(row.time)}</span></span>
        <span class="history-preview">${escapeHtml(row.preview)}</span>
        <span class="history-meta">${row.count || 0} messages · ${escapeHtml(row.sheets || "Forecast")}${state.historyScope === "all" ? ` · ${escapeHtml(row.workbook)}` : ""}</span>
        ${badges ? `<span class="history-badges">${badges}</span>` : ""}
      </span>
    </button>
    <button class="history-more" type="button" data-action="row-menu" data-id="${escapeHtml(row.id)}" aria-label="More actions for ${escapeHtml(row.title)}" aria-expanded="${menuOpen}">${icon("more")}</button>
    ${menuOpen ? `<div class="row-menu" role="menu">
      <button class="menu-row" type="button" role="menuitem" data-action="rename-conversation" data-id="${escapeHtml(row.id)}">${icon("pen", "icon menu-icon")}<span>Rename</span></button>
      <button class="menu-row" type="button" role="menuitem" data-action="export-conversation" data-id="${escapeHtml(row.id)}" data-format="md">${icon("download", "icon menu-icon")}<span>Export .md</span></button>
      <button class="menu-row" type="button" role="menuitem" data-action="export-conversation" data-id="${escapeHtml(row.id)}" data-format="json">${icon("download", "icon menu-icon")}<span>Export .json</span></button>
      ${row.movedFrom ? `<button class="menu-row" type="button" role="menuitem" data-action="unlink-conversation" data-id="${escapeHtml(row.id)}">${icon("link", "icon menu-icon")}<span>Move back to ${escapeHtml(row.movedFrom)}</span></button>` : ""}
      <div class="menu-separator"></div>
      <button class="menu-row danger-text" type="button" role="menuitem" data-action="delete-conversation" data-id="${escapeHtml(row.id)}">${icon("x", "icon menu-icon")}<span>Delete</span></button>
    </div>` : ""}
  </div>`;
}

/* -------------------------------------------------------------------------
   Link picker
   ------------------------------------------------------------------------- */

function linkableConversations() {
  const query = state.linkFilter.trim().toLowerCase();
  return state.conversations.filter((conversation) => {
    if (conversation.workbook === workbook().name) return false;
    if (!query) return true;
    return [conversation.title, conversation.workbook, conversation.path, conversation.preview]
      .some((value) => String(value || "").toLowerCase().includes(query));
  });
}

function renderLinkList() {
  if (state.unsaved) {
    return `<div class="link-empty">
      <div class="empty-mark">${icon("lock")}</div>
      <h2>Save the workbook first</h2>
      <p>Linking needs a saved workbook path. Your current draft stays in memory-only history until you save.</p>
    </div>`;
  }
  const rows = linkableConversations();
  if (!state.conversations.some((conversation) => conversation.workbook !== workbook().name)) {
    return `<div class="link-empty">
      <div class="empty-mark">${icon("link")}</div>
      <h2>No other conversations yet</h2>
      <p>Conversations from other workbooks appear here after they are saved.</p>
    </div>`;
  }
  if (!rows.length) {
    return `<div class="link-empty">
      <div class="empty-mark">${icon("search")}</div>
      <h2>No conversations match</h2>
      <p>Try a workbook name or part of a path.</p>
      <button class="button button-secondary button-small" type="button" data-action="clear-link-filter">Clear filter</button>
    </div>`;
  }
  const groups = new Map();
  for (const row of rows) {
    const key = `${row.workbook}\u0000${row.path || ""}`;
    if (!groups.has(key)) groups.set(key, { name: row.workbook, path: row.path || "", rows: [] });
    groups.get(key).rows.push(row);
  }
  return [...groups.values()].map((group) => `<section class="link-group" aria-label="${escapeHtml(group.name)}">
    <div class="link-group-head">
      <strong>${escapeHtml(group.name)}</strong>
      <span title="${escapeHtml(group.path)}">${escapeHtml(group.path)}</span>
    </div>
    ${group.rows.map((row) => `<div class="link-row">
      <div class="link-row-main">
        <div class="link-row-title">${escapeHtml(row.title)}</div>
        <div class="link-row-path" title="${escapeHtml(row.path)}">${escapeHtml(row.path)}</div>
        <div class="link-row-meta">${row.count || 0} messages · ${escapeHtml(row.sheets || "Forecast")} · ${escapeHtml(row.date || row.time || "")}</div>
      </div>
      <button class="button button-secondary button-small link-move" type="button" data-action="link-conversation" data-id="${escapeHtml(row.id)}">Move to this workbook</button>
    </div>`).join("")}
  </section>`).join("");
}

function renderLinkView() {
  return `
    <div class="pane-view pane-view-link pane-view-${state.variant} view-${state.variant}" data-view="link">
      <header class="view-header">
        <button class="icon-button" type="button" data-action="back-link" aria-label="Back">${icon("arrow-left")}</button>
        <h1 class="view-title" id="viewTitle" tabindex="-1">Link to an old conversation</h1>
        <span class="view-header-spacer"></span>
      </header>
      <div class="view-body link-body">
        <p class="link-intro">Moves a conversation from another workbook. It is not copied, and you can move it back later.</p>
        <label class="link-search">
          ${icon("search")}
          <input id="linkFilter" type="search" value="${escapeHtml(state.linkFilter)}" placeholder="Filter by path or name" aria-label="Filter conversations by path or name" autocomplete="off">
        </label>
        <div id="linkList">${renderLinkList()}</div>
      </div>
      ${renderSnackbarSlot()}
    </div>`;
}

/* -------------------------------------------------------------------------
   Settings
   ------------------------------------------------------------------------- */

function renderSettingsView() {
  const charCount = state.instructions.length;
  const savedLabel = state.instructionsState === "saving" ? "Saving…" : state.instructionsState === "error" ? "Not saved" : "Saved";
  const usedMb = historyUsedMb();
  const pct = historyPercent();
  const totalConversations = state.historyRows.reduce((total, row) => total + row.conversations, 0);
  const readOnlyForced = state.workbookReadOnly;
  const readOnlyPending = !!state.proposal && proposalGatesComposer();
  const readOnlyDisabled = readOnlyForced || readOnlyPending;
  const readOnlyDescription = readOnlyForced
    ? "Excel opened this workbook read-only, so edits are unavailable."
    : readOnlyPending
      ? "Apply or discard the pending changes before switching modes."
      : "No edits or proposals. Reads and questions still work.";
  const readOnlyFootnote = readOnlyForced
    ? "Read-only mode is forced on by Excel and can’t be turned off here."
    : "Saved with this workbook’s history; never written into the file.";
  return `
    <div class="pane-view pane-view-settings pane-view-${state.variant} view-${state.variant}" data-view="settings">
      ${renderViewHeader("Settings")}
      ${renderBudgetBanner()}
      ${renderBanner()}
      <div class="view-body settings-body">
        <div class="section-label">Instructions</div>
        <div class="settings-card settings-card-pad">
          <textarea class="instructions-input" id="instructions" rows="4" maxlength="8200" aria-label="Instructions for this workbook" placeholder="Format numbers with thousands separators. Bold column headers. Use € for currency.">${escapeHtml(state.instructions)}</textarea>
          ${!state.instructions.trim() && state.instructionsSource ? `<button class="instructions-chip" type="button" data-action="copy-instructions">
            ${icon("copy")}<span>Copy instructions from ${escapeHtml(state.instructionsSource.workbook)}</span>
          </button>` : ""}
          <div class="settings-help"><span>Applies to every conversation in this workbook.</span><span class="saved-state" id="instructionsState">${savedLabel === "Saved" ? `${icon("check")} ` : ""}${savedLabel}</span></div>
          <div class="char-counter" id="instructionsCount">${charCount.toLocaleString()} / 8,000</div>
        </div>

        <div class="section-label">Model</div>
        <div class="settings-card">
          <button class="setting-row" type="button" data-action="toggle-settings-model" aria-expanded="${state.settingsModelOpen}">
            <span class="setting-main"><span class="setting-label">Model</span><span class="setting-desc">Used for new messages. The current conversation keeps its model.</span></span>
            <span class="setting-value">${icon("grid", "icon icon-14")} ${escapeHtml(state.model)}</span>
            ${icon("chevron-right", "icon setting-chevron")}
          </button>
          ${state.settingsModelOpen ? `<div class="settings-model-list">${state.models.map((model) => `
            <button class="settings-model-row${model.name === state.model ? " is-selected" : ""}" type="button" data-action="select-model" data-model="${escapeHtml(model.name)}">
              <span class="model-check">${model.name === state.model ? icon("check") : ""}</span>
              <span>${escapeHtml(model.name)}</span>
              ${model.tag ? `<span class="model-tag">${escapeHtml(model.tag)}</span>` : ""}
            </button>`).join("")}</div>` : ""}
          <div class="setting-row static-row">
            <span class="setting-main"><span class="setting-label">Refresh model list</span><span class="setting-desc">Fetched from your gateway · ${escapeHtml(state.modelUpdatedAt)}</span></span>
            <button class="text-action primary-text" type="button" data-action="refresh-models" ${state.modelsLoading ? "disabled" : ""}>${state.modelsLoading ? "Refreshing…" : "Refresh"}</button>
          </div>
        </div>

        <div class="section-label">Appearance</div>
        <div class="settings-card">
          <div class="setting-row static-row">
            <span class="setting-main"><span class="setting-label">Theme</span><span class="setting-desc">Auto follows Excel’s light or dark theme.</span></span>
            <div class="theme-segmented" role="group" aria-label="Theme">
              ${THEMES.map((theme) => `<button class="theme-segment${state.themeMode === theme ? " is-active" : ""}" type="button" data-action="set-theme" data-theme="${theme}">${theme[0].toUpperCase() + theme.slice(1)}</button>`).join("")}
            </div>
          </div>
        </div>

        <div class="section-label">Startup</div>
        <div class="settings-card">
          <div class="setting-row static-row">
            <span class="setting-main"><span class="setting-label">Open with this workbook</span><span class="setting-desc">Adds a setting to ${escapeHtml(workbook().name)}. Off by default; turn it off here.</span></span>
            <button class="switch${state.autoOpen ? " is-on" : ""}" type="button" role="switch" aria-checked="${state.autoOpen}" data-action="toggle-auto-open" aria-label="Open with this workbook"><span></span></button>
          </div>
        </div>

        <div class="section-label" id="editingSection" tabindex="-1">Editing</div>
        <div class="settings-card" id="readOnlyCard">
          <div class="setting-row static-row">
            <span class="setting-main">
              <span class="setting-label">Read-only mode</span>
              <span class="setting-desc">${readOnlyDescription}</span>
              <span class="setting-desc setting-footnote">${readOnlyFootnote}</span>
            </span>
            <button id="readOnlySwitch" class="switch${isReadOnly() ? " is-on" : ""}${readOnlyDisabled ? " is-disabled" : ""}" type="button" role="switch" aria-checked="${isReadOnly()}" data-action="toggle-readonly" ${readOnlyDisabled ? "disabled" : ""} aria-label="Read-only mode"><span></span></button>
          </div>
          ${state.readOnlyJustSaved ? `<div class="settings-saved-row"><span class="saved-state">${icon("check")} Saved</span></div>` : ""}
        </div>

        <div class="section-label" id="historySection" tabindex="-1">History</div>
        <div class="settings-card" id="historyCard">
          <div class="history-usage">
            <div class="history-usage-line"><strong>${usedMb} MB of ${state.historyLimitMb} MB</strong><span class="${pct >= 100 ? "danger-text" : ""}">${pct}%</span></div>
            <div class="budget-track${pct >= 100 ? " is-over" : ""}" role="progressbar" aria-label="History budget" aria-valuemin="0" aria-valuemax="${state.historyLimitMb}" aria-valuenow="${Math.min(usedMb, state.historyLimitMb)}"><span style="width:${Math.min(100, pct)}%"></span></div>
            <div class="history-usage-meta">${totalConversations} conversation${totalConversations === 1 ? "" : "s"} · Stored on this machine</div>
          </div>
          ${state.historyRows.length ? state.historyRows.map((row) => `<div class="setting-row static-row history-manage-row" data-history-key="${escapeHtml(row.key)}">
            <span class="setting-main">
              <span class="setting-label">${escapeHtml(row.name)}</span>
              <span class="setting-desc history-path" title="${escapeHtml(row.path)}">${escapeHtml(row.path)}</span>
              <span class="setting-desc">${row.conversations} conversation${row.conversations === 1 ? "" : "s"} · ${row.sizeMb} MB</span>
            </span>
            <button class="text-action danger-text" type="button" data-action="delete-workbook-history" data-key="${escapeHtml(row.key)}">Delete</button>
          </div>`).join("") : `<div class="settings-empty">No local history. New conversations are saved automatically.</div>`}
          <div class="history-settings-actions">
            <button class="text-action" type="button" data-action="open-folder">${icon("folder")} Open folder</button>
            <button class="text-action danger-text" type="button" data-action="delete-all" ${state.historyRows.length ? "" : "disabled"}>Delete all history…</button>
          </div>
        </div>

        <div class="about-store">
          <div class="about-store-head">${icon("folder")}<span>History folder</span></div>
          <div class="about-store-path" title="${escapeHtml(HISTORY_STORE_PATH)}">${escapeHtml(HISTORY_STORE_PATH)}</div>
          <div class="about-store-actions">
            <button class="text-action" type="button" data-action="open-folder">Open folder</button>
            <button class="text-action" type="button" data-action="copy-store-path">Copy path</button>
          </div>
        </div>

        <div class="about-line"><span class="status-dot ${state.offline ? "is-error" : "is-done"}"></span> Excel Agent 0.1.0 · ${state.offline ? "service unreachable" : "localhost:3000"} · <button class="text-action" type="button" data-action="copy-details" data-copy="Excel Agent 0.1.0 · Excel 16.0 build 20527 · localhost:3000 · no gateway key in this prototype">Copy diagnostics</button></div>
      </div>
      ${renderSnackbarSlot()}
    </div>`;
}

function renderViewHeader(title, options = {}) {
  return `<header class="view-header">
    <button class="icon-button" type="button" data-action="back-to-chat" aria-label="Back to chat">${icon("arrow-left")}</button>
    <h1 class="view-title" id="viewTitle" tabindex="-1">${escapeHtml(title)}</h1>
    ${options.action ? `<button class="icon-button" type="button" data-action="${options.action}" aria-label="${escapeHtml(options.actionLabel || options.action)}" title="${escapeHtml(options.actionLabel || options.action)}">${icon(options.actionIcon || "plus")}</button>` : "<span class=\"view-header-spacer\"></span>"}
  </header>`;
}

/* -------------------------------------------------------------------------
   Dialog
   ------------------------------------------------------------------------- */

function renderDialog() {
  if (!state.dialog) return "";
  const dialog = state.dialog;
  return `<div class="dialog-scrim" data-action="dialog-backdrop">
    <div class="dialog" role="dialog" aria-modal="true" aria-labelledby="dialogTitle">
      <h2 id="dialogTitle" class="dialog-title">${escapeHtml(dialog.title)}</h2>
      <div class="dialog-body">${escapeHtml(dialog.body || "")}</div>
      ${dialog.input ? `<label class="dialog-field"><span>${escapeHtml(dialog.inputLabel || "Name")}</span><input id="dialogInput" type="text" value="${escapeHtml(dialog.inputValue || "")}" maxlength="80"></label>` : ""}
      <div class="dialog-actions">
        <button class="button button-ghost" type="button" data-action="dialog-cancel">Cancel</button>
        <button class="button ${dialog.danger ? "button-danger" : "button-primary"}" type="button" data-action="dialog-confirm">${escapeHtml(dialog.confirmLabel || "Confirm")}</button>
      </div>
    </div>
  </div>`;
}

/* -------------------------------------------------------------------------
   Render lifecycle
   ------------------------------------------------------------------------- */

function render() {
  const theme = resolveTheme();
  const currentTranscript = $("#transcript");
  if (currentTranscript) scrollMemory = currentTranscript.scrollTop;
  const currentViewBody = $(".view-body");
  if (lastView === state.view && currentViewBody) viewScrollMemory = currentViewBody.scrollTop;
  else viewScrollMemory = 0;
  lastView = state.view;
  document.documentElement.dataset.variant = state.variant;
  document.documentElement.dataset.theme = theme;
  document.documentElement.dataset.paneCompact = state.width <= 299 ? "true" : "false";
  document.documentElement.dataset.paneEmergency = state.width <= 239 ? "true" : "false";
  document.documentElement.style.setProperty("--pane-w", `${state.width}px`);
  app.innerHTML = `${renderExcelChrome()}${renderPrototypeBar()}${renderLiveRegion()}`;
  overlayRoot.innerHTML = renderDialog();
  requestAnimationFrame(afterRender);
}

function afterRender() {
  autoGrowComposer();
  if (state.focusAfter) {
    const selector = state.focusAfter;
    state.focusAfter = null;
    requestAnimationFrame(() => {
      const target = $(selector);
      if (!target) return;
      target.focus({ preventScroll: true });
      if (target.id === "composer" && state.focusCaret != null) {
        const caret = Math.min(state.focusCaret, target.value.length);
        target.setSelectionRange(caret, caret);
        state.focusCaret = null;
      }
    });
  }
  if (state.dialog) {
    requestAnimationFrame(() => {
      const input = $("#dialogInput");
      const cancel = $('[data-action="dialog-cancel"]');
      (input || cancel)?.focus();
    });
  }
  const shouldScrollBottom = state.scrollBottom;
  const shouldScrollToCard = state.scrollToCard;
  if (shouldScrollBottom) {
    state.scrollBottom = false;
    scrollMemory = null;
    requestAnimationFrame(() => {
      const transcript = $("#transcript");
      if (transcript) transcript.scrollTop = transcript.scrollHeight;
    });
  }
  if (shouldScrollToCard) {
    state.scrollToCard = false;
    scrollMemory = null;
    requestAnimationFrame(() => {
      const transcript = $("#transcript");
      const card = $("#editCard");
      if (!transcript || !card) return;
      const top = card.getBoundingClientRect().top - transcript.getBoundingClientRect().top + transcript.scrollTop;
      transcript.scrollTop = Math.max(0, top - 10);
    });
  }
  if (state.scrollViewTo) {
    const selector = state.scrollViewTo;
    state.scrollViewTo = null;
    requestAnimationFrame(() => {
      const target = $(selector);
      const body = $(".view-body");
      if (!target || !body) return;
      const top = target.getBoundingClientRect().top - body.getBoundingClientRect().top + body.scrollTop;
      body.scrollTop = Math.max(0, top - 10);
      viewScrollMemory = body.scrollTop;
    });
  } else if (viewScrollMemory > 0) {
    requestAnimationFrame(() => {
      const body = $(".view-body");
      if (body) body.scrollTop = viewScrollMemory;
    });
  }
  if (!shouldScrollBottom && !shouldScrollToCard && scrollMemory !== null) {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        const transcript = $("#transcript");
        if (transcript) transcript.scrollTop = scrollMemory;
        scrollMemory = null;
      });
    });
  }
  if (state.pendingAutoStream) {
    state.pendingAutoStream = false;
    window.setTimeout(() => startStream("question"), 320);
  }
}

function autoGrowComposer() {
  const composer = $("#composer");
  if (!composer) return;
  composer.style.height = "auto";
  composer.style.height = `${Math.min(composer.scrollHeight, 138)}px`;
}

function cancelLiveStream() {
  if (state.streamTimer) clearTimeout(state.streamTimer);
  state.streamTimer = null;
  state.live = null;
}

function clearAllTimers() {
  cancelLiveStream();
  if (state.applyTimer) clearTimeout(state.applyTimer);
  state.applyTimer = null;
  if (state.saveTimer) clearTimeout(state.saveTimer);
  state.saveTimer = null;
  if (state.snackbarTimer) clearTimeout(state.snackbarTimer);
  state.snackbarTimer = null;
}

/* -------------------------------------------------------------------------
   Scenario loading
   ------------------------------------------------------------------------- */

const QUESTION = "Why is the Q3 total lower than Q2 when units are up?";
const CLEANUP_PROMPT = "Update the totals with the new index.";
const REVIEW_STEPS = ["Read Forecast!A1:H400", "Read the selected table", "Staged 12 changes"];

function decorateAnswerMessages(messages = ANSWER_MESSAGES) {
  return clone(messages).map((message) => {
    if (message.role !== "assistant") return message;
    return {
      ...message,
      streaming: false,
      status: { state: "done", label: message.done || "3 steps · 4.8 s" },
      steps: message.steps || ["Read Forecast!A1:H400", "Read Forecast!B22:F25", "Checked 12 formulas"],
      showSteps: false,
      stopped: false,
    };
  });
}

function decorateThreadMessages(thread = []) {
  return clone(thread).map((message) => {
    if (message.role !== "assistant") return message;
    return {
      ...message,
      streaming: false,
      status: { state: "done", label: "2 steps · 3.2 s" },
      steps: ["Read Forecast!A1:H400", "Read the selection"],
      showSteps: false,
      stopped: false,
    };
  });
}

function reviewIntro() {
  return {
    role: "assistant",
    text: "I can fix the index and refresh the totals. Review the changes before anything is written.",
    streaming: false,
    status: { state: "done", label: "3 steps · 4.2 s" },
    steps: REVIEW_STEPS,
    showSteps: false,
    stopped: false,
  };
}

function readonlyRefusal() {
  return {
    role: "assistant",
    text: "I can’t stage an edit while read-only mode is on. Turn it off in Settings, then ask again — reads and questions still work.",
    streaming: false,
    status: { state: "done", label: "1 step · 0.8 s" },
    steps: ["Checked workbook state"],
    showSteps: false,
    stopped: false,
  };
}

function loadScenario(name, { initial = false } = {}) {
  clearAllTimers();
  state.scenario = SCENARIOS.some((scenario) => scenario.key === name) ? name : "empty";
  const viewByScenario = {
    history: "history",
    settings: "settings",
    link: "link",
    unsaved: "history",
    retention: "settings",
    "retention-full": "settings",
    storefail: "history",
    instructions: "settings",
    "readonly-forced": "settings",
  };
  state.view = viewByScenario[state.scenario] || "chat";
  state.messages = [];
  state.proposal = null;
  state.loading = false;
  state.offline = false;
  state.banner = null;
  state.snackbar = null;
  state.announcement = "";
  state.composerText = "";
  state.focusCaret = null;
  state.focusAfter = null;
  state.scrollBottom = true;
  state.scrollToCard = false;
  state.scrollViewTo = null;
  state.isScrolled = false;
  state.dialog = null;
  state.live = null;
  state.applyTimer = null;
  state.pendingAutoStream = false;
  state.modelMenuOpen = false;
  state.workbookMenuOpen = false;
  state.settingsModelOpen = false;
  state.rowMenuId = null;
  state.model = "gpt-5.2";
  state.models = clone(MODELS);
  state.modelsLoading = false;
  state.modelRefreshToken = null;
  state.modelUpdatedAt = "just now";
  state.instructions = INSTRUCTIONS_DEFAULT;
  state.instructionsState = "saved";
  state.autoOpen = false;
  state.workbook = clone(WORKBOOK);
  state.unsaved = false;
  state.historyRows = clone(HISTORY_WORKBOOKS);
  state.historyLimitMb = HISTORY_LIMIT_MB;
  state.linkFilter = "";
  state.linkFrom = "history";
  state.lastLink = null;
  state.compaction = null;
  state.storeFailure = false;
  state.historyNotice = null;
  state.recoveredId = null;
  state.instructionsSource = null;
  state.readOnly = false;
  state.workbookReadOnly = false;
  state.readOnlyJustSaved = false;
  state.sheets = clone(SHEETS);
  state.selection = { sheet: "Forecast", range: "B22:F25", active: "B22" };
  state.excelStatus = "Ready";
  state.conversations = clone(CONVERSATIONS).filter((conversation) => conversation.id !== "c-recovered");
  state.currentId = "c-current";
  state.historyScope = "workbook";
  state.lastPrompt = QUESTION;

  switch (state.scenario) {
    case "empty":
      state.focusAfter = "#composer";
      break;
    case "loading":
      state.loading = true;
      state.scrollBottom = false;
      break;
    case "streaming":
      state.messages = [{ role: "user", text: QUESTION }];
      state.pendingAutoStream = true;
      state.focusAfter = "#composer";
      break;
    case "answer":
      state.messages = decorateAnswerMessages();
      state.focusAfter = "#composer";
      break;
    case "review":
      state.messages = [{ role: "user", text: CLEANUP_PROMPT }, reviewIntro()];
      state.proposal = makeProposal("proposed");
      state.proposal.collapsed = false;
      state.focusAfter = "#editCard .edit-chevron";
      state.scrollBottom = false;
      state.scrollToCard = true;
      break;
    case "partial":
      state.messages = [{ role: "user", text: CLEANUP_PROMPT }, reviewIntro()];
      state.proposal = makeProposal("partial");
      state.proposal.collapsed = false;
      state.focusAfter = "#editCard .edit-chevron";
      state.scrollBottom = false;
      state.scrollToCard = true;
      break;
    case "applied":
      state.messages = [{ role: "user", text: CLEANUP_PROMPT }, reviewIntro()];
      state.proposal = makeProposal("applied");
      state.focusAfter = "#composer";
      state.scrollBottom = false;
      state.scrollToCard = true;
      break;
    case "history":
      state.messages = decorateAnswerMessages();
      state.focusAfter = "#viewTitle";
      break;
    case "settings":
      state.messages = decorateAnswerMessages();
      state.focusAfter = "#viewTitle";
      break;
    case "error":
      state.messages = [
        { role: "user", text: CLEANUP_PROMPT },
        {
          role: "error",
          title: "The model didn’t answer in time.",
          text: "Your workbook is unchanged. This is usually a gateway hiccup — try again when the banner clears.",
          detail: "Request timed out after 90 s · 22:14:06 · req_8f3a",
        },
      ];
      state.offline = true;
      state.banner = {
        tone: "warn",
        title: "No connection to the model gateway.",
        text: "Retrying in 8 s.",
        action: "retry-error",
        actionLabel: "Retry now",
      };
      break;
    case "link":
      state.messages = decorateAnswerMessages();
      state.focusAfter = "#linkFilter";
      break;
    case "unsaved":
      state.workbook = { name: "Book1", path: "", modified: "Not saved yet" };
      state.unsaved = true;
      state.messages = [];
      state.focusAfter = "#viewTitle";
      break;
    case "compaction":
      state.messages = decorateThreadMessages(COMPACTION_MESSAGES);
      state.compaction = {
        pct: 82,
        notice: true,
        summarizing: false,
        cuts: [],
        pending: { start: 0, end: 12, covered: 12 },
      };
      state.focusAfter = "#composer";
      state.scrollBottom = true;
      break;
    case "compaction-progress":
      state.messages = decorateThreadMessages(COMPACTION_MESSAGES);
      state.compaction = {
        pct: 82,
        notice: true,
        summarizing: true,
        cuts: [],
        pending: { start: 0, end: 12, covered: 12 },
      };
      state.focusAfter = ".compaction-notice";
      state.scrollBottom = true;
      break;
    case "compaction-stacked":
      state.messages = decorateThreadMessages(COMPACTION_MESSAGES);
      state.compaction = {
        pct: 84,
        notice: true,
        summarizing: false,
        cuts: [{ start: 0, end: 12, covered: 12, expanded: false, summary: "Earlier decisions, findings and the ranges that matter." }],
        pending: { start: 12, end: 16, covered: 4 },
      };
      state.focusAfter = "#composer";
      state.scrollBottom = true;
      break;
    case "retention":
      state.messages = decorateAnswerMessages();
      state.historyRows = clone(HISTORY_WORKBOOKS_RETENTION);
      state.focusAfter = "#historySection";
      state.scrollViewTo = "#historySection";
      break;
    case "retention-full":
      state.messages = decorateAnswerMessages();
      state.historyRows = clone(HISTORY_WORKBOOKS_FULL);
      state.focusAfter = "#historySection";
      state.scrollViewTo = "#historySection";
      break;
    case "storefail": {
      const recovered = CONVERSATIONS.find((conversation) => conversation.id === "c-recovered");
      if (recovered) state.conversations.unshift(clone(recovered));
      state.messages = decorateAnswerMessages();
      state.storeFailure = true;
      state.recoveredId = "c-recovered";
      state.banner = {
        tone: "warn",
        title: "History isn’t being saved.",
        text: "Your conversation is still open.",
        action: "retry-store",
        actionLabel: "Retry",
      };
      state.historyNotice = {
        title: "1 conversation couldn’t be read.",
        text: "It was set aside as 6f1c8e2a-3b7d-4f1a-9c2e-0a1b2c3d4e5f.json.bad.",
        action: "open-folder",
        actionLabel: "Open folder",
      };
      state.focusAfter = "#viewTitle";
      break;
    }
    case "instructions":
      state.messages = decorateAnswerMessages();
      state.instructions = "";
      state.instructionsSource = { workbook: "FY27 Budget.xlsx", text: INSTRUCTIONS_FROM_FY27 };
      state.focusAfter = "#instructions";
      state.scrollViewTo = "#instructions";
      break;
    case "readonly":
      state.messages = decorateAnswerMessages();
      state.messages.push({ role: "user", text: CLEANUP_PROMPT });
      state.messages.push(readonlyRefusal());
      state.readOnly = true;
      state.focusAfter = "#composer";
      state.scrollBottom = true;
      break;
    case "readonly-forced":
      state.messages = decorateAnswerMessages();
      state.workbookReadOnly = true;
      state.focusAfter = "#readOnlySwitch";
      state.scrollViewTo = "#editingSection";
      break;
    default:
      break;
  }

  if (state.view !== "chat") state.scrollBottom = false;
  if (state.scenario === "review" || state.scenario === "partial" || state.scenario === "applied") {
    state.lastPrompt = CLEANUP_PROMPT;
    state.currentId = "c-current";
  }
  if (!initial) syncUrl();
}

/* -------------------------------------------------------------------------
   Streaming
   ------------------------------------------------------------------------- */

function pinnedToBottom() {
  const transcript = $("#transcript");
  if (!transcript) return true;
  return transcript.scrollHeight - transcript.scrollTop - transcript.clientHeight < 96;
}

function scrollTranscript() {
  const transcript = $("#transcript");
  if (transcript) transcript.scrollTop = transcript.scrollHeight;
}

function startStream(key = "question") {
  cancelLiveStream();
  const response = RESPONSES[key] || RESPONSES.generic;
  const message = {
    role: "assistant",
    text: "",
    streaming: true,
    status: { state: "live", label: "Thinking…", step: "" },
    steps: [],
    showSteps: false,
    stopped: false,
  };
  state.messages.push(message);
  state.proposal = null;
  state.live = { message, response, index: 0, timer: null };
  state.scenario = "streaming";
  state.view = "chat";
  state.focusAfter = "#composer";
  state.scrollBottom = true;
  state.scrollToCard = false;
  syncUrl();
  render();
  announce("Response started");
  scheduleStreamStep();
}

function scheduleStreamStep() {
  if (!state.live) return;
  const live = state.live;
  const step = live.response.script[live.index];
  if (!step) {
    finishStream();
    return;
  }
  state.streamTimer = window.setTimeout(() => {
    if (state.live !== live) return;
    const wasPinned = pinnedToBottom();
    if (step.status) {
      live.message.status = { state: "live", ...step.status };
      updateLiveStatus();
    }
    if (step.text) {
      live.message.text += step.text;
      updateLiveBody();
    }
    live.index += 1;
    if (wasPinned) scrollTranscript();
    announce(step.status ? step.status.label : "Writing response");
    if (live.index >= live.response.script.length) {
      finishStream();
    } else {
      scheduleStreamStep();
    }
  }, step.delay || 360);
}

function updateLiveStatus() {
  const live = state.live;
  if (!live) return;
  const index = state.messages.indexOf(live.message);
  const current = $("#liveStatus");
  if (!current) return;
  current.outerHTML = renderStatusLine(live.message, index, true);
}

function updateLiveBody() {
  const live = state.live;
  if (!live) return;
  const body = $("#liveBody");
  if (!body) return;
  body.innerHTML = `${markdownToHtml(live.message.text, { partial: true })}<span class="caret" aria-hidden="true"></span>`;
}

function finishStream() {
  if (!state.live) return;
  const live = state.live;
  if (live.timer) clearTimeout(live.timer);
  const message = live.message;
  const response = live.response;
  message.streaming = false;
  message.status = { state: "done", label: response.done, step: "" };
  message.steps = response.steps;
  message.showSteps = false;
  state.live = null;
  state.streamTimer = null;
  if (response.stageProposal) {
    state.proposal = makeProposal("proposed");
    state.proposal.collapsed = false;
    state.scenario = "review";
    state.focusAfter = "#editCard .edit-chevron";
    state.scrollToCard = true;
    state.scrollBottom = false;
    announce("12 changes ready for review");
  } else {
    state.scenario = "answer";
    state.focusAfter = document.activeElement?.id === "composer" ? "#composer" : null;
  }
  if (!response.stageProposal) {
    state.scrollBottom = true;
  }
  state.isScrolled = false;
  touchConversation();
  syncUrl();
  render();
  if (!response.stageProposal) announce("Response complete");
}

function stopStream({ silent = false } = {}) {
  if (!state.live) return;
  const live = state.live;
  if (live.timer) clearTimeout(live.timer);
  live.message.streaming = false;
  live.message.stopped = true;
  live.message.status = { state: "stopped", label: "Stopped" };
  live.message.steps = [];
  state.live = null;
  state.streamTimer = null;
  state.scenario = "answer";
  state.scrollBottom = true;
  state.focusAfter = "#composer";
  touchConversation();
  syncUrl();
  if (!silent) {
    render();
    announce("Stopped");
  }
}

function retryError() {
  if (state.offline) {
    state.offline = false;
    state.banner = null;
    state.messages = state.messages.filter((message) => message.role !== "error");
    state.scenario = "answer";
    syncUrl();
    render();
    announce("Reconnected");
    window.setTimeout(() => startStream(responseForPrompt(state.lastPrompt)), 420);
  }
}

function retryMessage(index) {
  const message = state.messages[index];
  if (!message || message.role !== "assistant") return;
  const prompt = state.messages.slice(0, index).reverse().find((item) => item.role === "user")?.text || state.lastPrompt;
  state.messages = state.messages.slice(0, index);
  state.proposal = null;
  state.lastPrompt = prompt;
  state.composerText = "";
  startStream(responseForPrompt(prompt));
  announce("Retrying the turn");
}

/* -------------------------------------------------------------------------
   Send, apply, resolve
   ------------------------------------------------------------------------- */

function ensureCurrentConversation() {
  if (currentConversation()) return currentConversation();
  const id = `c-${Date.now()}`;
  const conversation = {
    id,
    title: "New conversation",
    preview: "No messages yet",
    time: "Just now",
    group: "Today",
    count: 0,
    sheets: "Forecast",
    workbook: workbook().name,
    path: workbook().path,
    current: false,
    pending: false,
  };
  state.conversations.unshift(conversation);
  state.currentId = id;
  return conversation;
}

function sendMessage() {
  if (composerDisabled()) {
    showSnackbar(composerHint()?.text || "The composer is gated right now.", { tone: "warn" });
    return;
  }
  if (isStreaming()) return;
  const prompt = state.composerText.trim();
  if (!prompt) return;
  state.messages.push({ role: "user", text: prompt });
  state.composerText = "";
  state.focusCaret = 0;
  state.lastPrompt = prompt;
  state.scrollBottom = true;
  const current = ensureCurrentConversation();
  if (current) {
    current.title = prompt.slice(0, 58);
    current.preview = "Waiting for the first response…";
    current.count = state.messages.length;
    current.pending = false;
  }
  startStream(responseForPrompt(prompt));
  announce("Message sent. Response started");
}

function fillComposer(prompt) {
  if (state.view !== "chat") {
    state.view = "chat";
    state.focusAfter = "#composer";
  }
  if (composerDisabled()) {
    showSnackbar(composerHint()?.text || "The composer is gated right now.", { tone: "warn" });
    return;
  }
  state.composerText = prompt;
  state.focusAfter = "#composer";
  state.focusCaret = prompt.length;
  render();
}

function touchConversation() {
  const current = currentConversation();
  if (!current) return;
  const firstUser = state.messages.find((message) => message.role === "user")?.text;
  const lastAssistant = [...state.messages].reverse().find((message) => message.role === "assistant");
  if (firstUser) {
    current.title = firstUser.slice(0, 58);
    current.preview = lastAssistant ? stripMarkdown(lastAssistant.text).slice(0, 92) : "No messages yet";
    current.count = state.messages.length + (state.proposal ? 1 : 0);
    current.pending = !!state.proposal && ["proposed", "applying", "partial", "dismissed"].includes(state.proposal.status);
  }
}

function saveCurrentConversation() {
  const current = currentConversation();
  if (!current) return;
  current._messages = clone(state.messages);
  current._proposal = state.proposal ? clone(state.proposal) : null;
  const firstUser = state.messages.find((message) => message.role === "user")?.text;
  const lastAssistant = [...state.messages].reverse().find((message) => message.role === "assistant");
  if (firstUser && !current.customTitle) current.title = firstUser.slice(0, 58);
  if (lastAssistant) current.preview = stripMarkdown(lastAssistant.text).slice(0, 92);
  current.count = state.messages.length + (state.proposal ? 1 : 0);
  current.pending = !!state.proposal && ["proposed", "applying", "partial", "dismissed"].includes(state.proposal.status);
}

function applyProposal() {
  const proposal = state.proposal;
  if (!proposal || proposal.status !== "proposed") return;
  if (overwriteSelectedCells(proposal) > 0 && !proposal.ack) {
    showSnackbar("Tick the overwrite acknowledgement first.", { tone: "warn" });
    state.focusAfter = '[data-action="toggle-ack"]';
    render();
    return;
  }
  if (!selectedCount(proposal)) return;
  proposal.ack = true;
  proposal.status = "applying";
  proposal.collapsed = false;
  proposal.appliedAt = new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  proposal.rows.forEach((row) => {
    row.outcome = row.included ? "queued" : "excluded";
  });
  state.scenario = "review";
  state.scrollBottom = false;
  state.scrollToCard = true;
  syncUrl();
  render();
  announce(`Applying ${selectedCount(proposal)} changes`);
  state.applyTimer = window.setTimeout(() => completeApply(proposal), 1850);
}

function completeApply(proposal) {
  if (!proposal) return;
  state.applyTimer = null;
  proposal.rows.forEach((row) => {
    row.outcome = row.included ? "applied" : "excluded";
  });
  proposal.appliedCount = selectedCount(proposal);
  proposal.excludedCount = proposal.rows.length - proposal.appliedCount;
  proposal.failedCount = 0;
  proposal.status = "applied";
  proposal.collapsed = true;
  state.scenario = "applied";
  state.focusAfter = "#composer";
  state.scrollBottom = false;
  state.scrollToCard = true;
  writeMockCells(proposal.rows);
  touchConversation();
  syncUrl();
  render();
  showSnackbar(`Applied ${proposal.appliedCount} changes. Ctrl+Z would undo this in Excel.`, { tone: "success" });
  announce(`Applied ${proposal.appliedCount} of ${proposal.rows.length} changes`);
}

function retryFailed() {
  const proposal = state.proposal;
  if (!proposal || proposal.status !== "partial") return;
  proposal.status = "applying";
  proposal.collapsed = false;
  proposal.rows.forEach((row) => {
    if (row.outcome === "failed") row.outcome = "queued";
  });
  state.scrollBottom = true;
  syncUrl();
  render();
  announce("Retrying the failed changes");
  state.applyTimer = window.setTimeout(() => completeRetry(proposal), 1500);
}

function completeRetry(proposal) {
  state.applyTimer = null;
  proposal.rows.forEach((row) => {
    if (row.outcome !== "excluded") row.outcome = "applied";
  });
  proposal.appliedCount = proposal.rows.filter((row) => row.outcome === "applied").length;
  proposal.failedCount = 0;
  proposal.excludedCount = proposal.rows.filter((row) => row.outcome === "excluded").length;
  proposal.status = "applied";
  proposal.collapsed = true;
  state.scenario = "applied";
  state.focusAfter = "#composer";
  state.scrollBottom = false;
  state.scrollToCard = true;
  writeMockCells(proposal.rows);
  touchConversation();
  syncUrl();
  render();
  showSnackbar(`Retried the failed rows. Applied ${proposal.appliedCount} changes.`, { tone: "success" });
  announce(`Applied ${proposal.appliedCount} changes after retry`);
}

function dismissProposal() {
  if (!state.proposal || state.proposal.status !== "partial") return;
  state.proposal.status = "dismissed";
  state.proposal.collapsed = true;
  state.scenario = "partial";
  state.focusAfter = "#composer";
  state.scrollBottom = true;
  touchConversation();
  syncUrl();
  render();
  showSnackbar("The failed changes stay pending in History.", { tone: "warn" });
  announce("Changes dismissed; failed rows remain pending");
}

function discardProposal() {
  if (!state.proposal) return;
  const count = selectedCount(state.proposal);
  openDialog({
    title: "Discard proposed changes?",
    body: `Discard ${count} proposed change${count === 1 ? "" : "s"}? Nothing has been written yet.`,
    confirmLabel: "Discard",
    danger: true,
    onConfirm: () => {
      state.proposal = null;
      state.scenario = "answer";
      state.focusAfter = "#composer";
      state.scrollBottom = true;
      syncUrl();
      render();
      showSnackbar("Discarded. Nothing was written.", { tone: "neutral" });
      announce("Proposal discarded");
    },
  });
}

function writeMockCells(rows) {
  for (const row of rows) {
    if (!row.included || row.outcome !== "applied") continue;
    if (row.sheet === "Forecast" && row.range === "B1") setCell("Forecast", "B1", "1.124");
    if (row.sheet === "Forecast" && row.range === "H8") setCell("Forecast", "H8", "1.124");
    if (row.sheet === "Inputs" && row.range === "B2") setCell("Inputs", "B2", "1.124");
  }
}

function setCell(sheet, address, value) {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match || !state.sheets[sheet]) return;
  const col = match[1].charCodeAt(0) - 65;
  const row = Number(match[2]) - 1;
  if (!state.sheets[sheet].rows[row]) state.sheets[sheet].rows[row] = [];
  state.sheets[sheet].rows[row][col] = value;
}

function revealApplied() {
  if (!state.proposal) return;
  const applied = state.proposal.rows.find((row) => row.outcome === "applied" || row.outcome === "excluded");
  if (!applied) return;
  navigateToSelection(applied.sheet, applied.range);
  const count = state.proposal.rows.filter((row) => row.outcome === "applied").length;
  if (count > 1) showSnackbar(`Selected the first of ${count} applied ranges.`, { tone: "neutral" });
}

/* -------------------------------------------------------------------------
   History navigation and deletion
   ------------------------------------------------------------------------- */

function scenarioForProposal(proposal) {
  if (!proposal) return null;
  if (proposal.status === "proposed" || proposal.status === "applying") return "review";
  if (proposal.status === "applied") return "applied";
  if (proposal.status === "partial" || proposal.status === "dismissed") return "partial";
  return null;
}

function loadConversation(target) {
  state.currentId = target.id;
  state.messages = target._messages ? clone(target._messages) : decorateThreadMessages(target.thread || []);
  state.proposal = target._proposal ? clone(target._proposal) : target.scenario === "partial" ? makeProposal("partial") : null;
  state.compaction = null;
  state.view = "chat";
  state.loading = false;
  state.offline = false;
  state.banner = null;
  state.composerText = "";
  state.scenario = scenarioForProposal(state.proposal) || (state.messages.length ? "answer" : "empty");
  state.focusAfter = proposalGatesComposer(state.proposal) ? "#editCard .edit-chevron" : "#composer";
  state.scrollBottom = false;
  state.scrollToCard = !!state.proposal;
  syncUrl();
  render();
  announce(`Opened ${target.title}`);
}

function openConversation(id) {
  if (state.proposal?.status === "applying") {
    showSnackbar("Wait for the apply to finish before switching conversations.", { tone: "warn" });
    return;
  }
  const target = state.conversations.find((conversation) => conversation.id === id);
  if (!target) return;
  if (target.workbook !== workbook().name) {
    showSnackbar(`That conversation belongs to ${target.workbook}. Open that workbook to continue it.`, { tone: "warn" });
    return;
  }
  if (isStreaming()) {
    openDialog({
      title: "Stop the current response?",
      body: "The partial answer will stay in this conversation. The new conversation will open at the same scroll position.",
      confirmLabel: "Stop and switch",
      onConfirm: () => {
        stopStream({ silent: true });
        saveCurrentConversation();
        loadConversation(target);
      },
    });
    return;
  }
  if (id === state.currentId) {
    state.view = "chat";
    state.focusAfter = "#composer";
    render();
    return;
  }
  saveCurrentConversation();
  loadConversation(target);
}

function startNewChat() {
  if (!state.messages.length && !state.proposal) {
    state.view = "chat";
    state.focusAfter = "#composer";
    render();
    return;
  }
  saveCurrentConversation();
  cancelLiveStream();
  const id = `c-${Date.now()}`;
  state.conversations.unshift({
    id,
    title: "New conversation",
    preview: "No messages yet",
    time: "Just now",
    group: "Today",
    count: 0,
    sheets: "Forecast",
    workbook: workbook().name,
    path: workbook().path,
    current: false,
    pending: false,
  });
  state.currentId = id;
  state.messages = [];
  state.proposal = null;
  state.compaction = null;
  state.view = "chat";
  state.composerText = "";
  state.offline = false;
  state.banner = null;
  state.scenario = "empty";
  state.focusAfter = "#composer";
  state.scrollBottom = true;
  syncUrl();
  render();
  announce("New chat");
}

function requestNewChat() {
  if (isStreaming()) {
    openDialog({
      title: "Stop the current response?",
      body: "Starting a new chat keeps this partial answer in History.",
      confirmLabel: "Stop and start new",
      onConfirm: () => {
        stopStream({ silent: true });
        startNewChat();
      },
    });
    return;
  }
  if (proposalGatesComposer()) {
    showSnackbar("Apply or discard the pending changes first.", { tone: "warn" });
    return;
  }
  startNewChat();
}

function renameConversation(id) {
  const target = state.conversations.find((conversation) => conversation.id === id);
  if (!target) return;
  state.rowMenuId = null;
  openDialog({
    title: "Rename conversation",
    body: "This only changes the label in local history.",
    input: true,
    inputLabel: "Conversation name",
    inputValue: target.customTitle || target.title,
    confirmLabel: "Rename",
    onConfirm: (value) => {
      if (value?.trim()) {
        target.customTitle = value.trim();
        target.title = value.trim();
        showSnackbar("Conversation renamed.", { tone: "neutral" });
      }
      render();
    },
  });
}

function deleteConversation(id) {
  const target = state.conversations.find((conversation) => conversation.id === id);
  if (!target) return;
  state.rowMenuId = null;
  openDialog({
    title: "Delete this conversation?",
    body: "It can’t be undone. Workbook content is not changed.",
    confirmLabel: "Delete",
    danger: true,
    onConfirm: () => {
      state.conversations = state.conversations.filter((conversation) => conversation.id !== id);
      if (state.currentId === id) {
        state.currentId = null;
        state.messages = [];
        state.proposal = null;
        state.scenario = state.view === "history" ? "history" : "empty";
      }
      syncUrl();
      render();
      showSnackbar("Conversation deleted.", { tone: "neutral" });
      announce("Conversation deleted");
    },
  });
}

function deleteAllHistory() {
  const count = state.historyRows.reduce((total, row) => total + row.conversations, 0);
  const size = historyUsedMb();
  if (!count) return;
  openDialog({
    title: `Delete all ${count} conversations?`,
    body: `This removes ${size} MB of local transcripts and audit records. Workbook content is not changed. It can’t be undone.`,
    confirmLabel: "Delete all",
    danger: true,
    onConfirm: () => {
      state.conversations = [];
      state.historyRows = [];
      state.currentId = null;
      state.messages = [];
      state.proposal = null;
      state.view = state.view === "settings" ? "settings" : "history";
      state.scenario = state.view === "settings" ? "settings" : "history";
      state.scrollViewTo = state.view === "settings" ? "#historySection" : null;
      state.focusAfter = "#viewTitle";
      syncUrl();
      render();
      showSnackbar("Local history deleted.", { tone: "neutral" });
      announce("All local history deleted");
    },
  });
}

/* -------------------------------------------------------------------------
   Snackbar, dialog, clipboard
   ------------------------------------------------------------------------- */

function showSnackbar(text, { tone = "neutral", duration = 4200, action = null } = {}) {
  if (state.snackbarTimer) clearTimeout(state.snackbarTimer);
  state.snackbar = { text, tone, action };
  const slot = $("#snackbarSlot");
  if (slot) {
    slot.innerHTML = renderSnackbarContent();
  } else {
    render();
  }
  if (duration > 0) {
    state.snackbarTimer = window.setTimeout(() => {
      state.snackbar = null;
      const current = $("#snackbarSlot");
      if (current) current.innerHTML = "";
    }, duration);
  }
  announce(text);
}

function dismissSnackbar() {
  if (state.snackbarTimer) clearTimeout(state.snackbarTimer);
  state.snackbar = null;
  const slot = $("#snackbarSlot");
  if (slot) slot.innerHTML = "";
}

function openDialog(dialog) {
  state.modelMenuOpen = false;
  state.workbookMenuOpen = false;
  state.rowMenuId = null;
  state.dialog = { ...dialog };
  render();
}

function closeDialog({ focus = true } = {}) {
  state.dialog = null;
  if (focus) state.focusAfter = "#composer";
  render();
}

function confirmDialog() {
  const dialog = state.dialog;
  if (!dialog) return;
  const value = $("#dialogInput")?.value;
  state.dialog = null;
  render();
  dialog.onConfirm?.(value);
}

async function copyText(text) {
  const value = String(text ?? "");
  if (!value) return;
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(value);
    } else {
      const area = document.createElement("textarea");
      area.value = value;
      area.setAttribute("readonly", "");
      area.style.position = "fixed";
      area.style.left = "-9999px";
      document.body.appendChild(area);
      area.select();
      document.execCommand("copy");
      area.remove();
    }
    showSnackbar("Copied.", { tone: "success" });
  } catch {
    showSnackbar("Couldn’t copy. Select the text instead.", { tone: "warn" });
  }
}

/* -------------------------------------------------------------------------
   Settings and chrome actions
   ------------------------------------------------------------------------- */

function openHistory() {
  state.view = "history";
  state.scenario = "history";
  syncUrl();
  state.modelMenuOpen = false;
  state.workbookMenuOpen = false;
  state.settingsModelOpen = false;
  state.rowMenuId = null;
  state.focusAfter = "#viewTitle";
  render();
  announce("History");
}

function openSettings(section = null) {
  state.view = "settings";
  state.scenario = "settings";
  syncUrl();
  state.modelMenuOpen = false;
  state.workbookMenuOpen = false;
  state.rowMenuId = null;
  state.settingsModelOpen = false;
  state.scrollViewTo = section === "history" ? "#historySection" : null;
  state.focusAfter = section === "history" ? "#historySection" : "#viewTitle";
  render();
  announce("Settings");
}

function openLink(from = "history") {
  if (state.unsaved) {
    showSnackbar("Save the workbook first.", { tone: "warn" });
    return;
  }
  state.view = "link";
  state.scenario = "link";
  state.linkFrom = from;
  syncUrl();
  state.workbookMenuOpen = false;
  state.modelMenuOpen = false;
  state.rowMenuId = null;
  state.linkFilter = "";
  state.focusAfter = "#linkFilter";
  render();
  announce("Link to an old conversation");
}

function backFromLink() {
  if (state.linkFrom === "menu") backToChat();
  else openHistory();
}

function openInstructions() {
  state.view = "settings";
  state.scenario = "settings";
  syncUrl();
  state.workbookMenuOpen = false;
  state.settingsModelOpen = false;
  state.scrollViewTo = null;
  state.focusAfter = "#instructions";
  render();
  announce("Instructions for this workbook");
}

function backToChat() {
  state.view = "chat";
  if (["history", "settings", "link"].includes(state.scenario)) {
    state.scenario = scenarioForProposal(state.proposal) || (state.messages.length ? "answer" : "empty");
    syncUrl();
  }
  state.focusAfter = "#composer";
  state.scrollBottom = true;
  render();
  announce("Conversation");
}

/* -------------------------------------------------------------------------
   #16: linking, history budget, store failure, compaction, read-only
   ------------------------------------------------------------------------- */

function linkConversation(id) {
  const target = state.conversations.find((conversation) => conversation.id === id);
  if (!target || target.workbook === workbook().name) return;
  const previous = {
    workbook: target.workbook,
    path: target.path,
    group: target.group,
    time: target.time,
    date: target.date,
  };
  target.workbook = workbook().name;
  target.path = workbook().path;
  target.group = "Today";
  target.time = "Just now";
  target.date = new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
  target.movedFrom = previous.workbook;
  target.movedFromPath = previous.path;
  target.movedFromGroup = previous.group;
  target.movedFromTime = previous.time;
  target.movedFromDate = previous.date;
  state.lastLink = { id, previous };
  state.focusAfter = "#linkFilter";
  render();
  showSnackbar("Moved to this workbook.", {
    tone: "success",
    duration: 8000,
    action: { label: "Undo", run: undoLastLink },
  });
  announce(`Moved ${target.title} to this workbook`);
}

function restoreLink(id) {
  const target = state.conversations.find((conversation) => conversation.id === id);
  const previous = state.lastLink?.id === id ? state.lastLink.previous : null;
  if (!target || !target.movedFrom) return null;
  const original = previous || {
    workbook: target.movedFrom,
    path: target.movedFromPath || "",
    group: target.movedFromGroup || "Earlier",
    time: target.movedFromTime || "—",
    date: target.movedFromDate || "",
  };
  const from = target.movedFrom;
  target.workbook = original.workbook;
  target.path = original.path;
  target.group = original.group;
  target.time = original.time;
  target.date = original.date;
  delete target.movedFrom;
  delete target.movedFromPath;
  delete target.movedFromGroup;
  delete target.movedFromTime;
  delete target.movedFromDate;
  if (state.lastLink?.id === id) state.lastLink = null;
  return from;
}

function undoLastLink() {
  const id = state.lastLink?.id;
  if (!id) return;
  const from = restoreLink(id);
  state.focusAfter = "#linkFilter";
  render();
  showSnackbar(`Moved back to ${from}.`, { tone: "neutral" });
  announce(`Moved back to ${from}`);
}

function unlinkConversation(id) {
  const from = restoreLink(id);
  state.rowMenuId = null;
  if (!from) return;
  render();
  showSnackbar(`Moved back to ${from}.`, { tone: "neutral" });
  announce(`Moved back to ${from}`);
}

function openFolder() {
  showSnackbar(`Opened ${HISTORY_STORE_PATH}.`, { tone: "neutral" });
  announce("Opened the history folder");
}

function copyStorePath() {
  copyText(HISTORY_STORE_PATH);
}

function exportConversation(id, format) {
  const target = state.conversations.find((conversation) => conversation.id === id);
  if (!target) return;
  state.rowMenuId = null;
  const safe = target.title.replace(/[\\/:*?"<>|]/g, "-");
  render();
  showSnackbar(`Exported ${safe}.${format}.`, { tone: "success" });
  announce(`Exported ${target.title} as ${format}`);
}

function deleteWorkbookHistory(key) {
  const row = state.historyRows.find((item) => item.key === key);
  if (!row) return;
  openDialog({
    title: `Delete ${row.conversations} conversation${row.conversations === 1 ? "" : "s"}?`,
    body: `This removes ${row.sizeMb} MB from ${row.name}. Workbook content is not changed. It can’t be undone.`,
    confirmLabel: "Delete",
    danger: true,
    onConfirm: () => {
      state.conversations = state.conversations.filter((conversation) => conversation.workbook !== row.name);
      state.historyRows = state.historyRows.filter((item) => item.key !== key);
      if (state.historyRows.length === 0) {
        state.currentId = null;
        state.messages = [];
        state.proposal = null;
      }
      state.view = "settings";
      state.scenario = "retention";
      state.scrollViewTo = "#historySection";
      state.focusAfter = "#historySection";
      syncUrl();
      render();
      showSnackbar(`Deleted ${row.name} from local history.`, { tone: "neutral" });
      announce(`Deleted history for ${row.name}`);
    },
  });
}

function retryStore() {
  state.storeFailure = false;
  state.banner = null;
  render();
  showSnackbar("History is being saved again.", { tone: "success" });
  announce("History is being saved again");
}

function summarizeCompaction() {
  const compaction = state.compaction;
  if (!compaction || compaction.summarizing) return;
  compaction.summarizing = true;
  state.focusAfter = ".compaction-notice";
  render();
  announce(`Summarizing ${compaction.pending?.covered || 0} turns`);
  window.setTimeout(() => {
    if (!state.compaction || state.compaction !== compaction) return;
    if (compaction.pending) {
      compaction.cuts.push({
        start: compaction.pending.start,
        end: compaction.pending.end,
        covered: compaction.pending.covered,
        expanded: false,
        summary: "Earlier decisions, findings and the ranges that matter.",
      });
    }
    compaction.summarizing = false;
    compaction.notice = false;
    compaction.pending = null;
    state.focusAfter = '[data-action="toggle-original"]';
    syncUrl();
    render();
    showSnackbar("Earlier turns summarized. The originals are still here.", { tone: "success" });
    announce("Earlier turns summarized; originals kept");
  }, 2600);
}

function toggleOriginal(cutIndex) {
  if (!state.compaction?.cuts?.[cutIndex]) return;
  const cut = state.compaction.cuts[cutIndex];
  cut.expanded = !cut.expanded;
  state.focusAfter = `[data-action="toggle-original"][data-cut="${cutIndex}"]`;
  render();
  announce(cut.expanded ? "Original turns shown read-only" : "Original turns hidden");
}

function copyInstructions() {
  const source = state.instructionsSource;
  if (!source) return;
  state.instructions = source.text;
  state.instructionsSource = null;
  state.instructionsState = "saved";
  state.focusAfter = "#instructions";
  render();
  showSnackbar(`Copied instructions from ${source.workbook}.`, { tone: "success" });
  announce("Instructions copied");
}

function toggleReadOnly() {
  if (state.workbookReadOnly) {
    showSnackbar("This workbook is open read-only.", { tone: "warn" });
    return;
  }
  if (state.proposal && proposalGatesComposer()) {
    showSnackbar("Apply or discard the pending changes first.", { tone: "warn" });
    return;
  }
  state.readOnly = !state.readOnly;
  state.readOnlyJustSaved = true;
  state.focusAfter = "#readOnlySwitch";
  render();
  showSnackbar(
    state.readOnly ? "Read-only mode on. Edits are unavailable." : "Read-only mode off. Edits are available.",
    { tone: "success" },
  );
  announce(state.readOnly ? "Read-only mode on" : "Read-only mode off");
}

function setTheme(mode) {
  if (!THEMES.includes(mode)) return;
  state.themeMode = mode;
  state.settingsModelOpen = false;
  state.focusAfter = state.view === "settings" ? ".theme-segment.is-active" : ".proto-theme button.is-active";
  syncUrl();
  render();
  announce(`Theme ${mode}`);
}

function toggleAutoOpen() {
  state.autoOpen = !state.autoOpen;
  state.focusAfter = '[data-action="toggle-auto-open"]';
  render();
  showSnackbar(
    state.autoOpen
      ? `Added Office.AutoShowTaskpaneWithDocument to ${workbook().name}.`
      : `Removed the auto-open setting from ${workbook().name}.`,
    { tone: state.autoOpen ? "success" : "neutral" },
  );
  announce(`Open with workbook ${state.autoOpen ? "on" : "off"}`);
}

function selectModel(name) {
  if (!state.models.some((model) => model.name === name)) return;
  state.model = name;
  state.modelMenuOpen = false;
  state.settingsModelOpen = false;
  state.focusAfter = state.view === "settings" ? "#viewTitle" : "#composer";
  render();
  showSnackbar(`Model set to ${name} for new messages.`, { tone: "success" });
  announce(`Model ${name}`);
}

function refreshModels() {
  const token = Symbol("models");
  state.modelRefreshToken = token;
  state.modelsLoading = true;
  render();
  window.setTimeout(() => {
    if (state.modelRefreshToken !== token) return;
    state.modelsLoading = false;
    state.models = clone(MODELS);
    state.modelUpdatedAt = "just now";
    render();
    showSnackbar("Model list refreshed from the gateway.", { tone: "success" });
    announce("Model list refreshed");
  }, 720);
}

function cycleVariant(direction) {
  const index = VARIANTS.indexOf(state.variant);
  state.variant = VARIANTS[(index + direction + VARIANTS.length) % VARIANTS.length];
  state.workbookMenuOpen = false;
  state.modelMenuOpen = false;
  state.rowMenuId = null;
  syncUrl();
  render();
  announce(`Variant ${VARIANT_META[state.variant].name}`);
}

function cycleTheme() {
  const index = THEMES.indexOf(state.themeMode);
  setTheme(THEMES[(index + 1) % THEMES.length]);
}

/* -------------------------------------------------------------------------
   Event handling
   ------------------------------------------------------------------------- */

function handleAction(element) {
  const action = element.dataset.action;
  switch (action) {
    case "prev-variant":
      cycleVariant(-1);
      break;
    case "next-variant":
      cycleVariant(1);
      break;
    case "reset-prototype":
      loadScenario(state.scenario);
      render();
      showSnackbar("Scenario reset.", { tone: "neutral" });
      break;
    case "set-theme":
      setTheme(element.dataset.theme);
      break;
    case "toggle-workbook-menu":
      state.workbookMenuOpen = !state.workbookMenuOpen;
      state.modelMenuOpen = false;
      state.rowMenuId = null;
      if (state.workbookMenuOpen) state.focusAfter = '.workbook-popover [role="menuitem"]';
      render();
      break;
    case "copy-path":
      state.workbookMenuOpen = false;
      render();
      copyText(workbook().path);
      break;
    case "open-instructions":
      openInstructions();
      break;
    case "toggle-auto-open":
      toggleAutoOpen();
      break;
    case "open-history":
      openHistory();
      break;
    case "open-settings":
      openSettings();
      break;
    case "open-settings-history":
      openSettings("history");
      break;
    case "open-readonly-settings":
      openSettings();
      state.scrollViewTo = "#editingSection";
      state.focusAfter = "#readOnlySwitch";
      render();
      break;
    case "open-link":
      openLink(element.dataset.from || "history");
      break;
    case "back-link":
      backFromLink();
      break;
    case "clear-link-filter":
      state.linkFilter = "";
      state.focusAfter = "#linkFilter";
      render();
      break;
    case "link-conversation":
      linkConversation(element.dataset.id);
      break;
    case "unlink-conversation":
      unlinkConversation(element.dataset.id);
      break;
    case "export-conversation":
      exportConversation(element.dataset.id, element.dataset.format || "md");
      break;
    case "delete-workbook-history":
      deleteWorkbookHistory(element.dataset.key);
      break;
    case "open-folder":
      openFolder();
      break;
    case "copy-store-path":
      copyStorePath();
      break;
    case "retry-store":
      retryStore();
      break;
    case "summarize":
      summarizeCompaction();
      break;
    case "toggle-original":
      toggleOriginal(Number(element.dataset.cut));
      break;
    case "copy-instructions":
      copyInstructions();
      break;
    case "toggle-readonly":
      toggleReadOnly();
      break;
    case "back-to-chat":
      backToChat();
      break;
    case "new-chat":
      requestNewChat();
      break;
    case "review-card": {
      if (!state.proposal) break;
      state.view = "chat";
      render();
      requestAnimationFrame(() => {
        const card = $("#editCard");
        card?.scrollIntoView({ block: "center", behavior: "smooth" });
        $("#editCard .edit-chevron")?.focus({ preventScroll: true });
      });
      break;
    }
    case "suggest":
      fillComposer(element.dataset.prompt || "");
      break;
    case "focus-composer":
      if (composerDisabled()) {
        showSnackbar(composerHint()?.text || "The composer is gated right now.", { tone: "warn" });
      } else {
        $("#composer")?.focus();
      }
      break;
    case "send":
      sendMessage();
      break;
    case "stop":
      stopStream();
      break;
    case "copy-message": {
      const message = state.messages[Number(element.dataset.index)];
      if (message) copyText(message.text || message.title || "");
      break;
    }
    case "copy-code":
      copyText(element.dataset.copy || "");
      break;
    case "copy-details":
      copyText(element.dataset.copy || "Excel Agent 0.1.0 · local host · req_8f3a");
      break;
    case "retry-message":
      retryMessage(Number(element.dataset.index));
      break;
    case "retry-error":
      retryError();
      break;
    case "toggle-steps": {
      const index = Number(element.dataset.index);
      const message = state.messages[index];
      if (message) {
        message.showSteps = !message.showSteps;
        render();
      }
      break;
    }
    case "toggle-card":
      if (state.proposal) {
        state.proposal.collapsed = !state.proposal.collapsed;
        state.focusAfter = "#editCard .edit-chevron";
        render();
      }
      break;
    case "toggle-row": {
      if (!state.proposal || state.proposal.status !== "proposed") break;
      const row = state.proposal.rows.find((item) => item.id === element.dataset.rowId);
      if (row) {
        row.included = !row.included;
        state.focusAfter = `[data-action="toggle-row"][data-row-id="${row.id}"]`;
        render();
      }
      break;
    }
    case "toggle-ack":
      if (state.proposal) {
        state.proposal.ack = !state.proposal.ack;
        state.focusAfter = '[data-action="toggle-ack"]';
        render();
      }
      break;
    case "toggle-all-rows":
      if (state.proposal) {
        state.proposal.showAllRows = !state.proposal.showAllRows;
        state.focusAfter = '[data-action="toggle-all-rows"]';
        render();
      }
      break;
    case "toggle-applied":
      if (state.proposal) {
        state.proposal.showApplied = !state.proposal.showApplied;
        state.focusAfter = '[data-action="toggle-applied"]';
        render();
      }
      break;
    case "apply":
      applyProposal();
      break;
    case "discard-proposal":
      discardProposal();
      break;
    case "retry-failed":
      retryFailed();
      break;
    case "dismiss-proposal":
      dismissProposal();
      break;
    case "reveal-applied":
      revealApplied();
      break;
    case "toggle-model-menu":
      state.modelMenuOpen = !state.modelMenuOpen;
      state.workbookMenuOpen = false;
      render();
      if (state.modelMenuOpen) {
        requestAnimationFrame(() => {
          $(".model-row.is-selected")?.focus({ preventScroll: true });
        });
      }
      break;
    case "select-model":
      selectModel(element.dataset.model);
      break;
    case "refresh-models":
      refreshModels();
      break;
    case "toggle-settings-model":
      state.settingsModelOpen = !state.settingsModelOpen;
      state.focusAfter = state.settingsModelOpen ? ".settings-model-row.is-selected" : "#viewTitle";
      render();
      break;
    case "history-scope":
      state.historyScope = element.dataset.scope === "all" ? "all" : "workbook";
      state.focusAfter = ".scope-segment.is-active";
      render();
      announce(`History scope: ${state.historyScope === "all" ? "all workbooks" : "this workbook"}`);
      break;
    case "open-conversation":
      openConversation(element.dataset.id);
      break;
    case "row-menu":
      state.rowMenuId = state.rowMenuId === element.dataset.id ? null : element.dataset.id;
      if (state.rowMenuId) state.focusAfter = '.row-menu [role="menuitem"]';
      render();
      break;
    case "rename-conversation":
      renameConversation(element.dataset.id);
      break;
    case "delete-conversation":
      deleteConversation(element.dataset.id);
      break;
    case "delete-all":
      deleteAllHistory();
      break;
    case "select-sheet":
      state.selection = { sheet: element.dataset.sheet, range: "A1", active: "A1" };
      state.excelStatus = `Selected ${state.selection.sheet}!A1`;
      refreshGridDom();
      announce(`${state.selection.sheet} sheet`);
      break;
    case "citation": {
      const target = element.dataset.cite || "";
      const bang = target.lastIndexOf("!");
      const sheet = bang >= 0 ? target.slice(0, bang) : state.selection.sheet;
      const range = bang >= 0 ? target.slice(bang + 1) : target;
      navigateToSelection(sheet, range);
      break;
    }
    case "dismiss-snackbar":
      dismissSnackbar();
      break;
    case "snackbar-action": {
      const snack = state.snackbar;
      dismissSnackbar();
      snack?.action?.run?.();
      break;
    }
    case "dialog-cancel":
      closeDialog();
      break;
    case "dialog-confirm":
      confirmDialog();
      break;
    case "dialog-backdrop":
      if (element.classList.contains("dialog-scrim")) closeDialog();
      break;
    default:
      break;
  }
}

function handleClick(event) {
  const actionElement = event.target.closest("[data-action]");
  if (actionElement && !actionElement.disabled) {
    if (actionElement.dataset.action === "dialog-backdrop" && event.target !== actionElement) {
      // Clicks inside the dialog should not dismiss it.
    } else {
      handleAction(actionElement);
    }
  }

  const cell = event.target.closest(".cell[data-cell]");
  if (cell) {
    const sheet = cell.dataset.sheet;
    const address = cell.dataset.cell;
    state.selection = { sheet, range: address, active: address };
    state.excelStatus = `Selected ${sheet}!${address}`;
    refreshGridDom();
  }

  if (state.dialog) return;
  if (state.modelMenuOpen && !event.target.closest(".model-popover") && !event.target.closest('[data-action="toggle-model-menu"]')) {
    state.modelMenuOpen = false;
    render();
  }
  if (state.workbookMenuOpen && !event.target.closest(".workbook-popover") && !event.target.closest('[data-action="toggle-workbook-menu"]')) {
    state.workbookMenuOpen = false;
    render();
  }
  if (state.rowMenuId && !event.target.closest(".history-row-wrap")) {
    state.rowMenuId = null;
    render();
  }
}

function handleInput(event) {
  if (event.target.id === "linkFilter") {
    state.linkFilter = event.target.value;
    const list = $("#linkList");
    if (list) list.innerHTML = renderLinkList();
    return;
  }
  if (event.target.id === "composer") {
    state.composerText = event.target.value;
    autoGrowComposer();
    const send = $('[data-action="send"]');
    if (send) send.disabled = !state.composerText.trim() || composerDisabled();
    return;
  }
  if (event.target.id === "instructions") {
    state.instructions = event.target.value;
    state.instructionsState = state.instructions.length > 8000 ? "error" : "saving";
    updateInstructionsDom();
    if (state.saveTimer) clearTimeout(state.saveTimer);
    if (state.instructionsState === "saving") {
      state.saveTimer = window.setTimeout(() => {
        state.instructionsState = "saved";
        updateInstructionsDom();
        announce("Instructions saved");
      }, 680);
    }
  }
}

function updateInstructionsDom() {
  const counter = $("#instructionsCount");
  if (counter) {
    counter.textContent = `${state.instructions.length.toLocaleString()} / 8,000`;
    counter.classList.toggle("is-over", state.instructions.length > 8000);
  }
  const saved = $("#instructionsState");
  if (saved) {
    if (state.instructionsState === "saving") saved.textContent = "Saving…";
    else if (state.instructionsState === "error") saved.textContent = "Over the 8,000 character limit";
    else saved.innerHTML = `${icon("check")} Saved`;
    saved.classList.toggle("is-error", state.instructionsState === "error");
  }
}

function handleChange(event) {
  if (event.target.id === "scenarioSelect") {
    const next = event.target.value;
    if (!SCENARIOS.some((scenario) => scenario.key === next)) return;
    loadScenario(next);
    render();
    announce(`Scenario ${next}`);
    return;
  }
  if (event.target.id === "widthSelect") {
    const width = Number(event.target.value);
    if (!Number.isFinite(width)) return;
    state.width = Math.max(260, Math.min(720, width));
    syncUrl();
    render();
    announce(`Pane width ${state.width} pixels`);
  }
}

function handleKeydown(event) {
  if (state.dialog) {
    if (event.key === "Escape") {
      event.preventDefault();
      closeDialog();
      return;
    }
    if (event.key === "Enter" && event.target.id === "dialogInput") {
      event.preventDefault();
      confirmDialog();
      return;
    }
    return;
  }

  if (event.key === "Escape") {
    if (state.modelMenuOpen || state.workbookMenuOpen || state.rowMenuId) {
      state.modelMenuOpen = false;
      state.workbookMenuOpen = false;
      state.rowMenuId = null;
      render();
      return;
    }
    if (isStreaming()) {
      event.preventDefault();
      stopStream();
      return;
    }
    if (state.snackbar) dismissSnackbar();
    return;
  }

  if (event.target.id === "composer" && event.key === "Enter" && !event.shiftKey && !event.isComposing) {
    event.preventDefault();
    if (!isStreaming()) sendMessage();
    return;
  }

  if (isTextInput(event.target) || event.metaKey || event.ctrlKey || event.altKey) return;
  if (state.modelMenuOpen || state.workbookMenuOpen || state.rowMenuId) return;

  if (event.key === "ArrowLeft") {
    event.preventDefault();
    cycleVariant(-1);
  } else if (event.key === "ArrowRight") {
    event.preventDefault();
    cycleVariant(1);
  }
}

function handleScroll(event) {
  if (event.target.classList?.contains("view-body")) {
    viewScrollMemory = event.target.scrollTop;
    return;
  }
  if (!event.target.classList?.contains("transcript")) return;
  const scrolled = event.target.scrollTop > 0;
  state.isScrolled = scrolled;
  $(".pane-header")?.classList.toggle("is-scrolled", scrolled);
}

/* -------------------------------------------------------------------------
   Boot
   ------------------------------------------------------------------------- */

document.addEventListener("click", handleClick);
document.addEventListener("input", handleInput);
document.addEventListener("change", handleChange);
document.addEventListener("keydown", handleKeydown);
document.addEventListener("scroll", handleScroll, true);
prefersDark.addEventListener("change", () => {
  if (state.themeMode === "auto") render();
});

loadScenario(state.scenario, { initial: true });
syncUrl();
render();

// Throwaway debug handle. Handy when testing the prototype; not production.
window.__excelAgentPrototype = {
  get state() { return state; },
};
