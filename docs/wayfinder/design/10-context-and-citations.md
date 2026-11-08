# Workbook context and cell citations

Ticket: [#10 · Workbook context and cell citations](https://github.com/mustaphaelou/excel-agent/issues/10)
Decided with the dev on 2026-11-08. Builds on the capability envelope ([#2](https://github.com/mustaphaelou/excel-agent/issues/2), [fact sheet](../research/01-officejs-capability-envelope.md)), the pane spike ([#6](https://github.com/mustaphaelou/excel-agent/issues/6), [record](../research/06-pane-sideload-spike.md)), the sidebar contract ([#8](https://github.com/mustaphaelou/excel-agent/issues/8), [design](08-sidebar-ux.md)), and the safety model ([#9](https://github.com/mustaphaelou/excel-agent/issues/9), [design](09-edit-safety.md)). #8's C17 citation chip and #9's read/write split are the UI and the policy; this record is the data contract between the workbook and the model.

---

## Summary

The model sees a map, not a dump. Every turn it receives a compact **inventory** — sheets, visibility, used ranges, tables, pivots, named ranges, selection — and cell data only through **paged reads** that name their sheet, address and revision. There is no full-workbook snapshot and no hidden pre-read.

Citations are sheet-qualified A1 addresses emitted as **structured tokens**, validated against the live workbook, and navigated with `activate`/`select` only. The protocol is **en-US canonical**; the pane localizes what it displays, never what it exchanges.

Caps are **2,000 cells per read** and **60K tokens of workbook data per turn**. Hidden sheets are readable and inventoried but never navigable. Staleness is surfaced with a revision counter, not guessed at. A citation is an address, not a tracked range: it is point-in-time, and the pane never silently re-maps it.

### Decisions at a glance

1. **Inventory + paged reads.** A compact inventory every turn, selection attached (contents inline at ≤500 cells), cell data on demand. No snapshot.
2. **No hidden pre-read.** The model calls `read_range` itself; the status line narrates the read.
3. **Citation token.** `[label](<cite:Sheet!A1>)`, always sheet-qualified, one sheet per token, multi-area comma-separated, named ranges resolved to addresses.
4. **Hidden sheets refuse.** The chip renders; the click explains. No unhide, no nearest-visible.
5. **Locale-proof protocol.** en-US addresses, formulas, format codes and ISO dates on the wire; `text`/`addressLocal`/`numberFormatLocal` only in the pane.
6. **Dependencies are tools.** Direct precedents/dependents by default, transitive behind a flag, capped; never auto-annotated onto a read.
7. **Raw + formula + format.** Dates become ISO, formats become a run-length legend, display text only where the legend does not already say it.
8. **Caps and budget.** 2,000 cells/read, 1,000 chars/cell, 60K tokens and 12,000 cells per turn, inventory capped at 200/100/100/200.
9. **Revision-tracked staleness.** `onChanged`/`onFormulaChanged`/`onCalculated` feed one counter; the model is told what changed; it re-reads on its own.
10. **Inventory depth.** Sheets, tables with headers, pivot output ranges, named items, selection. Merges per read block; charts and validation are not context.
11. **Broken citations are text.** Unparseable or missing targets never become chips; deleted, hidden and cell-edit clicks explain themselves.

---

## 1. Invariants

- **I1 — The model sees what it asked for.** The inventory is a map; cell values enter context only through a read tool call. There is no background workbook scanner.
- **I2 — Every cell is traceable.** Each block carries its sheet, address and revision. Nothing the model saw is unattributable to a tool result or the inventory.
- **I3 — Reads are bounded and honest.** The pane enforces the caps. A truncated read says it was truncated and offers the next address; nothing is silently dropped.
- **I4 — Citations are addresses, not ranges.** A citation is validated against the workbook, not against model memory. It is point-in-time and is never silently re-mapped after a structural edit.
- **I5 — Only structured tokens become chips.** Prose that looks like a cell reference never auto-links (#8).
- **I6 — Navigation never mutates.** No unhide, no write, no grid markup on a citation click. Select and activate only.
- **I7 — The protocol is locale-proof.** en-US canonical in both directions; localization is a display concern.
- **I8 — Staleness is surfaced, not guessed.** The pane tracks a revision, tells the model what moved, and re-validates at click. It never auto-re-reads and never pretends a stale block is current.

---

## 2. Context model

### 2.1 What is attached every turn

| Part | Contents | Cost |
| --- | --- | --- |
| **Inventory** | Workbook name/URL/read-only/dirty, display and content language, every sheet (id, name, visibility, used range, dimensions), tables (name, range, headers), pivots (name, output range), named items (name, formula, scope, visible), selection | 1–3K tokens typical; ≤10K at the caps |
| **Selection** | Address and areas always; the values block inline when the selection is ≤500 cells | ≤3K tokens |
| **System notes** | "Workbook changed since your last read: …" when the revision moved; budget warnings | <200 tokens |

The inventory is rebuilt at the start of every turn (one batched `Excel.run`), so it is current even after the pane was idle. It is cached for the duration of the turn; a structural change event (row/column insert or delete, sheet add/delete) marks it dirty and the next turn rebuilds it.

The selection is attached because it is the one piece of workbook state the user has already pointed at. "What's wrong with this?" must not cost a round trip. A selection over 500 cells is reduced to its address plus a note: the model reads the part it needs.

### 2.2 What is never attached

- Cell values the model did not read (no snapshot, no orientation dump, no hidden pre-read).
- Cell styles (fills, fonts, borders, colours). Formatting *codes* are attached because they change what a value means; styles are not.
- Charts, images, comments, notes, data validation, conditional formatting, row heights, column widths.
- Protection state and the undo stack; #9 owns both at write time.
- Anything from another workbook. The pane is document-scoped ([ADR 0001](../../adr/0001-runtime-topology.md)); there is no "active workbook" to follow.

### 2.3 Why not a snapshot

The VBA baseline sent a snapshot: 80×26 cells per sheet, 160 characters per cell, ~45K characters total ([README](../../../README.md), [gateway research](../research/03-gateway-streaming-and-tools.md)). It was a lie for any real workbook — it silently cut rows and columns — and it re-sent the same data every turn.

A snapshot also fights staleness and citations: the model would quote values it never re-read, and the pane could not say which revision a number came from. Paged reads cost a tool round trip, but each block is bounded, narrated, attributable and refreshable. That is the trade this ticket makes.

### 2.4 First turn, empty workbook, no selection

- **First turn:** inventory only. The model reads before it answers a workbook question (the system prompt says so; #12 owns the wording). "Hi" costs no cell data.
- **Empty workbook:** inventory says every used range is empty; a read returns an empty block, not an error.
- **No selection:** the selection field is `null`; nothing else changes.
- **Read-only workbook:** the inventory flags `readOnly: true`; reads are unaffected, writes are #9's refusal.

### 2.5 Scope and binding

- The pane's `Excel.run` context **is** the scope. There is no API that follows the active workbook, and no cross-workbook read ([capability envelope](../research/01-officejs-capability-envelope.md)).
- Multiple workbooks mean multiple panes with separate context and history. They share nothing.
- An unsaved workbook has `Office.context.document.url === null`; `Workbook.name` is a display label, not identity. Identity, persistence and rename handling are [#11](https://github.com/mustaphaelou/excel-agent/issues/11).
- A workbook in Protected View or with an IRM policy may refuse reads; the pane surfaces the Excel error and does not retry.

---

## 3. Context schema

### 3.1 `WorkbookContext`

Sent once per turn, before the first model call, as a system/context block.

```json
{
  "v": 1,
  "rev": 12,
  "at": "2026-11-08T10:55:00Z",
  "workbook": {
    "name": "Q3 Forecast.xlsx",
    "url": null,
    "readOnly": false,
    "isDirty": true,
    "displayLanguage": "en-US",
    "contentLanguage": "en-US",
    "valuesStale": false
  },
  "selection": {
    "sheet": "Forecast",
    "address": "Forecast!B4:F12",
    "areas": ["B4:F12"],
    "cells": null
  },
  "sheets": [
    {
      "id": "{00000000-0000-0000-0000-000000000000}",
      "name": "Forecast",
      "visibility": "Visible",
      "usedRange": "Forecast!A1:H240",
      "rows": 240,
      "cols": 8,
      "tables": [
        {
          "name": "ForecastTbl",
          "range": "Forecast!A1:H240",
          "headers": ["Month", "Revenue", "COGS", "Gross", "Opex", "EBITDA", "Margin", "Note"]
        }
      ],
      "pivots": [
        { "name": "ByRegion", "range": "Forecast!J2:P20" }
      ]
    }
  ],
  "names": [
    { "name": "TaxRate", "formula": "=Forecast!$B$2", "scope": "Workbook", "visible": true }
  ],
  "truncated": { "sheets": 0, "tables": 0, "pivots": 0, "names": 0 }
}
```

Rules:

- `rev` is the pane's revision counter (§9). Every read repeats it.
- `usedRange` is `Worksheet.getUsedRange(true)` — values only, so formatting-only cells do not inflate it ([Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21), 1.2). An empty sheet reports `null`.
- `visibility` is `Visible`, `Hidden` or `VeryHidden` ([Worksheet.visibility](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21)).
- `tables[].headers` is the header row's values; omitted when the header row is empty.
- `pivots[].range` is `PivotTable.layout.getRange()` ([PivotLayout](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivotlayout?view=excel-js-1.21), 1.8). Pivot fields are not inventoried; read the output range or ask for a field list.
- `names` is `Workbook.names` ([Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21), 1.1): `name`, `formula`, `scope`, `visible`. Names that resolve to `#REF!` are included as written.
- `selection.areas` is the per-area A1 list from `Workbook.getSelectedRanges()` ([Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21), 1.9); `selection.cells` is a `CellBlock` only when the selection is ≤500 cells.
- `truncated` counts what the caps removed (§5), never silent.
- Hidden and very-hidden sheets are listed with their visibility. Their contents are not read unless the model asks.

### 3.2 `CellBlock`

One read result. Values are canonical (§8); dates are ISO strings.

```json
{
  "v": 1,
  "sheet": "Forecast",
  "address": "Forecast!A1:H50",
  "rows": 50,
  "cols": 8,
  "rev": 12,
  "formats": {
    "columns": { "A": "General", "B:G": "#,##0", "H": "yyyy-mm-dd" },
    "cells": { "D7": "0.0%" }
  },
  "cells": [
    ["Month", "Revenue", "COGS", "Gross", "Opex", "EBITDA", "Margin", "Note"],
    ["Jan", 120000, 48000, 72000, 31000, 41000, 0.5694, "2026-01-16"]
  ],
  "formulas": { "D2": "=B2-C2", "F2": "=D2-E2" },
  "serials": { "H2": 46038 },
  "merged": ["A4:A5"],
  "spills": [],
  "hasMore": { "next": "Forecast!A51:H100" },
  "truncatedCells": 0,
  "notes": []
}
```

Abridged: `cells` shows two of the fifty rows.

- `cells` is row-major, exactly `rows × cols`. `null` is an empty cell; booleans are booleans; errors are strings beginning with `#` (`"#DIV/0!"`); dates are `YYYY-MM-DD` (or full ISO when the format carries a time).
- `formats.columns` is a run-length map of `numberFormat` codes per column range; `formats.cells` holds exceptions. `numberFormat` is the language-neutral code ([Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21), 1.1); `numberFormatLocal` is never sent.
- `formulas` is a sparse map of A1 (sheet-local) to formula text, present only when the read asked for formulas. Formulas are en-US and include `=` ([Range.formulas](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21)).
- `serials` is a sparse map of A1 to the raw Excel serial for date/time cells, so the ISO string is never the only copy (§8.4).
- `merged` is the block's merged areas, sheet-local, from `Range.getMergedAreasOrNullObject()` (1.13). Only the anchor carries a value; the other cells are `null`.
- `spills` lists spill extents when the block is a single cell that spills (`Range.hasSpill`, `getSpillingToRangeOrNullObject`, 1.12).
- `hasMore.next` is a ready-to-read address for the next page, not an opaque cursor.
- `truncatedCells` counts cells whose text hit the 1,000-character cap; `notes` carries read-level remarks ("3 cells are formulas", "date system assumed 1900").
- The block is sent as compact text (TSV grid plus a header and sidecars), not as JSON objects per cell: a measured 2,000-cell numeric block is ~18.7 KB ≈ 4.7K tokens as TSV, versus several times that as JSON. #12 owns the final wire encoding; the fields above are the contract.

### 3.3 Read request

```json
{ "sheet": "Forecast", "address": "A1:H50", "formulas": true, "text": false, "maxCells": 2000 }
```

The result is a `CellBlock` (§3.2) plus the tool-level `notes`.

- `sheet` is a name or id; names are matched case-insensitively and canonicalized.
- `address` is sheet-local A1, a named range, or a whole column/row (`A:A`, `1:1`).
- `formulas` adds the sparse formula map. `text` adds display text for cells whose format is not covered by the legend.
- `maxCells` is clamped to 2,000; a larger range is paged, not rejected.

---

## 4. Read protocol

### 4.1 Tools

#12 owns the schemas; #10 fixes the semantics.

| Tool | Purpose | Notes |
| --- | --- | --- |
| `read_range` | Read a sheet-local range | Default 2,000 cells; paged; values always, formulas/text on request |
| `get_precedents` | Cells a formula depends on | Direct by default; `transitive: true` opts into the closure |
| `get_dependents` | Cells that depend on a range | Same shape |

`read_range` is the only way cell data enters context. There is no `get_workbook`, no `read_sheet` that dumps a used range, and no implicit read beyond the ≤500-cell selection block (§2.1).

### 4.2 Modes and fields

- **Values** are always returned, in canonical form.
- **Formulas** are opt-in (`formulas: true`). The grid keeps computed values; the formula map carries the text. This avoids the "value or formula?" ambiguity and keeps formula-heavy sheets from doubling the grid.
- **Formats** are always returned as the run-length legend. It is what makes a raw number readable (`#,##0`, `0%`, `yyyy-mm-dd`, `$#,##0.00`).
- **Display text** is opt-in (`text: true`) and only for cells whose format is not covered by the column legend. The legend already says "this column is currency"; a one-off cell does not.
- **Merged** areas and **spill** extents are always included when present; both are cheap (a property load and one call on the read range).

### 4.3 Paging and truncation

- The pane counts **grid cells** (`rows × cols`), not non-empty cells. A sparse 2,000-cell block still costs 2,000 cells of budget; that is the honest unit because it is what the model must parse.
- A request larger than 2,000 cells returns the **first page** — whole rows, filled column by column — with `hasMore.next`. It does not error. The tool description tells the model to continue with `next` or narrow the columns.
- A single cell's text over 1,000 characters is cut at 1,000 and counted in `truncatedCells` with a `…[truncated]` marker. Nothing is cut silently.
- The **turn budget** is enforced before a read executes (§5). A read that would exceed it returns a `BudgetExhausted` tool error with the remaining budget; the model must summarize or ask for a narrower range. No partial block is returned for budget reasons.
- A read never spans sheets. Cross-sheet analysis is multiple reads.

### 4.4 Normalization

| Input | Behaviour |
| --- | --- |
| `A:A`, `1:1` | Intersect with the sheet's used range (`A1:A240`); an empty intersection returns an empty block |
| `B4:A1` (reversed) | Normalize to the bounding rectangle |
| Named range | Resolve through `Workbook.names` to its address, then read |
| `$A$1` absolute | Accepted; `$` stripped in labels |
| R1C1 | Rejected: `InvalidAddress` |
| `[Book1.xlsx]Sheet1!A1` | Rejected: single workbook only |
| `Table1[Column]` | Rejected: cite/read the table's range from the inventory |
| Empty sheet | `getUsedRangeOrNullObject` is null → empty block, not an error |
| Hidden sheet | Readable when named explicitly; not included in any default read |
| Merged block | Anchor value plus `merged` list; other cells `null` |
| Spill anchor | 1×1 read includes `spills`; a block covering the spill reads its cells normally |

`Worksheet.getUsedRange(true)` on a blank sheet returns the top-left cell rather than throwing ([Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21)); the pane uses `getUsedRangeOrNullObject` (1.4) to distinguish "empty" from "one cell".

### 4.5 Errors

Tool errors are structured and actionable, never prose:

| Code | When | Model should |
| --- | --- | --- |
| `InvalidSheet` | Name/id not found | Re-inventory, then retry or ask |
| `InvalidAddress` | Unparseable, R1C1, external, structured ref | Use the inventory's A1 address |
| `BudgetExhausted` | Turn token budget spent | Summarize, narrow, or end the turn |
| `WorkbookChanged` | The target sheet/range vanished mid-turn | Re-inventory, then retry |
| `ReadFailed` | Excel refused (protected view, IRM, cell edit) | Explain; do not retry in a loop |

`getPrecedents`/`getDependents` returning `ItemNotFound` is **empty**, not an error ([Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21)).

---

## 5. Budget

### 5.1 Caps

| Cap | Value | Why |
| --- | --- | --- |
| Cells per read | **2,000** | Claude's chunk size; one narrated step; ~5–14K tokens |
| Characters per cell | **1,000** | Bounds a pathological text cell; keeps a block under the 5 MB payload limit |
| Tokens per turn (workbook data) | **60K** | ~23% of the smallest gateway context (262K); leaves room for history and output |
| Cells per turn | **12,000** | Backstop for cheap-but-numerous reads |
| Selection inline | **500 cells** | Cheap enough to attach every turn |
| Inventory | 200 sheets, 100 tables, 100 pivots, 200 names | ~10K tokens worst case |
| Areas per citation | **20** | Keeps a chip and a `RangeAreas` sane |
| Citation target length | **512 chars** | Bounds the parser and the chip |

### 5.2 Token math

- A measured 2,000-cell dense numeric block (values, formulas, dates, format legend) is **18.7 KB ≈ 4.7K tokens** as TSV.
- Text-heavy blocks run **2–3×** that: 9–14K tokens for 2,000 cells.
- 60K tokens is therefore **6 dense reads** (the 12,000-cell backstop binds first) or **4–5 text-heavy reads** per turn.
- Inventory is 1–3K tokens for a normal workbook; the selection inline block is ≤3K.
- The smallest gateway context is 262,144 tokens ([gateway research](../research/03-gateway-streaming-and-tools.md)); 60K leaves 200K for the system prompt, history and the reply.

### 5.3 Enforcement

- The pane estimates tokens as `characters / 4` over everything it sends as workbook data (inventory + selection + blocks + notes). The estimate is advisory in accuracy but hard in effect: it is what the cap counts.
- Before a read, the pane checks the remaining turn budget. Over budget → `BudgetExhausted`, no read.
- At 80% of the budget the next tool result carries `budget: { used, remaining }` so the model can start wrapping up.
- The budget resets at each user turn. It is not a conversation budget; compaction and context-window management are [#11](https://github.com/mustaphaelou/excel-agent/issues/11).
- No silent eviction: the pane never drops an old read to make room. The model chooses what to re-read.

### 5.4 What the model can request next

- The next page (`hasMore.next`).
- A narrower range: fewer columns, a smaller row band.
- A different sheet.
- A dependency trace instead of a full read.
- A summary it writes itself; the pane does not summarize workbook data for it.

---

## 6. References

### 6.1 Formulas

- Formulas are the en-US text Excel returns from `Range.formulas`, `=`-prefixed, with sheet-qualified references (`=SUM(Forecast!B2:B11)`), structured table references (`=SUM(ForecastTbl[Revenue])`) and external references (`=[Book1.xlsx]Sheet1!A1`) as written.
- The pane does **not** parse formulas into an AST. It does not evaluate them, rewrite them, or resolve dependencies itself.
- `formulasLocal` is never sent. It exists for display and for #9's dangerous-function scan on non-English workbooks.
- A formula's computed value comes from the grid; a formula cell with no computed value (manual calc) shows its last cached value and the workbook context says `valuesStale`.

### 6.2 Named ranges

- `Workbook.names` is inventoried: `name`, `formula` (which for a range is `=Sheet!$A$1:$B$10`), `scope` (`Workbook`/`Worksheet`), `visible`.
- Formulas keep the names the user wrote; the model resolves them through the inventory. The pane does not substitute addresses into formula text.
- Citations do **not** accept names. The model cites the address the name resolves to, so a renamed or re-pointed name cannot silently change what a chip means.
- Sheet-scoped names are included with their scope; the inventory does not flatten them into workbook names.

### 6.3 Precedents and dependents

- Four APIs exist: `getDirectPrecedents` (1.12), `getDirectDependents` (1.13), `getPrecedents` (1.14), `getDependents` (1.15). All return `WorkbookRangeAreas` and throw `ItemNotFound` when empty ([Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21)).
- They are **tools**, not annotations. No read auto-attaches its dependency chain: one `SUM` over a 4,000-row column would add a 4,000-cell closure the model did not ask for.
- Default is **direct**; `transitive: true` opts into the closure. Results are capped at **500 cells / 20 areas**; over the cap returns the first 20 areas plus `truncated: { areas: N }`.
- Results are flattened from `WorkbookRangeAreas` per sheet: `areas` / `getRangeAreasBySheet` ([WorkbookRangeAreas](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbookrangeareas?view=excel-js-1.21), 1.12) → sheet-qualified A1 addresses. Empty is `{ areas: [] }`, not an error.

### 6.4 Tables and pivots

- Tables are inventoried with their range and header labels. The labels matter: they are what lets the model read `ForecastTbl[Revenue]` without a second call.
- Table data is read by A1 range; there is no separate table-read tool in v1. Table totals/header ranges are A1 too.
- Pivots are inventoried with name and output range. Their output cells are ordinary readable cells. Pivot **fields** are not inventoried; a field list is a future tool, and pivot editing is #9/#12.

### 6.5 Merged cells, hidden sheets, whole columns, spill

- **Merged:** the block lists its merged areas; only the anchor has a value, the rest are `null`. This is Excel's model, not a pane choice.
- **Hidden / very hidden:** inventoried with visibility, readable when named, never auto-read, never navigable (§7.7).
- **Whole column/row:** normalized to the used intersection (§4.4). A whole-column citation is legal; the click selects the used intersection.
- **Spill:** a 1×1 read of a spill anchor reports the spill extent; a larger block reads the spilled cells as ordinary values. Spill *writes* are #9's.

### 6.6 Not context

Charts, images, shapes, comments/notes, hyperlinks as properties, data validation, conditional formatting, cell styles, row/column hidden state, protection, page setup, the undo stack. Each is either irrelevant to a read or owned by another ticket.

---

## 7. Citation protocol

### 7.1 Grammar

A citation is a markdown link with the `cite:` scheme. The destination is wrapped in `<…>` so sheet names with spaces and parentheses survive.

```
citation   = "[", label, "](", "<cite:", target, ">)"
target     = area, { ",", area }                  ; all areas on one sheet
area       = sheet-ref, "!", range
sheet-ref  = quoted-name | bare-name
quoted-name= "'", { any-char-except-apostrophe | "''" }, "'"
bare-name  = any sheet name without "!" "," "<" ">" ")" and not shaped like a cell reference
range      = cell-ref, [ ":", cell-ref ]
           | column-ref, ":", column-ref
           | row-ref, ":", row-ref
cell-ref   = [ "$" ], letters, [ "$" ], digits
```

- **Sheet-qualified always.** The token carries the sheet even when the chip label omits it.
- **One sheet per token.** `RangeAreas` is single-sheet, and `RangeAreas.select()` can only select one sheet's areas ([RangeAreas](https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeareas?view=excel-js-1.21)). Cross-sheet answers use multiple chips.
- **Quote when needed:** the name contains `!`, `,`, `<`, `>`, `)` or an apostrophe, or the name looks like a cell reference. Apostrophes double inside quotes (`'O''Brien'!A1`).
- **No names, no R1C1, no workbook prefixes.** `cite:TaxRate`, `cite:R1C1`, `cite:[Book1.xlsx]Sheet1!A1` are invalid.
- **Absolute references** are accepted and normalized (`$B$4:$F$25` → `B4:F25`).
- **Limits:** 20 areas, 512 characters.

### 7.2 Examples

| Reply text | Chip |
| --- | --- |
| `[B22:F25](<cite:Forecast!B22:F25>)` | `B22:F25` (sheet elided: Forecast is the active sheet) |
| `[Summary!B12](<cite:Summary!B12>)` | `Summary!B12` |
| `[B22:F25, B30:F30](<cite:Forecast!B22:F25,Forecast!B30:F30>)` | `B22:F25 +1` (first area plus a count) |
| `[A:A](<cite:Forecast!A:A>)` | `A:A`; click selects `A1:A240` |
| `[Raw Data!B4](<cite:'Raw Data'!B4>)` | `Raw Data!B4`; click explains it is hidden |

The **label is ignored**. The chip always renders the canonical address; a model-supplied label can never make a chip say something other than where it goes. The label exists only so the raw markdown reads sensibly before enhancement.

### 7.3 Parsing

1. Parse the reply as markdown. Before sanitization, intercept link destinations that start with `cite:`.
2. Only a **complete** `](<cite:…>)` becomes a chip. A partial token mid-stream renders as text; the next delta re-renders it. No flicker, no half-chip.
3. Split the target on commas **outside quotes**; split each area on the **last** `!` outside quotes.
4. Resolve the sheet by name (case-insensitive, trimmed) against the inventory's canonical names.
5. Parse the range, normalize case and `$`, and validate bounds (`A1`–`XFD1048576`).
6. All areas must resolve to the same sheet; a mixed-sheet target is invalid.
7. Invalid tokens render as **plain text** (the label), never as chips.

The parser is the pane's, not the renderer's: a `cite:` URL must never reach the DOM as a real link. #12's system prompt must include the grammar and one example; #8's "only structured citations become chips" is enforced here.

### 7.4 Validation

- **At render:** sheet exists, range parses, bounds hold, areas share a sheet. Failure → plain text.
- **At click:** re-resolve the sheet and range against the live workbook. The render-time inventory may be stale; the click is authoritative.
- The pane does **not** verify that the model read the range it cites. Provenance is encouraged by the prompt, not enforced; a citation to an inventory address is legitimate.
- A citation is an **address**, not a tracked range. If rows were inserted above `B22:F25`, the chip still points at `B22:F25`; the pane does not re-map it and the model is told to re-read after structural changes.

### 7.5 Click

```
click(chip):
  sheet = live lookup by name            → missing: snackbar "no longer exists"
  if sheet.visibility != Visible         → snackbar "hidden; show it to jump"
  areas = areas.map(a => intersectWithUsedRange(a))   // A:A → A1:A240
  Excel.run:
    sheet.activate()
    if areas.length == 1: sheet.getRange(areas[0]).select()
    else:                 sheet.getRanges(areas.join(",")).select()   // RangeAreas.select, 1.18
    sync()
  catch InvalidOperationInCellEditMode   → snackbar "finish editing, then jump"
  catch other                            → snackbar with the reason
  focus stays in the grid; the pane does not steal it back
```

- **Activate, then select.** The sheet is activated first so the selection is visible where the user is looking; a hidden sheet is refused before either call. #6 proved both calls throw on a hidden sheet (`InvalidSelection` / `InvalidReference`).
- **Multi-area** uses `Worksheet.getRanges` (1.9) and `RangeAreas.select()` (1.18); the build supports 1.21 ([capability envelope](../research/01-officejs-capability-envelope.md)).
- **Scrolling** is Excel's: `select()` moves the selection and Excel brings it into view. There is no separate scroll API.
- **Protected sheets** are navigable; protection blocks writes, not selection.
- **Same-sheet label elision** uses the sheet that was active when the turn started, captured in the turn context, so a mid-stream sheet switch cannot rewrite an earlier chip.
- A multi-area chip with a dead area selects the live areas and says how many were skipped.

### 7.6 Broken citations

| Situation | Render | Click |
| --- | --- | --- |
| Unparseable token, unknown sheet, mixed sheets | Plain mono text (the label) | Nothing; not a button |
| Sheet deleted after render | Normal chip | Snackbar: "That sheet no longer exists." |
| Sheet renamed after render | Chip resolves through the session alias and updates its label | Navigates to the renamed sheet |
| Sheet hidden / very hidden | Normal chip | Snackbar naming the sheet; no unhide |
| Cell-edit mode | Normal chip | Snackbar: "Finish editing the cell, then jump." |
| At click, one area of several no longer valid | Normal chip | Selects the live areas; notes the skipped count |

Nothing broken ever looks clickable; nothing clickable ever fails silently.

### 7.7 Hidden sheets

- **Readable and inventoried** with `visibility`, because hidden data is still workbook data and a question may be about it.
- **Not navigable.** `Range.select()` and `Worksheet.activate()` both throw on a hidden sheet (#6), and unhiding is a non-undoable `Worksheet.visibility` write that #9 refuses (§2.2). v1 therefore refuses with a snackbar; it does not unhide and does not jump to a "nearest visible" lie.
- **Not auto-read.** A default read never includes hidden content; the model must name the sheet.
- Very hidden is treated identically; the sheet name is still shown so the user knows what was cited.

### 7.8 Security

- Only the `cite:` scheme is intercepted; every other link is an external link handled by #8 (externalized, marked).
- The label is ignored, so a chip cannot be made to point somewhere other than it says.
- The target is validated against the workbook; there is no navigation to arbitrary URLs, files, or named objects.
- Citations never write, never unhide, never change formatting.
- The parser runs on the pane's own markdown, not on HTML; no raw HTML reaches the DOM (#8).

---

## 8. Locale

### 8.1 Canonical (on the wire)

| Thing | Form | Source |
| --- | --- | --- |
| Address | en-US A1, sheet-qualified (`Forecast!B4:F25`) | `Range.address` |
| Formula | en-US, `=`-prefixed | `Range.formulas` |
| Number format | language-neutral code (`#,##0.00`, `yyyy-mm-dd`) | `Range.numberFormat` |
| Date | ISO 8601 (`2026-01-16`, `2026-01-16T09:30:00`) | pane conversion from the serial |
| Decimal | `.` | raw `values` |
| Boolean | `true` / `false` | raw `values` |
| Error | `#DIV/0!` and friends | raw `values` / `valueTypes` |
| Sheet name | the user's name, verbatim | `Worksheet.name` |

### 8.2 Display (in the pane)

- `Range.text` is the formatted display string; it is used for diffs (#9 §3.1), copy, and `text: true` read cells, never for logic.
- `addressLocal` and `numberFormatLocal` are used only if a pane surface needs the user's local notation; chips, citations and the protocol stay canonical.
- The pane UI is English in v1; the workbook's language does not change it.

### 8.3 Detection

- `Office.context.displayLanguage` (UI language) and `Office.context.contentLanguage` (editing language) are read at startup and recorded in the inventory ([localization](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/localization)).
- Detection is **informational**. It never switches the protocol, the citation grammar or the formula language. The spike machine runs English Excel (the #6 screenshot shows an English ribbon), but a French or Arabic workbook must not change what the model reads.
- Sheet names are user data in any script; citations quote them and the parser is Unicode-safe.

### 8.4 Dates and the 1904 caveat

- A date-formatted cell is sent as an ISO string in `cells` **and** its raw serial in `serials`. The serial is authoritative if the two ever disagree.
- The conversion assumes the **1900 date system**. The Excel JS API exposes no date-system property, so a 1904-system workbook (Mac legacy) would shift `d` by 1,462 days; the raw serial and the format legend still let a later version correct it. This is a known constraint (§11), not a silent assumption.
- Time-only formats (`hh:mm`) are sent as ISO times (`09:30:00`), not as a 1899-12-30 date.
- Currency and percent are not converted: the raw number plus the format code is the truth (`0.5694` + `0.0%`).

---

## 9. Staleness

### 9.1 Revision

- One pane-level counter, `rev`, starts at 0 when the pane loads.
- It increments on:
  - `WorksheetCollection.onChanged` (1.9) — any worksheet's data or comments;
  - `WorksheetCollection.onFormulaChanged` (1.13) — formula text;
  - `WorksheetCollection.onCalculated` (1.8) — a worksheet finished calculating;
  - `WorksheetCollection.onAdded` / `onDeleted` (1.7) — a sheet appeared or vanished;
  - `WorksheetCollection.onNameChanged` / `onVisibilityChanged` / `onMoved` (1.17) — a sheet was renamed, hidden/unhidden, or reordered.
- `WorksheetChangedEventArgs` (1.7) gives `worksheetId`, `address`, `changeType`, `source` and `triggerSource`; `WorksheetNameChangedEventArgs` (1.17) gives `nameBefore` / `nameAfter`; `changeType` is one of `RangeEdited`, `CellInserted`, `CellDeleted`, `RowInserted`, `RowDeleted`, `ColumnInserted`, `ColumnDeleted`, `Unknown` ([DataChangeType](https://learn.microsoft.com/en-us/javascript/api/excel/excel.datachangetype?view=excel-js-1.21)).
- Dirty areas are coalesced per sheet, capped at **10 addresses plus counts**; older entries are summarized as "+N more".
- Self-triggered events (the pane's own apply) still bump `rev` but are marked `self`, so the next note can say "your apply changed…" rather than blaming the user.

### 9.2 Turn notes

At the start of a turn, if `rev` moved since the previous turn's inventory:

```
Workbook changed since your last read: Forecast!B4:B9 (edited), Sheet2!A1:F40
(12 edits), rows inserted on Inputs. Re-read anything you rely on.
```

The note is capped, names sheets and addresses, and never claims more than the events reported. It does not include values — the model re-reads what it needs.

### 9.3 Mid-turn

- Every read repeats the current `rev`.
- If `rev` advanced since the previous read in the same turn, the tool result carries `changedSinceLastRead: true` and the pane adds a one-line note. The read itself is current as of its `sync()`.
- There is no automatic re-read and no lock. The user can edit while the model thinks; the model is told, not prevented.

### 9.4 Recalculation

- `onCalculated` sets `workbook.valuesStale` until the next read; the note says values were recalculated.
- Formulas and addresses are unchanged by a recalc, so citations stay valid.
- Manual calculation mode is not exposed by the API. In manual mode the cached values may lag; the workbook context sets `valuesStale` when the pane has seen a calculate event, and the next read's `notes` repeats it. Otherwise the pane cannot tell. This is a known constraint (§11).

### 9.5 Structural changes

- Row/column insert or delete and cell edits are reported by `changeType`; sheet add/delete/rename and visibility changes arrive through their own collection events (1.7 / 1.17).
- Citations are addresses (§7.4); the pane does not re-map cell addresses. The turn note says structure changed, and the model is expected to re-read before relying on an address.
- A **rename** is aliased for the session (`nameBefore → nameAfter`), so a chip rendered before the rename still resolves and its label updates on the next render. A rename that predates the pane is unresolvable by name; the click snackbar says the sheet is gone.
- A **visibility change** updates the chip's click state immediately; a sheet hidden mid-session stops navigating without waiting for a click to fail.

### 9.6 Pane lifecycle

- There is no document-close or workbook-deactivate event ([capability envelope](../research/01-officejs-capability-envelope.md)). When the pane closes, the runtime and the counter die.
- On reopen the pane rebuilds the inventory, resets `rev` to 0 and starts a fresh baseline. Dirty state is **not** persisted; persistence is [#11](https://github.com/mustaphaelou/excel-agent/issues/11).
- Workbook switching does not exist for a pane (§2.5). Another workbook is another pane with another counter.

---

## 10. Copy deck

| Context | Copy |
| --- | --- |
| Read cap | "That's 3,400 cells; the limit is 2,000 per read. Reading the first 2,000." |
| Budget | "This turn has used 58K of 60K tokens of workbook data. Narrow the range or start a new turn." |
| Whole column | "Reading the used range of column A (A1:A240)." |
| Empty sheet | "Sheet 'Notes' has no used range." |
| Stale | "Workbook changed since your last read: Forecast!B4:B9 and 12 other areas." |
| Recalc | "Values were recalculated; formulas are unchanged." |
| Hidden | "Sheet 'Raw Data' is hidden. Show it in Excel to jump." |
| Deleted | "That sheet no longer exists." |
| Cell edit | "Finish editing the cell, then jump." |
| Skipped areas | "Jumped to 2 of 3 areas; 1 no longer exists." |
| Unresolved | (no copy: renders as plain text) |
| Read-only | "Read-only workbook. Questions only." |
| Dependency cap | "That trace reaches 4,200 cells; showing the first 500." |

---

## 11. Known constraints

- **2,000 cells is a product cap, not an API one.** Office.js allows far more; the 5 MB `context.sync()` payload limit is the hard ceiling ([performance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/performance)), and 2,000 × 1,000 characters is designed to stay under it.
- **Selection is a UI side effect.** There is no scroll-without-select and no highlight API; a citation moves the user's cursor. This is the only navigation primitive.
- **No cross-sheet multi-area selection.** `RangeAreas` is single-sheet; a cross-sheet answer is multiple chips.
- **Hidden sheets cannot be navigated.** `select()`/`activate()` throw (#6); unhide is a non-undoable write #9 refuses. v1 refuses.
- **Citations are addresses, not tracked ranges.** Structural edits can shift what an address means; the pane tells the model, it does not re-map.
- **No calculation-mode API.** The pane cannot tell manual from automatic; `valuesStale` is best-effort.
- **No date-system API.** 1904 workbooks shift ISO dates; the raw serial is kept.
- **`text` is display text.** It can be width-dependent in Excel's UI and is never used for logic.
- **No provenance enforcement.** A citation to a valid address the model never read is accepted.
- **Renames are aliased only within the session.** A rename that predates the pane, or a persisted chip from #11, breaks by name.
- **Events coalesce.** Dirty areas are capped at 10; the note is a summary, not a journal.
- **No document-close event.** The revision dies with the pane; persistence is #11.
- **Inventory is rebuilt per turn.** A 200-sheet workbook costs a batch of calls at each turn start; that is the price of freshness.
- **Pivot fields are not inventoried.** The output range is; a field list is future work.
- **Very hidden sheets are listed by name.** Their existence is not secret from the model; their contents require an explicit read.
- **The model is remote.** Workbook data sent as context goes to the selected gateway model; "local-first" means no server-side store, not no egress. #11 owns the privacy copy.

---

## 12. Owned elsewhere

- **#11** — workbook identity, history, persistence, instructions, compaction, privacy copy.
- **#12** — tool schemas, SSE event shapes, system prompt (including the citation grammar), status line, error codes.
- **#13** — the prototype's citation chip states, snackbars and copy.
- **#9** — writes, approval, hidden-sheet writes, drift, dangerous formulas.
- **#8** — chip geometry, contrast, focus, keyboard activation, transcript layout.
- **#7 / ADR 0001** — runtime and pane lifecycle.

---

## 13. Sources

- [`Excel.Range`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) — `values`, `formulas`, `numberFormat`, `text`, `address`, `valueTypes`, `hasSpill` (1.12), `getSpillingToRangeOrNullObject` (1.12), `getMergedAreasOrNullObject` (1.13), `getDirectPrecedents` (1.12), `getDirectDependents` (1.13), `getPrecedents` (1.14), `getDependents` (1.15).
- [`Excel.Worksheet`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21) — `getUsedRange(valuesOnly)` (1.2), `getUsedRangeOrNullObject` (1.4), `getRanges` (1.9), `visibility`, `activate`, `onChanged` (1.7).
- [`Excel.WorksheetCollection`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-1.21) — `onChanged` (1.9), `onFormulaChanged` (1.13), `onCalculated` (1.8), `onAdded`/`onDeleted` (1.7), `onNameChanged`/`onVisibilityChanged`/`onMoved` (1.17).
- [`Excel.WorksheetNameChangedEventArgs`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetnamechangedeventargs?view=excel-js-1.21) — `nameBefore` / `nameAfter` (1.17).
- [`Excel.WorksheetChangedEventArgs`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetchangedeventargs?view=excel-js-1.21) / [`DataChangeType`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.datachangetype?view=excel-js-1.21) — the change event's fields and kinds.
- [`Excel.Workbook`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) — `names` (1.1), `pivotTables` (1.3), `getSelectedRanges` (1.9), `readOnly` (1.8), `isDirty` (1.9).
- [`Excel.RangeAreas`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeareas?view=excel-js-1.21) / [`Excel.WorkbookRangeAreas`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbookrangeareas?view=excel-js-1.21) — single-sheet areas, `select()` (1.18), per-sheet flattening (1.12).
- [`Excel.PivotLayout`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivotlayout?view=excel-js-1.21) — `getRange` (1.8).
- [Excel JavaScript API performance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/performance) — 5 MB `context.sync()` payload limit, batching, block reads.
- [Read or write to a large range](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-large) — split into blocks, `getSpecialCells`/`RangeAreas` for sparse data.
- [Localize your Office Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/localization) — `displayLanguage` / `contentLanguage`.
- [Capability envelope fact sheet](../research/01-officejs-capability-envelope.md) — requirement sets, hidden-sheet unknowns, the 5 MB/payload context.
- [Pane sideload spike](../research/06-pane-sideload-spike.md) — hidden-sheet `select()`/`activate()` throw, selection behaviour, build 20527.
- [Claude for Excel teardown](../research/04-claude-sidebar-teardown.md) — 2,000-cell chunked reads, inline chips, no sources footer, status line.
- [Sidebar UX and visual language](08-sidebar-ux.md) — C17 citation chip, same-sheet label elision, snackbar path.
- [Edit safety and approval model](09-edit-safety.md) — read/write split, 2,000-cell write cap, hidden-sheet write policy, diff display text.
