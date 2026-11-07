# Office.js capability envelope — Excel 16.0 build 20527 (Windows)

Ticket: [01 · Office.js capability envelope](https://github.com/mustaphaelou/excel-agent/issues/2)
Target: Excel 16.0 (Microsoft 365), build 20527, Windows desktop.
Method: Microsoft Learn only — Office Add-ins docs, the Office.js/Excel API reference, requirement-set and manifest reference, plus the Open XML `[MS-OWEXML]` spec. No blogs or secondary summaries. Version data read 2026-11-07.

**Confidence tags used throughout**

- **[C] confirmed** — stated in a Microsoft primary source listed in Sources.
- **[I] inferred** — follows from confirmed facts but is not stated directly.
- **[U] unknown** — not documented; must be settled by the sideload spike (ticket 05) or a live test.

> No code was run against Excel for this ticket. Everything below is documentation-derived; the build-number mapping in particular should be re-checked at runtime with `Office.context.requirements.isSetSupported` during the spike.

---

## Summary

The envelope is wide enough for the v1 scope, but it is shaped by document scope and by a non-transactional undo model.

- **The pane is document-scoped, not application-scoped.** The Excel object model entry point is `context.workbook` ([Workbook class](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21)), settings are "per file and add-in" ([SettingCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.settingcollection?view=excel-js-1.21)), and no API returns another workbook's object model (the desktop-only `Application.windows` surface is window metadata only). A single global sidebar that follows whichever workbook is foreground is **not** expressible; each workbook gets its own pane instance. **[C]**
- **Everything the v1 scope needs to read and write exists**: used ranges, values, formulas, number/text formats, tables, pivot tables (including creation), conditional formatting, data validation, sort/AutoFilter, sheet creation, range selection, and formula precedent/dependent tracing. All are inside requirement sets ≤ 1.21, which build 20527 satisfies. **[C]** (per-member evidence in [Workbook API](#workbook-api))
- **Undo exists, and 1.20 added explicit grouping.** Excel puts add-in API writes on the same undo stack as user edits; `Excel.run({ mergeUndoGroup: true })` collapses all `context.sync()` calls in one run into one Ctrl+Z step. **[C]** But a documented list of APIs is *not* undoable, and calling one **clears the user's undo stack from that call onward**. Structural work (delete/rename/reorder/hide sheets, protection) is in that list on Windows. There is no programmatic undo. **[C]**
- **Transports:** the task pane runs in the Chromium-based Edge WebView2 on Windows, which supports WebSockets, full CORS, cookies and local storage. **[C]** Server-sent events / `fetch` streaming are **not documented** by Microsoft either way — they are ordinary Chromium capabilities, so they are expected to work but must be proven in the spike. **[U]**
- **Task-pane pinning is Outlook-only.** `SupportsPinning` is documented as "Supported in Outlook only"; Excel has no pinning concept. Excel's equivalents are auto-open on document load and the shared-runtime startup behavior. **[C]**
- **Requirement-set ceiling on this build:** ExcelApi **1.21** (min Version 2606 / Build 20131.20090), ExcelApiDesktop **1.1** (min Build 19231.20156), TaskPaneApi **1.1** (min Build 19029.20004), SharedRuntime **1.1 and 1.2** (1.2 min Build 14326.20508). Build 20527 is later than all of these minimums. **[C]** for the table, **[I]** for "this build is later".
- **Strongest blockers** (detail in [Blockers and unknowns](#blockers-and-unknowns)): no cross-workbook sidebar; undo cannot be treated as rollback (structural ops are unrecoverable and poison the stack); hidden-sheet navigation for citations is undocumented; per-workbook history identity is weak for unsaved / renamed files; network-share sideload requires the XML manifest, not the unified manifest.

---

## Task pane

### Placement and sizing

| Fact | Value | Tag |
| --- | --- | --- |
| Default placement | Docked to the right side of the application window ("typically appear on the right side" for Word, PowerPoint, Excel, Outlook) | **[C]** [Task panes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/task-pane-add-ins) |
| Design size in Excel | 350 × 378 px at 1366 × 768 with the ribbon shown | **[C]** [same](https://learn.microsoft.com/en-us/office/dev/add-ins/design/task-pane-add-ins) |
| User resizing | The pane can be resized in the host UI; the file stores the resulting `width` as the pane's "default width". Exact drag behaviour and write-back timing are not documented. | **[I]** from [MS-OWEXML CT_OsfTaskpane](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/555d2968-2823-4cb3-b317-4511dfc97bef); **[U]** detail |
| Programmatic width | `Office.extensionLifeCycle.taskpane.setWidth(number)` — TaskPaneApi 1.1 | **[C]** [Office.TaskPane](https://learn.microsoft.com/en-us/javascript/api/office/office.taskpane?view=excel-js-1.21), [Office.ExtensionLifeCycle](https://learn.microsoft.com/en-us/javascript/api/office/office.extensionlifecycle?view=excel-js-1.21) |
| Width range on Windows | 86 px minimum, 50 % of the client window maximum; **default 320 px**; out-of-range values are ignored silently (no error) | **[C]** [Office.TaskPane remarks](https://learn.microsoft.com/en-us/javascript/api/office/office.taskpane?view=excel-js-1.21) |
| Width range on the web (Excel) | 350–500 px inclusive; default 350 px (irrelevant for this Windows-only target, but shows the API is platform-variable) | **[C]** [same](https://learn.microsoft.com/en-us/javascript/api/office/office.taskpane?view=excel-js-1.21) |
| `setWidth` availability | TaskPaneApi 1.1 on Windows needs **Version 2507 (Build 19029.20004)**; Excel and Word are supported on the web, PowerPoint on the web is not | **[C]** [Task Pane API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/task-pane-api-requirement-sets) |

### Per-workbook persistence and reopen behaviour

- **The document file format stores pane state per task pane.** `taskpanes.xml` (the "Taskpane Web Extension File") holds a `CT_OsfTaskpane` per instance with `dockstate` ("last-docked location"), `visibility` ("whether the Task Pane shows as visible by default when the document opens"), `width` ("default width value for this taskpane instance"), `row`, and `locked` ("whether the taskpane is locked to the document in the UI and cannot be closed by the user"). **[C]** [CT_OsfTaskpane](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/555d2968-2823-4cb3-b317-4511dfc97bef) · [Taskpane Web Extension File example](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/3d04f8ce-65f2-4dc3-bafa-636d0a7e41a1)
- **So width, dock state and "visible on open" are workbook-file properties, not app-level preferences.** Because they live in the `.xlsx`, they travel with the file. **[C]** for the schema, **[I]** for "Excel rewrites them when the user drags the splitter" — that write-back is not documented. **[U]**
- **Add-in data persists per file + add-in pairing.** `Excel.SettingCollection` is "a collection of key-value pair setting objects that are part of the workbook. The scope is limited to per file and add-in (task-pane or content) combination." **[C]** [SettingCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.settingcollection?view=excel-js-1.21)
- **Reopen, default runtime:** the runtime "shuts down when the task pane is closed", so reopening re-runs the page from scratch; nothing in JS survives. **[C]** [Runtimes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes)
- **Reopen, shared runtime:** the runtime "shuts down when the document is closed" and survives the pane being closed. `hide()`/`showAsTaskpane()` "only change the visibility of the task pane. They do not unload or reload it (or reinitialize its state)", and registered event listeners keep running while hidden. **[C]** [Runtimes](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes) · [Show or hide the task pane](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/show-hide-add-in)
- **Auto-open on document open (document-scoped):** tag the document with `Office.AutoShowTaskpaneWithDocument = true` via `Office.context.document.settings.set(...)` + `saveAsync()`, and mark one `ShowTaskpane` action with `<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>` (XML) or `"view": "Office.AutoShowTaskpaneWithDocument"` (unified). Requires the add-in to be installed; supported in Excel on Windows from Version 1705 (Build 8121.1000). Only one pane can be auto-opened. **[C]** [Automatically open a task pane with a document](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document)
- **Auto-load code on document open (app-level):** `Office.addin.setStartupBehavior(Office.StartupBehavior.load)` — "sets the startup behavior for the add-in for when the document is opened next time"; SharedRuntime 1.1. **[C]** [Office.Addin](https://learn.microsoft.com/en-us/javascript/api/office/office.addin?view=excel-js-1.21)
- **Pane visibility events:** `Office.addin.onVisibilityModeChanged(handler)` fires with `args.visibilityMode` (`Taskpane`, etc.), and returns a deregister function. SharedRuntime 1.1. **[C]** [Office.Addin](https://learn.microsoft.com/en-us/javascript/api/office/office.addin?view=excel-js-1.21) · [Show or hide](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/show-hide-add-in)
- **Close guard:** `Office.addin.beforeDocumentCloseNotification` shows a modal notification when the user closes the document and can cancel the close; Excel only; SharedRuntime **1.2**. **[C]** [Office.Addin](https://learn.microsoft.com/en-us/javascript/api/office/office.addin?view=excel-js-1.21) · [Shared runtime requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)

### `SupportsPinning`

- The manifest element is documented as: "Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection. **Supported in Outlook only.**" It is an optional child of a `ShowTaskpane` action and requires `VersionOverridesV1_1`. **[C]** [Action element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/action) (rows for `SupportsPinning`)
- The whole feature is about the Outlook Reading Pane closing when the selected message changes. **[C]** [Implement a pinnable task pane in Outlook](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/pinnable-taskpane)
- **Consequence for Excel:** there is no pin/unpin affordance to opt into. An Excel pane stays where it is until the user closes it or the workbook closes; the closest equivalents are auto-open on load, `setStartupBehavior(load)`, `showAsTaskpane()`, and the `locked` attribute in the taskpane part (file-level, not manifest-level). **[C]** for each piece; **[I]** for the equivalence.

### Other pane constraints

- A shared runtime supports **only one task pane** (one page, possibly with several views). Dialog runtimes cannot be shared. **[C]** [Configure a shared runtime](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime) · [Runtimes](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes)
- The task pane is a **browser runtime**; custom functions are JavaScript-only unless the runtime is shared. **[C]** [Runtimes](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes)

---

## Runtime/transport

### Shared runtime vs default

| Aspect | Default (no shared runtime) | Shared runtime (`lifetime: "long"`) |
| --- | --- | --- |
| Runtimes | Task pane, function commands, custom functions each get their own process | One browser runtime shared by task pane + commands (+ custom functions in Excel) **[C]** |
| Lifetime | Task pane runtime dies when the pane closes **[C]** | Lives until the document closes **[C]** |
| Global state across features | Not shared; message passing / storage needed **[C]** | Shared globals and libraries **[C]** |
| Task pane show/hide from code | Not available | `Office.addin.showAsTaskpane()` / `hide()` **[C]** |
| Auto-load on document open | No | `setStartupBehavior(load)` **[C]** |
| Custom functions | JavaScript-only runtime; **no** full CORS, no `localStorage` **[C]** | Browser runtime; full CORS and Office.js access **[C]** |
| Excel minimum | — | Excel 2002 (Build 12527.20092) for 1.1; Excel 2108 (Build 14326.20508) for 1.2 **[C]** |

Sources: [Configure a shared runtime](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime) · [Runtimes](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes) · [Shared runtime requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/shared-runtime-requirement-sets).

Manifest mechanics (XML): add `<Set Name="SharedRuntime" MinVersion="1.1"/>` under `<Requirements>`, and a `<Runtimes><Runtime resid="Taskpane.Url" lifetime="long" /></Runtimes>` inside the `<Host>` of `VersionOverrides`. The `resid` must match the task pane URL resource or the shared runtime won't load. **[C]** [Configure a shared runtime](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)

### Long-lived requests

- With a shared runtime, code and in-flight work survive the pane being closed, because the runtime is tied to the document, not the pane. **[C]** [Runtimes](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes)
- Without it, closing the pane tears down the runtime, so any open HTTP stream dies with it. **[I]** (from the lifetime rule above)
- Idle unload of a shared runtime before the document closes is not documented. **[U]**

### Transports and what the WebView allows

| Capability | Status | Tag / source |
| --- | --- | --- |
| `fetch` / `XMLHttpRequest` | Available; Microsoft's own docs describe "request external data by using an API like Fetch or by using XMLHttpRequest" | **[C]** [Activate add-ins with events](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation) |
| WebSocket | Explicitly listed as supported by the browser runtime ("support for WebSockets, Full CORS, and local storage, and cookies") and by the JavaScript-only runtime ("support for WebSockets, Full CORS") | **[C]** [Runtimes](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes) |
| CORS | "Full CORS" is a property of the browser runtime; a shared runtime gives "access to the task pane DOM and CORS from all parts of your add-in"; custom functions only get full CORS when shared | **[C]** [Runtimes](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes) · [Configure a shared runtime](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime) |
| Server-sent events / `ReadableStream` response bodies | Not mentioned anywhere in the Office Add-ins docs. Chromium supports streaming `fetch`; nothing in the Office layer is documented to block it. | **[U]** — verify in spike |
| Cookies, `localStorage` | Supported in browser runtimes; **not** in JavaScript-only runtimes | **[C]** [Runtimes](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes) |
| `sessionStorage` between runtimes | Not shared; each runtime is a separate session | **[C]** [Runtimes](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes) |

- CORS is enforced by the WebView as a normal browser. A cross-origin backend must return `Access-Control-Allow-Origin` (and handle preflight for non-simple requests). Microsoft does not document an Office-specific bypass for task panes. **[C]** for "full CORS"; **[I]** for "the standard headers are required".

### Edge WebView2 on Windows

- Windows Office hosts add-ins in **Microsoft Edge (Chromium-based) with WebView2**, and "WebView2 is installed with Office by default for supported versions of Office". Mac/iOS use WKWebView; the web uses the user's browser. **[C]** [Browsers and webview controls](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)
- Documented WebView2 quirks that matter for a chat pane: scroll bars are hidden until hovered unless CSS sets `overflow-y: scroll`; direct blob download of PDFs is unsupported (workaround: `Office.context.ui.openBrowserWindow`); WIP-protected documents require `msedgewebview2.exe` in the protected-apps list. **[C]** [same](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)
- IE/EdgeHTML is not in play on this build; the doc maps Windows to WebView2 unconditionally. **[C]** [same](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)

### HTTPS, localhost, mixed content

- **`SourceLocation` must be HTTPS:** "The source location must be an HTTPS address, not a file path." **[C]** [SourceLocation element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/sourcelocation)
- Learn's own guidance: "Office Add-ins should use HTTPS, not HTTP, even while you're developing", and on first run the dev tooling may prompt to grant Microsoft Edge WebView a localhost loopback exemption (needs admin). **[C]** [PowerPoint add-in tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/powerpoint-tutorial-yo?tabs=jsonmanifest)
- Therefore the pane page is always a secure context, and any `http://` backend call is mixed content. Chromium treats loopback origins as potentially trustworthy, but Microsoft does not document the behaviour inside the Office WebView. **[U]** — the spike should test `https://localhost:3000` → `http://127.0.0.1:8000` explicitly. Practical mitigation: serve the backend over HTTPS on localhost, or reverse-proxy it under the same origin as the pane.
- Cookies on a `Secure`/`SameSite` basis are normal Chromium behaviour; no Office-specific exceptions documented. **[I]**

---

## Workbook API

Every row below is read from the API reference page's own `[Api set: …]` remark for that member. Build 20527 is past all of these minimums per the [requirement-set table](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets).

### Reading

| Capability | API | Requirement set | Tag |
| --- | --- | --- | --- |
| Used range of a sheet | `Worksheet.getUsedRange(valuesOnly?)` | ExcelApi 1.2 (the `valuesOnly` overload; `getUsedRangeOrNullObject` is 1.4) | **[C]** [Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21) |
| Used range of a range | `Range.getUsedRange` (the `valuesOnly` parameter is 1.2), `Range.getUsedRangeOrNullObject` | 1.1 / 1.4 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Values | `Range.values` (2-D `any[][]`) | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Formulas (A1) | `Range.formulas` (returns value when no formula) | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Display text | `Range.text` | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Value types | `Range.valueTypes` | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Number formats | `Range.numberFormat` | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Locale number formats | `Range.numberFormatLocal` | 1.7 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Formatting object | `Range.format` (font, fill, borders, alignment, …) | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Merged cells | `Range.merge` / `unmerge` | 1.2 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Hidden rows | `Range.rowHidden` (tri-state: true/false/null if mixed) | 1.2 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Workbook identity | `Workbook.name` | 1.7 | **[C]** [Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) |
| Workbook metadata | `Workbook.properties` (`DocumentProperties`: author, creationDate, custom, …) | 1.7 | **[C]** [Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) · [DocumentProperties](https://learn.microsoft.com/en-us/javascript/api/excel/excel.documentproperties?view=excel-js-1.21) |
| Per-file settings | `Workbook.settings` (`SettingCollection`) | 1.4 | **[C]** [Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) |
| Stale/unsaved state | `Workbook.isDirty`, `previouslySaved`, `readOnly`, `autoSave` | 1.9 / 1.9 / 1.8 / 1.9 | **[C]** [Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) |
| Document URL (for identity) | `Office.context.document.url` — "Returns null if the URL is unavailable" | Common API | **[C]** [Office.Document](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=excel-js-1.21) |

### Writing

| Capability | API | Requirement set | Tag |
| --- | --- | --- | --- |
| Write values | `Range.values = [[…]]` | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Write formulas | `Range.formulas = [[…]]` | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Write number formats | `Range.numberFormat` / `numberFormatLocal` | 1.1 / 1.7 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Copy/paste semantics | `Range.copyFrom(source, "All" \| "Formulas" \| "Values" \| "Formats" \| "Link", skipBlanks?, transpose?)` | 1.9 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Move a range | `Range.moveTo(destination)` | 1.11 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Clear | `Range.clear(applyTo?)` | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |

### Structure and native operations

| Capability | API | Requirement set | Tag |
| --- | --- | --- | --- |
| Create a worksheet | `Workbook.worksheets.add(name?)` | 1.1 | **[C]** [WorksheetCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-1.21) |
| Create sheets from a base64 workbook | `Workbook.insertWorksheetsFromBase64` | 1.13 | **[C]** [Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) |
| Tables | `Worksheet.tables.add`, `Table.getRange`, `getDataBodyRange`, `convertToRange` | 1.1 (table sort and filters 1.2) | **[C]** [Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table?view=excel-js-1.21) |
| Table sort | `Table.sort` (`TableSort`) | 1.2 | **[C]** [Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table?view=excel-js-1.21) |
| Table filter | `Table.clearFilters()`, `reapplyFilters()`; `Table.autoFilter` | 1.2; 1.9 | **[C]** [Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table?view=excel-js-1.21) |
| Range sort | `Range.sort` | 1.2 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Sheet AutoFilter | `Worksheet.autoFilter` (`apply`, `clearCriteria`, `reapply`, `remove`, `getRange`, criteria) | 1.9 (some filter APIs 1.14) | **[C]** [AutoFilter](https://learn.microsoft.com/en-us/javascript/api/excel/excel.autofilter?view=excel-js-1.21) |
| Conditional formatting | `Range.conditionalFormats` / `ConditionalFormatCollection.add`; types: cell value, color scale, data bar, icon set, top/bottom, preset, contains text, custom | 1.6; rule-type conversion (`changeRuleTo…`) 1.17; `getRanges()` 1.9 | **[C]** [ConditionalFormat](https://learn.microsoft.com/en-us/javascript/api/excel/excel.conditionalformat?view=excel-js-1.21) · [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Data validation | `Range.dataValidation` (rule, prompt, errorAlert, ignoreBlanks, `valid`, `clear`, `getInvalidCells`) | 1.8 (`getInvalidCells` 1.9) | **[C]** [DataValidation](https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidation?view=excel-js-1.21) |
| Pivot tables — create | `Workbook.pivotTables.add(name, source: Range \| string \| Table, destination)` | 1.8 | **[C]** [PivotTableCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivottablecollection?view=excel-js-1.21) |
| Pivot tables — read/refresh | `PivotTable.refresh()`, `name`, `worksheet`, `layout`, hierarchy collections | class base 1.3; layout and hierarchies 1.8 | **[C]** [PivotTable](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivottable?view=excel-js-1.21) |
| Pivot tables — edit | `PivotLayout`, `RowColumnPivotHierarchyCollection`, `DataPivotHierarchyCollection`, `FilterPivotHierarchyCollection`; `enableDataValueEditing` 1.9; `refreshOnOpen` 1.13; `getDataSourceType`/`getSourceString` 1.15 | 1.8–1.15 | **[C]** [PivotTable](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivottable?view=excel-js-1.21) |
| Pivot tables — delete | `PivotTable.delete()` | 1.8 | **[C]** [PivotTable](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivottable?view=excel-js-1.21) |

### Navigation (citations / click-to-jump)

| Capability | API | Requirement set | Tag |
| --- | --- | --- | --- |
| Select a rectangular range | `Range.select()` | 1.1 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| Select multiple discontiguous areas | `RangeAreas.select()` | **1.18** | **[C]** [RangeAreas](https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeareas?view=excel-js-1.21) |
| Activate a sheet | `Worksheet.activate()` | 1.1 | **[C]** [Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21) |
| Current selection | `Workbook.getSelectedRange()`; `getSelectedRanges()` | 1.1; 1.9 | **[C]** [Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) |
| Active cell | `Workbook.getActiveCell()` | 1.7 | **[C]** [Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) |
| Active sheet | `Workbook.worksheets.getActiveWorksheet()` | 1.1 | **[C]** [WorksheetCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-1.21) |
| Sheet visibility | `Worksheet.visibility` (`Visible` \| `Hidden` \| `VeryHidden`) | 1.1 read, 1.2 set | **[C]** [Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21) |
| Stable sheet id | `Worksheet.id` (unchanged by rename/move), `Worksheet.tabId` | 1.1; 1.14 | **[C]** [Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21) |
| Bounding / offset ranges | `Range.getBoundingRect`, `getOffsetRange`, `getResizedRange`, `getCell` (all 1.1); `getRangeEdge` (1.13) | 1.1 / 1.13 | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |

- **`Range.select()` on a hidden or `VeryHidden` sheet is not documented.** Neither `Range.select()` nor `Worksheet.activate()` states what happens when the target sheet is hidden. Expect a failure or a silent no-op, which means a citation into a hidden sheet needs an explicit unhide/activate decision. **[U]** — test in the spike.
- Selection is a UI side effect: it moves the user's cursor and can disturb their work mid-typing. There is no "highlight without selecting" API. **[I]** from `select()` being the only navigation primitive.

### Formula precedents and dependents

| API | Semantics | Requirement set | Tag |
| --- | --- | --- | --- |
| `Range.getDirectPrecedents()` | Range containing all direct precedent cells, same or multiple sheets; throws `ItemNotFound` if none | ExcelApi **1.12** | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| `Range.getDirectDependents()` | Direct dependent cells | ExcelApi **1.13** | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| `Range.getPrecedents()` | All precedent cells (transitive) | ExcelApi **1.14** | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |
| `Range.getDependents()` | All dependent cells (transitive) | ExcelApi **1.15** | **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) |

All four return `WorkbookRangeAreas` and throw `ItemNotFound` when the result is empty — that error is expected control flow, not a failure. **[C]** [Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21)

---

## Undo

**Baseline: yes, add-in writes are on the normal undo stack.** "Excel add-ins support undo behavior. This preserves both actions performed by Excel JavaScript APIs and actions performed by the user in Excel. These actions are saved in the undo stack for an individual user." **[C]** [Undo support with the Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities)

**Grouping (ExcelApi 1.20).** `Excel.RunOptions.mergeUndoGroup`:
- `false` (default): "each `context.sync()` call creates an undo record."
- `true`: "all `context.sync()` calls in a single `Excel.run` are merged into one undo group."

**[C]** [Excel.RunOptions](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21) · ExcelApi 1.20 min Version 2509 (Build 19201.20002) per [requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets). Usage: `Excel.run({ mergeUndoGroup: true }, async (context) => { … })`.

**Failures and partial application.**
- If a grouped call includes an API without undo support, "the `UndoNotSupported` error is thrown to let you know that the operation can't be grouped." **[C]** [Undo doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities)
- Calls commit at each `context.sync()`. A run with several syncs where a later one throws leaves the earlier ones applied; `mergeUndoGroup` only groups the undo record, it does **not** roll back. **[I]** from the commit-per-sync model plus the absence of any transaction/rollback API. **[U]** to confirm the observable behaviour.
- `mergeUndoGroup: true` with multiple `Excel.run` calls still produces separate undo groups (the option is scoped to "a single `Excel.run`"). **[C]** [RunOptions](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21)

**Undo is not available programmatically.** No `Workbook.undo()`, no `Range.undo()`, no rollback API exists in the Excel namespace; `mergeUndoGroup` is the only undo-related member, and it only groups. Any "Undo this batch" button must either tell the user to press Ctrl+Z or restore from a snapshot the add-in took. **[C]** for the absence (Excel namespace listing), **[I]** for the conclusion.

**Documented non-undoable APIs relevant to v1** — calling any of these clears the undo stack from that point, so the user cannot undo past it: **[C]** [Undo doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities)

| API | Undo on Excel on Windows | Note |
| --- | --- | --- |
| `Worksheet.delete` | **No** | Deleting a sheet through the API cannot be undone at all |
| `Worksheet.copy` | **No** | |
| `Worksheet.name` | **No** (undo *is* supported on the web) | Rename via API kills the Windows undo trail |
| `Worksheet.position` | **No** (web: yes) | Reordering sheets |
| `Worksheet.visibility` | **No** (web: yes) | Hide/unhide |
| `WorksheetProtection.protect` / `unprotect` / `pauseProtection` / `setPassword` / `updateOptions` | **No** | |
| `WorkbookProtection.protect` / `unprotect` | **No** | |
| `Workbook.save`, `Workbook.close` | **No** | |
| `Workbook.insertWorksheetsFromBase64`, `WorksheetCollection.addFromBase64` | **No** | |
| `AllowEditRange.*`, `AllowEditRangeCollection.*` | **No** | |
| `PivotTableStyle.*`, `TableStyle.*`, `SlicerStyle.*`, `TimelineStyle.*` | Yes, but **no co-authoring undo**; some have no redo | Style-level pivot work |
| `Query.*` / `QueryCollection.refreshAll` | Yes, but **no redo** | |
| `DataConnectionCollection.refreshAll`, `LinkedWorkbook.refresh`, `LinkedWorkbookCollection.refreshAll` | **No** | |
| `DocumentProperties.author/category/...` | Yes on Windows (no on web) | |

**Not listed as unsupported:** cell/range value, formula and format writes, table creation and filters, AutoFilter, data validation, conditional formatting, sort, pivot creation/refresh/field layout (only the *style* APIs are called out), and `WorksheetCollection.add`. These are therefore undoable, subject to the grouping rules. **[C]** by omission from the unsupported table — treat as confirmed-but-negative evidence.

**`delayForCellEdit`.** A sibling `RunOptions` flag: when `false` (default), "if the user is in cell edit when the batch request is processed by the host, the batch automatically fails"; when `true`, the batch is delayed until the user exits edit mode. An agent that applies edits at arbitrary times needs this set, or must handle the failure. **[C]** [Excel.RunOptions](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21)

---

## Events

Source for the whole table (unless noted): [Work with Events using the Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events) plus the per-member `[Api set]` remarks on [Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21), [WorksheetCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-1.21) and [Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21).

| Event | On | Requirement set | Tag |
| --- | --- | --- | --- |
| `onSelectionChanged` | `Worksheet`, `Workbook`, `WorksheetCollection`, `Table`, `Binding` | 1.7 (Worksheet), 1.2 (Workbook), 1.9 (WorksheetCollection) | **[C]** |
| `onActivated` / `onDeactivated` | `Worksheet`, `WorksheetCollection` | 1.7 | **[C]** |
| `onActivated` | `Workbook` | 1.13 | **[C]** |
| `onChanged` | `Worksheet`, `WorksheetCollection`, `Table`, `TableCollection`, `CommentCollection` | 1.7 | **[C]** |
| `onFormulaChanged` | `Worksheet`, `WorksheetCollection` | 1.13 | **[C]** |
| `onFormatChanged` | `Worksheet`, `WorksheetCollection` | 1.9 | **[C]** |
| `onCalculated` | `Worksheet`, `WorksheetCollection` | 1.8 | **[C]** |
| `onRowSorted` / `onColumnSorted` | `Worksheet`, `WorksheetCollection` | 1.10 | **[C]** |
| `onRowHiddenChanged` | `Worksheet`, `WorksheetCollection` | 1.11 | **[C]** |
| `onNameChanged` / `onVisibilityChanged` / `onMoved` | `Worksheet`, `WorksheetCollection` | 1.17 | **[C]** |
| `onProtectionChanged` | `Worksheet`, `WorksheetCollection` | 1.14 | **[C]** |
| `onAdded` / `onDeleted` | `WorksheetCollection`, `TableCollection`, … | 1.7 | **[C]** |
| `onSingleClicked` | `Worksheet`, `WorksheetCollection` | 1.10 | **[C]** |
| `onAutoSaveSettingChanged` | `Workbook` | 1.9 | **[C]** |
| `onSettingsChanged` | `SettingCollection` | 1.4 | **[C]** |
| `onFiltered` | `Table`, `TableCollection`, `Worksheet`, `WorksheetCollection` | **preview only** | **[C]** [events doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events) |

**Event semantics**

- Handlers are destroyed "when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed"; they do not persist in the file. Registering them again after every reload is required. **[C]** [events doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events)
- Events fire for changes made by the Excel UI, by add-in JS, **and by VBA**. This matters if the old `.xlam` stays installed alongside. **[C]** [events doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events)
- `context.runtime.enableEvents = false` suppresses event firing during bulk edits (perf). **[C]** [events doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events)
- With coauthoring, `event.source` is `Local` or `Remote`. **[C]** [events doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events)
- A `Worksheet.onChanged` handler attached to a deleted sheet stops firing but leaks until reload. **[C]** [events doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events)

**Document open / close**

- **No Excel document-open event in the Excel JS API.** Open-time behaviour comes from three other mechanisms: the per-document auto-open tag, `Office.addin.setStartupBehavior(load)`, and the event-based activation `OnDocumentOpened` event. **[C]** for each mechanism; **[I]** for "no JS open event" (it is absent from the event table).
- `OnDocumentOpened` exists for Excel/PowerPoint/Word on Windows and the web, but it has **Office-application scope**: "After an add-in is installed by a Microsoft 365 admin in the Admin portal of the Microsoft 365 tenant, the add-in launches and runs code on every Office document that is opened." It runs in a JavaScript-only runtime (no UI APIs, no `Office.onReady`), times out around 300 s, and only one random add-in runs if several subscribe. Not a fit for a personally sideloaded sidebar. **[C]** [Activate add-ins with events](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation)
- **No Excel document-close event.** The only close hook is `Office.addin.beforeDocumentCloseNotification` (SharedRuntime 1.2, Excel only), a modal that can cancel the close; its `onCloseActionCancelled` handler fires if the user cancels. **[C]** [Office.Addin](https://learn.microsoft.com/en-us/javascript/api/office/office.addin?view=excel-js-1.21)
- Pane hide/show is observable via `Office.addin.onVisibilityModeChanged`. **[C]** [Office.Addin](https://learn.microsoft.com/en-us/javascript/api/office/office.addin?view=excel-js-1.21)

**Can the pane track the active workbook?**

- **Within its own workbook: partly.** `Workbook.onActivated` (1.13) fires when *this* workbook is activated, so the pane can tell when its workbook comes to the foreground. **[C]** [events doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events)
- **Across workbooks: no usable channel.** There is no `Workbook.onDeactivated` and no collection of open workbooks. The desktop-only `Application.windows`/`activeWindow` surface (see below) can indicate that the topmost window changed, but `Window` has no `workbook` property, so it cannot yield another workbook's ranges. **[C]** [events doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events) · [Excel.Application](https://learn.microsoft.com/en-us/javascript/api/excel/excel.application?view=excel-js-preview) · [Excel.Window](https://learn.microsoft.com/en-us/javascript/api/excel/excel.window?view=excel-js-1.21)
- Legacy common-API option: `Office.EventType.DocumentSelectionChanged` is documented as "Occurs when a document-level selection happens in Excel or Word". Its Excel behaviour versus `Workbook.onSelectionChanged` is not described. **[U]** [Office.EventType](https://learn.microsoft.com/en-us/javascript/api/office/office.eventtype?view=excel-js-1.21)

**Desktop-only window APIs (ExcelApiDesktop 1.1, superset of ExcelApi 1.21).** `Workbook.application` (ExcelApi 1.1) reaches `Excel.Application`, which adds `activeWindow` ("the active window (the window on top)") and `windows` ("Returns all the open Excel windows"). `Excel.Window` exposes `name`, `index`, `isVisible`, `type` (`"workbook" | "chartAsWindow" | "chartInPlace" | "clipboard"`), `activeCell`, `activeWorksheet`, `rangeSelection`, window position/size and side-by-side controls — **but no `workbook` property**, so there is no way to get another workbook's `Workbook` proxy from a window. The docs are internally ambiguous about cross-workbook scope: the `Application.windows` remark says "all the open Excel windows", while the `Window` class says "Represents a window in the workbook". **[C]** for the members; **[U]** for whether windows of *other* workbooks appear. Sources: [ExcelApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-desktop-1-1-requirement-set) · [Excel.Application](https://learn.microsoft.com/en-us/javascript/api/excel/excel.application?view=excel-js-preview) · [Excel.Window](https://learn.microsoft.com/en-us/javascript/api/excel/excel.window?view=excel-js-1.21) · [Excel.WindowCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.windowcollection?view=excel-js-1.21)

---

## Manifest and requirement sets

### Manifest types and sideloading on Windows

| Fact | Detail | Tag |
| --- | --- | --- |
| Two manifest formats | Add-in only (XML, `manifest.xml`) and unified manifest for Microsoft 365 (JSON, `manifest.json`) | **[C]** [Test Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins) |
| Windows sideload, XML manifest | Trusted network-share catalog: share a folder, then add its UNC path under **File → Options → Trust Center → Trusted Add-in Catalogs**, or write `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{GUID}` with `Url` and `Flags=1`, then restart Office | **[C]** [Sideload on Windows from a network share](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) |
| Network-share limits | "The add-in can only be installed on Windows computers"; "Add-ins that use the unified manifest for Microsoft 365 aren't supported when published to a network share"; ribbon changes require reinstall | **[C]** [same](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) |
| Unified manifest sideload | Supported on Office on Windows Version 2304 (Build 16320.20000) or later; sideloading on Windows also sideloads to the web; not on iPad | **[C]** [Sideload add-ins that use the unified manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest) |
| `SourceLocation` | "The source location must be an HTTPS address, not a file path", 1–2018 characters | **[C]** [SourceLocation](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/sourcelocation) |
| Requirements element | `<Requirements><Sets DefaultMinVersion="1.3"><Set Name="ExcelApi" MinVersion="1.3"/></Sets></Requirements>`; if the host doesn't satisfy it the add-in doesn't load and isn't listed | **[C]** [Excel requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets) |
| Desktop-only sets in XML | "The add-in only (XML) manifest can't specify desktop-only or online-only requirement sets as activation requirements. It isn't a valid value to use in the `Set` element." The unified manifest does support them | **[C]** [ExcelApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-desktop-1-1-requirement-set) |
| Runtime check | `Office.context.requirements.isSetSupported("ExcelApi", "1.3")` for feature-gating at runtime | **[C]** [Excel requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets) |

### Requirement sets available on build 20527

From the [Excel JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets) table (Windows / Microsoft 365 column):

| Requirement set | Documented minimum (Windows, M365) | Available on 20527? |
| --- | --- | --- |
| ExcelApi 1.21 | Version 2606 (Build 20131.20090) | Yes **[I]** (20527 > 20131) |
| ExcelApi 1.20 (adds undo grouping) | Version 2509 (Build 19201.20002) | Yes **[I]** |
| ExcelApi 1.19 | Version 2504 (Build 18730.20088) | Yes **[I]** |
| ExcelApi 1.18 (`RangeAreas.select`) | Version 2501 (Build 18429.20040) | Yes **[I]** |
| ExcelApi 1.17 | Version 2302 (Build 16130.20332) | Yes **[I]** |
| ExcelApi 1.12–1.16 (precedents/dependents) | Versions 2008–2208 (Build 13127.20408–15601.20148) | Yes **[I]** |
| ExcelApi 1.1–1.11 | Versions 1509–2002 | Yes **[I]** |
| ExcelApiDesktop 1.1 | Version 2509 (Build 19231.20156) | Yes **[I]** |
| TaskPaneApi 1.1 (`setWidth`) | Version 2507 (Build 19029.20004) | Yes **[I]** |
| SharedRuntime 1.1 | Excel Version 2002 (Build 12527.20092) | Yes **[I]** |
| SharedRuntime 1.2 (`beforeDocumentCloseNotification`) | Excel Version 2108 (Build 14326.20508) | Yes **[I]** |

Caveats:
- Build numbers are only comparable inside a channel; the table is the authority and the arrow comparison is mine. **[U]** until `isSetSupported` is run on the machine.
- "Preview" and `ExcelApiOnline` are not usable here (online-only; preview requires the preview CDN / Insider). **[C]** [requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets)
- `SharedRuntime 1.2` is Excel-only. **[C]** [shared runtime requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)

### Manifest shape for a shared-runtime task-pane add-in (XML)

Minimum pieces referenced by the docs above:
1. `<Requirements>` with `ExcelApi` (1.1 or the highest set you hard-depend on) and `SharedRuntime` 1.1. **[C]**
2. `<Hosts><Host xsi:type="Workbook">` with a `<Runtimes><Runtime resid="Taskpane.Url" lifetime="long"/></Runtimes>`. **[C]**
3. `<VersionOverrides>` → `<Action xsi:type="ShowTaskpane">` with `<TaskpaneId>` (`Office.AutoShowTaskpaneWithDocument` if auto-open) and `<SourceLocation resid="Taskpane.Url"/>`. **[C]**
4. `SourceLocation` resolves to an **https** URL. **[C]**

---

## Blockers and unknowns

Ordered by how much they constrain the v1 design. Each is a fact about the platform, not a design proposal.

1. **No cross-workbook object model — the pane is per workbook.** `context.workbook` is the root, settings are per file + add-in, and no API returns another workbook's `Workbook` proxy. Desktop-only `Application.windows` (ExcelApiDesktop 1.1) lists windows and `Application.activeWindow` gives the topmost one, but `Window` has no `workbook` property and the docs disagree on whether other workbooks' windows are included, so this is at best a hint that focus moved away — not a way to read or track another workbook. A Claude-style sidebar that follows the user across workbooks cannot be built as one instance; each workbook needs its own pane (auto-open per document), and there is no global view to aggregate them. **[C]** for the API surface, **[U]** for `Application.windows` scope; the strongest structural constraint found.
2. **Undo is not rollback.** Ctrl+Z does undo add-in writes, and `mergeUndoGroup` gives one undo step per `Excel.run` (1.20). But `Worksheet.delete`, `Worksheet.copy`, `Worksheet.name`, `Worksheet.position`, `Worksheet.visibility` (on Windows), all protection APIs, `Workbook.save/close`, and `AllowEditRange.*` are **not undoable**, and calling any of them **clears the user's undo stack from that call onward**. There is no programmatic undo, so a "revert this batch" affordance needs a pre-write snapshot taken by the add-in. Multi-sync runs also commit progressively, so a mid-run failure is a partial apply. **[C]** [Undo doc](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities) · [RunOptions](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21)
3. **Citations into hidden sheets are undefined.** `Range.select()`/`Worksheet.activate()` are the only navigation primitives, and neither documents behaviour for `Hidden`/`VeryHidden` sheets. A citation must either refuse hidden targets, unhide (an undo-poisons-the-stack operation, see #2), or navigate to the nearest visible ancestor. **[U]** — test in the spike.
4. **Streaming is undocumented at the Office layer.** WebSockets are explicitly supported; SSE/`fetch` streaming is not mentioned in any Office Add-ins doc. The WebView is Chromium, so it is expected to work, but CORS on the local backend and mixed content from `https://localhost` to an `http://` backend are unverified. **[U]**
5. **Per-workbook identity for history.** `Excel.SettingCollection` is correctly scoped per file + add-in, but no size cap is documented on the reference page, there is no export, and it only persists when the workbook is saved. `localStorage` on Windows is *not* partitioned (`Office.context.partitionKey` is "undefined in environments without partitioning, such as the webview controls for Office on Windows"), so keys must be built from `Office.context.document.url` (null when unavailable) or `Workbook.name` (changes on Save As). `Workbook.properties.custom` can hold a UUID but writing it is a visible workbook mutation that lands on the undo stack. **[C]** [Persist add-in state](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings) · [SettingCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.settingcollection?view=excel-js-1.21) · [Office.Document](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=excel-js-1.21)
6. **Pivot undo has holes.** Pivot creation, refresh, layout and hierarchy edits are not on the unsupported list, but every `PivotTableStyle.*` call lacks co-authoring undo and `PivotTableStyleCollection.add`/`setDefault` are explicitly unsupported. Pivot *field* undo (row/column/value/data hierarchy moves) is not itemised either way. **[C]** for the style entries, **[U]** for field-level undo.
7. **Pinning does not exist for Excel.** `SupportsPinning` is Outlook-only, so "the sidebar stays put" must come from auto-open + normal docked behaviour; the user can always close the pane. The only strong variant is the file-level `locked` attribute in the taskpane part, which prevents closing but is written via the OOXML part, not the manifest. **[C]** [Action element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/action) · [CT_OsfTaskpane](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/555d2968-2823-4cb3-b317-4511dfc97bef)
8. **Packaging is constrained to the XML manifest for a network-share install.** The unified manifest cannot be published to a network share, and the XML manifest cannot declare `ExcelApiDesktop` sets as activation requirements. Either the install path changes (dev-tool sideload) or the manifest stays XML. **[C]** [network share](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) · [ExcelApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-desktop-1-1-requirement-set)
9. **No document-close event.** Autosave/flush of conversation state cannot be triggered by a close hook; only `beforeDocumentCloseNotification` (Excel-only, modal, SharedRuntime 1.2) exists, and it is a UX-visible modal. Flushing must happen per turn or on a timer. **[C]** [Office.Addin](https://learn.microsoft.com/en-us/javascript/api/office/office.addin?view=excel-js-1.21)
10. **`delayForCellEdit` defaults to `false`.** A batch that lands while the user is typing in a cell "automatically fails". Agent-initiated applies need `delayForCellEdit: true` or explicit retry handling. **[C]** [RunOptions](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21)

**Unknowns to settle in the sideload spike (ticket 05)**

- `isSetSupported` for `ExcelApi` 1.20/1.21, `ExcelApiDesktop` 1.1, `TaskPaneApi` 1.1, `SharedRuntime` 1.2 on the actual machine.
- Hidden-sheet `Range.select()` / `Worksheet.activate()` behaviour.
- SSE through WebView2 from the pane to the local backend (CORS preflight, buffering, disconnect on pane close).
- Whether pane width/dock state survives close→reopen of the same workbook.
- Observable partial-apply behaviour when a later `context.sync()` in a `mergeUndoGroup` run fails.
- Whether the `locked` taskpane attribute is honoured for a sideloaded add-in (it is in the file format spec; no Office.js API exposes it).

---

## Sources

Office Add-ins documentation (Microsoft Learn)

1. [Excel JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets) — build/version → requirement-set table, runtime check snippet, `<Requirements>` sample.
2. [ExcelApi 1.20 requirement set](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-20-requirement-set) — "adds undo support … `RunOptions.mergeUndoGroup`".
3. [ExcelApi 1.21 requirement set](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-1-21-requirement-set) — newest set, linked-entity work.
4. [Excel JavaScript API desktop-only requirement set 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-desktop-1-1-requirement-set) — desktop-only set, superset of 1.21, XML-manifest restriction, window APIs.
5. [Task Pane API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/task-pane-api-requirement-sets) — TaskPaneApi 1.1 min build 19029.20004.
6. [Shared runtime requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/shared-runtime-requirement-sets) — SharedRuntime 1.1/1.2 minimums, 1.2 Excel-only.
7. [Task panes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/task-pane-add-ins) — placement, Excel design size, personality menu.
8. [Runtimes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes) — runtime types, lifespans, WebSockets/CORS/storage matrix.
9. [Configure your Office Add-in to use a shared runtime](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime) — manifest setup, one-task-pane rule, CORS/global sharing.
10. [Browsers and webview controls used by Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins) — WebView2 on Windows, PDF/scrollbar/WIP quirks.
11. [Automatically open a task pane with a document](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document) — `Office.AutoShowTaskpaneWithDocument`, document scope, min build 8121.1000.
12. [Show or hide the task pane](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/show-hide-add-in) — `showAsTaskpane`/`hide`, state preservation, `onVisibilityModeChanged`.
13. [Persist add-in state and settings](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings) — per-file + add-in settings, `partitionKey` undefined on Windows webviews, storage limits caveats.
14. [Undo support with the Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities) — undo baseline, grouping, `UndoNotSupported`, unsupported-API table.
15. [Work with Events using the Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events) — event catalogue, handler lifecycle, `enableEvents`, coauthoring `source`.
16. [Activate add-ins with events](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation) — `OnDocumentOpened`, admin scope, JS-only runtime, 300 s timeout, Fetch/XHR guidance.
17. [SourceLocation element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/sourcelocation) — HTTPS requirement.
18. [Action element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/action) — `SupportsPinning` "Supported in Outlook only".
19. [Implement a pinnable task pane in Outlook](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/pinnable-taskpane) — what pinning actually does.
20. [Sideload Office Add-ins on Windows from a network share](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) — trusted catalog, registry key, unified-manifest limitation.
21. [Sideload Office Add-ins that use the unified manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest) — Windows 2304 (16320.20000)+.
22. [Test Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins) — sideload matrix by platform/manifest.
23. [PowerPoint add-in tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/powerpoint-tutorial-yo?tabs=jsonmanifest) — "use HTTPS, not HTTP, even while developing", WebView2 localhost loopback exemption.

API reference (Microsoft Learn, `view=excel-js-1.21` unless noted)

24. [Excel.Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) — entry object, events (`onActivated` 1.13, `onSelectionChanged` 1.2), identity properties, `settings`, `close`/`save`.
25. [Excel.Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21) — `getUsedRange` 1.2, `activate` 1.1, event sets, `tabId` 1.14, `autoFilter` 1.9.
26. [Excel.WorksheetCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-1.21) — `add` 1.1, collection events.
27. [Excel.Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) — values/formulas/formats 1.1, `select` 1.1, `dataValidation` 1.8, `conditionalFormats` 1.6, `getDirectPrecedents` 1.12, `getDirectDependents` 1.13, `getPrecedents` 1.14, `getDependents` 1.15, `copyFrom` 1.9.
28. [Excel.RangeAreas](https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeareas?view=excel-js-1.21) — `select()` 1.18.
29. [Excel.PivotTable](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivottable?view=excel-js-1.21) — class base 1.3, layout/hierarchies 1.8, `refreshOnOpen` 1.13.
30. [Excel.PivotTableCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivottablecollection?view=excel-js-1.21) — `add` 1.8, `refreshAll` 1.3.
31. [Excel.Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table?view=excel-js-1.21) — table API 1.1, sort/filters 1.2, AutoFilter 1.9.
32. [Excel.AutoFilter](https://learn.microsoft.com/en-us/javascript/api/excel/excel.autofilter?view=excel-js-1.21) — 1.9 baseline, some criteria APIs 1.14.
33. [Excel.DataValidation](https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidation?view=excel-js-1.21) — 1.8, `getInvalidCells` 1.9.
34. [Excel.ConditionalFormat](https://learn.microsoft.com/en-us/javascript/api/excel/excel.conditionalformat?view=excel-js-1.21) — 1.6, rule conversion 1.17.
35. [Excel.SettingCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.settingcollection?view=excel-js-1.21) — per file + add-in scope, 1.4.
36. [Excel.DocumentProperties](https://learn.microsoft.com/en-us/javascript/api/excel/excel.documentproperties?view=excel-js-1.21) — author/creationDate/custom.
37. [Excel.Window](https://learn.microsoft.com/en-us/javascript/api/excel/excel.window?view=excel-js-1.21) — ExcelApiDesktop 1.1, "a window in the workbook", no `workbook` property.
38. [Excel.Application](https://learn.microsoft.com/en-us/javascript/api/excel/excel.application?view=excel-js-preview) — `activeWindow`, `windows` ("all the open Excel windows"), desktop-only.
39. [Excel.WindowCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.windowcollection?view=excel-js-1.21) — used by `Excel.Application: windows`.
40. [Excel.RunOptions](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21) — `mergeUndoGroup` and `delayForCellEdit` (1.20).
41. [Office.TaskPane](https://learn.microsoft.com/en-us/javascript/api/office/office.taskpane?view=excel-js-1.21) — `setWidth`, Windows 86 px–50 %, default 320 px.
42. [Office.ExtensionLifeCycle](https://learn.microsoft.com/en-us/javascript/api/office/office.extensionlifecycle?view=excel-js-1.21) — `Office.extensionLifeCycle.taskpane`.
43. [Office.Addin](https://learn.microsoft.com/en-us/javascript/api/office/office.addin?view=excel-js-1.21) — `showAsTaskpane`/`hide`/`setStartupBehavior`/`onVisibilityModeChanged` (SharedRuntime 1.1), `beforeDocumentCloseNotification` (1.2).
44. [Office.EventType](https://learn.microsoft.com/en-us/javascript/api/office/office.eventtype?view=excel-js-1.21) — `DocumentSelectionChanged` "Excel or Word".
45. [Office.Document](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=excel-js-1.21) — `url` returns null when unavailable, `settings`.

Format specification

46. [MS-OWEXML 2.2.8 CT_OsfTaskpane](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/555d2968-2823-4cb3-b317-4511dfc97bef) — persisted task pane: `dockstate`, `visibility`, `width`, `row`, `locked`.
47. [MS-OWEXML 3.2 Taskpane Web Extension File](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/3d04f8ce-65f2-4dc3-bafa-636d0a7e41a1) — `taskpanes.xml` example with `dockstate="right" visibility="1" width="408"`.
