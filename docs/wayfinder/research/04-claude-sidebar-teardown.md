# Claude for Excel sidebar — interaction teardown

Ticket: [`04-claude-sidebar-teardown`](../issues/04-claude-sidebar-teardown.md)
Researched: 2026-11-07 · Sources: Anthropic docs, Anthropic blog + newsroom, Anthropic help center (incl. an archived pre-GA version), Claude Academy, and official Anthropic screenshots/video.
Third-party material was used only to *locate* official assets; no third-party claim stands on its own below.

**Legend.** **[C]** = confirmed by a primary Anthropic source (quoted or directly shown). **[I]** = inferred — a reasonable read of primary material, but not stated by Anthropic. Anything unmarked is **[C]**.
Where a claim rests on a screenshot I give the asset URL; where it rests on a doc, the page URL.

**Version note.** Claude for Excel went beta as a research preview on 2025-10-27, beta to all Max/Team/Enterprise on 2025-11-24, and GA on 2026-05-07. The docs snapshot used here is dated 2026-10-31. The pre-GA help article (archived 2026-02-14) is quoted where it describes UI detail the current docs have dropped; treat those specifics as *historical* where flagged.

---

## Summary

1. **It is a right-docked Office.js task pane, not a custom window.** Served from `https://pivot.claude.ai/taskpane.html`, one pane per workbook, and the pane can auto-open with the file via `Office.AutoShowTaskpaneWithDocument` ([data storage](https://claude.com/docs/office-agents/data-storage)).
2. **The chrome is minimal.** The one official product screenshot shows a floating cream card with a single ⚙ at top-right, a scrolling transcript, thumbs up/down at the end of each answer, and a composer that carries a **workflow action row** — `⇄ Accept all edits ⌄` on the left, round send arrow on the right ([screenshot](https://cdn.sanity.io/images/4zrzovbb/website/450687b3c7a1189eb1a650ade77588c56ca01461-2560x1440.png)).
3. **Citations are inline chips, not footnotes.** Cell and range references render as bordered pills (`B22:F25`) inside the sentence and navigate the grid on click — "cell-level citations you can click to navigate to the referenced cell" ([docs](https://claude.com/docs/office-agents/excel)).
4. **Edits are batched and approved, not applied silently.** Anthropic's own screenshot shows a single "Accept all edits" control for a whole turn; docs add **overwrite protection** and **per-risky-tool confirmation pop-ups** ([docs](https://claude.com/docs/office-agents/excel), [archived help](https://web.archive.org/web/20260214042810/https://support.claude.com/en/articles/12650343-using-claude-in-excel)).
5. **The review surface is the spreadsheet itself**: Claude "highlights every cell it updates and provides explanatory comments" and the undo story is Excel's own Ctrl+Z. There is **no confirmed side-by-side diff** — that is our gap to fill, not theirs to copy.
6. **History is local-first and app-scoped**: IndexedDB, keyed to *user × organization × Office app*, capped at 50 conversations, cleared under Settings → Chat history → Delete all, with each Excel conversation also storing the workbook's path/URL ([data storage](https://claude.com/docs/office-agents/data-storage)).
7. **Instructions are a single free-text field in Settings**, scoped to Excel only, separate from Word/PowerPoint ([docs](https://claude.com/docs/office-agents/excel)).
8. **Streaming is real and visibly narrated.** The add-in makes a helper model call whose only job is to "write the status line shown while Claude thinks", and it streams tokens with a first-token event and a user-triggered abort ([data storage](https://claude.com/docs/office-agents/data-storage), [OpenTelemetry](https://claude.com/docs/office-agents/opentelemetry)).

---

## Layout

### The pane
- **Right-docked task pane inside Excel**, opened from Home → Add-ins on Windows or Tools → Add-ins on Mac, and also openable with a keyboard shortcut: **Ctrl+Option+C (Mac) / Ctrl+Alt+C (Windows)** ([docs](https://claude.com/docs/office-agents/excel), [Claude Academy](https://academy.claude.com/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning), [release notes 2025-11-24](https://support.claude.com/en/articles/12138966-release-notes)).
- **One pane per workbook.** "Workbooks that you open in the same instance of Excel share one Excel process, and each workbook has its own Claude pane" ([performance and limits](https://claude.com/docs/office-agents/performance)).
- **Width.** In Anthropic's official product screenshot the panel occupies **633 px of a 2560 px frame ≈ 25 %** of window width, full height minus ~10 % margins ([screenshot](https://cdn.sanity.io/images/4zrzovbb/website/450687b3c7a1189eb1a650ade77588c56ca01461-2560x1440.png)). Because it is an Office task pane it inherits Office's resizable-pane behavior. **[I]** Treat ~25–28 % of window (roughly 380–450 px at 1440–1600 px wide) as the target, user-resizable, with a min width that keeps citation chips on one line.

### Chrome inventory

| Element | State | Evidence |
|---|---|---|
| Settings (⚙) | **Confirmed** — top-right of the pane; opens a Settings surface containing Instructions, Chat history, Session logging, cross-app toggle | Screenshot shows ⚙ top-right; docs say "Open Settings in the add-in sidebar" and "cleared from Settings" ([docs](https://claude.com/docs/office-agents/excel), [data storage](https://claude.com/docs/office-agents/data-storage)) |
| Model picker | **Confirmed that it exists**; placement **inferred** | "the list you see in the add-in can be shorter than what you see in Claude.ai" ([docs](https://claude.com/docs/office-agents/excel)); the beta article answered "You can switch between Sonnet 4.5 and Opus 4.6" ([archived help](https://web.archive.org/web/20260214042810/https://support.claude.com/en/articles/12650343-using-claude-in-excel)). No primary source shows *where* it sits. **[I]** Likely a header chip or a Settings row — the Oct 2025 screenshot shows no persistent title row, only the gear. |
| New chat | **Confirmed as an action**; no documented button label | "deleted when the user … starts a new chat" ([data storage](https://claude.com/docs/office-agents/data-storage)) |
| History | **Confirmed as a concept and a settings section**; entry point **inferred** | Settings → "Chat history" → "Delete all" ([data storage](https://claude.com/docs/office-agents/data-storage)). No primary source shows a history *list* control; the natural reading is a clock/list icon in the header. **[I]** |
| Workbook / file context | **Confirmed at the data layer, invisible in chrome** | Each Excel conversation stores "the document's file path or URL"; `claude.fileId` (a random id written into the workbook) lets the add-in recognise the file after rename or Save As ([data storage](https://claude.com/docs/office-agents/data-storage)). Nothing indicates a persistent "workbook name" banner in the pane. |
| Connected-app indicators | **Confirmed**, cross-app mode only | "connected-app indicators appear in the sidebar when other Excel, PowerPoint, Word, or Outlook sessions are linked" ([work across apps](https://claude.com/docs/office-agents/work-across-apps)) |
| Composer | **Confirmed** | Large rounded field, placeholder "What can I do for you?", action row inside the field ([screenshot](https://cdn.sanity.io/images/4zrzovbb/website/450687b3c7a1189eb1a650ade77588c56ca01461-2560x1440.png)) |
| `+` button below the composer | **Confirmed** — opens **Connectors** | [Connectors and Skills](https://claude.com/docs/office-agents/connectors-and-skills) |
| `/` slash menu | **Confirmed** — Skills for the current app only | [Connectors and Skills](https://claude.com/docs/office-agents/connectors-and-skills) |
| Microphone | **Confirmed** — right side of the composer; hover reveals a device-picker arrow | [Dictation](https://claude.com/docs/office-agents/dictation) |
| Response feedback | **Confirmed** — thumbs up/down under the answer | Screenshot; docs describe rating a response and adding a comment ([data storage](https://claude.com/docs/office-agents/data-storage)) |
| Overflow `⋮` | **Confirmed** — contains **Download Transcript** | [Performance and limits](https://claude.com/docs/office-agents/performance) (troubleshooting flow) |

### What is *not* in the chrome
- No document tab strip, no sheet navigator, no formula bar, no minimap. The workbook is the canvas; the pane is a conversation column. The only structural affordance toward the grid is the citation chip.
- No visible "context attached: Sheet1!A1:F200" chip in any official asset. Context is implicit ("Claude reads all four sheets first" — [Claude Academy](https://academy.claude.com/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning)). **[I]** This is a deliberate removal of the "attach files" mental model that desktop chat UIs have; it costs discoverability but buys a clean pane.

---

## Chat and citations

### Message styling
- **User turn**: plain text, no bubble visible in the official assets; the transcript reads as a single column with generous line height on cream `#FAF9F5`.
- **Assistant turn**: body sans-serif in near-black `#141413`, with **bold serif run-in headings** for sections inside a single answer ("**Formulas remain in black** (as they should):") and bullet lists. Numbers and cell references are monospaced.
- **Emphasis conventions**: Claude explains financial-modeling color conventions in prose (blue = hardcoded input), which is a content habit worth noting, not chrome.

### Streaming and progress
- **Token streaming is confirmed.** Telemetry has an `agent.stream` span that records `first_token`, `stream_complete`, `stream_error`, plus `agent.message_count` ([OpenTelemetry](https://claude.com/docs/office-agents/opentelemetry)).
- **A status line narrates the wait.** The add-in issues a *separate* helper model request to "write the status line shown while Claude thinks" ([data storage](https://claude.com/docs/office-agents/data-storage)). This is the single most copyable detail: the wait is described in domain language ("Reading Headcount!A1:F400…") rather than a spinner.
- **Stop is a real control.** Aborts produce `APIUserAbortError` when "the user stops the call" ([OpenTelemetry](https://claude.com/docs/office-agents/opentelemetry)); docs describe "Stop and a timeout end Claude's wait, not the step Office already started" ([performance](https://claude.com/docs/office-agents/performance)).
- **Long reads are chunked and narrated**: one read returns at most 2,000 cells with data, so a big sheet is read "in several steps" ([performance](https://claude.com/docs/office-agents/performance)).

### Markdown, tables, code
- Official screenshots and the Academy walkthrough show **bullets, bold headings, and a flow diagram** ("returns a visual diagram showing how data flows from Assumptions through to Summary") ([Claude Academy](https://academy.claude.com/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning)).
- **Tables appear as before/after comparisons** ("a before/after comparison by department") — same Academy page.
- **Formula text renders inline in monospace**, e.g. the color-convention explanation in the official screenshot.
- **[I]** Rendered markdown is the safe assumption for the transcript (headings, lists, tables, inline code, fenced code). No primary source shows a fenced code block inside Excel, and nothing suggests syntax highlighting matters here.

### Cell-level citations
- **Confirmed and central to the product**: "Answers include cell-level citations you can click to navigate to the referenced cell" ([docs](https://claude.com/docs/office-agents/excel)); the launch post: "lets users navigate directly to the cells it references in its explanations" ([news, 2025-10-27](https://www.anthropic.com/news/advancing-claude-for-financial-services)).
- **Visual form** (from the official screenshot): a **rounded, cool-tinted pill** — fill `#EFF3F5`, 1 px hairline border, text `#698DC0`, monospace — carrying a range like `B22:F25`, followed by the human label ("Operating expenses (G&A, D&A, Pre-opening costs, Impairment)"). Each bullet begins with a ✓ or •, so a citation list doubles as a checklist.
- **Granularity**: ranges, not single cells (`B22:F25`, `B30:F30`, `B34:F34`) — because the model naturally reasons over row blocks. **[I]** Single-cell chips are a degenerate range; support both with one component.
- **No "Sources" footer.** Citations are inline only. There is no numbered reference list at the end of an answer in any official asset. **[I]** Anthropic chose in-place verification over footnotes because the user's next action is to look at the cell.
- **Click behavior**: navigate the grid to the referenced cell/range. **[I]** Selection + scroll, not a new sheet tab; combining sheets (e.g. `Summary!B12`) must work across tabs, which the docs imply ("Work across multi-tab workbooks").

### Conversational control
- **Claude asks before guessing, with clickable options.** In the Academy walkthrough the model "explains the root cause, shows you the options (0, 1, or a custom value), and lets you decide" ([Claude Academy](https://academy.claude.com/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning)). **[I]** That is a multiple-choice card in the transcript, not a free-text question — worth copying for ambiguity resolution.
- **It reads before it answers** ("Claude reads all four sheets first"), and a single answer can span tabs.

---

## Edit review

### How a change is proposed
1. **Claude acts, then the pane holds the result for approval.** The composer's action row becomes `⇄ Accept all edits ⌄` — one control for the whole turn, with a chevron implying a per-change list ([screenshot](https://cdn.sanity.io/images/4zrzovbb/website/450687b3c7a1189eb1a650ade77588c56ca01461-2560x1440.png)).
2. **Every touched cell is marked up in the grid.** "Claude highlights every cell it updates and provides explanatory comments" ([archived help](https://web.archive.org/web/20260214042810/https://support.claude.com/en/articles/12650343-using-claude-in-excel)). Earlier marketing says the same: "it tracks and explains its changes" ([news](https://www.anthropic.com/news/advancing-claude-for-financial-services)).
3. **Formula integrity is the promise, not a diff.** "Claude updates cell values while keeping formula relationships intact, so downstream cells recompute correctly" ([docs](https://claude.com/docs/office-agents/excel)).

### Batch vs per-change
- **Batch is the documented default** ("Accept all edits", one per turn).
- **Per-change exists only behind the chevron** **[I]** — the dropdown affordance is visible, the contents are not documented.
- **A separate, harsher gate exists for dangerous operations**: "Users can approve all of Claude's actions via a confirmation pop-up that appears when each tool is triggered", listing external fetch (`WEBSERVICE`, `STOCKHISTORY`, `RTD`…), imports (`IMPORTDATA`, `IMPORTXML`…), `INDIRECT`, `DDE`, `CALL`/`EVALUATE`, file-system functions (`FILES`, `FOPEN`, `FWRITE`), and `REGISTER.ID` ([archived help](https://web.archive.org/web/20260214042810/https://support.claude.com/en/articles/12650343-using-claude-in-excel)). Current docs compress this to "When Claude proposes a risky operation, you are asked to confirm before it runs" ([docs](https://claude.com/docs/office-agents/excel)).
- **Tool-call approval is a first-class state**: crash-detection records are deleted "while Claude waits for the user to answer a question or approve an action" ([data storage](https://claude.com/docs/office-agents/data-storage)).

### Preview / diff
- **Confirmed**: highlight + explanatory comments **in the workbook**. The explanation lives where the change lands, not in the pane.
- **Not confirmed**: any side-by-side or unified diff view, before/after value table for edits (the before/after table in the Academy material is a *scenario* result, not an edit diff), or a review list with per-row accept/reject.
- **[I]** Anthropic's model is "trust the grid": mark cells, comment them, let the user eyeball. Our v1 wants a diff, which is a genuine improvement, not a copy.

### Apply, overwrite protection, undo
- **Overwrite protection is an explicit feature**: "Claude warns you before overwriting existing data to avoid accidental data loss" ([docs](https://claude.com/docs/office-agents/excel)).
- **Undo is Excel's undo**: "undo (Cmd+Z or Ctrl+Z) reverses most actions" ([Claude Academy](https://academy.claude.com/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning)); "You can always undo changes using Excel's standard undo function" ([archived help](https://web.archive.org/web/20260214042810/https://support.claude.com/en/articles/12650343-using-claude-in-excel)). There is no in-pane undo stack documented.
- **Apply is through the Excel object model with hard limits**: 2,000 cells per read, 50 queued command batches, 90 s / 5 min step waits ([performance](https://claude.com/docs/office-agents/performance)) — the review UX has to survive a partial apply, so batching and idempotency are not optional.
- **No confirmation for ordinary writes** — only risky functions and overwrite. **[I]** That is the right default: approve the turn, not every cell.

### What this means for "formula-safe"
Anthropic never shows a formula-level diff. Their safety argument is structural (keep dependencies intact) plus post-hoc (highlight + comments + undo). Our planned "preview the exact cell writes, then approve" is strictly stronger; we should keep their *post-hoc* markup as the second line of defense after our pre-apply preview.

---

## History and instructions

### Storage model (confirmed)
- **Local-first**: chat history lives in **IndexedDB in the Office webview**, not on Anthropic's servers, and is not synced between devices ([data storage](https://claude.com/docs/office-agents/data-storage)).
- **Keyed to user × organization × Office app.** "Each conversation is stored against a user identity, an organization, and the Office app it was created in. The list shows only the conversations that match all three" ([data storage](https://claude.com/docs/office-agents/data-storage)). Signing out and back in preserves them; changing organizations hides them; changing devices loses them.
- **Per-workbook association exists at the record level**: in Excel each saved conversation keeps "the document's file path or URL", and `claude.fileId` inside the workbook survives rename/Save As and lets the pane re-associate ([data storage](https://claude.com/docs/office-agents/data-storage)). Marketing states the user-facing version: "Conversations persist with each file, so you can close the sidebar, reopen it the next day, and pick up where you left off" ([GA blog](https://claude.com/blog/collaborate-with-claude-across-excel-powerpoint-word-and-outlook)).
- **Retention**: the **50 most recent** conversations per user/org/app; older ones are deleted automatically; the browser may evict the oldest under storage pressure ([data storage](https://claude.com/docs/office-agents/data-storage)).
- **Clearing**: Settings → **Chat history** → **Delete all** removes every saved conversation for that user in that app and starts a new chat ([data storage](https://claude.com/docs/office-agents/data-storage)).
- **Opening the panel with the file**: the add-in writes `Office.AutoShowTaskpaneWithDocument` so Office opens the Claude pane when the document opens (works for centrally deployed/sideloaded manifests) ([data storage](https://claude.com/docs/office-agents/data-storage)).
- **Reopen** of a saved conversation replaces the working context with that conversation's stored parts ([data storage](https://claude.com/docs/office-agents/data-storage)) — i.e. history is a first-class resume, not a read-only log.

### Auto-compaction
- "Longer conversations are automatically compacted into new conversations to avoid running out of context" ([docs](https://claude.com/docs/office-agents/excel)); telemetry has `agent.compaction` spans with `compaction_start/complete/error` ([OpenTelemetry](https://claude.com/docs/office-agents/opentelemetry)).
- Note the phrasing: compaction **into new conversations** — the sibling thread is the escape hatch, not a silent summarisation in place.
- **[I]** The user is told something; whether it is a banner, a divider in the transcript, or a toast is not documented.

### Session logging → a log sheet in the workbook
- A setting, **off by default**, that keeps a sheet named **`Claude Log`** as the **first tab**, adding "a row for each exchange, with the date, a summary of what the user asked, what Claude did, and the result" ([data storage](https://claude.com/docs/office-agents/data-storage)). It is ordinary workbook content: it travels with the file and persists after the setting is turned off.
- This is the only place Claude writes prose into the file itself besides the edits. Useful as an audit trail for finance users; hostile to "the model never touches the workbook unasked".

### Instructions
- **Where**: Settings → **Instructions** field in the add-in sidebar ([docs](https://claude.com/docs/office-agents/excel)).
- **Scope**: "apply to every conversation in Excel"; "Instructions you set in Excel only apply to Excel. They are separate from Instructions you set in PowerPoint or Word" ([docs](https://claude.com/docs/office-agents/excel)). So: per-app, not per-workbook, not per-user-global.
- **Examples Anthropic gives**: "format numbers with thousand separators", "always bold column headers", currency or locale preferences, recurring workflow context ([docs](https://claude.com/docs/office-agents/excel)). Note: all *formatting and convention* instructions, none about analytical method.
- **Adjacent behaviour**: in Outlook, learning draft preferences writes "a summary of the draft preferences … to the Instructions field in Settings", between marker lines, and the user can delete it ([data storage](https://claude.com/docs/office-agents/data-storage)). This shows Instructions is a plain mutable text blob that other features append to — a design smell worth avoiding in our v1.

### Skills and connectors (entry points only — out of v1)
- **Connectors**: `+` below the chat input → **Connectors**; the pane calls remote MCP servers directly from the machine (CORS to `pivot.claude.ai`) ([connectors and skills](https://claude.com/docs/office-agents/connectors-and-skills), [third-party platforms](https://claude.com/docs/office-agents/third-party-platforms)).
- **Skills**: invoked implicitly, or explicitly by typing `/` for the current app's list; a Skills menu offers "Save to claude.ai" with a cloud icon per uploaded skill ([connectors and skills](https://claude.com/docs/office-agents/connectors-and-skills), [data storage](https://claude.com/docs/office-agents/data-storage)).
- Both are configuration surfaces, not pane chrome. Nothing here needs copying for v1.

---

## States

**Empty / first run**
- Composer placeholder is **"What can I do for you?"** ([screenshot](https://cdn.sanity.io/images/4zrzovbb/website/450687b3c7a1189eb1a650ade77588c56ca01461-2560x1440.png)) — same string as the Claude app, so the pane feels like Claude rather than like an Office add-in.
- A first-run flow exists ("the add-in … shows its first-run screens"), plus onboarding/terms flags in local storage ([data storage](https://claude.com/docs/office-agents/data-storage)).
- **Sign-in is a gate before the transcript**: "Open Excel, activate the add-in, and sign in with your Claude account" ([docs](https://claude.com/docs/office-agents/excel)); the sign-in/approval pages run an hCaptcha browser check ([third-party platforms](https://claude.com/docs/office-agents/third-party-platforms)).

**Loading / thinking**
- A live **status line** written by a helper model call (see Chat). **[I]** It replaces the first token's slot until streaming starts.
- Long operations show **step-by-step progress**, not a percentage: reads are chunked at 2,000 cells and the model narrates each step ([performance](https://claude.com/docs/office-agents/performance)).

**Timeout**
- Explicit, documented state: "Claude reports a timeout" while Office may still be running the step; waits are **90 s** for code-execution steps, **5 min** for other steps on desktop, **2 min** on the web ([performance](https://claude.com/docs/office-agents/performance)).
- The recovery advice is "wait, save, then ask for a smaller step" — the product does not retry behind the user's back. **[I]** An inline "Office is still working" card, not a modal.

**Error / connection failure**
- A **connection-failed screen** exists with a **"Copy error details"** paste that Anthropic's own debug tooling reads ([third-party platforms](https://claude.com/docs/office-agents/third-party-platforms)).
- Crash/hang detection: `claude-office-inflight` records are reconciled on next open after 10 minutes and reported as a usage event ([data storage](https://claude.com/docs/office-agents/data-storage)).
- Errors go to Sentry with the error text and the failing tool name ([data storage](https://claude.com/docs/office-agents/data-storage)).

**Rate limit / usage limit**
- The add-in **inherits the Claude account's usage limits**: "subject to the same usage limits" ([docs](https://claude.com/docs/office-agents/excel)); Anthropic links to the generic limits article rather than an add-in-specific one ([understanding usage and length limits](https://support.claude.com/en/articles/11647753-understanding-usage-and-length-limits)).
- Context exhaustion is handled by **auto-compaction into a new conversation**, framed as avoiding "length limit errors" ([docs](https://claude.com/docs/office-agents/excel), [release notes 2025-11-24](https://support.claude.com/en/articles/12138966-release-notes)).
- No primary source shows the rate-limit UI copy. **[I]** A dismissible banner plus a disabled composer, since the composer is the only input surface.

**Offline**
- Not documented anywhere. **[I]** The add-in is a webview calling `pivot.claude.ai` and `api.anthropic.com` ([third-party platforms](https://claude.com/docs/office-agents/third-party-platforms)), so offline lands in the same failure path as a blocked domain — the connection-failed screen, not a distinct offline state.
- **Local data survives**: chat history is on disk ([data storage](https://claude.com/docs/office-agents/data-storage)), so a good offline design would let history render and queue nothing. Anthropic does not claim this.

**Keyboard**
- **Ctrl+Option+C (Mac) / Ctrl+Alt+C (Windows)** opens Claude in Excel ([release notes](https://support.claude.com/en/articles/12138966-release-notes), [Claude Academy](https://academy.claude.com/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning)).
- **Enter** sends; during dictation Enter "stop[s] and send[s] in one step" ([dictation](https://claude.com/docs/office-agents/dictation)).
- **Ctrl/Cmd+Z** is the undo path for edits ([Claude Academy](https://academy.claude.com/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning)).
- No documented shortcut for new chat, history, or stop. **[I]** Stop is almost certainly Esc-or-button; the OTel span proves the abort path exists.

**Dictation**
- Mic icon on the right of the composer; placeholder switches to **"Listening…"** and the button highlights; words stream into the composer in real time; hover the mic for a device picker ([dictation](https://claude.com/docs/office-agents/dictation)).
- **Desktop only** (browser add-ins cannot reach the mic) and **direct-Claude-auth only** (not with Bedrock/Vertex/Foundry/gateway) ([dictation](https://claude.com/docs/office-agents/dictation)).
- Audio is streamed to Anthropic and not retained; only the transcript text stays ([dictation](https://claude.com/docs/office-agents/dictation)).
- **[I]** For a local-first v1, dictation is the one feature that forces an explicit "this leaves the machine" disclosure. Defer it or gate it behind a visible local/remote indicator.

---

## Visual language

Measured from Anthropic's official product screenshot ([asset](https://cdn.sanity.io/images/4zrzovbb/website/450687b3c7a1189eb1a650ade77588c56ca01461-2560x1440.png)) by pixel sampling, plus the docs site theme. Hex values are exact samples, not brand-book values.

| Token | Value | Use in the shot |
|---|---|---|
| Surface / pane background | `#FAF9F5` | The whole card — Anthropic's warm cream, not white |
| Composer surface | `#FFFFFF` | Composer sits *lighter* than the pane, one step forward |
| Ink | `#141413` | Headings and body; near-black with a warm cast |
| Muted body | `#6B6A68` | Secondary sentences, labels |
| Citation chip fill | `#EFF3F5` | Cool grey-blue, deliberately *outside* the warm ramp |
| Citation link text | `#698DC0` | Muted blue, monospace, no underline |
| Send button / accent | `#DF8B70` (with `#DDB4A4` ring) | Round 36 px-ish button, right-aligned in the composer |
| Backdrop accent (marketing) | `#D97757` | Claude's brand orange — appears in marketing frames, **not** inside the pane |
| Hairlines | ~`#E8E6E0` **[I]** | 1 px borders on chips and the composer |

**Type.**
- **Serif for in-answer run-in headings** (the bold "Formulas remain in black" line is a serif) and **sans for body/UI**. This serif-in-the-answer move is what makes it read as "a document Claude wrote" rather than a chat log.
- Cell references and numbers are **monospaced**, in the same size as body text — they read as data, not as code.
- Type scale is tight: body ≈ 15–16 px, headings one step up, generous line-height (~1.5). No giant H1 inside the transcript.

**Spacing.** One column, ~20–24 px side padding, ~16–20 px between message blocks, ~8 px inside citation bullets. The composer is inset and floats over the transcript with a subtle shadow.

**Icons.** Line icons, 1.5 px stroke, no fills: gear, chevron, thumbs, microphone, arrow-up, plus. The only filled/organic mark is the Claude starburst, used as the product identity in marketing and in the Claude Academy pane mock ([Academy og image](https://academy.claude.com/og/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning.png)).

**Dark mode.** Not documented for the add-in. The docs site supports light/dark with `#09090B` background and `#D4A27F` accent ([docs theme config](https://claude.com/docs/office-agents/excel)). **[I]** The task pane almost certainly follows the Office theme via `Office.context.officeTheme`; our v1 should read that rather than shipping a manual toggle. Office dark mode is near-black with `#D4A27F`-ish accents in Fluent — a warm-cream-to-charcoal inversion, not a blue-grey one.

**What the palette is doing, in one line.** Warm paper for the surface, cool blue only for *evidence*, orange only for *action*. Nothing else is coloured. That three-role discipline is the whole system.

---

## Copy vs avoid

Mapped to the four v1 pillars. "Copy" = do this; "Adapt" = same idea, our mechanics; "Avoid" = don't.

### 1. Streaming + citations

| Behaviour | Verdict | Why |
|---|---|---|
| Inline cell/range **chips** in the sentence (`B22:F25`) | **Copy** | Verification is the job. A chip is a one-click proof, a footnote is a scroll. |
| Chip = fill `#EFF3F5`, text `#698DC0`, mono, hairline | **Copy** | Reads as evidence, not a hyperlink; survives dark mode if you re-tint by role. |
| Click → select + scroll the range | **Copy** | The only navigation affordance in the pane. Support cross-sheet and cross-tab. |
| **Status line written by a helper call** in domain language | **Copy** | Cheapest perceived-latency win available; forces the model to say *what* it is doing. |
| 2,000-cell read chunks narrated as steps | **Copy** | Honest progress for big workbooks; matches our context caps. |
| No sources footer | **Copy the omission** | Inline only. A reference list at the bottom is dead weight in a spreadsheet. |
| Assistant headings in **serif**, body sans | **Copy** | Makes the answer feel authored; costs one font. |
| Thumbs up/down per answer | **Skip for v1** | No training loop to feed locally. |
| Multiple-choice decision card ("0, 1, or custom") | **Copy** | Turns ambiguity into one click instead of a round trip. |

### 2. Formula-safe approved edits

| Behaviour | Verdict | Why |
|---|---|---|
| **Batch approve the turn** ("Accept all edits ⌄") | **Copy** | Approval must be cheap or users will rubber-stamp. One gate per turn. |
| Chevron → per-change list | **Copy, and go further** | Anthropic hides it; we should make the list the primary surface with per-row toggles. |
| Highlight changed cells in the grid | **Copy** | Second line of defense after our preview; also the artifact users screenshot for review. |
| **Explanatory comments** on changed cells | **Adapt** | Write the rationale as an Excel note per changed block, not per cell. |
| **Pre-apply diff** (before → after values/formulas) | **Not present upstream — build it** | Our differentiator. Anthropic's "trust the grid" is weaker for formula rewrites. |
| **Overwrite protection** (warn before clobbering non-empty cells) | **Copy** | Cheap, and the single most credible safety signal in the docs. |
| Dangerous-operation confirmation (WEBSERVICE, INDIRECT, DDE, FILES…) | **Copy the tiering, narrow the list** | Two tiers (turn approval + risky-op confirm) is right. Our risky tier: external links, `INDIRECT`, volatile arrays, >N cells. |
| Undo = Excel's Ctrl+Z | **Copy, but add an in-pane "Undo this turn"** | Excel's undo stack is fine until a 400-cell apply; one-click turn undo is better. |
| Silent apply on ordinary writes | **Copy** | Per-cell confirmation would make the product unusable. |

### 3. Native Excel ops

| Behaviour | Verdict | Why |
|---|---|---|
| "Ask for these directly" — sort, filter, pivot edit, conditional format, data validation ([docs](https://claude.com/docs/office-agents/excel)) | **Copy the framing** | No tool menu, no slash commands for operations. Natural language is the API. |
| Ops reported back in the transcript with what changed | **Copy** | The user must be able to see "sorted A2:F500 by revenue desc" without hunting. |
| Hard platform limits surfaced in-product (50 queued batches, step timeouts) | **Copy** | Surface them as "this is big; split it" before the failure, not after. |
| Data tables / macros / VBA unsupported | **Copy the boundary** | Same exclusions are already agreed for us. |

### 4. Per-workbook history + instructions

| Behaviour | Verdict | Why |
|---|---|---|
| **Local IndexedDB**, no server copy, no cross-device sync | **Copy** | Exactly our local-first constraint; also removes a whole privacy story. |
| Conversation keyed to the **file** (path + stable `fileId`) | **Copy the `fileId` trick** | Survives rename/Save As; a path alone does not. |
| History scoped to **app**, not workbook | **Avoid** | The list bleeds across workbooks; our v1 should scope the list per workbook and keep an "all conversations" escape hatch. |
| **50-conversation cap**, oldest auto-deleted | **Adapt** | A cap is right; 50 is arbitrary. Make it configurable and never silently delete the conversation for the *current* workbook. |
| Settings → Chat history → **Delete all** | **Copy** | One obvious nuclear option; per-conversation delete alongside it. |
| **Auto-compaction into a new conversation** | **Adapt** | Compact in place with a visible divider; only branch when the user asks. Silent thread-splitting breaks "where did my conversation go". |
| Auto-open pane with the document (`AutoShowTaskpaneWithDocument`) | **Copy** | Big perceived-integration win; opt-out in settings. |
| **Session logging → `Claude Log` first sheet** | **Avoid by default** | Writing a log sheet into the user's workbook is invasive and fights "don't touch what I didn't ask for". Offer an export instead. |
| Instructions: single free-text field, **per app** | **Adapt** | Per-workbook is our scope; keep the same single-field simplicity, but scope it to the file and show which file it applies to. |
| Formatting/convention examples in the field | **Copy the examples** | "Thousands separators", "bold headers", locale — they teach the field's purpose better than a description does. |
| Other features appending text into Instructions (Outlook style summary) | **Avoid** | Never let a feature mutate the user's instruction text in place. |

### Cross-cutting: do not copy
- **Settings-only model picker.** If the model is swappable and user-visible, it belongs one click from the composer, not two menus deep.
- **The 25 %-of-screen floating card at full height.** It works in marketing; in a real Excel window at 1366 px it eats the grid. A docked, resizable ~380–420 px pane is kinder.
- **No visible workbook-context indicator.** "Which file am I talking to?" is a real question when three workbooks are open with three panes ([performance](https://claude.com/docs/office-agents/performance)). One filename line in the header is worth the pixels.
- **hCaptcha and sign-in inside the pane.** Local-first v1 should never block the transcript on a web round trip.

---

## Sources

**Product docs (current, snapshot 2026-10-31)**
1. https://claude.com/docs/office-agents/excel — capabilities, citations, instructions, models, overwrite protection, auto-compaction, limitations.
2. https://claude.com/docs/office-agents/overview — M365 overview, list of surfaces.
3. https://claude.com/docs/office-agents/work-across-apps — cross-app indicators, open-file constraint, settings toggle.
4. https://claude.com/docs/office-agents/dictation — mic, "Listening…", Enter-to-send, desktop/auth constraints.
5. https://claude.com/docs/office-agents/connectors-and-skills — `+` → Connectors, `/` → Skills.
6. https://claude.com/docs/office-agents/performance — 2,000-cell reads, 50-batch queue, 90 s / 5 min / 2 min waits, timeout UX, multi-pane model.
7. https://claude.com/docs/office-agents/data-storage — IndexedDB stores, keying, 50-conversation cap, Delete all, `claude.fileId`, `AutoShowTaskpaneWithDocument`, `Claude Log`, status-line helper call, inflight/approval records.
8. https://claude.com/docs/office-agents/third-party-platforms — connection-failed screen, "Copy error details", domains, network allowlist.
9. https://claude.com/docs/office-agents/opentelemetry — `agent.stream` (`first_token`, `stream_complete`, `APIUserAbortError`), `agent.compaction`, `agent.tool_execution`.

**Anthropic blog / newsroom**
10. https://claude.com/blog/collaborate-with-claude-across-excel-powerpoint-word-and-outlook — GA 2026-05-07; "conversations persist with each file"; one conversation across four apps.
11. https://www.anthropic.com/news/advancing-claude-for-financial-services — beta 2025-10-27; sidebar framing, "tracks and explains its changes", "navigate directly to the cells". Also the source page for the product screenshot asset below.
12. Product screenshot (Anthropic asset, 2560×1440): https://cdn.sanity.io/images/4zrzovbb/website/450687b3c7a1189eb1a650ade77588c56ca01461-2560x1440.png — pane chrome, citation chips, "Accept all edits", composer, palette samples.
13. Official video (Claude channel, 87 s): https://www.youtube.com/watch?v=F6dzjaBCBtU — embedded in the GA post; "Claude now works across your open Outlook, Word, Excel, and PowerPoint files in a single continuous session".

**Help center**
14. https://support.claude.com/en/articles/12138966-release-notes — 2025-11-24 beta to all paid plans + `ctrl+option+c`; 2026-02-05 Opus 4.6 + native Excel ops; 2026-03-11 cross-app sharing.
15. Archived pre-GA article (2026-02-14 snapshot): https://web.archive.org/web/20260214042810/https://support.claude.com/en/articles/12650343-using-claude-in-excel — change highlighting + explanatory comments, per-tool confirmation pop-up and the dangerous-function list, model switch between Sonnet 4.5 / Opus 4.6, .xlsx/.xlsm, auto-compaction, session logging, overwrite protection, Ctrl+Z, "chat history is not saved between sessions" (historical).
16. https://support.claude.com/en/articles/11647753-understanding-usage-and-length-limits — linked by the docs as the limits reference.

**Claude Academy (Anthropic)**
17. https://academy.claude.com/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning — "appears as a sidebar that can read your entire workbook"; reads before answering; clickable options ("0, 1, or a custom value"); before/after comparison table; diagram; Ctrl+Option+C / Ctrl+Alt+C; Cmd+Z/Ctrl+Z; "formulas stay formulas".
18. Academy OG illustration: https://academy.claude.com/og/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning.png — starburst + wordmark header, composer with round orange send button.

**Third-party, used only to locate official material (no claims rest on these)**
19. https://gigazine.net/gsc_news/en/20260312-claude-excel-powerpoint-share-context/ — pointer to the official cross-app video.
20. https://note.com/samurai_navi/n/ndeecc8c3c6f4 — Japanese practitioner note that first suggested looking for a Chat History control at the pane's top-right; the control's placement remains **[I]** here.
