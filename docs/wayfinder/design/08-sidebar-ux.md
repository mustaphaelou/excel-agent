# Sidebar UX and visual language

Ticket: [#8 Sidebar UX and visual language](https://github.com/mustaphaelou/excel-agent/issues/8)
Status: decided 2026-11-07 · amended 2026-11-08 after the #13 and #16 prototypes
Companion: [annotated wireframe sheet](08-wireframes.html) (open it in a browser; frames are 1:1 CSS px)
Depends on: [#2](https://github.com/mustaphaelou/excel-agent/issues/2) (capability envelope), [#5](https://github.com/mustaphaelou/excel-agent/issues/5) (Claude teardown), [#6](https://github.com/mustaphaelou/excel-agent/issues/6) (sideload spike), [#7](https://github.com/mustaphaelou/excel-agent/issues/7) (runtime, [ADR 0001](../../adr/0001-runtime-topology.md))

This record settles the pane's information architecture, visual language, components, states and keyboard behaviour. It is written to be implemented by the prototype ticket [#13](https://github.com/mustaphaelou/excel-agent/issues/13): every measurement, token, state and piece of copy below is meant to be used, not interpreted. The visual direction was validated by the prototype on 2026-11-08; see the amendment at the end of the Answer where it supersedes the original measurements.

---

## Answer

**Direction: a warm paper conversation column inside a Fluent pane.** Fluent supplies the geometry, control sizing, focus behaviour and host integration; Claude supplies the reading surface, the serif run-ins, the one-accent discipline and the inline evidence chips. The result is neither a Claude reskin nor an Excel command panel.

> **Prototype verdict — 2026-11-08.** The clickable prototype ([#13](https://github.com/mustaphaelou/excel-agent/issues/13)) was reviewed by the dev, who selected **B — Claude-close**. This supersedes the shell, user-turn, empty-state, composer, edit-card and view-header layout in §§2–5 below; the safety, state, history, copy, keyboard and accessibility contracts stand. The full amendment is at the end of this Answer.


1. **The pane is a reading column, not a second spreadsheet.** One column, max 720 px of content, 16 px gutters at the standard width, 48 characters per line at 380 px. The grid is the canvas; citations and edit range chips are the only bridge to it.
2. **Warm paper, one accent.** Surface `#FAF9F5` / ink `#1B1A18`; cool blue is reserved for evidence; terracotta is reserved for action. Excel green is not used inside the pane — it would read as a host command and collapse the evidence/action distinction.
3. **Default to a 380 px webview width and remember it.** Excel's default pane is 320 px. The pane targets 380 px of webview content (≈404 px outer; calibrate the chrome offset), stores the user's width in `localStorage`, and re-applies it on every open, because the spike proved width does not survive close→reopen. Width is a machine preference, never workbook state.
4. **One header line, not a banner.** 44 px: workbook name (truncating), unsaved dot, then history / new chat / settings at 32 px targets. New chat appears only once there is a conversation; the workbook menu is always available.
5. **The conversation is plain, the wait is narrated.** User messages are quiet tinted blocks; assistant answers are bubble-free markdown. A tool-derived status line narrates the wait in domain language ("Reading Forecast!A1:H400 · step 2"); the total is unknown until the turn ends, so #12 amended the "of 3" example ([#12 §15](12-agent-protocol.md)). No helper model call is spent on it.
6. **Citations are inline chips, with no sources footer.** 20 px pills, mono, cool fill, blue text at 5.32:1. Only structured citations become chips; cell-shaped prose never auto-links.
7. **The edit card is the review surface.** Inline after the proposing turn, expanded by default, grouped by sheet, one row per change with a per-row checkbox, a before→after diff, and risk notes attached to the row they belong to. Apply is one gate per turn; partial failure is a state of the same card.
8. **The composer is the only input surface.** Model picker and send/stop live inside it. While changes are pending the composer stays visible but send is disabled with a stated reason; it is never hidden.
9. **History, instructions and settings are full-height views**, not popovers. History defaults to this workbook; instructions are per workbook and say so; model, theme, startup and history live in Settings.
10. **No global keyboard shortcut in v1.** Office add-in shortcuts require SharedRuntime 1.1, and ADR 0001 chose the default task-pane runtime. Ribbon + Home > Add-ins + opt-in auto-open are the entry points. This is the one place where matching Claude costs an architectural reversal; it is called out for the dev rather than smuggled in.
11. **Auto-open is opt-in and discloses the write.** Settings → Startup sets `Office.AutoShowTaskpaneWithDocument` for this workbook (`settings.set` + `saveAsync`). Off by default, reversible, disabled for unsaved workbooks. The pane remains closable; `locked` is ignored by Excel and is not a pinning substitute.
12. **The design is a token system, not adjectives.** The light and dark palettes, type ramp, 4 px spacing, radii, motion and focus rules in §3 are the source; components never invent a colour or a size.


### Prototype verdict — 2026-11-08

The clickable prototype ([#13](https://github.com/mustaphaelou/excel-agent/issues/13)) was reviewed by the dev on 2026-11-08. The validated visual direction is **B — Claude-close**: the warm paper reading surface and Fluent host constraints, with the quieter shell, right-aligned user turns and floating composer trialled in the prototype. This section supersedes the original measurements below where they conflict. The prototype is throwaway; this is the production-facing decision. Source: [`prototype/clickable-sidebar-13`](https://github.com/mustaphaelou/excel-agent/tree/prototype/clickable-sidebar-13/docs/wayfinder/design/13-prototype).

| Element | Validated B — Claude-close | Supersedes |
| --- | --- | --- |
| Shell / header | A **46 px whisper topbar**, no permanent hairline: workbook name (11.5/16, truncating), muted file icon, unsaved dot and chevron; history / new chat / settings as **32 px hit targets** drawn as 30 px circles; 8 px side inset. New chat still appears only once there is a conversation. The workbook menu still opens from the name. | The 44 px bordered header and its 32 px square icon buttons. |
| Transcript / user turn | Assistant stays bubble-free. User prompts are a **right-aligned bubble**, max 88% width, 8/12 px padding, 13.5/20, radius `16px 16px 4px 16px`. Turn gap 18 px. | The full-width tinted user block; the left-aligned reading column. |
| Empty state | **Centered** composition: 38 px brand mark, serif question at 27/33, description 12.5/19, suggestion rows at 12 px radius, trust line centered. | The left-aligned empty state and 8 px suggestion radius. |
| Composer | A **floating card**: 8 px side inset, 4 px inner inset, 18 px radius, 1 px `--line` border, soft elevation; 11/12/8/14 px padding; 14/21 input; model chip without the grid glyph; **32 px hit target** send/stop (30 px visual circle). Gated, streaming, read-only and stop behavior are unchanged. | The 10 px composer card with its full-width bottom bar. |
| Edit card | **14 px radius, no outer border**, soft elevation; 14/13 px header; footer on `--surface`; same expanded-by-default body. The diff, risk notes, per-row inclusion, overwrite acknowledgement, one gate, apply/partial/retry/dismiss semantics are unchanged. | The 10 px bordered card and its footer treatment. |
| History / settings | Full-height views stay. 8 px outer / 4 px inner inset; **borderless view header**; serif view title at 16/21; same scope control, rows, footer and settings sections. | The bordered 44 px view header and flush body. |
| Shape / depth | Controls/checkboxes keep 4/6 px radii; **12 px** suggestions/menus, **14 px** cards, **16/4 px** user bubble, **18 px** composer. Hairline first, soft elevation second; no acrylic. | The 4/6/10 px-only radius rule in §3.4. |
| Action / brand | Claude's `#D97757` is the **brand mark and illustration only**. Filled actions keep #8's `#B55434` with white text (**4.90:1**); citation, warning, danger and success pairs are unchanged. Dark mode keeps the lightened action token with ink text. | Using the brand orange as a filled action colour. |

**What does not change.** One pane per workbook; 380 px default with a stored width; 260–720 px reflow; one column; ribbon + Home > Add-ins + opt-in auto-open; no global shortcut; tool-derived status; inline citation chips with no sources footer; per-workbook history/instructions; full-height History/Settings; no `Claude Log`; one approval gate per turn; per-row inclusion and overwrite acknowledgement; one `Excel.run` / undo group; partial failure with scoped Retry and Dismiss; the complete state matrix, copy deck, keyboard model, focus return, reduced-motion, forced-colors, 200 % zoom and contrast floors. The composer remains visible and explains its gate; the edit card remains the review surface, not a table.

**Implementation.** Production remains React 19 + TypeScript + Vite with CSS custom properties and bundled Source Serif 4; the prototype is not a code source. The two source snapshots for implementation review are the [B review frame](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/clickable-sidebar-13/docs/wayfinder/design/13-prototype/screenshots/claude-review.png) and the [B first-run frame](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/clickable-sidebar-13/docs/wayfinder/design/13-prototype/screenshots/claude-empty.png). Re-run the prototype at 260/380/520/720 in light and dark before changing component geometry.

### Safety and persistence amendment — 2026-11-08

The #16 extension of the throwaway prototype ([`prototype/safety-persistence-16`](https://github.com/mustaphaelou/excel-agent/tree/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype)) exercised the six surfaces #11 assigned to a prototype. This amendment is the production-facing decision; no new visual direction was introduced. All surfaces use B's shell, composer and full-height view anatomy. Screenshots: [link picker](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-link.png), [unsaved](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-unsaved.png), [compaction notice](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-compaction-notice.png), [divider](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-compaction-divider.png), [retention 80%](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-retention-80.png), [over budget](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-retention-100.png), [store failure](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-store-failure.png), [instructions copy](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-instructions-copy.png), [read-only chat](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-readonly-chat.png), [forced-on Settings](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-readonly-forced.png). Lighthouse accessibility is 100 on the link picker, retention and forced read-only samples.

**What was settled.**

1. **Link picker is a full-height view, reached from History's empty state and the workbook menu.** It lists every conversation from other workbooks with its full path and date, grouped by workbook, and filters on path or conversation name. Each row has one explicit "Move to this workbook" action. Moving is one-way; the original key is recorded on the conversation and shown in the row's overflow as "Move back to {workbook}", so the undo survives a pane reload as well as the toast's Undo. The current workbook's conversations are never listed. An unsaved workbook has no key, so History says "History starts when you save this workbook" and the picker entry is disabled with "Save the workbook first"; no memory-only linking.
2. **Compaction lives above the composer, not in the header banner.** The 80% notice is conversation-scoped: "This chat is {pct}% of the model's context. Summarize earlier turns or start a new chat." The button opens a determinate wait ("Summarizing {n} turns…") and then inserts a divider at the cut. "View original" expands the pre-cut turns read-only (no Copy, Retry or Steps actions); the originals remain in the transcript. A second compaction cuts from its own boundary and stacks a second divider. Nothing auto-summarizes.
3. **Retention management is a Settings section, not a second History screen.** The global 80% / 100% banner links there. It shows the store total against the 250 MB soft budget, then one row per workbook with path, conversation count, size and Delete, plus Open folder and "Delete all history…". Per-conversation `.md` and `.json` export lives in the History row overflow. Delete-all confirms with the conversation count and size. Nothing is deleted automatically.
4. **Store failure degrades; it never blocks the turn.** A persistent warning banner under the header says "History isn't being saved. Your conversation is still open." with Retry. A corrupt file is reported in History as "1 conversation couldn't be read. It was set aside as {file}." with Open folder. A conversation restored from the last good copy carries a "Recovered" badge. Retry clears the banner and announces "History is being saved again."
5. **Instructions migration is a one-click replacement, never a merge.** The "Copy instructions from {workbook}" chip appears only while the field is empty and offers the most recently used other workbook that has instructions. The click replaces the empty field, the chip disappears and the normal Saved state runs. There is no append path.
6. **Read-only is a persisted per-workbook mode with one control.** The header lock chip is a **status indicator and Settings shortcut**, not a toggle; the switch in Settings → Editing is the only on/off control. The preference is stored as `readOnly` in the workbook's `workbook.json` (never in the workbook file) and changed through `PUT /api/history/workbooks/:key/read-only`. Effective read-only is `setting.readOnly || Workbook.readOnly`, re-evaluated when the workbook opens. When Excel reports the workbook read-only, the switch is shown on and disabled with "Excel opened this workbook read-only, so edits are unavailable."; the persisted preference is unchanged, so a later writable copy opens in the user's chosen mode. Turning the mode on is disabled while a proposal is unresolved ("Apply or discard the pending changes before switching modes."). Chat and reads stay enabled; the composer hint says "Read-only mode. Edits are unavailable; questions still work." and write tools return the #9 `ReadOnly` error without staging a card.

**Component inventory amendments.**

| ID | Change |
| --- | --- |
| C4b | Header chip is now specified as **status + Settings entry**, not a toggle. 22 px pill, lock icon, "Read-only"; hidden when off. Tooltip: user mode "Read-only mode is on. Questions only. Open Settings to change it."; forced mode "This workbook is open read-only. Questions only." Clicking opens Settings → Editing and focuses the switch. At ≤299 px the label collapses to the lock icon with the same accessible name. |
| C6 | Global banner now carries two additional persistent statuses: **history budget** (warn at 80%, danger at 100%, action "Manage history") and **store failure** (warn, message + Retry). Both sit under the header and above the view body; neither overlays. |
| C25 / C27 | Read-only is an explicit composer state: enabled for questions, placeholder "Read-only mode — questions only", muted lock hint. It does not disable send; it gates writes. |
| C39 | History gets a link entry in the empty state and on the workbook menu; the empty state has a distinct unsaved variant. Scope and grouping are unchanged. |
| C41 | Row badges now cover **Pending**, **Recovered** and **Moved here**. A moved row's overflow gains "Move back to {workbook}". |
| C42 | Row overflow adds **Export .md** and **Export .json** beside Rename/Delete. |
| C43 | Footer unchanged: "Stored on this machine" + Delete all…; the link picker and export actions live in the body/row menu. |
| C44 / C49 | Settings gains an **Editing** section and a full **History** management section (total, per-workbook rows, Open folder, Delete all). Section order: Instructions, Model, Appearance, Startup, Editing, History, About. |
| C45 | Instructions field gains the **copy-from-workbook chip**, visible only when the field is empty. |
| C51 | **Link picker view** — full height, back to History (or chat when opened from the menu), intro line, search field, grouped rows. States: list, filtered, no match, no other conversations, unsaved. |
| C52 | **Link row** — title, full path, date, message count/sheets, one "Move to this workbook" button. No multi-select. |
| C53 | **Compaction notice** — above the composer, spark glyph, percentage line, "Summarize earlier turns" + "Start a new chat"; summarizing variant shows "Summarizing {n} turns…" and a progress bar. |
| C54 | **Compaction divider** — rule, "Earlier {n} turns summarized", "View original"/"Hide original"; expanded originals render read-only. |
| C55 | **Manage-history section** — total size / limit, budget track, per-workbook rows with path/count/size/Delete, Open folder, Delete all. |
| C56 | **History store notice** — danger inset with the corrupt-file copy and Open folder. |
| C57 | **Read-only switch row** — label, description, "Saved with this workbook's history; never written into the file." footnote, 34 × 18 switch; disabled for forced-on and pending proposals. |
| C58 | **Instructions copy chip** — pill with copy glyph and "Copy instructions from {workbook}". |

**State matrix additions.**

| State | Header | Banner / notice | Transcript | Composer |
| --- | --- | --- | --- | --- |
| Link picker | back · "Link to an old conversation" | — | old conversations grouped by path; filtered/no-match/empty states | n/a |
| Unsaved workbook | workbook name "Book1", unsaved dot | — | History empty: "History starts when you save this workbook"; picker disabled | normal |
| Compaction at 80% | workbook | notice above composer | full transcript; no divider yet | normal |
| Summarizing | workbook | "Summarizing {n} turns…" + progress | unchanged | normal |
| Compacted | workbook | — | divider at the cut; originals collapsed | normal |
| View original | workbook | — | pre-cut turns read-only under the divider | normal |
| Second compaction | workbook | notice above composer | two stacked dividers | normal |
| History at 80% | Settings | budget warn banner | n/a | n/a |
| History over budget | Settings | budget danger banner | n/a | n/a |
| Store failing | History or chat | persistent "History isn't being saved" banner + Retry | normal; recovered row in History | normal |
| Read-only (user) | lock chip | — | normal | enabled; hint; writes refused |
| Read-only (forced) | lock chip | — | normal | enabled; hint; switch disabled in Settings |
| Read-only + pending proposal | lock chip | — | proposal card | gated by the proposal; mode switch disabled |

**Copy additions (exact).**

| Context | Copy |
| --- | --- |
| Link picker title | Link to an old conversation |
| Link picker intro | Moves a conversation from another workbook. It is not copied, and you can move it back later. |
| Link picker search | Filter by path or name |
| Link picker empty | No other conversations yet |
| Link picker no match | No conversations match |
| Link action | Move to this workbook |
| Linked toast | Moved to this workbook. |
| Link-back action | Move back to {workbook} |
| Unsaved history | History starts when you save this workbook |
| Unsaved link | Save the workbook first |
| Compaction notice | This chat is {pct}% of the model's context. Summarize earlier turns or start a new chat. |
| Summarizing | Summarizing {n} turns… |
| Compaction divider | Earlier {n} turns summarized · View original |
| Compaction original toggle | Hide original |
| Budget banner, 80% | History is {pct}% of its {size} budget. Manage history |
| Budget banner, 100% | History is over its {size} budget. Nothing is deleted automatically. Manage history |
| Manage history total | {used} MB of {limit} MB |
| Manage history row | {n} conversations · {size} MB |
| Store failure | History isn't being saved. Your conversation is still open. |
| Store retry success | History is being saved again. |
| Corrupt file | 1 conversation couldn't be read. It was set aside as {file}. |
| Recovered badge | Recovered |
| Instructions chip | Copy instructions from {workbook} |
| Instructions copied | Copied instructions from {workbook}. |
| Read-only chip | Read-only |
| Read-only user tooltip | Read-only mode is on. Questions only. Open Settings to change it. |
| Read-only forced tooltip | This workbook is open read-only. Questions only. |
| Read-only settings label | Read-only mode |
| Read-only settings description | No edits or proposals. Reads and questions still work. |
| Read-only settings footnote | Saved with this workbook's history; never written into the file. |
| Read-only forced description | Excel opened this workbook read-only, so edits are unavailable. |
| Read-only forced footnote | Read-only mode is forced on by Excel and can't be turned off here. |
| Read-only pending | Apply or discard the pending changes before switching modes. |
| Read-only composer placeholder | Read-only mode — questions only |
| Read-only composer hint | Read-only mode. Edits are unavailable; questions still work. |
| Read-only refusal | I can't stage an edit while read-only mode is on. Turn it off in Settings, then ask again — reads and questions still work. |

**What does not change.** The B shell, composer, full-height History/Settings, one pane per workbook, no workbook identity marker, no auto-delete, no auto-compaction, no append-only instructions, and the #9 gate. #16 changes surfaces and adds the read-only persistence field; it does not reopen the visual direction.

**The one correction to the ticket body.** The ticket says width and dock state persist in `taskpanes.xml`. The [#6 spike](../../wayfinder/research/06-pane-sideload-spike.md) measured the opposite for this sideload: a saved `width="600"` reopened at the 326 px default, `dockstate` was empty, and `locked` was ignored. The design therefore treats width as a preference that must be re-applied, not state that can be relied on. `Office.extensionLifeCycle.taskpane.setWidth()` is the only width API; there is no `getWidth`.

---

## 1. Job and shape

**The pane's job.** Answer questions about the open workbook, explain what it finds, and propose edits that the user reviews before they land. It is not a chat app with an Excel theme; it is a review surface that happens to converse.

**One pane per workbook.** The pane is document-scoped: each open workbook gets its own instance and its own session, and there is no cross-workbook sidebar. Panes do not share conversation or context. The history view defaults to this workbook; the header names it. This is a hard Office.js constraint, not a preference.

**The three distances the layout has to keep:**

| Distance | What it means | Design consequence |
| --- | --- | --- |
| Pane ↔ grid | The workbook is the source of truth; the pane is a commentary on it | Cite ranges, navigate on click, never mirror the sheet |
| Answer ↔ evidence | A claim about a cell should be one click from that cell | Inline citation chips; no sources footer |
| Proposal ↔ write | The user must see the exact before/after before anything lands | Diff-first edit card; one approval gate per turn |

**What is deliberately absent:** document tabs, sheet navigator, formula bar, minimap, file attachment strip, account menu, sign-in, dictation, thumbs, connectors, slash commands, and any persistent workbook-context banner beyond the 44 px header line. The pane is small; every persistent element has to earn its pixels.

---

## 2. Layout

### 2.1 Webview widths

"Width" below means the webview's CSS width (`window.innerWidth`), which is what the frames and the composition use. The task pane's outer width is ~24 px larger; `setWidth` takes the outer width. Measure the offset once, then convert. The spike saw ~809 ms between the call and the new width.

| Range | Name | Layout |
| --- | --- | --- |
| 86–239 px | Emergency | Header collapses to file icon + one overflow menu; 8 px gutters; body 12/19; model label hidden; suggestion cards become one line; a single quiet line says "Drag the pane edge to widen." |
| 240–299 px | Compact | Header drops new chat (it stays in History); file name truncates before icons; 12 px gutters; table scrolls; model label may truncate; composer padding 8 px. |
| 300–519 px | Standard | The design target. 380 px webview / 348 px text column / ≈404 px outer; 16 px gutters; full header; 14/22 prose. |
| 520–719 px | Wide | 24 px gutters; content column max 720 px, centred; composer max 720 px; no second column. |
| ≥720 px | Wide | Same as wide; the pane never becomes a two-pane layout. A 720 px measure is the cap. |

- **Default nudge:** 380 px webview. The task pane API's Windows default is 320 px; at 320 the citation pill and the table's four columns start fighting. 380 webview is the smallest width where the standard frame is comfortable. Call `setWidth(380 + chromeOffset)`, not `setWidth(380)`.
- **The reflow is real.** `setWidth` is `void` and applies asynchronously (the spike saw ~809 ms). Do not animate the width and do not block first paint on it. Render the shell immediately at whatever width the host gives; one reflow after open is the cost of remembering width. The alternative — designing at the 320 default and dropping persistence — is worse for tables and citations.
- **Persistence:** store the *outer* width under `excel-agent.paneWidth`. On load, if TaskPaneApi 1.1 is present and a stored value exists, call `setWidth(stored)` once. On drag, debounce 300 ms and store `innerWidth + chromeOffset`, where `chromeOffset` is measured once from a known `setWidth` call (the spike measured 24 px; do not hard-code it if `outerWidth - innerWidth` is available). Clamp to `[86, 0.5 × client width]` (240 is the designed floor, not a hard clamp). The default target is an outer width that yields 380 webview.
- **Why re-apply every open:** the spike's close→reopen test reset 376 → 326. Without re-application the preference is a lie. Do not write width into `taskpanes.xml`; it is machine preference, not document content.
- **Do not fight the host:** no pinning (Outlook-only), no `locked` (ignored by this sideload), no custom close interception. Closing the pane ends the in-flight turn (default runtime). On reopen, restore the session and mark the interrupted assistant turn "Interrupted when the pane closed" with Retry.

### 2.2 Vertical layout

```
┌─────────────────────────────┐
│ host chrome (Excel's)       │  ~38–50 px, not ours
├─────────────────────────────┤
│ header                      │  44 px, sticky, hairline only when scrolled
├─────────────────────────────┤
│                             │
│ transcript / empty / view   │  flex 1, scrolls; 16 px gutters
│                             │
├─────────────────────────────┤
│ composer                    │  auto, ~86 px at rest
└─────────────────────────────┘
```

- The host draws the pane's title bar (the add-in name, not the workbook). That is why the in-pane 44 px line exists; it is not a duplicate.
- Header: 44 px, `padding-inline: 12px 8px`, 32 px icon buttons, 6 px gaps. The workbook button is the only flex item that shrinks.
- Transcript: `padding: 14px 16px 12px`, 20 px between turns, 10 px inside an assistant turn.
- Composer: `padding: 8px 12px 10px`; the composer card is 10 px radius, 1 px control border, 9/10/7/12 internal padding.
- At heights < 520 px, the empty state hides its trust line; at < 420 px it hides the suggestions and keeps the greeting + composer.
- The transcript scrolls under the header; the header only grows a hairline when `scrollTop > 0` (avoid a permanent rule).
- No sticky message headers, no floating "scroll to bottom" button until the user is more than one viewport away; then a 24 px pill above the composer.

### 2.3 The reading measure

| Webview width | Text column | Prose | Approx. characters |
| --- | --- | --- | --- |
| 260 px | 236 px | 13/20.5 | ~37 |
| 380 px | 348 px | 14/22 | ~48 |
| 520 px | 472 px | 14/22 | ~65 |
| ≥720 px | 720 px cap | 15/23 | ~85 |

At 380 px webview the measure is below the classic 65–75 character ideal. That is the correct trade: the pane is narrow, and a 380 px pane that refuses to use its width would be worse. The wide layout raises the cap but never exceeds 720 px.

### 2.4 Entry points and lifecycle

**Ribbon.** One custom tab, **Agent**, with one large button, **Open Agent** (`ShowTaskpane`). No secondary verbs in v1: "Explain selection" and "Summarise sheet" are prompts the empty state already offers once the pane is open, and duplicating them as ribbon actions would require either a shared runtime or a second task-pane URL whose state semantics are unproven. Home > Add-ins remains the fallback entry point.

- The button opens the pane; if it is already open, the host brings it forward. The button does not toggle it closed (closing is the host's chrome). Label it "Open Agent", never "Agent" alone — a verb is clearer in a ribbon of nouns.
- The manifest's `ShowTaskpane` action uses `<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>` so the same action can auto-open; no separate action is needed.
- **No global keyboard shortcut in v1.** Office add-in shortcuts require SharedRuntime 1.1 ([docs](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)), and ADR 0001 chose the default task-pane runtime. Adding Claude's `Ctrl+Alt+C` (or `Ctrl+Alt+A`) means amending #7/ADR 0001 and accepting that closing the pane no longer unloads the page; it is a lifecycle change, not a keybinding. In-pane shortcuts are in §7.
- **Auto-open is opt-in.** Settings → Startup sets `Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true)` and `saveAsync()` for this workbook. Off by default; reversible from the same switch or the workbook menu; disabled with "Save the workbook first" for an unsaved workbook. Only one pane can be auto-opened; the setting is ignored if the add-in is not installed ([docs](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document)).
- **Auto-open is not pinning.** `SupportsPinning` is Outlook-only; `locked` was ignored by this sideload; the user can always close the pane. The design does not pretend otherwise.

**Lifecycle.**

| Moment | Behaviour |
| --- | --- |
| Pane opens | `Office.onReady` → read theme → restore the most recent conversation for this workbook (or empty state) → restore width → focus the composer, or the card header if a proposal is pending. No workbook scan until the first prompt or an explicit context request. |
| Pane closes | Default runtime tears the page down; an in-flight turn stops. Persist the transcript and the last proposal state (#11). On reopen, an assistant turn that was mid-stream shows "Interrupted when the pane closed" + Retry. |
| Workbook closes | The pane closes with it; there are no workbook close/deactivate JS events. Same interrupted-turn rule. |
| User starts a new chat | If idle: new session, previous conversation stays in History. If streaming: confirm "Stop the current response and start a new chat?" first. If approval is pending: disabled with "Apply or discard first" — a new chat must not strand a pending proposal. |
| User switches conversations | Idle: allowed; the draft and scroll position are per conversation. Streaming: confirm stop. Pending approval: the proposal stays with its conversation; the History row shows a "Pending" badge; the new conversation's composer is not gated. |
| Citation click | Select + scroll the grid. The pane does not steal focus back. Hidden/protected sheet or cell-edit mode → snackbar with the reason and, where safe, an action. |
| Auto-open on file open | The pane opens docked right; the grid keeps focus; no modal, no toast. |

**Coexistence with the grid.** The pane is a right-docked column, not an overlay. It never covers the ribbon or the formula bar, never scrolls the grid on its own, and performs no grid mutations on load. Selection changes happen only on a citation or edit-range click. When the user is editing a cell, reads/writes that need Excel wait or fail with the cell-edit copy; the pane does not interrupt the edit.

---

## 3. Visual language

### 3.1 Direction and rationale

The reference is Claude's conversation layout, not Claude's chrome. The teardown found a three-role palette doing all the work: **warm paper for surface, cool blue for evidence, orange for action**. That discipline is the identity and it survives translation into Excel. Fluent supplies the parts a host integration needs: Segoe UI Variable, 32 px controls, a 4 px grid, focus rectangles, 4/6/10 px radii, and theme awareness.

What is **not** copied:

- Excel green as the accent. Green would make Apply look like a native ribbon command and would compete with the citation blue; the two-role distinction (evidence vs action) is more useful than host mimicry.
- A floating cream card at 25% of the window. This is a docked pane; 380 px is kinder at 1366 px than a 633 px card.
- Claude's low-contrast citation blue. The sampled `#698DC0` on `#EFF3F5` is 3.05:1. The role stays; the value becomes `#3F6599` (5.32:1).
- Serif for everything. The serif is reserved for assistant run-in headings and the empty-state question. UI and body stay sans.

### 3.2 Colour tokens

| Role | Light | Dark | Use | Contrast |
| --- | --- | --- | --- | --- |
| `--paper` | `#FAF9F5` | `#1B1B19` | Pane background, transcript | — |
| `--paper-2` | `#F2EEE5` | `#242420` | User message, recessed rows, tracks, hovers | ink 15.0:1 / 13.6:1 |
| `--surface` | `#FFFFFF` | `#22221F` | Composer, cards, menus | ink 17.4:1 / 13.9:1 |
| `--surface-2` | `#FBFAF7` | `#292824` | Row hover, disabled composer | — |
| `--ink` | `#1B1A18` | `#F2EFE7` | Primary text | 16.5:1 / 15.0:1 |
| `--ink-2` | `#4A4842` | `#CFCABE` | Secondary text, diff values | 8.7:1 / 10.6:1 |
| `--muted` | `#6B6862` | `#B8B3A8` | Metadata, helper text, placeholders | 5.27:1 / 8.26:1 |
| `--faint` | `#8A867C` | `#8E897E` | Icons, disabled text only — **never body copy** | 3.45:1 light (non-text ≥3); 4.95:1 dark |
| `--line` | `#E5E1D6` | `#35332D` | Decorative hairlines, dividers | — |
| `--line-strong` | `#D6D0C2` | `#46433B` | Card outline, table border | — |
| `--line-control` | `#8F897C` | `#7C776C` | Input/control boundary | 3.30:1 / 3.87:1 |
| `--accent` | `#B55434` | `#E08A6C` | Primary fill, action | white on light 4.90:1; ink on dark 6.59:1 |
| `--accent-hover` | `#9E4529` | `#EDA081` | Hover | 6.29:1 |
| `--accent-ink` | `#FFFFFF` | `#1B1B19` | Text/icon on accent | 4.90:1 / 6.59:1 |
| `--accent-soft` | `#F6E7DF` | `#3A2A23` | Selected chrome, focus glow | — |
| `--focus` | `#A94E2F` | `#F0B49B` | 2 px focus ring | 5.22:1 / 9.62:1 |
| `--cite-bg` | `#EEF3F6` | `#1F2C35` | Citation pill fill | — |
| `--cite-text` | `#3F6599` | `#A8C4E3` | Citation text | 5.32:1 / 7.94:1 |
| `--cite-border` | `#D9E4EB` | `#2C3B47` | Citation hairline | — |
| `--warn` / `--warn-bg` | `#7A4E00` / `#FFF4D6` | `#F0C674` / `#33280F` | Overwrite, busy, degraded | 6.57:1 / 8.98:1 |
| `--danger` / `--danger-bg` | `#A4262C` / `#FDE7E9` | `#FF9B9B` / `#3A1D1F` | Failed, destructive, risky | 6.15:1 / 7.57:1 |
| `--success` / `--success-bg` | `#0E6B0E` / `#DFF6DD` | `#7BD17B` / `#17301A` | Applied, saved | 5.87:1 / 7.63:1 |

**Rules.**

- Small text (≤14 px) must clear 4.5:1. `--faint` is for icons, disabled text and 1 px non-text marks only. This is not a suggestion; the muted token is the floor for metadata.
- Terracotta is never small text on paper (2.96:1). It is a fill, a 3 px accent bar, or an icon at ≥3:1. The focus ring is the darker `#A94E2F`, not the brand orange.
- Color is never the only carrier: overwrite/failure also carry words and icons; active history rows carry an accent bar *and* a tint.
- The background is not tinted by Office theme colours. Only light/dark is read from the host; the warm ramp is the product's.

### 3.3 Type

| Token | Family | Size / line | Weight | Use |
| --- | --- | --- | --- | --- |
| `display` | Source Serif 4 | 22 / 27 | 600 | Empty-state question, destructive-dialog title |
| `run-in` | Source Serif 4 | 16.5 / 21 | 600 | Assistant h2/h3 only |
| `prose` | Segoe UI Variable Text | 14 / 22 | 400 | Assistant body, list items |
| `ui` | Segoe UI Variable Text | 12.5 / 17 | 600 | Buttons, row labels, card titles |
| `small` | Segoe UI Variable Text | 10.5 / 15 | 400/550 | Metadata, helpers, status |
| `micro` | Segoe UI Variable Text | 9.5 / 13 | 700 caps 0.09em | Section labels, group heads |
| `mono` | Cascadia Mono / Consolas | 11.5 / 16 | 500 | Citations, formulas, diffs, error ids |

- **Families.** Body/UI: `"Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif`. Display/run-in: `"Source Serif 4", Cambria, Georgia, serif`, bundled (OFL) so the serif is identical on every machine. Mono: `"Cascadia Mono", Consolas, ui-monospace, monospace`; Cascadia ships with Windows Terminal, Consolas is the guaranteed fallback.
- **Serif scope.** Assistant `h2`/`h3` and the empty-state question only. Never UI labels, never table headers, never buttons. A serif button reads as a link; a serif table header reads as a quotation.
- **Mono scope.** Cell/range references, formulas, diff values, error ids. Numbers in tables use tabular figures (`font-variant-numeric: tabular-nums`) but remain sans. Inline code is mono with a `--paper-2` chip; it is not a citation.
- **No display font loading from a CDN.** Source Serif 4 is subset to Latin and shipped with the pane (~52 KB woff2). The body uses the system Segoe, so there is no FOUT on first paint.
- **Letter spacing.** Display `-0.022em`; run-in `-0.012em`; micro caps `+0.09em`; everything else 0. Body `-0.002em` is acceptable but not required.

### 3.4 Space, shape, depth, motion

| Token | Value | Notes |
| --- | --- | --- |
| Space | 2, 4, 8, 12, 16, 24, 32 | 4 px grid; 6 px gaps are allowed only inside a control row |
| Radius | 4 checkbox/range · 6 button/input · 12 suggestion/menu · 14 card · 16/4 user bubble · 18 composer · pill | Amended by #13. No larger radii inside the pane |
| Border | 1 px | Hairline for grouping; `--line-control` for anything clickable |
| Elevation 1 | `0 1px 2px rgba(27,26,24,.06), 0 6px 18px rgba(27,26,24,.06)` | Composer, edit card |
| Elevation 2 | `0 10px 30px rgba(27,26,24,.18), 0 2px 6px rgba(27,26,24,.08)` | Menus, dialogs |
| Motion | 120 ms hover · 180 ms disclosure · 240 ms view | `cubic-bezier(.2,0,0,1)`; no spring, no bounce |
| Reduced motion | opacity/color only; no height/travel/pulse | `prefers-reduced-motion: reduce` |
| Focus | 2 px `--focus`, 2 px offset | On every interactive element, including chips and rows |
| Tap/click | 32 px minimum | 44 px only for the primary action on touch (out of scope) |

- Depth is a hairline first, shadow second. No acrylic/mica: a translucent surface over a spreadsheet makes cell text harder to read and fights dark mode.
- Hover is background only (120 ms); no lift, no scale. The send button may scale to 0.96 on press, nothing else.
- The carat blinks 1.1 s steps(2); reduced motion makes it solid. The status dot pulses 1.4 s; reduced motion makes it static.
- View changes (chat ↔ history/settings) fade + translate 4 px over 120 ms; reduced motion cross-fades only. The pane must not feel like a page navigation inside a spreadsheet.

### 3.5 Icons

- 16 px at 1.5 px stroke, round caps/joins, no fills (the stop square is the one filled glyph).
- 32 px hit target; icon-only buttons always have `aria-label` and a tooltip after 400 ms.
- Icon set: chevron, clock (history), plus (new chat), gear, arrow-up (send), square (stop), copy, retry, alert-triangle, check, file, arrow-left, x, lock, grid (brand mark), pen, info, shield.
- The mark is a 2×2 grid in a rounded tile, not a starburst. It signals "cells", not "Claude".

### 3.6 Dark mode and host theme

- **Detection.** Prefer `matchMedia('(prefers-color-scheme: dark)')` for live changes (the spike's dark pane followed Excel); read `Office.context.officeTheme.isDarkTheme` at boot as a fallback. `OfficeThemeChanged` is documented for Outlook's mailbox object, not Excel's task pane, so do not rely on it. A manual override (Auto/Light/Dark) lives in Settings.
- **Inversion.** Warm charcoal, not blue-grey: paper `#1B1B19`, raised `#22221F`, ink `#F2EFE7`. The accent lightens to `#E08A6C` and its label flips to ink; citations lighten to `#A8C4E3`. Semantic backgrounds darken rather than invert to pastel.
- **Host chrome.** Excel's title bar around the pane is not ours. Do not try to match its exact grey; the warm surface is intentionally a different material, like paper on a desk.
- **First paint.** Set `data-theme` before React mounts from a tiny inline bootstrap that reads the media query + stored override; otherwise the pane flashes light on a dark host.

---

## 4. Information architecture

```
Pane (one per workbook)
├── Chat                         ← default
│   ├── Header: workbook · history · new chat · settings
│   ├── Transcript (or Empty)
│   │   ├── User turn
│   │   ├── Assistant turn
│   │   │   ├── status / step line
│   │   │   ├── markdown (prose, table, formula, links)
│   │   │   ├── inline citations
│   │   │   └── actions: copy · retry · steps
│   │   ├── (no decision card in v1; #12 asks in prose)
│   │   └── Edit proposal card (collapsed/expanded/applying/applied/failed)
│   ├── Approval bar (only when the card is scrolled out of view)
│   └── Composer: textarea · model · send/stop
├── History                      ← full-height view
│   ├── header: back · title · new chat
│   ├── scope: This workbook | All workbooks
│   ├── groups: Today / Yesterday / Earlier
│   ├── row: title · preview · time · message count · sheets
│   ├── overflow: rename · delete
│   └── footer: stored locally · delete all…
└── Settings                     ← full-height view
    ├── header: back · title
    ├── Instructions (per workbook)
    ├── Model + refresh
    ├── Appearance (Auto/Light/Dark)
    ├── Startup (auto-open, opt-in)
    ├── History summary + manage + delete all
    └── About: version · backend status · diagnostics
```

- The three surfaces are **views of one pane**, not routes. Back returns to Chat with the transcript scroll position preserved.
- There is no persistent navigation rail. The header's three icons are the entire navigation budget.
- The workbook menu (click the file name) contains: full path (copy), "Instructions for this workbook", "Open with this workbook" (same switch as Settings), and "Close pane" if the host exposes it. No file switcher; one pane is one workbook.
- The settings gear opens Settings; it is not a popover. At 380 px, a settings popover would be a scrolling card inside a card.

---

## 5. Component inventory

Each component lists what it is, its states, and the rules that are easy to get wrong. Measurements are CSS px at 380 px unless noted.

### 5.1 Shell and chrome

| ID | Component | Anatomy | States | Rules |
| --- | --- | --- | --- | --- |
| C1 | Pane shell | host chrome (Excel) + 44 px header + flex transcript + composer | boot, light, dark, narrow, wide | Host chrome is not ours. `color-scheme` follows theme. No custom title bar. |
| C2 | Header | workbook button (file icon, name, chevron), icon buttons | default, scrolled (hairline), no-conversation (no new chat), unsaved (dot), long name (ellipsis), streaming (new chat disabled with confirm), pending (new chat disabled) | 44 px; name is the only shrinking item; icons never shrink; 32 px targets; tooltip with full path. New chat is absent on first run. |
| C3 | Workbook menu | popover 232 px | closed/open | Shows full path with copy; workbook-scoped instructions; auto-open switch; no recent-file list. |
| C4 | Icon button | 32 × 32, 16 px icon | rest/hover/active/focus/disabled | `aria-label` always; active state for the current view; no icon-only destructive action without a confirm. |
| C4b | Read-only chip | 22 px pill, lock icon, "Read-only" | visible / hidden | Sits in the header when #9's read-only mode is on. Tooltip: "Questions only — no edits in this mode." Clicking opens Settings; it is a state, not a toggle in the header. |
| C5 | View header | back 32 px, title 13.5/18 650, optional action | History, Settings | Back returns to the previous scroll position; focus moves to the view title on open and back to the trigger on close. |
| C6 | Global banner | icon, message, one text action, 1 px bottom border | warning, danger, neutral | Under the header, above the transcript; never overlays; persists until the condition clears; `role="status"` (not `alert` unless it interrupts). |
| C7 | Toast/snackbar | 36 px, ink surface, one action | success, neutral, persistent-for-decision | Above the composer, 4 s for acknowledgements; a navigation decision (hidden sheet) gets 10 s and an explicit action. `role="status"`. Never for errors that need a decision beyond navigation. |
| C8 | Confirm dialog | `min(320px, calc(100vw - 24px))` max, title serif 18/24, body 13, Cancel + destructive | discard changes, delete conversation, delete all | Native `<dialog>` or focus-trapped primitive; initial focus on Cancel; Esc cancels; danger button uses `--danger` text on neutral, not a red fill, unless the action is irreversible. |
| C9 | Menu/popover | 232 px, 4 px padding, 32 px rows, check on selected | model, workbook, row overflow | Arrow keys + typeahead; Esc closes and returns focus; click-outside closes; no nested submenus. |
| C10 | Skeleton | 3–5 shimmer bars | boot, history loading, model loading | No spinner-only screens; the shell renders immediately. Reduced motion = static bars. |
| C11 | Tooltip | 12 px, ink surface, 240 ms delay | any icon-only control | Never the only source of a label; `aria-describedby` for the same text. |

### 5.2 Chat

| ID | Component | Anatomy | States | Rules |
| --- | --- | --- | --- | --- |
| C12 | Transcript | scroll container, 20 px turn gap | empty, populated, streaming, scrolled-up | A plain labelled `region`, not `role="log"` or `aria-live` — streaming tokens would be announced. Use one separate live region for turn milestones. `tabindex="0"` for keyboard scrolling. Auto-follow only when pinned to bottom. |
| C13 | User message | full-width tinted block, 10 radius, 10/12 padding, 13.5/20 | normal, long, code | No avatar/label. Align left; full width so long prompts use the pane. Markdown is not rendered in user messages (plain text, preserved newlines). |
| C14 | Assistant turn | status line, markdown, actions | streaming, complete, error, interrupted | No bubble. Max width 100%. The status line belongs to the turn, not the message. |
| C15 | Status/step line | 6 px dot, label, step count, chevron | live, done, stalled (>8 s), interrupted | Text comes from tool calls, not a helper model. Live: "Reading Forecast!A1:H400 · step 2" (the total is unknown until the turn ends; #12 §15). Done: "3 steps · 4.8 s" + chevron to the step list. Stalled: append "still working…" and an elapsed time; don't spin faster. |
| C16 | Markdown | headings, paragraphs, lists, tables, code, links, blockquotes | rendered, streaming (partial), empty | Sanitize; no raw HTML. Tables scroll horizontally; code blocks scroll; links get an external icon and `rel="noopener noreferrer"`. Partial markdown must never throw or show raw `**`. |
| C17 | Citation chip | pill, 20 px, 6 px padding, mono 11.5, 1 px border | default, hover, focus, pressed | Only explicit citations. Cross-sheet label always includes the sheet (`Summary!B12`); same-sheet may omit it (`B12`). Click selects + scrolls; Enter/Space too. Hidden sheet, protected sheet or cell-edit mode → snackbar with the reason and an action. |
| C18 | Table | 1 px border, 8 radius, scroll-x, 6/8 cells | short, wide, empty | `<table>` with `<th scope="col">`; numeric right-aligned, tabular; first column 550 weight; no zebra; scrollbar visible on hover/focus; never shrink columns below content. |
| C19 | Formula/code block | 1 px border, 8 radius, mono 11, scroll-x, copy button | default, copied | Formula text is not syntax-highlighted; copy copies the syntax verbatim, including `=`, so it can be pasted into Excel. A formula containing a range does not turn the range into a citation. |
| C20 | Message actions | text buttons 24 px: Copy, Retry, Steps | idle, hover, copied, disabled | Copy = markdown; Retry = last turn or the error card's turn; Steps = expand the tool trace. No thumbs in v1 (no training loop). |
| C21 | Decision card | question, 2–4 option buttons, selected state | not in v1 | #12 has no `ask_user` tool; v1 asks in prose. Reserved if a future tool adds a stream event and a persisted answer. |
| C22 | Empty state | mark, serif question, description, 3 suggestions, trust line | first run, returning, offline | Suggestions fill the composer; they never send. The third suggestion is selection-aware only when a range is selected. Trust line is one line, muted, with one small icon — no shield badge. |
| C23 | Suggestion row | 1 px border, 8 radius, 9/10 padding, 12.5 text | rest/hover/focus | Left-aligned; icon 13 px; arrow at the right; no more than 3. |
| C24 | Approval bar | 36 px, icon, "12 changes waiting", Review button | hidden, visible | Only when the edit card is off-screen; a sticky affordance, not a second decision surface. Clicking scrolls the card into view and focuses the card header (`tabindex="-1"`). |
| C25 | Composer | textarea, model chip, send/stop | empty, focused, typing, streaming, gated (approval), disabled (offline), read-only | Auto-grow 1→6 lines then scroll. Enter sends; Shift+Enter newline; IME-safe. The model chip is inside the composer, left. Send is a 32 px circle; Stop is the same circle in ink. While streaming, Enter does nothing and send is Stop; the draft is kept. Stop keeps the partial answer, marks it "Stopped", and does not roll back reads; writes never started without approval. |
| C26 | Model menu | 232 px, model name + check, "Refresh list", source line | loading, ready, stale, error, empty | No provider branding beyond the model name; the gateway is the source. Changing the model affects new messages only; the current turn keeps its model. On error the chip reads "Models unavailable · Retry" and keeps the last known model. |
| C27 | Composer hint | 10 px line under the composer | approval, offline, read-only | Explains why send is gated; uses `--muted`, not `--faint`. The hint is part of the composer surface, not a toast. |

### 5.3 Edit proposal

| ID | Component | Anatomy | States | Rules |
| --- | --- | --- | --- | --- |
| C28 | Edit card | header, body groups, footer | proposed (expanded), collapsed, applying, applied, partial failure, discarded, auto-applied | Inline after the proposing turn. Expanded by default. One card per turn. The header survives collapse with count + risk badges. |
| C29 | Card header | 26 px glyph, title, subtitle, risk badges, chevron | proposed, applied, partial | Title is the count ("Proposed changes" → "Applied 10 of 12"). Subtitle names ranges/sheets. Risk badges: warn "1 overwrite", danger "1 risky formula". |
| C30 | Change group | 22 px group head, sheet name, count | one or more sheets | Grouped by sheet, in workbook order; each group's changes in the order the model proposed them. |
| C31 | Change row | 16 px checkbox, 20 px range chip, operation, diff, note, status | included, excluded, risky, failed, applied | Checkbox defaults on. Range chip navigates on click. Diff: `was → now`, mono, wraps anywhere. Risk note sits under the diff. After apply: success check or failure text; checkbox disabled on applied rows. |
| C32 | Diff | inline mono line | value, formula, format, structural, native op | Values use the cell's displayed text (`Range.text`) so locale/format matches the sheet. Formulas use formula text. Format ops name the change ("General → $#,##0"). Structural ops use a sentence ("Insert sheet 'Summary' at position 3"). Native ops use a summary ("Sort A2:F400 by Revenue, descending"), not a fake cell diff. |
| C33 | Card footer | selection count, Discard ghost, Apply primary | ready, subset selected, applying, partial | Apply is the only primary; label includes the count ("Apply 12 changes" / "Apply 9 of 12"). Zero selected disables Apply. Discard opens a confirm if anything is selected. |
| C34 | Applying state | 2 px progress bar, per-row status, footer "Applying 3 of 12…" | applying | No cancel. Do not let the user edit the selection mid-apply. If the pane closes, the in-flight apply is gone; on reopen the card shows the last known state + "Interrupted". |
| C35 | Partial failure | danger inset alert, failed group, "Retry N failed", "Dismiss" | partial | Failed rows keep their checks and diffs. Retry is scoped to failures. Dismiss keeps them pending in the background; it never marks them applied. |
| C36 | Applied state | collapsed header "Applied 12 changes", Review, Ctrl+Z hint | applied | No celebratory confetti and no in-pane Undo. #9 provides no snapshot or programmatic undo; show the keyboard undo hint for all-undoable batches. |
| C37 | Large batch | inline card shows first 8 + "Review all 34 changes" | >20 changes | The full-height review view uses the same rows, grouped by sheet, with a sticky footer. The card summary keeps the risk badges. |
| C38 | Risk acknowledgement | card-level overwrite checkbox; row-level risk note; loss/confirm rows default excluded | overwrite, loss, confirm, blocked | #9 owns the policy. Apply stays disabled until the overwrite acknowledgement is ticked. Loss and confirm-tier rows start excluded and require ticking the row. Blocked rows are unselectable; protected/merged/pivot-output/spill-child reasons sit on the row. |

### 5.4 History and settings

| ID | Component | Anatomy | States | Rules |
| --- | --- | --- | --- | --- |
| C39 | History view | header, scope segmented, groups, rows, footer | empty, list, loading, deleting | Full height. The current conversation is active. Scope defaults to this workbook. |
| C40 | Scope segmented | two segments, 28 px | this workbook, all | "All workbooks" keeps the workbook name in each group; never merge files silently. |
| C41 | History row | title, time, preview, meta, optional Pending badge | active, hover, overflow open, pending | 56 px; title 12/16 600; preview one line; meta "6 messages · Forecast". Active = tint + 2 px accent bar + `aria-current="true"`. A conversation with an unresolved proposal carries a warn "Pending" badge. |
| C42 | Row overflow | 232 px menu: Rename, Delete | closed/open | Rename is inline/dialog; delete confirms; deleting the active conversation switches to a new chat. |
| C43 | History footer | lock icon, "Stored on this machine", "Delete all…" | default | Delete all confirms with the count; it does not delete workbook content, only local history. |
| C44 | Settings view | header + sections | default, saving, error | Full height; sections in the order: Instructions, Model, Appearance, Startup, History, About. |
| C45 | Instructions field | 3-row textarea, helper, saved state, counter | empty, editing, saving, saved, error, over limit | Per workbook; 8,000 char soft limit; save on blur + 500 ms debounce; "Saved" inline; no feature appends text. |
| C46 | Setting row | label, description, control, chevron | default, disabled | 38 px min; label 11.5/16 550; description 9.5/14 muted; the control is right-aligned. |
| C47 | Theme segmented | Auto / Light / Dark | selected | Auto follows the host; override stored locally. |
| C48 | Auto-open switch | 34 × 18 visual switch inside a 38 px row target, label, disclosure | off, on, disabled (unsaved) | Writes `Office.AutoShowTaskpaneWithDocument` into the workbook; says so; off by default; disabled with "Save the workbook first" when unsaved. |
| C49 | History summary row | "12 conversations · 1.4 MB", Manage | default | Links to History; the destructive "Delete all history…" sits below as a text row, not a filled red button. |
| C50 | About | version, backend status dot, copy diagnostics | connected, offline, error | Diagnostics copy includes version, host build, backend URL/health, last error id — never the gateway key. The copy is plain text suitable for a bug report. |

---

## 6. State model

### 6.1 Session states

| State | Header | Banner | Transcript | Composer |
| --- | --- | --- | --- | --- |
| Booting | skeleton | — | skeleton | skeleton |
| Ready / empty | workbook | — | empty state | focused, enabled |
| Streaming | workbook | — | user turn + assistant status + text + caret | editable; send = Stop; Enter does nothing |
| Tool running | workbook | — | status line ("Reading …"), step count | same; Stop |
| Awaiting approval | workbook | — | assistant turn + edit card expanded | gated; hint "Apply or discard…" |
| Applying | workbook | — | card progress per row | gated; hint |
| Applied | workbook | — | card collapsed "Applied N" | enabled |
| Stopped | workbook | — | partial answer + "Stopped" line | enabled |
| Partial failure | workbook | — | card partial; failed rows | gated until retry/dismiss |
| Turn error | workbook | health banner if global | error card | enabled if service healthy; retry |
| Offline / backend down | workbook | warning banner | transcript stays | disabled; "Reconnect to continue…" |
| Rate limited | workbook | warning banner with retry-at | transcript stays | disabled until the window resets; queue nothing |
| Gateway auth error | workbook | danger banner | transcript stays | disabled; "Update the key" + copy path |
| Excel busy (cell edit) | workbook | — | status "Waiting for Excel…" or failed row | enabled for chat; writes retry |
| Context limit | workbook | — | inline notice + "Start a new chat" | enabled (new chat) |
| Interrupted (pane closed) | workbook | — | assistant turn "Interrupted when the pane closed" + Retry | enabled |
| Read-only | lock chip | — | normal | enabled; edits refused with a reason |

### 6.2 Approval state machine (UI)

```
proposed ──discard──► discarded
   │
   ├──toggle rows──► proposed (subset)
   │
   └──apply──► applying ──success──► applied
                  │
                  └──failure──► partial ──retry──► applying
                                  │
                                  └──dismiss──► pending (card collapsed, failures retained)
```

- The composer is gated from `proposed` until `applied`, `discarded`, or `dismissed`.
- `partial` is not an error screen; it is the same card with an applied count, a danger inset, and the failed rows.
- `dismissed` keeps the failures pending in the background; the next turn may propose a fix, but it must not silently re-apply.
- Closing the pane during `applying` is not recoverable; on reopen the card says "Interrupted" and offers "Review" / "Retry" only if the state was persisted. #11 owns persistence.

### 6.3 Connection states

- **Service health** is separate from **turn outcome**. The banner is about the local service/gateway; the error card is about the turn.
- Health probe: `/api/health` on boot and after a failure; exponential retry (1, 2, 5, 10, 30 s), then every 60 s. Show a countdown only when `retry-after` is known.
- Offline: `navigator.onLine === false` is a hint, not a verdict; the local service can be reachable while the internet is not. The banner copy names which one failed.
- A 401/403 from the gateway is not "offline"; it gets its own copy and a "Copy config path" action.
- Timeouts: 90 s for Excel steps, 5 min for the model turn on desktop (per the Claude limits; our own #12 values may differ). Copy: "The model didn't answer in time. Your workbook is unchanged."

---

## 7. Keyboard and focus

| Key | Where | Action |
| --- | --- | --- |
| `Enter` | Composer | Send (ignored while IME composing) |
| `Shift+Enter` | Composer | Newline |
| `Esc` | Streaming | Stop the turn (announce "Stopped") |
| `Esc` | Menu/dialog | Close and return focus |
| `Esc` | Approval card | Collapse (does not discard) |
| `Tab` / `Shift+Tab` | Pane | Normal focus order; focus ring always visible |
| `Space` | Change row checkbox | Toggle include |
| `Space` / `Enter` | Citation chip | Select the range in the grid |
| `Enter` | History row | Open conversation |
| `↑`/`↓` | Model menu, history list | Move; `Home`/`End` jump |
| `PageUp`/`PageDown` | Transcript (focused) | Scroll one viewport |
| `Ctrl/Cmd+C` | Selection | Browser copy; copy buttons are additive |
| `Ctrl/Cmd+Z` | Grid after apply | Excel's undo; the pane does not intercept |

- **No global add-in shortcut in v1.** Custom keyboard shortcuts require SharedRuntime 1.1, and ADR 0001 chose the default task-pane runtime. The ribbon is the entry point. If the dev wants `Ctrl+Alt+A`, the correct path is to amend #7/ADR 0001 to shared runtime; it also changes the close-kills-turn behaviour. This is a real trade, not a missing detail.
- **Focus order:** header → transcript scroll region → first interactive element in the transcript → composer → model → send. The transcript scroll region is focusable so keyboard users can scroll it; interactive chips/rows are tab stops inside.
- **Focus management:** opening History/Settings moves focus to the view title (`tabindex="-1"`); closing returns it to the trigger. Opening the workbook/model menu focuses the first selected item; Esc returns. A citation click intentionally moves focus to the Excel grid; the pane does not steal it back.
- **Composer after actions:** after Apply/Discard, return focus to the composer. After Retry, focus stays on the Retry button until the new turn starts.
- **No focus traps** outside dialogs. No autofocus on streaming content. `aria-live` announcements never move focus.

---

## 8. Accessibility

- **Contrast:** body/small text ≥4.5:1; large text ≥3:1; control boundaries, icons, focus rings ≥3:1. The token table in §3.2 is the audit. `--faint` is not a text token.
- **Live regions:** one visually hidden `role="status"` `aria-live="polite"` region announces: "Response started", "Reading Sheet1 A1 to H400", "Step 2 of 3", "Response complete", "12 changes ready for review", "Applied 10 of 12; 2 failed", "Stopped", "Reconnected". Never announce token-by-token. The transcript itself is not an `aria-live` region.
- **Semantics:** header `banner`, view body `main`, composer `form` with a labelled textarea, transcript `region` labelled "Conversation", lists as lists, tables as tables with `scope`, dialogs as dialogs. Buttons are `<button>`; citations are buttons (they select, not navigate to a URL); external links are `<a>`.
- **Labels:** send "Send message"; stop "Stop generating"; citation "Go to Forecast, B22 to F25"; range chip "Select D4 to D40 in the grid"; include checkbox "Include change to B4: set formula"; apply "Apply 12 changes"; model "Model: gpt-5.2. Change model"; file "Workbook: Q3 Forecast.xlsx. Change".
- **Targets:** 32 × 32 minimum; 2 px focus ring; no keyboard trap; no hover-only actions. Row actions appear on focus as well as hover.
- **State without colour:** risk = icon + text; failure = icon + "Failed" + tint; active history = bar + tint + `aria-current`; selected segment = pill + weight.
- **Reduced motion:** status dot static, caret solid, no height animation; the state change still happens. No parallax/scale.
- **Forced colors:** use system colours (`Canvas`, `CanvasText`, `LinkText`, `Highlight`); do not force backgrounds; keep borders visible.
- **Zoom:** the layout is designed to survive 200% zoom (≈190 CSS px effective width) with the emergency header. Test 200% at 380 and 260.
- **Locale:** numbers/dates in diffs come from the cell's displayed text, not `Intl` on the raw value; relative times use `Intl.RelativeTimeFormat`; the UI language is `Office.context.displayLanguage` when localization lands. LTR only for v1; use logical properties (`padding-inline`, `inset-inline`) so RTL is not a rewrite.
- **Screen-reader order in the edit card:** header (count + risks) → group → each row (checkbox label includes the change) → footer. The diff is part of the checkbox's description, not a separate tab stop.
- **Tables:** if a table scrolls, it gets `tabindex="0"` and `role="region"` with a label ("Revenue by segment, scrollable"); a screen reader user can read the table without horizontal scrolling.

---

## 9. Copy deck

Voice: specific, calm, second person, sentence case. Name the object and the consequence. No "Oops", no exclamation marks, no anthropomorphism, no "I think". The model's prose is its own; this deck is the chrome.

| Context | Copy |
| --- | --- |
| Empty title | What are we working on? |
| Empty body | I can read this workbook, answer questions, and make changes — with your approval. |
| Suggestions | Summarise the Forecast sheet · Find formulas that need a second look · Clean up the selected table |
| Trust line | Stored locally · Nothing is written without approval |
| Composer placeholder | Ask about this workbook… |
| Follow-up placeholder | Ask a follow-up… |
| Gated placeholder | Review the proposed changes above |
| Gated hint | Apply or discard to keep chatting. |
| Offline placeholder | Reconnect to continue… |
| Read-only placeholder | Read-only mode — questions only |
| Status / tool | Reading {range}… · Checking formulas… · Writing {n} cells… · Waiting for Excel… |
| Status / no tool | Thinking… (only after 1.2 s with no tool) |
| Status / done | {n} steps · {seconds} s |
| Interrupted | Interrupted when the pane closed. |
| Proposal title | Proposed changes |
| Proposal subtitle | {n} changes in {m} ranges · {sheet} |
| Risk badges | 1 overwrite · 1 risky formula · Uses a volatile function |
| Diff overwrite | Overwrites a non-empty input cell |
| Diff risky | INDIRECT can reach outside this workbook |
| Apply | Apply {n} changes |
| Apply subset | Apply {n} of {m} |
| Discard | Discard |
| Discard confirm | Discard {n} proposed changes? Nothing has been written yet. |
| Applying | Applying {n} of {m}… |
| Applied | Applied {n} changes |
| Partial title | Applied {n} of {m} |
| Partial alert | Excel was still in cell-edit mode. {n} changes weren't applied. Finish the edit, then retry. |
| Retry failed | Retry {n} failed |
| Dismiss | Dismiss |
| Dismiss hint | The failed changes stay pending. |
| History title | History |
| History scope | This workbook · All workbooks |
| History empty | No conversations yet. Your chats are saved automatically. |
| History footer | Stored on this machine |
| Delete conversation | Delete this conversation? It can't be undone. |
| Delete all | Delete all {n} conversations? This can't be undone. |
| Settings / instructions | Instructions · Applies to every conversation in this workbook. |
| Instructions placeholder | Format numbers with thousands separators. Bold column headers. Use € for currency. |
| Instructions saved | Saved |
| Model | Model · Used for new messages. The current conversation keeps its model. |
| Model source | From your gateway. |
| Theme | Theme · Auto follows Excel's light or dark theme. |
| Auto-open | Open with this workbook · Adds a setting to {file}. Off by default; turn it off here. |
| Auto-open unsaved | Save the workbook first. |
| History summary | {n} conversations · {size} |
| About | Excel Agent {version} · {backend status} |
| Connection banner | No connection to the model gateway. Retrying in {n} s. |
| Backend banner | Can't reach the local service. Start Excel Agent, then retry. |
| Auth banner | The gateway key was rejected. Update config.env and restart the service. |
| Rate limit | Model rate limit reached. Try again in {time}. |
| Timeout | The model didn't answer in time. |
| Timeout body | Your workbook is unchanged. This is usually a gateway hiccup — try again. |
| Excel busy | Excel is busy with a cell edit. The change will apply when you finish. |
| Hidden sheet | That sheet is hidden. Show it and go there? |
| Protected sheet | Can't navigate: the sheet is protected. |
| Context limit | This chat is full. Start a new chat to keep going. Your history is saved. |
| Copy details | Copy details |
| Copied | Copied |
| Retry | Retry |
| Stop | Stop generating |

**Do not write:** "Something went wrong", "An error occurred", "Oops", "Uh oh", "Please try again later" (with no path), "The AI", "I'm sorry", "Great question", "As an AI". Prefer the concrete noun and the next action.

---

## 10. What we keep, adapt and reject from Claude

| Claude behaviour | Verdict | Our version |
| --- | --- | --- |
| Minimal chrome, one gear | **Keep** | One header line with workbook + three icons; gear opens a full Settings view. |
| No persistent workbook banner | **Adapt** | A quiet 44 px line with the file name; at 380 px the identity question is real with several panes open. |
| Inline citation chips, no sources footer | **Keep** | Same role, stronger contrast, structured citation tokens only. |
| Model-written status line | **Adapt** | Tool-derived status; no helper model call. Cheaper, honest, and it names the range. |
| Batch approval in the composer | **Adapt** | Inline edit card, expanded by default, with per-row inclusion. The card survives partial failure; the composer is gated but explains itself. |
| Highlight changed cells + comments | **Keep** as the post-apply second line of defence | #9 owns the exact mechanism; the card is the pre-apply surface. |
| No pre-apply diff | **Reject the omission** | Diff-first card; this is our differentiator. |
| Batch approve, chevron hides the list | **Adapt** | The list is primary; the header still collapses for scanning. |
| Ctrl+Z as undo | **Keep** | We do not intercept it. #9 rules out an in-pane Undo button and snapshot; Ctrl+Z is the only undo. |
| History app-scoped, 50 conversations | **Reject the scoping** | Per workbook first, all workbooks as an escape hatch; #11 owns storage and cap. |
| Instructions per app | **Adapt** | Per workbook, one textarea, examples; #11 confirms the scope. |
| Settings-only model picker | **Reject** | Model chip in the composer; Settings has the same control. |
| Auto-open pane with the document | **Keep, opt-in** | Settings switch, writes the document setting, off by default. |
| `Claude Log` sheet | **Reject** | No prose written into the workbook unless it is an approved edit. An export can come later. |
| Sign-in / captcha | **Reject** | Local-first; no auth gate inside the pane. |
| Dictation | **Defer** | Not v1; it forces an explicit "audio leaves the machine" disclosure. |
| Thumbs up/down | **Defer** | No training loop; it would be chrome. |
| Connectors / skills / slash menu | **Reject v1** | Out of scope; no `+`, no `/`. |
| 25% floating card | **Reject** | Docked, resizable, 380 px target. |
| Dark mode via host theme | **Keep** | Warm charcoal inversion, follows Excel; manual override. |

---

## 11. Implementation contract for #13

**Stack recommendation.** React 19 + TypeScript + Vite. No visual UI kit: the language is custom and Fluent's components would fight the warm ramp. Use unstyled accessible primitives for the hard parts (menu, dialog, tooltip, focus trap) — Radix or the platform's `<dialog>`/`popover` — and hand-roll the rest with the tokens above. Markdown via `marked` + `DOMPurify` (no raw HTML, no syntax highlighting). Icons as an inline SVG sprite (1.5 px stroke), not an icon package. Fonts: system Segoe + bundled Source Serif 4 (OFL subset, ~52 KB).

**Suggested source shape.**

```
app/
  src/
    shell/        Pane, Header, ViewHeader, Banner, Menu, Dialog
    chat/         Transcript, UserMessage, AssistantTurn, StatusLine, Markdown,
                  Citation, Table, Formula, MessageActions, EmptyState
    edits/        EditCard, ChangeRow, Diff, RiskNote, ApplyFooter, ReviewView
    history/      HistoryView, HistoryRow, ScopeSegmented
    settings/     SettingsView, Instructions, ModelRow, ThemeRow, StartupRow
    composer/     Composer, ModelChip, SendStop, GatedHint
    tokens.css    the §3 tables as custom properties
    a11y/         LiveRegion, useFocusReturn, useRovingTabIndex
    state/        session reducer, stream store, health probe
  public/fonts/source-serif-4-600-latin.woff2
```

**State.** One reducer per pane session with discriminated unions for `Turn`, `Proposal`, `Connection`, `ModelCatalog`. Streaming tokens go into a separate mutable store and are committed to the turn at ~10 fps; do not `setState` per token. Completed messages are memoized. The transcript is not virtualized in v1; if #11's history makes a conversation >100 messages, virtualize rows only after measuring.

**Rendering.** The pane is a single React root. Use `contain: layout paint` on the transcript; avoid measuring during streaming; tables and code blocks own their scroll. Set `data-theme` before mount from `matchMedia` + stored override. `prefers-reduced-motion` is read by CSS; no JS animation library.

**Prototype scope.** Build the wireframe states with mock data and a `?state=` switch; #13 extended them to ten presets: `empty`, `loading`, `streaming`, `answer`, `review`, `partial`, `applied`, `history`, `settings`, `error`. The prototype is clickable for: suggestions fill composer, send/stop, approval expand/collapse, include/exclude, apply progress, partial retry, history open, settings sections, theme toggle. The workbook is mocked; Office.js loading is not required for #13. Keep the exact 380 px and 260 px webview widths.

**Performance budgets.** First paint < 300 ms in WebView2 on this machine; initial JS < 180 KB gzip; streaming commit ≤ 100 ms per frame at 10 fps; no horizontal page scroll at 240–720 px; no layout shift when the model chip loads.

**Out of scope for #13.** Real Office.js calls, real gateway streaming, persistence, auth, the ribbon manifest, and production packaging. Those belong to the build effort and #12/#11.

---

## 12. Open questions owned elsewhere

| Ticket | Question | What this design assumes |
| --- | --- | --- |
| [#9](https://github.com/mustaphaelou/excel-agent/issues/9) | Approval policy, auto-apply tiers, undo/rollback, dangerous functions, audit, cell-edit retry | The card can render review / applied / partial; auto-applied is a compact variant; risky row notes are slots for #9's policy. |
| [#10](https://github.com/mustaphaelou/excel-agent/issues/10) | Context assembly, citation encoding, hidden-sheet navigation, token budget | Citations are structured tokens; chips are buttons; hidden/protected sheets use the snackbar path. |
| [#11](https://github.com/mustaphaelou/excel-agent/issues/11) | Storage, workbook identity, retention cap, instructions scope, compaction | History is per workbook with an "all" escape hatch; instructions are per workbook; no `Claude Log` sheet. |
| [#12](https://github.com/mustaphaelou/excel-agent/issues/12) | Tool surface, SSE parsing, model catalog, status source, error codes, question cards | Status is tool-derived; model menu reads `/api/models`; no `ask_user` tool, so v1 asks in prose; errors map to the copy deck. |
| [#7](https://github.com/mustaphaelou/excel-agent/issues/7) | Default vs shared runtime | Default runtime. A global shortcut would require amending ADR 0001; it is not in v1. |
| [#13](https://github.com/mustaphaelou/excel-agent/issues/13) | Framework, prototype scope, mock data | Production React 19 + TypeScript + Vite; throwaway prototype had 10 scenario presets and 260/380/520/720 widths; no Office.js. |

**Confirmed with the dev on 2026-11-07:**

1. **A — warm paper, Fluent bones.** Not a Fluent-neutral pane with only Claude layout, and not a full Claude surface.
2. **Inline edit card, expanded by default.** Not a Claude-style composer action bar with a hidden list, and not a dedicated review sheet for every batch.
3. **No global keyboard shortcut in v1.** The default runtime stays; the ribbon and opt-in auto-open are the entry points.
4. **History and instructions are per workbook first,** with "All workbooks" as an escape hatch. Not app-scoped like Claude, and not global defaults with per-file overrides.

---

## 13. Evidence and corrections

- Pane width range/default/website: [`Office.TaskPane`](https://learn.microsoft.com/en-us/javascript/api/office/office.taskpane?view=excel-js-1.21): Windows 86 px–50%, default 320 px; web Excel 350–500 px; `setWidth` only; out-of-range values are ignored silently.
- Width persistence: [#6 spike, "Pane width, dock, and `locked`"](../../wayfinder/research/06-pane-sideload-spike.md): 376 → 326 after close/reopen; saved `width="600"` not applied; `dockstate` empty; `locked` ignored. This contradicts the ticket body's persistence claim and is the basis for the stored-preference design.
- Pinning: [`SupportsPinning`](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/action) is Outlook-only. Excel has no pinning; `Office.addin.setStartupBehavior`/`showAsTaskpane` are SharedRuntime APIs and do not create a pin.
- Auto-open: [`Automatically open a task pane with a document`](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document): `Office.AutoShowTaskpaneWithDocument` via document settings; one pane; only centrally deployed or sideloaded add-ins; ignored if the add-in is not installed.
- Keyboard shortcuts: [`Add custom keyboard shortcuts`](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts): requires SharedRuntime 1.1 and an `ExecuteFunction` action. ADR 0001 chose the default task-pane runtime, so a global shortcut is an architectural change, not a UX toggle.
- Theme: [`Office.OfficeTheme`](https://learn.microsoft.com/en-us/javascript/api/office/office.officetheme?view=common-js-preview): `isDarkTheme`, `themeId`, `fluentThemeData`; `OfficeThemeChanged` is documented on the Outlook mailbox object. Use `prefers-color-scheme` for live changes in Excel.
- Citations, status line, approval model, history/instructions scoping, palette samples and the copy/avoid table: [Claude for Excel sidebar teardown](../../wayfinder/research/04-claude-sidebar-teardown.md).
- One pane per workbook, hidden-sheet `select`/`activate` throws, cell-edit mode fails `Excel.run`: [#2 capability envelope](../../wayfinder/research/01-officejs-capability-envelope.md) and [#6 spike](../../wayfinder/research/06-pane-sideload-spike.md).
- Contrast ratios in §3.2 were computed with the WCAG 2.2 relative-luminance formula; the citation pair deliberately departs from Claude's sampled 3.05:1.

---

## 14. Annotated wireframes

These are the **original A frames**; the validated B visual source is the [#13 prototype](https://github.com/mustaphaelou/excel-agent/tree/prototype/clickable-sidebar-13/docs/wayfinder/design/13-prototype), with the amendment at the top of this record. Do not implement the 44 px header or full-width user blocks from these frames. Open [08-wireframes.html](08-wireframes.html) only for the original A layout record. Quick view: [contact sheet of all eight frames](assets/08-sheet-contact.jpg). Full sheet: [08-sheet-preview.jpg](assets/08-sheet-preview.jpg). The sheet contains:

1. **First run** — header, greeting, three suggestions, trust line, focused composer.
2. **Streaming answer** — status line, markdown, table, inline citations, stop.
3. **Review a batch** — expanded edit card, risk badges, per-row diff, footer decision, gated composer.
4. **Partial apply** — applied count, cell-edit alert, failed rows, scoped retry, gated composer.
5. **History** — scope control, active row, grouping, footer delete.
6. **Settings** — instructions scope, model, theme, auto-open disclosure, history.
7. **Narrow, dark** — 260 px, warm charcoal, scrolling table, wrapped citations.
8. **Connection lost** — health banner, timeout card, details, disabled composer.

The frames are 1:1 CSS px at 380 px webview (260 px for frame 7). Annotation pins sit outside the pane so they never cover the UI. The sheet is static and shows A; the #13 prototype was the clickable test and validated B.

---

*This document is the ticket's supporting record. The authoritative decision summary is the `## Answer` comment on [#8](https://github.com/mustaphaelou/excel-agent/issues/8); the map gists it.*
