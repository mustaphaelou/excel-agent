# Clickable sidebar prototype — tickets #13 and #16

**Throwaway.** This directory is the primary source for the validated direction, not production code. No framework, no dependencies, no Office.js, no gateway, no persistence. Everything is mocked in the browser.

> **Validated 2026-11-08:** the dev chose **B — Claude-close** as the visual direction. The prototype uses #8's accessible action token `#B55434` for filled actions; Claude's `#D97757` is the mark/illustration only. See [issue #13's Answer](https://github.com/mustaphaelou/excel-agent/issues/13).
>
> **Extended for #16** with the six safety and persistence surfaces: link picker, compaction, retention, store failure, instructions migration and read-only mode. The states reuse B's shell and composer; no new visual direction. See [issue #16's Answer](https://github.com/mustaphaelou/excel-agent/issues/16).

## Run

```bash
cd docs/wayfinder/design/13-prototype
npm start
# http://127.0.0.1:4178/?variant=hybrid&state=empty&width=380&theme=auto
```

`node serve.mjs` and `node serve.mjs 8080` work too. There is one route; all state is in the query string. No `npm install` is needed: the prototype has no dependencies.

**Stack note.** The prototype is dependency-free HTML/CSS/JS so a reviewer can run it without a toolchain. The production pane should still follow #8's React 19 + TypeScript + Vite recommendation; this is a visual/interaction probe, not the implementation.

## The question

Which visual direction should the pane take, and which interaction patterns survive contact with a real, clickable conversation?

- **A · Hybrid** (`?variant=hybrid`) — the #8 direction: warm paper reading column, Fluent geometry, inline expanded diff card. This is the primary candidate.
- **B · Claude-close** (`?variant=claude`) — **validated**: floating paper surface, editorial serif, right-aligned user bubbles, a large floating composer.
- **C · Fluent-native** (`?variant=fluent`) — command bar, message avatars, denser type, Excel-blue action role, a review table instead of the diff card. Tests how much Excel should win.

Variant keys are shareable and reload-stable. The floating bottom bar cycles them with the arrow buttons or `←` / `→` when focus is outside a field.

## Scenarios

`?state=` selects a preset: `empty`, `loading`, `streaming`, `answer`, `review`, `partial`, `applied`, `history`, `settings`, `error`.
The #16 extension adds `link`, `unsaved`, `compaction`, `compaction-progress`, `compaction-stacked`, `retention`, `retention-full`, `storefail`, `instructions`, `readonly` and `readonly-forced`.

The preset is a starting point, not a lock. The prototype is clickable:

- suggestions fill the composer (they do not send)
- send / stop, mocked streaming, markdown, tables, citation chips
- citation chips select the matching sheet and range in the mock grid; the formula bar and selection follow
- the edit card supports per-row inclusion, overwrite acknowledgement, apply progress, applied, partial failure, retry, dismiss, discard confirmation
- history opens, scopes to this workbook or all workbooks, renames, deletes, exports `.md`/`.json`, and blocks opening another workbook's conversation (one pane per workbook)
- settings takes per-workbook instructions with a debounced Saved state, model menu, theme, auto-open, per-workbook history management, and the read-only switch
- **link picker**: filter old conversations by path or name, move one here, then Undo or "Move back to {workbook}"; the current workbook's empty state shows the unsaved-workbook block
- **compaction**: 80% notice, "Summarize earlier turns" progress, a transcript divider, "View original" read-only, and a second stacked cut
- **retention**: 80%/100% budget banners, per-workbook size/count/delete, Open folder, export actions, delete-all confirmation
- **store failure**: persistent "History isn't being saved" banner, Retry, the corrupt-file notice, and a recovered conversation
- **instructions migration**: the "Copy instructions from FY27 Budget.xlsx" chip appears only while the field is empty
- **read-only**: header lock chip opens Settings, Settings switch is the control, forced-on for an Excel read-only workbook, and the composer stays usable for questions while writes are refused

`?width=260|380|520|720` changes the webview width; the pane is a container, so the layout reflows. `?theme=auto|light|dark` changes the pane and the mock host. Keyboard `←` / `→` cycles variants; `Esc` stops a stream or closes a menu/dialog.

## Snapshots

At 380 px / 1366 × 768, light:

- [Hybrid — review](screenshots/hybrid-review.png)
- [Claude-close — first run](screenshots/claude-empty.png)
- [Claude-close — review](screenshots/claude-review.png)
- [Fluent-native — review](screenshots/fluent-review.png)
- [#16 link picker](screenshots/16-link.png)
- [#16 unsaved workbook](screenshots/16-unsaved.png)
- [#16 compaction notice](screenshots/16-compaction-notice.png) · [progress](screenshots/16-compaction-progress.png) · [divider](screenshots/16-compaction-divider.png) · [original](screenshots/16-compaction-original.png) · [second cut](screenshots/16-compaction-second.png)
- [#16 retention at 80%](screenshots/16-retention-80.png) · [over budget](screenshots/16-retention-100.png)
- [#16 store failure](screenshots/16-store-failure.png)
- [#16 instructions copy](screenshots/16-instructions-copy.png)
- [#16 read-only chat](screenshots/16-readonly-chat.png) · [forced-on Settings](screenshots/16-readonly-forced.png)

## What to judge

1. **Reading rhythm at 380 px** — does the warm paper + serif run-in + 14/22 body hold up with a table and citations in the same turn?
2. **The diff card** — is one inline gate with per-row inclusion fast enough to use every turn, or does the Fluent table scan better?
3. **Citation chips** — do they read as evidence rather than decoration? Do they compete with links/code?
4. **The Fluent chrome** — does the command bar make the pane feel native, or does it turn the conversation into a spreadsheet command panel?
5. **Claude-close** — how much of the floating, editorial surface is worth carrying into a docked pane?
6. **States** — is the gated composer explained, is partial failure honest, do history and instructions belong in this pane?
7. **#16 surfaces** — does the link picker make a Save As orphan recoverable without making moving files feel casual? Is the compaction notice calm enough to leave in place? Does Manage history belong in Settings or in History? Does the read-only chip earn its header pixels?

## Notes

- The host window, spreadsheet, sheet tabs, ribbon, and formula bar are a mock. The pane is the prototype.
- The sidebar is intentionally 260–720 px; the default is 380.
- The baseline and research for this direction live in [#8's UX record](../08-sidebar-ux.md) and [#9's safety model](../09-edit-safety.md).
- Capture the winning direction in the ticket's `## Answer` and the map's Decisions, then move this directory to the throwaway branch (it should not be merged to `main`).
