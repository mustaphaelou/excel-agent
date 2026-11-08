# Clickable sidebar prototype — ticket #13

**Throwaway.** This directory is the primary source for the validated direction, not production code. No framework, no dependencies, no Office.js, no gateway, no persistence. Everything is mocked in the browser.

> **Validated 2026-11-08:** the dev chose **B — Claude-close** as the visual direction. The prototype uses #8's accessible action token `#B55434` for filled actions; Claude's `#D97757` is the mark/illustration only. See [issue #13's Answer](https://github.com/mustaphaelou/excel-agent/issues/13).

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

The preset is a starting point, not a lock. The prototype is clickable:

- suggestions fill the composer (they do not send)
- send / stop, mocked streaming, markdown, tables, citation chips
- citation chips select the matching sheet and range in the mock grid; the formula bar and selection follow
- the edit card supports per-row inclusion, overwrite acknowledgement, apply progress, applied, partial failure, retry, dismiss, discard confirmation
- history opens, scopes to this workbook or all workbooks, renames, deletes, and blocks opening another workbook's conversation (one pane per workbook)
- settings takes per-workbook instructions with a debounced Saved state, model menu, theme, auto-open, history management

`?width=260|380|520|720` changes the webview width; the pane is a container, so the layout reflows. `?theme=auto|light|dark` changes the pane and the mock host. Keyboard `←` / `→` cycles variants; `Esc` stops a stream or closes a menu/dialog.

## Snapshots

At 380 px / 1366 × 768, light:

- [Hybrid — review](screenshots/hybrid-review.png)
- [Claude-close — first run](screenshots/claude-empty.png)
- [Fluent-native — review](screenshots/fluent-review.png)

## What to judge

1. **Reading rhythm at 380 px** — does the warm paper + serif run-in + 14/22 body hold up with a table and citations in the same turn?
2. **The diff card** — is one inline gate with per-row inclusion fast enough to use every turn, or does the Fluent table scan better?
3. **Citation chips** — do they read as evidence rather than decoration? Do they compete with links/code?
4. **The Fluent chrome** — does the command bar make the pane feel native, or does it turn the conversation into a spreadsheet command panel?
5. **Claude-close** — how much of the floating, editorial surface is worth carrying into a docked pane?
6. **States** — is the gated composer explained, is partial failure honest, do history and instructions belong in this pane?

## Notes

- The host window, spreadsheet, sheet tabs, ribbon, and formula bar are a mock. The pane is the prototype.
- The sidebar is intentionally 260–720 px; the default is 380.
- The baseline and research for this direction live in [#8's UX record](../08-sidebar-ux.md) and [#9's safety model](../09-edit-safety.md).
- Capture the winning direction in the ticket's `## Answer` and the map's Decisions, then move this directory to the throwaway branch (it should not be merged to `main`).
