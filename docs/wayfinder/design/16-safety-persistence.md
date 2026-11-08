# Safety and persistence surfaces

Ticket: [#16 · Prototype the missing safety and persistence surfaces](https://github.com/mustaphaelou/excel-agent/issues/16)
Status: answered · 2026-11-08
Prototype: [`prototype/safety-persistence-16`](https://github.com/mustaphaelou/excel-agent/tree/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype) (throwaway, not merged)
Depends on: [#8](https://github.com/mustaphaelou/excel-agent/issues/8), [#9](https://github.com/mustaphaelou/excel-agent/issues/9), [#11](https://github.com/mustaphaelou/excel-agent/issues/11), [#13](https://github.com/mustaphaelou/excel-agent/issues/13)

## Question

#11 assigned the link picker, compaction, retention, store-failure, instructions-copy and read-only surfaces to #13, but #13 did not implement them; #8's component inventory predates #11. Validate the states against the validated **B — Claude-close** direction, settle the read-only control, and fold the result into #8/#11.

## Verdict

**The B shell absorbs all six surfaces without a new visual direction.** The link picker, compaction notice/divider, retention section, store-failure banner, instructions chip and read-only chip/switch all use B's topbar, full-height views, composer and card language. No new screen geometry or colour role was needed.

### Link picker

A full-height view, not a menu, entered from History's empty state and the workbook menu. It lists every conversation from other workbooks with path and date, grouped by workbook, and filters by path or name. "Move to this workbook" is the only action. The move is one-way; the previous key is stored on the conversation and surfaced as "Move back to {workbook}" in the row overflow, so undo survives a reload as well as the immediate toast. An unsaved workbook has no key: History shows "History starts when you save this workbook" and the link entry is disabled with "Save the workbook first".

### Compaction

The 80% notice belongs above the composer, not in the global banner: it is conversation-scoped. "Summarize earlier turns" shows "Summarizing {n} turns…" with progress, then inserts a divider at the cut. "View original" expands the pre-cut turns read-only, with no Copy/Retry/Steps actions. A second compaction stacks a second divider. Nothing auto-summarizes; originals are removed only with the conversation.

### Retention

Manage history is a Settings section, opened by the 80%/100% global budget banner: total against 250 MB, one row per workbook with path/count/size/Delete, Open folder, and "Delete all history…". Per-conversation `.md`/`.json` export lives in the History row overflow. Delete-all confirms with the count and size. Nothing is deleted automatically.

### Store failure

A persistent warning banner says "History isn't being saved. Your conversation is still open." with Retry. History shows the corrupt-file notice and a recovered conversation with a "Recovered" badge. Retry clears the banner and confirms "History is being saved again."

### Instructions migration

The "Copy instructions from {workbook}" chip appears only while the instructions field is empty. One click replaces the whole field; the chip disappears and the normal Saved state runs. There is no append path.

### Read-only control — settled

- **Chip vs Settings was a false choice.** The header lock chip is a **status indicator and Settings shortcut**; the Settings → Editing switch is the only on/off control. This keeps #8 C4b's "not a toggle in the header" rule and gives #9's forced-on state somewhere to explain itself.
- **Persisted field/API:** `readOnly` in `workbook.json`, changed through `PUT /api/history/workbooks/:key/read-only` with `{readOnly: boolean}`; never written into the workbook.
- **Effective state:** `setting.readOnly || Workbook.readOnly`, recomputed when the workbook opens. When Excel reports the workbook read-only, the switch is shown on and disabled; the persisted preference is unchanged, so a later writable copy opens in the user's chosen mode.
- **Pending proposal:** the switch is disabled until Apply/Discard. A restored proposal cannot be applied while the workbook is read-only.
- **Composer:** chat and reads stay enabled, with "Read-only mode — questions only" and "Read-only mode. Edits are unavailable; questions still work." Write tools return #9's `ReadOnly` error without staging a card. The composer is write-gated, not send-disabled.
- **Copy:** chip "Read-only"; user tooltip "Read-only mode is on. Questions only. Open Settings to change it."; forced tooltip "This workbook is open read-only. Questions only."; forced footnote "Read-only mode is forced on by Excel and can't be turned off here."

## What changed in the records

| Record | Change |
| --- | --- |
| [#8](08-sidebar-ux.md) | Added the **Safety and persistence amendment — 2026-11-08**: six surfaces, C4b/C6/C25/C27/C39/C41-C45/C49 amendments, new C51-C58, state matrix and copy deck. |
| [#9](09-edit-safety.md) | §2.4/§4.2 now name the chip as status + Settings entry, the Settings switch as the control, the persisted field/API and the forced-on/pending states. |
| [#11](11-history-and-instructions.md) | Added `readOnly` to `workbook.json`, the read-only endpoint, `linkedFrom` on the conversation envelope, the link-picker/unsaved detail and the read-only preference section. |
| [#14](14-handoff.md) | Gap closed; phase 5 can start after the rest of the handoff. |

## Prototype evidence

Run:

```bash
cd docs/wayfinder/design/13-prototype
npm start
# http://127.0.0.1:4178/?variant=claude&state=link&width=380&theme=light
```

New presets: `link`, `unsaved`, `compaction`, `compaction-progress`, `compaction-stacked`, `retention`, `retention-full`, `storefail`, `instructions`, `readonly`, `readonly-forced`. Every surface is clickable: filter/move/undo, summarize/view original/second cut, budget delete flows, retry store, copy instructions, toggle/forced read-only.

Screenshots, 1366 × 768, pane at 380 px, light:

- [Link picker](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-link.png) · [unsaved workbook](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-unsaved.png)
- [Compaction notice](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-compaction-notice.png) · [progress](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-compaction-progress.png) · [divider](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-compaction-divider.png) · [original](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-compaction-original.png) · [second cut](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-compaction-second.png)
- [Retention 80%](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-retention-80.png) · [over budget](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-retention-100.png)
- [Store failure](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-store-failure.png)
- [Instructions copy](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-instructions-copy.png)
- [Read-only chat](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-readonly-chat.png) · [forced-on Settings](https://raw.githubusercontent.com/mustaphaelou/excel-agent/prototype/safety-persistence-16/docs/wayfinder/design/13-prototype/screenshots/16-readonly-forced.png)

Checks: all 21 presets load without console errors at 380 px and no pane/doc horizontal overflow; the new states were also checked at 260 px and in dark mode; Lighthouse accessibility/best-practices/SEO/agentic browsing score 100 on the link picker, retention and forced read-only samples. This is a prototype check, not the #15 test suite.

## Boundaries

- No production code was added. The prototype remains dependency-free HTML/CSS/JS on a throwaway branch.
- No new visual direction. The screenshots are the B shell with the #16 states.
- The exact store paths, copy and state names in the prototype are mocks; #11's API and #8's copy deck are the source for implementation.
- Read-only, link-back metadata and store recovery still need real-Excel/store verification in the build phase.
