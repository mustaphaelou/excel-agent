# Edit safety and approval model

Ticket: [#9 · Edit safety and approval model](https://github.com/mustaphaelou/excel-agent/issues/9)
Decided with the dev on 2026-11-08. Builds on the capability envelope ([#2](https://github.com/mustaphaelou/excel-agent/issues/2), [fact sheet](../research/01-officejs-capability-envelope.md)), the live pane spike ([#6](https://github.com/mustaphaelou/excel-agent/issues/6), [record](../research/06-pane-sideload-spike.md)), and the sidebar contract ([#8](https://github.com/mustaphaelou/excel-agent/issues/8), [design](08-sidebar-ux.md)). #8's edit card (C28–C38) and approval state machine (§6.2) are the UI; this document is the policy behind them.

---

## Summary

The pane reads freely and writes only through a reviewed proposal. In v1 every write is undoable with Excel's own undo, staged as a proposal, reviewed as an exact diff, and applied in one undo group. The model never writes directly, nothing is auto-applied, and non-undoable operations are refused.

The safety model is code-enforced, not prompt-enforced. The model proposes; the pane classifies risk, renders the diff, owns the acknowledgement, and owns the apply path. Workbook content and fetched pages can influence what the model proposes, but they cannot approve it.

### Decisions at a glance

1. **Staged writes only.** Write tools append to a proposal; there is no direct-write tool and no model-callable apply.
2. **Undoable scope.** v1 permits only operations Excel can undo. Sheet delete/copy/rename/reorder/visibility, protection, `save`/`close`, base64 sheet import, `AllowEditRange`, and data-connection/linked-workbook refresh are refused.
3. **Review + read-only.** Review every batch is the default; read-only is a per-workbook mode (forced on when `Workbook.readOnly` is true). Auto-apply is not in v1.
4. **One gate per turn.** One card per turn, one Apply/Discard, per-row inclusion. Excluded rows close as excluded when the card resolves.
5. **Risk defaults.** Data overwrites need one card-level acknowledgement; loss rows and confirm-tier formulas are excluded by default and require ticking the row.
6. **One run, one undo group.** Apply is one `Excel.run({ mergeUndoGroup: true })`, one `sync()` per row, stop at the first failure, "Applied N of M", user-initiated scoped retry, no auto-retry.
7. **Drift fails closed.** Apply re-reads every target; a row whose pre-state changed is not written. "Update diff" recomputes against the new state; no auto-merge.
8. **Ctrl+Z is the undo.** No programmatic undo, no in-pane undo button, no snapshot. The card names the undo step.
9. **No grid markup.** No post-apply highlights or comments; the card is the record, and the row range chip (plus a card-level Reveal) navigates.
10. **Local audit only.** Append-only record in local storage; no `Claude Log` sheet in the workbook.
11. **Dangerous formulas are tiered.** Exec/file/DDE functions are blocked; network/external-data functions need an explicit row tick; volatile functions are warnings.
12. **Caps.** 2,000 written cells and 50 card rows per turn; larger work is split across turns.
13. **Injection containment.** Deterministic diffs and risk classes, code-enforced staging, no auto-apply, and a code-level formula scan.

---

## 1. Invariants

- **I1 — Reads are free; writes are staged.** Any write tool produces proposal rows. It never queues a write against the workbook.
- **I2 — The model has no write path.** Apply is a pane action on a frozen proposal. The model cannot call it, extend it, or alter a rendered card.
- **I3 — The card is the source of truth.** Diffs, risk classes, counts, and ranges are derived by pane code from the structured operation and a pre-read. Model prose is commentary, not the review surface.
- **I4 — Everything applied is undoable, or it is not applied.** A non-undoable API is refused before it reaches a run. `UndoNotSupported` is the runtime backstop.
- **I5 — The user's workbook is authoritative at apply time.** Every selected row is re-read and compared; drift fails that row. There is no merge and no "apply as reviewed" override.
- **I6 — The grid is never repurposed as UI.** No highlight fills, no comments, no selection changes, no focus steals during apply.
- **I7 — Nothing is auto-applied.** Every write needs a gesture after review. Read-only refuses before staging.
- **I8 — A failed apply is reported as partial, not as atomic.** Earlier `sync()` calls persist; the card says what landed and scopes the retry.

---

## 2. Writable scope

### 2.1 Allowed in v1

| Area | Operations | Undo notes |
| --- | --- | --- |
| Cells | values, formulas, number formats, fills/fonts/borders/alignment, `clear` | Standard undo; `clear` is a loss row |
| Structure | row/column insert and delete; worksheet add | Insert and add are undoable; row/column delete is a loss row |
| Tables | create, resize, sort, filter, totals row edits | Header rename and totals overwrite are loss rows |
| Native operations | range/table sort, AutoFilter, pivot create/edit/refresh/delete, conditional formatting, data validation | Pivot/table **style** APIs may not redo; the card notes it |
| Sheet lifecycle | add only | Delete/copy/rename/reorder/visibility are refused |

### 2.2 Refused in v1

| API | Why |
| --- | --- |
| `Worksheet.delete`, `Worksheet.copy` | Not undoable |
| `Worksheet.name`, `Worksheet.position`, `Worksheet.visibility` | Not undoable on Windows; they also clear the undo stack from that call onward |
| `WorksheetProtection.*`, `WorkbookProtection.*` | Not undoable; a security boundary the agent should not cross |
| `Workbook.save`, `Workbook.close` | Not undoable; the user's file lifecycle |
| `Workbook.insertWorksheetsFromBase64`, `WorksheetCollection.addFromBase64` | Not undoable |
| `AllowEditRange.*` | Not undoable |
| `DataConnectionCollection.refreshAll`, `LinkedWorkbook.refresh` | Not undoable |
| `Range.copyFrom` as a tool | Undoable, but a copy hides its source in the diff; v1 writes are explicit |

The load-bearing fact is Excel's own undo table: the lifecycle APIs above are **not undoable**, and calling one **clears the undo stack from that call onward**. `Range.copyFrom` is excluded for diff transparency, not because it is non-undoable. There is no programmatic undo in the Excel JS API. `mergeUndoGroup` only groups the syncs of a single `Excel.run`; it does not roll back. See [Undo support](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities), [`Excel.RunOptions`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21), and the [envelope fact sheet](../research/01-officejs-capability-envelope.md).

### 2.3 Why no snapshot

`Office.context.document.getFileAsync(Office.FileType.Compressed)` can return the document file in slices ([Office.Document](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=excel-js-1.21)), but it is a point-in-time **file copy**, not an in-memory transaction:

- restoring it is a file-level action that discards every later edit, including the user's;
- it cannot be applied from the pane, and the pane cannot know whether the file's saved state matches the in-memory state;
- it does not answer "what did this apply do" the way the audit and per-row outcome do.

So v1 has no rollback machinery beyond Excel's undo. The safety comes from refusing non-undoable operations, not from trying to reconstruct a workbook after the fact. A v1.1 structural mode can be designed around an explicit "save a copy first" workflow, separate from the ordinary apply path.

### 2.4 Read-only

- A per-workbook local setting, persisted as `readOnly` in the workbook's `workbook.json` with the conversation store (#11), never written into the workbook. The API is `PUT /api/history/workbooks/:key/read-only` with `{ "readOnly": boolean }`; the value stores the user's preference only.
- The header lock chip is a **status indicator and Settings shortcut**, not a toggle. The switch in Settings → Editing is the only on/off control. This reconciles #8 C4b's non-toggle chip with #11's missing field; the #16 verdict is folded into #8.
- Effective mode is `setting.readOnly || Workbook.readOnly`, re-evaluated whenever the workbook opens. Forced on when `Workbook.readOnly` is true ([Workbook](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21)); the Settings switch is shown on and disabled, and its footnote explains why it cannot be turned off. The persisted preference is not changed, so a later writable copy opens in the user's chosen mode.
- The switch is disabled while a proposal is unresolved ("Apply or discard the pending changes before switching modes."); turning the mode on must not strand a card.
- When on, every write tool returns a structured `ReadOnly` error. No proposal is staged, no card is created. The assistant explains and offers to turn the mode off; the next turn can propose.
- Chat, reads and citations are unaffected; the composer stays enabled with "Read-only mode — questions only" and "Read-only mode. Edits are unavailable; questions still work." Protected sheets are not a mode; they are per-row blockers (§3.5).

---

## 3. Risk model

### 3.1 Pre-state and diff

At staging, the pane reads the exact fields the operation will touch:

- **values / formulas**: `values`, `formulas`, `valueTypes`, and `text` for the displayed comparison;
- **formats**: the specific property being written (`numberFormat`, fill, font, borders, alignment);
- **structure**: sheet id/name/position, table/pivot identity, row/column existence, and the used-range content summary;
- **native ops**: the target range, the operation summary, and the structural anchors (sheet/table/pivot, header/key column).

The diff is derived from that pre-read, never from the model:

| Operation | Diff |
| --- | --- |
| Value | `was → now`, using the cell's displayed text so locale and number format match the sheet |
| Formula | formula text, old → new |
| Format | a sentence (`General → $#,##0`) |
| Structural | a sentence (`Delete rows 12–18; 14 non-empty cells`) |
| Native | a summary (`Sort A2:F400 by Revenue, descending`) |

This is #8's C32. The pre-read is also the drift baseline.

### 3.2 Classes

| Class | What lands here | Default | Gate |
| --- | --- | --- | --- |
| **Normal** | Values/formulas into empty cells; number formats; fills/fonts/borders; sheet add; row/column insert; table create; sort/AutoFilter; pivot create/edit/refresh; conditional formatting; data validation | included | Apply |
| **Overwrite** | A non-empty value or formula is replaced, including formula→formula | included | card-level acknowledgement |
| **Loss** | Formula→value; clear contents or formats; row/column delete; table header rename; table totals row; hyperlink property add/change; delete a conditional-format or data-validation rule | **excluded** | tick the row |
| **Confirm formula** | `WEBSERVICE`, `RTD`, `IMPORTDATA/XML/HTML/FEED`, `STOCKHISTORY`, `INDIRECT`, external-workbook references, `HYPERLINK` formula | **excluded** | tick the row |
| **Blocked formula** | `CALL`, `EVALUATE`, `REGISTER.ID`, `EXEC`, `FILES/FOPEN/FWRITE/FCLOSE`, DDE (`|`), `CMD` | — | tool error |
| **Blocked row** | Protected sheet; merged target; pivot output; spill child; missing anchor | unselectable | fix, then Re-check |

"Non-empty" means the pre-state has a value or a formula, tested on `valueTypes`/`formulas`, not on displayed text. A formula returning `""` is non-empty. Format-only writes are not overwrites.

### 3.3 Classification order

1. Not in the v1 allowlist, non-undoable, over cap, overlapping, or read-only → **tool error**; no card.
2. Blocked formula → **tool error**; no card.
3. Target condition (protected, merged, pivot output, spill child, missing anchor) → **blocked row**.
4. Confirm-tier formula or hyperlink → **confirm** row.
5. Loss operation → **loss** row.
6. Non-empty pre-state replaced → **overwrite** row.
7. Otherwise → **normal** row.

Loss wins over overwrite: a row that clears a formula is one loss row, not two acknowledgements.

### 3.4 Acknowledgements

- **Overwrite**: one checkbox in the card footer — "Overwrite 12 cells that contain data" — disabled Apply until checked. The count is recomputed as rows are toggled.
- **Loss / confirm**: the row is excluded by default. Ticking its include checkbox is the acknowledgement; the note names what is lost or where the formula reaches. No second checkbox.
- Formula→formula is an overwrite, not a loss: the dependency chain stays intact and the diff shows both formulas. A formula→value row is a loss.
- Changing a format over an existing format is normal; clearing a format is a loss.
- The acknowledgement is per card and per row, never modal, never persisted across cards.

### 3.5 Blocked rows

Staged with a reason, unselectable, and never written:

- **Protected sheet**: v1 does not attempt writes to unlocked cells; unprotect the sheet and press **Re-check**.
- **Merged target**: allowed only when the target is exactly the merged area, or a single anchor cell; otherwise blocked.
- **Pivot output**: any target intersecting `PivotTable.layout.getRange()` ([PivotLayout](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivotlayout?view=excel-js-1.21)); refresh the pivot or edit its source instead.
- **Spill child**: a target inside a spilled range but not its anchor.
- **Missing anchor**: a sheet, table, pivot, row, or column that no longer exists.

Blocked rows do not gate the composer; they are a preview. A card with at least one selectable row gates as #8 specifies. Fixing the workbook and pressing **Re-check** re-runs classification and unblocks rows; nothing is retried silently.

### 3.6 Rejected at staging

Tool errors, not card rows:

- refused operation (§2.2);
- read-only mode;
- a blocked formula;
- more than 2,000 written cells or 50 rows;
- overlapping rows (§5.2);
- a no-op row is silently dropped; if every row is a no-op, the tool result says "no change" and no card is created.

The tool result is structured (`{ staged, dropped, rejected, risk }`) so the model can adapt without guessing. It never says "applied".

---

## 4. Approval modes

### 4.1 Review (default)

Every turn's writes are staged, rendered as one card, and resolved by Apply or Discard. There is no setting to skip the gate in v1.

### 4.2 Read-only

As §2.4. The header chip is the only header affordance and opens Settings; the Settings switch is the only control. The mode is per workbook; the model is told on the next turn. A read-only card does not exist because nothing is staged. When Excel reports the workbook read-only, the switch is forced on and cannot be turned off; the chip tooltip explains it. A restored pending proposal is shown but cannot be applied until the workbook is writable and the mode is off.

### 4.3 Auto-apply (deferred)

Not shipped. #8's `auto-applied` card state stays reserved and unreachable. Revisit only after the formula scanner and risk classifier have real mileage; any future auto-apply would be limited to closed-set operations with no formulas, no overwrites, and no loss, and would still be logged.

---

## 5. Staging contract

### 5.1 Write tools

The tool surface (#12) exposes declarative operations: set values, set formulas, set number format, set format properties, clear, insert/delete rows or columns, add sheet, table operations, sort, filter, pivot operations, conditional formatting, data validation. Each carries a sheet id, an A1 range, and a payload.

Not in v1: `copyFrom`, paste, defined names, macro sheets, data tables, external workbook writes. `Range.hyperlink` is a writable property ([Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21)) and is a loss row; the `HYPERLINK` formula is a confirm row.

### 5.2 Overlaps

Two rows may share cells only when they write **disjoint fields** — value/formula, number format, fill, font, border, alignment. Two rows writing the same field to overlapping cells, or a native op overlapping a cell write, is a **tool error**: "Split into separate turns." This keeps drift, acknowledgement, and diff independent.

Native operations claim their whole range: a sort of `A2:F400` cannot share the turn with a cell write inside `A2:F400`.

### 5.3 No-ops

A cell whose proposed value/formula/format already matches its pre-state is dropped from the row. A row with no remaining cells is dropped. Dropped rows are counted in the tool result, never shown as fake changes.

### 5.4 Turn end

- The proposal is frozen when the turn ends (complete, stopped, or error).
- A stopped or errored turn keeps its staged rows as a card labelled "Turn interrupted"; the exact diff still requires review.
- A pane close mid-turn loses the in-memory staging; nothing was written. There is no card to recover.
- A card is persisted when rendered (before any apply) so an interrupted apply can be verified (#11).

### 5.5 What the model is told

- The write tool result reports staged/dropped/rejected counts and the risk classes, never "applied".
- The system prompt says the user applies changes and the model must not claim a change landed.
- After the card resolves, the next turn receives a system note with the outcome (`applied`, `partial`, `discarded`, `excluded`) so the model does not re-propose blindly. #11/#12 own the mechanism; #9 requires it.

---

## 6. Approval state machine

#8's component states stay. The policy adds a card flag (`needs-ack`), a row flag (`blocked`), and apply-time states (`stale`, `waiting`, `interrupted`); no new screen.

```
write tool ──read-only──► tool error (no card)

staging ──turn ends──► proposed ──discard──► discarded
                          │
                          └──apply──► preflight
                                        ├──all rows drift──► stale ──Update diff──► proposed
                                        ├──some drift──────► applying (valid rows) ──► partial
                                        ├──cell editing────► proposed (refuse; nothing queued)
                                        └──ok──────────────► applying
                                                               ├──all ok──► applied
                                                               ├──row fails──► partial ──retry──► preflight
                                                               │                  └──dismiss──► dismissed
                                                               └──pane closes──► interrupted ──verify──► applied/partial
```

Drifted rows use **Update diff**; **Retry** applies only rows that failed at write time or were never attempted.

| State | What the user sees | Writes | Composer |
| --- | --- | --- | --- |
| `staging` | Status line; no card yet | none | enabled until the turn ends |
| `proposed` | Card expanded; risk badges; Apply/Discard | none | gated |
| `proposed` + `needs-ack` | Footer checkbox; Apply disabled | none | gated |
| `proposed` + blocked rows | Blocked notes; Re-check; selectable rows still gated | none | gated if any row is selectable |
| `stale` | Danger inset "Workbook changed"; Update diff | none | gated |
| `applying` | Per-row status; progress | writing | gated |
| `waiting` (race guard) | "Waiting for Excel to finish editing…" | queued | gated |
| `applied` | Collapsed "Applied N"; excluded/blocked counts; undo hint | done | enabled |
| `partial` | Applied N; failed rows; Retry N; Dismiss | partial | gated until retry/dismiss |
| `dismissed` | Card collapsed; History "Pending" | none | enabled |
| `interrupted` | "Apply was interrupted…"; Review/Retry | unknown | enabled |
| `discarded` | Card closed; no changes | none | enabled |
| `auto-applied` | reserved | — | — |

Rules:

- **One gating card per conversation.** After `dismissed`, the old card is background-pending; a later overlapping card supersedes its rows so they can never be retried into a double write.
- **Excluded rows close with the card.** Applying 9 of 12 resolves the card as "Applied 9, 3 excluded"; the excluded rows are logged and must be re-proposed in a new turn. They do not keep the conversation gated.
- **Blocked rows close with the card** if the card resolves; the user can fix and Re-check before applying, or discard and re-ask.
- **Discard never writes.** It closes pending and blocked rows; already-applied rows are untouched (there is no discard after a partial apply — the failure path is Retry/Dismiss).
- **Dismiss is not discard.** Failed rows remain pending in History; the composer ungates; they are never auto-applied.
- **Interrupted is verified, not trusted.** See §7.6.
- The model cannot mutate a `proposed` card; a later turn creates a new card.

---

## 7. Apply protocol

### 7.1 Pre-flight

Before any write:

1. Card is `proposed`/`partial retry`; at least one selectable row; confirmations complete.
2. Read-only is re-checked as defence in depth.
3. Every selected row is re-read and compared to its drift baseline (§7.3).
4. Targets are re-checked: sheet exists, not protected, not merged illegally, not pivot output, not a spill child; anchors exist.
5. Blocked rows are excluded; they never enter the run.
6. Caps are re-checked.
7. The cell-edit probe runs (§7.4); editing refuses the apply with nothing queued.

If every selected row drifts, the card enters `stale` and nothing is written. If some rows drift, the valid rows proceed and the drifted rows are failed with "Changed since proposed"; the card ends `partial`, and those rows are excluded from Retry (Update diff first).

### 7.2 Execution

```js
await Excel.run({ mergeUndoGroup: true, delayForCellEdit: true }, async (context) => {
  // one sync per card row, in card order
});
```

- **One run, one undo group.** All syncs in the run collapse into one Ctrl+Z step ([RunOptions](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21)).
- **One `sync()` per row**, so the card knows exactly which rows landed. The #6 spike confirmed that a failed later sync leaves earlier writes applied; per-row sync turns that into an honest "Applied N of M".
- **Stop at the first failure.** Do not sync again. The failed row and every unattempted row stay pending.
- **No non-undoable API in the run.** `UndoNotSupported` aborts the run if the scope ever leaks; that is a bug, not a retryable state.
- **No selection, activation, or focus changes.** Apply writes cells; it does not move the user's cursor.
- The run's rejection is caught; the card is updated from the per-row records, not from the promise alone.

### 7.3 Drift

At staging the pane captured the exact field-level pre-state. At apply it re-reads the same fields and compares:

- values/formulas: raw `values`/`formulas`/`valueTypes`, not displayed text;
- formats: the specific property;
- structure: sheet identity/position, table/pivot identity, row/column existence, used-range content summary;
- native ops: structural anchors and the header/key column; a native op is declarative and re-runs on current data, so it does not hash every cell.

A mismatch fails that row with "Changed since proposed" and writes nothing for it. **Update diff** re-reads the drifted rows, recomputes the same intended post-state against the new pre-state, reclassifies risk (a new overwrite/loss may appear), and returns the card to `proposed` for re-approval. No model call, no auto-merge, no "apply anyway".

### 7.4 Cell-edit collision

`delayForCellEdit` defaults to `false`: a batch processed while the user is in cell edit fails ([RunOptions](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21)); the #6 spike observed `InvalidOperationInCellEditMode`.

- **Probe first.** A trivial read with the default flag. If it throws `InvalidOperationInCellEditMode`, refuse: "Finish editing the cell, then Apply." Nothing is queued.
- **Race guard.** The apply itself uses `delayForCellEdit: true`. If the user enters edit mode in the small window after the probe, the batch waits and the card shows "Waiting for Excel to finish editing…". This batch cannot be cancelled once sent; closing the pane abandons it.
- **Reads too.** Foreground reads use the probe/delay; no background polling while editing.
- **No focus or selection changes** during apply, so the pane never causes the collision it is avoiding.

### 7.5 Partial failure and retry

- Applied rows are locked and their checkboxes disabled; the card shows "Applied N of M".
- The failed row keeps its diff and error; unattempted rows keep their diffs.
- **Retry N** re-runs only failed-at-write + unattempted rows, after a fresh drift check, with the same pre-flight. Drifted rows are not eligible for Retry; they must go through **Update diff** first. A retry is a new `Excel.run` and therefore a **new undo group**.
- **No auto-retry.** No backoff loop, no silent second attempt. Transient failures (cell edit, Excel busy) are retried by the user.
- **Dismiss** collapses the card and leaves the failures background-pending; the composer ungates. A later overlapping card supersedes them; they never apply themselves.
- A row that fails because its target vanished is a blocked/drift failure; Update diff or discard, not Retry.

### 7.6 Interrupted recovery

If the pane closes during `applying`, the default runtime tears the page down ([ADR 0001](../../adr/0001-runtime-topology.md)); the in-flight run dies and the workbook may be partially changed. The proposal was persisted before apply (#11), so on reopen the pane verifies rather than trusts:

- current state equals the intended post-state → **applied (verified)**;
- equals the pre-state → **not applied**;
- neither → **unknown** (the user changed it).

The card shows the counts and offers Review/Retry for the not-applied rows only; nothing is retried automatically. The audit is updated with the verification. If the card was never persisted, there is no recovery card; the audit's intent record is the only trace.

### 7.7 Spill

A written formula can spill outside its target range. Pre-apply, the row notes "This formula may spill; the diff shows the target, not the spill extent." Excel will not spill over non-empty cells (`#SPILL!`), so the overwrite gate still protects the target. Post-apply, the pane reads `getSpillingToRangeOrNullObject()` for each formula row ([Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21), 1.12) and reports "Spilled into B2:B40" in the applied row and the audit. The spill is part of the same undo group.

### 7.8 Concurrent edits

The pre-flight drift check is the only guard. Apply is sub-second and the composer is gated, but the grid stays live. A user edit that lands between rows can be overwritten by a later row in the same run. Accepted for v1; the audit records exactly what was written. Copy: "Don't edit the grid while Apply runs." Revisit with `Worksheet.onChanged` if a real collision shows up.

### 7.9 Recalculation

No forced calculation. Formula writes recalculate at commit in automatic mode; in manual mode the displayed values lag until the user recalculates. The diff shows formulas, not computed results, so the review is correct in either mode. The pane never calls `Application.calculate`.

### 7.10 Caps

- **2,000 written cells** per turn for cell-level writes (values, formulas, formats, clear); **50 card rows**.
- Native and structural operations count as one row and are exempt from the 2,000-cell count; the card names the exact range, and their drift check is structural (§7.3).
- Whole-column cell writes must be narrowed to the used range; a whole-column fill is over cap.
- Over cap → tool error with the count; the assistant splits the work across turns. The cap keeps the review surface, the undo group, and the per-row sync count bounded.

---

## 8. Undo and rollback

- **Ctrl+Z is the only undo.** There is no `Workbook.undo()`; `mergeUndoGroup` groups, it does not reverse. See [Undo support](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities) and the [envelope fact sheet](../research/01-officejs-capability-envelope.md).
- One apply attempt = one undo group. Retry is a new run and therefore a new group. Whether a **failed** run leaves a usable undo group for the rows that committed is not documented; the card does not promise a clean single step after a partial, only that Ctrl+Z is the recovery path.
- The applied card says "This apply is one Ctrl+Z step." After a retry, it says "Ctrl+Z undoes the last attempt; earlier attempts undo next."
- No in-pane Undo button, no snapshot, no file restore (§2.3).
- Because non-undoable operations are refused, every applied row is undoable. Style APIs may not redo; the card notes it.
- The undo stack can also contain the user's own edits. The hint means "before other edits, Ctrl+Z removes this apply"; it does not promise the apply is the next entry forever.
- A partial apply followed by Ctrl+Z is intended to remove the landed rows; the failed/unattempted rows were never written and remain pending. If the failed run's group is not intact, the card still shows which rows landed, so the user can repair them manually. This is why partial states are reported honestly.

---

## 9. Dangerous formulas

### 9.1 Tiers

| Tier | Functions / forms | Behaviour |
| --- | --- | --- |
| **Block** | `CALL`, `EVALUATE`, `REGISTER.ID`, `EXEC`, `FILES`, `FOPEN`, `FWRITE`, `FCLOSE`, `CMD`, DDE (`|`), Excel 4 macro functions | Tool error; never written |
| **Confirm** | `WEBSERVICE`, `RTD`, `IMPORTDATA`, `IMPORTXML`, `IMPORTHTML`, `IMPORTFEED`, `STOCKHISTORY`, `INDIRECT`, external-workbook references, `HYPERLINK` | Row excluded by default; tick to include; note names the reach |
| **Warn** | `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET`, `CELL`, `INFO` | Risk note only; no extra gate |

The block tier is code execution or file access with no legitimate v1 use. The confirm tier is a deliberate user choice: `WEBSERVICE` can send workbook data to an external URL; `IMPORT*` and `RTD` pull external data; `INDIRECT` can build external references; `HYPERLINK` adds a clickable destination. The warn tier is volatile but ordinary.

### 9.2 Scanner

- Runs on every formula write and on the `hyperlink` property.
- Normalizes case and whitespace; strips `_xlfn.`/`_xlws.` prefixes.
- Ignores function names inside double-quoted string literals.
- Distinguishes external references from structured table references: an external reference has `]` followed by a sheet name and `!` (e.g. `[Book1.xlsx]Sheet1!A1`), while `Table1[Column]` does not.
- Flags DDE by `|` outside string literals.
- `Range.formulas` is the A1/English formula text; `formulasLocal` is the user's language. English Excel is the v1 target; on a non-English workbook, also scan `formulasLocal` against a localized list (open item for #12).
- Defined names are not in the v1 tool surface; if they are added, their formulas must be scanned too.

---

## 10. Audit

- **Local only.** Append-only records in the local store — `audit.jsonl`, one line per transition, in the workbook's directory; #11 owns the store, retention and deletion ([11-history-and-instructions.md](11-history-and-instructions.md)). No `Claude Log` sheet; Claude's sheet is a workbook write and travels with the file ([teardown](../research/04-claude-sidebar-teardown.md)).
- **Written at intent, updated at outcome.** The record is created when the card is rendered and marked `applying` before the first write, so a crash or pane close leaves a trace. No document-close event exists to flush on ([envelope](../research/01-officejs-capability-envelope.md)); transitions are the flush points.
- **Shape** (v1):

```json
{
  "v": 1,
  "id": "turn-…",
  "conversationId": "…",
  "workbook": { "name": "Q3 Forecast.xlsx", "url": "…" },
  "model": "…",
  "at": "2026-11-08T10:00:00Z",
  "request": "truncated user message",
  "rows": [{
    "i": 0, "sheetId": "…", "sheetName": "Forecast",
    "range": "B4:B12", "op": "setFormula", "fields": ["formulas"],
    "before": "…", "after": "…", "risk": "overwrite",
    "included": true, "ack": true,
    "outcome": "applied|failed|skipped|excluded|blocked",
    "error": null, "spill": null, "durationMs": 12
  }],
  "decision": "pending|applied|partial|discarded|dismissed|interrupted",
  "attempts": [{ "n": 1, "startedAt": "…", "finishedAt": "…", "outcome": "partial" }]
}
```

- `before`/`after` are truncated (500 chars per row) and stored as a hash plus dimensions above 32 cells, so a large range cannot bloat the store.
- The audit is a record, not a security boundary. It does not make an unsafe apply safe; it makes an applied change reviewable after the fact.
- Clearing history clears the audit; it is not written into the workbook and does not outlive the local store.

---

## 11. Prompt injection

Workbook content and fetched pages are attacker-controlled input. The model may be persuaded to propose a harmful write; it must not be able to make that write land.

- **Data, not instructions.** Cell values, sheet names, comments, defined names, and fetched pages are wrapped as untrusted data with provenance; the system prompt says so; tool results are structured, not prose commands.
- **The gate is code.** Write tools stage; only the pane applies; the model cannot call apply, extend a rendered card, or alter its risk class. Approval cannot be prompt-injected because it is not a prompt.
- **The review surface is code.** Diffs, counts, ranges, and risk notes are derived from the structured operation and the pre-read. The model's prose sits above the card; the card does not quote it as evidence. Injected text cannot hide a row or mark it safe.
- **Dangerous formulas are scanned in code.** A block-tier formula is a tool error; a confirm-tier formula needs a row tick.
- **No auto-apply.** The cheapest exfiltration path — write a network formula and let recalculation send data — requires a user gesture.
- **Fetched pages.** The backend fetch must allow only `http(s)`, reject loopback/private/link-local targets and redirects to them, and attach no credentials (SSRF boundary; #12 owns the fetch). The pane renders sanitized markdown only; no raw HTML; links are externalized and marked (#8 C16).
- **Claims.** The assistant must not say a change is applied; the tool result says "staged". The card/apply state is authoritative. The next turn receives a system note with the real outcome.
- **No new badge.** The gate is constant, so a per-turn "this read untrusted content" warning would be noise. The existing trust line stays: "Stored locally · Nothing is written without approval."

---

## 12. Post-apply navigation

- **No highlights, no comments.** They are workbook writes, they clobber or pollute the user's file, and they exist in Claude because Claude has no confirmed diff ([teardown](../research/04-claude-sidebar-teardown.md)). Our diff is the review surface.
- **Row chips navigate.** Each applied row keeps its range chip (#8 C31); clicking selects the range.
- **Reveal.** The applied card's Reveal selects all applied ranges on visible sheets via `RangeAreas.select` (1.18, [RangeAreas](https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeareas?view=excel-js-1.21)). Hidden targets are skipped with a snackbar naming them; hidden-sheet navigation is #10's policy, and #6 proved `select()`/`activate()` throw there.
- **No selection during apply.** Reveal is user-initiated, after the run.
- If the user wants a visible marker, that is a normal format proposal through the gate, not an automatic write.

---

## 13. Copy deck additions

| Context | Copy |
| --- | --- |
| Read-only refusal | "Read-only mode. Turn it off in the header to propose edits." |
| Overwrite acknowledgement | "Overwrite 12 cells that contain data." |
| Formula loss | "Replaces formula =SUM(B2:B9)." |
| Clear | "Clears 5 cells, including 2 formulas." |
| Row/column delete | "Deletes rows 12–18 (14 non-empty cells)." |
| Confirm-tier formula | "Uses WEBSERVICE — can send workbook data to an external service." |
| Block-tier formula | "CALL isn't allowed: it can run external commands." |
| Cell edit | "Finish editing the cell, then Apply." |
| Waiting for Excel | "Waiting for Excel to finish editing…" |
| Drift | "Workbook changed since this was proposed. Update diff to review." |
| Applied | "Applied 9 of 12. 3 excluded." |
| Partial | "Applied 9 of 12. 3 failed. Retry 3." |
| Dismissed | "3 changes still pending." |
| Interrupted | "Apply was interrupted when the pane closed. Review before retrying." |
| Undo | "This apply is one Ctrl+Z step." / after retry: "Ctrl+Z undoes the last attempt; earlier attempts undo next." |
| Spill | "This formula may spill; the diff shows the target." / after: "Spilled into B2:B40." |
| Cap | "That's 3,400 cells; the limit is 2,000 per turn. I can split it." |
| Overlap | "Two changes target the same cells. Split them into separate turns." |
| Hidden sheet | "Targets hidden sheet 'Raw Data'." |
| Protected sheet | "Sheet 'Inputs' is protected. Unprotect it, then Re-check." |
| Merged cells | "B4 is merged. Target the anchor cell or unmerge first." |
| Pivot output | "C10:F40 is a pivot output. Refresh the pivot or edit its source." |

---

## 14. Known constraints

- **Concurrent edits during apply** are not locked (§7.8).
- **Spill extents** are known only after the formula is written (§7.7).
- **Manual calculation** leaves displayed values stale (§7.9).
- **Localized function names** may need a second scan on non-English workbooks (§9.2).
- **Protected sheets** are blocked, not partially written; v1 does not inspect `locked` per cell.
- **No document-close event**; the audit flushes on transitions, not on close.
- **Pane close during apply** is an interrupted apply, verified on reopen; the default runtime dies with the page.
- **Per-row sync** adds latency; the 50-row cap bounds it.
- **Style APIs** (pivot/table style) may undo without redo.
- **`UndoNotSupported`** is a backstop against a scope leak; it is not a user-facing flow.
- **The pane cannot lock the grid**; the user can type during apply.
- **Hidden sheets** are writable with a warning; Reveal skips them.
- **The audit is local** and tied to the conversation store; it is not a workbook artifact.

---

## 15. Owned elsewhere

- **#10** — context assembly, citation encoding, hidden-sheet navigation, selection semantics.
- **#11** — storage, workbook identity, retention, instructions, persistence of cards and audit.
- **#12** — tool schemas, SSE, model catalog, status text, error codes, the fetch/SSRF boundary.
- **#13** — the clickable prototype's states and copy.
- **#14** — blueprint completeness.
- **#7 / ADR 0001** — runtime; a global shortcut or a shared runtime would reopen that decision, not this one.

---

## 16. Sources

- [Undo support with the Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities) — undo baseline, unsupported-API table, `UndoNotSupported`, no programmatic undo.
- [`Excel.RunOptions`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-1.21) — `mergeUndoGroup` (1.20), `delayForCellEdit`.
- [`Excel.Range`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-1.21) — `values`/`formulas`/`valueTypes`/`text`, `numberFormat`, `format`, `clear`, `insert`/`delete`, `hyperlink` (1.7), `getMergedAreasOrNullObject` (1.13), `getSpillingToRange` (1.12), `getSpillParentOrNullObject`.
- [`Excel.Worksheet`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-1.21) / [`WorksheetProtection`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetprotection?view=excel-js-1.21) — `protection.protected` (1.2), the refused lifecycle APIs.
- [`Excel.Workbook`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.21) — `readOnly` (1.8), `isDirty` (1.9), `save`/`close`.
- [`Excel.PivotLayout`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.pivotlayout?view=excel-js-1.21) — `getRange` (1.8), used to block pivot output.
- [`Excel.Table`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table?view=excel-js-1.21) — `getHeaderRowRange`, `getDataBodyRange`, `getTotalRowRange`.
- [`Excel.RangeAreas`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeareas?view=excel-js-1.21) — `select()` (1.18).
- [`Office.Document`](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=excel-js-1.21) — `getFileAsync`, considered and rejected as a rollback mechanism.
- [Capability envelope fact sheet](../research/01-officejs-capability-envelope.md) — the undo/API facts, the non-undoable list, and the unknowns.
- [Pane sideload spike](../research/06-pane-sideload-spike.md) — partial apply across `sync()`, `InvalidOperationInCellEditMode`, hidden-sheet selection.
- [Sidebar UX and visual language](08-sidebar-ux.md) — C28–C38, §6.2, the copy and state contract.
- [Claude for Excel teardown](../research/04-claude-sidebar-teardown.md) — highlight/comments, risky-function list, `Claude Log` sheet, overwrite protection.
