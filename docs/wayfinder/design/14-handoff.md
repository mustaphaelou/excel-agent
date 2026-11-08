# Blueprint handoff review (#14)

Ticket: [#14 · Confirm the blueprint is decision-complete](https://github.com/mustaphaelou/excel-agent/issues/14)
Reviewed: 2026-11-08
Status: **review complete · blueprint not yet decision-complete** — two gaps remain, ticketed below. No architecture reversal is implied.

The map is an index; this is the handoff, not a second copy of the decisions. The durable records are linked, not restated.

---

## Verdict

Every **v1 product axis** has a settled decision: runtime (ADR 0001), UX (#8 + the #13 amendment), safety (#9), context and citations (#10), persistence (#11), and protocol/tools (#12). The #13 prototype validated the visual direction and is folded into #8.

Two gaps block a clean "start building" handoff:

| Gap | Ticket | Why it is a blocker |
| --- | --- | --- |
| **Test strategy** | [#15](https://github.com/mustaphaelou/excel-agent/issues/15) (open) | No unit-test runner, mocked Office.js harness, SSE fixtures, real-Excel smoke test or CI gate. Implementation cannot be called verifiable without it. |
| **Missing safety/persistence surfaces** | [#16](https://github.com/mustaphaelou/excel-agent/issues/16) (open) | #11 assigns the link picker, compaction, retention, store-failure and instructions-copy states to #13, but #13 did not implement them; #8's component inventory predates #11. The read-only control is also inconsistent: #8 puts a chip in the header and sends it to Settings, #9 calls it the only control, and #11's `workbook.json`/API has no field for it. |

Neither gap changes the architecture, the safety model or the validated B direction. #15 can run in parallel with early build work; #16 must land before the persistence phase.

**What "decision-complete" now means:** close #15 and #16, record their answers, and amend #8/#11 where #16 resolves the surfaces. Then this handoff is clean.

---

## Coverage

| v1 scope | Record | Status |
| --- | --- | --- |
| Local HTTPS runtime, sideload, lifecycle, cutover | [#7 / ADR 0001](../adr/0001-runtime-topology.md) | Settled |
| Sidebar shell, visual language, states, keyboard, accessibility | [#8 / amended spec](08-sidebar-ux.md#prototype-verdict--2026-11-08) | Settled; B validated |
| Clickable visual direction | [#13](https://github.com/mustaphaelou/excel-agent/issues/13) / `prototype/clickable-sidebar-13` | Settled; throwaway, not merged |
| Formula-safe edits, approval, undo, risk, audit | [#9](https://github.com/mustaphaelou/excel-agent/issues/9) / [09-edit-safety.md](09-edit-safety.md) | Settled |
| Context inventory, paged reads, citations, locale, staleness | [#10](https://github.com/mustaphaelou/excel-agent/issues/10) / [10-context-and-citations.md](10-context-and-citations.md) | Settled |
| History store, identity, instructions, compaction, retention | [#11](https://github.com/mustaphaelou/excel-agent/issues/11) / [11-history-and-instructions.md](11-history-and-instructions.md) | Data model settled; surfaces gap #16 |
| Protocol, SSE, 29-tool catalog, validation, errors, prompt | [#12](https://github.com/mustaphaelou/excel-agent/issues/12) / [12-agent-protocol.md](12-agent-protocol.md) · [12-tools.json](12-tools.json) | Settled |
| Test strategy | [#15](https://github.com/mustaphaelou/excel-agent/issues/15) | **Open** |
| Read-only control and #11 UI surfaces | [#16](https://github.com/mustaphaelou/excel-agent/issues/16) | **Open** |

### What was exercised in #13

The prototype exercised empty, loading, streaming, answer, review, partial, applied, history, settings and error states; 260/380/520/720 widths; light/dark; mocked streaming/stop; citation-driven sheet selection; per-row inclusion and overwrite acknowledgement; apply/applied; partial failure and scoped retry; history scope, rename and delete; instructions save; model menu; auto-open; and error retry. Lighthouse accessibility scored 100 on the sampled hybrid and Claude-close review screens, Claude empty and Fluent dark.

It is not a test suite and does not count toward #15.

---

## Contradictions reconciled in this review

Commit [`d41be25`](https://github.com/mustaphaelou/excel-agent/commit/d41be25) corrects records that had drifted from settled decisions. No new decision was made.

| Record | Drift | Resolution |
| --- | --- | --- |
| ADR 0001 | Said the pane parses raw OpenAI SSE frames. | §12 made the server normalize into the pane event stream; ADR now says so. |
| #9 answer | Said the audit was IndexedDB. | The durable #9/#11 record is the append-only `audit.jsonl`; answer corrected. |
| #11 answer/doc | Said no settings entry is ever written. | The explicit, opt-in auto-open setting is the one exception; content writes still go through #9. |
| #10 doc | Included `maxCells` and called pivot fields a future tool. | #12's schema always pages at 2,000 and adds `get_pivot`; #10 updated. |
| #8 doc | Status line "step 2 of 3", optional decision card, optional Undo, conditional risk ack, stale A wireframes. | #12 made the step total unknown; no `ask_user`; #9 has no snapshot/Undo but does require ack; wireframes are labelled original A. |
| #12 doc | `GatewayAuth` copy said "Update it in Settings". | The key lives in `config.env`; copy now says update and restart. |
| Map | #12 decision and open #15/#16 were absent. | Added a #12 gist and an Open tickets section. |

**Not contradictions, just boundaries:** compaction/linking/retention/export/read-only are #16; tests are #15; model reliability remains a build spike.

---

## Unverified assumptions and risks

Facts already proved on this machine: Excel build 20527's API set, hidden-sheet selection failure, partial-apply across `sync()`, cell-edit failure, same-origin SSE through WebView2, one pane per workbook, and width reset. These are not re-opened.

Still to verify or accept:

1. **Gateway tool-call fidelity** across model families through the OpenAI-compatible endpoint; #12 defaults to `deepseek-v4.1-flash` and records fallbacks but has not benchmarked them. This is the top model risk.
2. **Windows runtime prerequisites** — Node 24 on the Windows side, Task Scheduler one-time elevation, CA trust, port 3000 cutover. The WSL checkout used for this review reports Node 22.23.4, so the build/test runtime is not the Windows runtime; verify at install.
3. **History store on NTFS under antivirus/indexer** — `.tmp` + rename atomicity, `EPERM`/`EBUSY` retries, corrupt-file recovery. Designed, not exercised.
4. **Auto-open setting** — writing `Office.AutoShowTaskpaneWithDocument` and verifying it survives save/reopen; designed, not in the #6 spike.
5. **#16 surfaces** — link picker, compaction, retention/Manage history, store-failure, instructions copy, read-only control. Data model exists; UI and read-only persistence do not.
6. **Tests** — no runner, harness, fixtures or smoke test yet (#15).
7. **Web fetch SSRF boundary** — private/link-local blocking, redirect re-checks, DNS pinning and the allowlist are specified; the production implementation is not tested.
8. **Performance budgets** — first paint, streaming frame budget, table/citation rendering at 48 characters per line; measured as targets, not on the final build.
9. **No external CI**; personal machine, manual cutover, one Node process and one port.
10. **Date/locale edge cases** — 1904 date system and manual calculation remain best-effort (#10 known constraints); localized formula scanning is #9/#12 follow-up.

**Accepted risks already explicit in the records:** one live turn per pane; closing the pane kills the turn; no programmatic undo or snapshot; non-undoable operations refused; concurrent grid edits during Apply; local history is plaintext and readable by anything running as this user; prompt or workbook content goes to the selected gateway model; no cross-workbook or background operations; no WebSocket; no auto-apply.

**Out of scope** is recorded on the [map](https://github.com/mustaphaelou/excel-agent/issues/1#out-of-scope): building the add-in in this effort, connectors/skills, distribution/auth, non-Windows hosts, VBA parity, data tables and macros.

---

## Build sequence

This is a sequence, not a schedule. The hard gates are what matter.

1. **Harness and skeleton — #15 + runtime shell.**
   Stand up the test runner, SSE fixtures, mocked `Excel.run` harness and real-Excel smoke checklist. In parallel, create the Node 24 HTTPS origin, cert, manifest, registry sideload, Task Scheduler supervisor and install layout from ADR 0001.
   **Gate:** test command runs; a blank sideloaded pane loads in real Excel.

2. **Read-only vertical slice.**
   `/api/chat` with server-normalized SSE, `/api/models`, model catalog/fallback; pane shell, transcript, composer, markdown and stop; inventory and selection; `read_range` and citations via `cite:` tokens with click-to-navigate.
   **Gate:** a real workbook question streams, renders citations and navigates; a hidden-sheet citation refuses cleanly.

3. **Write gate.**
   Canonical tool catalog + validators; proposal staging; risk classes, diffs, per-row inclusion and ack; the B edit card; apply preflight, drift, one `Excel.run`, one sync per row, partial/retry/dismiss; audit.
   **Gate:** stage → review → apply → Ctrl+Z; a partial failure is reported honestly and retried only by the user; drift fails closed.

4. **Native operations.**
   Tables, sort/AutoFilter, pivot create/update/refresh/delete, conditional formatting, data validation, hyperlinks, structure. One operation per family through the same gate.
   **Gate:** each family has a real-Excel smoke case and a risk class.

5. **Persistence and #16 surfaces.**
   History backend, resume, proposals, instructions, deletion/export; link picker, compaction, retention, store-failure and read-only control after #16 closes.
   **Gate:** pane close/reopen resumes; Save As/link works; backend-down degrades with a banner; read-only refuses writes; compaction leaves originals.

6. **Hardening and cutover.**
   Error/offline/rate-limit/auth states, Excel-busy and timeout paths, dark/narrow/200 % zoom, focus/keyboard/accessibility, performance budgets, packaging, uninstall/renew-cert, rollback to the `.xlam`.
   **Gate:** the #15 smoke matrix passes on the installed build; rollback is exercised.

Do not start phase 5 before #16 closes. Do not call the build verifiable before #15 closes. Phases 1–4 can proceed once the test harness is usable.

---

## Handoff checklist

- [x] Every v1 axis has exactly one durable decision record.
- [x] Prototype verdict is folded into #8; the old A wireframes are marked superseded.
- [x] Cross-ticket contradictions reconciled in `d41be25`.
- [x] Risks, unknowns and out-of-scope are explicit.
- [x] Build sequence and hard gates recorded above.
- [ ] [#15 test strategy](https://github.com/mustaphaelou/excel-agent/issues/15) closed.
- [ ] [#16 safety/persistence surfaces](https://github.com/mustaphaelou/excel-agent/issues/16) closed and folded into #8/#11.

**Bottom line:** the architecture and product decisions are stable. The blueprint is *not yet* decision-complete; close #15 and #16, then the build can start with a clean gate.
