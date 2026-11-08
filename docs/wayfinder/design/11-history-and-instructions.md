# Conversation, history and instructions

> Ticket [#11](https://github.com/mustaphaelou/excel-agent/issues/11). The authoritative summary is the `## Answer` comment on the issue; this file is the durable record.

## Summary

A conversation belongs to a workbook and outlives the pane. The backend owns one local store of plain JSON files under `%LOCALAPPDATA%\ExcelAgent\history\`; the pane owns nothing durable. Identity is the workbook's path — no identity marker is written into the workbook, so Save As starts a new thread and an old one is linked back by hand. The stored transcript is exactly the non-system messages sent to the gateway, tool results included; the system prompt and inventory are rebuilt on resume, so a resumed conversation is the same conversation. Instructions are one free-text field per workbook that only the user writes. Compaction is user-initiated and leaves a visible divider with the originals intact. Nothing is deleted without an explicit action.

### Decisions at a glance

| # | Decision | Why |
| --- | --- | --- |
| 1 | Identity is the canonical document URL; no ID is written into the workbook | #9 forbids unreviewed writes; a hidden file ID is one |
| 2 | The backend owns the store; the pane keeps no durable copy | The pane's origin storage dies on every Office cache clear |
| 3 | One directory per workbook, one file per conversation, one append-only audit log | Independent writers, a crash-safe audit, trivial export |
| 4 | The pane sends the URL; the server derives the directory name | A client path must never reach the filesystem |
| 5 | A conversation is workbook-bound, resumed on open, ended only by the user | Claude's "pick up where you left off" without app-scoped bleed |
| 6 | The full proposal is persisted at render and verified at review | #9 needs a trace before the first write |
| 7 | Instructions: one field per workbook, replace-only, no appends | Claude's Outlook-style appends are a smell |
| 8 | Compaction is user-initiated, behind a divider, originals kept | Silent in-place summarization is forbidden; auto-branching loses the thread |
| 9 | Soft 250 MB budget, no silent deletion, delete-all plus export | A personal tool should not eat its own record |
| 10 | Store failure degrades, never blocks the turn | Losing a turn to a disk hiccup is worse than a banner |
| 11 | Read-only is a per-workbook preference plus Excel's current `Workbook.readOnly`; the header chip is status, the Settings switch is the control | One persisted field, no workbook write, and a forced state the user cannot accidentally turn off |

## 1. Invariants

1. **One owner.** The backend owns every durable byte. The pane holds the live conversation in memory and writes it through; it keeps no second copy, so there is no divergence to reconcile.
2. **No hidden workbook state.** No file ID, no log sheet, and no settings entry except the explicit, opt-in auto-open switch ([#8 C48](08-sidebar-ux.md)). The auto-open switch is a user gesture, never agent-authored; every content write goes through #9's approval gate.
3. **A conversation is never silently rewritten.** Compaction adds; deletion is explicit; linking moves only on a user click.
4. **Reproducible turns.** The stored user/assistant/tool messages are exactly the non-system bytes sent. A resumed conversation replays them plus a fresh system prompt and inventory, not the system prompt that was live.
5. **Nothing is deleted on a timer.** Retention is a warning, never a policy.
6. **Instructions belong to the user.** No feature appends to them, and the API has no append verb.
7. **The store is not a security boundary.** It is local, single-user, and readable by anything running as this user.

## 2. Identity and binding

### 2.1 What identifies a workbook

`Office.context.document.url` — a `file:///C:/…` URL for a local file, an `https://` URL for OneDrive or SharePoint, `null` when unsaved ([#10 §2.5](10-context-and-citations.md)).

- Canonical form: percent-decoded, NFC-normalized, trimmed; lowercased for hashing, with the original string kept for display.
- The URL is the key, not the display name: `Workbook.name` changes on Save As and is not identity.
- Case-insensitivity is Windows' contract — `Forecast.XLSX` and `forecast.xlsx` are the same workbook.

### 2.2 Save As, rename, move

All three change the URL, so all three start a new conversation. The old conversation is not lost: it stays under its old key and appears in History → All workbooks with its old path. History's empty state and the workbook menu offer **Link to an old conversation…** (§2.4). There is no automatic matching — matching without an in-file ID is guesswork.

### 2.3 Unsaved workbooks

`url === null` → the conversation is memory-only. The empty state says "History starts when you save this workbook." At every turn start the pane re-reads `document.url`; the first non-null value creates the record with every turn so far. If the pane closes before that, the draft is gone — stated in the empty state, not discovered later.

### 2.4 Linking

- Entry points: History's empty state, and the workbook menu.
- A full-height picker (not a menu) titled **Link to an old conversation**, with a **Filter by path or name** field.
- Lists every conversation on the machine except the current workbook's, with its path and date, newest first, grouped by workbook, filterable by path or name.
- Linking **moves** the conversation file under the current workbook's key and appends an audit note. One-way; linking it back to the original key is the undo. No copy, no merge, no duplicate.
- The conversation records the original key, path, group and time in `linkedFrom`, so the "Move back to {workbook}" action in its History overflow works after a pane reload, not only through the immediate Undo toast.
- An unsaved workbook has no key: the picker entry is disabled with "Save the workbook first", and History says "History starts when you save this workbook." No memory-only link is offered.
- The picker is the only cross-workbook operation in v1.

### 2.5 One pane per workbook

There is no API that follows the active workbook, so identity is never in doubt inside a pane. Two workbooks are two panes with two keys. The same workbook opened from two paths is two keys — a copy is a different file.

### 2.6 Read-only preference

- `readOnly` is a per-workbook **user preference**, stored in `workbook.json` and changed through `PUT /api/history/workbooks/:key/read-only` with `{ "readOnly": boolean }`. It is never written into the workbook file.
- Effective read-only is `setting.readOnly || Workbook.readOnly`, recomputed when the workbook opens. `Workbook.readOnly` is Excel's read-only state, not a stored preference; the client does not overwrite the preference when forced.
- The header lock chip is a status indicator and Settings shortcut, not a toggle; the switch in Settings → Editing is the only control (#16 resolves the #8 C4b / #9 wording).
- Turning the mode on is disabled while a proposal is unresolved; a restored pending proposal is shown read-only and cannot be applied until the workbook is writable and the mode is off.

## 3. The store

### 3.1 Owner

The backend owns `%LOCALAPPDATA%\ExcelAgent\history\`. The pane's own origin storage was the obvious alternative — Claude's IndexedDB — and is rejected for one concrete reason: this install clears it routinely. `clear-office-cache.ps1` moves `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` aside, and that is the WebView2 cache holding the pane origin's storage; a port, certificate or hostname change orphans the origin just as completely. A history that dies during ordinary add-in troubleshooting is not a history.

Consequences, accepted: the backend is no longer stateless (ADR 0001 §10 amended); history is unavailable while the backend is down; the store must be excluded from deploy, upgrade and uninstall.

### 3.2 Layout

```
%LOCALAPPDATA%\ExcelAgent\history\
├── q3-forecast-3f2a9c1e\           ← <slug>-<8 hex of sha256(canonical url)>
│   ├── workbook.json               ← path, name, instructions, timestamps, schema v
│   ├── audit.jsonl                 ← #9's records, one JSON object per line, append-only
│   └── conversations\
│       ├── 6f1c8e2a-….json         ← one envelope per conversation
│       └── …
└── usage.json                      ← last measured size and timestamp
```

`workbook.json`:

```json
{
  "v": 1,
  "key": "q3-forecast-3f2a9c1e",
  "url": "file:///C:/Users/me/Documents/Q3 Forecast.xlsx",
  "name": "Q3 Forecast.xlsx",
  "instructions": "",
  "instructionsUpdatedAt": null,
  "readOnly": false,
  "createdAt": "2026-11-08T10:00:00Z",
  "updatedAt": "2026-11-08T10:00:00Z"
}
```

### 3.3 Directory naming and the trust boundary

The pane sends the raw URL; the server derives the name. `<slug>` is the basename without extension, lowercased, non-alphanumerics collapsed to `-`, trimmed, capped at 40 characters, `workbook` when empty. The 8-hex hash of the canonical URL disambiguates and cannot be steered by a crafted filename.

The server never joins a client string into a path. Every later call names a server-issued key and a conversation UUID validated against `^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$`. Reads are a server-side directory walk: no client-supplied glob, no `..`, no absolute path, no symlink following.

### 3.4 Writes, atomicity, concurrency

- Whole-file rewrite for `workbook.json` and each conversation; append for `audit.jsonl`.
- Write `<file>.tmp`, `fsync`, then `rename` over the target (Node maps this to `MoveFileEx` with replace, which is atomic on NTFS). A crash leaves either the old file or the new one, never half of each; a leftover `.tmp` is removed at startup.
- `EPERM` or `EBUSY` on rename (antivirus, indexer) retries twice with 20 ms backoff, then surfaces as a store failure (§3.5).
- No lock file and no cross-pane sync: one file per conversation means writers never share a file. Two panes can only collide by writing the same conversation, which would require the same workbook open twice — impossible with one pane per workbook.

### 3.5 Failure modes

| Failure | Behavior |
| --- | --- |
| Unreadable or corrupt file | Renamed `.bad`, skipped in listings; History shows "1 conversation couldn't be read. It was set aside as {file}." |
| Newer schema `v` | Refused with a message; never downgraded or rewritten |
| Older schema `v` | Migrated on read, written back on the next save |
| Disk full, permission error, backend restart | The turn continues in memory; a persistent banner says "History isn't being saved. Your conversation is still open."; retried on the next transition |
| Directory deleted while running | Recreated on the next write; only the conversation being written comes back |
| `history\` missing at startup | Recreated empty, no error |

### 3.6 What is never stored

- The workbook file itself, or any sheet the user did not put in a prompt or a tool result.
- The gateway key, the config, or anything from `config.env`.
- Full-sheet dumps: only the ranges a tool actually read (§4.2).
- Telemetry of any kind.

## 4. Conversation model

### 4.1 Lifecycle

- A conversation is created by the first user message, never by opening the pane.
- It is bound to the workbook, not the pane. Closing the pane does not end it; reopening resumes it.
- Reopening loads the most recent conversation for the key, renders its transcript, and restores the scroll position.
- **New chat** starts another conversation and leaves the old one in History. It is disabled while streaming (with a confirm) and while a proposal is unresolved, and absent on first run (#8 C2).
- Rename is inline; delete confirms and, when it is the active conversation, drops the pane to the empty state.

### 4.2 Envelope

```json
{
  "v": 1,
  "id": "6f1c8e2a-…",
  "workbookKey": "q3-forecast-3f2a9c1e",
  "title": "Why does the forecast miss Q4?",
  "createdAt": "2026-11-08T10:00:00Z",
  "updatedAt": "2026-11-08T10:04:00Z",
  "turns": [],
  "compactions": [],
  "linkedFrom": null
}
```

`linkedFrom` is present only on a conversation moved by `POST /api/history/link`: `{ "key", "url", "name", "linkedAt" }`. It is the source for "Move back to {workbook}"; it is cleared when the conversation is linked back. It is metadata, not transcript content, and is never sent to the model.

A turn:

```json
{
  "id": "t-…",
  "role": "user",
  "at": "2026-11-08T10:00:00Z",
  "model": "…",
  "status": "complete",
  "text": "Why does the forecast miss Q4?",
  "toolCalls": [
    { "name": "read_range", "args": {}, "result": {}, "status": "ok", "durationMs": 41 }
  ],
  "citations": [
    { "token": "[B22:F25](<cite:Forecast!B22:F25>)", "sheet": "Forecast", "address": "B22:F25" }
  ],
  "usage": { "promptTokens": 0, "completionTokens": 0, "estimated": false },
  "proposal": null
}
```

- `toolCalls[].result` is **exactly** what the tool returned to the model, untruncated. That is what makes a resume faithful, and it is the bulk of the store.
- `text` is the display copy. Tool results are never rendered as transcript text.
- `usage` comes from the gateway when it reports it, otherwise a `chars / 4` estimate with `estimated: true`.
- `proposal` is present only on a turn that proposed changes (§5).

### 4.3 Titles

First user message, markdown stripped, whitespace collapsed, 60 characters, ellipsis. Renameable. A model-written title is not v1 (#12 may add one later).

### 4.4 Resume

On boot: resolve the URL → if the key exists, load the newest envelope → render. Nothing goes to the model until the user sends a message. A resumed conversation sends its stored turns verbatim — minus any compacted prefix (§7.3) — plus a fresh #10 inventory.

## 5. Pending proposals across reloads

### 5.1 Persistence

#9 requires the card to exist before the first write. The proposal is part of its turn, written when the card renders and updated at each transition (`applying` → `applied` / `partial` / `discarded` / `dismissed`), with the same record mirrored into `audit.jsonl`.

```json
{
  "id": "turn-…",
  "state": "proposed",
  "rows": [
    { "i": 0, "sheet": "Forecast", "range": "B4:B12", "op": "setFormula",
      "before": "…", "after": "…", "risk": "overwrite", "included": true,
      "ack": false, "outcome": null }
  ],
  "attempts": []
}
```

### 5.2 Restore

- Restored only when its own conversation is the one resumed. A proposal never appears in another conversation.
- The card returns in the state it was written in; `applying` becomes `interrupted`.
- No verification at boot: a drift check across many ranges is slow, can fail on a locked or hidden sheet, and would spend a workbook read before the user asks for one.

### 5.3 Drift at review

Review, Retry and Apply each run #9's pre-flight check for the rows in play — per row **unchanged**, **changed**, or **unknown** (unreadable). Changed and unknown rows are excluded by default and must be re-included explicitly; nothing is retried or applied automatically. The verification outcome is appended to the audit, as #9 §7.6 requires.

## 6. Instructions

### 6.1 Scope and shape

- One free-text field per workbook, stored in `workbook.json`, 8,000-character soft limit (#8 C45).
- No global instructions and no inheritance in v1: a new workbook starts empty.
- The empty field offers one chip — **Copy instructions from {workbook}**, the most recently used other workbook that has instructions. Explicit, one click, and it replaces the field's current (empty) text.
- Saved on blur plus 500 ms debounce, with an inline "Saved" and a character counter past 6,000.

### 6.2 No appends

The API exposes one `PUT` that replaces the whole text. There is no append, patch or merge verb, so no feature can accrete into the field — the failure mode Claude shows in Outlook ([#4](../research/04-claude-sidebar-teardown.md)), where a draft-preference summary is written between marker lines.

### 6.3 Framing to the model

Sent every turn as a labelled block — `## Workbook instructions` — inside the system prompt, before the workbook inventory. #12 owns the prompt; #11 owns the rule that the block is framed as **user instructions**, never as data, never as protocol, and cannot override the citation grammar or the write rules. Empty instructions omit the block entirely.

## 7. Compaction

### 7.1 Trigger

At 80% of the selected model's context window (window size from #12's `/api/models`), the composer shows "This chat is {pct}% of the model's context. Summarize earlier turns or start a new chat." Nothing happens automatically, ever.

### 7.2 Summary record

```json
{
  "at": "2026-11-08T11:20:00Z",
  "cutTurnId": "t-…",
  "turnsCovered": 14,
  "summary": "…",
  "model": "…",
  "tokensBefore": 214000,
  "tokensAfter": 61000,
  "estimated": true
}
```

The summarizer is one model call over the turns being cut, asked for decisions, findings, unresolved questions and the ranges that matter — not a narrative recap. It runs only on a user click, with "Summarizing 14 turns…" while it does.

### 7.3 What the model sees after

`system prompt (with instructions)` + `## Earlier in this conversation` summary block + the turns after `cutTurnId` + fresh inventory + selection. The pre-cut turns are not sent again.

### 7.4 Divider and originals

- The transcript shows a divider at the cut: "Earlier 14 turns summarized · View original".
- "View original" expands the pre-cut turns read-only; nothing is hidden or deleted.
- A second compaction cuts from its own boundary and stacks a second divider.
- Originals are removed only by deleting the conversation.

### 7.5 Token accounting

`tokensBefore` uses the gateway's reported usage when available, otherwise `chars / 4`, flagged `estimated`. The pane re-evaluates at each turn start; the notice appears once per crossing, not once per turn.

## 8. Retention, export and deletion

### 8.1 Budget

- 250 MB soft budget across the store, measured by a directory walk cached for 60 s.
- At 80%: banner "History is {pct}% of its {size} budget. Manage history".
- At 100%: the same banner in danger styling. Nothing is blocked, nothing is deleted.
- Revisit only if it fires.

### 8.2 Manage history

A section of Settings (#8 C49): total size, then one row per workbook — name, path, conversation count, size, Delete. Footer: "Delete all history…" with the count and size, and "Open folder".

### 8.3 Deletion

- Delete conversation: its file, its audit lines, and its place in the list. When active, the pane drops to the empty state.
- Delete a workbook's history: the directory.
- Delete all: every directory, then an empty `history\`.
- Every deletion is confirmed with a count, none is reversible, and none touches the workbook.
- A conversation deleted by another pane lives on in that pane's memory until it writes again, at which point only that conversation is recreated. Deliberate: silently discarding an on-screen conversation is worse than a resurrected file.

### 8.4 Export

- Per conversation: `conversation.md` — the rendered transcript, tool calls reduced to one-line summaries with ranges and no payloads — and `conversation.json`, the full envelope, so it can be imported later.
- The audit exports separately as `audit.jsonl` on request, never by default.
- Import is not v1; the JSON shape is versioned so it can become import later.

## 9. Privacy

### 9.1 On disk

Conversation text, tool calls **including the cell data the model saw**, proposal diffs (`before`/`after` truncated at 500 characters per row, hashes above 32 cells — #9), the audit, and instructions. All of it is workbook-derived content in plain JSON, readable by anything running as this user.

### 9.2 Logs

`logs\server.log` records method, path, status, duration and an error id. It never records prompt text, cell values or the gateway key. The error id links a pane message to a log line.

### 9.3 Off the machine

The prompt and the workbook data inside it go to the OpenCode Go gateway — already decided by [ADR 0001](../../adr/0001-runtime-topology.md). Nothing else leaves: no telemetry, no sync, no crash reporting. History is never proxied; the history endpoints are loopback-only and are not part of the swappable surface.

### 9.4 Disclosure

- History footer: "Stored on this machine" (#8 C43).
- Settings → About: "History folder · {path}" with **Open folder** and **Copy path**.
- Nothing is disclosed inside the workbook, because nothing is written into it (§2.1).

## 10. Multi-window

- Several workbooks open → several panes → one origin → one store. Each pane writes only its own key.
- No live sync. History re-reads the directory when the view opens and when it regains focus, so a second pane's new workbook appears without a restart.
- Concurrent writes are independent because they are different files (§3.4).

## 11. Store API

Loopback-only, same-origin, JSON, all under `/api/history`. Keys are server-derived; ids are UUIDs.

| Method | Path | Purpose |
| --- | --- | --- |
| `POST` | `/api/history/resolve` | `{url}` → `{key, exists}`; never creates |
| `GET` | `/api/history/workbooks` | every workbook with counts and sizes |
| `GET` | `/api/history/workbooks/:key` | `workbook.json` plus conversation summaries |
| `PUT` | `/api/history/workbooks/:key/instructions` | replace the whole field |
| `PUT` | `/api/history/workbooks/:key/read-only` | replace the `readOnly` preference (`{readOnly: boolean}`); never changes the workbook file |
| `GET` | `/api/history/workbooks/:key/conversations/:id` | one envelope |
| `PUT` | `/api/history/workbooks/:key/conversations/:id` | write the envelope; creates on first write |
| `DELETE` | `/api/history/workbooks/:key/conversations/:id` | delete one |
| `POST` | `/api/history/link` | `{fromKey, conversationId, toKey}` → move |
| `POST` | `/api/history/workbooks/:key/audit` | append one audit record |
| `GET` | `/api/history/workbooks/:key/audit` | read the log |
| `GET` | `/api/history/usage` | `{bytes, budget, pct}` |
| `DELETE` | `/api/history/workbooks/:key` | delete one workbook's history |
| `DELETE` | `/api/history` | delete everything |
| `GET` | `/api/history/export/:key/:id?format=md\|json` | download |
| `POST` | `/api/history/reveal` | open the store folder in Explorer |

Every mutation carries the same `Origin`/`Host` checks as the rest of the API ([ADR 0001 §9](../../adr/0001-runtime-topology.md)). `reveal` shells out to `explorer.exe` on the fixed store path only, never on a client-supplied one.

## 12. Copy deck

| Context | Copy |
| --- | --- |
| Unsaved workbook, empty state | History starts when you save this workbook. |
| Link picker title | Link to an old conversation |
| Link picker empty | No other conversations yet. |
| Link picker row meta | {path} · {date} |
| Linked toast | Moved to this workbook. |
| Compaction divider | Earlier {n} turns summarized · View original |
| Compaction notice | This chat is {pct}% of the model's context. Summarize earlier turns or start a new chat. |
| Summarizing | Summarizing {n} turns… |
| Budget banner | History is {pct}% of its {size} budget. Manage history |
| Budget danger | History is over its {size} budget. Nothing is deleted automatically. Manage history |
| Unreadable file | 1 conversation couldn't be read. It was set aside as {file}. |
| History not saving | History isn't being saved. Your conversation is still open. |
| About, store | History folder · {path} · Open folder · Copy path |
| Instructions, empty-field chip | Copy instructions from {workbook} |
| Context limit (amends #8) | This chat is full. Summarize earlier turns or start a new chat. Your history is saved. |

Voice unchanged from [#8 §9](08-sidebar-ux.md): specific, calm, second person, sentence case.

## 13. Known constraints

- **Origin independence is the point.** Because the store is not origin-scoped, port, certificate and hostname changes are harmless — and because it is not, nothing in the pane can be the source of truth.
- **The backend is now a dependency for history.** Backend down → the transcript on screen survives, new turns are not persisted, and the banner says so.
- **A resume is only as faithful as its stored tool results**, which is why they are stored whole and why the store grows.
- **Compaction cannot rescue a turn that already started.** A single turn that exceeds the window is #12's error path, not compaction's.
- **No import, no cross-device sync, no search across workbooks** in v1.
- **Linking is manual.** Save As and rename produce orphans by design; the picker is the only healing path.
- **`document.url` is the only identity Office offers.** A workbook copied to another path is a different workbook.
- **The auto-open setting is the one explicit workbook setting write.** It is opt-in, user-triggered, and not history; nothing else outside #9 writes to the workbook.

## 14. Owned elsewhere

- **#12** — tool schemas, SSE events, the system prompt and the instructions block's exact framing, model catalog and context windows, token counting, error codes.
- **#13** — visual direction prototype (validated B — Claude-close).
- **#16** — prototype states for History, the link picker, compaction, the budget banner, store failure, the instructions copy chip and the read-only control; closed 2026-11-08. Surface amendment folded into #8; read-only field/API settled here.
- **#9** — approval, audit record content, drift checks, interrupted recovery.
- **#10** — inventory, reads, citations, and the revision counter (not persisted).
- **#8** — surfaces, component geometry, copy voice.
- **#7 / ADR 0001** — topology; §10 amended by this ticket.

## 15. Sources

1. [Claude data storage](https://claude.com/docs/office-agents/data-storage) — IndexedDB, user × org × app keying, 50-conversation cap, `claude.fileId`, `Claude Log`, delete-all.
2. [Claude for Excel](https://claude.com/docs/office-agents/excel) — instructions scope, auto-compaction, limitations.
3. [Claude for Excel sidebar teardown](../research/04-claude-sidebar-teardown.md) — history and instructions evidence, and the keep/adapt/reject table.
4. [Office.js capability envelope](../research/01-officejs-capability-envelope.md) — one pane per workbook, no document-close event.
5. [Workbook context and cell citations](10-context-and-citations.md) — `document.url`, inventory, revision counter.
6. [Edit safety and approval model](09-edit-safety.md) — audit shape, persisted card, interrupted recovery.
7. [ADR 0001](../../adr/0001-runtime-topology.md) — one origin, loopback API, install layout.
8. [`Office.context.document`](https://learn.microsoft.com/en-us/javascript/api/office/office.document) — `url`, `settings`, `customProperties`.
9. `.scratch/spike-06/setup/clear-office-cache.ps1` (gitignored, local archive) — the cache clear that motivates the store's owner.
