# Add-in test strategy

Ticket: [#15 · Decide the add-in test strategy](https://github.com/mustaphaelou/excel-agent/issues/15)
Decided: 2026-11-08
Builds on: [ADR 0001](../../adr/0001-runtime-topology.md), the [#6 pane spike](../research/06-pane-sideload-spike.md), [09-edit-safety.md](09-edit-safety.md), [12-agent-protocol.md](12-agent-protocol.md), and [14-handoff.md](14-handoff.md).

The add-in is tested in three layers: **unit and contract tests** in one Vitest run, **harness integration** that drives the real agent loop and approval gate against a hand-written fake Excel, and a **real-Excel smoke** on the installed build at phase gates. The gateway is never live in the deterministic suite; recordings are the contract, and one opt-in live check watches for gateway drift.

Two things are deliberately *not* tested by a runner: Office.js and WebView2 semantics (they only exist in Excel) and model answer quality (unscheduled, per the map). Everything below says what replaces them.

---

## Summary

- **Vitest 5** runs the whole deterministic suite on Node 22.12+/24, in two projects: `app` (happy-dom, React Testing Library) and `node` (server, shared, integration).
- One root `package.json`, no workspaces. Unit tests are co-located; `tests/` holds only cross-boundary suites, the harness and fixtures.
- A **hand-written narrow fake `Excel.run`** over an in-memory workbook, with a journal and a snapshot. No `office-addin-mock`, no production port.
- A write is proven absent at three levels: **FSM state**, **journal**, **snapshot**. A staged write must leave the workbook byte-identical.
- Gateway fixtures are **recorded raw SSE** plus a hand-reviewed normalized golden. Pane fixtures are normalized event lists. An opt-in live test is the only thing that contacts the real gateway.
- The smoke **launches** the installed build through `office-addin-debugging` and **asserts** through Playwright over CDP; the human drives the chat and the Apply click. Pass is a fixed must-pass list, not a vibe.
- **GitHub Actions** runs typecheck, lint and the deterministic suite. The smoke never enters CI, because CI cannot see Excel.
- The gate is a **named list of risky modules** with no global coverage percentage. A coverage report is produced and read, not enforced.

### Decisions at a glance

1. Vitest 5, one root config, two projects, one `npm test`.
2. One root package, no workspaces; co-located unit tests; `tests/` for boundaries.
3. Hand-written narrow fake `Excel.run` (+ the `Office` globals); no `office-addin-mock`, no `ExcelHost` port.
4. Three-level absence assertion: FSM state, journal, snapshot.
5. Recorded raw gateway SSE + reviewed goldens; two fixture families; live check is opt-in.
6. Scripted smoke launch and DOM/Office assertions; human owns chat, review and Ctrl+Z.
7. Smoke runs on the **installed build**, after `deploy.ps1`, with cache cleared and version asserted.
8. GitHub Actions for typecheck + lint + unit/integration; local pre-push hook runs the same `verify`.
9. Named mandatory suites; coverage report with no threshold.
10. Every untested area has a named manual check or an accepted-risk line — never silence.

---

## 1. Layers

| Layer | What it proves | Runs | Time | Blocks |
| --- | --- | --- | --- | --- |
| Unit and contract | Pure logic, schemas, parsers, FSM, store, URL policy, component behaviour | Every change, CI | < 20 s | Merge |
| Harness integration | The real controller + real adapters against the fake Excel and scripted streams; `/api/chat` over a real socket | Every change, CI | < 20 s | Merge |
| Live gateway | Gateway auth, model catalog, SSE transport, one tool call | Opt-in, before phase gates | < 60 s | Nothing automatic |
| Real-Excel smoke | Install, manifest, cert, Office.js, WebView2, one read, one write, one citation, Stop, audit | Phase gates 1–6, cutover | 10–15 min | Phase gate |

The first two share one command. The last two are deliberately outside it: one needs the internet, the other needs Excel.

---

## 2. Layout

```text
package.json                 one root package; devDependencies only
package-lock.json            committed
vitest.config.ts             two projects + coverage
tsconfig.json                app, shared, tests; server stays JS
eslint.config.js             flat config
.githooks/pre-push           runs npm run verify (core.hooksPath)

app/
  src/**/*.test.tsx          co-located component and unit tests
server/
  src/**/*.test.mjs          co-located unit tests
shared/
  tools.json                 generated from docs/.../12-tools.json; drift-tested
  errors.json                code -> retryable -> copy key
  protocol.d.ts              types only; the server stays plain ESM JS

tests/
  harness/
    fake-excel.ts            Excel.run, sync, proxies, collections, events
    fake-office.ts           Office.onReady, requirements, settings, taskpane
    workbook.ts              in-memory sheets/cells/formats/tables/pivots
    journal.ts               read/write/select/apply entries
    assertions.ts            expectNoWrites, expectSnapshot, expectJournal
    stream.ts                toSse(events) for pane-side fixtures
    README.md                capability matrix
  integration/
    turn.test.ts             loop, tool results, step limits, cancellation
    approval.test.ts         staging -> review -> apply -> partial -> retry
    chat-http.test.ts        real server socket, SSE framing, abort
  fixtures/
    gateway/**/*.sse         raw upstream frames, provenance header
    gateway/**/*.expected.json  normalized event golden, reviewed
    gateway/**/*.http.json   pre-stream errors (401/429/400)
    streams/*.json           normalized event lists for the pane
  smoke/
    setup.ps1                copies the probe and package.json to the smoke tool dir, npm ci
    run.ps1                  launch, cache clear, registry snapshot/restore, teardown
    probe.mjs                CDP assertions and evidence capture
    package.json             office-addin-debugging + playwright-core
    CHECKLIST.md             living must-pass list
    runs/<date>-phase<n>/    record.md, console.txt, logs, screenshots
  setup/
    app.ts                   jest-dom, RTL cleanup, matchMedia stub
    node.ts                  network guard, fake timers helper, temp dirs
    network.ts               throws on unmocked fetch; loopback opt-in

.github/workflows/verify.yml typecheck + lint + tests on Node 24
```

Rules:

- `app/src/**` and `server/src/**` co-locate `*.test.tsx` / `*.test.mjs`. A test moves with the module it guards.
- `tests/` is only for things that cross a boundary: the controller + fake, the HTTP surface, the fixtures, the smoke.
- `tests/harness/` is not collected by Vitest; it is imported.
- No workspaces: the server has zero runtime dependencies, nothing is published, and one `engines` field keeps WSL and Windows honest.
- The install tree in `%LOCALAPPDATA%\ExcelAgent\` stays free of `node_modules`. The smoke's `office-addin-debugging` and Playwright live in `%LOCALAPPDATA%\ExcelAgentSmoke\`, which is tooling, not the install.

The durable record is this file. `tests/smoke/CHECKLIST.md` is the living checklist the build edits as phases land; `tests/README.md` says how to run each command.

---

## 3. Runner and config

**Vitest 5** (`^22.12.0 || ^24.0.0`) on the pane's Vite pipeline. Two projects in one config:

```ts
// vitest.config.ts (shape)
test: {
  projects: [
    { test: { name: "app",  environment: "happy-dom",
              include: ["app/src/**/*.test.{ts,tsx}"],
              setupFiles: ["tests/setup/app.ts"] } },
    { test: { name: "node", environment: "node",
              include: ["server/src/**/*.test.mjs", "shared/**/*.test.ts",
                        "tests/integration/**/*.test.ts"],
              setupFiles: ["tests/setup/node.ts"] } },
  ],
  coverage: { provider: "v8", include: ["app/src/**", "server/src/**", "shared/**"],
              reporter: ["text", "html"] },
  allowOnly: !process.env.CI,   // CI fails on .only
}
```

House rules, enforced by `tests/setup/node.ts` and reviewed:

- **No unmocked `fetch`.** The global is replaced with a stub that throws unless the URL is loopback and the test called `allowLoopback()`. The one real-socket integration test opts in; nothing else can reach the internet.
- **No wall clock.** Time-sensitive tests call `vi.setSystemTime(...)` with a fixed instant; no test reads `Date.now()` for a fixture value.
- **No shared state.** Each test builds its own fake workbook, store directory and controller; no module-level mutable singletons. Files may run in parallel.
- **Temp stores only.** History and audit tests use `mkdtemp` under the OS temp dir and remove it in `afterEach`.
- **No import-time Office access.** Pane modules read `Office`/`Excel` inside functions, never at module scope, so tests can assign `globalThis.Excel`/`globalThis.Office` before driving them.
- **No `.only`.** CI sets `allowOnly: false`.

`tsc --noEmit` covers `app/`, `shared/` and the tests. The server is plain JS with JSDoc and is not type-checked; its safety comes from the tests and lint. That is a deliberate cost of ADR 0001's zero-dependency server.

Scripts:

| Script | Does | Target |
| --- | --- | --- |
| `npm test` | `vitest run` | < 20 s |
| `npm run test:watch` | `vitest` in watch mode | — |
| `npm run verify` | `tsc --noEmit && eslint . && vitest run --coverage` | < 60 s |
| `npm run test:live` | real gateway: models, one stream, one tool call | < 60 s |
| `npm run fixtures:record` | refresh `tests/fixtures/gateway/` from the gateway | — |
| `npm run sync:catalog` | regenerate `shared/tools.json` from the #12 record | — |
| `.\tests\smoke\run.ps1 -Phase <n>` | installed-build smoke (Windows) | 10–15 min |

---

## 4. The fake Excel and Office harness

### 4.1 Why hand-written

`office-addin-mock` 4.0.1 emulates `load`/`sync` and the "property not loaded" error well, but Microsoft documents that it **does not support collection-type objects** (`WorksheetCollection`, `TableCollection`, …), and it does not provide `Excel.run` — the test adds that itself. The pane's inventory needs exactly those collections, and the approval gate needs workbook semantics (pre-state, drift, partial apply, cell-edit) rather than proxy semantics. A hand-written fake over an in-memory workbook models the semantics that matter, and the capability matrix keeps it honest.

No production port is introduced. Tests assign `globalThis.Excel` and `globalThis.Office`; the adapter reads the globals at call time. The code under test is the code that ships.

### 4.2 Surface

| API | Fake behaviour |
| --- | --- |
| `Excel.run(options, cb)` | Creates a context, runs `cb`, returns its result; records `options` (`mergeUndoGroup`, `delayForCellEdit`) but does not honour timing |
| `context.sync()` | Commits queued writes in call order; resolves loaded properties; throws an injected error instead |
| `range.load("values,address")` | Records requested properties; reading before `sync` throws `PropertyNotLoaded` like Office |
| `range.values/formulas/numberFormat` | Assignments queue; reads return committed state after sync |
| `worksheet.getRange/getUsedRangeOrNullObject/getRanges` | Resolve against the in-memory sheet; null objects carry `isNullObject` |
| `worksheets`, `tables`, `pivotTables`, `conditionalFormats`, `dataValidation` | Collection surfaces with the members the adapter calls |
| `range.select()`, `worksheet.activate()` | Queued and applied at sync; journaled |
| `worksheet.onChanged` | Fires for committed writes (`source: "apply"`) and `externalEdit()` (`source: "external"`) |
| `range.getSpillingToRangeOrNullObject()` | Null object unless `simulateSpill()` was configured |
| `Office.onReady` | Resolves `{ host: "Excel", platform: "PC" }` |
| `Office.context.requirements.isSetSupported` | True for the four sets in #6; overridable to simulate an older build |
| `Office.context.document.settings` | In-memory get/set for the auto-open setting |
| `Office.extensionLifeCycle.taskpane.setWidth` | Recorded, not applied |

Anything the adapter calls that the fake does not model **throws a loud "unmodelled API" error**. A silently `undefined` return would let a test pass while production breaks.

### 4.3 Journal, snapshot, injection

```ts
type JournalEntry =
  | { kind: "read";   sheet: string; address: string }
  | { kind: "write";  sheet: string; address: string; before: Cell[][]; after: Cell[][]; sync: number }
  | { kind: "select"; sheet: string; address: string; sync: number }
  | { kind: "apply";  rows: number; sync: number };

interface FakeExcel {
  Excel: typeof Excel;
  Office: typeof Office;
  workbook: FakeWorkbook;                 // direct seeding for tests
  journal: Journal;
  snapshot(): WorkbookSnapshot;           // deterministic, sorted, deep-cloned
  failSyncAt(n: number, error: FakeExcelError): void;
  throwOn(call: string, error: FakeExcelError): void;
  enterCellEdit(): void;                  // InvalidOperationInCellEditMode
  setReadOnly(value: boolean): void;
  protectSheet(name: string): void;
  simulateSpill(ref: string, target: string): void;
  externalEdit(ref: string, values: unknown[][]): void;
  requirements(map: Record<string, boolean>): void;
}
```

`sync()` follows #6's observed semantics: a failed sync commits nothing **from that sync**, earlier syncs stay applied, and the failure is the injected Excel error code. That is what makes the partial-apply tests real rather than theatrical.

### 4.4 Capability matrix

Written in `tests/harness/README.md` and updated with each tool family:

| Area | Modelled | Not modelled |
| --- | --- | --- |
| Cells | values, formulas (as text, not evaluated), number formats, used range | recalculation, `text` derived from formats, rich values |
| Structure | insert/delete cells with shifting, add sheet | delete/rename/hide sheet (refused by #9 anyway) |
| Tables | create, resize, totals, header detection | table styles, structured-reference semantics |
| Pivots | create/update/refresh/delete, field metadata | OLAP, slicers, layout rendering |
| CF / DV / hyperlinks | rule payloads, ranges, clearing | rendering, priority reordering semantics |
| Sort / filter | reorders and filters the in-memory grid | locale collation, dynamic arrays |
| Protection, merged cells, spill, cell edit | injectable errors only | their real API behaviour |
| Undo | not modelled | `mergeUndoGroup`, Ctrl+Z |
| Events | onChanged for writes and external edits | batching, debouncing, cross-window events |

Anything in the right column is smoke-only by definition. The fake models *our* contract with Excel, not Excel.

### 4.5 Maintenance rule

If the `load`/`sync` emulation grows past roughly 150 lines, stop and re-evaluate `office-addin-mock` for the adapter's primitive tests. The fake exists to keep the gate honest, not to become a second Office.js.

---

## 5. Proving a write did not land

Three assertions, strongest first. `expectNoWrites(fake, before)` in `tests/harness/assertions.ts` runs all three.

```ts
const fake = createFakeExcel(seed);
globalThis.Excel = fake.Excel;
globalThis.Office = fake.Office;
const session = createSession({ transport: scriptedStream("set-values") });
const before = fake.snapshot();

await session.send("Set Data!B2 to 99");              // the model streams a set_values call

expect(session.proposal.state).toBe("proposed");      // 1. the gate held
expect(fake.journal.writes()).toEqual([]);            // 2. nothing reached the host
expect(fake.snapshot()).toEqual(before);              // 3. the workbook is untouched
```

`Journal` exposes `entries` plus the convenience filters `writes()`, `applies()` and `selects()`; `expectNoWrites` is the one helper tests should call.

- **FSM state** is the real contract: the proposal exists, is reviewable, and the apply path was never entered. Component tests add the UI half — Apply is disabled until an acknowledgement, and clicking a disabled control calls nothing.
- **Journal** catches a write that was queued but not committed, and distinguishes a read from a write when the workbook state happens to coincide.
- **Snapshot** is the backstop: it catches any mutation that bypassed the journal, including a direct assignment in the adapter.

The cases that must use it:

| Case | Assertion |
| --- | --- |
| Staged write (I4) | Snapshot equals pre-state; one or more proposal rows exist; journal has no write |
| Blocked row (formula tier, protection, merged, spill child) | No row is selectable; no write; the tool result carries the code |
| Overwrite/loss without acknowledgement | Apply disabled; no write even if the handler is invoked |
| Discard | Proposal closes; snapshot equals pre-state |
| No-op / all-unchanged | No card; no write |
| Apply, all rows | Snapshot equals pre-state with every row applied; journal has one apply with N writes |
| Apply, row 3 fails | Snapshot equals pre + rows 1–2; journal sync indexes are 1,2; FSM is `partial`; Retry touches only rows 3..N |
| Drift before apply | Snapshot equals pre-state; FSM goes `stale`, not `applied` |
| Interrupted apply | FSM is `interrupted`; verification compares state to pre/post, never retries automatically |

The fake's journal is also how tests assert the *positive*: one `apply` entry, one sync per row, no selection changes.

---

## 6. Gateway faking

### 6.1 Two fixture families

**Raw gateway frames** (`tests/fixtures/gateway/**/*.sse`) feed the server adapter. Each file starts with a provenance header:

```text
# scenario: parallel tool calls; null id/name on continuation frames
# model: deepseek-v4.1-flash
# captured: 2026-11-08
# provenance: recorded (sanitized: messages, ids and headers replaced)
# expected: parallel-tool-calls.expected.json
```

The matching `*.expected.json` is the normalized event sequence the adapter must emit, with ids and timestamps stripped by a normalizer. It is written only on `--update` and reviewed like code; it is never computed live from the parser under test, or the test proves nothing.

Recorded scenarios: basic answer, parallel tool calls, null continuations, `reasoning_content`, usage-only chunk, `[DONE]`, a fallback-shaped 400, a stream cut mid-frame. Authored edge cases: CRLF, comments, multi-line `data:`, malformed arguments, empty deltas, extra choices. Pre-stream HTTP failures live in `*.http.json` (status, headers, body).

**Normalized pane streams** (`tests/fixtures/streams/*.json`) feed the pane client. They are hand-authored arrays of the #12 §3.3 events; `toSse(events)` in the harness encodes them, so the pane's frame decoder is still exercised. Framing edge cases (CRLF, comments, multi-line data) are inline strings in the client test, not another fixture family.

The two families never share a golden: one tests the server's normalization of the gateway, the other tests the pane's consumption of the normalized contract.

### 6.2 Refresh and the live check

- `npm run fixtures:record -- --model deepseek-v4.1-flash` calls the real gateway, writes raw frames, strips the key, headers and prompt text, and drafts the expected file. The diff is reviewed and committed deliberately.
- `npm run test:live` starts the server on an ephemeral loopback port with the real gateway config and checks three things: `/api/models` returns a catalog with the default model and sane limits; one fixed prompt streams to `done`; one prompt asking for `read_range` produces a `tool_call` with valid arguments. It prints latency and the chosen model. Without a key it skips with a clear message.
- The first two are hard assertions. The tool-call check is **reported**, not failed, unless `--strict` is passed: a model that will not call a tool is a fidelity result, and the map keeps model benchmarking outside this ticket. The fallback chain is exercised by unit tests, not by the live model.

---

## 7. Mandatory suites

These modules get tests before they are considered done. The list is the gate; the percentage is not.

| # | Module | Layer | What breaks without it |
| --- | --- | --- | --- |
| 1 | Tool catalog contract | node | A tool the pane cannot execute, a schema the model cannot satisfy, a renamed tool, prompt/catalog drift |
| 2 | Pane schema + semantic validators | unit | A malformed call stages a write; caps, overlap, read-only and no-op stop being enforced |
| 3 | Formula scanner | unit | A blocked `CALL`/`EXEC`/DDE formula reaches the workbook, or every `HYPERLINK` is refused |
| 4 | Risk classifier + diff/pre-state | unit | Overwrites and losses apply without acknowledgement; excluded rows are misreported |
| 5 | Citation parser + revision | unit | A citation navigates the wrong sheet, a hidden sheet is unhidden, a stale read is presented as current |
| 6 | Approval FSM | unit | Apply/Discard/Retry/Interrupted states leak writes; the composer gates wrongly |
| 7 | No-write invariants | integration | A "staged" write lands, or a failed apply reports rows it never wrote |
| 8 | Server SSE adapter | unit | Null continuations, usage chunks and final frames corrupt the event stream |
| 9 | Retry and fallback | unit | A retry after partial text replays a turn; fallback crosses the 3-model cap or ignores 401 |
| 10 | Pane SSE client | unit | A seq gap is ignored; a dropped stream auto-reconnects and doubles a step |
| 11 | History store | unit | A torn JSON file, an `EPERM` rename or a path-traversal id loses or corrupts history |
| 12 | URL and SSRF policy | unit | A loopback/metadata host, a redirect or a DNS rebind gets fetched |
| 13 | Error catalog + copy | unit | A code reaches the UI with no string, or a retryable error offers no retry |
| 14 | Gating components | component | The composer sends during a proposal; Apply fires without an acknowledgement; a citation chip is not a button |

Notes per suite:

- **Catalog contract** — `shared/tools.json` is generated from `docs/wayfinder/design/12-tools.json` (`npm run sync:catalog`); a test fails if they differ. It also asserts 29 unique snake_case names, meta-schema validity, one handler + validator + risk class + family per tool, and that the system-prompt template names only catalog tools.
- **Validators** — table-driven per family, asserting the #12 §8.2 order: shape, then sheet/address, caps, read-only, overlap, formula tier, no-op.
- **Formula scanner** — the #9 §9.1 tiers, function names inside string literals, `_xlfn.`/`_xlws.` prefixes, DDE pipes, external references versus `Table1[Column]`.
- **FSM** — every transition in #9 §6, including `stale`/Update diff, `partial`/Retry eligibility, `dismissed` versus `discarded`, and `interrupted` verification.
- **History store** — atomic `.tmp` + rename, retry on `EPERM`/`EBUSY`, corrupt-file recovery, retention accounting, audit append.
- **URL policy** — scheme, port, userinfo, IPv4 and IPv6 private/link-local/metadata, `.local`, DNS pinning, max five redirects, allowlist, content type and size.
- **Gating components** — React Testing Library plus `jest-dom`, no snapshots of markup; assert roles, states and callbacks.

Not mandatory: presentational components, token tables, copy-only changes, and anything whose only observable is a screenshot.

---

## 8. Real-Excel smoke

### 8.1 Preconditions

The smoke is a **Windows** ritual against the **installed build**:

1. Excel is closed. The script refuses to start otherwise; it never kills a user session.
2. The backend is the installed Task Scheduler service, not a dev server, and `/api/health` reports the same git short SHA that `deploy.ps1` wrote (the pane build exposes the same stamp).
3. The normal build + `deploy.ps1` has run. There is no dev-server smoke path.
4. The Office WebView2 cache is cleared while Excel is closed — remove the `CustomUIValidationCache` entries and move `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` aside, as the #6 setup did. Registry sideloads survive.
5. The `WEF\Developer\{Id}` value is exported before launch.
6. The smoke workbook is the `office-addin-debugging` generated document, whose `taskpanes.xml` has `visibility="1"`, so the pane opens without a ribbon click. The probe seeds `Data!A1:C3` through `Excel.run` from the page before the first turn.

### 8.2 Launch and assert

```powershell
# run.ps1, after the preconditions
$smoke = "$env:LOCALAPPDATA\ExcelAgentSmoke"   # setup.ps1 npm-ci's the probe here
Push-Location $smoke                            # the CLI needs a package.json in the cwd
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=$port"
& "$smoke\node_modules\.bin\office-addin-debugging.cmd" start "$env:LOCALAPPDATA\ExcelAgent\manifest.xml" --no-debug
node "$smoke\probe.mjs" --port $port --phase 3
& "$smoke\node_modules\.bin\office-addin-debugging.cmd" stop "$env:LOCALAPPDATA\ExcelAgent\manifest.xml"
Pop-Location
```

`office-addin-debugging` needs a `package.json` in the working directory (the #6 finding); the smoke tool dir has one, with no add-in config, so the CLI only registers and opens the document. `stop` **removes the add-in registration**, so the script restores the exported value and asserts it is back — otherwise the installed add-in stops loading after a smoke run.

WebView2 honours `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`, and Microsoft documents the same variable for add-in debugging; Playwright connects with `chromium.connectOverCDP`. Both the CDP attach and the pane's auto-open are **phase-1 spikes**. If either fails, the fallback is the same checklist run by hand with DevTools, and the script still launches and captures evidence.

The probe needs stable selectors: `data-testid` on the composer, send/stop, model chip, edit card, apply/discard, citation chip, error banner and status line. That is the one production affordance this strategy asks for; it is also useful to accessibility and automation.

### 8.3 Must-pass checklist

| # | Check | Who | Evidence |
| --- | --- | --- | --- |
| 1 | Pane loads over HTTPS, no certificate interstitial | script | screenshot, page URL |
| 2 | Health stamp matches the deployed build | script | `/api/health` JSON |
| 3 | Empty state, composer enabled, model chip lists `/api/models` | script | screenshot |
| 4 | One read streams to completion and cites `Data!A1:C3` | script + human | screenshot, transcript |
| 5 | Clicking the citation selects the cited range | script | selection address via `Excel.run` |
| 6 | One write shows rows, risk/overwrite notes; Apply lands it | human reviews and applies; script asserts | screenshot, cell value |
| 7 | Ctrl+Z restores the pre-apply cell | human | screenshot |
| 8 | Stop mid-stream marks the turn stopped; no write is applied and a staged proposal is frozen | script | snapshot, screenshot |
| 9 | Backend down shows the banner; restart recovers | script | screenshot, logs |
| 10 | Reopen restores the transcript and pending proposal (phase 5+) | human | screenshot |
| 11 | `audit.jsonl` has one apply entry with the expected counts | script | log tail |
| 12 | `server.log` has no error or stack trace; console is clean | script | logs |
| 13 | Add-in registration is restored after `stop` | script | registry export |

### 8.4 Pass rule, evidence, cadence

- **Pass** = every must-pass check green on the installed build. Any red blocks the phase gate.
- **Flake policy** = one immediate re-run is allowed with a note in the record. A check that flakes twice across runs opens an issue; it is not silently tolerated.
- **Evidence** = `record.md` (checklist, outcome, model, build stamp), `console.txt`, `server.log` tail, `audit.jsonl` tail, and four screenshots: empty, review card, applied, citation. The markdown and text logs are committed under `tests/smoke/runs/`; screenshots are committed only when they explain a decision, as in #6.
- **Cadence** = phase gates 1–6 and cutover, plus any deploy that changes the manifest, certificate, adapter or agent loop. Not per commit.

---

## 9. CI and local gates

`.github/workflows/verify.yml`:

```yaml
name: verify
on: [push, pull_request]
concurrency:
  group: verify-${{ github.ref }}
  cancel-in-progress: true
jobs:
  verify:
    runs-on: ubuntu-latest
    timeout-minutes: 10
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with: { node-version: 24, cache: npm }
      - run: npm ci
      - run: npm run verify
```

- The repo is public, Actions is enabled, and the suite is seconds long. No secrets are needed: the live test is not run.
- **What CI cannot validate:** Office.js, WebView2, the certificate, the registry entry, Task Scheduler, Windows paths, the pane's real rendering. That list is in a comment at the top of the workflow so nobody mistakes green for shippable.
- Locally, `.githooks/pre-push` runs `npm run verify` (enable with `git config core.hooksPath .githooks` in phase 1). `npm test` is the TDD loop.
- WSL runs Node 22.23 and Windows runs Node 24.15 at decision time; `engines: ">=22.12"` and CI on Node 24 keep both honest. A test that needs a 24-only feature is a bug.
- **Definition of done for a change:** the owning suite has a test for the behaviour; `npm run verify` is green; no `.only`; any fixture change is deliberate and noted; the smoke runs at the phase gate, not per change.

---

## 10. Deliberately untested in v1

| Not tested | Covered instead |
| --- | --- |
| Excel API behaviour beyond the smoke cases | One smoke case per tool family at phase 4; real Excel only |
| Undo stack and `mergeUndoGroup` internals | One human Ctrl+Z check per write smoke |
| Model answer quality and tool-call fidelity per family | `test:live` catches transport; quality stays the map's unscheduled benchmark |
| WebView2 certificate, cache and version quirks | Install checklist plus the smoke's no-interstitial and cache-clear steps |
| Installer, Task Scheduler, UAC, uninstall, rollback | Cutover checklist, run once |
| Layout at 260/380/520/720, dark mode, 200 % zoom, IME | #13 evidence plus manual review; no visual-regression suite |
| Full accessibility and screen readers | `axe` on key components only; Lighthouse stays manual |
| Performance budgets | Measured at phase 6, not asserted |
| Multi-pane, multi-window, hosted backend | Out of scope by ADR 0001 |
| Provider-side cancellation, 1904 dates, manual calculation, locales | Accepted risk, already recorded in #10/#12; en-US only |

New gaps must arrive with a named manual check or an accepted-risk line. Silence is not a test decision.

---

## 11. Gates in the build sequence

Mapping to [#14's phases](14-handoff.md#build-sequence):

| Phase | Test work | Gate |
| --- | --- | --- |
| 1. Harness and skeleton | Vitest projects, network guard, fake skeleton, catalog contract, CI workflow, pre-push hook, smoke `run.ps1` skeleton; prove CDP attach and pane auto-open | `npm run verify` green; blank pane smoke passes |
| 2. Read-only slice | SSE adapter + client, citation parser, revision, `read_url` policy, read tools; live check | Read, citation jump and SSE pass the smoke |
| 3. Write gate | Validators, scanner, risk, FSM, no-write invariants, apply/partial/retry | Stage → review → apply → Ctrl+Z; partial reported honestly |
| 4. Native operations | One validator + fake family per tool family; one smoke case each | Each family has a smoke case and a risk class |
| 5. Persistence and #16 | Store atomicity/recovery, resume, proposal restore, retention, deletion/export | Reopen resumes; backend-down degrades; read-only refuses |
| 6. Hardening and cutover | Error/offline/timeout, axe, measured budgets, full smoke matrix, rollback | Smoke matrix passes on the installed build; rollback exercised |

Do not call the build verifiable before phase 1's gate passes.

---

## 12. Risks and open spikes

1. **CDP attach to Excel's WebView2 is unproven here.** WebView2 honours the env var and Playwright connects over CDP, but Office may override arguments or the port. Phase-1 spike; fallback is the manual checklist with DevTools.
2. **`office-addin-debugging --document` auto-open is unproven.** #6 proved the generated document; a prepared workbook may not carry `taskpanes.xml visibility="1"`. Use the generated document first; a committed smoke workbook only if the spike proves it.
3. **The fake can drift from Excel.** Mitigated by the capability matrix, throw-on-unmodelled, and the smoke being the source of truth for real semantics.
4. **WSL 22.23 is not the Windows 24.15 runtime.** CI runs 24; local WSL runs 22; both must pass. Real runtime behaviour is only exercised on Windows.
5. **No Windows CI.** The smoke is the only Windows signal, and it is manual. Accepted.
6. **Fixtures go stale as the gateway changes.** `test:live` and a deliberate `fixtures:record` run before each phase gate.
7. **The smoke depends on a live model.** Fixed prompt and model limit variance; a failure is a signal, not noise. The tool-call check is reported, not gating.
8. **`office-addin-debugging stop` removes the add-in registration.** The script exports and restores the `WEF\Developer` value and asserts it.

---

## 13. Rejected alternatives

| Alternative | Why rejected |
| --- | --- |
| `node:test` for the server, Vitest for the app | Two runners, two assertion styles, no DOM for React; the only gain is one fewer dev dependency |
| Jest | A second transform stack for ESM/TS on top of Vite; slower and redundant |
| `office-addin-mock` as the harness | No collection mocking; no `Excel.run`; it tests proxy usage, not workbook semantics |
| An `ExcelHost` production port | Leaves the real adapter — the subtle code — covered only by smoke, and adds an abstraction the blueprint never asked for |
| Live model in the deterministic suite | Non-deterministic, slow, needs a key, and turns a gateway outage into a red build |
| Local stub gateway only | Cannot produce the edge cases (null continuations, usage chunks, malformed args) without rebuilding them anyway |
| Full smoke automation, including Apply and Ctrl+Z | The human review is the behaviour under test; Ctrl+Z targets Excel's native stack outside the page |
| `office-addin-test-server` | Requires test hooks and a polling channel compiled into the pane; duplicates what CDP gives |
| Manual-only smoke | #6 showed the launch is scriptable and DOM assertions are cheap; hand-reading screenshots is the expensive part |
| Global coverage threshold | Rewards testing trivial code and creates busywork; the named suite list is the real gate |
| No CI | A public repo with a free runner; local-only loses the Node 22/24 and clean-clone signals |
| npm workspaces | Nothing is published and the server has no runtime dependencies; the ceremony buys nothing |
| Playwright browser mode / Vitest browser | The browser that matters is WebView2 inside Excel; a headless Chromium would test a different renderer |

---

## 14. Sources

- [Unit testing in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/unit-testing) — `office-addin-mock`, its `load`/`sync` emulation, and the documented collection limitation.
- [`office-addin-mock` on npm](https://www.npmjs.com/package/office-addin-mock) — version and API at decision time.
- [`office-addin-debugging` README](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging) — `start`, `stop`, `--no-debug`, `--document`.
- [Test and debug Office Add-ins on a non-local server](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-non-local-server) — always run `stop` to remove registrations.
- [Playwright: WebView2](https://playwright.dev/docs/webview2) — `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`, `connectOverCDP`.
- [Debug add-ins using developer tools for Edge WebView2](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-chromium) — the same env var in Office's own debugging flow.
- [06-pane-sideload-spike.md](../research/06-pane-sideload-spike.md) — auto-open via the generated document, same-origin SSE, partial apply across `sync()`, cell-edit failure, hidden-sheet failure.
- [ADR 0001](../../adr/0001-runtime-topology.md) — install layout, one origin, pane-owned loop, zero-dependency server.
- [09-edit-safety.md](09-edit-safety.md) — approval states and the apply protocol the FSM tests encode.
- [12-agent-protocol.md](12-agent-protocol.md) — event stream, tool catalog, validation layers, retry/fallback, error catalog.
- [14-handoff.md](14-handoff.md) — the build sequence and phase gates this strategy slots into.
