# Ticket 06 — Minimal Office.js pane spike

Target: Windows 11, Excel 16.0.20527.20112 (Microsoft 365, x64), Node 24.15.0.
Worked: 2026-11-07. Throwaway code lives in `.scratch/spike-06/` (gitignored).

**Verdict: it works.** A pane served from `https://localhost:3000` loaded inside
Excel's WebView2, docked on the right, read the active range, selected `A1`, and
consumed a same-origin `fetch` + SSE stream. The risky path is proven, with four
caveats that matter to the architecture: hidden-sheet navigation throws,
`locked` is ignored, pane width is not durable, and Excel-in-cell-edit mode
fails every `Excel.run` until the edit is committed.

Evidence (all taken on this machine):

| Evidence | File |
| --- | --- |
| Pane docked, `Spike` tab, `isSetSupported`, active-range read | [01-pane-docked.png](assets/06/01-pane-docked.png) |
| Scrolled pane: hidden-sheet, partial apply, `setWidth`, SSE | [02-pane-probes.png](assets/06/02-pane-probes.png) |
| Close → reopen: width resets and cell-edit mode fails the probes | [03-reopen-cell-edit.png](assets/06/03-reopen-cell-edit.png) |
| Saved workbook with `locked="true"`: pane open | [04-locked-open.png](assets/06/04-locked-open.png) |
| Same workbook after clicking the pane's enabled close button | [05-locked-closed.png](assets/06/05-locked-closed.png) |
| Probe JSON from the generated sideload document | [probes-generated-doc.json](assets/06/probes-generated-doc.json) |
| Probe JSON from the saved/locked workbook | [probes-locked-workbook.json](assets/06/probes-locked-workbook.json) |
| Certificate + Task Scheduler + health output | [verification.txt](assets/06/verification.txt) |

The manifest is [assets/06/manifest.xml](assets/06/manifest.xml).

## What was proven

- **HTTPS works.** `https://localhost:3000/taskpane.html` loaded in WebView2
  with a locally generated CA in `CurrentUser\Root`; no warning page and no
  add-in-level cert bypass.
- **Sideload works.** The manifest was registered at
  `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, and the standard
  `office-addin-debugging` sideload document opened Excel with the pane already
  docked.
- **The pane is real Office.js.** `Office.onReady` returned
  `{host: Excel, platform: PC}`, diagnostics reported
  `16.0.20527.20112`, and `Excel.run` read `SpikeData!B2 = 2`, then
  `Range.select()` moved the selection to `SpikeData!A1`.
- **Same-origin SSE works.** A `POST` from the pane to `/api/stream` returned
  `text/event-stream; charset=utf-8`; the pane read 6 `data:` events plus a
  `done` event (630 bytes) in ~1.5 s through `response.body.getReader()`.
- **All four requirement sets are present:** ExcelApi 1.21, ExcelApiDesktop
  1.1, TaskPaneApi 1.1, SharedRuntime 1.2 (`isSetSupported` all `true`).
- **`setWidth` works.** `Office.extensionLifeCycle.taskpane.setWidth(400)`
  moved the content from 326 CSS px to 376 CSS px (the 24 px difference is
  pane chrome).
- **Partial apply is observable.** With
  `Excel.run({mergeUndoGroup: true})`, the first `context.sync()` write
  persisted; a later `context.sync()` failed with `ItemNotFound`, and the first
  write was still in `B2`. `mergeUndoGroup` groups undo records; it does not
  roll back.

## The one path that reproduced immediately

The reliable loop was:

```powershell
# Windows, elevated only once for the Task Scheduler task
cd "$env:LOCALAPPDATA\ExcelAgentSpike"
npx office-addin-debugging start manifest.xml --no-debug
```

`office-addin-debugging` does two things: it writes the same
`WEF\Developer` registry value this report describes, and it opens
`%TEMP%\Excel add-in <Id>.xlsx` — a workbook whose
`xl/webextensions/webextension1.xml` references the registry entry
(`store="developer" storeType="Registry"`) and whose
`xl/webextensions/taskpanes.xml` has `visibility="1"`. Excel then loads the
add-in and opens the pane without a click.

A plain workbook created over COM *did* load the add-in eventually — the ribbon
appeared and a report arrived about 30 s after launch — but it was not visible
10–12 s after launch, and the pane only opened when it was clicked. For a
repeatable spike, the generated document is the deterministic trigger; the
registry entry alone is necessary but not sufficient to make the UI appear
promptly.

The `office-addin-debugging` command needs a `package.json` in the working
directory; without one it fails with
`ENOENT ... package.json` before launching Excel.

## Exact setup

### 1. Certificate (long-lived, current user)

```powershell
# make-cert.ps1, run once from Windows PowerShell
$ca = New-SelfSignedCertificate `
  -Subject "CN=Excel Agent Spike Dev CA" `
  -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
  -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable `
  -KeyUsage CertSign, CRLSign, DigitalSignature `
  -TextExtension @("2.5.29.19={critical}{text}ca=1&pathlength=1") `
  -NotAfter (Get-Date).AddYears(10)

Export-Certificate -Cert $ca -FilePath ca.cer -Type CERT
Import-Certificate -FilePath ca.cer -CertStoreLocation Cert:\CurrentUser\Root

$leaf = New-SelfSignedCertificate `
  -Subject "CN=localhost" -Signer $ca `
  -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable `
  -KeyUsage DigitalSignature, KeyEncipherment `
  -TextExtension @(
    "2.5.29.17={text}DNS=localhost&IPAddress=127.0.0.1&IPAddress=::1",
    "2.5.29.37={text}1.3.6.1.5.5.7.3.1",
    "2.5.29.19={critical}{text}ca=0"
  ) `
  -NotAfter (Get-Date).AddDays(825)

$pw = ConvertTo-SecureString "spike-passphrase" -AsPlainText -Force
Export-PfxCertificate -Cert $leaf -FilePath server.pfx -Password $pw
```

Observed values:

| | |
| --- | --- |
| CA subject | `CN=Excel Agent Spike Dev CA` |
| CA thumbprint | `F49DC8FFB539D8FA73C5324C051BD5D1B252E1F6` |
| CA not after | 2036-11-07 |
| CA store | `CurrentUser\Root` |
| Leaf subject | `CN=localhost` |
| Leaf thumbprint | `F06618D4608D20CFDB9D0A65E764D713B6726315` |
| Leaf not after | 2029-02-09 (825 days) |
| Leaf SAN | `DNS=localhost`, `IP=127.0.0.1`, `IP=::1` |
| Leaf EKU | Server Authentication |
| PFX | `%LOCALAPPDATA%\ExcelAgentSpike\cert\server.pfx` |

The leaf is served by Node (PFX + passphrase). WebView2 built a chain to the
CA in `CurrentUser\Root`; Edge and Chrome headless loaded the same URL with no
interstitial, confirming the trust before Excel was involved.

The 825-day leaf is long enough for a spike but still expires; the certificate
and port-3000 bind remain the two blank-pane failure modes from ticket 03.

### 2. Server

`server/server.mjs` is Node standard library only:

```js
https.createServer({ pfx, passphrase }, handler)
  .listen(3000, "127.0.0.1");
https.createServer(options, handler)
  .listen(3000, "::1");
```

Routes:

```text
GET  /taskpane.html   static pane
GET  /styles.css      static
GET  /app.js          static
GET  /assets/*        ribbon icons
GET  /api/health      200 {"ok":true,"pid":...,"uptimeSeconds":...}
POST /api/stream      text/event-stream, 6 data events + event: done
POST /api/report      stores the pane's probe JSON for external reading
GET  /api/report      last probe JSON
```

Every response carries `Cache-Control: no-store`. The SSE route sends
`: spike sse open`, one `data:` frame every 250 ms, then `event: done`.

Start it with `%LOCALAPPDATA%\ExcelAgentSpike\run-backend.cmd`:

```bat
@echo off
cd /d "%~dp0"
"C:\Program Files\nodejs\node.exe" server\server.mjs >> logs\backend.log 2>&1
```

### 3. Manifest and registry sideload

The manifest (full copy at [assets/06/manifest.xml](assets/06/manifest.xml)):

- `<Id>7c1d2a64-5b3e-4f8a-9d20-6e4c8b1f7a35</Id>`, version `1.0.0.1`
- `SourceLocation` = `https://localhost:3000/taskpane.html`
- `IconUrl` / `HighResolutionIconUrl` = `https://localhost:3000/assets/icon-32.png` / `icon-80.png`
- `Permissions` = `ReadWriteDocument`
- `VersionOverrides 1.0` with a custom tab **Spike** and a **Spike Pane** button
  (`<Action xsi:type="ShowTaskpane">`)

Registry write:

```powershell
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key `
  -Name "7c1d2a64-5b3e-4f8a-9d20-6e4c8b1f7a35" `
  -Value "$env:LOCALAPPDATA\ExcelAgentSpike\manifest.xml" `
  -PropertyType String -Force
```

`npx office-addin-manifest validate manifest.xml` reported **The manifest is
valid.** The ID, tab, and button are deliberately a custom tab: the first
attempt put the group on `TabHome`, but the add-in was slow enough to initialize
that the Home tab showed no sign of it during the 10 s observation window. The
custom tab made the add-in's UI trivially visible once the add-in loaded.

### 4. Launch and open

```powershell
cd "$env:LOCALAPPDATA\ExcelAgentSpike"
npx office-addin-debugging start manifest.xml --no-debug
```

Excel opens `%TEMP%\Excel add-in 7c1d2a64-5b3e-4f8a-9d20-6e4c8b1f7a35.xlsx`,
loads the add-in and auto-opens the pane. This was the run captured in
[01](assets/06/01-pane-docked.png) and [02](assets/06/02-pane-probes.png).

For UIA automation, the pane's close button is exposed as
`Button "Close pane"` (`InvokePattern`); the ribbon button is not exposed while
the "Get Started" callout is on screen, so it was clicked by coordinate
(`66,171` inside the Excel window). The "Get Started" and "Using multiple
panes?" callouts are Excel onboarding, not add-in UI; click **Got it** to
dismiss.

### 5. Task Scheduler backend

Non-elevated creation, as expected from ticket 03's research:

```powershell
schtasks /create /tn "ExcelAgentSpikeBackend" `
  /tr "\"C:\Users\musta\AppData\Local\ExcelAgentSpike\run-backend.cmd\"" `
  /sc onlogon /delay 0000:10 /rl LIMITED /f
# ERROR: Access is denied.
```

One elevated prompt (UAC) was enough:

```powershell
schtasks /create /tn "ExcelAgentSpikeBackend" `
  /tr "\"C:\Users\musta\AppData\Local\ExcelAgentSpike\run-backend.cmd\"" `
  /sc onlogon /delay 0000:10 /rl LIMITED /f
# SUCCESS
schtasks /run /tn "ExcelAgentSpikeBackend"      # SUCCESS
```

Observed task state:

```text
TaskName:      \ExcelAgentSpikeBackend
Status:        Running
Schedule Type: At logon time
Task To Run:   C:\Users\musta\AppData\Local\ExcelAgentSpike\run-backend.cmd
Run As User:   musta
```

`GET /api/health` immediately answered with `pid: 3136`, `startedAt:
2026-11-07T21:38:37Z`. The task is per-user and needs no elevation to run;
only the initial create did. This matches ticket 03's recommendation and its
"admin once" caveat.

## Probe results

The pane auto-runs the probes on `Office.onReady` and POSTs the JSON to
`/api/report`. The full JSON is in
[probes-generated-doc.json](assets/06/probes-generated-doc.json); the values
below are from that run and the saved-workbook run.

### `isSetSupported`

```json
{
  "ExcelApi 1.21": true,
  "ExcelApi 1.20": true,
  "ExcelApiDesktop 1.1": true,
  "TaskPaneApi 1.1": true,
  "SharedRuntime 1.1": true,
  "SharedRuntime 1.2": true
}
```

### Active range and selection

| Probe | Result |
| --- | --- |
| `getSelectedRange()` | `SpikeData!B2`, values `[[2]]`, text `[["2"]]`, 1×1 |
| `Range.select()` on `A1` | OK; `getSelectedRange().address` = `SpikeData!A1` |
| Later run on `Sheet1!A1` with `"m"` typed | read `"m"` (1×1); the selected cell in the screenshot is `A1` |

The screenshot also shows the formula bar and the name box on `A1` after the
pane's selection, so the selection is a real Excel UI selection, not just a
proxy.

### Hidden sheet

The probe adds `SpikeHidden`, writes `A1`, hides it, then:

| Operation | Result |
| --- | --- |
| `Range.select()` on hidden sheet | **FAIL** `InvalidSelection`: "The current selection is invalid for this operation." |
| `Worksheet.activate()` on hidden sheet | **FAIL** `InvalidReference`: "This reference is invalid for the current operation." |
| Active sheet afterwards | the visible sheet (`SpikeData` / `Sheet1`) |
| Selection afterwards | unchanged (`SpikeData!A1` / `Sheet1!A1`) |

The hidden sheet was made visible and deleted at the end of the probe. A
citation into a hidden sheet therefore has to **refuse** or **unhide with an
explicit warning**; silent `select()` is not an option.

### Partial apply across `context.sync()`

```js
await Excel.run({ mergeUndoGroup: true }, async (context) => {
  const sheet = context.workbook.worksheets.getItem(originalSheet);
  sheet.getRange("B2").values = [["write-before-failure"]];
  await context.sync();                 // commits
  const missing = context.workbook.worksheets.getItem("SpikeNoSuchSheet");
  missing.getRange("A1").values = [["x"]];
  await context.sync();                 // fails: ItemNotFound
});
```

Observed: `B2` still contained `"write-before-failure"` after the failed run;
the error was `RichApi.Error / ItemNotFound` at `WorksheetCollection.getItem`.
`mergeUndoGroup: true` did **not** roll the first sync back. The screenshot
shows `write-before-failure` in `B2`.

Consequence: "Apply" must either be a single `Excel.run` with one `sync()`, or
the design must show a partial-apply state and offer a snapshot restore. There
is no transaction API and no programmatic undo.

### Same-origin SSE

The pane does not use `EventSource` (GET-only). It sends:

```js
const res = await fetch("/api/stream", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({ prompt: "spike #6" }),
});
const reader = res.body.getReader();
```

and parses `\n\n`-delimited frames.

| | |
| --- | --- |
| Status | 200 |
| Content-Type | `text/event-stream; charset=utf-8` |
| Data events | 6 |
| Done event | `{"ok":true,"events":6}` |
| Bytes | 630 |
| Elapsed | 1,544 ms |
| CORS preflight | none (same origin) |

### Pane width, dock, and `locked`

- `setWidth(400)` works at runtime: content 326 → 376 px, outer ~400 px.
- **Close → reopen does not preserve width.** After closing the pane and
  clicking **Spike Pane** again, the new page started at `innerWidth 326`;
  [03](assets/06/03-reopen-cell-edit.png) captures that state. The pane's own
  width probe then widened it again to 376.
- Saving the workbook wrote
  `xl/webextensions/taskpanes.xml` with
  `visibility="1" width="600" dockstate=""` and no `locked` attribute.
  Reopening that file auto-opened the pane (`visibility` was honored), but the
  first report measured `innerWidth 326`, so the saved `width="600"` was not
  applied, and `dockstate` was empty; the pane opened right-docked.
- **`locked` is not honored.** I edited the saved part to
  `locked="1"` and, separately, `locked="true"`, reopened the workbook, and
  checked the pane header. The `Close pane` button was present and
  `IsEnabled=True`; invoking it closed the pane both times. See
  [04](assets/06/04-locked-open.png) and [05](assets/06/05-locked-closed.png).

So: width and dock are not a reliable persistence mechanism for a sideloaded
developer add-in here; `locked` cannot replace pinning. Excel has no pane
pinning for Excel anyway (`SupportsPinning` is Outlook-only).

### Cell-edit mode

If Excel is still editing a cell when the pane runs, every `Excel.run` throws:

```text
RichApi.Error
InvalidOperationInCellEditMode
Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER
or TAB or selecting another cell, and then try again.
```

[03](assets/06/03-reopen-cell-edit.png) shows the reopened pane in this state;
after committing the edit, the probes passed. This confirms the ticket-01
`delayForCellEdit` concern: agent-initiated reads/writes need
`Excel.run({ delayForCellEdit: true })` (ExcelApi 1.20) or explicit retry
handling. The pane itself should also not fire probes on load while the user is
typing.

## Settled unknowns from ticket 01

| Unknown | Result |
| --- | --- |
| `isSetSupported` ExcelApi 1.21 | true |
| `isSetSupported` ExcelApiDesktop 1.1 | true |
| `isSetSupported` TaskPaneApi 1.1 | true |
| `isSetSupported` SharedRuntime 1.2 | true |
| Hidden-sheet `Range.select()` | throws `InvalidSelection` |
| Hidden-sheet `Worksheet.activate()` | throws `InvalidReference` |
| SSE through WebView2, same origin | works: 6 events + done, 630 bytes, 1.5 s |
| Partial apply when a later `sync()` fails | first write persists (`ItemNotFound`); no rollback |
| Pane width survives close → reopen | no; 376 → 326 (default) |
| Pane dock survives close → reopen | always right-docked; no other dock state observed |
| `locked` taskpane attribute honored | no; close button enabled and closes the pane |

## Failure modes / gotchas observed

1. **The old VBA add-in interferes.** `ExcelAgent.xlam` in
   `%APPDATA%\Microsoft\Excel\XLSTART` adds an "Assistant" group on Home and a
   modeless "Excel Agent" panel. During UIA automation a coordinate click meant
   for **Spike Pane** hit the VBA **Chat** button because the Spike tab had
   disappeared with the Office.js pane. The VBA add-in was temporarily renamed
   to `ExcelAgent.xlam.spike-bak` for a clean run and restored afterwards. The
   spike's own add-in is independent; the repo's shipped baseline is untouched.
2. **COM-launched Excel starts minimized in this environment.** UIA gave
   `-32000` coordinates until `ShowWindow(SW_MAXIMIZE)` + `SetForegroundWindow`.
   A DPI-unaware screenshot process captures the wrong region at 150 % scaling;
   call `SetProcessDPIAware()` before `GetWindowRect`/`CopyFromScreen`.
   `PrintWindow` does not reliably capture WebView2 content; `CopyFromScreen`
   on the visible window did.
3. **The pane's ribbon button is not always exposed by UIA.** The "Get Started"
   callout makes the ribbon's button tree unavailable in the snapshot. Click by
   coordinate inside the custom tab, or dismiss the callout first. The pane's
   `Close pane` button *is* exposed (`InvokePattern`).
4. **The first report POST hung on the unlogged server.** The pane had already
   posted a 1,454-byte report, then a later 3,643-byte POST never reached
   `report received`. Restarting the server with request-body logging fixed it;
   every later run posted 3,643/3,644 bytes. Treat a stuck report as a server
   lifecycle bug, not an Office.js limitation — the SSE route on the same
   origin had already completed.
5. **The generated document is the deterministic dev loop.** Registry-only
   launch eventually loaded, but with a ~30 s invisible window; the
   `office-addin-debugging` document loaded and auto-opened the pane in ~2 s.
6. **The PFX passphrase is throwaway.** `spike-passphrase` is in the server
   code by default; do not reuse this cert/server for anything but the spike.
7. **Task Scheduler create needs elevation once.** Non-elevated `schtasks
   /create` returns `ERROR: Access is denied.`; after the one elevated create,
   the task runs as the user at logon and can be started with `schtasks /run`
   without elevation.

## What this changes in the architecture

- **Citations into hidden sheets need a policy.** `Range.select()` and
  `Worksheet.activate()` both throw; do not silently call them. Refuse, or
  unhide with a visible warning (and accept the undo-stack cost).
- **`locked` is not a pinning substitute.** It was ignored for this
  sideloaded developer add-in. The pane is closable; the shared runtime's
  `Office.addin.onVisibilityModeChanged` is the way to react, not a lock.
- **Width is a preference, not state.** Don't store layout in the pane width;
  it resets on close/reopen. Use `setWidth` only as a one-shot nudge.
- **Cell-edit mode is a hard failure.** Every read/write path needs
  `delayForCellEdit: true` or a retry; agent turns that fire while the user is
  typing will otherwise fail at the first `sync()`.
- **Partial apply must be designed for.** A multi-sync apply can leave half the
  workbook changed; use one sync where possible, otherwise snapshot and offer
  restore. The pane should render "applied N of M" rather than claiming
  atomicity.
- **Dev sideload:** keep the registry entry and use
  `office-addin-debugging start manifest.xml --no-debug` (with a `package.json`)
  for a predictable auto-open; don't rely on a plain workbook and hope the
  ribbon renders.

## Cleanup state

- The old VBA add-in was restored to
  `%APPDATA%\Microsoft\Excel\XLSTART\ExcelAgent.xlam`; Excel was left closed at
  the end of the spike.
- The spike registry entry remains at
  `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\7c1d2a64-...`; remove it to
  uninstall the spike add-in.
- The Task Scheduler task `ExcelAgentSpikeBackend` remains installed and
  running (it serves the pane). Remove it with an **elevated**
  `schtasks /delete /tn "ExcelAgentSpikeBackend" /f`.
- The deploy folder `%LOCALAPPDATA%\ExcelAgentSpike` (server, cert, logs,
  screenshots) remains; the cert can be removed from
  `CurrentUser\My` / `CurrentUser\Root` by thumbprint if the spike is finished.
- The spike code and screenshots are in `.scratch/spike-06/` (gitignored); only
  this report, its assets, and the manifest are tracked.

## Sources

- Ticket 02's local HTTPS / sideload baseline:
  [`02-local-hosting-and-sideload.md`](02-local-hosting-and-sideload.md)
- Ticket 01's unknowns list:
  [`01-officejs-capability-envelope.md`](01-officejs-capability-envelope.md)
- Microsoft's dev-cert and sideload tooling:
  [`office-addin-dev-certs`](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-dev-certs),
  [`office-addin-debugging`](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging)
- WebView2 certificate behaviour:
  [ICoreWebView2_14](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/icorewebview2_14)
- Registry sideload semantics:
  [Uninstall add-ins under development](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/uninstall-add-in)
- Task panes and the taskpanes part:
  [CT_OsfTaskpane](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/555d2968-2823-4cb3-b317-4511dfc97bef),
  [Taskpane Web Extension File example](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/3d04f8ce-65f2-4dc3-bafa-636d0a7e41a1)
