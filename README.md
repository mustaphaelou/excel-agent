# Excel Agent

Chat with an LLM agent from inside Excel. It reads the open workbook, can read
web pages and APIs you link to, and can edit cells, formulas and formatting —
with your approval before anything is written.

```
Excel  ──►  "Agent" ribbon tab  ──►  chat panel (VBA UserForm)
                                        │
                                        │  workbook snapshot + your question
                                        ▼
                            http://127.0.0.1:8765  (bridge.py, WSL)
                                        │
                                        ├─ fetches links you paste (GET only)
                                        ├─ builds the prompt, calls OpenCode Go
                                        └─ returns a reply + a list of edits
                                        │
                                        ▼
                             edits queued ──► Apply / Discard
```

> **Planning:** the Claude-for-Excel-style Office.js rewrite is charted in the [wayfinder map](https://github.com/mustaphaelou/excel-agent/issues/1). Its runtime is decided in [ADR 0001](docs/adr/0001-runtime-topology.md): one local Node origin serves the pane and API, with the agent loop in the pane. Research fact sheets live in [`docs/wayfinder/`](docs/wayfinder/).

## Using it

1. Restart Excel (close every Excel window, reopen). The ribbon tab **Agent**
   appears next to *Automate*.
2. Click **Chat**.
3. Type a request, press **Enter** (Shift+Enter for a new line).

Pick the answering model from the **Model** dropdown at the top right of the
panel. It lists what the gateway currently offers, and remembers your choice.

Examples:

- *"What does this sheet contain?"*
- *"Add a Total column in D that multiplies units by price, and bold the header."*
- *"Format C2:C40 as currency with thousands separators."*
- *"Sort by revenue descending, keep the header row."*
- *"Create a Summary sheet with the total revenue per region."*
- *"Read https://example.com and put the key points in A2:A4."*
- *"Summarise https://api.github.com/repos/python/cpython and tell me the star count."*

Buttons on the ribbon: **Chat**, **Ask about selection**, **Summarise sheet**,
**Clear conversation**.

### The panel

| Control | Meaning |
|---|---|
| **Model** (top right) | Choose which model answers new messages. The list comes live from the gateway; the choice is remembered between sessions |
| **Send** | Submit (Enter does the same) |
| **Apply** | Write the proposed edits into the workbook. Failed edits stay pending so you can fix the sheet and retry |
| **Discard** | Throw the queued edits away |
| **Clear conversation** | Forget the chat history and pending edits |
| **Confirm edits before applying** | On by default: you see the edit list first. Turn it off to let edits land immediately |
| **Read-only** | Ask questions with no edits at all |

While proposed edits are waiting, **Send** is held back until you Apply or
Discard them. The proposals stay bound to the workbook that was active when
they were made, so switching workbooks cannot silently move them.

The model picker is populated from `GET /models` on the bridge, so it shows
what your key can actually use. Until the bridge answers it lists a small
built-in set; the first send starts the bridge if needed and the live list
arrives a moment later. Changing the model only affects new messages — the
conversation carries on.

### Reading web pages and APIs

Paste an `http(s)` link in your message and the bridge fetches it before the
model is called. The readable text is added to the prompt, and the panel logs
`Read <url> (… chars)` under the reply. The agent can then answer from the page
and write what it found into cells.

- GET only, public hosts only. Local/private addresses (`127.0.0.1`, `10.x`,
  `169.254.x`, `.local`) are refused, and every redirect is re-checked.
- Up to 3 links per message, 12,000 characters per page, 30,000 in total,
  12-second timeout. HTML, text, JSON, XML/CSV are readable; PDFs and images
  are not.
- Page text is sent to the model gateway as part of the prompt. Don't paste
  URLs whose contents you don't want sent to the model.
- It cannot search the web, follow links found on a page, or use a login.
  Paste the exact URL you want read. Links that only appear inside workbook
  cells are not fetched.
- Page text is treated as untrusted: the model is told to ignore instructions
  inside it, and any edits it proposes still need your approval.

To turn it off or restrict it, set these before starting the bridge (or put
hosts in `web_allowlist.txt` next to `bridge.py`):

```bash
EXCEL_AGENT_WEB=0 python3 bridge.py              # disable web reading
EXCEL_AGENT_FETCH_HOSTS=example.com,*.python.org python3 bridge.py
```

With an allowlist, only those hosts are fetched (`*.example.com` matches
subdomains). An empty list means any public host, and
`EXCEL_AGENT_FETCH_HOSTS` overrides the file when both are set. The bridge
running under Excel inherits the WSL environment, so the simplest permanent
setup is the allowlist file, one host per line:

```
# web_allowlist.txt
example.com
*.wikipedia.org
```

### Resizing the panel

The panel is a normal window: drag any edge or corner to resize it, double-click
the title bar to maximise, or use the maximise button. The layout recomputes from
the window size — the log box and text fields grow with it, the buttons stay
pinned to the bottom. It will not shrink below 380×340.

The size you leave it at is remembered for next time (`SaveSetting` under
`HKCU\Software\VB and VBA Program Settings\ExcelAgent\Panel`). Height is capped
by your screen; that cap is Windows', not the add-in's.

`modWindow.bas` adds `WS_THICKFRAME` / `WS_MAXIMIZEBOX` to the MSForms window
(`ThunderDFrame`), because VBA UserForms are created with a fixed border and MSForms
exposes no API for it. MSForms also does not reliably fire `UserForm_Resize` while
a border is being dragged, so a once-a-second `Application.OnTime` tick re-checks
the size while the panel is visible.

Edits made by the agent **cannot be undone with Ctrl+Z** — VBA writes are not
in Excel's undo stack. That is why confirmation is on by default. Save before a
big change if you are nervous.

If an edit fails after Apply, it stays pending and the panel reports how many
are left; fix the sheet and press Apply again, or Discard them.

## How the pieces work

| File | Role |
|---|---|
| `bridge.py` | Local HTTP service. Fetches linked pages, builds the prompt, calls the model gateway, validates every edit, returns a plain-text protocol. Standard library only. |
| `vba/Turn.bas` | Conversation state: history, pending edits, bridge call, apply/discard/clear |
| `vba/TurnResult.cls` | One turn's result: reply, notes, proposed-edit labels, error/status |
| `vba/modAgent.bas` | Panel host: ribbon callbacks, modeless form, layout timer, self-test |
| `vba/modContext.bas` | Turns the workbook into a snapshot (sheets, used ranges, values, selection) |
| `vba/modApply.bas` | Executes edits, re-validating every sheet name and range |
| `vba/modHttp.bas` | UTF-8 HTTP to the bridge, plus starting it if it is not running; owns the bridge script path |
| `vba/modProto.bas` | The escaping and record format shared with `bridge.py` |
| `vba/modModels.bas` | Model list, selected model, registry persistence, background refresh |
| `vba/modWindow.bas` | Win32 window styles that give the panel resize edges and a maximise button |
| `vba/frmAgent.frm.txt` | The chat panel. Controls are created at runtime and wired with `WithEvents`; renders `TurnResult` |
| `ribbon/customUI14.xml` | The Agent ribbon tab |
| `build.ps1` | Drives Excel to emit the `.xlam` from the VBA sources, and compiles it as a gate |
| `build/compilegate.ps1` | Forces a full VBA compile and fails the build on a compile error |
| `inject_ribbon.py` | Grafts the ribbon part into the `.xlam` package |
| `deploy.ps1` | Copies the add-in into `%APPDATA%\Microsoft\Excel\XLSTART` |
| `selftest.ps1` | Headless end-to-end test (real model call, real cell writes, form build) |
| `web_allowlist.txt` | Optional host allowlist for web reading, one host per line |
| `make.sh` | build → inject → deploy → test in one command |

### Design notes

- **No JSON in VBA.** VBA has no JSON parser and 64-bit Office has no
  `ScriptControl`. Both directions use a line/unit-separator protocol; JSON
  exists only inside `bridge.py`.
- **Two layers of validation.** The model proposes; `bridge.py` drops actions
  for unknown sheets and bad ops; `modApply.bas` re-checks every range before
  touching a cell.
- **The conversation is a module.** `Turn.bas` owns history, pending edits, the
  bridge call, and apply/discard/clear; `frmAgent` renders `TurnResult` and
  `modAgent` hosts the panel. Pending edits are bound to the workbook that was
  active when they were proposed, a new message waits for Apply/Discard, and a
  failed apply keeps its edits pending. Closing the panel keeps the
  conversation and any proposed edits; reopening shows how many are still
  waiting. **Clear conversation** clears both and unloads the form.
- **The agent never sees the whole workbook.** Sheets are capped at 80×26
  cells, cells at 160 characters, ~45k characters total.
- **The model list is live.** `bridge.py` asks the gateway's `/models` (cached
  for ten minutes) and exposes it as `GET /models`; the panel's dropdown is
  built from that. If the gateway or the key is unavailable, a small static
  catalog keeps the picker usable. A failed model falls over to the fallback
  list and the panel shows which one actually answered.
- **Local only.** The bridge binds `127.0.0.1` and is the only thing that
  talks to the internet. It reads the API key from the environment or
  `~/.dsh/.credentials.yaml` and never logs it.
- **Web reading is explicit and capped.** Only links written by the user in the
  current message are fetched — never links the model invents, so workbook data
  cannot be exfiltrated. Private/local addresses are refused, redirects are
  re-checked, and each page is treated as untrusted text. The model cannot
  search, and the page cannot trigger an edit without approval.
- **Loaded from XLSTART**, which is already a Trusted Location, so no macro
  security prompt and no registry entries. (An AddIns registration was tried
  first: a COM-launched Excel writes the startup entry as a relative path that
  does not resolve.)
- **VBA compiles on demand**, which hides compile errors. A declaration placed
  after a procedure (`Private gLayoutRunning As Boolean` below an `End Sub`) is
  rejected with *"Only comments may appear after End Sub, End Function, or End
  Property"* — but the add-in still loads, and the error only appears the first
  time something calls into that module, as a modal dialog that
  `Application.Run` never returns from. Excel just hangs. `build.ps1` therefore
  forces `Debug > Compile VBAProject` and fails the build if the IDE raises a
  dialog. Keep every module-level `Dim` / `Private` / `Public` above the first
  procedure.
- **The ribbon part must be `application/xml`.** Per [MS-OI29500] 3.4.1.3 the
  CustomUI part carries the plain XML content type - it is covered by the
  package's `<Default Extension="xml">`. Giving it the vendor type
  `application/vnd.ms-office.customUI+xml` makes Excel discard the part with no
  error and no tab. `inject_ribbon.py` therefore adds no Override.

## Troubleshooting

**No Agent tab after restarting Excel.** Check the file exists:
`%APPDATA%\Microsoft\Excel\XLSTART\ExcelAgent.xlam`. If it is there, Excel may
have it in *Disabled Items* — File → Options → Add-ins → Manage: Disabled
Items → Enable.

**"Agent service did not start."** The bridge could not come up. Start it by
hand from WSL:

```bash
cd ~/hay2010agent/excel-agent
python3 bridge.py --port 8765
```

**Replies are slow or fail with 429.** The gateway is rate limiting that model.
Pick another one in the panel, or change the fallback order in `bridge.py`.
`bridge.py` retries, then falls back (`deepseek-v4.1-flash`, `glm-5.3-flash`,
`qwen3.8-flash`, `beam-free`), and the panel notes when it did. To change the
default choice, edit `DEFAULT_MODEL` in `bridge.py` **and** in
`vba/modModels.bas`, then rerun `./make.sh`.

**A link was not read.** The panel shows the reason after the reply. Common
ones: the page needs JavaScript/login (little readable text), it is a PDF or
image, the host is not on your allowlist, or the address resolves to a private
IP. Only `http(s)` GET is supported; paste the exact page URL rather than a
search result link. `web_allowlist.txt` is re-read on every request; if you
changed `EXCEL_AGENT_WEB` or `EXCEL_AGENT_FETCH_HOSTS`, restart the bridge
(`pkill -f bridge.py`, then send again — the add-in auto-starts it).

**Debugging.** `bridge.log` (next to `bridge.py`) has one line per request.
The add-in also writes `%TEMP%\excelagent_trace.log` during self-tests.

**When Excel hangs.** `Application.Run` never returns while VBA is sitting in a
modal dialog, so a hang shows up as a script that stops rather than an error.
`build/loop.sh` builds a throwaway `.xlam` and runs one probe macro under a
timeout, leaving breadcrumbs and a VBE screenshot:

```bash
./build/loop.sh PbGuard          # does EnsureBridgeOrWarn survive?
./build/loop.sh PbFormLayout --form   # dump control geometry at two sizes
```

Probe macros live in `build/probe.bas`. **Do not name a macro `<letters><digits>`**
(`S1`, `P2`): `Application.Run` resolves that as a cell reference and fails with
`0x800A03EC` before your code runs.

## Rebuilding

```bash
cd ~/hay2010agent/excel-agent
./make.sh          # build, deploy, run both self-tests
./make.sh --build-only
```

Building rewrites the add-in from the VBA files in `vba/`, so **edit those, not
the installed copy** — the installed `.xlam` is a build artifact.

Requires: Excel installed, "Trust access to the VBA project object model"
enabled, Python 3 in WSL, and `OPENCODE_API_KEY` or `OPENCODE_GO_API_KEY` in
`~/.dsh/.credentials.yaml`.

**Close Excel before rebuilding.** Excel caches its security settings at
startup and writes them back on exit, so a running Excel can silently reset
`AccessVBOM` to 0 - which makes `build.ps1` fail with *"You cannot call a method
on a null-valued expression"* at `VBComponents.Add`. `build.ps1` now checks and
re-enables the bit, and warns if Excel is running. `./settrust.ps1` does it by
hand too.

## Uninstalling

```bash
rm "$APPDATA/Microsoft/Excel/XLSTART/ExcelAgent.xlam"
```

In PowerShell: `Remove-Item "$env:APPDATA\Microsoft\Excel\XLSTART\ExcelAgent.xlam"`.
Then stop the bridge (close Excel, or `pkill -f bridge.py` in WSL). Your
workbooks are not touched — the add-in stores nothing inside them.
