# Ticket 02 — Local HTTPS, sideload, and backend startup

Target: Windows 11, Excel 16.0 (Microsoft 365) build 20527, personal use, no org admin.
Researched: 2026-11-07.

Confidence tags used inline:

- **[C]** confirmed — a primary source cited on the same line says it.
- **[L]** likely — follows from cited primary sources, but no source states it verbatim.
- **[?]** unknown — not confirmed in docs; must be verified on the machine.

## Summary

- **Excel's pane is Edge WebView2 on Windows, so it uses normal Chromium TLS, cache, CORS, and streaming rules.** [C] ([browsers used by Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins))
- **`https://localhost:<port>` works, with any port, provided the certificate chains to a root the user trusts.** Microsoft says self-signed certs are fine "so long as the certificate is trusted on the local machine", and its own dev-cert tool installs the CA into `cert:\CurrentUser\Root`. [C] ([requirements](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins), [dev-certs install.ps1](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-certs/scripts/install.ps1))
- **Manifest `SourceLocation` must be HTTPS (not a file path), 1–2018 chars; no port or hostname allowlist is documented, and `localhost` is used in Microsoft's own examples.** [C] ([SourceLocation](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/sourcelocation), [IconUrl](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/iconurl))
- **The least fragile sideload is the per-user registry `WEF\Developer` key, not the network share and not "Upload My Add-in".** Office reloads every manifest listed there on each launch, it needs no admin, and it survives cache clears. [C] ([Uninstall add-ins under development](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/uninstall-add-in))
- **Serve the pane and the agent API from one HTTPS origin.** Same-origin removes CORS, mixed-content, and Chrome Local Network Access from the equation in one move. Cross-origin to a *different* loopback port would still work, but only with CORS headers, and a public origin calling `127.0.0.1` is where this design gets fragile. [C] ([same-origin policy](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/addressing-same-origin-policy-limitations), [Secure Contexts §3.1](https://w3c.github.io/webappsec-secure-contexts/#is-origin-trustworthy), [Local Network Access](https://developer.chrome.com/blog/local-network-access))
- **No Excel event can start the backend for a personal, non-admin install.** Event-based activation (`OnDocumentOpened` exists for Excel) "work[s] only when deployed by an administrator". Use Task Scheduler at logon + a health check in the pane. [C] ([event-based activation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation))
- **Streaming: use `fetch` + `POST` + SSE framing on the same origin.** `EventSource` is GET-only and cannot set headers; `fetch` streaming can, and WebSocket buys nothing here. [C] ([HTML SSE](https://html.spec.whatwg.org/multipage/server-sent-events.html))
- **Recommendation: single local HTTPS server (`https://localhost:3000`) serving both the pane and the API, a long-lived localhost cert in `CurrentUser\Root`, registry sideload, Task Scheduler at logon.** Biggest failure mode is a broken/expired certificate or a port that will not bind — either one takes the *whole pane* down, because the backend that serves it is the thing that failed. [L]

## Serving over HTTPS

### Which webview, and what it trusts

Office on Windows renders add-ins in Edge WebView2, which Office installs by default:

> "Windows | any | Microsoft Edge (Chromium-based) with WebView2* — * WebView2 is installed with Office by default for supported versions of Office."
> — [Browsers and webview controls used by Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins) [C]

WebView2's default certificate behaviour is Chromium's, with no add-in-level override:

> "If you don't handle the event, WebView2 will show the default TLS interstitial error page to the user for navigations, and for non-navigations the web request is cancelled."
> — [ICoreWebView2_14 / ServerCertificateErrorDetected](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/icorewebview2_14) [C]

Handling that event is a *host app* API. Office is the host here; Office.js exposes no certificate-error hook, so the certificate must actually be trusted. [L — absence of any such Office.js API; Office controls the WebView2 instance]

WebView2 trust is anchored in the operating-system certificate store:

> "If your environment uses TLS inspection, ensure that the inspection certificates are trusted by the operating system's certificate store."
> — [Prevent security tools from blocking WebView2 apps](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/measures) [C]

The Office docs state the acceptance rule directly:

> "Self-signed certificates can be used for development and testing, so long as the certificate is trusted on the local machine."
> — [Requirements for running Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins) [C]

### Which store, and how Microsoft's own tooling does it

Microsoft's `office-addin-dev-certs` package (used by Yo Office / `npm start`) generates a development CA plus a `localhost` certificate and imports the **CA** into the per-user Root store:

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1 CurrentUser "<ca-cert-path>"
# install.ps1: Import-Certificate -CertStoreLocation cert:\CurrentUser\Root
```

> — [`install.ts`](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-certs/src/install.ts), [`install.ps1`](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-certs/scripts/install.ps1) [C]

So the store that works is **`cert:\CurrentUser\Root`** (Trusted Root Certification Authorities, current user). `--machine` switches to `LocalMachine\Root` and "must be an Administrator" ([README](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-certs/README.md)). [C]

Practical consequences:

- A **chain** works: root CA in `Root`, leaf for `localhost` in `My` (or served directly by the app). Putting a bare self-signed leaf into `Root` also works, but then every renewal is a new explicit trust decision. [L]
- **No public CA is required, and none will issue for `localhost` anyway.** Office's requirement is "trusted on the local machine", not "publicly trusted". [C/L]
- **Validity is short by default**: "The certificate is valid for 30 days by default" (`--days` overrides). A local server that is expected to survive months needs a longer `--days`, or a re-install step in the logon script. [C] ([README](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-certs/README.md))
- `.NET`'s `dotnet dev-certs https --trust` is the other common route: it stores the cert in `CurrentUser\My` and `--trust` "trusts the certificate on the local machine". In .NET 10+ the generated SANs are `localhost`, `*.dev.localhost`, `*.dev.internal`, `host.docker.internal`, `host.containers.internal`, `127.0.0.1`, `::1`; older SDKs' SAN sets are not documented. [C/L] ([dotnet dev-certs](https://learn.microsoft.com/en-us/dotnet/core/tools/dotnet-dev-certs))

### Hostname rules

- **`localhost` and `127.0.0.1` are different origins** (scheme + host + port), so pick one and use it in the cert SAN, the manifest, and every URL. [C] ([URL Standard, origin](https://url.spec.whatwg.org/#origin))
- The certificate **must match the hostname you actually navigate to**. A cert with only `DNS:localhost` will fail on `https://127.0.0.1/...`; a cert with only `IP:127.0.0.1` will fail on `https://localhost/...`. [L — standard TLS name verification; the .NET SAN table shows both are issued separately]
- **IPv6 gotcha:** `localhost` can resolve to `::1` first on Windows. A server bound only to `127.0.0.1` can be reachable at `https://127.0.0.1:3000` but refused at `https://localhost:3000`. Bind loopback on both families (or bind `localhost` and verify both). [?] — not documented by Microsoft; standard dual-stack behaviour, verify once on the machine.
- Ports: **no restriction is documented** for `SourceLocation`. Microsoft's examples use `https://localhost:3000` and `https://localhost:44300`. [C] ([IconUrl example](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/iconurl), [Visual Studio web sideload steps](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing))

### What a broken certificate looks like

- **Top-level navigation to the pane** → WebView2 shows the TLS interstitial; the pane does not load. Office may additionally report an add-in load error; the exact wording is not documented. [C for the interstitial; ? for Office's exact error text]
- **Subresources** (icons, `fetch`, SSE) → request cancelled. Result: blank/default ribbon icon, and API calls that fail with a network error rather than an HTTP status. [C for cancellation; L for the visible symptom]
- Office also warns generally that non-HTTPS endpoints "generate unsecure content warnings and errors during use". [C] ([Requirements](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins))
- There is **no add-in-side bypass** (no Office.js cert API). The only fix is trusting the cert. [L]

## Sideloading

No org admin means: no Centralized Deployment / Integrated Apps, and no event-based activation (see [Backend startup](#backend-startup)). What is left, per Microsoft's own list of Windows sideload paths ([Test Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins)):

| Path | Admin? | Persists across Excel restarts? | Notes |
| --- | --- | --- | --- |
| Registry `WEF\Developer` key | No (per-user HKCU) | **Yes** — "When an Office application launches, it reloads all add-ins listed in this key" [C] | Used by `office-addin-debugging` / Yo Office. Best for a stable personal install. |
| Network-share trusted catalog | No | Yes (TrustedCatalogs key + manifest file in share) | Supported, but documented as non-production and it forces a reinstall when ribbon changes. [C] |
| `Insert > My Add-ins > Upload My Add-in` | No | Only documented for **Office on the web** (manifest kept in browser local storage) | Desktop presence in build 20527 is unconfirmed. [?] |
| Centralized Deployment | **Yes** | n/a | Out of scope (personal use, no admin). |

### Registry sideload (recommended)

Location and format, from the docs:

> "Navigate to **Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer**. This key lists the add-ins that are currently sideloaded... The **Data** value for each entry is the path to the add-in's manifest. The **Name** value varies... typically the add-in's ID. When an Office application launches, it reloads all add-ins listed in this key."
> — [Uninstall add-ins under development](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/uninstall-add-in) [C]

So: create `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer` if missing ([docs include](https://github.com/officedev/office-js-docs-pr/blob/main/docs/includes/developer-registry-key.md)), then add a **String value** whose name is the manifest's `<Id>` GUID and whose data is the full path to the manifest XML. Restart Excel. [C for key semantics; L for the exact value type — the doc says "Name"/"Data" but not REG_SZ explicitly]

Why this is the sturdiest option here:

- **Per-user** (HKCU) → no elevation, no admin, no machine-wide policy. [C]
- **Re-read on every launch** → editing the manifest in place is enough; you do not re-sideload. [C]
- **Not removed by clearing the Office cache.** The docs treat "clear the cache" and "delete the registry entry" as two separate uninstall steps. [C] ([Clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache), [Uninstall](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/uninstall-add-in))
- Survives Office Click-to-Run updates: updates rewrite Program Files, not HKCU. [L — not stated in docs]
- Cost: it is a manual registry write, or one call to Microsoft's tooling:
  `npx office-addin-debugging start manifest.xml` registers it and launches Excel; `npx office-addin-debugging stop` removes it. [C] ([Test and debug on a non-local server](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-non-local-server))

Uninstall is three deletions: the `Developer` value, the `Developer\<ID>` subkey if present, and the `...\Common\CustomUIValidationCache` entries that start with the add-in ID. [C] ([Uninstall](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/uninstall-add-in))

### Network-share trusted catalog

Steps: share a folder (Read/Write), then either **File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs** with the UNC path + "Show in Menu", or write:

```text
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{GUID}]
"Id"="{GUID}"
"Url"="\\\\server\\share"
"Flags"=dword:00000001
```

then put the manifest in the share and use **Home > Add-ins > Advanced > SHARED FOLDER**. [C] ([Sideload on Windows from a network share](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins))

Documented limitations: Windows only; **not supported for production**; **unified manifest not supported** from a network share; and "If a new version of an add-in changes the ribbon... each user will have to reinstall the add-in." [C] (same page)

For a one-machine install this adds moving parts (SMB share, catalog trust, "Show in Menu") without buying anything the registry key does not already give you. [L]

### Upload My Add-in

- **Web client:** documented, and the manifest lives in the browser's local storage — "if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again." [C] ([Sideload to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing))
- **Windows desktop:** not in the current Windows sideload docs. Older Microsoft docs describe `Insert > Office Add-ins > MY ORGANIZATION > Upload My Add-in` for Office 365 in a browser, not the desktop client. [C for the web flow; ? for desktop build 20527]
- How to check on the machine: Excel → **Insert > Add-ins** (or **Home > Add-ins > More Add-ins**). If an "Upload My Add-in" / "My Add-ins" upload control exists, it works for a quick test; do not build the install story on it. [?]

### What persists across restarts and updates

| Thing | Excel restart | Office update | Cache clear |
| --- | --- | --- | --- |
| `WEF\Developer` registry entry | Yes [C] | Expected yes (HKCU) [L] | Survives [C] |
| Network-share catalog trust | Yes [C] | Expected yes (HKCU) [L] | Survives [C] |
| Web upload (web client) | Browser storage [C] | n/a | Lost [C] |
| Pane HTML/JS in Wef/WebView2 cache | Yes [L] | Yes [L] | Cleared [C] |
| Manifest edits (same path/ID) | Re-read from registry path [C] | Yes | Need a cache clear to be safe [C] |

## Manifest URLs

### `SourceLocation`

> "Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path."
> — [SourceLocation element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/sourcelocation) [C]

- **HTTPS required.** `http://localhost:3000` is not a valid `SourceLocation`; `https://localhost:3000/taskpane.html` is. [C]
- **No localhost/port restriction documented.** Microsoft's samples use `https://localhost:3000` and `https://localhost:44300`. Any port is fine as long as the URL is https and the cert matches. [C/L]
- **2018-character cap** — keep query strings short. [C]
- **A certificate error kills the top-level navigation.** The pane never renders; there is no graceful in-pane fallback. [C] (WebView2 interstitial, above)

### `IconUrl` / `HighResolutionIconUrl`

> "Specifies the full, absolute URL of the image... `<IconUrl DefaultValue="https://localhost:3000/assets/images/icon-32.png" />`"
> — [IconUrl element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/iconurl) [C]

- Must be a **full, absolute URL** (a relative path is not valid). [C]
- Task pane icons: **32×32** (content add-ins also need 64×64). Formats: BMP, EXIF, GIF, JPG, PNG, TIFF. [C]
- **HTTPS is not stated as a hard requirement for icons** the way it is for `SourceLocation`, but the documented example is `https://localhost:3000`, and the general "add-ins that are not SSL-secured generate errors" warning applies. Treat icons as HTTPS-only. [C for absolute URL/sizes; L for HTTPS]
- **Changing `IconUrl` at runtime is not supported** — it is a manifest + cache-clear change. [C]
- Failure symptom: a broken/default ribbon icon. Microsoft's own troubleshooting step is "paste the URLs for the add-in's ribbon icons into a browser... see if the icon files open". [C] ([Troubleshoot development errors](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-development-errors))

### `AppDomains` — not a CORS tool

`AppDomains` lists extra domains Office should trust for **pages and iframes that call Office.js** ("This enables pages in the domains to make calls to Office.js APIs from IFrames within the add-in"). It does **not** grant `fetch` permission; CORS governs that. [C] ([AppDomains](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/appdomains), [same-origin policy](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/addressing-same-origin-policy-limitations))

### Manifest hygiene that bites later

- `<Id>` is the add-in's identity. Changing it after install produces "This add-in is no longer available" until the old registration is removed. [C] ([Troubleshoot development errors](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-development-errors))
- `resid` values (and matching resource ids) are capped at **32 characters**. [C] (same page)
- Keep the **manifest file path stable** if you use the registry key, so updates never require re-sideloading. [L]

## CORS and streaming

### The rule

Office add-ins live in a webview, so the ordinary same-origin policy applies:

> "Because Office Add-ins are hosted in a webview control, the same-origin policy applies to script running in their web pages as well."
> — [Address same-origin policy limitations](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/addressing-same-origin-policy-limitations) [C]

Two deployment shapes:

1. **Same origin (recommended):** one server at `https://localhost:3000` serves `/taskpane.html`, `/assets/*`, and `/api/*`. The pane's `fetch('/api/chat')` is same-origin → **no CORS, no preflight, no mixed content, no Local Network Access check, one certificate, one port.** [C for the mechanics; L that this is the simplest]
2. **Split origins:** pane at `https://localhost:3000`, API at `https://127.0.0.1:8787` (or `http://127.0.0.1:8787`). Different host ⇒ different origin ⇒ the API must answer CORS (`Access-Control-Allow-Origin`, `Vary: Origin`, and an `OPTIONS` preflight for `Content-Type: application/json` or custom headers). [C] ([Fetch/CORS](https://fetch.spec.whatwg.org/#http-cors-protocol))

### Mixed content and loopback

The Mixed Content spec only blocks requests whose URL is **not** potentially trustworthy. Loopback is explicitly trustworthy:

> "If origin's host matches one of the CIDR notations 127.0.0.0/8 or ::1/128, return 'Potentially Trustworthy'." ... "Neither origin's domain nor port has any effect on whether or not it is considered to be a secure context."
> — [Secure Contexts §3.1](https://w3c.github.io/webappsec-secure-contexts/#is-origin-trustworthy) [C]

So `http://127.0.0.1:8787` from an HTTPS pane is **not** mixed content and **is** a secure context, per spec. Chromium's *Local Network Access* work is the newer constraint: it gates requests from a **public** origin to local/loopback destinations behind a permission prompt (Chrome 138+ behind a flag at time of writing). A pane served from `localhost` is itself local, so LNA does not apply; a pane served from a tunnel/Worker **does** hit it when it calls the local API. [C] ([Local Network Access](https://developer.chrome.com/blog/local-network-access)) — [L for "a localhost pane is exempt"; the blog frames LNA as public→local]

### SSE

The `EventSource` API is a GET-only stream with no header control:

- Constructor takes a URL and only `{ withCredentials }`; there is no method or headers option. [C] ([HTML SSE §9.2.2](https://html.spec.whatwg.org/multipage/server-sent-events.html#the-eventsource-interface))
- Server response must be `Content-Type: text/event-stream`, UTF-8; messages are `data:` lines terminated by a blank line; `id:` drives `Last-Event-ID` reconnect. [C] (same)
- Auto-reconnects on drop; `204 No Content` tells it to stop. [C] (same)

For a chat pane that needs `POST` + JSON + an `AbortController`, use **`fetch` with a streaming response body** instead: `fetch('/api/chat', {method:'POST', signal, body})` then read `response.body.getReader()` / `TextDecoderStream`, parsing the same `data:` framing. Chromium (WebView2) supports fetch streaming; no Office-specific limitation is documented. [C for fetch streams (Fetch/Streams standards), L for "no Office restriction"]

Server-side requirements for both: disable buffering/proxy compression for that route, flush each chunk, send a heartbeat comment (`: ping`) every ~15s to survive idle timeouts (the SSE spec explicitly recommends this for proxies). [C] ([HTML SSE authoring notes](https://html.spec.whatwg.org/multipage/server-sent-events.html#authoring-notes))

### WebSocket

- Not subject to CORS; the browser sends an `Origin` header and the server should validate it. [C for the handshake model (RFC 6455); L for "validate Origin" as guidance]
- From an HTTPS page, a plaintext `ws://` to a *non-loopback* host is mixed content and blocked; loopback is exempt per the trustworthy-origin rule above. Cross-origin `wss://` avoids the question entirely. [L]
- It buys nothing for this design: SSE over a same-origin `fetch` is one connection, restartable, and easier to proxy. Cloudflare's own tunnel FAQ notes full WebSocket support, but its Quick Tunnels explicitly **do not support SSE** — relevant only if you go the tunnel route. [C] ([Tunnels FAQ](https://developers.cloudflare.com/cloudflare-one/faq/cloudflare-tunnels-faq/), [Quick Tunnels](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/do-more-with-tunnels/trycloudflare/))

### Concrete shape that avoids all of it

```
https://localhost:3000/
  /taskpane.html          static pane
  /assets/*               icons, JS, CSS
  /api/health             GET  -> 200 {"ok":true}
  /api/chat               POST -> text/event-stream
```

No CORS headers, no preflight, no second cert, no second port, nothing for a public origin to trip over. [L]

## Backend startup

### The structural constraint

The backend **serves the pane**. If the backend is not running, Excel has nothing to load — the task pane fails before any of your JavaScript runs, so the pane cannot even show "starting…". That makes startup reliability part of the install, not a nice-to-have. [L — follows directly from `SourceLocation` being served by the backend]

(Alternative: host the pane on a static host and keep only the API local. That decouples "pane loads" from "backend up", but adds a public origin + LNA/CORS problems — see [Hosting options compared](#hosting-options-compared).)

### Why no Excel event can do it

- Office.js has no process-launch API; the pane is a sandboxed webview. [L]
- Event-based activation *does* have an Excel `OnDocumentOpened` event, but:

> "Event-based add-ins work only when deployed by an administrator. If users install them directly from Microsoft Marketplace or the Office Store, they will not automatically launch... Admin deployments are done by uploading the manifest to the Microsoft 365 admin center."
> — [Activate add-ins with events](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation) [C]

So for a personal, sideloaded install: **no Excel-start event is available.** [C]

### Task Scheduler at logon (recommended)

```bat
schtasks /create /tn "ExcelAgentBackend" ^
  /tr "\"C:\Tools\excel-agent\run-backend.cmd\"" ^
  /sc onlogon /delay 0000:30 /rl LIMITED /f
```

- `/sc onlogon` = "runs whenever a user (any user) logs on"; `/delay` is valid for ONLOGON and takes `mmmm:ss`; `/rl LIMITED` runs with least privilege. [C] ([schtasks create](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/schtasks-create))
- **Elevation:** the same doc says "Only Administrators can schedule tasks". Create it once from an elevated prompt (or via Task Scheduler GUI → Create Task). After that it runs as you at every logon without prompting. [C for the admin statement; L for "runs without prompting thereafter"]
- **Restart on failure** is not a `schtasks /create` flag; set it in the GUI (Settings → "If the task fails, restart every…") or in the task XML. [L]
- Point `/tr` at a `.cmd` wrapper so you can redirect logs and set the working directory:

```bat
@echo off
cd /d C:\Tools\excel-agent
python server.py >> "%LOCALAPPDATA%\ExcelAgent\logs\backend.log" 2>&1
```

- Task Scheduler records `Last Run Result` and history; that plus your file log is enough to diagnose "did it even start". [C for Task Scheduler history; L for the combination]

### Windows service (heavier, needs admin)

```
sc.exe create ExcelAgentBackend binpath= "C:\Tools\excel-agent\backend.exe" start= auto
```

- `start= auto` = "starts automatically each time the computer is restarted and runs even if no one logs on". Default account is `LocalSystem`. Creating a service requires elevation. [C for the flags and default account ([sc create](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/sc-create)); C/L for elevation]
- Consequence: it runs in session 0 as `LocalSystem`, so a cert trusted in **your** `CurrentUser\Root` is invisible to it — it needs `LocalMachine\Root` (admin) or a cert file with its own ACL. That is exactly the extra failure surface a personal single-user tool does not need. [L]
- Use a service only if you want the backend up before login (e.g., it must serve other machines). Not the fit here.

### Startup-folder shortcut (no admin, least machinery)

Put a shortcut to the `.cmd` in `shell:startup`. No elevation, trivially removable, but no delay and no restart-on-failure. Good fallback if the Task Scheduler task cannot be created (e.g., locked-down machine). [L — standard Windows behaviour, not covered by the Office docs]

### Port conflicts

- **Pick a fixed port below 49152.** Windows' dynamic/ephemeral range is 49152–65535, so a server on 3000/8443 is outside the range clients grab from. [C] ([Service overview and network port requirements](https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/service-overview-and-network-port-requirements))
- **Check for reserved ranges first** (Hyper-V, WSL, Docker and some VPN clients reserve blocks):

```bat
netsh int ipv4 show excludedportrange protocol=tcp
```

  Binding into an excluded range fails with `WSAEACCES (10013)`; Microsoft's workaround is to "use a port that is not included in the default dynamic port range (from 49,152 to 65,535)". [C] ([Error 10013](https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/error-10013-wsaeacces-is-returned)) — [? for the exact `show excludedportrange` output on this machine]
- **Single instance:** on `EADDRINUSE`, probe `https://localhost:3000/api/health`; if it answers with your app's shape, exit 0 (already running). If it answers with something else, log loudly and exit non-zero — never silently pick another port, because the manifest hard-codes this one. [L]

### Health check and logs

- `GET /api/health` → `200 {"ok":true,"version":"…","pid":…}`. The pane calls it on load with a ~1.5s timeout and shows one of three states: ready / starting (retry with backoff) / not running (show the exact command to start it and a link to the log). [L — design guidance]
- Log to a file, not stdout only, because Task Scheduler swallows console output: `%LOCALAPPDATA%\ExcelAgent\logs\backend-YYYYMMDD.log`, with the bind address, cert path/thumbprint, and port on startup. Never log workbook content or API keys. [L]
- Distinguish Office's **runtime logging** (manifest/host diagnostics; `npx office-addin-dev-settings runtime-log --enable`) from your backend log — it does not capture your `console.log`. [C] ([runtime logging](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtime-logging))

## Dev loop

| Change | What to do | Source |
| --- | --- | --- |
| HTML/CSS/JS in the pane | Focus the pane, **Ctrl+F5** — "you shouldn't need to clear the cache" | [C] ([Clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache)) |
| Manifest (icons, command text, URLs) | Clear the Office cache, then reopen | [C] (same) |
| Cache, automated | `npx office-addin-cache clear` | [C] (same) |
| Cache, automatic every start | File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs > "Next time Office starts, clear all previously-started web add-ins cache" (Office 2108+; build 20527 is newer) | [C] (same) |
| Cache, manual | Delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` and (if present) `%LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` | [C] (same) |
| Stale assets from the web server | Send `Cache-Control: private, no-cache, no-store`, `Pragma: no-cache`, `Expires: -1` | [C] ([Troubleshoot development errors](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-development-errors)) |
| Edge cache that will not clear | `del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` | [C] (same) |
| Debugger | Focus the pane, **Ctrl+Shift+I** (or right-click → Inspect) | [C] ([Debug with Edge DevTools](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-chromium)) |
| Debug initialization | Close Office, set `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--auto-open-devtools-for-tabs`, reopen | [C] (same) |
| Manifest doesn't load at all | `npx office-addin-dev-settings runtime-log --enable`, search the log for the add-in ID (`SolutionId`) | [C] ([runtime logging](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtime-logging)) |

Notes:

- **Reloading the pane does not reload the manifest.** Manifest edits need a cache clear. [C]
- **Changing the add-in `<Id>`** needs remove-then-re-add, or Excel reports "This add-in is no longer available". [C]
- **Cache clear is destructive to sideloads installed via the cache**; registry-registered add-ins survive it. [C/L]
- **Cache busting for production:** content-hash filenames (`app.9f3c1.js`) plus a small non-cached `taskpane.html`. Avoid changing the `SourceLocation` itself just to bust cache — that is a manifest change and therefore a cache clear. [L]
- In dev, disable HTTP caching entirely rather than fighting the WebView2 cache. [C for the headers, L for the strategy]

## Hosting options compared

| | **A. Local HTTPS, same origin** | **B. Cloudflare named Tunnel → local server** | **C. Cloudflare Quick Tunnel** | **D. Deployed Worker (static pane only)** | **E. Worker pane + Tunnel API** |
| --- | --- | --- | --- | --- | --- |
| Setup | Dev cert + one server + registry key | cloudflared install, account, domain, DNS route, service | One command, no account | `wrangler deploy` + domain | Both B and D |
| Admin needed | No | No (local install) | No | No | No |
| Internet dependency | **None** | Yes — tunnel must stay up | Yes | Yes | Yes |
| Pane loads when backend is down | **No** (backend serves it) | Yes | Yes | Yes | Yes |
| Data leaves the machine | **No** | Yes (through Cloudflare) | Yes | Pane only | Yes |
| SSE | **Native** | Supported in principle; Cloudflare docs confirm WebSockets, do **not** document SSE for named tunnels [?] | **Documented as not supported** [C] | n/a | same as B |
| Stable URL | Yes (localhost) | Yes, once DNS is set | **No** — "hostname changes each time" [C] | Yes | Yes |
| Extra failure modes | cert trust/expiry; port bind | tunnel down, DNS, account, origin cert mode, access rules | URL churn, 200 in-flight cap, no SSE, no uptime [C] | local API calls from a public origin hit LNA/CORS | both |
| Manifest churn | None | None | **Every restart** → cache clear each time | None | None |
| Verdict for this project | **Recommended** | Fallback for remote access only | Not viable for streaming | Breaks local-first for the API | Over-engineered |

Sources for the table: [Tunnel overview](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/), [Create a tunnel](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/get-started/create-remote-tunnel/), [Quick Tunnels](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/do-more-with-tunnels/trycloudflare/), [Tunnels FAQ](https://developers.cloudflare.com/cloudflare-one/faq/cloudflare-tunnels-faq/), [Workers static assets](https://developers.cloudflare.com/workers/static-assets/), [Local Network Access](https://developer.chrome.com/blog/local-network-access). [C for quoted limits; L for the comparison judgements]

The decisive point: **a tunnel or Worker makes the pane a public origin, and a public origin talking to `127.0.0.1` is exactly the pattern Local Network Access is designed to gate.** Keeping the pane on `localhost` keeps the whole app on the trusted side of that line. [L]

## Recommendation

**Least fragile local-first setup: one HTTPS origin on `https://localhost:3000`, a long-lived locally-trusted certificate, registry sideload, Task Scheduler at logon.**

1. **Certificate** — create a dev CA + `localhost` cert with a long validity and trust the CA for the current user:
   ```bat
   npx office-addin-dev-certs install --days 825
   ```
   Verify: `npx office-addin-dev-certs verify`; or check `certmgr.msc` → Trusted Root Certification Authorities → Certificates for "Developer CA for Microsoft Office Add-ins". Interim: `dotnet dev-certs https --trust`. [C for the commands ([README](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-certs/README.md)); ? for the maximum `--days` accepted]
2. **One server, two routes** — bind loopback on **both** IPv4 and IPv6, port **3000** (below the ephemeral range, unlikely to be reserved). Serve `/taskpane.html`, `/assets/*`, `/api/*`. Same origin by construction. [L]
3. **Health + logs** — `/api/health`, a file log, single-instance bind check. [L]
4. **Manifest** — `SourceLocation` = `https://localhost:3000/taskpane.html`; `IconUrl` = `https://localhost:3000/assets/icon-32.png`; `<Id>` fixed forever; `resid`s ≤ 32 chars. [C]
5. **Sideload** — write `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` → `{<manifest-Id>}` = `C:\Tools\excel-agent\manifest.xml`, then start Excel. (Or run `npx office-addin-debugging start manifest.xml` once.) [C]
6. **Survive reboots** — Task Scheduler `/sc onlogon /delay 0000:30` pointing at a `.cmd` that `cd`s to the install dir and redirects output to the log. Create it from an elevated prompt once. [C for the flags]
7. **Fail visibly** — the pane's first action is `fetch('/api/health', {signal: AbortSignal.timeout(1500)})`; on failure show "backend not running" + the exact start command + log path, and retry with backoff. Since the pane is served by the backend, this only helps for a *late* failure; for a cold start the fix is the logon task. [L]
8. **Dev loop** — `Ctrl+F5` for code, cache clear for manifest, `no-store` headers, `Ctrl+Shift+I` for DevTools. [C]

**Biggest failure mode: the certificate.** An expired/renewed/untrusted cert or a hostname mismatch produces a blank pane with a WebView2 TLS interstitial, which looks identical to "the add-in is broken" and is invisible from inside Excel. Microsoft's dev cert defaults to **30 days**, so a "finished" install quietly dies about a month later. Mitigate with a long `--days`, a startup check that logs cert subject/expiry, and a browser tab open to `https://localhost:3000/taskpane.html` as the outside-Excel smoke test. [C for the 30-day default and interstitial; L for the mitigation]

**Second failure mode: the port.** If something else owns 3000 (Hyper-V/WSL reservation, a leftover process), the server never binds, so the pane never loads. Mitigate with the excluded-port-range check, a single-instance probe, and a log line that names the failing bind. [C for excluded ranges/10013; L for the mitigation]

**Deliberately rejected:** Cloudflare Quick Tunnel (no SSE, URL churn per start, no uptime); a deployed Worker for the pane (turns local API calls into public→local requests subject to LNA, and moves UI code off-machine); event-based activation (admin-only). A **named tunnel** is the right choice only if you later want to open the pane from another device; it is not the right default for a one-machine local-first tool. [C/L]

## Sources

Office Add-ins (Microsoft Learn):

1. [Browsers and webview controls used by Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins) — WebView2 on Windows; WIP; PDF workaround.
2. [Requirements for running Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins) — HTTPS recommended; self-signed OK if trusted on the machine.
3. [SourceLocation element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/sourcelocation) — HTTPS only; 1–2018 chars; not a file path.
4. [IconUrl element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/iconurl) — absolute URL; sizes/formats; localhost example.
5. [AppDomains element](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/appdomains) — extra trusted domains for iframes/Office.js, not CORS.
6. [Address same-origin policy limitations](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/addressing-same-origin-policy-limitations) — SOP applies in the webview.
7. [Sideload Office Add-ins on Windows from a network share](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) — share + TrustedCatalogs registry; limitations.
8. [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) — Upload My Add-in; localStorage.
9. [Test Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins) — canonical list of Windows sideload paths.
10. [Test and debug Office Add-ins on a non-local server](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-non-local-server) — office-addin-debugging registers in the Windows registry; `stop` removes.
11. [Uninstall add-ins under development](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/uninstall-add-in) — `WEF\Developer` layout and reload-at-launch behaviour; CustomUIValidationCache cleanup.
12. [Clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache) — Ctrl+F5; auto-clear checkbox; Wef/INetCache paths; office-addin-cache.
13. [Troubleshoot development errors with Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-development-errors) — no-cache headers; icon URL check; ID change errors; resid limits.
14. [Debug add-ins using developer tools in Microsoft Edge](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-chromium) — Ctrl+Shift+I; WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS.
15. [Debug your add-in with runtime logging](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtime-logging) — manifest/host diagnostics only.
16. [Activate add-ins with events](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation) — Excel `OnDocumentOpened`; admin-only deployment.
17. [Automatically open a task pane with a document](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document) — autoopen is document-scoped; registry `storeType` "developer".
18. [Excel task pane quickstart](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery) and the [start-server include](https://github.com/officedev/office-js-docs-pr/blob/main/docs/includes/quickstart-yo-start-server-excel.md) — `npm start`, localhost dev server.

Microsoft Edge WebView2:

19. [Prevent security tools from blocking WebView2 apps](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/measures) — trust follows the OS certificate store; TLS inspection.
20. [ICoreWebView2_14 (ServerCertificateErrorDetected)](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/icorewebview2_14) — default interstitial / cancelled subresources; host-app-only override.
21. [Manage user data folders](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/user-data-folder) — UDF holds cache/cookies/localStorage.

Microsoft tooling source:

22. [`office-addin-dev-certs` README](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-certs/README.md) — 30-day default; `--days`; `--machine` needs admin.
23. [`install.ts`](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-certs/src/install.ts) and [`install.ps1`](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-certs/scripts/install.ps1) — CA imported to `cert:\CurrentUser\Root`.
24. [`sideload.ts`](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-dev-settings/src/sideload.ts) — SourceLocation must be https and localhost/127.0.0.1 for the web path; registry registration for desktop.
25. [docs include: developer-registry-key.md](https://github.com/officedev/office-js-docs-pr/blob/main/docs/includes/developer-registry-key.md) — creating the `Developer` key.

Windows platform:

26. [schtasks create](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/schtasks-create) — `/sc onlogon`, `/delay`, `/rl`, admin requirement.
27. [sc.exe create](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/sc-create) — `start= auto`, LocalSystem default.
28. [Service overview and network port requirements](https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/service-overview-and-network-port-requirements) — dynamic range 49152–65535.
29. [Error 10013 (WSAEACCES)](https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/error-10013-wsaeacces-is-returned) — excluded port ranges; use a port outside the dynamic range.

Web platform / Cloudflare:

30. [Secure Contexts §3.1](https://w3c.github.io/webappsec-secure-contexts/#is-origin-trustworthy) — loopback is potentially trustworthy; port irrelevant.
31. [Mixed Content](https://w3c.github.io/webappsec-mixed-content/) — definition keys off potentially-trustworthy URLs.
32. [HTML: Server-sent events](https://html.spec.whatwg.org/multipage/server-sent-events.html) — EventSource interface, framing, reconnect, authoring notes.
33. [Fetch Standard](https://fetch.spec.whatwg.org/) — CORS protocol and streaming response bodies.
34. [URL Standard: origin](https://url.spec.whatwg.org/#origin) — host+port+scheme define the origin.
35. [Local Network Access](https://developer.chrome.com/blog/local-network-access) — public→local/loopback permission; Chrome 138+.
36. [dotnet dev-certs](https://learn.microsoft.com/en-us/dotnet/core/tools/dotnet-dev-certs) — CurrentUser\My; `--trust`; .NET 10 SANs.
37. [Cloudflare Tunnel overview](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/) — outbound-only connector model.
38. [Create a tunnel](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/get-started/create-remote-tunnel/) — named tunnel, public hostname route.
39. [Quick Tunnels](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/do-more-with-tunnels/trycloudflare/) — no SSE, 200 in-flight cap, hostname churn, no uptime.
40. [Tunnels FAQ](https://developers.cloudflare.com/cloudflare-one/faq/cloudflare-tunnels-faq/) — WebSocket support.
41. [Workers static assets](https://developers.cloudflare.com/workers/static-assets/) — deploying pane assets to a Worker.
