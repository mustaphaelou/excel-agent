# 03 — Gateway streaming and tool calling

Research for ticket 03, "Gateway streaming and tool calling". Read-only: no code was
changed. All observations are against the OpenCode Go gateway (`https://opencode.ai/zen/go/v1`)
with this machine's key; no key material or workbook data appears below.

**Evidence tags used throughout:** `[confirmed]` = observed directly or stated by the owning
docs/source; `[inferred]` = follows from a confirmed fact but was not exercised; `[unknown]` =
not established, and say so rather than guessing.

## Summary

The gateway is a plain OpenAI-compatible `/chat/completions` endpoint. It supports SSE
streaming (`stream: true`, standard `chat.completion.chunk` frames, `[DONE]`, and a final
usage-only chunk), OpenAI-style `tools`/`tool_choice`, parallel tool calls, and
`response_format: json_schema`. `[confirmed]` All 30 models in the live catalog advertise
tool support; 20 advertise structured output. Context windows are 262,144–1,050,000 tokens and
output caps up to 384K — none of this is the constraint on the add-in.

The current bridge uses none of it: `bridge.py` posts a non-streaming request with no
`tools`, no `response_format`, `max_tokens: 3000`, and turns the model's free-text JSON
reply into a line/unit-separator action protocol that VBA parses. `[confirmed]` The gap is
therefore entirely client-side: no tool loop, no executor, no streaming path, and no CORS on
the local bridge. The gateway blocks direct browser calls with CORS (preflight 404, no
`Access-Control-Allow-Origin` on `/chat/completions`), so the key has to stay behind a local
backend regardless. Only the pane can execute Excel operations, so the tool loop belongs
there and the backend becomes a stateless proxy.

## Method

- Read `bridge.py` (1243 lines), `README.md`, and the VBA protocol/HTTP/model modules.
- One direct `GET /models` and one `GET /models` through a temporarily started local bridge
  (`--port 8799`, killed afterwards); both returned the same 30-model live catalog. The
  running user bridge on `:8765` was not touched.
- Four **free-model** completion requests against `beam-free` (usage reports `cost: 0`):
  one SSE stream with a tool, one non-streaming parallel-tool request, and two
  `response_format: json_schema` requests. Two failure probes (fake key, unknown model) exercised
  the error shapes without invoking a model. No paid-model completion was made.
- Public sources: OpenCode Zen/Go docs and one gateway issue on tool-call streaming.

## Current call path

All line references are `bridge.py` unless noted.

| Fact | Value | Evidence |
|---|---|---|
| Base URL | `https://opencode.ai/zen/go/v1` | `bridge.py:50` `[confirmed]` |
| Chat endpoint | `POST {GATEWAY}/chat/completions` | `bridge.py:860` `[confirmed]` |
| Auth header | `Authorization: Bearer <key>`, key from env `OPENCODE_API_KEY` / `OPENCODE_GO_API_KEY`, then `~/.dsh/.credentials.yaml` or `~/.dsh/.env` | `bridge.py:120-142`, `68-72` `[confirmed]` |
| Other headers | `Content-Type: application/json`, `Accept: application/json`, `User-Agent: opencode/1.2.3`, `x-opencode-session: <one UUID per process>` | `bridge.py:836-842`, `825` `[confirmed]` |
| Default model | `beam-free` | `bridge.py:51`; `vba/modModels.bas:12` `[confirmed]` |
| Fallback order | requested model → `deepseek-v4.1-flash` → `glm-5.3-flash` → `qwen3.8-flash` → `beam-free`, de-duplicated | `bridge.py:53`, `844-847` `[confirmed]` |
| Retries | 2 attempts per model; 429/500/502/503/504 sleep `1.5 * attempt` s and retry; 401/403 abort the turn immediately; other 4xx break to the next model | `bridge.py:850-889` `[confirmed]` |
| Socket timeout | 180 s per attempt (whole non-streamed body) | `bridge.py:863` `[confirmed]` |
| Model-catalog timeout | 10 s, cached 600 s | `bridge.py:231`, `66` `[confirmed]` |
| Payload | `{model, messages, max_tokens: 3000, temperature: 0.2}` — no `tools`, no `stream`, no `response_format` | `bridge.py:855-858` `[confirmed]` |
| Reply parsing | non-streaming `json.load(resp)` → `choices[0].message.content`, falling back to `reasoning_content` | `bridge.py:864-868` `[confirmed]` |
| Empty text | treated as `last_err = "empty completion"` and the next model is tried | `bridge.py:871` `[confirmed]` |
| User-visible error | `REPLY "The model gateway did not answer: <err>"` + `NOTE "gateway error - see bridge.log"`; HTTP status is not surfaced | `bridge.py:1136-1142` `[confirmed]` |
| Logging | model, requested model, web count, action count, sheets, elapsed; never the key | `bridge.py:1165-1168` `[confirmed]` |
| VBA timeouts | `/chat` POST: `SetTimeouts 3000, 3000, 20000, timeoutSec*1000` = 3 s resolve, 3 s connect, 20 s send, 180 s receive; `/models` GET: 8 s receive | `vba/modHttp.bas:99`, `76` `[confirmed]` |

Model list: `bridge.py` fetches `GET {GATEWAY}/models`, normalizes `{id, name, free}` from the
top-level `data` array, drops non-`type == "chat"` entries, sorts default → preferred → free →
rest, and caches for 10 minutes (`bridge.py:164-249`). When the call fails (offline, no key,
401) it falls back to a six-entry static list (`bridge.py:58-65`, `239-242`) and the panel
marks the source as `fallback` (`bridge.py:252-262`). Free-ness is `cost.text.input.price == 0`
(`bridge.py:156-161`). `[confirmed]` The live payload is much richer than the bridge keeps —
it also carries `capabilities`, `limit`, `cost`, `modalities`, and per-model `api` metadata,
all discarded by `_normalize_models`.

The bridge hands the model exactly one system prompt and up to 12 history messages
(`bridge.py:802-818`), with a workbook snapshot capped at 80×26 cells / 160 chars per cell /
45 000 chars total (`bridge.py:74-78`, `372-414`). Web pages the user links are fetched first
(`bridge.py:1130`).

### Error shapes actually observed

`[confirmed]` Unauthenticated request to `/chat/completions`:

```http
HTTP/1.1 401 Unauthorized
Content-Type: application/json

{"error":{"type":"authentication_error","message":"Invalid credential"}}
```

`[confirmed]` Unknown model (same shape seen earlier in `bridge.log:71`):

```http
HTTP/1.1 400 Bad Request
Content-Type: application/json

{"error":{"type":"server_error","message":"Upstream request failed: Model is unavailable."}}
```

Note the `type` is `server_error` even on a 400; a client must key off the HTTP status and the
`message`, not `type`. No `Retry-After` header was present on any response.

### Rate limits and quotas

`[confirmed]` No rate-limit headers were returned on `/models` or on completions (only
`Date`, `Content-Type`, `Server: cloudflare`, `CF-RAY`, plus gateway diagnostics
`x-opencode-log-id`, `x-zen-model`, `server-timing`).

`[confirmed]` Documented limits are monetary, not RPM. Go is a $10/month plan (Go Plus $40,
Go Teams $20/member); each model has a monthly dollar allowance, enforced as 5-hour = 20 %,
weekly = 50 %, monthly = 100 %. Free models are listed as "Unlimited" for a limited time. Each
response reports what it counted in `usage.opencode.allowance` (observed as `0` for free
models, alongside `usage.opencode.billing_source: "go"`, `list_cost`, and per-bucket token
counts). Over-limit behaviour / 429 body: `[unknown]` — not exercised.

`[confirmed]` Docs also state Go is "designed for OpenCode and other coding agents", traffic
is monitored for abuse, clients should identify themselves with their own user agent, and
should send a stable per-conversation `x-opencode-session`.

## Streaming

`[confirmed]` SSE works on `POST /zen/go/v1/chat/completions` with `stream: true`:
`Content-Type: text/event-stream`, `data:`-prefixed JSON frames, terminated by `data: [DONE]`.
Observed on `beam-free` with a tool definition: 200, first byte ≈ 906 ms, full turn ≈ 1.1 s,
frames at ~35–40 ms spacing.

Exact frame sequence (abridged; content strings shown verbatim, ids shortened):

```
data: {"id":"chatcmpl-…","object":"chat.completion.chunk","created":1794081193,"model":"beam-free",
       "choices":[{"index":0,"finish_reason":null,"delta":{"role":"assistant"}}],"usage":null}
data: {… "delta":{"reasoning_content":"The"}}
data: {… "delta":{"reasoning_content":" user wants me to read"}}
data: {… "delta":{"tool_calls":[{"index":0,"id":"call-7e90…","type":"function",
       "function":{"name":"read_cell","arguments":""}}]}}
data: {… "delta":{"tool_calls":[{"index":0,"function":{"arguments":"{\"ref\": \"A1\"}"}}]}}
data: {… "choices":[{"index":0,"finish_reason":"tool_calls","delta":{}}]}
data: {"id":"chatcmpl-…","choices":[],"usage":{"prompt_tokens":87,"completion_tokens":40,
       "total_tokens":127,"prompt_tokens_details":{"cached_tokens":0},
       "completion_tokens_details":{"reasoning_tokens":20},"cost":0,
       "opencode":{"billing_source":"go","list_cost":0,"allowance":0,"tokens":{…}}}}
data: [DONE]
```

Notes for an implementer:

- The bridge currently sends `Accept: application/json` (`bridge.py:839`); a streaming client
  must send `Accept: text/event-stream` and stop calling `json.load(resp)` (`bridge.py:864`).
  `[confirmed]`
- The model's hidden reasoning streams as `delta.reasoning_content` (DeepSeek convention), not
  `delta.content`. A naive renderer that only reads `content` shows nothing until the tool call.
  `[confirmed]`
- Usage arrives by default in a final chunk with `"choices": []`; `stream_options` is not
  required. `[confirmed]`
- Continuation tool-call deltas normally carry only `index` + `function.arguments`; the gateway
  has historically emitted explicit `"id": null` / `"function": {"name": null}` on some upstream
  backends, which breaks presence-based accumulators. `[confirmed]` by a gateway issue with raw
  SSE; not reproduced in today's probe. An accumulator should ignore null fields.
- First `tool_calls` delta carries `index`, `id`, `type`, `function.name`, `arguments: ""`.
  `[confirmed]`
- `finish_reason` is `tool_calls`; `stop` for normal text. `[confirmed]`
- Cancellation is not documented; the standard approach is to abort the HTTP request (close the
  stream). The current bridge cannot cancel: it blocks in `urlopen` until the full body arrives
  (`bridge.py:863`). `[unknown]` server-side cancellation semantics; `[inferred]` closing the
  connection stops generation.
- Translation: several models are natively served on the Responses (`gpt-6-luna`,
  `crouton-alpha-free`, `grok-4.7`, `muse-spark-*`) or Anthropic Messages (`claude-haiku-5-5`,
  `qwen3.8-*`, `minimax-m3`) or Google APIs; the docs state every text model also works through
  the other APIs and Console translates. `[confirmed]` for text: `bridge.log` shows
  `crouton-alpha-free` answered through the bridge's `/chat/completions` path. Tool-call fidelity
  through translation for those families: `[unknown]`.

## Tool calling

`[confirmed]` OpenAI-style function calling works on `/chat/completions`.
Request used in the probe (free model, `cost: 0`):

```json
{
  "model": "beam-free",
  "max_tokens": 128,
  "parallel_tool_calls": true,
  "messages": [{"role": "user", "content": "In one turn, read A1 and write 'done' to B1 using both tools."}],
  "tools": [
    {"type": "function", "function": {"name": "read_cell", "description": "Read a cell value",
      "parameters": {"type": "object", "properties": {"ref": {"type": "string"}},
                     "required": ["ref"], "additionalProperties": false}}},
    {"type": "function", "function": {"name": "write_cell", "description": "Write a value to a cell",
      "parameters": {"type": "object", "properties": {"ref": {"type": "string"}, "value": {"type": "string"}},
                     "required": ["ref", "value"], "additionalProperties": false}}}
  ],
  "tool_choice": "auto"
}
```

Response (200, ≈1.12 s, abridged):

```json
{
  "choices": [{
    "finish_reason": "tool_calls",
    "message": {
      "role": "assistant", "content": null, "refusal": null,
      "tool_calls": [
        {"id": "call-3073…", "type": "function",
         "function": {"name": "read_cell", "arguments": "{\"ref\": \"A1\"}"}},
        {"id": "call-eea2…", "type": "function",
         "function": {"name": "write_cell", "arguments": "{\"ref\": \"B1\", \"value\": \"done\"}"}}
      ]
    }
  }],
  "usage": {"prompt_tokens": 120, "completion_tokens": 82, "total_tokens": 202,
            "cost": 0, "opencode": {"billing_source": "go", "allowance": 0, "tokens": {…}}}
}
```

- **Parallel tool calls:** confirmed — two independent calls in one assistant message when
  `parallel_tool_calls: true`. `[confirmed]`
- **`tool_choice`:** catalog metadata lists per-model allowed values. `deepseek-v4.1-flash`,
  `deepseek-v4-pro`, `grok-4.7`, `gpt-5.6-luna`, `gpt-6-luna`, `claude-haiku-5-5`, `beam-free`,
  `gemini-4-argon` expose `auto | none | required | tool`; `kimi-k3` and `kimi-k3.1-pro` expose
  `auto | none | required` (`kimi-k2.7-code` only `auto | none`); `qwen3.8-*` and `qwen3.7-plus`
  expose `auto | none | tool`; `glm-5.2`, `glm-5.3`, `glm-5.3-flash`, `mimo-v2.6-*` expose only
  `auto`; and `hy3`, `hy4`, `longcat-2.0`, `step-5-preview`, `iceman-contributor-free`,
  `crouton-alpha-free` advertise no choice list at all. `[confirmed]` from `GET /models`.
- **Strict schemas:** `api.<family>.capabilities.tools.strict: true` for `gpt-5.6-luna`,
  `gpt-6-luna`, `claude-haiku-5-5`, `kimi-k2.7-code`, `kimi-k3`, `kimi-k3.1-pro`,
  `mimo-v2.6-flash`, `mimo-v2.6-pro`, `beam-free`; absent for the rest. `[confirmed]`
- **Structured outputs:** `capabilities.structured_output: true` for 20 of 30 models
  (`deepseek-v4.1-flash`, `deepseek-v4-pro`, `glm-5.2/5.3/5.3-flash`, `grok-4.7`, `muse-spark-*`,
  `gpt-5.6-luna`, `gpt-6-luna`, `claude-haiku-5-5`, `kimi-k2.7-code`, `kimi-k3`, `kimi-k3.1-pro`,
  `step-5-preview`, `beam-free`, `qwen3.8-max`, `qwen3.8-flash`, `gemini-4-argon`,
  `crouton-alpha-free`); `null` (not advertised) for `hy3`, `hy4-preview`, `mimo-v2.6-*`,
  `minimax-m3`, `longcat-2.0`, `qwen3.7-plus`, `iceman-contributor-free`, `hy4`. `[confirmed]`
  `response_format: {"type":"json_schema","json_schema":{"name":…,"strict":true,"schema":…}}`
  returned 200 and exactly `{"answer":"ok"}` on `beam-free`. `[confirmed]` On DeepSeek routes
  there is a [third-party report of a 400 for `json_schema`](https://github.com/lidge-jun/opencodex/pull/4274);
  not reproducible without a paid call, so treat DeepSeek structured output as `[unknown]`
  despite the catalog flag.
- **Tool results:** the standard follow-up is `{"role":"tool","tool_call_id":"call-…","content":"…"}`
  and re-posting the conversation. `[inferred]` from OpenAI compatibility; not exercised.
- **Empty `content` is normal in a tool turn.** The current bridge treats no text as
  `"empty completion"` and falls back to another model (`bridge.py:866-871`), so a native
  tool-call response would today be misread as a failure. `[confirmed]` by code path.

## Models and limits

Live catalog, 2026-11-07 (`GET /models`, 30 entries; `ctx`/`out` tokens; prices $/1M).
All entries advertise `capabilities.tools.supported: true`.

| id | API family | struct | ctx | out | in | out $ |
|---|---|---|---|---|---|---|
| deepseek-v4.1-flash | chat-completions | yes | 1 000 000 | 384 000 | 0.15 | 0.60 |
| deepseek-v4-pro | chat-completions | yes | 1 000 000 | 384 000 | 0.66 | 1.98 |
| glm-5.2 | chat-completions | yes | 1 000 000 | 131 072 | 1.40 | 4.40 |
| glm-5.3 | chat-completions | yes | 1 000 000 | 131 072 | 1.40 | 4.40 |
| glm-5.3-flash | chat-completions | yes | 1 000 000 | 131 072 | 0.15 | 0.50 |
| qwen3.8-max | messages | yes | 1 000 000 | 131 072 | 2.00 | 6.00 |
| qwen3.8-flash | messages | yes | 1 000 000 | 131 072 | 0.15 | 0.47 |
| qwen3.7-plus | messages | no | 1 000 000 | 65 536 | 0.40 | 1.60 |
| grok-4.7 | responses | yes | 500 000 | 500 000 | 2.00 | 6.00 |
| gpt-5.6-luna | responses | yes | 1 050 000 | 128 000 | 0.20 | 1.20 |
| gpt-6-luna | responses | yes | 1 050 000 | 128 000 | 0.10 | 0.50 |
| claude-haiku-5-5 | messages | yes | 1 000 000 | 128 000 | 0.10 | 0.50 |
| kimi-k2.7-code | chat-completions | yes | 262 144 | 262 144 | 0.95 | 4.00 |
| kimi-k3 | chat-completions | yes | 1 048 576 | 131 072 | 3.00 | 15.00 |
| kimi-k3.1-pro | chat-completions | yes | 1 048 576 | 131 072 | 2.00 | 10.00 |
| mimo-v2.6-flash | chat-completions | no | 1 048 576 | 131 072 | 0.14 | 0.28 |
| mimo-v2.6-pro | chat-completions | no | 1 048 576 | 131 072 | 0.435 | 0.87 |
| minimax-m3 | messages | no | 1 000 000 | 131 072 | 0.30 | 1.20 |
| minimax-m3.1-flash | chat-completions | no | 1 000 000 | 512 000 | 0.15 | 0.60 |
| longcat-2.0 | chat-completions | no | 1 000 000 | 131 072 | 0.30 | 1.20 |
| step-5-preview | chat-completions | yes | 1 000 000 | 65 536 | 1.00 | 2.70 |
| hy3 | chat-completions | no | 256 000 | 128 000 | 0.14 | 0.58 |
| hy4 | chat-completions | no | 1 048 576 | 65 536 | 0.834 | 2.501 |
| hy4-preview | chat-completions | no | 1 024 000 | 64 000 | 0.834 | 2.501 |
| muse-spark-1.2-contributor | responses | yes | 1 048 576 | 131 072 | 0.10 | 0.20 |
| muse-spark-1.3-contributor | responses | yes | 1 048 576 | 131 072 | 0.10 | 0.20 |
| gemini-4-argon | generate-content | yes | 1 048 576 | 65 536 | 2.00 | 10.00 |
| beam-free | chat-completions | yes | 1 048 576 | 131 072 | 0 | 0 |
| crouton-alpha-free | responses | yes | 1 048 576 | 270 000 | 0 | 0 |
| iceman-contributor-free | chat-completions | no | 262 144 | 131 072 | 0 | 0 |

Source: one `GET /models`; every field above comes from `limit.context`, `limit.output`,
`cost.text.*.price`, `capabilities.structured_output`, and the single key of `api`.
`[confirmed]` Context is not the practical limit for this add-in: the bridge already caps the
snapshot at ~45 K characters (`bridge.py:74-78`), roughly 11–15 K tokens.

Suitability for an Excel tool loop, from capability metadata plus observed behaviour:

- **Best default:** `deepseek-v4.1-flash` — 1 M context, tools with `required`/`tool` choice,
  structured output, 0.15/0.60 pricing, described as "reasoning and agentic coding". It is
  already the bridge's first fallback. `[confirmed]` capabilities; tool quality on this model
  `[unknown]` (no paid call made).
- **Cheap alternative:** `glm-5.3-flash` (same price band, 1 M, structured output) but
  `tool_choice` is `auto` only, so it cannot be forced to call a tool. `[confirmed]`
- **Free:** `beam-free` — strict schemas, `required`/`tool` choice, 1 M context, and the only
  model whose streaming/tool/structured paths were exercised here. Good for development and
  demos; "free for a limited time" and region-restricted. `[confirmed]`
- **Highest tool fidelity (strict-schema flags):** `gpt-6-luna` (0.10/0.50), `claude-haiku-5-5`
  (0.10/0.50), `kimi-k3`/`kimi-k3.1-pro` (expensive), at the cost of higher prices and, for the
  Responses/Messages families, translation through `/chat/completions`. `[confirmed]` flags;
  fidelity `[unknown]`.
- **Avoid for forced tools:** `glm-5.2`, `glm-5.3`, `glm-5.3-flash`, `mimo-v2.6-*` (`auto` only).
- Models are region-gated: Beam Free's published exclusion list does not include Morocco, while
  Crouton Alpha Free's excludes the EU along with a similar list. `[confirmed]` docs; this
  machine's region works for both `[confirmed]` by probe/log.

## Protocol implications

**What the current protocol does.** VBA sends a plain-text body of records: `VERSION`,
`READONLY`, `MODEL`, `USER`, `ACTIVESHEET`, `SELECTION`, `SHEET`, `GRID`, `HISTORY`, `END`,
with fields joined by `0x1F` and records by `\n`, escaped as `\\ \n \r \t \s`
(`bridge.py:266-306`, `308-364`; `vba/modProto.bas:10-67`). The bridge replies with `REPLY`,
`NOTE`, `ACTION`, `END` records (`bridge.py:1110-1169`; parsed at `vba/Turn.bas:93-109`).
`ACTION` records are `op, label, …fields` (`bridge.py:989-1096`). The model never sees this
protocol: it is told to emit one JSON object `{"reply": …, "actions": […]}` (`bridge.py:755-799`),
which the bridge extracts with a balanced-brace scanner (`bridge.py:892-922`), validates, and
converts (`bridge.py:947-1102`). VBA collects `ACTION` records as pending edits and only applies
them on **Apply** (`vba/Turn.bas:104-124`, `142-212`) — the approval gate is a protocol-level
two-phase design.

**Why it exists.** VBA has no JSON parser and 64-bit Office lost `ScriptControl`, so the line
format was the cheapest thing VBA could split (`bridge.py:11-15`, `README.md:155-175`). That
constraint disappears in an Office.js task pane: JavaScript has `JSON.parse` and `fetch`.

**What native tool calls change.**

1. **Request shape.** The bridge gains `tools`, `tool_choice`, and optionally
   `parallel_tool_calls`; `max_tokens: 3000` becomes a per-model output cap. The system prompt's
   JSON contract (`bridge.py:760-799`) is replaced by tool schemas.
2. **Response shape.** A tool turn returns `content: null`, `tool_calls[]`, and
   `finish_reason: "tool_calls"`. The current `content`-or-`reasoning_content` read
   (`bridge.py:865-868`) and `"empty completion"` fallback (`bridge.py:871`) would misfire;
   the parser must branch on `finish_reason`.
3. **Loop ownership.** This is the real architectural change. The model asks for a tool; only
   the Excel side can execute it. Today's single POST (`bridge.py:829-889`) becomes N requests:
   model → tool_calls → execute in the pane → `role:"tool"` results → model again until `stop`.
   The bridge cannot execute Excel operations, so the loop must live in the pane (or a backend
   that owns both the model and an Excel channel). The bridge becomes a proxy, not an agent.
4. **Approval gate.** A tool call is a natural proposal: intercept before `Excel.run`, show the
   diff, execute only on approval, keep the current "Apply / Discard" semantics. A dry-run mode
   would have to be a tool-level convention (return a preview instead of writing), because the
   model cannot know the user approved.
5. **Streaming.** The line protocol has no incremental frame. Streaming needs either SSE passed
   through to the pane (simplest: `fetch` + `ReadableStream`, since `EventSource` cannot POST),
   or new record types (`DELTA`, `TOOL_CALL`, `TOOL_RESULT`) if VBA were kept. The pane renders
   token deltas; the `[DONE]` + usage chunk closes the turn.
6. **Structured outputs.** `response_format: json_schema` can replace `extract_json` and the
   "return exactly one JSON object" prompt discipline for the final answer, but not for tool
   arguments (those are already schema-constrained by the tool definition).
7. **State.** `x-opencode-session` is one UUID per bridge process (`bridge.py:825`), shared by
   every conversation. The docs ask for a stable session id **per conversation**; a native loop
   should send one per chat to help routing and prompt caching.

## Client placement

**Browser task pane → gateway directly: blocked, and wrong on secrets.**

- `[confirmed]` `OPTIONS /zen/go/v1/chat/completions` (with `Origin` and
  `Access-Control-Request-Method: POST`) returns **404** with no CORS headers; a `GET` on the
  same path is also 404. A real `POST` with an `Origin` header returns 200 **without**
  `Access-Control-Allow-Origin`. So a page cannot call the completion endpoint cross-origin.
- `[confirmed]` `OPTIONS /zen/go/v1/models` returns 200 with
  `Access-Control-Allow-Origin: *`, `Access-Control-Allow-Headers: Content-Type, Authorization`,
  `Access-Control-Allow-Methods: GET, POST, OPTIONS`. Only the catalog is browser-readable.
- Key exposure: the key would live in pane JavaScript, be visible in WebView2 devtools/network,
  and be reachable by any script injected into the add-in. It is a subscription credential with
  a monthly dollar allowance. `[confirmed]` as a risk; not acceptable for a personal machine
  that also runs other add-ins.

**Browser task pane → local backend → gateway: required.** `[confirmed]` The current bridge is
not CORS-ready: it sends only `Content-Type` and `Content-Length` (`bridge.py:1178-1184`), has no
`do_OPTIONS`, and returns **501** for `OPTIONS /chat` (tested against `--port 8799`). Any
Office.js pane on an `https://…` origin will fail preflight against `http://127.0.0.1:8765`
until the bridge adds `Access-Control-Allow-Origin` (the pane's exact origin, not `*` if
credentials are ever used), `Access-Control-Allow-Headers: content-type`, and an `OPTIONS`
handler. Chromium treats loopback HTTP as potentially trustworthy, so this is CORS, not mixed
content; still worth confirming inside the WebView2 pane in the sideload spike (ticket 05).
`[inferred]` from browser security rules.

**Latency.** `[confirmed]` Gateway time-to-first-byte on a free model was ≈0.9 s and a short
non-streamed turn ≈1.1 s; real bridge chats in `bridge.log` run 1–15 s (one 79 s outlier),
and the fallback path once took 58.6 s (`bridge.log:13:14:08`). A localhost hop is sub-millisecond
and irrelevant next to that. Streaming improves *perceived* latency (first token in ~1 s instead
of a 5–15 s blank panel), not total time. The bridge currently destroys that benefit by buffering
the whole body (`bridge.py:863`, `vba/Turn.bas:85`).

## Recommendation

**Native tool calling is viable; adopt it in the Office.js pane, keep the local backend, and drop
the line protocol there.** The gateway side is not the risk: tools, parallel calls, streaming, and
structured outputs all work on the same key and the same `/chat/completions` endpoint, and the
cheapest useful models (`deepseek-v4.1-flash`, `glm-5.3-flash`, free `beam-free`) have 1 M-token
contexts and tool support. The gap is the client: nothing in this repo issues a tool request, no
code consumes SSE, and the only executor is VBA behind a two-phase `ACTION` protocol.

Concretely:

1. **Keep the local backend** (`127.0.0.1`) as the only holder of the API key, and add CORS plus
   `OPTIONS` for the add-in origin. Browser-direct calls are blocked by the gateway today and
   would expose the key.
2. **Put the agent loop in the pane**, not the backend: the model asks, the pane executes with
   `Excel.run`, the backend relays `role:"tool"` results. The backend stays stateless per request.
3. **Stream SSE end-to-end** (`fetch` + `ReadableStream`), ignoring null `id`/`name` fields in
   continuation deltas and consuming the final `choices: []` usage chunk.
4. **Define a small tool surface** mirroring today's ops (`read_range`, `write_range`,
   `set_formula`, `format_range`, `sort`, `clear`, `create_sheet`, `select_range`), strict schemas
   where the model advertises `strict`, and `parallel_tool_calls: true` for read-only batches.
5. **Preserve the approval gate** by executing write tools only after Apply; reads can run
   immediately.
6. **Model policy:** default `deepseek-v4.1-flash`; fall back `glm-5.3-flash` → `qwen3.8-flash` →
   `beam-free`, matching the existing order; use `beam-free` for development because its
   tool/stream/JSON paths are the ones actually observed here. Send a per-conversation
   `x-opencode-session`.
7. **For the existing VBA add-in, change nothing**: its custom JSON-action protocol is the right
   shape for a client that cannot parse JSON or hold a socket open, and it already has the
   two-phase approval that native tools would have to re-create.

Remaining unknowns worth a spike, not a rewrite: tool-call fidelity of Responses/Messages-native
models through `/chat/completions`, DeepSeek `json_schema` acceptance, the 429/quota-exhausted
error body, and cancellation semantics.

## Sources

Local, primary:

- [`bridge.py`](../../../bridge.py) — gateway config `50-72`; credentials `120-142`; catalog
  `156-262`; protocol `266-364`; prompt `755-818`; gateway client `829-889`; JSON extraction
  `892-922`; actions `947-1102`; request handling `1110-1216`.
- [`vba/modProto.bas`](../../../vba/modProto.bas) — escaping and record parsing, `1-94`.
- [`vba/modHttp.bas`](../../../vba/modHttp.bas) — Timeouts `59`, `76`, `99`; POST `90-128`.
- [`vba/modModels.bas`](../../../vba/modModels.bas) — default model `12`; catalog load `131-192`.
- [`vba/Turn.bas`](../../../vba/Turn.bas) — one-shot POST and parse `84-113`; apply `142-212`.
- [`README.md`](../../../README.md) — design notes, gateway/model picker, troubleshooting `204-232`.
- `bridge.log` — observed 400 shape line 71, timeout line 41, model/fallback history.
- `GET https://opencode.ai/zen/go/v1/models` (2026-11-07) — capabilities, limits, cost, `api`
  metadata for all 30 models.

External, primary:

- OpenCode Zen (Console) docs — models, endpoints, translation between APIs, models endpoint:
  <https://opencode.ai/docs/zen/>
- OpenCode Go docs — plans, per-model monthly limits and 5-hour/weekly buckets, session header
  requirement, validated clients, usage/`opencode.allowance`, region/privacy notes:
  <https://opencode.ai/docs/go/>
- Gateway issue with raw SSE tool-call chunks (explicit-null continuation deltas):
  <https://github.com/anomalyco/opencode/issues/43328>
- Third-party report of `json_schema` rejection on DeepSeek routes (unverified here):
  <https://github.com/lidge-jun/opencodex/pull/4274>
