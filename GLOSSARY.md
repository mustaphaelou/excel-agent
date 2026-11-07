# Excel Agent

An Excel add-in: a docked task pane that chats with a model, reads the workbook, and edits it only after approval. The shipped VBA add-in and the Office.js replacement share this vocabulary.

## Language

**Pane**:
The Office.js task pane for one workbook. It is document-scoped: each workbook gets its own instance, and there is no cross-workbook sidebar.
_Avoid_: sidebar (when the document scope matters), panel

**Backend**:
The local Node server at `https://localhost:3000` that serves the pane and proxies model and web-fetch calls. It is stateless and holds the gateway key.
_Avoid_: bridge, proxy (for the whole backend)

**Bridge**:
The legacy Python process (`bridge.py`) that served the VBA add-in. It is reference material, not part of the new runtime.
_Avoid_: backend

**Agent loop**:
The cycle in the pane where the model asks for a tool, the pane executes it, the result goes back to the model, and the model continues until it stops.
_Avoid_: agent (for the loop), orchestration

**Turn**:
One user message through the streamed reply and any tool calls, ending when the model stops.
_Avoid_: request, exchange

**Tool call**:
A structured request from the model to read or write Excel or to fetch a URL. Excel tools execute in the pane; the web fetch executes in the backend.
_Avoid_: action, command

**Approval gate**:
The pane's review step between a proposed write and the write landing. Reads do not pass through it.
_Avoid_: confirmation dialog, apply step

**Citation**:
A clickable cell or range reference in a reply that navigates the workbook to that location.
_Avoid_: reference, link

**Session**:
The conversation context owned by one pane instance, identified by a per-pane session id. It is not server state.
_Avoid_: conversation (when the storage boundary matters)
