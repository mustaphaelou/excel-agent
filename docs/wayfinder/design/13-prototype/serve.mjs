// Zero-dependency static server for the #13 prototype.
// Usage: node serve.mjs [port]   (defaults to 4178)
// One route: /?variant=hybrid|claude|fluent&state=...&width=...&theme=...
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const port = Number(process.argv[2] || process.env.PORT || 4178);
const host = process.env.HOST || "127.0.0.1";

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".woff2": "font/woff2",
};

function send(res, status, body, headers = {}) {
  res.writeHead(status, {
    "Cache-Control": "no-store, no-cache, must-revalidate",
    "X-Content-Type-Options": "nosniff",
    ...headers,
  });
  res.end(body);
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://${host}:${port}`);
  if (url.pathname === "/" || url.pathname === "/index.html") {
    return fs.readFile(path.join(here, "index.html"), (err, data) => {
      if (err) return send(res, 500, "index.html missing");
      send(res, 200, data, { "Content-Type": MIME[".html"] });
    });
  }
  if (url.pathname === "/health") return send(res, 200, "ok", { "Content-Type": "text/plain" });
  const rel = decodeURIComponent(url.pathname).replace(/^\/+/, "");
  const abs = path.resolve(here, rel);
  if (!abs.startsWith(`${here}${path.sep}`)) return send(res, 403, "forbidden");
  fs.readFile(abs, (err, data) => {
    if (err) return send(res, 404, "not found");
    const type = MIME[path.extname(abs).toLowerCase()] || "application/octet-stream";
    send(res, 200, data, { "Content-Type": type });
  });
});

server.listen(port, host, () => {
  console.log(`Excel Agent #13 prototype: http://${host}:${port}/?variant=hybrid&state=empty&width=380&theme=auto`);
  console.log("Variants: hybrid, claude, fluent · Scenarios: empty, loading, streaming, answer, review, partial, applied, history, settings, error");
});
