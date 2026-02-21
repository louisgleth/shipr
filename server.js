const http = require("http");
const fs = require("fs");
const path = require("path");

const HOST = "0.0.0.0";
const PORT = Number(process.env.PORT) || 4173;
const ROOT = __dirname;
const INDEX_FILE = path.join(ROOT, "index.html");

const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".geojson": "application/geo+json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".ico": "image/x-icon",
  ".pdf": "application/pdf",
  ".txt": "text/plain; charset=utf-8",
  ".html": "text/html; charset=utf-8",
};

function send(res, status, headers, body) {
  res.writeHead(status, headers);
  res.end(body);
}

function safeJoin(root, requestPath) {
  const fullPath = path.normalize(path.join(root, requestPath));
  if (!fullPath.startsWith(root)) return null;
  return fullPath;
}

function sendFile(res, filePath) {
  fs.readFile(filePath, (readError, data) => {
    if (readError) {
      send(res, 500, { "Content-Type": "text/plain; charset=utf-8" }, "Internal Server Error");
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    send(
      res,
      200,
      {
        "Content-Type": MIME_TYPES[ext] || "application/octet-stream",
        "Cache-Control": ext === ".html" ? "no-cache" : "public, max-age=3600",
      },
      data
    );
  });
}

const server = http.createServer((req, res) => {
  const requestUrl = new URL(req.url || "/", "http://localhost");
  let pathname = decodeURIComponent(requestUrl.pathname);

  if (pathname === "/") {
    sendFile(res, INDEX_FILE);
    return;
  }

  const requested = safeJoin(ROOT, pathname.replace(/^\/+/, ""));
  if (!requested) {
    send(res, 403, { "Content-Type": "text/plain; charset=utf-8" }, "Forbidden");
    return;
  }

  fs.stat(requested, (statError, stats) => {
    if (!statError && stats.isFile()) {
      sendFile(res, requested);
      return;
    }

    const normalizedPathname = pathname.replace(/\/+$/, "");
    const hasFileExtension = path.extname(normalizedPathname) !== "";
    if (hasFileExtension) {
      send(res, 404, { "Content-Type": "text/plain; charset=utf-8" }, "Not Found");
      return;
    }

    // SPA history fallback: routes like /account or /reports serve index.html.
    sendFile(res, INDEX_FILE);
  });
});

server.listen(PORT, HOST, () => {
  console.log(`Shipr app running at http://${HOST}:${PORT}`);
});
