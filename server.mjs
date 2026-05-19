import crypto from "node:crypto";
import fs from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { buildWorkbookFromPayload } from "./src/reporting.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicDir = path.join(__dirname, "public");
const port = Number(process.env.PORT || 3030);
const downloadTtlMs = 10 * 60 * 1000;
const downloads = new Map();

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

function json(res, status, payload) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(payload));
}

function cleanupDownloads() {
  const now = Date.now();
  for (const [token, item] of downloads.entries()) {
    if (item.expiresAt <= now) {
      downloads.delete(token);
    }
  }
}

function storeDownload(result) {
  cleanupDownloads();
  const token = crypto.randomUUID();
  const expiresAt = Date.now() + downloadTtlMs;
  downloads.set(token, {
    bytes: result.bytes,
    contentType: result.contentType || mimeTypes[".xlsx"],
    expiresAt,
    fileName: result.fileName,
  });

  return {
    token,
    expiresAt,
  };
}

function encodeDownloadName(fileName) {
  return encodeURIComponent(fileName).replace(/['()*]/g, (char) => `%${char.charCodeAt(0).toString(16).toUpperCase()}`);
}

function contentDisposition(fileName) {
  const fallbackName = fileName.replace(/[^\x20-\x7E]/g, "_");
  return `attachment; filename="${fallbackName}"; filename*=UTF-8''${encodeDownloadName(fileName)}`;
}

async function serveStatic(pathname, res) {
  const requestedPath = pathname === "/" ? "/index.html" : pathname;
  const normalizedPath = path.normalize(requestedPath).replace(/^(\.\.(\/|\\|$))+/, "");
  const relativePath = normalizedPath.replace(/^[/\\]+/, "");
  const filePath = path.join(publicDir, relativePath);

  if (!filePath.startsWith(publicDir)) {
    json(res, 403, { error: "Forbidden" });
    return;
  }

  try {
    const ext = path.extname(filePath);
    const content = await fs.readFile(filePath);
    res.writeHead(200, {
      "Content-Type": mimeTypes[ext] || "application/octet-stream",
    });
    res.end(content);
  } catch {
    json(res, 404, { error: "Not found" });
  }
}

async function readJsonBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }

  const raw = Buffer.concat(chunks).toString("utf8");
  return raw ? JSON.parse(raw) : {};
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);
    const { pathname } = url;

    if (req.method === "GET" && pathname === "/api/health") {
      json(res, 200, { ok: true });
      return;
    }

    if (req.method === "POST" && pathname === "/api/generate") {
      const payload = await readJsonBody(req);
      const result = await buildWorkbookFromPayload(payload);
      const download = storeDownload(result);

      json(res, 200, {
        downloadUrl: `/api/download/${download.token}`,
        expiresAt: download.expiresAt,
        fileName: result.fileName,
        summary: result.summary,
      });
      return;
    }

    if (req.method === "GET" && pathname.startsWith("/api/download/")) {
      cleanupDownloads();
      const token = pathname.slice("/api/download/".length);
      const item = downloads.get(token);

      if (!item) {
        json(res, 404, {
          error: "다운로드 링크가 만료되었거나 존재하지 않습니다. 다시 생성해 주세요.",
        });
        return;
      }

      res.writeHead(200, {
        "Cache-Control": "no-store",
        "Content-Disposition": contentDisposition(item.fileName),
        "Content-Length": item.bytes.length,
        "Content-Type": item.contentType,
      });
      res.end(item.bytes);
      return;
    }

    if (req.method === "GET") {
      await serveStatic(pathname, res);
      return;
    }

    json(res, 404, { error: "Not found" });
  } catch (error) {
    json(res, 500, {
      error: error instanceof Error ? error.message : "Unexpected error",
    });
  }
});

setInterval(cleanupDownloads, 60 * 1000).unref();

server.listen(port, () => {
  console.log(`Bymom report generator running on http://localhost:${port}`);
});
