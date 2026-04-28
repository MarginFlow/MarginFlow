const fs = require("fs");
const http = require("http");
const path = require("path");
const { URL } = require("url");
const { parseCarrierEmail, dedupeVisibleMessages } = require("./lib/lane-parser");
const {
  hasHistoryStoreConfig,
  listLaneHistory,
  saveLaneHistory,
  deleteLaneHistory,
} = require("./lib/history-store");

const PORT = Number(process.env.PORT || 3000);
const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, "public");
const MSAL_BUNDLE = path.join(ROOT, "node_modules", "@azure", "msal-browser", "lib", "msal-browser.min.js");

const CONTENT_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml; charset=utf-8",
};

function send(res, statusCode, body, contentType = "application/json; charset=utf-8") {
  res.writeHead(statusCode, {
    "Content-Type": contentType,
    "Cache-Control": "no-store",
  });
  res.end(body);
}

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function sendConfigError(res) {
  send(
    res,
    503,
    JSON.stringify({
      error: "Lane history storage is not configured. Set TURSO_DATABASE_URL and TURSO_AUTH_TOKEN.",
    })
  );
}

function parseMoneyValue(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const parsed = Number(String(value).replace(/,/g, "").replace(/\$/g, "").trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function serveStatic(req, res) {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);
  const pathname = requestUrl.pathname === "/" ? "/index.html" : requestUrl.pathname;
  const filePath = path.normalize(path.join(PUBLIC_DIR, pathname));

  if (!filePath.startsWith(PUBLIC_DIR)) {
    send(res, 403, "Forbidden", "text/plain; charset=utf-8");
    return;
  }

  fs.readFile(filePath, (error, buffer) => {
    if (error) {
      send(res, 404, "Not found", "text/plain; charset=utf-8");
      return;
    }

    send(res, 200, buffer, CONTENT_TYPES[path.extname(filePath).toLowerCase()] || "application/octet-stream");
  });
}

const server = http.createServer(async (req, res) => {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);

  if (req.method === "GET" && req.url === "/vendor/msal-browser.min.js") {
    fs.readFile(MSAL_BUNDLE, (error, buffer) => {
      if (error) {
        send(res, 404, "MSAL bundle not found", "text/plain; charset=utf-8");
        return;
      }
      send(res, 200, buffer, "application/javascript; charset=utf-8");
    });
    return;
  }

  if (req.method === "POST" && req.url === "/api/parse-lane") {
    try {
      const rawBody = await readRequestBody(req);
      const payload = JSON.parse(rawBody);
      const messages = Array.isArray(payload.messages) ? payload.messages : [];
      const lane = payload.lane || {};

      const parsed = messages.map((message) => parseCarrierEmail(message, lane));
      const visible = dedupeVisibleMessages(
        parsed.filter((item) => !item.ignored)
      )
        .sort((left, right) => {
          if (left.carrierRate === null && right.carrierRate === null) {
            return 0;
          }
          if (left.carrierRate === null) {
            return 1;
          }
          if (right.carrierRate === null) {
            return -1;
          }
          return left.carrierRate - right.carrierRate;
        });

      send(
        res,
        200,
        JSON.stringify({
          results: visible,
          hidden: parsed.filter((item) => item.ignored),
        })
      );
    } catch (error) {
      send(res, 500, JSON.stringify({ error: error.message || "Parse failed" }));
    }
    return;
  }

  if (requestUrl.pathname === "/api/lane-history") {
    if (!hasHistoryStoreConfig()) {
      sendConfigError(res);
      return;
    }

    try {
        if (req.method === "GET") {
          const userKey = String(requestUrl.searchParams.get("userKey") || "").trim();
          const environment = String(requestUrl.searchParams.get("environment") || "production").trim() || "production";
          if (!userKey) {
            send(res, 400, JSON.stringify({ error: "userKey is required" }));
            return;
          }

          const items = await listLaneHistory(userKey, environment);
          send(res, 200, JSON.stringify({ items, debug: { action: "list", userKey, environment, count: items.length } }));
          return;
        }

      const rawBody = await readRequestBody(req);
      const payload = rawBody ? JSON.parse(rawBody) : {};
      const userKey = String(payload.userKey || payload.userEmail || "local-margin-flow-user").trim();
      const environment = String(payload.environment || "production").trim() || "production";
      if (!userKey) {
        send(res, 400, JSON.stringify({ error: "userKey is required" }));
        return;
      }

      if (req.method === "POST") {
        const entry = payload.entry || {};
        const shipperRate = parseMoneyValue(entry.shipperRate);
        const quotedRate = parseMoneyValue(entry.quotedRate);
        const negotiatedRate = parseMoneyValue(entry.negotiatedRate);
        const finalRate = parseMoneyValue(entry.finalRate);
        const fallbackFinalRate = finalRate ?? negotiatedRate ?? quotedRate ?? 0;
        const sourceConversationId = String(entry.sourceConversationId || entry.sourceMessageId || Date.now());
        const sourceMessageId = entry.sourceMessageId ? String(entry.sourceMessageId) : sourceConversationId;
        const entryId = String(entry.id || `${userKey}::${sourceConversationId}`);

        await saveLaneHistory({
          id: entryId,
          userKey,
          userEmail: String(payload.userEmail || ""),
          environment,
          carrierName: String(entry.carrierName || "Unknown carrier"),
          carrierEmail: String(entry.carrierEmail || ""),
          carrierPhone: entry.carrierPhone ? String(entry.carrierPhone) : null,
          mcNumber: entry.mcNumber ? String(entry.mcNumber) : null,
          pickupZip: entry.pickupZip ? String(entry.pickupZip) : null,
          deliveryZip: entry.deliveryZip ? String(entry.deliveryZip) : null,
          shipperRate,
          quotedRate,
          negotiatedRate,
          finalRate: fallbackFinalRate,
          savedAt: String(entry.savedAt || new Date().toISOString()),
          sourceMessageId,
          sourceConversationId,
        });

        const items = await listLaneHistory(userKey, environment);
        send(
          res,
          200,
          JSON.stringify({
            items,
            debug: {
              action: "save",
              userKey,
              environment,
              count: items.length,
              sourceConversationId,
              sourceMessageId,
              entryId,
            },
          })
        );
        return;
      }

      if (req.method === "DELETE") {
        const sourceConversationId = String(payload.sourceConversationId || "").trim();
        if (!sourceConversationId) {
          send(res, 400, JSON.stringify({ error: "sourceConversationId is required" }));
          return;
        }

        await deleteLaneHistory(userKey, environment, sourceConversationId);
        const items = await listLaneHistory(userKey, environment);
        send(
          res,
          200,
          JSON.stringify({
            items,
            debug: { action: "delete", userKey, environment, count: items.length, sourceConversationId },
          })
        );
        return;
      }

      send(res, 405, JSON.stringify({ error: "Method not allowed" }));
    } catch (error) {
      send(res, 500, JSON.stringify({ error: error.message || "Lane history request failed" }));
    }
    return;
  }

  serveStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`Margin Flow running at http://localhost:${PORT}`);
});
