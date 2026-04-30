const fs = require("fs");
const http = require("http");
const path = require("path");
const { URL } = require("url");
const { getBearerToken, resolveMicrosoftUser } = require("./lib/auth");
const {
  MAX_JSON_BODY_BYTES,
  assertBodyWithinLimit,
  buildSecurityHeaders,
  checkRateLimit,
  getClientAddress,
  assertAllowedOrigin,
  toPublicErrorMessage,
  sanitizeLaneHistoryEntry,
  validateParseLanePayload,
} = require("./lib/security");
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
    ...buildSecurityHeaders(),
    "Content-Type": contentType,
  });
  res.end(body);
}

function readRequestBody(req, maxBytes = MAX_JSON_BODY_BYTES) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let totalBytes = 0;
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("data", (chunk) => {
      totalBytes += chunk.length;
      if (totalBytes > maxBytes) {
        const error = new Error(`Request body too large. Limit is ${maxBytes} bytes.`);
        error.statusCode = 413;
        req.destroy(error);
      }
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function sendConfigError(res) {
  send(
    res,
    503,
    JSON.stringify({
      error: toPublicErrorMessage(
        "Lane history storage is not configured. Set TURSO_DATABASE_URL and TURSO_AUTH_TOKEN.",
        "Lane history storage is not configured."
      ),
    })
  );
}

async function resolveAuthenticatedUser(req, res) {
  const accessToken = getBearerToken(req.headers);
  if (!accessToken) {
    send(res, 401, JSON.stringify({ error: "Microsoft sign-in token is required." }));
    return null;
  }

  try {
    return await resolveMicrosoftUser(accessToken);
  } catch (error) {
    send(
      res,
      error.statusCode || 401,
      JSON.stringify({
        error: error.statusCode >= 500 ? "Microsoft authentication failed." : (error.message || "Microsoft authentication failed."),
      })
    );
    return null;
  }
}

function enforceRateLimit(req, res, scope, limit, windowMs) {
  const clientAddress = getClientAddress(req.headers, req.socket && req.socket.remoteAddress);
  const result = checkRateLimit({
    key: `${scope}::${clientAddress}`,
    limit,
    windowMs,
  });

  if (result.allowed) {
    return true;
  }

  res.setHeader("Retry-After", String(result.retryAfterSeconds));
  send(res, 429, JSON.stringify({ error: "Too many requests. Please try again shortly." }));
  return false;
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
      assertAllowedOrigin(req.headers, req.headers.host);
      if (!enforceRateLimit(req, res, "parse-lane", 45, 60 * 1000)) {
        return;
      }

      const rawBody = await readRequestBody(req);
      assertBodyWithinLimit(req.headers, rawBody);
      const payload = rawBody ? JSON.parse(rawBody) : {};
      const { messages, lane } = validateParseLanePayload(payload);

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
      const statusCode = error.statusCode || 500;
      send(
        res,
        statusCode,
        JSON.stringify({
          error: statusCode >= 500 ? "Lane parsing failed." : (error.message || "Lane parsing failed."),
        })
      );
    }
    return;
  }

  if (requestUrl.pathname === "/api/lane-history") {
    if (!hasHistoryStoreConfig()) {
      sendConfigError(res);
      return;
    }

    try {
        assertAllowedOrigin(req.headers, req.headers.host);
        if (!enforceRateLimit(req, res, "lane-history", 90, 60 * 1000)) {
          return;
        }

        const authenticatedUser = await resolveAuthenticatedUser(req, res);
        if (!authenticatedUser) {
          return;
        }

        if (req.method === "GET") {
          const userKey = authenticatedUser.userKey;
          const environment = String(requestUrl.searchParams.get("environment") || "production").trim() || "production";

          const items = await listLaneHistory(userKey, environment);
          send(res, 200, JSON.stringify({ items, debug: { action: "list", userKey, environment, count: items.length } }));
          return;
        }

      const rawBody = await readRequestBody(req);
      assertBodyWithinLimit(req.headers, rawBody);
      const payload = rawBody ? JSON.parse(rawBody) : {};
      const userKey = authenticatedUser.userKey;
      const userEmail = authenticatedUser.userEmail;
      const environment = String(payload.environment || "production").trim() || "production";

      if (req.method === "POST") {
        const skipItems = Boolean(payload.skipItems);
        const entry = sanitizeLaneHistoryEntry(payload.entry || {}, userKey, environment);

        await saveLaneHistory({
          id: entry.id,
          userKey,
          userEmail,
          environment,
          carrierName: entry.carrierName,
          carrierEmail: entry.carrierEmail,
          carrierPhone: entry.carrierPhone,
          mcNumber: entry.mcNumber,
          pickupZip: entry.pickupZip,
          deliveryZip: entry.deliveryZip,
          shipperRate: entry.shipperRate,
          quotedRate: entry.quotedRate,
          negotiatedRate: entry.negotiatedRate,
          finalRate: entry.finalRate,
          savedAt: entry.savedAt,
          sourceMessageId: entry.sourceMessageId,
          sourceConversationId: entry.sourceConversationId,
        });

        if (skipItems) {
          send(res, 200, JSON.stringify({ ok: true }));
          return;
        }

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
              sourceConversationId: entry.sourceConversationId,
              sourceMessageId: entry.sourceMessageId,
              entryId: entry.id,
            },
          })
        );
        return;
      }

      if (req.method === "DELETE") {
        const skipItems = Boolean(payload.skipItems);
        const sourceConversationId = String(payload.sourceConversationId || "").trim();
        if (!sourceConversationId) {
          send(res, 400, JSON.stringify({ error: "sourceConversationId is required" }));
          return;
        }

        await deleteLaneHistory(userKey, environment, sourceConversationId);

        if (skipItems) {
          send(res, 200, JSON.stringify({ ok: true }));
          return;
        }

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
      const statusCode = error.statusCode || 500;
      send(
        res,
        statusCode,
        JSON.stringify({
          error:
            statusCode >= 500
              ? "Lane history request failed."
              : (error.message || "Lane history request failed."),
        })
      );
    }
    return;
  }

  serveStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`Margin Flow running at http://localhost:${PORT}`);
});
