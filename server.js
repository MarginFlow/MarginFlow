const fs = require("fs");
const http = require("http");
const path = require("path");
const { URL } = require("url");

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

function normalizeZip(value) {
  if (value === null || value === undefined) {
    return "";
  }

  const digits = String(value).match(/\d/g);
  return digits ? digits.join("").slice(0, 5) : "";
}

function normalizeMoney(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const parsed = Number(String(value).replace(/,/g, "").replace(/\$/g, "").trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function roundTo(value, digits) {
  if (!Number.isFinite(value)) {
    return null;
  }

  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function extractMcNumber(text) {
  const match = String(text || "").match(/\bMC\s*#?\s*-?\s*(\d{4,8})\b/i);
  return match ? match[1] : null;
}

function extractRateCandidates(text) {
  const candidates = [];
  const source = String(text || "");
  const patterns = [
    {
      type: "per_mile",
      regex: /(?:rate|can|do|cover|all in|offer|quote|my rate)?\s*[:=-]?\s*\$?\s*((?:\d{1,3}(?:,\d{3})+)|(?:\d{1,6}))(?:\.(\d{1,3}))?\s*(?:\/\s*(?:mi|mile)|per\s*mile|rpm)\b/gi,
      baseScore: 10,
    },
    {
      type: "total",
      regex: /(?:rate(?: is)?|my rate(?: is)?|can do|can|cover|all in|offer|quote|for|flat|do)\s*[:=-]?\s*\$?\s*((?:\d{1,3}(?:,\d{3})+)|(?:\d{2,6}))(?:\.(\d{1,2}))?\b/gi,
      baseScore: 12,
    },
    {
      type: "total",
      regex: /\$\s*((?:\d{1,3}(?:,\d{3})+)|(?:\d{2,6}))(?:\.(\d{1,2}))?\b/gi,
      baseScore: 7,
    },
    {
      type: "total",
      regex: /\b((?:\d{1,3}(?:,\d{3})+)|(?:\d{2,6}))(?:\.(\d{1,2}))?\b/gi,
      baseScore: 2,
    },
  ];

  patterns.forEach((pattern) => {
    let match;
    while ((match = pattern.regex.exec(source)) !== null) {
      const wholePart = match[1] || "";
      const decimalPart = match[2] ? `.${match[2]}` : "";
      const value = normalizeMoney(`${wholePart}${decimalPart}`);
      const matchedText = match[0] || "";
      const matchStart = match.index;
      const matchEnd = match.index + matchedText.length;
      const prevChar = matchStart > 0 ? source[matchStart - 1] : "";
      const nextChar = matchEnd < source.length ? source[matchEnd] : "";
      if (!Number.isFinite(value) || value <= 1) {
        continue;
      }
      if (prevChar === "," || nextChar === ",") {
        continue;
      }

      const context = source
        .slice(Math.max(0, match.index - 36), Math.min(source.length, pattern.regex.lastIndex + 36))
        .toLowerCase();

      let score = pattern.baseScore;
      if (/\brate\b|\bcan\b|\bdo\b|\bcover\b|\ball in\b|\bquote\b|\boffer\b/.test(context)) {
        score += 4;
      }
      if (/\bmc\b|\bdot\b|\bphone\b|\bref\b|\bzip\b/.test(context)) {
        score -= 6;
      }
      if (/\bpickup\b|\bdeliver\b|\bdrop\b/.test(context)) {
        score -= 3;
      }
      if (value >= 100 && value <= 4000) {
        score += 2;
      }
      if (value < 100 && pattern.type === "total") {
        score -= 5;
      }
      if (value > 10000) {
        score -= 10;
      }

      candidates.push({
        type: pattern.type,
        value,
        score,
        index: match.index,
      });
    }
  });

  return candidates.sort((left, right) => {
    if (right.score !== left.score) {
      return right.score - left.score;
    }
    return right.index - left.index;
  });
}

function detectLaneMatch(text, pickupZip, deliveryZip) {
  const source = String(text || "").toLowerCase();
  const pickup = normalizeZip(pickupZip);
  const delivery = normalizeZip(deliveryZip);
  const pickupMatch = pickup ? source.includes(pickup) : false;
  const deliveryMatch = delivery ? source.includes(delivery) : false;

  if (pickup && delivery) {
    return pickupMatch || deliveryMatch
      ? pickupMatch && deliveryMatch
      : false;
  }

  if (pickup) {
    return pickupMatch;
  }

  if (delivery) {
    return deliveryMatch;
  }

  return false;
}

function classifyMargin(marginPercent) {
  if (!Number.isFinite(marginPercent)) {
    return "UNKNOWN";
  }
  if (marginPercent >= 15) {
    return "GREEN";
  }
  if (marginPercent >= 5) {
    return "YELLOW";
  }
  return "RED";
}

function parseCarrierEmail(message, lane) {
  const shipperRate = normalizeMoney(lane.shipperRate);
  const pickupZip = normalizeZip(lane.pickupZip);
  const deliveryZip = normalizeZip(lane.deliveryZip);
  const body = message.bodyText || "";
  const haystack = [
    message.subject || "",
    message.fromName || "",
    message.fromAddress || "",
    body,
  ].join("\n");

  const laneMatch = detectLaneMatch(haystack, pickupZip, deliveryZip);
  const rateCandidates = extractRateCandidates(haystack);
  const bestRate = rateCandidates[0] || null;
  const carrierRate = bestRate ? roundTo(bestRate.value, 2) : null;
  const marginPercent =
    Number.isFinite(shipperRate) && shipperRate > 0 && Number.isFinite(carrierRate)
      ? roundTo(((shipperRate - carrierRate) / shipperRate) * 100, 2)
      : null;
  const ignored = !laneMatch || (Number.isFinite(marginPercent) && marginPercent <= -50);

  return {
    id: message.id,
    conversationId: message.conversationId || "",
    webLink: message.webLink || "",
    subject: message.subject || "",
    receivedDateTime: message.receivedDateTime || "",
    fromName: message.fromName || "",
    fromAddress: message.fromAddress || "",
    carrierName: message.fromName || message.fromAddress || "Unknown carrier",
    mcNumber: extractMcNumber(haystack),
    carrierRate,
    marginPercent,
    classification: classifyMargin(marginPercent),
    laneMatch,
    ignored,
    ignoreReason: !laneMatch
      ? "Wrong lane"
      : Number.isFinite(marginPercent) && marginPercent <= -50
        ? "Margin <= -50%"
        : null,
    bodyPreview: body.slice(0, 280),
  };
}

function compareReceivedDate(left, right) {
  const leftTime = left && left.receivedDateTime ? Date.parse(left.receivedDateTime) : 0;
  const rightTime = right && right.receivedDateTime ? Date.parse(right.receivedDateTime) : 0;
  return leftTime - rightTime;
}

function pickRepresentativeMessage(existing, candidate) {
  if (!existing) {
    return candidate;
  }

  const existingScore = Number.isFinite(existing.carrierRate) ? existing.carrierRate : Number.POSITIVE_INFINITY;
  const candidateScore = Number.isFinite(candidate.carrierRate) ? candidate.carrierRate : Number.POSITIVE_INFINITY;

  if (compareReceivedDate(candidate, existing) < 0) {
    return candidate;
  }

  if (compareReceivedDate(candidate, existing) === 0 && candidateScore < existingScore) {
    return candidate;
  }

  return existing;
}

function dedupeVisibleMessages(items) {
  const byConversation = new Map();

  items.forEach((item) => {
    const key = item.conversationId || item.id;
    byConversation.set(key, pickRepresentativeMessage(byConversation.get(key), item));
  });

  return Array.from(byConversation.values());
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

  serveStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`Margin Flow running at http://localhost:${PORT}`);
});
