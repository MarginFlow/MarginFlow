"use strict";

const { getBearerToken, resolveMicrosoftUser } = require("../lib/auth");
const {
  MAX_JSON_BODY_BYTES,
  assertBodyWithinLimit,
  buildSecurityHeaders,
  checkRateLimit,
  getClientAddress,
  assertAllowedOrigin,
  toPublicErrorMessage,
  sanitizeLaneHistoryEntry,
} = require("../lib/security");
const {
  hasHistoryStoreConfig,
  listLaneHistory,
  saveLaneHistory,
  deleteLaneHistory,
} = require("../lib/history-store");

function applySecurityHeaders(res) {
  const headers = buildSecurityHeaders();
  Object.entries(headers).forEach(([key, value]) => {
    res.setHeader(key, value);
  });
}

function sendConfigError(res) {
  res.status(503).json({
    error: toPublicErrorMessage(
      "Lane history storage is not configured. Set TURSO_DATABASE_URL and TURSO_AUTH_TOKEN.",
      "Lane history storage is not configured."
    ),
  });
}

async function resolveAuthenticatedUser(req, res) {
  const accessToken = getBearerToken(req.headers);
  if (!accessToken) {
    res.status(401).json({ error: "Microsoft sign-in token is required." });
    return null;
  }

  try {
    return await resolveMicrosoftUser(accessToken);
  } catch (error) {
    res.status(error.statusCode || 401).json({
      error:
        error.statusCode >= 500
          ? "Microsoft authentication failed."
          : (error.message || "Microsoft authentication failed."),
    });
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
  res.status(429).json({ error: "Too many requests. Please try again shortly." });
  return false;
}

module.exports = async (req, res) => {
  applySecurityHeaders(res);

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
      const environment = String((req.query && req.query.environment) || "production").trim() || "production";

      const items = await listLaneHistory(userKey, environment);
      res.status(200).json({ items });
      return;
    }

    const rawBody = typeof req.body === "string" ? req.body : "";
    assertBodyWithinLimit(req.headers, rawBody, MAX_JSON_BODY_BYTES);
    const payload = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
    const userKey = authenticatedUser.userKey;
    const userEmail = authenticatedUser.userEmail;
    const environment = String(payload.environment || "production").trim() || "production";

    if (req.method === "POST") {
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

      const items = await listLaneHistory(userKey, environment);
      res.status(200).json({ items });
      return;
    }

    if (req.method === "DELETE") {
      const sourceConversationId = String(payload.sourceConversationId || "").trim();
      if (!sourceConversationId) {
        res.status(400).json({ error: "sourceConversationId is required" });
        return;
      }

      await deleteLaneHistory(userKey, environment, sourceConversationId);
      const items = await listLaneHistory(userKey, environment);
      res.status(200).json({ items });
      return;
    }

    res.status(405).json({ error: "Method not allowed" });
  } catch (error) {
    const statusCode = error.statusCode || 500;
    res.status(statusCode).json({
      error:
        statusCode >= 500
          ? "Lane history request failed."
          : (error.message || "Lane history request failed."),
    });
  }
};
