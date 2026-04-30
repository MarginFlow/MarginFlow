"use strict";

const {
  MAX_JSON_BODY_BYTES,
  assertBodyWithinLimit,
  buildSecurityHeaders,
  checkRateLimit,
  getClientAddress,
  assertAllowedOrigin,
  validateParseLanePayload,
} = require("../lib/security");
const { parseCarrierEmail, dedupeVisibleMessages } = require("../lib/lane-parser");

function applySecurityHeaders(res) {
  const headers = buildSecurityHeaders();
  Object.entries(headers).forEach(([key, value]) => {
    res.setHeader(key, value);
  });
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

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    assertAllowedOrigin(req.headers, req.headers.host);
    if (!enforceRateLimit(req, res, "parse-lane", 45, 60 * 1000)) {
      return;
    }

    const rawBody = typeof req.body === "string" ? req.body : "";
    assertBodyWithinLimit(req.headers, rawBody, MAX_JSON_BODY_BYTES);
    const payload = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
    const { messages, lane } = validateParseLanePayload(payload);

    const parsed = messages.map((message) => parseCarrierEmail(message, lane));
    const visible = dedupeVisibleMessages(
      parsed.filter((item) => !item.ignored)
    ).sort((left, right) => {
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

    res.status(200).json({
      results: visible,
      hidden: parsed.filter((item) => item.ignored),
    });
  } catch (error) {
    const statusCode = error.statusCode || 500;
    res.status(statusCode).json({
      error: statusCode >= 500 ? "Lane parsing failed." : (error.message || "Lane parsing failed."),
    });
  }
};
