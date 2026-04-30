"use strict";

const MAX_JSON_BODY_BYTES = 256 * 1024;
const MAX_MESSAGES_PER_PARSE = 250;

const rateLimitBuckets = new Map();

function isProductionRuntime() {
  return process.env.NODE_ENV === "production" || process.env.VERCEL_ENV === "production";
}

function toPublicErrorMessage(message, fallback) {
  return isProductionRuntime() ? fallback : message;
}

function buildSecurityHeaders() {
  return {
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "strict-origin-when-cross-origin",
    "X-Frame-Options": "DENY",
    "Permissions-Policy": "camera=(), microphone=(), geolocation=()",
    "Content-Security-Policy":
      "default-src 'self'; script-src 'self' https://cdn.vercel-insights.com; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self' https://graph.microsoft.com https://login.microsoftonline.com https://vitals.vercel-insights.com; frame-ancestors 'none'; base-uri 'self'; form-action 'self'",
  };
}

function getClientAddress(headersLike, fallback = "") {
  const forwardedFor =
    (headersLike && typeof headersLike.get === "function" && headersLike.get("x-forwarded-for")) ||
    (headersLike && headersLike["x-forwarded-for"]) ||
    "";
  return String(forwardedFor || fallback || "unknown")
    .split(",")[0]
    .trim()
    .slice(0, 120);
}

function checkRateLimit({ key, limit, windowMs }) {
  const now = Date.now();
  const bucketKey = `${key}::${windowMs}`;
  const current = rateLimitBuckets.get(bucketKey);
  if (!current || current.resetAt <= now) {
    rateLimitBuckets.set(bucketKey, { count: 1, resetAt: now + windowMs });
    return { allowed: true };
  }

  if (current.count >= limit) {
    return {
      allowed: false,
      retryAfterSeconds: Math.max(1, Math.ceil((current.resetAt - now) / 1000)),
    };
  }

  current.count += 1;
  return { allowed: true };
}

function assertAllowedOrigin(headersLike, host) {
  const origin =
    (headersLike && typeof headersLike.get === "function" && headersLike.get("origin")) ||
    (headersLike && headersLike.origin) ||
    "";
  if (!origin) {
    return;
  }

  let parsedOrigin;
  try {
    parsedOrigin = new URL(String(origin));
  } catch (error) {
    const invalidOriginError = new Error("Invalid request origin.");
    invalidOriginError.statusCode = 403;
    throw invalidOriginError;
  }

  const normalizedHost = String(host || "").trim().toLowerCase();
  if (!normalizedHost) {
    return;
  }

  if (parsedOrigin.host.toLowerCase() !== normalizedHost) {
    const mismatchError = new Error("Cross-origin requests are not allowed.");
    mismatchError.statusCode = 403;
    throw mismatchError;
  }
}

function sanitizeString(value, maxLength, fallback = "") {
  const normalized = String(value == null ? fallback : value).trim().replace(/\s+/g, " ");
  return normalized.slice(0, maxLength);
}

function sanitizeOptionalString(value, maxLength) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const normalized = sanitizeString(value, maxLength, "");
  return normalized || null;
}

function sanitizeEmail(value) {
  const normalized = sanitizeOptionalString(value, 320);
  if (!normalized) {
    return "";
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized)) {
    return "";
  }

  return normalized.toLowerCase();
}

function sanitizeZip(value) {
  const normalized = sanitizeOptionalString(value, 10);
  if (!normalized) {
    return null;
  }

  if (!/^\d{5}(?:-\d{4})?$/.test(normalized)) {
    return null;
  }

  return normalized;
}

function sanitizeMcNumber(value) {
  const normalized = sanitizeOptionalString(value, 16);
  if (!normalized) {
    return null;
  }

  const digitsOnly = normalized.replace(/[^\d]/g, "");
  return digitsOnly ? digitsOnly.slice(0, 12) : null;
}

function sanitizePhone(value) {
  const normalized = sanitizeOptionalString(value, 32);
  if (!normalized) {
    return null;
  }

  return normalized.replace(/[^\d+\-(). xext]/gi, "").slice(0, 32) || null;
}

function sanitizeIdentifier(value, maxLength = 160) {
  const normalized = sanitizeString(value, maxLength, "");
  return normalized || "";
}

function parseMoneyValue(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const parsed = Number(String(value).replace(/,/g, "").replace(/\$/g, "").trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function assertBodyWithinLimit(headersLike, rawBody, maxBytes = MAX_JSON_BODY_BYTES) {
  const contentLengthHeader =
    (headersLike && typeof headersLike.get === "function" && headersLike.get("content-length")) ||
    (headersLike && headersLike["content-length"]) ||
    "";
  const declaredLength = Number(contentLengthHeader);
  if (Number.isFinite(declaredLength) && declaredLength > maxBytes) {
    const error = new Error(`Request body too large. Limit is ${maxBytes} bytes.`);
    error.statusCode = 413;
    throw error;
  }

  if (typeof rawBody === "string" && Buffer.byteLength(rawBody, "utf8") > maxBytes) {
    const error = new Error(`Request body too large. Limit is ${maxBytes} bytes.`);
    error.statusCode = 413;
    throw error;
  }
}

function sanitizeMoneyValue(value) {
  const parsed = parseMoneyValue(value);
  if (parsed === null) {
    return null;
  }

  if (parsed < 0 || parsed > 1000000) {
    return null;
  }

  return Number(parsed.toFixed(2));
}

function sanitizeLaneHistoryEntry(entry, userKey, environment) {
  const safeEntry = entry || {};
  const shipperRate = sanitizeMoneyValue(safeEntry.shipperRate);
  const quotedRate = sanitizeMoneyValue(safeEntry.quotedRate);
  const negotiatedRate = sanitizeMoneyValue(safeEntry.negotiatedRate);
  const finalRate =
    sanitizeMoneyValue(safeEntry.finalRate) ??
    negotiatedRate ??
    quotedRate;

  if (finalRate === null) {
    const rateError = new Error("A valid final rate is required.");
    rateError.statusCode = 400;
    throw rateError;
  }

  const sourceConversationId = sanitizeIdentifier(
    safeEntry.sourceConversationId || safeEntry.sourceMessageId || `${Date.now()}`,
    180
  );
  const sourceMessageId = sanitizeIdentifier(
    safeEntry.sourceMessageId || sourceConversationId,
    180
  );

  return {
    id: sanitizeIdentifier(safeEntry.id || `${environment}::${userKey}::${sourceConversationId}`, 220),
    carrierName: sanitizeString(safeEntry.carrierName || "Unknown carrier", 160, "Unknown carrier"),
    carrierEmail: sanitizeEmail(safeEntry.carrierEmail),
    carrierPhone: sanitizePhone(safeEntry.carrierPhone),
    mcNumber: sanitizeMcNumber(safeEntry.mcNumber),
    pickupZip: sanitizeZip(safeEntry.pickupZip),
    deliveryZip: sanitizeZip(safeEntry.deliveryZip),
    shipperRate,
    quotedRate,
    negotiatedRate,
    finalRate,
    savedAt: sanitizeIdentifier(safeEntry.savedAt || new Date().toISOString(), 64),
    sourceMessageId,
    sourceConversationId,
  };
}

function validateParseLanePayload(payload) {
  const messages = Array.isArray(payload && payload.messages) ? payload.messages : [];
  if (messages.length > MAX_MESSAGES_PER_PARSE) {
    const error = new Error(`Too many messages requested. Limit is ${MAX_MESSAGES_PER_PARSE}.`);
    error.statusCode = 413;
    throw error;
  }

  const lane = (payload && payload.lane) || {};
  return {
    messages,
    lane: {
      pickupZip: sanitizeZip(lane.pickupZip),
      deliveryZip: sanitizeZip(lane.deliveryZip),
      shipperRate: sanitizeMoneyValue(lane.shipperRate),
    },
  };
}

module.exports = {
  MAX_JSON_BODY_BYTES,
  assertBodyWithinLimit,
  buildSecurityHeaders,
  checkRateLimit,
  getClientAddress,
  assertAllowedOrigin,
  toPublicErrorMessage,
  parseMoneyValue,
  sanitizeMoneyValue,
  sanitizeLaneHistoryEntry,
  validateParseLanePayload,
};
