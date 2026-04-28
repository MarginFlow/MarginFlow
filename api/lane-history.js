"use strict";

const {
  getMissingHistoryStoreConfigKeys,
  hasHistoryStoreConfig,
  listLaneHistory,
  saveLaneHistory,
  deleteLaneHistory,
} = require("../lib/history-store");

function sendConfigError(res) {
  const missing = getMissingHistoryStoreConfigKeys();
  res.status(503).json({
    error: missing.length
      ? `Lane history storage is not configured. Missing: ${missing.join(", ")}`
      : "Lane history storage is not configured.",
  });
}

function parseMoneyValue(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const parsed = Number(String(value).replace(/,/g, "").replace(/\$/g, "").trim());
  return Number.isFinite(parsed) ? parsed : null;
}

module.exports = async (req, res) => {
  if (!hasHistoryStoreConfig()) {
    sendConfigError(res);
    return;
  }

  try {
    if (req.method === "GET") {
      const userKey = String((req.query && req.query.userKey) || "").trim();
      const environment = String((req.query && req.query.environment) || "production").trim() || "production";
      if (!userKey) {
        res.status(400).json({ error: "userKey is required" });
        return;
      }

      const items = await listLaneHistory(userKey, environment);
      res.status(200).json({ items });
      return;
    }

    const payload = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
    const userKey = String(payload.userKey || payload.userEmail || "local-margin-flow-user").trim();
    const environment = String(payload.environment || "production").trim() || "production";

    if (!userKey) {
      res.status(400).json({ error: "userKey is required" });
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
    res.status(500).json({ error: error.message || "Lane history request failed" });
  }
};
