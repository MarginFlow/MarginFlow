"use strict";

const { parseCarrierEmail, dedupeVisibleMessages } = require("../lib/lane-parser");

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    const payload = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
    const messages = Array.isArray(payload.messages) ? payload.messages : [];
    const lane = payload.lane || {};

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
    res.status(500).json({ error: error.message || "Parse failed" });
  }
};
