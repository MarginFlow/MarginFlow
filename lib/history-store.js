"use strict";

const { createClient } = require("@libsql/client");

let client;
let schemaReadyPromise;

function getConfig() {
  return {
    url: process.env.TURSO_DATABASE_URL || "",
    authToken: process.env.TURSO_AUTH_TOKEN || "",
  };
}

function getMissingHistoryStoreConfigKeys() {
  const config = getConfig();
  const missing = [];

  if (!config.url) {
    missing.push("TURSO_DATABASE_URL");
  }

  if (!config.authToken) {
    missing.push("TURSO_AUTH_TOKEN");
  }

  return missing;
}

function hasHistoryStoreConfig() {
  return getMissingHistoryStoreConfigKeys().length === 0;
}

function getClient() {
  if (client) {
    return client;
  }

  const config = getConfig();
  if (!config.url || !config.authToken) {
    throw new Error("Turso is not configured. Set TURSO_DATABASE_URL and TURSO_AUTH_TOKEN.");
  }

  client = createClient({
    url: config.url,
    authToken: config.authToken,
  });

  return client;
}

async function ensureSchema() {
  if (!schemaReadyPromise) {
    const db = getClient();
    schemaReadyPromise = (async () => {
      await db.execute(`CREATE TABLE IF NOT EXISTS lane_history (
        id TEXT PRIMARY KEY,
        user_key TEXT NOT NULL,
        user_email TEXT NOT NULL,
        environment TEXT NOT NULL DEFAULT 'production',
        carrier_name TEXT NOT NULL,
        carrier_email TEXT,
        carrier_phone TEXT,
        mc_number TEXT,
        pickup_zip TEXT,
        delivery_zip TEXT,
        shipper_rate REAL,
        quoted_rate REAL,
        negotiated_rate REAL,
        final_rate REAL NOT NULL,
        saved_at TEXT NOT NULL,
        source_message_id TEXT,
        source_conversation_id TEXT
      )`);
      try {
        await db.execute("ALTER TABLE lane_history ADD COLUMN shipper_rate REAL");
      } catch (error) {
        const message = String(error && error.message ? error.message : "");
        if (!/duplicate column name|already exists/i.test(message)) {
          throw error;
        }
      }
      try {
        await db.execute("ALTER TABLE lane_history ADD COLUMN environment TEXT NOT NULL DEFAULT 'production'");
      } catch (error) {
        const message = String(error && error.message ? error.message : "");
        if (!/duplicate column name|already exists/i.test(message)) {
          throw error;
        }
      }
      await db.execute("CREATE INDEX IF NOT EXISTS idx_lane_history_user_saved ON lane_history(user_key, saved_at DESC)");
      await db.execute("CREATE INDEX IF NOT EXISTS idx_lane_history_user_conversation ON lane_history(user_key, source_conversation_id)");
      await db.execute("CREATE INDEX IF NOT EXISTS idx_lane_history_user_env_saved ON lane_history(user_key, environment, saved_at DESC)");
    })();
  }

  await schemaReadyPromise;
}

function toRowObject(row) {
  return {
    id: row.id,
    userKey: row.user_key,
    userEmail: row.user_email,
    environment: row.environment,
    carrierName: row.carrier_name,
    carrierEmail: row.carrier_email,
    carrierPhone: row.carrier_phone,
    mcNumber: row.mc_number,
    pickupZip: row.pickup_zip,
    deliveryZip: row.delivery_zip,
    shipperRate: row.shipper_rate,
    quotedRate: row.quoted_rate,
    negotiatedRate: row.negotiated_rate,
    finalRate: row.final_rate,
    savedAt: row.saved_at,
    sourceMessageId: row.source_message_id,
    sourceConversationId: row.source_conversation_id,
  };
}

async function listLaneHistory(userKey, environment) {
  await ensureSchema();
  const db = getClient();
  const result = await db.execute({
    sql: `SELECT
      id,
      user_key,
      user_email,
      environment,
      carrier_name,
      carrier_email,
      carrier_phone,
      mc_number,
      pickup_zip,
      delivery_zip,
      shipper_rate,
      quoted_rate,
      negotiated_rate,
      final_rate,
      saved_at,
      source_message_id,
      source_conversation_id
    FROM lane_history
    WHERE user_key = ? AND environment = ?
    ORDER BY datetime(saved_at) DESC`,
    args: [userKey, environment],
  });

  return (result.rows || []).map(toRowObject);
}

async function saveLaneHistory(entry) {
  await ensureSchema();
  const db = getClient();

  await db.execute({
    sql: `INSERT OR REPLACE INTO lane_history (
      id,
      user_key,
      user_email,
      environment,
      carrier_name,
      carrier_email,
      carrier_phone,
      mc_number,
      pickup_zip,
      delivery_zip,
      shipper_rate,
      quoted_rate,
      negotiated_rate,
      final_rate,
      saved_at,
      source_message_id,
      source_conversation_id
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)` ,
    args: [
      entry.id,
      entry.userKey,
      entry.userEmail,
      entry.environment,
      entry.carrierName,
      entry.carrierEmail,
      entry.carrierPhone,
      entry.mcNumber,
      entry.pickupZip,
      entry.deliveryZip,
      entry.shipperRate,
      entry.quotedRate,
      entry.negotiatedRate,
      entry.finalRate,
      entry.savedAt,
      entry.sourceMessageId,
      entry.sourceConversationId,
    ],
  });
}

async function deleteLaneHistory(userKey, environment, sourceConversationId) {
  await ensureSchema();
  const db = getClient();
  await db.execute({
    sql: "DELETE FROM lane_history WHERE user_key = ? AND environment = ? AND source_conversation_id = ?",
    args: [userKey, environment, sourceConversationId],
  });
}

module.exports = {
  getMissingHistoryStoreConfigKeys,
  hasHistoryStoreConfig,
  listLaneHistory,
  saveLaneHistory,
  deleteLaneHistory,
};
