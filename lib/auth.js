"use strict";

async function readJsonSafely(response) {
  const text = await response.text();
  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    return { raw: text };
  }
}

function getBearerToken(headersLike) {
  const headerValue =
    (headersLike && typeof headersLike.get === "function" && headersLike.get("authorization")) ||
    (headersLike && (headersLike.authorization || headersLike.Authorization)) ||
    "";

  const match = String(headerValue).match(/^Bearer\s+(.+)$/i);
  return match ? match[1].trim() : "";
}

async function resolveMicrosoftUser(accessToken) {
  if (!accessToken) {
    throw new Error("Missing Microsoft bearer token.");
  }

  const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,mail", {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    const payload = await readJsonSafely(response);
    const message =
      payload &&
      payload.error &&
      (payload.error.message || payload.error.code) ||
      `Microsoft identity lookup failed with ${response.status}`;
    const error = new Error(message);
    error.statusCode = response.status === 401 || response.status === 403 ? 401 : 502;
    throw error;
  }

  const profile = await response.json();
  const userKey = String(profile.id || "").trim();
  const userEmail = String(profile.userPrincipalName || profile.mail || "").trim();

  if (!userKey) {
    const error = new Error("Microsoft identity lookup returned no stable user id.");
    error.statusCode = 502;
    throw error;
  }

  return {
    userKey,
    userEmail,
  };
}

module.exports = {
  getBearerToken,
  resolveMicrosoftUser,
};
