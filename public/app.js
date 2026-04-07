const state = {
  msalInstance: null,
  msalReady: false,
  authBusy: false,
  account: null,
  messages: [],
  results: [],
  hidden: [],
  selectedIds: new Set(),
};

function getElement(id) {
  return document.getElementById(id);
}

function setStatus(message) {
  getElement("status-text").textContent = message;
}

function normalizeMoney(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function formatMoney(value) {
  if (!Number.isFinite(value)) {
    return "-";
  }

  return `$${value.toFixed(2)}`;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getConfig() {
  return window.FIREFLY_CONFIG || {};
}

function hasUsableConfig() {
  const config = getConfig();
  return Boolean(config.clientId && config.clientId !== "REPLACE_ME");
}

function hasLaneInputs() {
  const lane = getLaneInput();
  return Boolean(lane.pickupZip || lane.deliveryZip);
}

function updateActionState() {
  const signedIn = Boolean(state.account);
  getElement("refresh-button").disabled = !signedIn;
  getElement("scan-lane-button").disabled = !(signedIn && hasLaneInputs());
  getElement("sign-out-button").disabled = !signedIn;
}

function updateSummary() {
  getElement("visible-count").textContent = String(state.results.length);
  getElement("hidden-count").textContent = String(state.hidden.length);
  getElement("selected-count").textContent = String(state.selectedIds.size);

  const lowestRate = state.results
    .map((item) => item.carrierRate)
    .filter((value) => Number.isFinite(value))
    .sort((left, right) => left - right)[0];
  getElement("lowest-rate").textContent = Number.isFinite(lowestRate) ? formatMoney(lowestRate) : "-";
}

function sortResults(results) {
  const sortValue = getElement("sort-by") ? getElement("sort-by").value : "rate-asc";
  const sorted = [...results];

  sorted.sort((left, right) => {
    const leftRate = Number.isFinite(left.carrierRate) ? left.carrierRate : null;
    const rightRate = Number.isFinite(right.carrierRate) ? right.carrierRate : null;

    if (sortValue === "rate-desc") {
      if (leftRate === null && rightRate === null) {
        return 0;
      }
      if (leftRate === null) {
        return 1;
      }
      if (rightRate === null) {
        return -1;
      }
      return rightRate - leftRate;
    }

    if (leftRate === null && rightRate === null) {
      return 0;
    }
    if (leftRate === null) {
      return 1;
    }
    if (rightRate === null) {
      return -1;
    }
    return leftRate - rightRate;
  });

  return sorted;
}

function getLaneInput() {
  return {
    pickupZip: getElement("pickup-zip").value.trim(),
    deliveryZip: getElement("delivery-zip").value.trim(),
    shipperRate: normalizeMoney(getElement("shipper-rate").value),
  };
}

async function getAccessToken() {
  const scopes = getConfig().graphScopes || ["User.Read", "Mail.Read", "Mail.Send"];
  let tokenResponse;

  try {
    tokenResponse = await state.msalInstance.acquireTokenSilent({
      account: state.account,
      scopes,
    });
  } catch (error) {
    const message = String(error.message || "");
    if (message.includes("interaction_required") || message.includes("consent_required")) {
      tokenResponse = await state.msalInstance.acquireTokenRedirect({
        account: state.account,
        scopes,
      });
    } else {
      throw error;
    }
  }

  return tokenResponse.accessToken;
}

async function graphFetch(path, options = {}) {
  const token = await getAccessToken();
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Prefer: 'outlook.body-content-type="text"',
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || `Graph request failed with ${response.status}`);
  }

  if (response.status === 202 || response.status === 204) {
    return null;
  }

  const text = await response.text();
  if (!text) {
    return null;
  }

  return JSON.parse(text);
}

async function signIn() {
  if (!state.msalReady) {
    setStatus("Microsoft sign-in is still initializing. Try again in a moment.");
    return;
  }

  if (state.authBusy) {
    setStatus("A Microsoft sign-in window is already in progress. Finish it or close it, then try again.");
    return;
  }

  if (!hasUsableConfig()) {
    setStatus("Add your Microsoft app registration values to public/config.js before signing in.");
    return;
  }

  try {
    state.authBusy = true;
    setStatus("Redirecting to Microsoft sign-in...");
    await state.msalInstance.loginRedirect({
      scopes: getConfig().graphScopes,
      prompt: "select_account",
    });
    return;
  } catch (error) {
    const message = String(error.message || "");
    if (message.includes("interaction_in_progress")) {
      setStatus("Microsoft sign-in is already in progress. Close any open login popup and try once more.");
    } else {
      setStatus(`Sign-in failed: ${error.message}`);
    }
  } finally {
    state.authBusy = false;
  }
}

async function signOut() {
  if (!state.msalReady) {
    setStatus("Microsoft sign-in is still initializing.");
    return;
  }

  if (!state.account) {
    return;
  }

  await state.msalInstance.logoutRedirect({
    account: state.account,
    postLogoutRedirectUri: getConfig().redirectUri,
  });
}

async function fetchInboxMessages() {
  const limit = Number(getElement("scan-limit").value || 50);
  setStatus(`Loading the last ${limit} inbox emails from Microsoft Graph...`);

  const result = await graphFetch(
    `/me/mailFolders/inbox/messages?$top=${limit}&$select=id,subject,from,receivedDateTime,bodyPreview,body,webLink`
  );

  state.messages = (result.value || []).map((message) => ({
    id: message.id,
    subject: message.subject || "",
    fromName: message.from && message.from.emailAddress ? message.from.emailAddress.name || "" : "",
    fromAddress: message.from && message.from.emailAddress ? message.from.emailAddress.address || "" : "",
    receivedDateTime: message.receivedDateTime || "",
    bodyText:
      message.body && typeof message.body.content === "string" && message.body.content.trim()
        ? message.body.content
        : message.bodyPreview || "",
    webLink: message.webLink || "",
  }));

  setStatus(`Loaded ${state.messages.length} inbox emails. Scan the active lane when you're ready.`);
}

async function scanLane() {
  const lane = getLaneInput();
  if (!lane.pickupZip && !lane.deliveryZip) {
    setStatus("Enter at least one lane ZIP before scanning.");
    return;
  }

  if (!state.messages.length) {
    await fetchInboxMessages();
  }

  setStatus("Parsing inbox emails for the active lane...");

  const response = await fetch("/api/parse-lane", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      lane,
      messages: state.messages,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || "Lane parse failed");
  }

  const payload = await response.json();
  state.results = payload.results || [];
  state.hidden = payload.hidden || [];
  state.selectedIds.clear();
  renderResults();
  updateSummary();
  setStatus(`Found ${state.results.length} relevant emails for this lane and hid ${state.hidden.length}.`);
}

function classificationClass(value) {
  if (value === "GREEN") {
    return "green";
  }
  if (value === "YELLOW") {
    return "yellow";
  }
  return "red";
}

function buildReplyComment(rawValue) {
  const trimmed = String(rawValue || "").trim();
  if (!trimmed) {
    return "";
  }

  if (/^\$?\d+(?:\.\d{1,2})?$/.test(trimmed)) {
    const amount = trimmed.startsWith("$") ? trimmed : `$${trimmed}`;
    return `Can you do ${amount} all in?`;
  }

  return trimmed;
}

async function sendReply(messageId) {
  const textarea = getElement(`reply-${messageId}`);
  const statusNode = getElement(`reply-status-${messageId}`);
  const comment = buildReplyComment(textarea.value);
  if (!comment) {
    setStatus("Type a reply price or reply message before sending.");
    if (statusNode) {
      statusNode.textContent = "Type a reply before sending.";
    }
    return;
  }

  try {
    setStatus("Sending reply through Microsoft Graph...");
    if (statusNode) {
      statusNode.textContent = "Sending reply...";
    }
    await graphFetch(`/me/messages/${messageId}/reply`, {
      method: "POST",
      body: JSON.stringify({ comment }),
    });
    setStatus("Reply sent.");
    textarea.value = "";
    if (statusNode) {
      statusNode.textContent = "Reply sent.";
    }
  } catch (error) {
    setStatus(`Reply failed: ${error.message}`);
    if (statusNode) {
      statusNode.textContent = `Reply failed: ${error.message}`;
    }
  }
}

function toggleSelection(messageId, checked) {
  if (checked) {
    state.selectedIds.add(messageId);
  } else {
    state.selectedIds.delete(messageId);
  }
  updateSummary();
}

async function massReplyToSelected() {
  const rawValue = getElement("bulk-reply-input").value;
  const prepared = buildReplyComment(rawValue);

  if (!prepared) {
    setStatus("Type a price or message in the mass reply box first.");
    return;
  }

  if (!state.selectedIds.size) {
    setStatus("Select at least one carrier before sending a mass reply.");
    return;
  }

  const selectedIds = Array.from(state.selectedIds);
  let sentCount = 0;

  setStatus(`Sending replies to ${selectedIds.length} selected carriers...`);

  for (const messageId of selectedIds) {
    const textarea = getElement(`reply-${messageId}`);
    const statusNode = getElement(`reply-status-${messageId}`);

    if (textarea) {
      textarea.value = rawValue;
    }

    try {
      if (statusNode) {
        statusNode.textContent = "Sending mass reply...";
      }

      await graphFetch(`/me/messages/${messageId}/reply`, {
        method: "POST",
        body: JSON.stringify({ comment: prepared }),
      });

      sentCount += 1;
      if (statusNode) {
        statusNode.textContent = "Mass reply sent.";
      }
      if (textarea) {
        textarea.value = "";
      }
    } catch (error) {
      if (statusNode) {
        statusNode.textContent = `Mass reply failed: ${error.message}`;
      }
    }
  }

  setStatus(`Mass reply finished. Sent ${sentCount} of ${selectedIds.length} selected replies individually.`);
}

function renderResults() {
  const container = getElement("results-list");
  const sortedResults = sortResults(state.results);

  if (!sortedResults.length) {
    container.innerHTML = '<p class="empty-state">No visible carrier emails for the active lane yet.</p>';
    updateSummary();
    return;
  }

  container.innerHTML = sortedResults
    .map(
      (item) => `
        <article class="result-card">
          <div class="result-top">
            <div>
              <div class="carrier-line">
                <input class="selection" type="checkbox" data-select-id="${escapeHtml(item.id)}" />
                <h3 class="carrier-name">${escapeHtml(item.carrierName)}</h3>
              </div>
              <p class="subject">${escapeHtml(item.subject || "(No subject)")}</p>
              <p class="meta-line">${escapeHtml(item.fromAddress || "")}</p>
            </div>
            <div>
              <p class="rate-value ${classificationClass(item.classification)}">${escapeHtml(formatMoney(item.carrierRate))}</p>
              <p class="rate-label">Quoted Rate</p>
              <p class="margin-value ${classificationClass(item.classification)}">${
                item.marginPercent === null ? "Margin N/A" : `Margin ${escapeHtml(item.marginPercent)}%`
              }</p>
            </div>
          </div>
          <div class="pill-row">
            <span class="pill">MC ${escapeHtml(item.mcNumber || "Unknown")}</span>
            <span class="pill">${escapeHtml(item.classification)}</span>
          </div>
          <p class="preview">${escapeHtml(item.bodyPreview || "No preview available.")}</p>
          <div class="pill-row">
            <a class="pill mono" href="${escapeHtml(item.webLink)}" target="_blank" rel="noreferrer">Open email</a>
            <span class="pill mono">${escapeHtml(item.receivedDateTime || "")}</span>
          </div>
          <div class="reply-box">
            <textarea id="reply-${escapeHtml(item.id)}" class="reply-input" placeholder="Type a new price or message. Example: 1800"></textarea>
            <button class="primary-button send-button" type="button" data-send-id="${escapeHtml(item.id)}">Send Reply</button>
          </div>
          <p id="reply-status-${escapeHtml(item.id)}" class="reply-status"></p>
        </article>
      `
    )
    .join("");

  container.querySelectorAll("[data-select-id]").forEach((node) => {
    node.addEventListener("change", (event) => {
      toggleSelection(node.getAttribute("data-select-id"), event.target.checked);
    });
  });

  container.querySelectorAll("[data-send-id]").forEach((node) => {
    node.addEventListener("click", () => {
      sendReply(node.getAttribute("data-send-id"));
    });
  });

  updateSummary();
}

function wireUi() {
  getElement("sign-in-button").addEventListener("click", signIn);
  getElement("sign-out-button").addEventListener("click", signOut);
  getElement("refresh-button").addEventListener("click", fetchInboxMessages);
  getElement("mass-reply-button").addEventListener("click", () => {
    massReplyToSelected().catch((error) => {
      setStatus(`Mass reply failed: ${error.message}`);
    });
  });
  getElement("sort-by").addEventListener("change", renderResults);
  getElement("scan-lane-button").addEventListener("click", async () => {
    try {
      await scanLane();
    } catch (error) {
      setStatus(`Lane scan failed: ${error.message}`);
    }
  });

  ["pickup-zip", "delivery-zip", "shipper-rate"].forEach((id) => {
    getElement(id).addEventListener("input", updateActionState);
  });
}

async function initializeApp() {
  if (typeof window.msal === "undefined") {
    setStatus("Microsoft sign-in library failed to load. Refresh the page and confirm the local server is running.");
    return;
  }

  if (!hasUsableConfig()) {
    setStatus("Add your Microsoft app registration values to public/config.js, then reload this page.");
  }

  const config = getConfig();
  state.msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: config.clientId || "REPLACE_ME",
      authority: `https://login.microsoftonline.com/${config.tenantId || "common"}`,
      redirectUri: config.redirectUri || window.location.origin,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true,
    },
  });

  await state.msalInstance.initialize();
  const redirectResult = await state.msalInstance.handleRedirectPromise().catch(() => null);
  state.msalReady = true;

  if (redirectResult && redirectResult.account) {
    state.account = redirectResult.account;
  }

  const accounts = state.msalInstance.getAllAccounts();
  if (!state.account && accounts.length > 0) {
    state.account = accounts[0];
    state.msalInstance.setActiveAccount(accounts[0]);
  }

  if (state.account) {
    setStatus(`Signed in as ${state.account.username}. Ready to scan the inbox.`);
  }

  updateActionState();
  updateSummary();
  renderResults();
  wireUi();
}

document.addEventListener("DOMContentLoaded", () => {
  initializeApp().catch((error) => {
    setStatus(`Initialization failed: ${error.message}`);
  });
});
