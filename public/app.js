const state = {
  msalInstance: null,
  msalReady: false,
  authBusy: false,
  autoRefreshBusy: false,
  account: null,
  messages: [],
  loadedMessageCount: 0,
  results: [],
  hidden: [],
  selectedIds: new Set(),
  conversationThreads: {},
  lastSyncedAt: null,
  autoRefreshTimer: null,
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
  const connectionBadge = getElement("connection-badge");
  if (connectionBadge) {
    connectionBadge.textContent = signedIn ? "Microsoft inbox connected" : "Microsoft inbox not connected";
    connectionBadge.className = signedIn ? "trust-pill connected" : "trust-pill";
  }
}

function startAutoRefresh() {
  stopAutoRefresh();

  if (!state.account) {
    return;
  }

  state.autoRefreshTimer = window.setInterval(() => {
    if (document.hidden || state.autoRefreshBusy || !state.account) {
      return;
    }

    autoRefreshInbox().catch(() => {
      // Keep background refresh silent so the page never feels interrupted.
    });
  }, 5000);
}

function stopAutoRefresh() {
  if (state.autoRefreshTimer) {
    window.clearInterval(state.autoRefreshTimer);
    state.autoRefreshTimer = null;
  }
}

function updateSyncText() {
  const syncNode = getElement("sync-text");
  if (!syncNode) {
    return;
  }

  if (!state.lastSyncedAt) {
    syncNode.textContent = "Last synced: not yet loaded";
    return;
  }

  syncNode.textContent = `Last synced: ${state.lastSyncedAt.toLocaleString()}`;
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

function normalizeThreadText(value) {
  return String(value || "")
    .replace(/\r\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function stripQuotedReplyText(value) {
  const normalized = normalizeThreadText(value);
  if (!normalized) {
    return "";
  }

  const cutoffPatterns = [
    /\n_{5,}[\s\S]*$/i,
    /\n-+\s*Original Message\s*-+[\s\S]*$/i,
    /\nFrom:\s.*\nSent:\s.*\nTo:\s.*\nSubject:\s.*$/is,
    /\nOn .*wrote:[\s\S]*$/i,
  ];

  let cleaned = normalized;
  for (const pattern of cutoffPatterns) {
    cleaned = cleaned.replace(pattern, "");
  }

  return cleaned.trim();
}

function truncateText(value, limit = 260) {
  const normalized = stripQuotedReplyText(value);
  if (!normalized) {
    return "No message body available.";
  }
  return normalized.length > limit ? `${normalized.slice(0, limit).trim()}...` : normalized;
}

function sleep(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function isMailboxThrottleError(error) {
  const message = String((error && error.message) || "");
  return message.includes("ApplicationThrottled") || message.includes("MailboxConcurrency");
}

function extractRateCandidates(text) {
  const candidates = [];
  const source = String(text || "");
  const patterns = [
    /(?:rate(?: is)?|my rate(?: is)?|can do|can|cover|all in|offer|quote|for|flat|do)\s*[:=-]?\s*\$?\s*((?:\d{1,3}(?:,\d{3})+)|(?:\d{2,6}))(?:\.(\d{1,2}))?\b/gi,
    /\$\s*((?:\d{1,3}(?:,\d{3})+)|(?:\d{2,6}))(?:\.(\d{1,2}))?\b/gi,
    /\b((?:\d{1,3}(?:,\d{3})+)|(?:\d{2,6}))(?:\.(\d{1,2}))?\b/gi,
  ];

  patterns.forEach((regex, patternIndex) => {
    let match;
    while ((match = regex.exec(source)) !== null) {
      const wholePart = match[1] || "";
      const decimalPart = match[2] ? `.${match[2]}` : "";
      const parsed = Number(`${wholePart}${decimalPart}`.replace(/,/g, "").replace(/\$/g, "").trim());
      const matchedText = match[0] || "";
      const matchStart = match.index;
      const matchEnd = matchStart + matchedText.length;
      const prevChar = matchStart > 0 ? source[matchStart - 1] : "";
      const nextChar = matchEnd < source.length ? source[matchEnd] : "";

      if (!Number.isFinite(parsed) || parsed <= 1) {
        continue;
      }
      if (prevChar === "," || nextChar === ",") {
        continue;
      }

      const context = source
        .slice(Math.max(0, matchStart - 36), Math.min(source.length, regex.lastIndex + 36))
        .toLowerCase();

      let score = 4 - patternIndex;
      if (/\brate\b|\bcan\b|\bdo\b|\bcover\b|\ball in\b|\bquote\b|\boffer\b/.test(context)) {
        score += 4;
      }
      if (/\bmc\b|\bdot\b|\bphone\b|\bref\b|\bzip\b/.test(context)) {
        score -= 6;
      }
      if (parsed >= 100 && parsed <= 4000) {
        score += 2;
      }
      if (parsed > 10000) {
        score -= 10;
      }

      candidates.push({
        value: parsed,
        score,
        index: matchStart,
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

function computeMarginPercent(rate) {
  const shipperRate = normalizeMoney(getElement("shipper-rate").value);
  if (!Number.isFinite(shipperRate) || shipperRate <= 0 || !Number.isFinite(rate)) {
    return null;
  }
  return Math.round((((shipperRate - rate) / shipperRate) * 100) * 100) / 100;
}

function buildResultsSignature(results, hidden) {
  return JSON.stringify({
    results: (results || []).map((item) => ({
      id: item.id,
      conversationId: item.conversationId || "",
      carrierRate: item.carrierRate,
      marginPercent: item.marginPercent,
      mcNumber: item.mcNumber || "",
    })),
    hidden: (hidden || []).map((item) => item.id),
  });
}

function getNegotiatedDetails(item) {
  const thread = state.conversationThreads[item.conversationId] || [];
  if (!thread.length) {
    return null;
  }

  const originalRate = Number.isFinite(item.carrierRate) ? item.carrierRate : null;
  const carrierMessages = thread.filter((message) => classifyThreadAuthor(message) === "carrier");

  for (let index = carrierMessages.length - 1; index >= 0; index -= 1) {
    const message = carrierMessages[index];
    const candidates = extractRateCandidates(message.bodyText || message.subject || "");
    const candidate = candidates[0];
    if (!candidate || !Number.isFinite(candidate.value)) {
      continue;
    }
    if (originalRate !== null && Math.abs(candidate.value - originalRate) < 0.01) {
      continue;
    }

    const negotiatedRate = Math.round(candidate.value * 100) / 100;
    const marginPercent = computeMarginPercent(negotiatedRate);
    return {
      negotiatedRate,
      marginPercent,
      classification: classificationClass(
        marginPercent === null
          ? "RED"
          : marginPercent >= 15
            ? "GREEN"
            : marginPercent >= 5
              ? "YELLOW"
              : "RED"
      ),
    };
  }

  return null;
}

function sortResults(results) {
  const sortValue = getElement("sort-by") ? getElement("sort-by").value : "rate-asc";
  const sorted = [...results];

  sorted.sort((left, right) => {
    const leftRate = Number.isFinite(left.carrierRate) ? left.carrierRate : null;
    const rightRate = Number.isFinite(right.carrierRate) ? right.carrierRate : null;
    const leftNegotiated = getNegotiatedDetails(left);
    const rightNegotiated = getNegotiatedDetails(right);
    const leftNegotiatedRate = leftNegotiated && Number.isFinite(leftNegotiated.negotiatedRate) ? leftNegotiated.negotiatedRate : null;
    const rightNegotiatedRate = rightNegotiated && Number.isFinite(rightNegotiated.negotiatedRate) ? rightNegotiated.negotiatedRate : null;

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

    if (sortValue === "negotiated-asc") {
      if (leftNegotiatedRate === null && rightNegotiatedRate === null) {
        return 0;
      }
      if (leftNegotiatedRate === null) {
        return 1;
      }
      if (rightNegotiatedRate === null) {
        return -1;
      }
      return leftNegotiatedRate - rightNegotiatedRate;
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

  stopAutoRefresh();

  await state.msalInstance.logoutRedirect({
    account: state.account,
    postLogoutRedirectUri: getConfig().redirectUri,
  });
}

async function fetchInboxMessages(options = {}) {
  const { silent = false } = options;
  const limit = Number(getElement("scan-limit").value || 50);
  if (!silent) {
    setStatus(`Loading the last ${limit} inbox emails from Microsoft Graph...`);
  }

  const result = await graphFetch(
    `/me/mailFolders/inbox/messages?$top=${limit}&$select=id,subject,from,receivedDateTime,bodyPreview,body,webLink,conversationId`
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
    conversationId: message.conversationId || "",
  }));
  state.loadedMessageCount = limit;
  state.lastSyncedAt = new Date();
  updateSyncText();

  if (!silent) {
    setStatus(`Loaded ${state.messages.length} inbox emails. Scan the active lane when you're ready.`);
  }
}

function captureUiSnapshot() {
  const replyDrafts = {};
  document.querySelectorAll(".reply-input").forEach((node) => {
    if (node.id) {
      replyDrafts[node.id] = node.value;
    }
  });

  const activeElement = document.activeElement;
  return {
    scrollX: window.scrollX,
    scrollY: window.scrollY,
    bulkReply: getElement("bulk-reply-input") ? getElement("bulk-reply-input").value : "",
    replyDrafts,
    activeId: activeElement && activeElement.id ? activeElement.id : null,
    selectionStart:
      activeElement && typeof activeElement.selectionStart === "number" ? activeElement.selectionStart : null,
    selectionEnd:
      activeElement && typeof activeElement.selectionEnd === "number" ? activeElement.selectionEnd : null,
  };
}

function restoreUiSnapshot(snapshot) {
  if (!snapshot) {
    return;
  }

  if (getElement("bulk-reply-input")) {
    getElement("bulk-reply-input").value = snapshot.bulkReply || "";
  }

  Object.entries(snapshot.replyDrafts || {}).forEach(([id, value]) => {
    const node = getElement(id);
    if (node) {
      node.value = value;
    }
  });

  if (snapshot.activeId) {
    const activeNode = getElement(snapshot.activeId);
    if (activeNode) {
      activeNode.focus({ preventScroll: true });
      if (
        typeof snapshot.selectionStart === "number" &&
        typeof snapshot.selectionEnd === "number" &&
        typeof activeNode.setSelectionRange === "function"
      ) {
        activeNode.setSelectionRange(snapshot.selectionStart, snapshot.selectionEnd);
      }
    }
  }

  window.scrollTo(snapshot.scrollX, snapshot.scrollY);
}

async function parseCurrentLane(options = {}) {
  const {
    silent = false,
    preserveUi = false,
    preserveSelected = false,
  } = options;
  const lane = getLaneInput();

  if (!lane.pickupZip && !lane.deliveryZip) {
    if (!silent) {
      setStatus("Enter at least one lane ZIP before scanning.");
    }
    return false;
  }

  if (!silent) {
    setStatus("Parsing inbox emails for the active lane...");
  }

  const previousSignature = buildResultsSignature(state.results, state.hidden);
  const uiSnapshot = preserveUi ? captureUiSnapshot() : null;
  const selectedBefore = preserveSelected ? new Set(state.selectedIds) : null;
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
  const nextResults = payload.results || [];
  const nextHidden = payload.hidden || [];
  const nextSignature = buildResultsSignature(nextResults, nextHidden);
  const changed = previousSignature !== nextSignature;

  state.results = nextResults;
  state.hidden = nextHidden;
  if (changed) {
    state.conversationThreads = {};
  }
  if (selectedBefore) {
    state.selectedIds = new Set(
      Array.from(selectedBefore).filter((id) => state.results.some((item) => item.id === id))
    );
  } else {
    state.selectedIds.clear();
  }
  updateSummary();
  if (!changed && silent) {
    restoreUiSnapshot(uiSnapshot);
    return false;
  }

  renderResults();
  restoreUiSnapshot(uiSnapshot);

  try {
    await loadConversationThreads();
    restoreUiSnapshot(uiSnapshot);
  } catch (error) {
    if (!silent) {
      setStatus(`Loaded lane matches, but conversation threads failed: ${error.message}`);
    }
    return true;
  }

  if (!silent) {
    setStatus(`Found ${state.results.length} relevant emails for this lane and hid ${state.hidden.length}.`);
  }
  return changed;
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

  const requestedLimit = Number(getElement("scan-limit").value || 50);
  if (state.loadedMessageCount !== requestedLimit) {
    await fetchInboxMessages();
  }

  await parseCurrentLane();
}

async function autoRefreshInbox() {
  if (!state.account) {
    return;
  }

  state.autoRefreshBusy = true;
  try {
    await fetchInboxMessages({ silent: true });

    if (state.results.length && hasLaneInputs()) {
      const changed = await parseCurrentLane({
        silent: true,
        preserveUi: true,
        preserveSelected: true,
      });

      if (changed) {
        setStatus(`Auto-refreshed ${state.messages.length} inbox emails.`);
      }
    }
  } finally {
    state.autoRefreshBusy = false;
  }
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

function findResultById(messageId) {
  return state.results.find((item) => item.id === messageId) || null;
}

function classifyThreadAuthor(message) {
  const accountEmail = String(state.account && state.account.username ? state.account.username : "").toLowerCase();
  const fromAddress = String(message.fromAddress || "").toLowerCase();
  return accountEmail && fromAddress === accountEmail ? "broker" : "carrier";
}

async function fetchConversationThread(conversationId) {
  if (!conversationId) {
    return [];
  }

  const safeConversationId = String(conversationId).replace(/'/g, "''");
  const filter = encodeURIComponent(`conversationId eq '${safeConversationId}'`);
  const result = await graphFetch(
    `/me/messages?$filter=${filter}&$top=25&$select=id,subject,from,receivedDateTime,bodyPreview,body,webLink,conversationId`
  );

  return (result && result.value ? result.value : [])
    .map((message) => ({
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
      conversationId: message.conversationId || "",
    }))
    .sort((left, right) => {
      const leftTime = left.receivedDateTime ? Date.parse(left.receivedDateTime) : 0;
      const rightTime = right.receivedDateTime ? Date.parse(right.receivedDateTime) : 0;
      return leftTime - rightTime;
    });
}

async function refreshThreadForMessage(messageId) {
  const result = findResultById(messageId);
  if (!result || !result.conversationId) {
    return;
  }

  try {
    const thread = await fetchConversationThread(result.conversationId);
    state.conversationThreads[result.conversationId] = thread;
    updateConversationUiForItem(result);
  } catch (error) {
    if (isMailboxThrottleError(error)) {
      setStatus("Reply sent, but Outlook is rate-limiting thread refreshes right now. Try Refresh Inbox in a moment.");
      return;
    }
    throw error;
  }
}

async function loadConversationThreads() {
  const conversationIds = Array.from(
    new Set(
      state.results
        .map((item) => item.conversationId)
        .filter(Boolean)
    )
  );

  if (!conversationIds.length) {
    return;
  }

  const nextThreads = {};
  for (const conversationId of conversationIds) {
      try {
        nextThreads[conversationId] = await fetchConversationThread(conversationId);
        await sleep(150);
      } catch (error) {
        if (isMailboxThrottleError(error)) {
          state.conversationThreads = {
            ...state.conversationThreads,
            ...nextThreads,
          };
          updateConversationUi();
          throw new Error("Outlook is rate-limiting conversation history. Try Refresh Inbox in a moment.");
        }
        throw error;
      }
  }

  state.conversationThreads = {
    ...state.conversationThreads,
    ...nextThreads,
  };
  updateConversationUi();
}

function renderThread(item) {
  return `<div id="thread-container-${escapeHtml(item.id)}">${renderThreadMarkup(item)}</div>`;
}

function renderThreadMarkup(item) {
  const thread = (state.conversationThreads[item.conversationId] || []).filter((message) => message.id !== item.id);
  if (!thread.length) {
    return '<div class="thread-empty">Conversation replies will appear here.</div>';
  }

  return `
    <div class="thread-panel">
      <p class="thread-title">Conversation</p>
      <div class="thread-list">
        ${thread
          .map((message) => {
            const authorType = classifyThreadAuthor(message);
            const authorLabel =
              authorType === "broker"
                ? "You"
                : escapeHtml(message.fromName || message.fromAddress || "Carrier");

            return `
              <article class="thread-item ${authorType}">
                <div class="thread-head">
                  <span class="thread-author">${authorLabel}</span>
                  <span class="thread-date">${escapeHtml(message.receivedDateTime || "")}</span>
                </div>
                <p class="thread-body">${escapeHtml(truncateText(message.bodyText || message.subject || ""))}</p>
              </article>
            `;
          })
          .join("")}
      </div>
    </div>
  `;
}

function renderNegotiatedSectionMarkup(item) {
  const negotiated = getNegotiatedDetails(item);
  if (!negotiated) {
    return "";
  }

  return `
      <div class="economics-block negotiated-block">
        <p class="rate-value ${negotiated.classification}">${escapeHtml(formatMoney(negotiated.negotiatedRate))}</p>
        <p class="rate-label">Negotiated Rate</p>
        <p class="margin-value ${negotiated.classification}">${
          negotiated.marginPercent === null ? "New Margin N/A" : `New Margin ${escapeHtml(negotiated.marginPercent)}%`
        }</p>
      </div>
    `;
}

function renderNegotiatedSection(item) {
  return `<div id="negotiated-container-${escapeHtml(item.id)}">${renderNegotiatedSectionMarkup(item)}</div>`;
}

function updateConversationUiForItem(item) {
  const economicsNode = getElement(`economics-grid-${item.id}`);
  const hasNegotiated = Boolean(getNegotiatedDetails(item));
  if (economicsNode) {
    economicsNode.className = hasNegotiated ? "economics-grid has-negotiated" : "economics-grid";
  }

  const threadNode = getElement(`thread-container-${item.id}`);
  if (threadNode) {
    threadNode.innerHTML = renderThreadMarkup(item);
  }

  const negotiatedNode = getElement(`negotiated-container-${item.id}`);
  if (negotiatedNode) {
    negotiatedNode.innerHTML = renderNegotiatedSectionMarkup(item);
  }
}

function updateConversationUi() {
  state.results.forEach((item) => {
    updateConversationUiForItem(item);
  });
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
    await refreshThreadForMessage(messageId);
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
      await refreshThreadForMessage(messageId);
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
              <div id="economics-grid-${escapeHtml(item.id)}" class="economics-grid ${getNegotiatedDetails(item) ? "has-negotiated" : ""}">
                <div class="economics-block">
                  <p class="rate-value ${classificationClass(item.classification)}">${escapeHtml(formatMoney(item.carrierRate))}</p>
                  <p class="rate-label">Quoted Rate</p>
                  <p class="margin-value ${classificationClass(item.classification)}">${
                    item.marginPercent === null ? "Margin N/A" : `Margin ${escapeHtml(item.marginPercent)}%`
                  }</p>
                </div>
                ${renderNegotiatedSection(item)}
              </div>
            </div>
          <div class="pill-row">
            <span class="pill">MC ${escapeHtml(item.mcNumber || "Unknown")}</span>
          </div>
          <p class="preview">${escapeHtml(item.bodyPreview || "No preview available.")}</p>
          <div class="pill-row">
            <a class="pill mono" href="${escapeHtml(item.webLink)}" target="_blank" rel="noreferrer">Open email</a>
            <span class="pill mono">${escapeHtml(item.receivedDateTime || "")}</span>
          </div>
          ${renderThread(item)}
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
    startAutoRefresh();
  }

  updateActionState();
  updateSyncText();
  updateSummary();
  renderResults();
  wireUi();
}

document.addEventListener("DOMContentLoaded", () => {
  initializeApp().catch((error) => {
    setStatus(`Initialization failed: ${error.message}`);
  });
});
