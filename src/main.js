// src/main.js — Teams Chat Export Portal
// Clean production-focused build: login + chats + queue export + offline HTML + forwarded resolution

import * as msal from "@azure/msal-browser";

/** =========================
 * CONFIG
 * ========================= */
const TENANT_ID = "2c51b27f-3900-46f3-b208-a8b66df531e3";
const CLIENT_ID = "77934d56-04dd-4faf-b26b-d97275751bf8";
const REDIRECT_URI = window.location.origin + "/";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/** =========================
 * SESSION POLICY
 * ========================= */
const SESSION_PREEXPORT_WINDOW_MS = 300000 * 60 * 1000;
const SESSION_CHECK_EVERY_MS = 5000;

const APP_SESSION_KEYS = {
  loginTs: "teams_export_login_ts",
  exportStarted: "teams_export_started",
};

/** =========================
 * RETRY POLICY
 * ========================= */
const RETRY = {
  GRAPH_MAX_ATTEMPTS: 5,
  BINARY_MAX_ATTEMPTS: 5,
  BASE_DELAY_MS: 900,
  MAX_DELAY_MS: 8000,
};

/** =========================
 * LIMITS
 * ========================= */
const LIMITS = {
  MAX_IMAGE_BYTES_EACH: 10 * 1024 * 1024,
  MAX_IMAGE_BYTES_TOTAL: 120 * 1024 * 1024,

  MAX_ATTACH_BYTES_EACH: 250 * 1024 * 1024,
  MAX_ATTACH_BYTES_TOTAL: 1200 * 1024 * 1024,

  MAX_PAGES: 200000,
  DL_CONCURRENCY: 6,

  DENY_EXT: new Set([
    // "exe", "msi"
  ]),
};

/** =========================
 * SCOPES (delegated)
 * ========================= */
const SCOPES = [
  "User.Read",
  "Chat.Read",
  "offline_access",
  "Files.Read.All",
  // "Sites.Read.All",
];

/** =========================
 * MSAL
 * ========================= */
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
});

console.log("[BOOT] main.js loaded", new Date().toISOString());

/** =========================
 * DOM
 * ========================= */
const $ = (id) => document.getElementById(id);

const ui = {
  loginView: null,
  appView: null,

  btnLogin: null,
  btnChats: null,
  btnExport: null,
  btnQueueAll: null,
  btnRetryFailed: null,

  mePill: null,
  countStat: null,
  selectedStat: null,

  q: null,
  chatList: null,
  queueStatus: null,
};

function refreshUiRefs() {
  ui.loginView = $("loginView");
  ui.appView = $("appView");

  ui.btnLogin = $("btnLogin");
  ui.btnChats = $("btnChats");
  ui.btnExport = $("btnExport");
  ui.btnQueueAll = $("btnQueueAll");
  ui.btnRetryFailed = $("btnRetryFailed");

  ui.mePill = $("mePill");
  ui.countStat = $("countStat");
  ui.selectedStat = $("selectedStat");

  ui.q = $("q");
  ui.chatList = $("chatList");
  ui.queueStatus = $("queueStatus");
}

function log(obj) {
  const text = typeof obj === "string" ? obj : JSON.stringify(obj, null, 2);
  console.log("[APP]", text);
}

function appendLogLine(line) {
  console.log("[APP]", String(line));
}

function assertUiIds() {
  refreshUiRefs();

  const must = [
    "loginView",
    "appView",
    "btnLogin",
    "btnChats",
    "btnExport",
    "mePill",
    "countStat",
    "selectedStat",
    "q",
    "chatList",
    "queueStatus",
  ];

  const missing = must.filter((id) => !$(id));
  if (missing.length) {
    console.warn("[UI] missing ids:", missing);
    log("UI missing element IDs:\n" + missing.join(", "));
    return false;
  }
  return true;
}

function showLoginView() {
  ui.loginView?.classList.remove("hidden");
  ui.appView?.classList.add("hidden");
}

function showAppView() {
  ui.loginView?.classList.add("hidden");
  ui.appView?.classList.remove("hidden");
}

function ensureExtraButtons() {
  const host = ui.btnExport?.parentElement;
  if (!host) return;

  if (!ui.btnQueueAll) {
    const btn = document.createElement("button");
    btn.id = "btnQueueAll";
    btn.type = "button";
    btn.textContent = "Queue all for export";
    btn.className = "btn";
    host.appendChild(btn);
  }

  if (!ui.btnRetryFailed) {
    const btn = document.createElement("button");
    btn.id = "btnRetryFailed";
    btn.type = "button";
    btn.textContent = "Retry failed only";
    btn.className = "btn btnDanger";
    host.appendChild(btn);
  }

  refreshUiRefs();
}

/** =========================
 * STATE
 * ========================= */
let _me = null;
let _chats = [];
let _selectedChatIds = new Set();

let _queueItems = [];
let _queueRunning = false;

let _loginAtMs = 0;
let _exportStartedSinceLogin = false;
let _sessionWatchHandle = null;
let _sessionReauthInProgress = false;
let _authInited = false;
let _uiEventsBound = false;

/** =========================
 * HELPERS
 * ========================= */
function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function formatDDMMYYYY_HHMMSS(d) {
  return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
}

function formatShortDateTime(value) {
  if (!value) return "Няма данни";
  const d = new Date(value);
  if (isNaN(d.getTime())) return "Няма данни";
  return formatDDMMYYYY_HHMMSS(d);
}

function getChatStartedLabel(chat) {
  return formatShortDateTime(chat?.createdDateTime);
}

function getChatLastMessageLabel(chat) {
  return formatShortDateTime(chat?.lastMessagePreview?.createdDateTime || "");
}

function esc(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function safeFileName(name) {
  const n = String(name || "").trim() || "Teams_Chat_Archive";
  return n.replace(/[\/\\:*?"<>|]/g, "_").replace(/\s+/g, " ").trim();
}

function makeUniqueFileName(desiredName, usedSet) {
  const safe = safeFileName(desiredName || "Teams_Chat_Archive");
  const m = safe.match(/^(.*?)(\.[a-z0-9]{1,12})$/i);
  const base = m ? m[1] : safe;
  const ext = m ? m[2] : "";

  let candidate = safe;
  let n = 2;

  while (usedSet.has(candidate.toLowerCase())) {
    candidate = `${base} (${n++})${ext}`;
  }

  usedSet.add(candidate.toLowerCase());
  return candidate;
}

function extOf(name) {
  const m = String(name || "").toLowerCase().match(/\.([a-z0-9]{1,12})$/);
  return m ? m[1] : "";
}

function normalizeErrorMessage(errorLike) {
  if (!errorLike) return "Unknown error";
  if (typeof errorLike === "string") return errorLike;
  return errorLike?.message || String(errorLike) || "Unknown error";
}

function stripToPreviewText(html, maxLen = 240) {
  try {
    const doc = new DOMParser().parseFromString(String(html || ""), "text/html");
    const t = (doc.body?.textContent || "").replace(/\s+/g, " ").trim();
    if (!t) return "";
    return t.length > maxLen ? `${t.slice(0, maxLen - 1)}…` : t;
  } catch {
    const t = String(html || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
    return t.length > maxLen ? `${t.slice(0, maxLen - 1)}…` : t;
  }
}

function normalizeTeamsHtml(html) {
  let src = String(html || "");
  if (!src) return "";

  src = src.replace(
    /<emoji\b[^>]*\balt="([^"]+)"[^>]*><\/emoji>/gi,
    (_, alt) => esc(alt || "")
  );

  src = src.replace(
    /<emoji\b[^>]*\balt='([^']+)'[^>]*><\/emoji>/gi,
    (_, alt) => esc(alt || "")
  );

  src = src.replace(/<emoji\b[^>]*><\/emoji>/gi, "🙂");

  return src;
}

function normalizeId(value) {
  return String(value || "").trim().toLowerCase();
}

function collectMemberCandidateIds(member) {
  return [
    member?.userId,
    member?.id,
    member?.user?.id,
    member?.email,
    member?.user?.email,
    member?.upn,
    member?.userPrincipalName,
  ]
    .map((x) => String(x || "").trim())
    .filter(Boolean);
}

function mergeMembers(primaryMembers, secondaryMembers) {
  const out = [];
  const seen = new Set();

  for (const member of [...(primaryMembers || []), ...(secondaryMembers || [])]) {
    if (!member) continue;

    const keys = collectMemberCandidateIds(member)
      .map((x) => normalizeId(x))
      .filter(Boolean);

    const stableKey =
      keys[0] ||
      normalizeId(member?.displayName) ||
      Math.random().toString(36).slice(2);

    if (seen.has(stableKey)) continue;

    seen.add(stableKey);
    out.push(member);
  }

  return out;
}

function getReactionVisual(reactionType) {
  const raw = String(reactionType || "").trim();
  const k = raw.toLowerCase();

  const known = {
    like: "👍",
    heart: "❤️",
    laugh: "😆",
    surprised: "😮",
    sad: "😢",
    angry: "😡",
  };

  if (known[k]) return known[k];

  // If API already gives an actual emoji, keep it as-is
  const looksLikeEmoji =
    /[\p{Extended_Pictographic}\u2600-\u27BF]/u.test(raw);

  if (looksLikeEmoji) return raw;

  // Keep unknown reaction types distinct instead of collapsing to 👍
  return raw || "❔";
}

function resolveUserDisplayNameById(userId, members, me) {
  const uid = normalizeId(userId);
  if (!uid) return "";

  const myIds = [
    me?.id,
    me?.userPrincipalName,
    me?.mail,
  ]
    .map((x) => normalizeId(x))
    .filter(Boolean);

  if (myIds.includes(uid)) {
    return me?.displayName || me?.userPrincipalName || "";
  }

  const hit = (members || []).find((m) => {
    const ids = collectMemberCandidateIds(m)
      .map((x) => normalizeId(x))
      .filter(Boolean);

    return ids.includes(uid);
  });

  return (
    hit?.displayName ||
    hit?.user?.displayName ||
    hit?.email ||
    hit?.userPrincipalName ||
    ""
  );
}

function resolveReactionUserDisplayName(reaction, members, me) {
  const directName = String(reaction?.user?.displayName || "").trim();
  if (directName) return directName;

  const userId =
    reaction?.user?.id ||
    reaction?.userId ||
    "";

  const resolved = resolveUserDisplayNameById(userId, members, me);
  if (resolved) return resolved;

  if (me?.id && normalizeId(me.id) === normalizeId(userId)) {
    return me.displayName || me.userPrincipalName || "";
  }

  return "";
}

function isForwardedReferenceAttachment(a) {
  return String(a?.contentType || "").toLowerCase() === "forwardedmessagereference";
}

function parseRetryAfterMs(value) {
  if (!value) return 0;
  const n = Number(value);
  if (Number.isFinite(n) && n > 0) return n * 1000;
  const d = Date.parse(value);
  if (Number.isFinite(d)) {
    const delta = d - Date.now();
    return delta > 0 ? delta : 0;
  }
  return 0;
}

function getBackoffDelayMs(attempt, retryAfterMs = 0) {
  if (retryAfterMs > 0) return Math.min(retryAfterMs, RETRY.MAX_DELAY_MS);
  const exp = Math.min(
    RETRY.BASE_DELAY_MS * Math.pow(2, Math.max(0, attempt - 1)),
    RETRY.MAX_DELAY_MS
  );
  return exp + Math.floor(Math.random() * 350);
}

function isRetriableHttpStatus(status) {
  return status === 408 || status === 409 || status === 425 || status === 429 || (status >= 500 && status <= 599);
}

function classifyFailure(message) {
  const m = String(message || "").toLowerCase();

  if (m.includes("429")) {
    return {
      code: "throttled",
      human: "Graph/API throttling (429). Service temporarily limited request rate.",
    };
  }
  if (m.includes("401")) {
    return {
      code: "unauthorized",
      human: "Unauthorized (401). Sign-in session or token is no longer valid.",
    };
  }
  if (m.includes("403") || m.includes("accessdenied")) {
    return {
      code: "access_denied",
      human: "Access denied (403). Current account cannot access this resource.",
    };
  }
  if (m.includes("404")) {
    return {
      code: "not_found",
      human: "Resource not found (404). Chat, message, file or drive item no longer exists.",
    };
  }
  if (m.includes("timeout") || m.includes("network") || m.includes("failed to fetch")) {
    return {
      code: "network",
      human: "Network or connectivity issue during request execution.",
    };
  }
  if (m.includes("too_large_each") || m.includes("too_large_total")) {
    return {
      code: "size_limit",
      human: "File or image exceeded configured embedding size limits.",
    };
  }
  if (m.includes("driveitem_missing_ids")) {
    return {
      code: "drive_item_resolution_failed",
      human: "Could not resolve the shared file into a valid drive item.",
    };
  }
  if (m.includes("blob_empty") || m.includes("inline_empty")) {
    return {
      code: "empty_content",
      human: "The API returned empty content for this resource.",
    };
  }
  if (m.includes("denied_extension")) {
    return {
      code: "blocked_extension",
      human: "Attachment type is blocked by extension policy.",
    };
  }

  return {
    code: "generic_failure",
    human: `Operation failed: ${normalizeErrorMessage(message)}`,
  };
}

function formatFailureForLog(queueItem) {
  const stage = queueItem?.errorStage ? `stage=${queueItem.errorStage}` : "stage=unknown";
  const code = queueItem?.errorCode ? `code=${queueItem.errorCode}` : "code=unknown";
  const msg = queueItem?.error ? `message=${queueItem.error}` : "message=Unknown error";
  return `${stage} • ${code} • ${msg}`;
}

function getArchiveRange(items) {
  const valid = (items || []).filter((x) => Number.isFinite(x?.dtMs) && x.dtMs > 0);
  if (!valid.length) {
    return {
      fromMs: 0,
      toMs: 0,
      fromStr: "",
      toStr: "",
      label: "Няма валидни timestamps",
    };
  }

  const fromMs = Math.min(...valid.map((x) => x.dtMs));
  const toMs = Math.max(...valid.map((x) => x.dtMs));

  return {
    fromMs,
    toMs,
    fromStr: formatDDMMYYYY_HHMMSS(new Date(fromMs)),
    toStr: formatDDMMYYYY_HHMMSS(new Date(toMs)),
    label: `${formatDDMMYYYY_HHMMSS(new Date(fromMs))} → ${formatDDMMYYYY_HHMMSS(new Date(toMs))}`,
  };
}

async function asyncPool(limit, items, iteratorFn) {
  const ret = [];
  const executing = [];

  for (const item of items) {
    const p = Promise.resolve().then(() => iteratorFn(item));
    ret.push(p);

    if (limit <= items.length) {
      const e = p.then(() => executing.splice(executing.indexOf(e), 1));
      executing.push(e);
      if (executing.length >= limit) {
        await Promise.race(executing);
      }
    }
  }

  return Promise.all(ret);
}

/** =========================
 * QUEUE UI HELPERS
 * ========================= */
function getQueueSummaryCounts() {
  const counts = { queued: 0, running: 0, done: 0, failed: 0 };
  for (const item of _queueItems) {
    if (counts[item.status] != null) counts[item.status] += 1;
  }
  return counts;
}

function updateQueueStatus() {
  if (!ui.queueStatus) return;

  if (!_queueItems.length) {
    ui.queueStatus.classList.add("hidden");
    ui.queueStatus.textContent = "";
    return;
  }

  const counts = getQueueSummaryCounts();
  const total = _queueItems.length;
  const done = counts.done || 0;
  const running = counts.running || 0;
  const failed = counts.failed || 0;

  ui.queueStatus.textContent =
    `Queue progress: ${done} done • ${running} running • ${failed} failed • ${total} total`;

  ui.queueStatus.classList.remove("hidden");
}

function getQueueProgressText() {
  if (!_queueRunning) return "";

  const total = _queueItems.length;
  const done = _queueItems.filter((x) => x.status === "done").length;
  const failed = _queueItems.filter((x) => x.status === "failed").length;
  const running = _queueItems.filter((x) => x.status === "running").length;
  const processed = done + failed;

  if (running > 0) {
    return `Exporting ${Math.min(processed + 1, total)} / ${total}`;
  }

  if (processed < total) {
    return `Preparing queue… ${processed} / ${total}`;
  }

  return `Queue finished ${total} / ${total}`;
}

function setMePill() {
  if (!ui.mePill) return;

  if (!_me) {
    ui.mePill.textContent = "Not signed in";
    return;
  }

  ui.mePill.textContent = `Signed in: ${_me.displayName || _me.userPrincipalName || _me.id}`;
}

function setCounts() {
  if (ui.countStat) ui.countStat.textContent = String(_chats.length);
  if (ui.selectedStat) ui.selectedStat.textContent = String(_selectedChatIds.size);

  setMePill();
  updateQueueStatus();

  if (!_queueRunning) {
    setBusy(false);
  }
}

function setBusy(isBusy) {
  const b = !!isBusy || _queueRunning;

  if (ui.btnLogin) ui.btnLogin.disabled = b || _sessionReauthInProgress;
  if (ui.btnChats) ui.btnChats.disabled = b;

  if (ui.btnExport) {
    ui.btnExport.disabled = b || !_me || _selectedChatIds.size === 0 || _chats.length === 0;
    ui.btnExport.textContent = _queueRunning
      ? getQueueProgressText()
      : `Export selected${_selectedChatIds.size > 0 ? ` (${_selectedChatIds.size})` : ""}`;
  }

  if (ui.btnQueueAll) {
    ui.btnQueueAll.disabled = b || !_me || _chats.length === 0;
    ui.btnQueueAll.textContent = _queueRunning
      ? getQueueProgressText()
      : `Queue all for export${_chats.length > 0 ? ` (${_chats.length})` : ""}`;
  }

  if (ui.btnRetryFailed) {
    const failedCount = getQueueSummaryCounts().failed;
    ui.btnRetryFailed.disabled = b || !_me || failedCount === 0;
    ui.btnRetryFailed.textContent = _queueRunning
      ? getQueueProgressText()
      : `Retry failed only${failedCount > 0 ? ` (${failedCount})` : ""}`;
  }
}

/** =========================
 * SESSION HELPERS
 * ========================= */
function loadAppSessionMeta() {
  const loginTs = Number(sessionStorage.getItem(APP_SESSION_KEYS.loginTs) || "0");
  const exportStarted = sessionStorage.getItem(APP_SESSION_KEYS.exportStarted) === "1";

  _loginAtMs = Number.isFinite(loginTs) ? loginTs : 0;
  _exportStartedSinceLogin = !!exportStarted;
}

function saveAppSessionMeta() {
  if (_loginAtMs > 0) sessionStorage.setItem(APP_SESSION_KEYS.loginTs, String(_loginAtMs));
  else sessionStorage.removeItem(APP_SESSION_KEYS.loginTs);

  if (_exportStartedSinceLogin) sessionStorage.setItem(APP_SESSION_KEYS.exportStarted, "1");
  else sessionStorage.removeItem(APP_SESSION_KEYS.exportStarted);
}

function clearAppSessionMeta() {
  _loginAtMs = 0;
  _exportStartedSinceLogin = false;
  sessionStorage.removeItem(APP_SESSION_KEYS.loginTs);
  sessionStorage.removeItem(APP_SESSION_KEYS.exportStarted);
}

function stopSessionWatcher() {
  if (_sessionWatchHandle) {
    clearInterval(_sessionWatchHandle);
    _sessionWatchHandle = null;
  }
}

function recordAuthenticatedSession({ reset = false } = {}) {
  loadAppSessionMeta();

  if (reset || !_loginAtMs) _loginAtMs = Date.now();
  if (reset) _exportStartedSinceLogin = false;

  saveAppSessionMeta();
  armSessionWatcher();
  setMePill();
}

function markExportStartedForSession() {
  _exportStartedSinceLogin = true;
  saveAppSessionMeta();
  setMePill();
}

function getPreExportMsRemaining() {
  if (_exportStartedSinceLogin) return Number.POSITIVE_INFINITY;
  if (!_loginAtMs) return 0;
  return _loginAtMs + SESSION_PREEXPORT_WINDOW_MS - Date.now();
}

function resetUiForSignedOutState() {
  _me = null;
  _chats = [];
  _selectedChatIds = new Set();
  _queueItems = [];
  renderChats();
  setCounts();
  showLoginView();
}

async function forceReLogin(reason) {
  if (_queueRunning) return;
  if (_sessionReauthInProgress) return;

  _sessionReauthInProgress = true;
  stopSessionWatcher();
  appendLogLine(reason);

  try {
    const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;

    clearAppSessionMeta();
    resetUiForSignedOutState();

    if (account) {
      await msalInstance.logoutRedirect({ account });
      return;
    }

    await login();
  } catch (e) {
    console.error(e);
    appendLogLine(`Re-login redirect failed: ${normalizeErrorMessage(e)}`);
    alert(`Re-login failed:\n${normalizeErrorMessage(e)}`);
  } finally {
    _sessionReauthInProgress = false;
  }
}

function checkSessionPolicy() {
  if (!_me) return;
  if (_queueRunning) return;
  if (_exportStartedSinceLogin) {
    setMePill();
    return;
  }

  const remaining = getPreExportMsRemaining();
  if (remaining <= 0) {
    void forceReLogin("Session expired: no export was started within 3 minutes. Please sign in again.");
    return;
  }

  setMePill();
}

function armSessionWatcher() {
  stopSessionWatcher();
  if (!_me) return;
  _sessionWatchHandle = setInterval(checkSessionPolicy, SESSION_CHECK_EVERY_MS);
  checkSessionPolicy();
}

/** =========================
 * MSAL AUTH
 * ========================= */
async function initAuth() {
  if (_authInited) return msalInstance.getActiveAccount() || null;

  await msalInstance.initialize();
  loadAppSessionMeta();

  const resp = await msalInstance.handleRedirectPromise();
  if (resp?.account) {
    msalInstance.setActiveAccount(resp.account);
    recordAuthenticatedSession({ reset: true });
  }

  const acc = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (acc) msalInstance.setActiveAccount(acc);

  _authInited = true;
  return msalInstance.getActiveAccount() || null;
}

async function login() {
  try {
    if (!_authInited) {
      await initAuth();
    }

    appendLogLine("Starting Microsoft login redirect…");

    await msalInstance.loginRedirect({
      scopes: SCOPES,
      prompt: "select_account",
      redirectStartPage: window.location.href,
    });
  } catch (e) {
    console.error(e);
    const msg = normalizeErrorMessage(e);
    appendLogLine(`Login failed: ${msg}`);
    alert(`Login failed:\n${msg}`);
  }
}

async function getAccessToken() {
  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!account) throw new Error("No active account. Click Login first.");

  try {
    const res = await msalInstance.acquireTokenSilent({ account, scopes: SCOPES });
    return res.accessToken;
  } catch {
    await msalInstance.acquireTokenRedirect({ account, scopes: SCOPES });
    throw new Error("Redirecting for token...");
  }
}

/** =========================
 * GRAPH HELPERS
 * ========================= */
async function graphFetch(url, { method = "GET", headers = {}, body } = {}) {
  const token = await getAccessToken();
  const fullUrl = url.startsWith("http") ? url : `${GRAPH_BASE}${url}`;

  let lastError = null;

  for (let attempt = 1; attempt <= RETRY.GRAPH_MAX_ATTEMPTS; attempt++) {
    try {
      const res = await fetch(fullUrl, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          ...headers,
        },
        body,
      });

      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        if (attempt < RETRY.GRAPH_MAX_ATTEMPTS && isRetriableHttpStatus(res.status)) {
          const retryAfterMs = parseRetryAfterMs(res.headers.get("retry-after"));
          const delay = getBackoffDelayMs(attempt, retryAfterMs);
          appendLogLine(`[retry] Graph ${res.status} on ${method} ${fullUrl} — attempt ${attempt}/${RETRY.GRAPH_MAX_ATTEMPTS}, waiting ${delay}ms`);
          await sleep(delay);
          continue;
        }

        throw new Error(`Graph ${res.status} ${res.statusText}: ${txt.slice(0, 900)}`);
      }

      const ct = res.headers.get("content-type") || "";
      if (ct.includes("application/json")) return await res.json();
      return res;
    } catch (e) {
      lastError = e;
      const msg = normalizeErrorMessage(e).toLowerCase();
      const retriableNetwork =
        msg.includes("failed to fetch") ||
        msg.includes("networkerror") ||
        msg.includes("network") ||
        msg.includes("load failed") ||
        msg.includes("timeout");

      if (attempt < RETRY.GRAPH_MAX_ATTEMPTS && retriableNetwork) {
        const delay = getBackoffDelayMs(attempt);
        appendLogLine(`[retry] Graph/network on ${method} ${fullUrl} — attempt ${attempt}/${RETRY.GRAPH_MAX_ATTEMPTS}, waiting ${delay}ms`);
        await sleep(delay);
        continue;
      }

      throw lastError;
    }
  }

  throw lastError || new Error("Graph request failed");
}

async function graphFetchAllPages(firstUrl) {
  const out = [];
  let url = firstUrl;
  let pages = 0;

  while (url && pages++ < LIMITS.MAX_PAGES) {
    const data = await graphFetch(url);
    out.push(...(data.value || []));
    url = data["@odata.nextLink"] || null;
  }

  return out;
}

/** =========================
 * DATA
 * ========================= */
async function loadMe() {
  return await graphFetch("/me");
}

async function loadMyChats() {
  return await graphFetchAllPages("/me/chats?$top=50&$expand=members,lastMessagePreview");
}

async function loadChatMembers(chatId) {
  return await graphFetchAllPages(`/chats/${encodeURIComponent(chatId)}/members?$top=50`);
}

async function loadAllChatMessages(chatId) {
  return await graphFetchAllPages(`/chats/${encodeURIComponent(chatId)}/messages?$top=50`);
}

const _chatMembersCache = new Map();
const _chatMessageByIdCache = new Map();

async function loadChatMembersCached(chatId) {
  const key = String(chatId || "").trim();
  if (!key) return [];

  if (_chatMembersCache.has(key)) {
    return _chatMembersCache.get(key);
  }

  const members = await loadChatMembers(key);
  _chatMembersCache.set(key, members || []);
  return members || [];
}

async function loadSingleChatMessage(chatId, messageId) {
  const cid = String(chatId || "").trim();
  const mid = String(messageId || "").trim();
  if (!cid || !mid) return null;

  const key = `${cid}::${mid}`;
  if (_chatMessageByIdCache.has(key)) {
    return _chatMessageByIdCache.get(key);
  }

  const msg = await graphFetch(`/chats/${encodeURIComponent(cid)}/messages/${encodeURIComponent(mid)}`);
  _chatMessageByIdCache.set(key, msg || null);
  return msg || null;
}

/** =========================
 * CHAT / MESSAGE HELPERS
 * ========================= */
function chatDisplayName(chat, members) {
  const topic = (chat?.topic || "").trim();
  if (topic) return topic;

  const names = (members || [])
    .map((m) => (m?.displayName || "").trim())
    .filter(Boolean);

  return names.length ? names.join(", ") : "Chat";
}

function getChatTitle(chat) {
  const topic = (chat?.topic || "").trim();
  if (topic) return topic;

  const chatType = String(chat?.chatType || "").toLowerCase();
  const names = (chat?.members || [])
    .map((m) => (m?.displayName || "").trim())
    .filter(Boolean);

  if (chatType === "oneonone") {
    const myName = (_me?.displayName || "").trim();
    const other = names.filter((n) => !myName || n !== myName);

    if (other.length === 1) return other[0];
    if (other.length > 1) return other.join(", ");
    return names[0] || "Chat";
  }

  return names.length ? names.join(", ") : "Chat";
}

function fileBaseForChat(chat, me, members) {
  const chatType = String(chat?.chatType || "").toLowerCase();

  if (chatType === "group" || chatType === "meeting") {
    return safeFileName(chat?.topic || "Group chat");
  }

  if (chatType === "oneonone") {
    const myName = (me?.displayName || "").trim();
    const names = (members || [])
      .map((m) => (m?.displayName || "").trim())
      .filter(Boolean)
      .filter((n) => !myName || n !== myName);

    return safeFileName(names.length ? names.join(" & ") : "1on1_chat");
  }

  return safeFileName(chatDisplayName(chat, members) || "Chat");
}

function extractSender(msg) {
  return msg?.from?.user?.displayName || msg?.from?.application?.displayName || "";
}

function isFromMe(msg, myUserId) {
  const uid = msg?.from?.user?.id || "";
  return !!(uid && myUserId && uid === myUserId);
}

function extractBodyHtml(msg) {
  const b = msg?.body;
  if (!b?.content) return "";
  if ((b.contentType || "").toLowerCase() === "text") {
    return `<div>${esc(b.content)}</div>`;
  }
  return String(b.content);
}

function shouldAttemptHosted(bodyHtml) {
  const h = String(bodyHtml || "");
  return h.includes("<img") || h.includes("img ");
}

/** =========================
 * REACTIONS
 * ========================= */
function buildReactionsFromGraph(msg, members, me) {
  const rx = Array.isArray(msg?.reactions) ? msg.reactions : [];
  if (!rx.length) return "";

  const acc = new Map();

  for (const r of rx) {
    const reactionType = String(r?.reactionType || "").trim();
    const visual = getReactionVisual(reactionType);
    const key = reactionType || visual;

    const name = resolveReactionUserDisplayName(r, members, me);

    if (!acc.has(key)) {
      acc.set(key, {
        visual,
        reactionType,
        count: 0,
        names: [],
      });
    }

    const v = acc.get(key);
    v.count += 1;
    if (name) v.names.push(name);
  }

  let out = `<div class="reactions">`;

  for (const [, v] of [...acc.entries()].sort((a, b) => {
    return (
      b[1].count - a[1].count ||
      String(a[1].visual).localeCompare(String(b[1].visual))
    );
  })) {
    const uniqNames = [...new Set(v.names)].slice(0, 100);
    const joinedNames = uniqNames.join(" | ");

    out += `
      <span
        class="reactChip"
        data-emoji="${esc(v.visual)}"
        data-reaction-type="${esc(v.reactionType || v.visual)}"
        data-count="${v.count}"
        data-names="${esc(joinedNames)}"
        title="Show reactions"
      >
        <span class="e">${esc(v.visual)}</span>
        <span class="c">${v.count}</span>
      </span>
    `;
  }

  out += `</div>`;
  return out;
}

/** =========================
 * REPLY / QUOTES / FORWARDED
 * ========================= */
function buildReplyPreview(dto, byId) {
  const rid = dto.replyToId;
  if (!rid) return null;
  const target = byId.get(rid);
  if (!target) return null;

  return {
    author: target.sender || "Microsoft Teams",
    ts: target.dtStr || "",
    previewText: stripToPreviewText(target.bodyHtml || "", 220),
    refId: rid,
  };
}

function extractBlockquoteQuotes(bodyHtml) {
  const quotes = [];
  if (!bodyHtml) return quotes;

  try {
    const doc = new DOMParser().parseFromString(String(bodyHtml), "text/html");
    const nodes = Array.from(doc.querySelectorAll("blockquote")).slice(0, 8);
    for (const bq of nodes) {
      const txt = (bq.textContent || "").replace(/\s+/g, " ").trim();
      if (!txt) continue;
      quotes.push({ author: "", ts: "", previewText: txt, refId: "" });
    }
  } catch {}

  return quotes;
}

function extractForwardFromBody(bodyHtml) {
  const html = String(bodyHtml || "");
  if (!html) return null;

  const plain = html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
  const hasForwardWord =
    plain.includes("forwarded") ||
    plain.includes("препрат") ||
    plain.includes("препратено");

  if (!hasForwardWord) return null;

  let originalAuthor = "";
  const m1 = plain.match(/\bfrom\s+([a-z0-9а-яёіїє\-_.\s]{2,60})\b/i);
  const m2 = plain.match(/\bот\s+([a-z0-9а-яёіїє\-_.\s]{2,60})\b/i);
  const cand = (m1?.[1] || m2?.[1] || "").trim();
  if (cand && cand.length <= 60) originalAuthor = cand;

  return { isForwarded: true, originalAuthor, originalTs: "", originalBodyHtml: "" };
}

function renderForwardBlock(fwd) {
  if (!fwd?.isForwarded) return "";

  const meta = [fwd.originalAuthor, fwd.originalTs].filter(Boolean).join(" • ") || "—";
  const bodyHtml = fwd.originalBodyHtml || "";

  return `
    <div class="fwdWrap">
      <div class="fwdHeader">
        <span class="fwdChip">Препратено</span>
        <div class="fwdMeta">${esc(meta)}</div>
      </div>
      <div class="fwdBody">
        <div class="fwdInner">
          ${bodyHtml || `<div class="fwdEmpty">Няма съдържание</div>`}
        </div>
      </div>
    </div>
  `;
}

function renderQuoteStack(quotes) {
  const q = Array.isArray(quotes) ? quotes.filter(Boolean) : [];
  if (!q.length) return "";

  let out = `<div class="quotePreview">`;
  for (const item of q.slice(0, 10)) {
    const headMeta = [item.author, item.ts].filter(Boolean).join(" • ");
    const head = headMeta
      ? `<div class="quoteHead">
          <span class="qAuthor">${esc(item.author || "Microsoft Teams")}</span>
          <span class="qTs">${esc(item.ts || "")}</span>
        </div>`
      : `<div class="quoteHead">
          <span class="qAuthor">${esc(item.author || "Цитат")}</span>
        </div>`;

    const jump = item.refId ? ` <a class="qJump" href="#mid-${esc(item.refId)}">Виж</a>` : "";

    out += `
      <div class="quoteBox">
        ${head}${jump}
        <div class="qText">${esc(item.previewText || "")}</div>
      </div>
    `;
  }
  out += `</div>`;
  return out;
}

/** =========================
 * HOSTED CONTENT
 * ========================= */
async function listHostedContents(chatId, messageId) {
  const data = await graphFetch(`/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/hostedContents`);
  return data.value || [];
}

async function getHostedContentBlob(chatId, messageId, hostedId) {
  const token = await getAccessToken();
  const url =
    `${GRAPH_BASE}/chats/${encodeURIComponent(chatId)}` +
    `/messages/${encodeURIComponent(messageId)}` +
    `/hostedContents/${encodeURIComponent(hostedId)}/$value`;

  let lastError = null;

  for (let attempt = 1; attempt <= RETRY.BINARY_MAX_ATTEMPTS; attempt++) {
    try {
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });

      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        if (attempt < RETRY.BINARY_MAX_ATTEMPTS && isRetriableHttpStatus(res.status)) {
          const retryAfterMs = parseRetryAfterMs(res.headers.get("retry-after"));
          const delay = getBackoffDelayMs(attempt, retryAfterMs);
          appendLogLine(`[retry] hostedContent ${res.status} — attempt ${attempt}/${RETRY.BINARY_MAX_ATTEMPTS}, waiting ${delay}ms`);
          await sleep(delay);
          continue;
        }

        throw new Error(`HostedContent $value ${res.status}: ${txt.slice(0, 700)}`);
      }

      return {
        blob: await res.blob(),
        contentType: res.headers.get("content-type") || "application/octet-stream",
      };
    } catch (e) {
      lastError = e;
      const msg = normalizeErrorMessage(e).toLowerCase();
      const retriableNetwork = msg.includes("failed to fetch") || msg.includes("network") || msg.includes("timeout");

      if (attempt < RETRY.BINARY_MAX_ATTEMPTS && retriableNetwork) {
        const delay = getBackoffDelayMs(attempt);
        appendLogLine(`[retry] hostedContent/network — attempt ${attempt}/${RETRY.BINARY_MAX_ATTEMPTS}, waiting ${delay}ms`);
        await sleep(delay);
        continue;
      }

      throw lastError;
    }
  }

  throw lastError || new Error("Hosted content download failed");
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onerror = () => reject(new Error("FileReader failed"));
    r.onload = () => resolve(String(r.result || ""));
    r.readAsDataURL(blob);
  });
}

function tokenizeHostedImages(bodyHtml, hostedList) {
  const html = String(bodyHtml || "");
  const hostedIds = (hostedList || []).map((h) => h?.id).filter(Boolean);
  if (!html || !hostedIds.length) return { html, hostedIds: [] };

  try {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const imgs = Array.from(doc.querySelectorAll("img"));

    for (const img of imgs) {
      const src = img.getAttribute("src") || "";
      const hit = hostedIds.find((id) => src.includes(id));
      if (!hit) continue;
      img.setAttribute("data-hid", hit);
      if (!img.getAttribute("data-src-orig")) img.setAttribute("data-src-orig", src);
    }

    return { html: doc.body.innerHTML || html, hostedIds };
  } catch {
    return { html, hostedIds };
  }
}

async function embedHostedImages(chatId, messageId, bodyHtml, hostedList, stats) {
  const html = String(bodyHtml || "");
  if (!html || !hostedList?.length) return html;

  let doc;
  try {
    doc = new DOMParser().parseFromString(html, "text/html");
  } catch {
    return html;
  }

  const imgs = Array.from(doc.querySelectorAll('img[data-hid]'));
  if (!imgs.length) return doc.body.innerHTML || html;

  let total = stats.imagesBytes || 0;

  await asyncPool(LIMITS.DL_CONCURRENCY, imgs, async (img) => {
    if (total >= LIMITS.MAX_IMAGE_BYTES_TOTAL) return;
    const hid = img.getAttribute("data-hid");
    if (!hid) return;

    try {
      const got = await getHostedContentBlob(chatId, messageId, hid);
      const blob = got?.blob;
      if (!blob) return;

      const size = blob.size || 0;
      if (size <= 0) throw new Error("blob_empty");
      if (size > LIMITS.MAX_IMAGE_BYTES_EACH) throw new Error("too_large_each");
      if (total + size > LIMITS.MAX_IMAGE_BYTES_TOTAL) throw new Error("too_large_total");

      const ct = got.contentType || blob.type || "application/octet-stream";
      const dataUrl = await blobToDataUrl(new Blob([blob], { type: ct }));
      if (!dataUrl.startsWith("data:")) throw new Error("dataurl_invalid");

      img.setAttribute("src", dataUrl);
      img.removeAttribute("crossorigin");

      stats.imagesEmbedded = (stats.imagesEmbedded || 0) + 1;
      total += size;
      stats.imagesBytes = total;
    } catch (e) {
      stats.failures.push({
        kind: "hostedImage",
        stage: "embedHostedImages",
        messageId,
        hostedId: hid,
        error: normalizeErrorMessage(e),
      });
      const orig = img.getAttribute("data-src-orig");
      if (orig) img.setAttribute("src", orig);
    }
  });

  return doc.body.innerHTML || html;
}

/** =========================
 * ATTACHMENTS
 * ========================= */
function toShareIdFromUrl(sharingUrl) {
  const b64 = btoa(unescape(encodeURIComponent(sharingUrl)));
  const b64url = b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return `u!${b64url}`;
}

async function getDriveItemFromSharingUrl(sharingUrl) {
  const shareId = toShareIdFromUrl(sharingUrl);
  return await graphFetch(`/shares/${encodeURIComponent(shareId)}/driveItem`);
}

async function downloadDriveItemContent(driveId, itemId) {
  const token = await getAccessToken();
  const url = `${GRAPH_BASE}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`;

  let lastError = null;

  for (let attempt = 1; attempt <= RETRY.BINARY_MAX_ATTEMPTS; attempt++) {
    try {
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });

      if (!res.ok) {
        const t = await res.text().catch(() => "");
        if (attempt < RETRY.BINARY_MAX_ATTEMPTS && isRetriableHttpStatus(res.status)) {
          const retryAfterMs = parseRetryAfterMs(res.headers.get("retry-after"));
          const delay = getBackoffDelayMs(attempt, retryAfterMs);
          appendLogLine(`[retry] driveContent ${res.status} — attempt ${attempt}/${RETRY.BINARY_MAX_ATTEMPTS}, waiting ${delay}ms`);
          await sleep(delay);
          continue;
        }

        throw new Error(`DriveItem content ${res.status}: ${t.slice(0, 700)}`);
      }

      return await res.blob();
    } catch (e) {
      lastError = e;
      const msg = normalizeErrorMessage(e).toLowerCase();
      const retriableNetwork = msg.includes("failed to fetch") || msg.includes("network") || msg.includes("timeout");

      if (attempt < RETRY.BINARY_MAX_ATTEMPTS && retriableNetwork) {
        const delay = getBackoffDelayMs(attempt);
        appendLogLine(`[retry] driveContent/network — attempt ${attempt}/${RETRY.BINARY_MAX_ATTEMPTS}, waiting ${delay}ms`);
        await sleep(delay);
        continue;
      }

      throw lastError;
    }
  }

  throw lastError || new Error("Drive item download failed");
}

function isLikelyQuoteOrReferenceAttachment(a) {
  const ct = String(a?.contentType || "").toLowerCase();
  return (
    (ct.includes("message") && ct.includes("reference")) ||
    ct.includes("application/vnd.microsoft.card") ||
    ct.includes("application/vnd.microsoft.teams.card") ||
    ct.includes("adaptive") ||
    ct.includes("quote")
  );
}

function tryParseAttachmentContentJson(a) {
  const raw = a?.content;
  if (!raw) return null;
  if (typeof raw === "object") return raw;
  try {
    return JSON.parse(String(raw));
  } catch {
    return null;
  }
}

function extractForwardedAttachmentData(attachments, members, me) {
  const forwarded = [];

  for (const a of Array.isArray(attachments) ? attachments : []) {
    if (!isForwardedReferenceAttachment(a)) continue;

    const obj = tryParseAttachmentContentJson(a);
    if (!obj) continue;

    const originalUser = obj?.originalMessageSender?.user || null;
    const originalUserId = originalUser?.id || "";
    const resolvedName =
  String(originalUser?.displayName || "").trim() ||
  resolveUserDisplayNameById(originalUserId, members, me) ||
  (me?.id && normalizeId(me.id) === normalizeId(originalUserId)
    ? (me.displayName || me.userPrincipalName || "")
    : "") ||
  "Unknown";

    const originalContent = normalizeTeamsHtml(obj?.originalMessageContent || "");
    const originalSentDateTime = obj?.originalSentDateTime || "";
    const originalConversationId = obj?.originalConversationId || "";
    const originalMessageId = obj?.originalMessageId || "";

    let originalTs = "";
    if (originalSentDateTime) {
      const d = new Date(originalSentDateTime);
      if (!isNaN(d.getTime())) originalTs = formatDDMMYYYY_HHMMSS(d);
      else originalTs = String(originalSentDateTime);
    }

    forwarded.push({
      isForwarded: true,
      originalAuthor: resolvedName,
      originalAuthorId: originalUserId,
      originalTs,
      originalBodyHtml: originalContent,
      originalConversationId,
      originalMessageId,
    });
  }

  return forwarded;
}

async function enrichForwardedAuthors(forwardedItems, currentMembers, me) {
  const items = Array.isArray(forwardedItems) ? forwardedItems : [];
  if (!items.length) return items;

  for (const item of items) {
    if (!item?.isForwarded) continue;

    const alreadyResolved =
      item.originalAuthor &&
      String(item.originalAuthor).trim() &&
      String(item.originalAuthor).trim().toLowerCase() !== "unknown";

    if (alreadyResolved) continue;

    const originalUserId = String(item.originalAuthorId || "").trim();
    const originalConversationId = String(item.originalConversationId || "").trim();
    const originalMessageId = String(item.originalMessageId || "").trim();

    let resolved = "";

    if (originalUserId) {
      resolved = resolveUserDisplayNameById(originalUserId, currentMembers, me) || "";
    }

    if (!resolved && originalConversationId && originalMessageId) {
      try {
        const originalMsg = await loadSingleChatMessage(originalConversationId, originalMessageId);
        resolved =
          originalMsg?.from?.user?.displayName ||
          originalMsg?.from?.application?.displayName ||
          "";
      } catch (e) {
        appendLogLine(
          `Forwarded original message lookup failed (${originalConversationId}/${originalMessageId}): ${normalizeErrorMessage(e)}`
        );
      }
    }

    if (!resolved && originalUserId && originalConversationId) {
      try {
        const originalMembers = await loadChatMembersCached(originalConversationId);
        resolved = resolveUserDisplayNameById(originalUserId, originalMembers, me) || "";
      } catch (e) {
        appendLogLine(
          `Forwarded author member lookup failed (${originalConversationId}): ${normalizeErrorMessage(e)}`
        );
      }
    }

    item.originalAuthor = resolved || "Unknown";
  }

  return items;
}

function splitQuotesFromAttachments(attachments) {
  const quotes = [];
  const fileAttachments = [];

  for (const a of Array.isArray(attachments) ? attachments : []) {
    if (isForwardedReferenceAttachment(a)) {
      continue;
    }

    if (!isLikelyQuoteOrReferenceAttachment(a)) {
      fileAttachments.push(a);
      continue;
    }

    const obj = tryParseAttachmentContentJson(a);

    const refId =
      obj?.messageId ||
      obj?.messageid ||
      obj?.replyToId ||
      obj?.replytoid ||
      obj?.referenceMessageId ||
      obj?.referencedMessageId ||
      obj?.id ||
      "";

    const author =
      obj?.from?.user?.displayName ||
      obj?.from?.user?.name ||
      obj?.author ||
      obj?.senderDisplayName ||
      obj?.sender ||
      "";

    const tsRaw =
      obj?.createdDateTime ||
      obj?.createddatetime ||
      obj?.timestamp ||
      obj?.time ||
      "";

    let ts = "";
    if (tsRaw) {
      const d = new Date(tsRaw);
      if (!isNaN(d.getTime())) ts = formatDDMMYYYY_HHMMSS(d);
      else ts = String(tsRaw);
    }

    const preview =
      obj?.preview ||
      obj?.text ||
      obj?.content ||
      obj?.bodyPreview ||
      obj?.messagePreview ||
      obj?.summary ||
      "";

    const previewText = String(preview || "").trim();

    if (refId || previewText || author || ts) {
      quotes.push({
        author: author || "Microsoft Teams",
        ts,
        previewText: previewText ? previewText.slice(0, 220) : "",
        refId: refId ? String(refId) : "",
      });
    }
  }

  return { quotes, fileAttachments };
}

function tokenizeFileAttachments(bodyHtml, attachments, fileTokenToMeta) {
  let html = bodyHtml || "";
  if (!Array.isArray(attachments) || !attachments.length) return html;

  let seq = fileTokenToMeta.__seq || 0;

  for (const a of attachments) {
    const name = (a?.name || a?.contentUrl || a?.id || "attachment").toString().trim();

    if (a?.contentBytes && a?.contentType) {
      const token = `FILE_${++seq}`;
      fileTokenToMeta.set(token, {
        fileName: safeFileName(name),
        kind: "inline",
        contentType: a.contentType,
        contentBytes: a.contentBytes,
        url: "",
      });
      html += `<div data-file-token="${token}">📎 ${esc(name)}</div>`;
      continue;
    }

    const contentUrl = (a?.contentUrl || "").trim();
    if (contentUrl && /^https?:\/\//i.test(contentUrl)) {
      const token = `FILE_${++seq}`;
      fileTokenToMeta.set(token, {
        fileName: safeFileName(name),
        kind: "shareLink",
        url: contentUrl,
      });
      html += `<div data-file-token="${token}">📎 ${esc(name)}</div>`;
      continue;
    }

    const token = `FILE_${++seq}`;
    fileTokenToMeta.set(token, {
      fileName: safeFileName(name),
      kind: "unknown",
      url: "",
    });
    html += `<div data-file-token="${token}">📎 ${esc(name)}</div>`;
  }

  fileTokenToMeta.__seq = seq;
  return html;
}

async function embedFileAttachments(fileTokenToMeta, stats) {
  const tokenToDataUrl = new Map();
  let total = stats.attachBytes || 0;

  const entries = [...fileTokenToMeta.entries()].filter(([k]) => /^FILE_\d+$/.test(k));

  for (const [token, meta] of entries) {
    if (total >= LIMITS.MAX_ATTACH_BYTES_TOTAL) break;

    const fileName = meta.fileName || "File";
    const ext = extOf(fileName);

    if (LIMITS.DENY_EXT.has(ext)) {
      stats.failures.push({
        kind: "attachment",
        stage: "embedFileAttachments",
        token,
        fileName,
        error: "denied_extension",
      });
      continue;
    }

    try {
      if (meta.kind === "inline" && meta.contentBytes && meta.contentType) {
        const b64 = meta.contentBytes;
        const size = Math.floor((b64.length * 3) / 4);
        if (size <= 0) throw new Error("inline_empty");
        if (size > LIMITS.MAX_ATTACH_BYTES_EACH) throw new Error("inline_too_large_each");
        if (total + size > LIMITS.MAX_ATTACH_BYTES_TOTAL) throw new Error("inline_too_large_total");

        const dataUrl = `data:${meta.contentType};base64,${b64}`;
        tokenToDataUrl.set(token, { dataUrl, fileName, size });
        stats.attachEmbedded = (stats.attachEmbedded || 0) + 1;
        total += size;
        stats.attachBytes = total;
        continue;
      }

      if (meta.kind === "shareLink" && meta.url) {
        const driveItem = await getDriveItemFromSharingUrl(meta.url);
        const driveId = driveItem?.parentReference?.driveId;
        const itemId = driveItem?.id;
        if (!driveId || !itemId) throw new Error("driveItem_missing_ids");

        const blob = await downloadDriveItemContent(driveId, itemId);
        const size = blob.size || 0;
        if (size <= 0) throw new Error("blob_empty");
        if (size > LIMITS.MAX_ATTACH_BYTES_EACH) throw new Error("too_large_each");
        if (total + size > LIMITS.MAX_ATTACH_BYTES_TOTAL) throw new Error("too_large_total");

        const dataUrl = await blobToDataUrl(blob);
        if (!dataUrl.startsWith("data:")) throw new Error("dataurl_invalid");

        tokenToDataUrl.set(token, { dataUrl, fileName, size });
        stats.attachEmbedded = (stats.attachEmbedded || 0) + 1;
        total += size;
        stats.attachBytes = total;
        continue;
      }

      throw new Error("unknown_attachment_kind");
    } catch (e) {
      stats.failures.push({
        kind: "attachment",
        stage: "embedFileAttachments",
        token,
        fileName,
        url: meta.url || "",
        error: normalizeErrorMessage(e),
      });
    }
  }

  return tokenToDataUrl;
}

function replaceFileTokensOffline(html, tokenToDataUrl, fileTokenToMeta) {
  let out = html || "";

  out = out.replaceAll(
    /<div\b[^>]*\bdata-file-token=(["'])(FILE_\d+)\1[^>]*>[\s\S]*?<\/div>/gi,
    (full, q, token) => {
      const meta = fileTokenToMeta.get(token) || {};
      const fileName = meta.fileName || "Файл";
      const online = meta.url || "";
      const got = tokenToDataUrl.get(token);

      if (got?.dataUrl) {
        const sizeMb = got.size ? (got.size / 1024 / 1024).toFixed(2) : "";
        return `
          <div class="attachmentCard embedded">
            <div class="attIcon">📎</div>
            <div class="attName" title="${esc(fileName)}">${esc(fileName)}</div>
            <a class="attDownload" download="${esc(fileName)}" href="${esc(got.dataUrl)}">Свали офлайн</a>
            <div class="attMeta">${sizeMb ? `${esc(sizeMb)} MB` : ""}</div>
          </div>
        `;
      }

      if (online) {
        return `
          <div class="attachmentCard missingLocal">
            <div class="attIcon">📎</div>
            <div class="attName" title="${esc(fileName)}">${esc(fileName)}</div>
            <a class="attDownload" href="${esc(online)}" target="_blank" rel="noreferrer">Свали онлайн</a>
            <div class="attMissing">не е вграден</div>
          </div>
        `;
      }

      return `
        <div class="attachmentCard missingLocal">
          <div class="attIcon">📎</div>
          <div class="attName" title="${esc(fileName)}">${esc(fileName)}</div>
          <div class="attMissing">липсва линк</div>
        </div>
      `;
    }
  );

  return out;
}

/** =========================
 * EXPORT HTML BUILDER
 * ========================= */
function buildHtml(items, stats, exportTitle, archiveRange) {
  const generatedAt = new Date().toLocaleString();

  const getInitials = (name) => {
    const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
    const a = parts[0]?.[0] || "?";
    const b = parts[1]?.[0] || "";
    return (a + b).toUpperCase();
  };

  let htmlMsgs = "";
  let lastSender = null;

  for (const m of items) {
    const sameSender = m.sender === lastSender;
    lastSender = m.sender;

    htmlMsgs += `
      <div class="msgRow ${m.isMe ? "me" : "other"}" id="mid-${esc(m.mid)}" data-mid="${esc(m.mid)}">
        <div class="avatar ${sameSender ? "ghost" : ""}">${esc(getInitials(m.sender))}</div>
        <div class="bubbleWrap">
          <div class="msgHead ${sameSender ? "compact" : ""}">
            <div class="senderLine">
              ${sameSender ? "" : `<span class="sender">${esc(m.sender)}</span>`}
            </div>
            <div class="time">${esc(m.dtStr)}</div>
          </div>
          <div class="bubble">
            <div class="body">
              ${m.forwardHtml || ""}
              ${m.quoteHtml || ""}
              ${m.bodyHtml || ""}
              ${m.reactionsHtml || ""}
            </div>
          </div>
        </div>
      </div>
    `;
  }

  return `<!doctype html>
<html lang="bg">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>${esc(exportTitle || "Teams Chat Archive (OFFLINE)")}</title>
<style>
  :root{
    --bg:#0b1220;
    --panel:rgba(15,23,42,.55);
    --line:rgba(255,255,255,.10);
    --text:#e5e7eb;
    --muted:#94a3b8;
    --bubble:rgba(255,255,255,.06);
    --bubble2:rgba(255,255,255,.04);
  }
  *{box-sizing:border-box}
  body{
    margin:0;
    font-family: ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, Arial;
    color:var(--text);
    background:
      radial-gradient(1200px 800px at 20% -10%, rgba(124,58,237,.35), transparent 60%),
      radial-gradient(1000px 700px at 90% 0%, rgba(37,99,235,.25), transparent 60%),
      var(--bg);
    height:100vh;
    overflow:hidden;
  }
  .topbar{
    padding:12px 16px;
    border-bottom:1px solid var(--line);
    background: var(--panel);
    backdrop-filter: blur(10px);
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:12px;
  }
  .title{ font-weight:1000; font-size:15px; line-height:1.2; }
  .meta{
    color:var(--muted);
    font-size:12px;
    margin-top:4px;
    display:flex;
    gap:10px;
    flex-wrap:wrap;
  }
  .pill{
    display:inline-flex;
    align-items:center;
    gap:6px;
    padding:4px 10px;
    border-radius:999px;
    border:1px solid var(--line);
    background: rgba(0,0,0,.22);
    color:var(--muted);
    font-weight:900;
    font-size:12px;
    user-select:none;
  }
  .actions{ display:flex; gap:10px; align-items:center; }
  .btn{
    padding:8px 10px;
    border-radius:14px;
    border:1px solid var(--line);
    background: rgba(255,255,255,.03);
    color:var(--text);
    font-weight:1000;
    cursor:pointer;
  }
  .btn:hover{ background: rgba(255,255,255,.06); }

  .chatScroller{
    height: calc(100vh - 70px);
    overflow:auto;
    padding:14px 18px 18px 18px;
  }
  .chatInner{
    max-width:980px;
    margin:0 auto;
    display:flex;
    flex-direction:column;
    gap:10px;
  }

  .msgRow{ display:flex; gap:10px; align-items:flex-start; }
  .avatar{
    width:34px;height:34px;border-radius:999px;
    display:grid;place-items:center;
    font-weight:1000;font-size:12px;
    background: rgba(255,255,255,.08);
    border:1px solid var(--line);
    user-select:none;
    flex:0 0 auto;
  }
  .avatar.ghost{ visibility:hidden; }
  .bubbleWrap{ flex:1; min-width:0; }
  .msgHead{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:12px;
    padding:0 4px;
  }
  .msgHead.compact{ opacity:.85; }
  .sender{ font-weight:1000; font-size:13px; }
  .time{ color:var(--muted); font-size:11px; white-space:nowrap; }

  .bubble{
    margin-top:6px;
    padding:10px 12px;
    border-radius:14px;
    background: linear-gradient(180deg, var(--bubble), var(--bubble2));
    border:1px solid var(--line);
    box-shadow:0 18px 40px rgba(0,0,0,.22);
    overflow:hidden;
  }
  .body{ line-height:1.35; overflow-wrap:anywhere; }
  .body a{ color:#93c5fd; }
  .body a:hover{ text-decoration:underline; }
  .body img{
    display:block;
    max-width:min(520px, 100%);
    max-height:420px;
    width:auto;
    height:auto;
    object-fit:contain;
    border-radius:12px;
    border:1px solid var(--line);
    margin:8px 0;
  }

  .attachmentCard{
    display:flex; align-items:center; gap:12px;
    padding:10px 12px; border-radius:14px; margin-top:10px;
    border:1px solid var(--line); background: rgba(255,255,255,.04);
    max-width:780px;
  }
  .attachmentCard.embedded{ background: rgba(37,99,235,.10); border-color: rgba(37,99,235,.25); }
  .attachmentCard.missingLocal{ background: rgba(239,68,68,.08); border-color: rgba(239,68,68,.20); }
  .attIcon{font-size:18px; flex-shrink:0;}
  .attName{
    flex:1; min-width:0; font-weight:1000; font-size:13px;
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
  }
  .attDownload{
    font-size:12px; font-weight:1000; text-decoration:none; color:var(--text);
    padding:7px 10px; border-radius:12px; border:1px solid var(--line);
    background: rgba(0,0,0,.25); white-space:nowrap;
  }
  .attDownload:hover{ background: rgba(255,255,255,.06); }
  .attMissing{ font-size:12px; color:#fecaca; font-weight:1000; white-space:nowrap; }
  .attMeta{ font-size:11px; color:var(--muted); font-weight:900; white-space:nowrap; }

  .reactions{ display:flex; gap:6px; flex-wrap:wrap; margin-top:10px; }
  .reactChip{
    display:inline-flex; align-items:center; gap:6px;
    padding:4px 8px; border-radius:999px;
    border:1px solid var(--line); background: rgba(0,0,0,.25);
    font-weight:1000; font-size:12px; user-select:none;
    cursor:default;
  }
  .reactChip .e{ font-size:14px; line-height:1; }
  .reactChip .c{ color: var(--muted); font-weight:1000; }

  .quotePreview{
    margin:0 0 10px 0;
    padding:10px 12px;
    border-radius:12px;
    border:1px solid rgba(255,255,255,.14);
    background: rgba(0,0,0,.22);
  }
  .quoteBox{
    border-left:3px solid rgba(147,197,253,.55);
    padding-left:10px;
    margin-top:8px;
    color:rgba(229,231,235,.92);
    font-size:12.5px;
  }
  .quoteHead{ display:flex; gap:8px; align-items:baseline; flex-wrap:wrap; opacity:.95; }
  .qAuthor{ font-weight:1000; font-size:12px; color:rgba(229,231,235,.92); }
  .qTs{ font-weight:900; font-size:11px; color:rgba(148,163,184,.95); }
  .qJump{ color:#93c5fd; font-size:12px; font-weight:1000; text-decoration:none; }
  .qJump:hover{ text-decoration:underline; }
  .qText{ margin-top:6px; font-size:12.5px; color:rgba(229,231,235,.92); }

  .fwdWrap{ margin:0 0 10px 0; }
  .fwdHeader{
    display:flex;
    gap:10px;
    align-items:flex-start;
    flex-wrap:wrap;
  }
  .fwdChip{
    display:inline-flex;
    align-items:center;
    padding:4px 10px;
    border-radius:999px;
    border:1px solid rgba(255,255,255,.14);
    background: rgba(124,58,237,.12);
    color: rgba(229,231,235,.96);
    font-weight:1000;
    font-size:12px;
  }
  .fwdMeta{
    color: var(--muted);
    font-size:12px;
    font-weight:900;
    opacity:.95;
    padding-top:3px;
  }
  .fwdBody{ margin-top:8px; }
  .fwdInner{
    padding:12px 14px;
    border-radius:12px;
    border:1px solid rgba(255,255,255,.14);
    background: rgba(0,0,0,.22);
    line-height:1.45;
  }
  .fwdInner p{ margin:0 0 8px 0; }
  .fwdInner p:last-child{ margin-bottom:0; }
  .fwdInner a{ color:#93c5fd; }
  .fwdInner .fwdEmpty{
    color: var(--muted);
    font-size:12px;
    font-style:italic;
  }

  .msgRow.me{ flex-direction:row-reverse; }
  .msgRow.me .msgHead{ flex-direction:row-reverse; }

  .reactPopover{
    position:fixed;
    z-index:999999;
    min-width:180px;
    max-width:420px;
    background: rgba(15,23,42,.96);
    color: var(--text);
    border:1px solid rgba(255,255,255,.10);
    border-radius:12px;
    box-shadow:0 18px 50px rgba(0,0,0,.45);
    padding:10px 12px;
    display:none;
  }
  .reactPopover .tTitle{ font-weight:1000; margin-bottom:8px; opacity:.95; }
  .reactPopover .tItem{
    display:flex; gap:10px; align-items:center;
    padding:6px 0;
    border-top:1px solid rgba(255,255,255,.08);
  }
  .reactPopover .tItem:first-of-type{ border-top:none; }
  .reactPopover .pAvatar{
    width:26px;height:26px;border-radius:999px;
    display:grid;place-items:center;
    background: rgba(255,255,255,.08);
    border:1px solid rgba(255,255,255,.10);
    font-weight:1000;
    font-size:11px;
  }
  .reactPopover .pName{ font-weight:1000; }
</style>
</head>
<body>
  <div class="topbar">
    <div>
      <div class="title">${esc(exportTitle || "Teams Chat Archive (OFFLINE)")}</div>
      <div class="meta">
        <span class="pill">Елементи: ${items.length}</span>
        <span class="pill">Период: ${esc(archiveRange?.label || "—")}</span>
        <span class="pill">Генериран: ${esc(generatedAt)}</span>
        <span class="pill">Снимки: ${stats.imagesEmbedded || 0} embedded • ${((stats.imagesBytes || 0) / 1024 / 1024).toFixed(1)} MB</span>
        <span class="pill">Файлове: ${stats.attachEmbedded || 0} embedded • ${((stats.attachBytes || 0) / 1024 / 1024).toFixed(1)} MB</span>
      </div>
    </div>
    <div class="actions">
      <button class="btn" id="btnBottom">⬇ Най-нови</button>
    </div>
  </div>

  <div class="chatScroller" id="chatScroller">
    <div class="chatInner" id="chatInner">${htmlMsgs}</div>
  </div>

  <div class="reactPopover" id="reactPopover" role="dialog" aria-modal="false">
    <div class="tTitle" id="rpTitle">Реакции</div>
    <div id="rpList"></div>
  </div>

<script>
(function(){
  const sc = document.getElementById("chatScroller");
  const btn = document.getElementById("btnBottom");
  const pop = document.getElementById("reactPopover");
  const list = document.getElementById("rpList");
  const title = document.getElementById("rpTitle");

  if (sc && btn) {
    const toBottom = () => { sc.scrollTop = sc.scrollHeight; };
    requestAnimationFrame(() => requestAnimationFrame(toBottom));
    btn.addEventListener("click", toBottom);
  }

  if (!pop || !list || !title) return;

  let activeChip = null;
  let hideTimer = null;

  const initials = (name) => {
    const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
    const a = parts[0]?.[0] || "?";
    const b = parts[1]?.[0] || "";
    return (a + b).toUpperCase();
  };

  const parseNames = (chip) => {
    const raw = (chip.getAttribute("data-names") || "").trim();
    if (!raw) return [];
    return raw.split("|").map(s => s.trim()).filter(Boolean);
  };

  const getEmoji = (chip) => {
    return chip.getAttribute("data-emoji") ||
      chip.querySelector(".e")?.textContent?.trim() ||
      "👍";
  };

  const getCount = (chip) => {
    return chip.getAttribute("data-count") ||
      chip.querySelector(".c")?.textContent?.trim() ||
      "0";
  };

  const position = (chip) => {
    const rect = chip.getBoundingClientRect();
    pop.style.display = "block";

    const w = pop.offsetWidth;
    const h = pop.offsetHeight;

    let left = window.scrollX + rect.left;
    let top = window.scrollY + rect.bottom + 10;

    if (left + w > window.scrollX + window.innerWidth - 12) {
      left = window.scrollX + window.innerWidth - w - 12;
    }
    if (left < window.scrollX + 12) {
      left = window.scrollX + 12;
    }

    if (top + h > window.scrollY + window.innerHeight - 12) {
      top = window.scrollY + rect.top - h - 10;
    }

    if (top < window.scrollY + 12) {
      top = window.scrollY + 12;
    }

    pop.style.left = left + "px";
    pop.style.top = top + "px";
  };

  const renderList = (chip) => {
    const emoji = getEmoji(chip);
    const count = getCount(chip);
    const names = parseNames(chip);

    const reactionType = chip.getAttribute("data-reaction-type") || emoji;
title.textContent = emoji + " " + count + " reaction" + (count === "1" ? "" : "s") + " • " + reactionType;
    list.innerHTML = "";

    if (!names.length) {
      const row = document.createElement("div");
      row.className = "tItem";

      const av = document.createElement("div");
      av.className = "pAvatar";
      av.textContent = emoji;

      const nm = document.createElement("div");
      nm.className = "pName";
      nm.textContent = "No participant names available";

      row.appendChild(av);
      row.appendChild(nm);
      list.appendChild(row);
      return;
    }

    for (const n of names.slice(0, 100)) {
      const row = document.createElement("div");
      row.className = "tItem";

      const av = document.createElement("div");
      av.className = "pAvatar";
      av.textContent = initials(n);

      const nm = document.createElement("div");
      nm.className = "pName";
      nm.textContent = n;

      row.appendChild(av);
      row.appendChild(nm);
      list.appendChild(row);
    }
  };

  const show = (chip) => {
    clearTimeout(hideTimer);
    activeChip = chip;
    renderList(chip);
    position(chip);
  };

  const hide = () => {
    clearTimeout(hideTimer);
    hideTimer = setTimeout(() => {
      pop.style.display = "none";
      activeChip = null;
    }, 140);
  };

  document.addEventListener("mouseover", (e) => {
    const chip = e.target?.closest?.(".reactChip");
    if (!chip) return;
    show(chip);
  }, true);

  document.addEventListener("mouseout", (e) => {
    const chip = e.target?.closest?.(".reactChip");
    if (!chip) return;
    hide();
  }, true);

  pop.addEventListener("mouseenter", () => {
    clearTimeout(hideTimer);
  });

  pop.addEventListener("mouseleave", () => {
    hide();
  });

  document.addEventListener("scroll", () => {
    if (activeChip) position(activeChip);
  }, true);

  window.addEventListener("resize", () => {
    if (activeChip) position(activeChip);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      pop.style.display = "none";
      activeChip = null;
    }
  });
})();
</script>
</body>
</html>`;
}

/** =========================
 * FILE DOWNLOAD
 * ========================= */
function downloadTextFile(fileName, text, mime = "text/html;charset=utf-8") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 8000);
}

/** =========================
 * EXPORT PIPELINE
 * ========================= */
async function exportChatToOfflineHtml(chat, me, usedFileNames = null) {
  let members = Array.isArray(chat?.members) ? [...chat.members] : [];

  try {
    const fullMembers = await loadChatMembersCached(chat.id);
    members = mergeMembers(fullMembers, members);
  } catch {
    members = mergeMembers(members, []);
  }

  const exportTitle = chatDisplayName(chat, members);
  const fileBase = fileBaseForChat(chat, me, members);
  const initialName = `${safeFileName(fileBase)}.html`;
  const outName = usedFileNames ? makeUniqueFileName(initialName, usedFileNames) : initialName;

  let messages;
  try {
    messages = await loadAllChatMessages(chat.id);
  } catch (e) {
    const err = new Error(`load_messages_failed: ${normalizeErrorMessage(e)}`);
    err.stage = "load_messages";
    throw err;
  }

  messages.sort((a, b) => {
    const at = Date.parse(a.createdDateTime || "") || 0;
    const bt = Date.parse(b.createdDateTime || "") || 0;
    if (at !== bt) return at - bt;
    return String(a.id || "").localeCompare(String(b.id || ""));
  });

  const stats = {
    imagesEmbedded: 0,
    imagesBytes: 0,
    attachEmbedded: 0,
    attachBytes: 0,
    failures: [],
  };

  const fileTokenToMeta = new Map();
  const dtos = [];

  for (const msg of messages) {
    const created = new Date(msg.createdDateTime);
    const dtMs = Date.parse(msg.createdDateTime) || 0;
    const dtStr = isNaN(created.getTime()) ? "(няма дата/час)" : formatDDMMYYYY_HHMMSS(created);

    const sender = extractSender(msg) || "Unknown";
    let bodyHtml = extractBodyHtml(msg);

    if (shouldAttemptHosted(bodyHtml)) {
      let hosted = [];
      try {
        hosted = await listHostedContents(chat.id, msg.id);
      } catch (e) {
        stats.failures.push({
          kind: "hostedImageList",
          stage: "listHostedContents",
          messageId: msg.id,
          error: normalizeErrorMessage(e),
        });
        hosted = [];
      }

      const tok = tokenizeHostedImages(bodyHtml, hosted);
      bodyHtml = tok.html;
      bodyHtml = await embedHostedImages(chat.id, msg.id, bodyHtml, hosted, stats);
    }

    let forwardedAttachments = extractForwardedAttachmentData(msg.attachments || [], members, me);
    forwardedAttachments = await enrichForwardedAuthors(forwardedAttachments, members, me);

    const { quotes: quotesFromAttachments, fileAttachments } =
      splitQuotesFromAttachments(msg.attachments || []);

    bodyHtml = tokenizeFileAttachments(bodyHtml, fileAttachments, fileTokenToMeta);

    dtos.push({
      kind: "message",
      mid: msg.id,
      sender,
      dtStr,
      dtMs,
      isMe: isFromMe(msg, me?.id),
      bodyHtml,
      replyToId: msg.replyToId || "",
      reactionsHtml: buildReactionsFromGraph(msg, members, me),
      forwardHtml: "",
      quoteHtml: "",
      _quotesFromAttachments: quotesFromAttachments,
      _forwardedAttachments: forwardedAttachments,
      _rawBodyHtml: bodyHtml,
    });
  }

  const byId = new Map(dtos.filter((d) => d.kind === "message").map((d) => [d.mid, d]));

  for (const dto of dtos) {
    const quotes = [];

    const reply = buildReplyPreview(dto, byId);
    if (reply) quotes.push(reply);

    for (const q of dto._quotesFromAttachments || []) {
      if (q?.refId) {
        const t = byId.get(String(q.refId));
        if (t) {
          q.author =
            q.author && q.author !== "Unknown" && q.author !== "Microsoft Teams"
              ? q.author
              : t.sender || "Unknown";
          q.ts = q.ts || t.dtStr || "";
          if (!q.previewText) q.previewText = stripToPreviewText(t.bodyHtml || "", 220);
        }
      }
      quotes.push(q);
    }

    for (const q of extractBlockquoteQuotes(dto._rawBodyHtml)) {
      quotes.push(q);
    }

    dto.quoteHtml = renderQuoteStack(quotes);

    const forwarded =
      (dto._forwardedAttachments && dto._forwardedAttachments[0]) ||
      extractForwardFromBody(dto._rawBodyHtml);

    dto.forwardHtml = renderForwardBlock(forwarded);
  }

  const tokenToDataUrl = await embedFileAttachments(fileTokenToMeta, stats);
  for (const dto of dtos) {
    dto.bodyHtml = replaceFileTokensOffline(dto.bodyHtml, tokenToDataUrl, fileTokenToMeta);
  }

  const archiveRange = getArchiveRange(dtos);
  const html = buildHtml(dtos, stats, exportTitle, archiveRange);
  downloadTextFile(outName, html);

  return { outName, stats, items: dtos.length, archiveRange };
}

/** =========================
 * QUEUE HELPERS
 * ========================= */
function getQueueStatusLabel(status) {
  if (status === "queued") return "Queued";
  if (status === "running") return "Running";
  if (status === "done") return "Done";
  if (status === "failed") return "Failed";
  return status || "";
}

function getQueueItemByChatId(chatId) {
  return _queueItems.find((x) => x.chatId === chatId) || null;
}

function clearQueueErrorState(item) {
  item.error = null;
  item.errorCode = null;
  item.errorStage = null;
  item.errorHuman = null;
}

function buildQueueFromSelection() {
  const selectedChats = _chats.filter((c) => _selectedChatIds.has(c.id));
  const existingIds = new Set(_queueItems.map((x) => x.chatId));

  let added = 0;
  for (const chat of selectedChats) {
    if (existingIds.has(chat.id)) continue;
    _queueItems.push({
      chatId: chat.id,
      title: getChatTitle(chat),
      status: "queued",
      result: null,
      error: null,
      errorCode: null,
      errorStage: null,
      errorHuman: null,
      startedAt: null,
      finishedAt: null,
    });
    existingIds.add(chat.id);
    added += 1;
  }

  return { added, totalSelected: selectedChats.length };
}

function buildQueueFromAllChats() {
  _selectedChatIds = new Set(_chats.map((c) => c.id));
  return buildQueueFromSelection();
}

function requeueFailedOnly() {
  let changed = 0;
  for (const item of _queueItems) {
    if (item.status !== "failed") continue;
    item.status = "queued";
    item.result = null;
    item.startedAt = null;
    item.finishedAt = null;
    clearQueueErrorState(item);
    changed += 1;
  }
  return changed;
}

async function runExportQueue() {
  if (_queueRunning) return;
  if (!_me) {
    log("Please login first.");
    return;
  }

  const queued = _queueItems.filter((x) => x.status === "queued");
  if (!queued.length) {
    log("No queued chats to export.");
    return;
  }

  _queueRunning = true;
  renderChats();
  setCounts();
  setBusy(true);

  appendLogLine(`Queue start: ${queued.length} chat(s).`);
  appendLogLine("Note: browser may ask to allow multiple automatic downloads for this site.");

  const usedFileNames = new Set(
    _queueItems
      .filter((x) => x.status === "done" && x.result?.outName)
      .map((x) => String(x.result.outName).toLowerCase())
  );

  try {
    for (const item of _queueItems) {
      if (item.status !== "queued") continue;

      const chat = _chats.find((c) => c.id === item.chatId);
      if (!chat) {
        item.status = "failed";
        item.errorStage = "queue_resolution";
        item.error = "chat_not_found";
        const cls = classifyFailure(item.error);
        item.errorCode = cls.code;
        item.errorHuman = cls.human;
        item.finishedAt = new Date().toISOString();

        renderChats();
        setCounts();
        appendLogLine(`FAILED: ${item.title} -> ${formatFailureForLog(item)}`);
        continue;
      }

      item.status = "running";
      item.startedAt = new Date().toISOString();
      item.result = null;
      item.finishedAt = null;
      clearQueueErrorState(item);

      renderChats();
      setCounts();
      setBusy(true);
      appendLogLine(`RUNNING: ${item.title}`);

      try {
        const res = await exportChatToOfflineHtml(chat, _me, usedFileNames);
        item.status = "done";
        item.result = res;
        item.finishedAt = new Date().toISOString();

        appendLogLine(
          `DONE: ${item.title} -> ${res.outName} (${res.items} items, period ${res.archiveRange?.label || "—"})`
        );
        setBusy(true);
      } catch (e) {
        item.status = "failed";
        item.finishedAt = new Date().toISOString();

        item.errorStage = e?.stage || "export_chat";
        item.error = normalizeErrorMessage(e);

        const cls = classifyFailure(item.error);
        item.errorCode = cls.code;
        item.errorHuman = cls.human;

        appendLogLine(`FAILED: ${item.title} -> ${formatFailureForLog(item)}`);
        setBusy(true);
      }

      renderChats();
      setCounts();
      await sleep(180);
    }

    const qc = getQueueSummaryCounts();
    appendLogLine(`Queue finished. Done: ${qc.done}, Failed: ${qc.failed}, Total: ${_queueItems.length}`);
  } finally {
    _queueRunning = false;
    renderChats();
    setCounts();
    setBusy(false);
    armSessionWatcher();
  }
}

/** =========================
 * CHAT LIST RENDERING
 * ========================= */
function renderChats() {
  if (!ui.chatList) return;

  const q = (ui.q?.value || "").trim().toLowerCase();
  ui.chatList.innerHTML = "";

  const list = _chats
    .map((c) => ({ chat: c, title: getChatTitle(c) }))
    .filter((x) => !q || x.title.toLowerCase().includes(q));

  for (const item of list) {
    const c = item.chat;
    const queueItem = getQueueItemByChatId(c.id);

    const row = document.createElement("div");
    row.className = "chatItem";
    row.style.cursor = _queueRunning ? "default" : "pointer";
    row.style.background =
      queueItem?.status === "running"
        ? "rgba(37,99,235,.10)"
        : queueItem?.status === "done"
        ? "rgba(34,197,94,.08)"
        : queueItem?.status === "failed"
        ? "rgba(239,68,68,.08)"
        : "rgba(255,255,255,.03)";

    const left = document.createElement("div");
    left.className = "chatItemLeft";

    const checkWrap = document.createElement("div");
    checkWrap.className = "chatCheckWrap";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "chatCheckbox";
    cb.checked = _selectedChatIds.has(c.id);
    cb.disabled = _queueRunning;

    cb.addEventListener("click", (ev) => ev.stopPropagation());
    cb.addEventListener("change", () => {
      if (cb.checked) _selectedChatIds.add(c.id);
      else _selectedChatIds.delete(c.id);

      setCounts();
      renderChats();
    });

    checkWrap.appendChild(cb);

    const main = document.createElement("div");
    main.className = "chatMain";

    const title = document.createElement("div");
    title.className = "chatTitle";
    title.textContent =
      `${String(c.chatType).toLowerCase() === "group" ? "👥 " : "💬 "}${item.title}`;

    const meta = document.createElement("div");
    meta.className = "chatMeta";
    const started = getChatStartedLabel(c);
    const lastWritten = getChatLastMessageLabel(c);
    meta.textContent =
      `Started: ${started} • Last message: ${lastWritten}${queueItem ? ` • Queue: ${getQueueStatusLabel(queueItem.status)}` : ""}`;

    main.appendChild(title);
    main.appendChild(meta);

    if (queueItem?.status === "failed") {
      const err1 = document.createElement("div");
      err1.className = "chatError";
      err1.textContent = `Reason: ${queueItem.errorHuman || queueItem.error || "Unknown error"}`;

      const err2 = document.createElement("div");
      err2.className = "chatErrorMeta";
      err2.textContent = `Stage: ${queueItem.errorStage || "unknown"} • Code: ${queueItem.errorCode || "unknown"}`;

      main.appendChild(err1);
      main.appendChild(err2);
    }

    left.appendChild(checkWrap);
    left.appendChild(main);

    const right = document.createElement("div");
    right.className = "chatRight";

    if (queueItem) {
      const badge = document.createElement("span");
      badge.className = "statusBadge";
      badge.textContent = getQueueStatusLabel(queueItem.status);
      badge.style.background =
        queueItem.status === "running"
          ? "rgba(37,99,235,.18)"
          : queueItem.status === "done"
          ? "rgba(34,197,94,.18)"
          : queueItem.status === "failed"
          ? "rgba(239,68,68,.18)"
          : "rgba(255,255,255,.06)";
      right.appendChild(badge);
    }

    row.appendChild(left);
    row.appendChild(right);

    row.addEventListener("click", () => {
      if (_queueRunning) return;

      if (_selectedChatIds.has(c.id)) _selectedChatIds.delete(c.id);
      else _selectedChatIds.add(c.id);

      setCounts();
      renderChats();
    });

    ui.chatList.appendChild(row);
  }
}

/** =========================
 * UI ACTIONS
 * ========================= */
async function onExportSelectedClick() {
  if (!_me) {
    log("Please login first.");
    return;
  }

  if (_selectedChatIds.size === 0) {
    log("Please select at least one chat first.");
    return;
  }

  markExportStartedForSession();

  const info = buildQueueFromSelection();
  renderChats();
  setCounts();

  appendLogLine(`Queue prepared: added ${info.added} / selected ${info.totalSelected}.`);
  await runExportQueue();
}

async function onQueueAllClick() {
  if (!_me) {
    log("Please login first.");
    return;
  }

  if (_chats.length === 0) {
    log("Load chats first.");
    return;
  }

  markExportStartedForSession();

  const info = buildQueueFromAllChats();
  renderChats();
  setCounts();

  appendLogLine(`Queue ALL prepared: added ${info.added} / total chats ${_chats.length}.`);
  await runExportQueue();
}

async function onRetryFailedOnlyClick() {
  if (!_me) {
    log("Please login first.");
    return;
  }

  const failedCount = getQueueSummaryCounts().failed;
  if (!failedCount) {
    log("There are no failed chats to retry.");
    return;
  }

  markExportStartedForSession();

  const changed = requeueFailedOnly();
  renderChats();
  setCounts();

  appendLogLine(`Retry FAILED only: re-queued ${changed} chat(s).`);
  await runExportQueue();
}

/** =========================
 * UI EVENTS
 * ========================= */
function bindUiEvents() {
  if (_uiEventsBound) return;
  _uiEventsBound = true;

  ui.btnLogin?.addEventListener("click", async () => {
    try {
      setBusy(true);
      appendLogLine("Login button clicked.");
      await login();
    } catch (e) {
      console.error(e);
      const msg = normalizeErrorMessage(e);
      appendLogLine(`Login error: ${msg}`);
      alert(`Login error:\n${msg}`);
    } finally {
      setBusy(false);
      setCounts();
    }
  });

  ui.btnChats?.addEventListener("click", async () => {
    try {
      if (!_me) {
        log("Please login first.");
        return;
      }

      setBusy(true);
      log("Loading chats…");

      _chats = await loadMyChats();

      const validIds = new Set(_chats.map((c) => c.id));
      _selectedChatIds = new Set([..._selectedChatIds].filter((id) => validIds.has(id)));
      _queueItems = _queueItems.filter((x) => validIds.has(x.chatId));

      renderChats();
      setCounts();

      log({ chats: _chats.length, sample: _chats.slice(0, 3) });
      appendLogLine(`Chats loaded successfully: ${_chats.length}`);
    } catch (e) {
      console.error(e);
      log(`Load chats failed:\n${normalizeErrorMessage(e)}`);
    } finally {
      setBusy(false);
      setCounts();
    }
  });

  ui.q?.addEventListener("input", () => renderChats());

  ui.btnExport?.addEventListener("click", async () => {
    try {
      await onExportSelectedClick();
    } catch (e) {
      console.error(e);
      log(`Export queue failed:\n${normalizeErrorMessage(e)}`);
    } finally {
      setBusy(false);
      setCounts();
    }
  });

  document.addEventListener("click", async (ev) => {
    if (ev.target?.id === "btnQueueAll") {
      try {
        await onQueueAllClick();
      } catch (e) {
        console.error(e);
        log(`Queue all failed:\n${normalizeErrorMessage(e)}`);
      } finally {
        setBusy(false);
        setCounts();
      }
      return;
    }

    if (ev.target?.id === "btnRetryFailed") {
      try {
        await onRetryFailedOnlyClick();
      } catch (e) {
        console.error(e);
        log(`Retry failed only error:\n${normalizeErrorMessage(e)}`);
      } finally {
        setBusy(false);
        setCounts();
      }
    }
  });
}

/** =========================
 * INIT
 * ========================= */
async function doInit() {
  if (!assertUiIds()) return;

  ensureExtraButtons();
  refreshUiRefs();
  bindUiEvents();
  showLoginView();

  try {
    setBusy(true);
    log("Initializing auth…");

    const acc = await initAuth();

    if (acc) {
      log("Signed in. Loading /me…");
      const data = await loadMe();
      _me = {
        id: data.id,
        displayName: data.displayName,
        userPrincipalName: data.userPrincipalName,
      };

      if (!_loginAtMs) recordAuthenticatedSession({ reset: true });
      else recordAuthenticatedSession({ reset: false });

      showAppView();
      appendLogLine(`Ready. Signed in as ${_me.displayName || _me.userPrincipalName || _me.id}`);
    } else {
      _me = null;
      clearAppSessionMeta();
      showLoginView();
      log("Not signed in. Use the Microsoft sign-in button.");
    }

    setCounts();
  } catch (e) {
    console.error(e);
    showLoginView();
    log(`Init failed:\n${normalizeErrorMessage(e)}`);
    alert(`Init failed:\n${normalizeErrorMessage(e)}`);
  } finally {
    setBusy(false);
    setCounts();
  }
}

doInit();