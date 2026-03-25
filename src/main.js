// src/main.js — Teams Chat Export Portal
// Clean production-focused build: login + chats + queue export + offline HTML + forwarded resolution

import * as msal from "@azure/msal-browser";
import {
  t,
  tFor,
  getCurrentLanguage,
  setCurrentLanguage,
  getSupportedLanguages,
  applyI18nToDocument,
} from "./i18n";

/** =========================
 * CONFIG
 * ========================= */
const TENANT_ID = String(import.meta.env.VITE_TENANT_ID || "").trim();
const CLIENT_ID = String(import.meta.env.VITE_CLIENT_ID || "").trim();
const REDIRECT_URI = window.location.origin + "/";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function assertRequiredConfig() {
  const missing = [];

  if (!TENANT_ID) missing.push("VITE_TENANT_ID");
  if (!CLIENT_ID) missing.push("VITE_CLIENT_ID");

  if (!missing.length) return;

  const msg =
    `Missing required Vite environment variables: ${missing.join(", ")}.\n` +
    `Create/update the .env file in the project root and restart the Vite dev server.`;

  console.error("[CONFIG]", {
    TENANT_ID,
    CLIENT_ID,
    missing,
    redirectUri: REDIRECT_URI,
  });

  throw new Error(msg);
}

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

  // New file policy:
  // - embed everything automatically up to 600 MB per file
  // - no confirmation prompts during export
  MAX_ATTACH_BYTES_EACH: 600 * 1024 * 1024,
  MAX_ATTACH_BYTES_TOTAL: 20 * 1024 * 1024 * 1024,

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
  "Sites.Read.All",
];

/** =========================
 * MSAL
 * ========================= */
assertRequiredConfig();

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

  langSelect: null,
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

  ui.langSelect = $("langSelect");
}

function log(obj) {
  const text = typeof obj === "string" ? obj : JSON.stringify(obj, null, 2);
  console.log("[APP]", text);
}

function appendLogLine(line) {
  console.log("[APP]", String(line));
}

function applyLanguageToUi() {
  applyI18nToDocument(document);

  if (ui.langSelect) {
    ui.langSelect.value = getCurrentLanguage();
  }

  ensureExtraButtons();
  setCounts();
  renderChats();
}

function populateLanguageSelector() {
  if (!ui.langSelect) return;

  ui.langSelect.innerHTML = "";

  for (const lang of getSupportedLanguages()) {
    const opt = document.createElement("option");
    opt.value = lang.code;
    opt.textContent = lang.label;
    ui.langSelect.appendChild(opt);
  }

  ui.langSelect.value = getCurrentLanguage();
}

function bindLanguageSelector() {
  ui.langSelect?.addEventListener("change", () => {
    setCurrentLanguage(ui.langSelect.value);
    applyLanguageToUi();
  });
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
    "langSelect",
  ];

  const missing = must.filter((id) => !$(id));
  if (missing.length) {
    console.warn("[UI] missing ids:", missing);
    log(t("uiMissingElementIds", { ids: missing.join(", ") }));
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
    btn.className = "btn";
    host.appendChild(btn);
  }

  if (!ui.btnRetryFailed) {
    const btn = document.createElement("button");
    btn.id = "btnRetryFailed";
    btn.type = "button";
    btn.className = "btn btnDanger";
    host.appendChild(btn);
  }

  refreshUiRefs();

  if (ui.btnQueueAll) {
    ui.btnQueueAll.textContent = t("queueAll");
  }

  if (ui.btnRetryFailed) {
    ui.btnRetryFailed.textContent = t("retryFailedOnly");
  }
}

/** =========================
 * STATE
 * ========================= */
let _me = null;
let _chats = [];
let _selectedChatIds = new Set();

let _queueItems = [];
let _queueRunning = false;
let _lastExportSummary = null;

let _loginAtMs = 0;
let _exportStartedSinceLogin = false;
let _sessionWatchHandle = null;
let _sessionReauthInProgress = false;
let _authInited = false;
let _uiEventsBound = false;
let _authInitPromise = null;

let _exportDirectoryHandle = null;
let _currentBatchFolderHandle = null;
let _currentBatchFolderName = "";

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

function getMemberBestLabel(member) {
  const raw =
    String(
      member?.displayName ||
      member?.user?.displayName ||
      member?.email ||
      member?.user?.email ||
      member?.userPrincipalName ||
      member?.upn ||
      member?.user?.userPrincipalName ||
      member?.id ||
      member?.userId ||
      ""
    ).trim();

  if (!raw) return "";

  // If it is an email/UPN, keep only the local readable part when possible
  const emailMatch = raw.match(/^([^@]+)@/);
  if (emailMatch && emailMatch[1]) {
    const local = emailMatch[1]
      .replace(/[._-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (local) {
      return local
        .split(" ")
        .filter(Boolean)
        .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
        .join(" ");
    }
  }

  return raw;
}

function getMemberIdentityStrings(member) {
  return [
    member?.displayName,
    member?.user?.displayName,
    member?.email,
    member?.user?.email,
    member?.userPrincipalName,
    member?.upn,
    member?.user?.userPrincipalName,
    member?.id,
    member?.userId,
    member?.user?.id,
  ]
    .map((x) => String(x || "").trim())
    .filter(Boolean);
}

function isSamePersonLike(member, me) {
  const memberKeys = buildIdentityKeySet(getMemberIdentityStrings(member));
  const myKeys = buildIdentityKeySet([
    me?.id,
    me?.displayName,
    me?.userPrincipalName,
    me?.mail,
  ].filter(Boolean));

  if (!memberKeys.size || !myKeys.size) return false;

  for (const k of memberKeys) {
    if (myKeys.has(k)) return true;
  }

  return false;
}

function removeQueuedItemsThatNeverStarted() {
  _queueItems = _queueItems.filter((item) => item.status !== "queued");
}

function buildAttachmentFailureMap(stats) {
  const map = new Map();
  const list = Array.isArray(stats?.failures) ? stats.failures : [];

  for (const item of list) {
    if (!item?.token) continue;
    map.set(item.token, item);
  }

  return map;
}

function getDeletedNoticeAuthorLabel(sender) {
  const s = String(sender || "").trim();
  if (!s || s.toLowerCase() === "unknown") return "A participant";
  return s;
}

function getMessageMinuteKey(dtMs, dtStr = "") {
  if (Number.isFinite(dtMs) && dtMs > 0) {
    const d = new Date(dtMs);
    return [
      d.getFullYear(),
      pad2(d.getMonth() + 1),
      pad2(d.getDate()),
      pad2(d.getHours()),
      pad2(d.getMinutes()),
    ].join("-");
  }

  const raw = String(dtStr || "").trim();
  if (!raw) return "";

  const m = raw.match(/^(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2})/);
  if (!m) return raw;

  const [, dd, mm, yyyy, hh, mi] = m;
  return `${yyyy}-${mm}-${dd}-${hh}-${mi}`;
}

function getOtherParticipantLabelForOneOnOne(members, me) {
  const list = Array.isArray(members) ? members : [];
  if (!list.length) return "";

  const others = list.filter((m) => !isSamePersonLike(m, me));

  const candidates = (others.length ? others : list)
    .map((m) => getMemberBestLabel(m))
    .filter(Boolean);

  // Remove my display name as plain-text fallback as well
  const myName = String(me?.displayName || "").trim().toLowerCase();
  const filtered = candidates.filter((name) => !myName || name.toLowerCase() !== myName);

  return filtered[0] || candidates[0] || "";
}

function computeAttachmentsTotalSizeBytes(atts) {
  const list = Array.isArray(atts) ? atts : [];
  let total = 0;

  for (const a of list) {
    if (!a) continue;

    // 1. Direct size от Graph
    if (typeof a.size === "number" && a.size > 0) {
      total += a.size;
      continue;
    }

    // 2. Embedded contentBytes fallback
    if (typeof a.contentBytes === "string") {
      // base64 length → приблизителен размер
      const len = a.contentBytes.length;
      total += Math.floor((len * 3) / 4);
      continue;
    }
  }

  return total;
}

function isDeletedMessage(msg, bodyHtml = "") {
  if (msg?.deletedDateTime) return true;

  const mt = String(msg?.messageType || "").trim().toLowerCase();
  if (mt.includes("deleted")) return true;

  const text = extractPlainTextFromHtml(bodyHtml).toLowerCase();

  if (
    text === "this message has been deleted." ||
    text === "this message has been deleted" ||
    text === "message deleted" ||
    text === "съобщението е изтрито" ||
    text === "това съобщение е изтрито"
  ) {
    return true;
  }

  return false;
}

function formatBytesHuman(bytes) {
  const n = Number(bytes || 0);
  if (!Number.isFinite(n) || n <= 0) return "0 B";

  const gb = 1024 * 1024 * 1024;
  const mb = 1024 * 1024;
  const kb = 1024;

  if (n >= gb) return `${(n / gb).toFixed(2)} GB`;
  if (n >= mb) return `${(n / mb).toFixed(2)} MB`;
  if (n >= kb) return `${(n / kb).toFixed(2)} KB`;
  return `${n} B`;
}

function base64ToBlob(base64, mime = "application/octet-stream") {
  const clean = String(base64 || "").trim();
  if (!clean) throw new Error("base64_empty");

  const byteChars = atob(clean);
  const byteNumbers = new Array(byteChars.length);

  for (let i = 0; i < byteChars.length; i++) {
    byteNumbers[i] = byteChars.charCodeAt(i);
  }

  const byteArray = new Uint8Array(byteNumbers);
  return new Blob([byteArray], { type: mime });
}

function formatBytesToMB(bytes) {
  if (!bytes || bytes <= 0) return "";
  const mb = bytes / (1024 * 1024);
  return `${mb.toFixed(2)} MB`;
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

function formatFileTimestampYYYYMMDDHHMMSS(value = new Date()) {
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return "00000000000000";

  return [
    d.getFullYear(),
    pad2(d.getMonth() + 1),
    pad2(d.getDate()),
    pad2(d.getHours()),
    pad2(d.getMinutes()),
    pad2(d.getSeconds()),
  ].join("");
}

function trimFileBasePreserveWords(value, maxLen = 140) {
  const s = String(value || "").trim();
  if (!s) return "Chat";

  if (s.length <= maxLen) return s;

  let cut = s.slice(0, maxLen).trim();

  const lastSpace = cut.lastIndexOf(" ");
  if (lastSpace >= 24) {
    cut = cut.slice(0, lastSpace).trim();
  }

  return cut.replace(/[._\-\s]+$/g, "").trim() || "Chat";
}

function getParticipantInitials(displayName) {
  const parts = String(displayName || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  if (!parts.length) return "?";

  const first = parts[0]?.[0] || "";
  const last = parts.length > 1 ? (parts[parts.length - 1]?.[0] || "") : "";

  return (first + last).toUpperCase() || "?";
}

function getGroupChatFallbackTitleFromMembers(members) {
  const list = Array.isArray(members) ? members : [];

  const uniqueNames = [];
  const seen = new Set();

  for (const member of list) {
    const name = String(
      member?.displayName ||
      member?.user?.displayName ||
      member?.email ||
      member?.user?.email ||
      member?.userPrincipalName ||
      ""
    ).trim();

    if (!name) continue;

    const key = name.toLowerCase();
    if (seen.has(key)) continue;

    seen.add(key);
    uniqueNames.push(name);
  }

  if (!uniqueNames.length) return "Group Chat";

  const initials = uniqueNames.map(getParticipantInitials).filter(Boolean);
  return `Group Chat(${initials.join(", ")})`;
}

function isUnnamedGroupChat(chat) {
  const topic = String(chat?.topic || "").trim();
  return !topic || topic.toLowerCase() === "group chat";
}

function buildExportFileBase(chat, me, members) {
  const chatType = String(chat?.chatType || "").toLowerCase();
  const topic = String(chat?.topic || "").trim();

  if (chatType === "oneonone") {
    const otherLabel = getOtherParticipantLabelForOneOnOne(members, me);
    if (otherLabel) return otherLabel;

    if (topic) return topic;

    return "Former Participant";
  }

  if (chatType === "group" || chatType === "meeting") {
    if (!isUnnamedGroupChat(chat)) {
      return topic;
    }

    return getGroupChatFallbackTitleFromMembers(members);
  }

  return chatDisplayName(chat, members) || "Chat";
}

function buildExportFileName(chat, me, members) {
  const prefix = "Teams_Chat_Export_";
  const base = buildExportFileBase(chat, me, members);
  const ts = formatFileTimestampYYYYMMDDHHMMSS(new Date());

  // Reserve space for prefix + "_" + timestamp + ".html"
  const maxBaseLen = 140;
  const trimmedBase = trimFileBasePreserveWords(base, maxBaseLen);
  const safeBase = safeFileName(trimmedBase);

  return `${prefix}${safeBase}_${ts}.html`;
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

function isMsalNoTokenRequestCacheError(errorLike) {
  const msg = normalizeErrorMessage(errorLike).toLowerCase();
  return msg.includes("no_token_request_cache_error");
}

function isMsalInteractionInProgressError(errorLike) {
  const msg = normalizeErrorMessage(errorLike).toLowerCase();
  return msg.includes("interaction_in_progress");
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

function isGuidLike(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || "").trim());
}

function extractCanonicalIdentityKeys(value) {
  const raw = String(value || "").trim();
  if (!raw) return [];

  const out = new Set();
  const lower = raw.toLowerCase();

  out.add(lower);

  // Common URI-like wrappers
  if (lower.startsWith("mailto:")) out.add(lower.slice("mailto:".length));
  if (lower.startsWith("sip:")) out.add(lower.slice("sip:".length));

  // Common Teams / Graph identity prefixes
  const knownPrefixes = [
    "8:orgid:",
    "orgid:",
    "8:teamsvisitor:",
    "8:skypeid:",
    "skypeid:",
    "8:",
  ];

  for (const prefix of knownPrefixes) {
    if (lower.startsWith(prefix)) {
      out.add(lower.slice(prefix.length));
    }
  }

  // Extract GUID anywhere inside the identifier
  const guidMatch = lower.match(/[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}/i);
  if (guidMatch) {
    out.add(guidMatch[0].toLowerCase());
  }

  // Extract email if present anywhere
  const emailMatch = lower.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
  if (emailMatch) {
    out.add(emailMatch[0].toLowerCase());
  }

  return [...out].filter(Boolean);
}

function buildIdentityKeySet(values) {
  const set = new Set();

  for (const value of Array.isArray(values) ? values : [values]) {
    for (const k of extractCanonicalIdentityKeys(value)) {
      set.add(k);
    }
  }

  return set;
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
  const targetKeys = buildIdentityKeySet(userId);
  if (!targetKeys.size) return "";

  const myValues = [
    me?.id,
    me?.userPrincipalName,
    me?.mail,
  ].filter(Boolean);

  const myKeys = buildIdentityKeySet(myValues);

  for (const k of targetKeys) {
    if (myKeys.has(k)) {
      return me?.displayName || me?.userPrincipalName || "";
    }
  }

  const hit = (members || []).find((m) => {
    const ids = collectMemberCandidateIds(m);
    const memberKeys = buildIdentityKeySet(ids);

    for (const k of targetKeys) {
      if (memberKeys.has(k)) return true;
    }
    return false;
  });

  return (
    hit?.displayName ||
    hit?.user?.displayName ||
    hit?.email ||
    hit?.userPrincipalName ||
    ""
  );
}

function extractReactionActor(reaction) {
  const root = reaction?.user || {};
  const nestedUser = root?.user || {};

  const displayName =
    String(
      root?.displayName ||
      nestedUser?.displayName ||
      ""
    ).trim();

  const id =
    String(
      root?.id ||
      nestedUser?.id ||
      reaction?.userId ||
      ""
    ).trim();

  const email =
    String(
      root?.email ||
      nestedUser?.email ||
      root?.userPrincipalName ||
      nestedUser?.userPrincipalName ||
      ""
    ).trim();

  return {
    displayName,
    id,
    email,
  };
}

function resolveReactionUserDisplayName(reaction, members, me) {
  const actor = extractReactionActor(reaction);

  if (actor.displayName) {
    return actor.displayName;
  }

  const actorIdentityCandidates = [
    actor.id,
    actor.email,
    reaction?.userId,
    reaction?.user?.user?.id,
    reaction?.user?.id,
    reaction?.user?.user?.email,
    reaction?.user?.email,
    reaction?.user?.user?.userPrincipalName,
    reaction?.user?.userPrincipalName,
  ].filter(Boolean);

  for (const candidate of actorIdentityCandidates) {
    const resolved = resolveUserDisplayNameById(candidate, members, me);
    if (resolved) return resolved;
  }

  const myKeys = buildIdentityKeySet([
    me?.id,
    me?.userPrincipalName,
    me?.mail,
  ]);

  const actorKeys = buildIdentityKeySet(actorIdentityCandidates);

  for (const k of actorKeys) {
    if (myKeys.has(k)) {
      return me?.displayName || me?.userPrincipalName || "";
    }
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

  ui.queueStatus.textContent = t("queueProgress", {
    done,
    running,
    failed,
    total,
  });

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
    return t("queueExporting", {
      current: Math.min(processed + 1, total),
      total,
    });
  }

  if (processed < total) {
    return t("queuePreparing", {
      processed,
      total,
    });
  }

  return t("queueFinished", { total });
}

function setMePill() {
  if (!ui.mePill) return;

  if (!_me) {
    ui.mePill.textContent = t("notSignedIn");
    return;
  }

  ui.mePill.textContent = t("signedInAs", {
    name: _me.displayName || _me.userPrincipalName || _me.id,
  });
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
      : `${t("exportSelected")}${_selectedChatIds.size > 0 ? ` (${_selectedChatIds.size})` : ""}`;
  }

  if (ui.btnQueueAll) {
    ui.btnQueueAll.disabled = b || !_me || _chats.length === 0;
    ui.btnQueueAll.textContent = _queueRunning
      ? getQueueProgressText()
      : `${t("queueAll")}${_chats.length > 0 ? ` (${_chats.length})` : ""}`;
  }

  if (ui.btnRetryFailed) {
    const failedCount = getQueueSummaryCounts().failed;
    ui.btnRetryFailed.disabled = b || !_me || failedCount === 0;
    ui.btnRetryFailed.textContent = _queueRunning
      ? getQueueProgressText()
      : `${t("retryFailedOnly")}${failedCount > 0 ? ` (${failedCount})` : ""}`;
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
    appendLogLine(t("reloginRedirectFailed", { message: normalizeErrorMessage(e) }));
    alert(`${t("reloginFailed")}:\n${normalizeErrorMessage(e)}`);
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
    void forceReLogin(t("sessionExpired"));
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
  if (_authInited) {
    return msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
  }

  if (_authInitPromise) {
    return await _authInitPromise;
  }

  _authInitPromise = (async () => {
    await msalInstance.initialize();
    loadAppSessionMeta();

    let resp = null;

    try {
      resp = await msalInstance.handleRedirectPromise();
    } catch (e) {
      if (isMsalNoTokenRequestCacheError(e)) {
        console.warn("[MSAL] Ignoring no_token_request_cache_error during redirect handling:", e);
        resp = null;
      } else {
        throw e;
      }
    }

    if (resp?.account) {
      msalInstance.setActiveAccount(resp.account);
      recordAuthenticatedSession({ reset: true });
    }

    const acc = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
    if (acc) {
      msalInstance.setActiveAccount(acc);
    }

    _authInited = true;
    return acc;
  })();

  try {
    return await _authInitPromise;
  } finally {
    _authInitPromise = null;
  }
}

async function login() {
  try {
    await initAuth();

    if (msalInstance.getActiveAccount()) {
      appendLogLine(t("alreadySignedIn"));
      return;
    }

    const interactionStatus = sessionStorage.getItem("msal.interaction.status");
    if (interactionStatus) {
      appendLogLine(t("interactionAlreadyInProgress", { status: interactionStatus }));
      return;
    }

    appendLogLine(t("loginRedirectStart"));

    await msalInstance.loginRedirect({
      scopes: SCOPES,
      prompt: "select_account",
      redirectStartPage: REDIRECT_URI,
    });
  } catch (e) {
    if (isMsalNoTokenRequestCacheError(e)) {
      console.warn("[MSAL] Suppressed no_token_request_cache_error in login():", e);
      appendLogLine(t("msalCacheMismatchRetry"));
      return;
    }

    if (isMsalInteractionInProgressError(e)) {
      appendLogLine(t("loginInProgress"));
      return;
    }

    console.error(e);
    const msg = normalizeErrorMessage(e);
    appendLogLine(`${t("loginFailed")}: ${msg}`);
    alert(`${t("loginFailed")}:\n${msg}`);
  }
}

async function getAccessToken() {
  await initAuth();

  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!account) throw new Error("No active account. Click Login first.");

  try {
    const res = await msalInstance.acquireTokenSilent({ account, scopes: SCOPES });
    return res.accessToken;
  } catch (e) {
    if (isMsalInteractionInProgressError(e)) {
      throw new Error("Authentication interaction already in progress.");
    }

    await msalInstance.acquireTokenRedirect({
      account,
      scopes: SCOPES,
      redirectStartPage: REDIRECT_URI,
    });

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

  if (isUnsupportedMembersEndpointChatId(key)) {
    _chatMembersCache.set(key, []);
    return [];
  }

  try {
    const members = await loadChatMembers(key);
    _chatMembersCache.set(key, members || []);
    return members || [];
  } catch (e) {
    const msg = normalizeErrorMessage(e);
    console.warn("[CHAT MEMBERS LOAD FAILED]", { chatId: key, error: msg });

    _chatMembersCache.set(key, []);
    return [];
  }
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

function extractSender(msg) {
  return msg?.from?.user?.displayName || msg?.from?.application?.displayName || "";
}

function isFromMe(msg, myUserId) {
  const uid = msg?.from?.user?.id || "";
  return !!(uid && myUserId && uid === myUserId);
}

function isEmojiOnlyHtml(html) {
  const text = extractPlainTextFromHtml(normalizeTeamsHtml(html || ""))
    .replace(/\uFE0F/g, "")
    .replace(/\s+/g, "")
    .trim();

  if (!text) return false;

  return /^[\p{Extended_Pictographic}\u2600-\u27BF\u{1F1E6}-\u{1F1FF}]+$/u.test(text);
}

function extractBodyHtml(msg) {
  const b = msg?.body;
  if (!b?.content) return "";

  if ((b.contentType || "").toLowerCase() === "text") {
    return normalizeTeamsHtml(`<div>${esc(b.content)}</div>`);
  }

  return normalizeTeamsHtml(String(b.content));
}

function extractPlainTextFromHtml(html) {
  const src = String(html || "").trim();
  if (!src) return "";

  try {
    const doc = new DOMParser().parseFromString(src, "text/html");
    return (doc.body?.textContent || "").replace(/\s+/g, " ").trim();
  } catch {
    return src.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  }
}

function extractTextFromAnyHtmlish(value) {
  const src = String(value || "").trim();
  if (!src) return "";

  try {
    const doc = new DOMParser().parseFromString(src, "text/html");
    return (doc.body?.textContent || "").replace(/\s+/g, " ").trim();
  } catch {
    return src.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  }
}

function collectEventDetailTextCandidates(eventDetail) {
  if (!eventDetail || typeof eventDetail !== "object") return [];

  const candidates = [];

    const push = (v) => {
    const s = extractTextFromAnyHtmlish(v);
    if (!s) return;
    if (isGuidLike(s)) return;
    if (isIsoDateTimeLike(s)) return;
    if (s === "0001-01-01T00:00:00Z") return;
    if (s === "aadUser") return;
    if (s.startsWith("#microsoft.graph.")) return;

    candidates.push(s);
  };

  for (const [key, value] of Object.entries(eventDetail)) {
    if (key === "@odata.type") continue;

    if (value == null) continue;

    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
      push(String(value));
      continue;
    }

    if (Array.isArray(value)) {
      for (const item of value) {
        if (item == null) continue;

        if (typeof item === "string" || typeof item === "number" || typeof item === "boolean") {
          push(String(item));
          continue;
        }

        if (typeof item === "object") {
          push(item.displayName);
          push(item.userPrincipalName);
          push(item.email);
          push(item.id);

          for (const nested of Object.values(item)) {
            if (typeof nested === "string" || typeof nested === "number" || typeof nested === "boolean") {
              push(String(nested));
            }
          }
        }
      }
      continue;
    }

    if (typeof value === "object") {
      push(value.displayName);
      push(value.userPrincipalName);
      push(value.email);
      push(value.id);

      for (const nested of Object.values(value)) {
        if (typeof nested === "string" || typeof nested === "number" || typeof nested === "boolean") {
          push(String(nested));
        }
      }
    }
  }

  return [...new Set(candidates.filter(Boolean))];
}

function getEventDetailActorId(eventDetail) {
  if (!eventDetail || typeof eventDetail !== "object") return "";

  const candidates = [
    eventDetail?.initiator?.user?.id,
    eventDetail?.initiator?.id,
    eventDetail?.from?.user?.id,
    eventDetail?.from?.id,
    eventDetail?.sender?.user?.id,
    eventDetail?.sender?.id,
    eventDetail?.user?.id,
    eventDetail?.userId,
    eventDetail?.actor?.user?.id,
    eventDetail?.actor?.id,
  ]
    .map((x) => String(x || "").trim())
    .filter(Boolean);

  return candidates[0] || "";
}

function getEventDetailTargetIds(eventDetail) {
  if (!eventDetail || typeof eventDetail !== "object") return [];

  const out = [];

  const add = (value) => {
    const id = String(value || "").trim();
    if (id) out.push(id);
  };

  const arraysToCheck = [
    eventDetail?.members,
    eventDetail?.memberDetails,
    eventDetail?.participants,
    eventDetail?.users,
    eventDetail?.targets,
    eventDetail?.addedMembers,
    eventDetail?.removedMembers,
  ];

  for (const arr of arraysToCheck) {
    if (!Array.isArray(arr)) continue;

    for (const item of arr) {
      add(item?.user?.id);
      add(item?.id);
      add(item?.userId);
      add(item?.member?.id);
      add(item?.member?.user?.id);
    }
  }

  const directCandidates = [
    eventDetail?.member?.user?.id,
    eventDetail?.member?.id,
    eventDetail?.target?.user?.id,
    eventDetail?.target?.id,
    eventDetail?.targetUserId,
    eventDetail?.addedMember?.user?.id,
    eventDetail?.addedMember?.id,
    eventDetail?.removedMember?.user?.id,
    eventDetail?.removedMember?.id,
  ];

  for (const c of directCandidates) add(c);

  return [...new Set(out)];
}

function resolveNameListByIds(ids, members, me) {
  const out = [];

  for (const id of ids) {
    const name = resolveUserDisplayNameById(id, members, me);
    if (name) out.push(name);
  }

  return [...new Set(out)];
}

function buildKnownMemberIdSet(members, me) {
  const set = new Set();

  for (const member of Array.isArray(members) ? members : []) {
    for (const value of collectMemberCandidateIds(member)) {
      for (const key of extractCanonicalIdentityKeys(value)) {
        set.add(key);
      }
    }
  }

  for (const value of [me?.id, me?.userPrincipalName, me?.mail].filter(Boolean)) {
    for (const key of extractCanonicalIdentityKeys(value)) {
      set.add(key);
    }
  }

  return set;
}

function extractAllGuidLikeIdsDeep(obj) {
  const out = new Set();

  function walk(value) {
    if (value == null) return;

    if (typeof value === "string") {
      const s = value.trim();
      if (isGuidLike(s)) out.add(s);
      return;
    }

    if (Array.isArray(value)) {
      for (const item of value) walk(item);
      return;
    }

    if (typeof value === "object") {
      for (const v of Object.values(value)) walk(v);
    }
  }

  walk(obj);
  return [...out];
}

function filterIdsToKnownChatMembers(ids, members, me) {
  const known = buildKnownMemberIdSet(members, me);

  return [...new Set(
    (Array.isArray(ids) ? ids : [])
      .map((x) => String(x || "").trim())
      .filter(Boolean)
      .filter((id) => {
        const keys = extractCanonicalIdentityKeys(id);
        return keys.some((k) => known.has(k));
      })
  )];
}

function getDeepStringValues(obj) {
  const out = [];

  function walk(value) {
    if (value == null) return;

    if (typeof value === "string") {
      const s = value.trim();
      if (s) out.push(s);
      return;
    }

    if (Array.isArray(value)) {
      for (const item of value) walk(item);
      return;
    }

    if (typeof value === "object") {
      for (const v of Object.values(value)) walk(v);
    }
  }

  walk(obj);
  return out;
}

function extractRenameTargetName(eventDetail) {
  const isTeamsOpaqueId = (value) => {
    const s = String(value || "").trim();
    if (!s) return false;

    const lower = s.toLowerCase();

    // Common Teams / chat thread identifiers
    if (/^19:[a-z0-9._:-]+@thread\.v2$/i.test(s)) return true;
    if (/^19:[a-z0-9._:-]+@unq\.gbl\.spaces$/i.test(s)) return true;
    if (/^28:[a-z0-9._:-]+$/i.test(s)) return true;
    if (/^8:[a-z0-9._:-]+$/i.test(s)) return true;

    // Opaque ids that look like service routing identifiers, not titles
    if (lower.includes("@thread.v2")) return true;
    if (lower.includes("@unq.gbl.spaces")) return true;

    return false;
  };

  const isBadRenameCandidate = (value) => {
    const s = String(value || "").trim();
    if (!s) return true;

    if (isIsoDateTimeLike(s)) return true;

    const lower = s.toLowerCase();

    if (s === "aadUser") return true;
    if (s === "0001-01-01T00:00:00Z") return true;
    if (isGuidLike(s)) return true;
    if (isTeamsOpaqueId(s)) return true;

    // OData / Graph type markers
    if (s.startsWith("#microsoft.graph.")) return true;
    if (lower.includes("microsoft.graph.")) return true;
    if (lower.includes("@odata.type")) return true;

    // Known non-title event markers
    const blocked = new Set([
      "unknownfuturevalue",
      "chatrenamedeventmessagedetail",
      "membersaddedeventmessagedetail",
      "membersremovedeventmessagedetail",
      "historysharedeventmessagedetail",
      "chatrenamed",
      "membersadded",
      "membersremoved",
      "historyshared",
      "eventmessage",
      "systemevent",
      "thread.v2",
    ]);

    if (blocked.has(lower)) return true;

    return false;
  };

  const preferredKeys = [
    "topic",
    "chatName",
    "newTopic",
    "newChatName",
    "displayName",
    "title",
    "name",
  ];

  for (const key of preferredKeys) {
    const v = String(eventDetail?.[key] || "").trim();
    if (v && !isBadRenameCandidate(v)) return v;
  }

  const strings = getDeepStringValues(eventDetail)
    .map((x) => String(x || "").trim())
    .filter(Boolean)
    .filter((x) => !isBadRenameCandidate(x));

  const humanish = strings.filter((x) => {
    if (x.length < 2) return false;
    if (!/[a-zа-я0-9]/i.test(x)) return false;
    return true;
  });

  humanish.sort((a, b) => b.length - a.length);
  return humanish[0] || "";
}
function resolveLikelyActorAndTargets(eventDetail, members, me) {
  const actorIdRaw = getEventDetailActorId(eventDetail);
  const deepIds = extractAllGuidLikeIdsDeep(eventDetail);
  const knownIds = filterIdsToKnownChatMembers(deepIds, members, me);

  let actorId = "";
  if (actorIdRaw) {
    const actorKeys = extractCanonicalIdentityKeys(actorIdRaw);
    const known = buildKnownMemberIdSet(members, me);

    if (actorKeys.some((k) => known.has(k))) {
      actorId = actorIdRaw;
    }
  }

  let targetIds = getEventDetailTargetIds(eventDetail);
  targetIds = filterIdsToKnownChatMembers(targetIds, members, me);

  // Fallback: if explicit targets are missing, use known deep IDs
  if (!targetIds.length) {
    targetIds = knownIds;
  }

  // If actor is still missing, infer from current user universe:
  // choose one ID as actor and the rest as targets only when we have >1 known IDs.
  if (!actorId && knownIds.length > 1) {
    actorId = knownIds[0];
    targetIds = knownIds.slice(1);
  }

  // Remove actor from targets if present
  if (actorId) {
    const actorNorm = normalizeId(actorId);
    targetIds = targetIds.filter((id) => normalizeId(id) !== actorNorm);
  }

  const actorName = actorId
    ? (resolveUserDisplayNameById(actorId, members, me) || "")
    : "";

  const targetNames = resolveNameListByIds(targetIds, members, me);

  return {
    actorId,
    actorName,
    targetIds,
    targetNames,
  };
}

function isIsoDateTimeLike(value) {
  const s = String(value || "").trim();
  if (!s) return false;

  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?Z$/i.test(s);
}

function extractSystemEventText(msg, fallbackBodyHtml = "", members = [], me = null) {
  const fromBody = extractPlainTextFromHtml(fallbackBodyHtml);
  if (fromBody) return fromBody;

  const eventDetail = msg?.eventDetail || {};
  const type = String(eventDetail?.["@odata.type"] || "").toLowerCase();
  const mt = String(msg?.messageType || "").trim().toLowerCase();

  const resolved = resolveLikelyActorAndTargets(eventDetail, members, me);
  const actorName = resolved.actorName;
  const targetNames = resolved.targetNames;

  const historyShared =
    eventDetail?.historyShared === true ||
    String(eventDetail?.historyShared || "").toLowerCase() === "true";

  if (type.includes("membersadded") || mt.includes("membersadded")) {
    if (actorName && targetNames.length) {
      return `${actorName} added ${targetNames.join(", ")} to the chat${historyShared ? " and shared all chat history" : ""}`;
    }

    if (targetNames.length) {
      return `${targetNames.join(", ")} joined the chat`;
    }
  }

  if (type.includes("membersremoved") || mt.includes("membersremoved")) {
    if (actorName && targetNames.length) {
      return `${actorName} removed ${targetNames.join(", ")} from the chat`;
    }

    if (targetNames.length === 1) {
      return `${targetNames[0]} was removed from the chat`;
    }

    if (targetNames.length > 1) {
      return `${targetNames.join(", ")} were removed from the chat`;
    }
  }

    if (type.includes("chatrenamed") || mt.includes("chatrenamed")) {
  const renamedTo = extractRenameTargetName(eventDetail);

  if (renamedTo) {
    const looksBad =
      renamedTo.startsWith("#microsoft.graph.") ||
      /^19:[a-z0-9._:-]+@thread\.v2$/i.test(renamedTo) ||
      /^19:[a-z0-9._:-]+@unq\.gbl\.spaces$/i.test(renamedTo);

    if (!looksBad) {
      if (actorName) {
        return `${actorName} changed the chat name to "${renamedTo}"`;
      }
      return `Chat name changed to "${renamedTo}"`;
    }
  }

  return actorName
    ? `${actorName} changed the chat name`
    : "Chat name changed";
}

  if (type.includes("historyshared") || mt.includes("historyshared")) {
    if (actorName) return `${actorName} shared chat history`;
    return "Chat history was shared";
  }

    const detailCandidates = collectEventDetailTextCandidates(eventDetail)
    .filter((x) => x !== "aadUser")
    .filter((x) => x !== "0001-01-01T00:00:00Z")
    .filter((x) => !isGuidLike(x))
    .filter((x) => !isIsoDateTimeLike(x))
    .filter((x) => !/^#microsoft\.graph\./i.test(String(x || "").trim()));

  if (detailCandidates.length) {
    const combined = detailCandidates.join(" ").replace(/\s+/g, " ").trim();
    if (combined) return combined;
  }

  const bodyContentDirect = extractTextFromAnyHtmlish(msg?.body?.content);
  if (bodyContentDirect) return bodyContentDirect;

  // 🔥 SPECIAL HANDLING FOR unknownFutureValue
if (mt.includes("unknownfuturevalue") || type.includes("unknownfuturevalue")) {

  if (targetNames.length && actorName) {
    // Try heuristic: if only 1 target → likely remove
    if (targetNames.length === 1) {
      return `${actorName} removed ${targetNames[0]} from the chat`;
    }

    // Multiple → assume add
    if (targetNames.length > 1) {
      return `${actorName} updated chat participants (${targetNames.join(", ")})`;
    }
  }

  if (targetNames.length === 1) {
    return `${targetNames[0]} was removed from the chat`;
  }

  if (targetNames.length > 1) {
    return `Chat participants updated (${targetNames.join(", ")})`;
  }

  return "Chat participants updated";
}
}

function normalizeSystemEventHtml(input, lang) {
  const src = String(input || "").trim();
  const fallback = tFor(lang, "exportSystemEventFallback");

  if (!src) return `<div class="systemEventText">${esc(fallback)}</div>`;

  try {
    const doc = new DOMParser().parseFromString(src, "text/html");
    const text = (doc.body?.textContent || "").replace(/\s+/g, " ").trim();
    if (!text) return `<div class="systemEventText">${esc(fallback)}</div>`;

    return `<div class="systemEventText">${esc(text)}</div>`;
  } catch {
    const text = src.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
    return text
      ? `<div class="systemEventText">${esc(text)}</div>`
      : `<div class="systemEventText">${esc(fallback)}</div>`;
  }
}

function isSystemEventMessage(msg, bodyHtml = "") {
  const mt = String(msg?.messageType || "").trim().toLowerCase();
  const eventType = String(msg?.eventDetail?.["@odata.type"] || "").trim().toLowerCase();
  const text = extractPlainTextFromHtml(bodyHtml).toLowerCase();

  if (mt && mt !== "message") return true;
  if (eventType) return true;

  const systemPatterns = [
    " added ",
    " removed ",
    " joined the chat",
    " left the chat",
    " renamed the chat",
    " changed the chat",
    " changed the topic",
    " shared all chat history",
    " shared chat history",
  ];

  return systemPatterns.some((p) => text.includes(p));
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

    if (!name) {
      console.warn("[REACTION UNRESOLVED]", {
        messageId: msg?.id || "",
        reactionType,
        reaction: r,
        me: {
          id: me?.id || "",
          userPrincipalName: me?.userPrincipalName || "",
          displayName: me?.displayName || "",
        },
        memberSample: (members || []).slice(0, 8).map((m) => ({
          displayName: m?.displayName || "",
          id: m?.id || "",
          userId: m?.userId || "",
          email: m?.email || "",
          user: {
            id: m?.user?.id || "",
            displayName: m?.user?.displayName || "",
            email: m?.user?.email || "",
          },
        })),
      });
    }

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

    if (name) {
      v.names.push(name);
    }
  }

  let out = `<div class="reactions">`;

  for (const [, v] of [...acc.entries()].sort((a, b) => {
    return (
      b[1].count - a[1].count ||
      String(a[1].visual).localeCompare(String(b[1].visual))
    );
  })) {
    const uniqNames = [...new Set(v.names.map((x) => String(x).trim()).filter(Boolean))].slice(0, 100);
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

function renderForwardBlock(fwd, lang) {
  if (!fwd?.isForwarded) return "";

  const meta = [fwd.originalAuthor, fwd.originalTs].filter(Boolean).join(" • ") || "—";
  const bodyHtml = fwd.originalBodyHtml || "";

  return `
    <div class="fwdWrap">
      <div class="fwdHeader">
        <span class="fwdChip">${esc(tFor(lang, "exportForwarded"))}</span>
        <div class="fwdMeta">${esc(meta)}</div>
      </div>
      <div class="fwdBody">
        <div class="fwdInner">
          ${bodyHtml || `<div class="fwdEmpty">${esc(tFor(lang, "exportNoContent"))}</div>`}
        </div>
      </div>
    </div>
  `;
}

function renderQuoteStack(quotes, lang) {
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
          <span class="qAuthor">${esc(item.author || tFor(lang, "exportQuote"))}</span>
        </div>`;

    const jump = item.refId ? ` <a class="qJump" href="#mid-${esc(item.refId)}">${esc(tFor(lang, "exportView"))}</a>` : "";

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
  const hosted = Array.isArray(hostedList) ? hostedList.filter(Boolean) : [];
  const hostedIds = hosted.map((h) => String(h?.id || "").trim()).filter(Boolean);

  if (!html || !hostedIds.length) {
    return { html, hostedIds: [] };
  }

  try {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const imgs = Array.from(doc.querySelectorAll("img"));

    if (!imgs.length) {
      return { html: doc.body.innerHTML || html, hostedIds };
    }

    const usedHostedIds = new Set();

    // Pass 1: exact/strong match by hosted id inside src or outerHTML
    for (const img of imgs) {
      const src = String(img.getAttribute("src") || "").trim();
      const outer = String(img.outerHTML || "");

      const hit = hostedIds.find((id) => {
        if (usedHostedIds.has(id)) return false;
        return src.includes(id) || outer.includes(id);
      });

      if (!hit) continue;

      img.setAttribute("data-hid", hit);
      if (!img.getAttribute("data-src-orig")) {
        img.setAttribute("data-src-orig", src);
      }

      usedHostedIds.add(hit);
    }

    // Pass 2: fallback sequential assignment for unmatched images
    const remainingHostedIds = hostedIds.filter((id) => !usedHostedIds.has(id));
    const unmatchedImgs = imgs.filter((img) => !img.getAttribute("data-hid"));

    if (remainingHostedIds.length && unmatchedImgs.length) {
      const count = Math.min(remainingHostedIds.length, unmatchedImgs.length);

      for (let i = 0; i < count; i++) {
        const img = unmatchedImgs[i];
        const hid = remainingHostedIds[i];
        const src = String(img.getAttribute("src") || "").trim();

        img.setAttribute("data-hid", hid);
        if (!img.getAttribute("data-src-orig")) {
          img.setAttribute("data-src-orig", src);
        }
      }
    }

    return {
      html: doc.body.innerHTML || html,
      hostedIds,
    };
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

function isUnsupportedMembersEndpointChatId(chatId) {
  const s = String(chatId || "").trim().toLowerCase();
  if (!s) return false;

  return s.includes("@unq.gbl.spaces");
}

async function tryResolveDriveItemFromMeDrivePath(relativePath) {
  const p = String(relativePath || "").trim();
  if (!p) throw new Error("me_drive_relative_path_missing");

  return await graphFetch(`/me/drive/root:/${p}`);
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

    const originalContent = normalizeTeamsHtml(String(obj?.originalMessageContent || ""));
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

async function enrichForwardedAuthors(forwardedItems, currentMembers, me, stats = null) {
  const items = Array.isArray(forwardedItems) ? forwardedItems : [];
  if (!items.length) return items;

  for (const item of items) {
    if (!item?.isForwarded) continue;

    const originalUserId = String(item.originalAuthorId || "").trim();
    const originalConversationId = String(item.originalConversationId || "").trim();
    const originalMessageId = String(item.originalMessageId || "").trim();

    let originalMsg = null;

    if (originalConversationId && originalMessageId) {
      try {
        originalMsg = await loadSingleChatMessage(originalConversationId, originalMessageId);
      } catch (e) {
        appendLogLine(
          `Forwarded original message lookup failed (${originalConversationId}/${originalMessageId}): ${normalizeErrorMessage(e)}`
        );
      }
    }

    // Resolve author
    const alreadyResolved =
      item.originalAuthor &&
      String(item.originalAuthor).trim() &&
      String(item.originalAuthor).trim().toLowerCase() !== "unknown";

    if (!alreadyResolved) {
      let resolved = "";

      if (originalUserId) {
        resolved = resolveUserDisplayNameById(originalUserId, currentMembers, me) || "";
      }

      if (!resolved && originalMsg) {
        resolved =
          originalMsg?.from?.user?.displayName ||
          originalMsg?.from?.application?.displayName ||
          "";
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

    // Resolve timestamp
    if ((!item.originalTs || !String(item.originalTs).trim()) && originalMsg?.createdDateTime) {
      const d = new Date(originalMsg.createdDateTime);
      if (!isNaN(d.getTime())) {
        item.originalTs = formatDDMMYYYY_HHMMSS(d);
      }
    }

    const payloadBodyHtml = normalizeTeamsHtml(String(item.originalBodyHtml || ""));
    const messageBodyHtml = originalMsg
      ? normalizeTeamsHtml(String(extractBodyHtml(originalMsg) || ""))
      : "";

    let originalBodyHtml = payloadBodyHtml || messageBodyHtml || "";

    // Prefer original message body when it contains images or looks richer
    const payloadHasImg = /<img\b/i.test(payloadBodyHtml);
    const messageHasImg = /<img\b/i.test(messageBodyHtml);

    if (messageHasImg) {
      originalBodyHtml = messageBodyHtml;
    } else if (!payloadBodyHtml && messageBodyHtml) {
      originalBodyHtml = messageBodyHtml;
    } else if (
      messageBodyHtml &&
      stripToPreviewText(messageBodyHtml, 500).length > stripToPreviewText(payloadBodyHtml, 500).length
    ) {
      originalBodyHtml = messageBodyHtml;
    }

    // Embed hosted images from original message
    if (originalConversationId && originalMessageId && shouldAttemptHosted(originalBodyHtml)) {
      try {
        const hosted = await listHostedContents(originalConversationId, originalMessageId);

        if (Array.isArray(hosted) && hosted.length) {
          const tok = tokenizeHostedImages(originalBodyHtml, hosted);
          originalBodyHtml = tok.html;

          originalBodyHtml = await embedHostedImages(
            originalConversationId,
            originalMessageId,
            originalBodyHtml,
            hosted,
            stats || {
              imagesEmbedded: 0,
              imagesBytes: 0,
              failures: [],
            }
          );
        }
      } catch (e) {
        if (stats?.failures) {
          stats.failures.push({
            kind: "forwardedHostedImage",
            stage: "enrichForwardedAuthors",
            originalConversationId,
            originalMessageId,
            error: normalizeErrorMessage(e),
          });
        }

        appendLogLine(
          `Forwarded hosted content lookup failed (${originalConversationId}/${originalMessageId}): ${normalizeErrorMessage(e)}`
        );
      }
    }

    item.originalBodyHtml = originalBodyHtml || "";
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

function tokenizeFileAttachments(bodyHtml, attachments, fileTokenToMeta, context = {}) {
  let html = bodyHtml || "";
  if (!Array.isArray(attachments) || !attachments.length) return html;

  let seq = fileTokenToMeta.__seq || 0;

  const chatTitle = String(context.chatTitle || "").trim();
  const messageCreatedDateTime = String(context.messageCreatedDateTime || "").trim();
  const messageCreatedLabel = String(context.messageCreatedLabel || "").trim();

  for (const a of attachments) {
    const name = (a?.name || a?.contentUrl || a?.id || "attachment").toString().trim();
    const contentType = String(a?.contentType || "").trim().toLowerCase();
    const contentUrl = String(a?.contentUrl || "").trim();
    const size =
      typeof a?.size === "number" && a.size > 0
        ? a.size
        : (typeof a?.contentBytes === "string" ? Math.floor((a.contentBytes.length * 3) / 4) : 0);

    const commonMeta = {
      chatTitle,
      messageCreatedDateTime,
      messageCreatedLabel,
      size,
      rawAttachment: a,
    };

    if (a?.contentBytes && a?.contentType) {
      const token = `FILE_${++seq}`;
      fileTokenToMeta.set(token, {
        fileName: safeFileName(name),
        kind: "inline",
        contentType: a.contentType,
        contentBytes: a.contentBytes,
        url: "",
        ...commonMeta,
      });
      html += `<div data-file-token="${token}">📎 ${esc(name)}</div>`;
      continue;
    }

    if (contentUrl && /^https?:\/\//i.test(contentUrl)) {
      const token = `FILE_${++seq}`;
      fileTokenToMeta.set(token, {
        fileName: safeFileName(name),
        kind: "shareLink",
        url: contentUrl,
        contentType,
        ...commonMeta,
      });
      html += `<div data-file-token="${token}">📎 ${esc(name)}</div>`;
      continue;
    }

    const token = `FILE_${++seq}`;
    fileTokenToMeta.set(token, {
      fileName: safeFileName(name),
      kind: "unknown",
      url: "",
      contentType,
      ...commonMeta,
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
    const fileName = meta?.fileName || "File";
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
      let blob = null;
      let mime = String(meta?.contentType || "").trim() || "application/octet-stream";
      let fileSize = Number(meta?.size || 0);

      if (meta?.kind === "inline" && meta?.contentBytes && meta?.contentType) {
        mime = meta.contentType || "application/octet-stream";
        blob = base64ToBlob(meta.contentBytes, mime);
        fileSize = blob?.size || Math.floor((meta.contentBytes.length * 3) / 4);
      } else if (meta?.kind === "shareLink" && meta?.url) {
        const driveItem = await getDriveItemFromSharingUrl(meta.url);
        const driveId = driveItem?.parentReference?.driveId;
        const itemId = driveItem?.id;

        if (!driveId || !itemId) {
          throw new Error("driveItem_missing_ids");
        }

        blob = await downloadDriveItemContent(driveId, itemId);
        mime =
          blob?.type ||
          meta?.contentType ||
          driveItem?.file?.mimeType ||
          "application/octet-stream";
        fileSize = blob?.size || fileSize || 0;
      } else if (meta?.kind === "unknown") {
        console.warn("[ATTACHMENT UNRESOLVED]", {
          token,
          fileName,
          contentType: meta?.contentType || "",
          rawAttachment: meta?.rawAttachment || null,
        });

        throw new Error(
          meta?.contentType
            ? `attachment_unresolved_no_url_or_bytes (${meta.contentType})`
            : "attachment_unresolved_no_url_or_bytes"
        );
      } else {
        throw new Error("unknown_attachment_kind");
      }

      if (!blob) {
        throw new Error("blob_missing");
      }

      if (fileSize <= 0) {
        throw new Error("blob_empty");
      }

      if (fileSize > LIMITS.MAX_ATTACH_BYTES_EACH) {
        if (!Array.isArray(stats.skippedLargeFiles)) {
          stats.skippedLargeFiles = [];
        }

        stats.skippedLargeFiles.push({
          chatTitle: meta?.chatTitle || "",
          fileName,
          sizeBytes: fileSize,
          sentAt: meta?.messageCreatedDateTime || "",
          sentAtLabel: meta?.messageCreatedLabel || "",
        });

        stats.failures.push({
          kind: "attachment",
          stage: "embedFileAttachments",
          token,
          fileName,
          error: "too_large_each",
        });

        continue;
      }

      if (total + fileSize > LIMITS.MAX_ATTACH_BYTES_TOTAL) {
        throw new Error("too_large_total");
      }

      const dataUrl = await blobToDataUrl(blob);
      if (!dataUrl.startsWith("data:")) {
        throw new Error("dataurl_invalid");
      }

      tokenToDataUrl.set(token, {
        dataUrl,
        fileName,
        size: fileSize,
        mime,
      });

      stats.attachEmbedded = (stats.attachEmbedded || 0) + 1;
      stats.attachBytes = (stats.attachBytes || 0) + fileSize;
      total += fileSize;
    } catch (e) {
      stats.failures.push({
        kind: "attachment",
        stage: "embedFileAttachments",
        token,
        fileName,
        error: normalizeErrorMessage(e),
      });
    }
  }

  return tokenToDataUrl;
}

function replaceFileTokensOffline(html, tokenToDataUrl, fileTokenToMeta, stats = null, lang = "en") {
  let out = html || "";
  const failureMap = buildAttachmentFailureMap(stats);

  out = out.replaceAll(
    /<div\b[^>]*\bdata-file-token=(["'])(FILE_\d+)\1[^>]*>[\s\S]*?<\/div>/gi,
    (full, q, token) => {
      const meta = fileTokenToMeta.get(token) || {};
      const fileName = meta.fileName || tFor(lang, "attachmentFileFallback");
      const got = tokenToDataUrl.get(token);
      const failure = failureMap.get(token);

      if (got?.dataUrl) {
        const sizeLabel = got.size ? formatBytesHuman(got.size) : "";
        const mime = got.mime || meta.contentType || "application/octet-stream";

        return `
          <div class="attachmentCard embedded">
            <div class="attIcon">📎</div>
            <div class="attMain">
              <div class="attName" title="${esc(fileName)}">${esc(fileName)}</div>
              ${sizeLabel ? `<div class="attMeta">${esc(sizeLabel)}</div>` : ``}
            </div>
            <div class="attActions">
              <button
                type="button"
                class="attPreviewBtn"
                data-file-url="${esc(got.dataUrl)}"
                data-file-name="${esc(fileName)}"
                data-file-mime="${esc(mime)}"
              >
                ${esc(tFor(lang, "attachmentPreview"))}
              </button>
              <a class="attDownload" download="${esc(fileName)}" href="${esc(got.dataUrl)}">${esc(tFor(lang, "attachmentDownloadOffline"))}</a>
            </div>
          </div>
        `;
      }

      let reasonText = tFor(lang, "attachmentMissingGeneric");
      const err = String(failure?.error || "").toLowerCase();

      if (
        err.includes("404") ||
        err.includes("notfound") ||
        err.includes("accessdenied") ||
        err.includes("share_access_denied_or_expired")
      ) {
        reasonText = tFor(lang, "attachmentMissingUnavailable");
      } else if (err.includes("user_declined_large_embed")) {
        reasonText = tFor(lang, "attachmentMissingDeclinedLarge");
      } else if (err.includes("denied_extension")) {
        reasonText = tFor(lang, "attachmentMissingBlockedType");
      } else if (err.includes("too_large_each")) {
        reasonText = tFor(lang, "attachmentMissingTooLarge");
      } else if (err.includes("too_large_total")) {
        reasonText = tFor(lang, "attachmentMissingTotalLimit");
      } else if (err.includes("attachment_unresolved")) {
        reasonText = tFor(lang, "attachmentMissingUnresolved");
      }

      return `
        <div class="attachmentCard missingLocal">
          <div class="attIcon">📎</div>
          <div class="attMain">
            <div class="attName" title="${esc(fileName)}">${esc(fileName)}</div>
            <div class="attMissing">${esc(reasonText)}</div>
          </div>
        </div>
      `;
    }
  );

  return out;
}

function parsePersonalSharePointUrl(rawUrl) {
  const s = String(rawUrl || "").trim();
  if (!s) return null;

  try {
    const u = new URL(s);
    const host = String(u.hostname || "").toLowerCase();
    const path = String(u.pathname || "");

    if (!host.includes("-my.sharepoint.com")) return null;

    const m = path.match(/^\/personal\/([^/]+)\/Documents\/(.+)$/i);
    if (!m) return null;

    const ownerAlias = decodeURIComponent(m[1] || "").trim();
    const relativePathAfterDocuments = decodeURIComponent(m[2] || "")
      .replace(/^\/+/, "")
      .replace(/\\/g, "/")
      .trim();

    if (!ownerAlias || !relativePathAfterDocuments) return null;

    const personalSitePath = `/personal/${ownerAlias}`;

    return {
      host,
      personalSitePath,
      ownerAlias,
      relativePathAfterDocuments,
      fileName: relativePathAfterDocuments.split("/").pop() || "",
    };
  } catch {
    return null;
  }
}

function buildPossibleUpnsFromPersonalAlias(alias) {
  const raw = String(alias || "").trim().toLowerCase();
  const out = new Set();

  if (raw) {
    const parts = raw.split("_").filter(Boolean);

    if (parts.length >= 3) {
      const domainLast = parts[parts.length - 1];
      const domainFirst = parts[parts.length - 2];
      const localParts = parts.slice(0, -2);

      const domain = `${domainFirst}.${domainLast}`;
      const localUnderscore = localParts.join("_");
      const localDot = localParts.join(".");
      const localFlat = localParts.join("");

      if (localUnderscore) out.add(`${localUnderscore}@${domain}`);
      if (localDot) out.add(`${localDot}@${domain}`);
      if (localFlat) out.add(`${localFlat}@${domain}`);
    }
  }

  // Add exact signed-in user identity as strongest fallback
  if (_me?.userPrincipalName) out.add(String(_me.userPrincipalName).trim().toLowerCase());
  if (_me?.mail) out.add(String(_me.mail).trim().toLowerCase());

  return [...out].filter(Boolean);
}

async function trySearchDriveItemInMeDrive(fileName, folderHint = "") {
  const name = String(fileName || "").trim();
  const hint = String(folderHint || "").trim().toLowerCase();

  if (!name) throw new Error("me_drive_search_name_missing");

  const data = await graphFetch(
    `/me/drive/root/search(q='${name.replaceAll("'", "''")}')`
  );

  const list = Array.isArray(data?.value) ? data.value : [];
  if (!list.length) throw new Error("me_drive_search_empty");

  const ranked = [...list].sort((a, b) => {
    const aPath = String(a?.parentReference?.path || "").toLowerCase();
    const bPath = String(b?.parentReference?.path || "").toLowerCase();

    const aHint = hint && aPath.includes(hint) ? 1 : 0;
    const bHint = hint && bPath.includes(hint) ? 1 : 0;
    if (aHint !== bHint) return bHint - aHint;

    const aName = String(a?.name || "").toLowerCase();
    const bName = String(b?.name || "").toLowerCase();
    const want = name.toLowerCase();

    const aExact = aName === want ? 1 : 0;
    const bExact = bName === want ? 1 : 0;
    if (aExact !== bExact) return bExact - aExact;

    return 0;
  });

  return ranked[0] || null;
}

async function trySearchDriveItemInUserDrive(userIdOrUpn, fileName, folderHint = "") {
  const u = String(userIdOrUpn || "").trim();
  const name = String(fileName || "").trim();
  const hint = String(folderHint || "").trim().toLowerCase();

  if (!u) throw new Error("user_drive_search_user_missing");
  if (!name) throw new Error("user_drive_search_name_missing");

  const data = await graphFetch(
    `/users/${encodeURIComponent(u)}/drive/root/search(q='${name.replaceAll("'", "''")}')`
  );

  const list = Array.isArray(data?.value) ? data.value : [];
  if (!list.length) throw new Error("user_drive_search_empty");

  const ranked = [...list].sort((a, b) => {
    const aPath = String(a?.parentReference?.path || "").toLowerCase();
    const bPath = String(b?.parentReference?.path || "").toLowerCase();

    const aHint = hint && aPath.includes(hint) ? 1 : 0;
    const bHint = hint && bPath.includes(hint) ? 1 : 0;
    if (aHint !== bHint) return bHint - aHint;

    const aName = String(a?.name || "").toLowerCase();
    const bName = String(b?.name || "").toLowerCase();
    const want = name.toLowerCase();

    const aExact = aName === want ? 1 : 0;
    const bExact = bName === want ? 1 : 0;
    if (aExact !== bExact) return bExact - aExact;

    return 0;
  });

  return ranked[0] || null;
}

async function tryResolvePersonalSite(siteHost, personalSitePath) {
  const host = String(siteHost || "").trim().toLowerCase();
  const sitePath = String(personalSitePath || "").trim();

  if (!host) throw new Error("personal_site_host_missing");
  if (!sitePath) throw new Error("personal_site_path_missing");

  const normalizedPath = sitePath.startsWith("/") ? sitePath : `/${sitePath}`;

  return await graphFetch(`/sites/${host}:${normalizedPath}`);
}

async function getDriveItemFromSharingUrl(sharingUrl) {
  const url = String(sharingUrl || "").trim();
  if (!url) {
    throw new Error("sharing_url_missing");
  }

  // 1) Primary path: shares API
  try {
    const shareId = toShareIdFromUrl(url);
    return await graphFetch(`/shares/${encodeURIComponent(shareId)}/driveItem`);
  } catch (shareErr) {
    const personalInfo = parsePersonalSharePointUrl(url);

    if (!personalInfo) {
      throw shareErr;
    }

    const {
      host,
      personalSitePath,
      relativePathAfterDocuments,
      fileName,
      ownerAlias,
    } = personalInfo;

    let mePathErr = null;
    let meSearchErr = null;
    let ownerPathErr = null;
    let ownerSearchErr = null;
    let siteResolveErr = null;
    let sitePathErr = null;
    let siteSearchErr = null;

    // 2) current signed-in drive exact path
    try {
      return await tryResolveDriveItemFromMeDrivePath(relativePathAfterDocuments);
    } catch (err) {
      mePathErr = err;
    }

    // 3) current signed-in drive search
    try {
      const foundInMe = await trySearchDriveItemInMeDrive(
        fileName,
        "microsoft teams chat files"
      );

      if (foundInMe?.id && foundInMe?.parentReference?.driveId) {
        return foundInMe;
      }

      meSearchErr = new Error("me_drive_search_no_matching_item");
    } catch (err) {
      meSearchErr = err;
    }

    // 4) user-drive candidates
    const ownerCandidates = buildPossibleUpnsFromPersonalAlias(ownerAlias);

    console.warn("[SHARE OWNER CANDIDATES]", {
      ownerAlias,
      me: _me,
      ownerCandidates,
      relativePathAfterDocuments,
      fileName,
      host,
      personalSitePath,
    });

    for (const candidate of ownerCandidates) {
      try {
        return await tryResolveDriveItemFromUserDrivePath(
          candidate,
          relativePathAfterDocuments
        );
      } catch (err) {
        ownerPathErr = err;
      }
    }

    for (const candidate of ownerCandidates) {
      try {
        const found = await trySearchDriveItemInUserDrive(
          candidate,
          fileName,
          "microsoft teams chat files"
        );

        if (found?.id && found?.parentReference?.driveId) {
          return found;
        }

        ownerSearchErr = new Error("user_drive_search_no_matching_item");
      } catch (err) {
        ownerSearchErr = err;
      }
    }

    // 5) personal site resolution
    let site = null;
    try {
      site = await tryResolvePersonalSite(host, personalSitePath);
    } catch (err) {
      siteResolveErr = err;
    }

    const siteId = String(site?.id || "").trim();

    // 6) site drive exact path
    if (siteId) {
      try {
        return await tryResolveDriveItemFromSiteDrivePath(
          siteId,
          relativePathAfterDocuments
        );
      } catch (err) {
        sitePathErr = err;
      }

      // 7) site drive search
      try {
        const foundInSite = await trySearchDriveItemInSiteDrive(
          siteId,
          fileName,
          "microsoft teams chat files"
        );

        if (foundInSite?.id && foundInSite?.parentReference?.driveId) {
          return foundInSite;
        }

        siteSearchErr = new Error("site_drive_search_no_matching_item");
      } catch (err) {
        siteSearchErr = err;
      }
    }

    throw new Error(
      `share_me_full_resolution_failed | share=${normalizeErrorMessage(shareErr)} | mePath=${normalizeErrorMessage(mePathErr)} | meSearch=${normalizeErrorMessage(meSearchErr)} | ownerCandidates=${ownerCandidates.join(",")} | ownerPath=${normalizeErrorMessage(ownerPathErr)} | ownerSearch=${normalizeErrorMessage(ownerSearchErr)} | siteResolve=${normalizeErrorMessage(siteResolveErr)} | sitePath=${normalizeErrorMessage(sitePathErr)} | siteSearch=${normalizeErrorMessage(siteSearchErr)}`
    );
  }
}

async function tryResolveDriveItemFromSiteDrivePath(siteId, relativePath) {
  const sid = String(siteId || "").trim();
  const p = String(relativePath || "").trim();

  if (!sid) throw new Error("site_id_missing");
  if (!p) throw new Error("site_drive_relative_path_missing");

  return await graphFetch(`/sites/${encodeURIComponent(sid)}/drive/root:/${p}`);
}

async function trySearchDriveItemInSiteDrive(siteId, fileName, folderHint = "") {
  const sid = String(siteId || "").trim();
  const name = String(fileName || "").trim();
  const hint = String(folderHint || "").trim().toLowerCase();

  if (!sid) throw new Error("site_search_site_id_missing");
  if (!name) throw new Error("site_search_name_missing");

  const data = await graphFetch(
    `/sites/${encodeURIComponent(sid)}/drive/root/search(q='${name.replaceAll("'", "''")}')`
  );

  const list = Array.isArray(data?.value) ? data.value : [];
  if (!list.length) throw new Error("site_drive_search_empty");

  const ranked = [...list].sort((a, b) => {
    const aPath = String(a?.parentReference?.path || "").toLowerCase();
    const bPath = String(b?.parentReference?.path || "").toLowerCase();

    const aHint = hint && aPath.includes(hint) ? 1 : 0;
    const bHint = hint && bPath.includes(hint) ? 1 : 0;
    if (aHint !== bHint) return bHint - aHint;

    const aName = String(a?.name || "").toLowerCase();
    const bName = String(b?.name || "").toLowerCase();
    const want = name.toLowerCase();

    const aExact = aName === want ? 1 : 0;
    const bExact = bName === want ? 1 : 0;
    if (aExact !== bExact) return bExact - aExact;

    return 0;
  });

  return ranked[0] || null;
}

async function tryResolveDriveItemFromUserDrivePath(userIdOrUpn, relativePath) {
  const u = String(userIdOrUpn || "").trim();
  const p = String(relativePath || "").trim();

  if (!u) throw new Error("user_drive_user_missing");
  if (!p) throw new Error("user_drive_relative_path_missing");

  return await graphFetch(`/users/${encodeURIComponent(u)}/drive/root:/${p}`);
}

function buildParticipantsPanel(members, me, lang) {
  const list = Array.isArray(members) ? members : [];
  if (list.length <= 1) return "";

  const normalizedMeIds = new Set(
    [
      me?.id,
      me?.userPrincipalName,
      me?.mail,
    ]
      .map((x) => normalizeId(x))
      .filter(Boolean)
  );

  const unique = [];
  const seen = new Set();

  for (const member of list) {
    if (!member) continue;

    const ids = collectMemberCandidateIds(member)
      .map((x) => normalizeId(x))
      .filter(Boolean);

    const stableKey =
      ids[0] ||
      normalizeId(member?.displayName) ||
      normalizeId(member?.email) ||
      Math.random().toString(36).slice(2);

    if (seen.has(stableKey)) continue;
    seen.add(stableKey);

    const displayName =
      String(
        member?.displayName ||
        member?.user?.displayName ||
        member?.email ||
        member?.userPrincipalName ||
        member?.user?.email ||
        "Unknown participant"
      ).trim();

    const secondary =
      String(
        member?.email ||
        member?.user?.email ||
        member?.userPrincipalName ||
        member?.upn ||
        ""
      ).trim();

    const isMe = ids.some((id) => normalizedMeIds.has(id));

    unique.push({
      displayName,
      secondary,
      isMe,
    });
  }

  unique.sort((a, b) => {
    if (a.isMe !== b.isMe) return a.isMe ? -1 : 1;
    return a.displayName.localeCompare(b.displayName);
  });

  const pluralSuffix = unique.length === 1 ? "" : "s";

  let html = `
    <section class="participantsPanel">
      <div class="participantsHead">
        <div class="participantsTitle">${esc(tFor(lang, "exportParticipantsTitle"))}</div>
        <div class="participantsCount">${esc(tFor(lang, "exportParticipantsCount", { count: unique.length, s: pluralSuffix }))}</div>
      </div>
      <div class="participantsGrid">
  `;

  for (const p of unique) {
    const initials = p.displayName
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((x) => x[0] || "")
      .join("")
      .toUpperCase() || "?";

    html += `
      <div class="participantCard ${p.isMe ? "me" : ""}">
        <div class="participantAvatar">${esc(initials)}</div>
        <div class="participantMeta">
          <div class="participantNameRow">
            <div class="participantName">${esc(p.displayName)}</div>
            ${p.isMe ? `<span class="participantBadge">${esc(tFor(lang, "exportYou"))}</span>` : ``}
          </div>
          ${
            p.secondary
              ? `<div class="participantSecondary">${esc(p.secondary)}</div>`
              : `<div class="participantSecondary empty">—</div>`
          }
        </div>
      </div>
    `;
  }

  html += `
      </div>
    </section>
  `;

  return html;
}

/** =========================
 * EXPORT HTML BUILDER
 * ========================= */
function buildHtml(items, stats, exportTitle, archiveRange, members, me, lang) {
  const generatedAt = new Date().toLocaleString();
  const participantsHtml = buildParticipantsPanel(members, me, lang);

  const getInitials = (name) => {
    const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
    const a = parts[0]?.[0] || "?";
    const b = parts[1]?.[0] || "";
    return (a + b).toUpperCase();
  };

  let htmlMsgs = "";
  let lastSender = null;
  let lastMinuteKey = null;

    for (const m of items) {
    if (m.isSystemEvent) {
      lastSender = null;
      lastMinuteKey = null;

      htmlMsgs += `
        <div class="systemEventRow" id="mid-${esc(m.mid)}" data-mid="${esc(m.mid)}">
          <div class="systemEventInner">
            <div class="systemEventBody">${m.bodyHtml || ""}</div>
            <div class="systemEventTime">${esc(m.dtStr)}</div>
          </div>
        </div>
      `;
      continue;
    }

    if (m.isDeletedNotice) {
      lastSender = null;
      lastMinuteKey = null;

      const deletedAuthor = getDeletedNoticeAuthorLabel(m.sender);

      htmlMsgs += `
        <div class="deletedNoticeRow" id="mid-${esc(m.mid)}" data-mid="${esc(m.mid)}">
          <div class="deletedNoticeInner">
            <div class="deletedNoticeAuthor">${esc(deletedAuthor)}</div>
            <div class="deletedNoticeBody">${m.bodyHtml || ""}</div>
            <div class="deletedNoticeTime">${esc(m.dtStr)}</div>
          </div>
        </div>
      `;
      continue;
    }

    const currentMinuteKey = getMessageMinuteKey(m.dtMs, m.dtStr);
    const sameSenderSameMinute =
      m.sender === lastSender &&
      currentMinuteKey &&
      lastMinuteKey &&
      currentMinuteKey === lastMinuteKey;

    lastSender = m.sender;
    lastMinuteKey = currentMinuteKey || null;

    htmlMsgs += `
      <div class="msgRow ${m.isMe ? "me" : "other"}" id="mid-${esc(m.mid)}" data-mid="${esc(m.mid)}">
        <div class="avatar ${sameSenderSameMinute ? "ghost" : ""}">${esc(getInitials(m.sender))}</div>
        <div class="bubbleWrap">
          <div class="msgHead ${sameSenderSameMinute ? "compact" : ""}">
            <div class="senderLine">
              ${sameSenderSameMinute ? "" : `<span class="sender">${esc(m.sender)}</span>`}
            </div>
            <div class="time">${esc(m.dtStr)}</div>
          </div>
          <div class="bubble ${m.isEmojiOnly ? "emojiOnlyBubble" : ""}">
            <div class="body ${m.isEmojiOnly ? "emojiOnlyBody" : ""}">
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
<title>${esc(exportTitle || tFor(lang, "exportArchiveDefaultTitle"))}</title>
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

  .deletedNoticeRow{
    display:flex;
    justify-content:center;
    padding:6px 0 4px 0;
  }

  .deletedNoticeInner{
    max-width:760px;
    text-align:center;
    padding:0;
  }

  .deletedNoticeAuthor{
    margin-bottom:6px;
    color:rgba(229,231,235,.92);
    font-size:12px;
    font-weight:1000;
    line-height:1.3;
  }

  .deletedNoticeBody{
    display:inline-flex;
    align-items:center;
    gap:8px;
    padding:9px 14px;
    border-radius:10px;
    background:rgba(124,58,237,.18);
    border:1px solid rgba(124,58,237,.28);
    color:#f3f4f6;
    font-size:13px;
    line-height:1.45;
    font-style:italic;
    font-weight:800;
  }

  .deletedNoticeText{
    display:inline;
  }

  .deletedNoticeTime{
    margin-top:6px;
    color:var(--muted);
    font-size:11px;
    font-weight:900;
  }

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

    .emojiOnlyBubble{
    padding:12px 16px;
  }

  .emojiOnlyBody{
    font-size:36px;
    line-height:1.2;
    text-align:left;
  }

  .emojiOnlyBody > div,
  .emojiOnlyBody > p,
  .emojiOnlyBody > span{
    font-size:inherit;
    line-height:inherit;
    margin:0;
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

  .fileWarning{
  border:1px solid rgba(255,120,120,.3);
  background:rgba(255,120,120,.08);
}

.fileWarningText{
  font-size:12px;
  color:#fca5a5;
  margin-top:4px;
}

  .systemEventRow{
    display:flex;
    justify-content:center;
    padding:6px 0 4px 0;
  }

  .systemEventInner{
    max-width:760px;
    text-align:center;
    padding:4px 14px;
  }

  .systemEventBody{
    color:rgba(229,231,235,.78);
    font-size:13px;
    line-height:1.55;
    font-weight:700;
  }

  .systemEventText{
    display:inline;
  }

  .systemEventTime{
    margin-top:6px;
    color:var(--muted);
    font-size:11px;
    font-weight:900;
  }

    .attachmentCard{
    display:flex;
    align-items:center;
    gap:12px;
    padding:10px 12px;
    border-radius:14px;
    margin-top:10px;
    border:1px solid var(--line);
    background: rgba(255,255,255,.04);
    max-width:900px;
  }
  .attachmentCard.embedded{
    background: rgba(37,99,235,.10);
    border-color: rgba(37,99,235,.25);
  }
  .attachmentCard.missingLocal{
    background: rgba(239,68,68,.08);
    border-color: rgba(239,68,68,.20);
  }
  .attIcon{
    font-size:18px;
    flex-shrink:0;
  }
  .attMain{
    flex:1;
    min-width:0;
  }
  .attName{
    min-width:0;
    font-weight:1000;
    font-size:13px;
    white-space:nowrap;
    overflow:hidden;
    text-overflow:ellipsis;
  }
  .attMeta{
    margin-top:4px;
    font-size:11px;
    color:var(--muted);
    font-weight:900;
    line-height:1.3;
    word-break:break-word;
  }
  .attActions{
    display:flex;
    gap:8px;
    align-items:center;
    flex-wrap:wrap;
    flex-shrink:0;
  }
  .attDownload,
  .attPreviewBtn{
    font-size:12px;
    font-weight:1000;
    text-decoration:none;
    color:var(--text);
    padding:7px 10px;
    border-radius:12px;
    border:1px solid var(--line);
    background: rgba(0,0,0,.25);
    white-space:nowrap;
    cursor:pointer;
  }
  .attDownload:hover,
  .attPreviewBtn:hover{
    background: rgba(255,255,255,.06);
  }
  .attMissing{
    margin-top:4px;
    font-size:12px;
    color:#fecaca;
    font-weight:1000;
    line-height:1.4;
    word-break:break-word;
  }

    .filePreviewBackdrop{
    position:fixed;
    inset:0;
    background:rgba(0,0,0,.55);
    opacity:0;
    pointer-events:none;
    transition:opacity .2s ease;
    z-index:99992;
  }

  .filePreviewBackdrop.open{
    opacity:1;
    pointer-events:auto;
  }

  .filePreviewModal{
    position:fixed;
    inset:24px;
    max-width:min(1400px, calc(100vw - 48px));
    max-height:calc(100vh - 48px);
    margin:auto;
    display:flex;
    flex-direction:column;
    background:rgba(11,18,32,.98);
    border:1px solid rgba(255,255,255,.10);
    border-radius:18px;
    box-shadow:0 30px 70px rgba(0,0,0,.45);
    transform:scale(.98);
    opacity:0;
    pointer-events:none;
    transition:transform .2s ease, opacity .2s ease;
    z-index:99993;
    overflow:hidden;
  }

  .filePreviewModal.open{
    transform:scale(1);
    opacity:1;
    pointer-events:auto;
  }

  .filePreviewHead{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:12px;
    padding:16px 18px 14px 18px;
    border-bottom:1px solid rgba(255,255,255,.08);
    background:rgba(255,255,255,.03);
  }

  .filePreviewTitle{
    font-size:15px;
    font-weight:1000;
    line-height:1.25;
    word-break:break-word;
  }

  .filePreviewSub{
    margin-top:6px;
    color:var(--muted);
    font-size:12px;
    line-height:1.4;
    word-break:break-word;
  }

  .filePreviewClose{
    flex-shrink:0;
  }

  .filePreviewBody{
    flex:1;
    min-height:0;
    overflow:auto;
    padding:16px;
    background:rgba(255,255,255,.02);
  }

  .filePreviewFrame,
  .filePreviewObject{
    width:100%;
    min-height:72vh;
    border:none;
    border-radius:14px;
    background:#0b0f19;
  }

  .filePreviewImage{
    display:block;
    max-width:100%;
    max-height:72vh;
    margin:0 auto;
    border-radius:14px;
    border:1px solid rgba(255,255,255,.08);
    background:#0b0f19;
  }

  .filePreviewText{
    margin:0;
    padding:16px;
    border-radius:14px;
    border:1px solid rgba(255,255,255,.08);
    background:#111827;
    color:#e5e7eb;
    font:13px/1.55 Consolas, "Cascadia Mono", "Courier New", monospace;
    white-space:pre-wrap;
    word-break:break-word;
    overflow:auto;
  }

  .filePreviewMedia{
    width:100%;
    max-height:72vh;
    border-radius:14px;
    background:#000;
  }

  .filePreviewUnsupported{
    max-width:760px;
    margin:0 auto;
    padding:18px;
    border-radius:16px;
    border:1px solid rgba(255,255,255,.10);
    background:rgba(255,255,255,.04);
  }

  .filePreviewUnsupportedTitle{
    font-size:15px;
    font-weight:1000;
    margin-bottom:8px;
  }

  .filePreviewUnsupportedText{
    color:var(--muted);
    font-size:13px;
    line-height:1.55;
    margin-bottom:14px;
  }

  .filePreviewActions{
    display:flex;
    gap:10px;
    flex-wrap:wrap;
  }

    .imageZoomBackdrop{
    position:fixed;
    inset:0;
    background:rgba(0,0,0,.72);
    opacity:0;
    pointer-events:none;
    transition:opacity .2s ease;
    z-index:99994;
  }

  .imageZoomBackdrop.open{
    opacity:1;
    pointer-events:auto;
  }

  .imageZoomModal{
    position:fixed;
    inset:24px;
    max-width:min(1500px, calc(100vw - 48px));
    max-height:calc(100vh - 48px);
    margin:auto;
    display:flex;
    flex-direction:column;
    background:rgba(11,18,32,.98);
    border:1px solid rgba(255,255,255,.10);
    border-radius:18px;
    box-shadow:0 30px 70px rgba(0,0,0,.50);
    transform:scale(.985);
    opacity:0;
    pointer-events:none;
    transition:transform .2s ease, opacity .2s ease;
    z-index:99995;
    overflow:hidden;
  }

  .imageZoomModal.open{
    transform:scale(1);
    opacity:1;
    pointer-events:auto;
  }

  .imageZoomHead{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:12px;
    padding:16px 18px 14px 18px;
    border-bottom:1px solid rgba(255,255,255,.08);
    background:rgba(255,255,255,.03);
  }

  .imageZoomTitle{
    font-size:15px;
    font-weight:1000;
    line-height:1.25;
    word-break:break-word;
  }

  .imageZoomSub{
    margin-top:6px;
    color:var(--muted);
    font-size:12px;
    line-height:1.4;
    word-break:break-word;
  }

  .imageZoomBody{
    flex:1;
    min-height:0;
    overflow:auto;
    display:flex;
    align-items:center;
    justify-content:center;
    padding:8px;
    background:rgba(255,255,255,.02);
  }

  .imageZoomImg{
    display:block;
    max-width:calc(100vw - 40px);
    max-height:calc(100vh - 40px);
    width:auto;
    height:auto;
    object-fit:contain;
    border-radius:12px;
    border:1px solid rgba(255,255,255,.08);
    background:#0b0f19;
    box-shadow:0 30px 80px rgba(0,0,0,.5);
  }

  .imageZoomClose{
    flex-shrink:0;
  }

  .body img,
  .fwdInner img,
  .quotePreview img{
    cursor:zoom-in;
  }

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

    .participantsPanel{
    margin: 0 0 14px 0;
    padding: 14px 16px;
    border-radius: 18px;
    border: 1px solid var(--line);
    background: linear-gradient(180deg, rgba(255,255,255,.05), rgba(255,255,255,.03));
    box-shadow: 0 18px 40px rgba(0,0,0,.18);
  }

  .participantsHead{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:12px;
    flex-wrap:wrap;
    margin-bottom:12px;
  }

  .participantsTitle{
    font-size:14px;
    font-weight:1000;
    letter-spacing:.01em;
  }

  .participantsCount{
    color:var(--muted);
    font-size:12px;
    font-weight:900;
    padding:5px 10px;
    border-radius:999px;
    border:1px solid var(--line);
    background:rgba(0,0,0,.20);
  }

  .participantsGrid{
    display:grid;
    grid-template-columns:repeat(auto-fit, minmax(240px, 1fr));
    gap:10px;
  }

  .participantCard{
    display:flex;
    align-items:flex-start;
    gap:12px;
    padding:12px 13px;
    border-radius:14px;
    border:1px solid rgba(255,255,255,.08);
    background:rgba(255,255,255,.035);
  }

  .participantCard.me{
    background:rgba(37,99,235,.10);
    border-color:rgba(37,99,235,.24);
  }

  .participantAvatar{
    width:38px;
    height:38px;
    border-radius:999px;
    display:grid;
    place-items:center;
    flex:0 0 auto;
    font-size:12px;
    font-weight:1000;
    border:1px solid rgba(255,255,255,.10);
    background:rgba(255,255,255,.08);
  }

  .participantMeta{
    min-width:0;
    flex:1;
  }

  .participantNameRow{
    display:flex;
    align-items:center;
    gap:8px;
    flex-wrap:wrap;
  }

  .participantName{
    font-size:13px;
    font-weight:1000;
    line-height:1.25;
    word-break:break-word;
  }

  .participantBadge{
    display:inline-flex;
    align-items:center;
    padding:3px 8px;
    border-radius:999px;
    font-size:11px;
    font-weight:1000;
    color:#dbeafe;
    background:rgba(37,99,235,.18);
    border:1px solid rgba(37,99,235,.30);
  }

  .participantSecondary{
    margin-top:5px;
    color:var(--muted);
    font-size:12px;
    line-height:1.35;
    word-break:break-word;
  }

  .participantSecondary.empty{
    opacity:.55;
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

    .drawerBackdrop{
    position:fixed;
    inset:0;
    background:rgba(0,0,0,.45);
    opacity:0;
    pointer-events:none;
    transition:opacity .22s ease;
    z-index:99990;
  }

  .drawerBackdrop.open{
    opacity:1;
    pointer-events:auto;
  }

  .participantsDrawer{
    position:fixed;
    top:0;
    right:0;
    width:min(420px, 92vw);
    height:100vh;
    background:rgba(11,18,32,.98);
    border-left:1px solid rgba(255,255,255,.10);
    box-shadow:-20px 0 50px rgba(0,0,0,.35);
    transform:translateX(100%);
    transition:transform .24s ease;
    z-index:99991;
    display:flex;
    flex-direction:column;
  }

  .participantsDrawer.open{
    transform:translateX(0);
  }

  .participantsDrawerHead{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:12px;
    padding:18px 18px 14px 18px;
    border-bottom:1px solid rgba(255,255,255,.08);
    background:rgba(255,255,255,.03);
    backdrop-filter:blur(8px);
  }

  .participantsDrawerTitle{
    font-size:16px;
    font-weight:1000;
    line-height:1.2;
  }

  .participantsDrawerSub{
    margin-top:6px;
    color:var(--muted);
    font-size:12px;
    line-height:1.4;
  }

  .participantsDrawerBody{
    flex:1;
    overflow:auto;
    padding:16px;
  }

  .drawerCloseBtn{
    padding:8px 12px;
    min-width:auto;
  }

  .participantsDrawer .participantsPanel{
    margin:0;
    padding:0;
    border:none;
    background:transparent;
    box-shadow:none;
  }

  .participantsDrawer .participantsHead{
    margin-bottom:14px;
  }
  .reactPopover .pName{ font-weight:1000; }
</style>
</head>
<body>
  <div class="topbar">
    <div>
      <div class="title">${esc(exportTitle || tFor(lang, "exportArchiveDefaultTitle"))}</div>
      <div class="meta">
        <span class="pill">${esc(tFor(lang, "exportItems"))}: ${items.length}</span>
        <span class="pill">${esc(tFor(lang, "exportPeriod"))}: ${esc(archiveRange?.label || "—")}</span>
        <span class="pill">${esc(tFor(lang, "exportGenerated"))}: ${esc(generatedAt)}</span>
        <span class="pill">${esc(tFor(lang, "exportImages"))}: ${stats.imagesEmbedded || 0} ${esc(tFor(lang, "exportEmbedded"))} • ${((stats.imagesBytes || 0) / 1024 / 1024).toFixed(1)} MB</span>
        <span class="pill">
          ${esc(tFor(lang, "exportFiles"))}:
          • ${stats.attachEmbedded || 0} ${esc(tFor(lang, "exportEmbedded"))} • ${esc(formatBytesToMB(stats.attachBytes || 0) || "0.00 MB")}
        </span>
      </div>
    </div>
    <div class="actions">
      <button class="btn" id="btnParticipants">👥 ${esc(tFor(lang, "exportParticipants"))}</button>
      <button class="btn" id="btnBottom">⬇ ${esc(tFor(lang, "exportLatest"))}</button>
    </div>
  </div>

    <div class="chatScroller" id="chatScroller">
      <div class="chatInner" id="chatInner">
        ${htmlMsgs}
      </div>
    </div>

    <div class="drawerBackdrop" id="drawerBackdrop"></div>

    <aside class="participantsDrawer" id="participantsDrawer" aria-hidden="true">
      <div class="participantsDrawerHead">
        <div>
          <div class="participantsDrawerTitle">${esc(tFor(lang, "exportParticipantsTitle"))}</div>
          <div class="participantsDrawerSub">${esc(tFor(lang, "exportParticipantsSubtitle"))}</div>
        </div>
        <button class="btn drawerCloseBtn" id="btnCloseParticipants">✕</button>
      </div>

      <div class="participantsDrawerBody">
        ${participantsHtml}
      </div>
    </aside>

    <div class="reactPopover" id="reactPopover" role="dialog" aria-modal="false">
      <div class="tTitle" id="rpTitle">${esc(tFor(lang, "exportReactions"))}</div>
      <div id="rpList"></div>
    </div>

        <div class="filePreviewBackdrop" id="filePreviewBackdrop"></div>

    <section class="filePreviewModal" id="filePreviewModal" aria-hidden="true">
      <div class="filePreviewHead">
        <div>
          <div class="filePreviewTitle" id="filePreviewTitle">${esc(tFor(lang, "exportFilePreview"))}</div>
          <div class="filePreviewSub" id="filePreviewSub">${esc(tFor(lang, "exportOfflinePreview"))}</div>
        </div>
        <button class="btn filePreviewClose" id="btnCloseFilePreview">✕</button>
      </div>
      <div class="filePreviewBody" id="filePreviewBody"></div>
    </section>

        <div class="imageZoomBackdrop" id="imageZoomBackdrop"></div>

    <section class="imageZoomModal" id="imageZoomModal" aria-hidden="true">
      <div class="imageZoomHead">
        <div>
          <div class="imageZoomTitle" id="imageZoomTitle">${esc(tFor(lang, "exportImagePreview"))}</div>
          <div class="imageZoomSub" id="imageZoomSub">${esc(tFor(lang, "exportCloseHint"))}</div>
        </div>
        <button class="btn imageZoomClose" id="btnCloseImageZoom">✕</button>
      </div>
      <div class="imageZoomBody" id="imageZoomBody"></div>
    </section>

<script>
(function(){
  const TXT = {
    reactionListTitle: ${JSON.stringify(tFor(lang, "exportReactionListTitle"))},
    noParticipantNames: ${JSON.stringify(tFor(lang, "exportNoParticipantNames"))},
    filePreview: ${JSON.stringify(tFor(lang, "exportFilePreview"))},
    offlinePreview: ${JSON.stringify(tFor(lang, "exportOfflinePreview"))},
    imagePreview: ${JSON.stringify(tFor(lang, "exportImagePreview"))},
    closeHint: ${JSON.stringify(tFor(lang, "exportCloseHint"))},
    unsupportedPreviewTitle: ${JSON.stringify(tFor(lang, "exportUnsupportedPreviewTitle"))},
    unsupportedPreviewText: ${JSON.stringify(tFor(lang, "exportUnsupportedPreviewText"))},
    fileType: ${JSON.stringify(tFor(lang, "exportFileType"))},
    downloadOffline: ${JSON.stringify(tFor(lang, "exportDownloadOffline"))},
  };

  function fmt(template, vars) {
    var out = String(template || "");
    vars = vars || {};

    for (var k in vars) {
      if (!Object.prototype.hasOwnProperty.call(vars, k)) continue;
      var val = vars[k] == null ? "" : String(vars[k]);
      out = out.split("{" + k + "}").join(val);
    }

    return out;
  }
  const sc = document.getElementById("chatScroller");
  const btn = document.getElementById("btnBottom");
  const btnParticipants = document.getElementById("btnParticipants");
  const btnCloseParticipants = document.getElementById("btnCloseParticipants");
  const participantsDrawer = document.getElementById("participantsDrawer");
  const drawerBackdrop = document.getElementById("drawerBackdrop");

  const pop = document.getElementById("reactPopover");
  const list = document.getElementById("rpList");
  const title = document.getElementById("rpTitle");

  const filePreviewModal = document.getElementById("filePreviewModal");
  const filePreviewBackdrop = document.getElementById("filePreviewBackdrop");
  const filePreviewTitle = document.getElementById("filePreviewTitle");
  const filePreviewSub = document.getElementById("filePreviewSub");
  const filePreviewBody = document.getElementById("filePreviewBody");
  const btnCloseFilePreview = document.getElementById("btnCloseFilePreview");

  const imageZoomModal = document.getElementById("imageZoomModal");
  const imageZoomBackdrop = document.getElementById("imageZoomBackdrop");
  const imageZoomTitle = document.getElementById("imageZoomTitle");
  const imageZoomSub = document.getElementById("imageZoomSub");
  const imageZoomBody = document.getElementById("imageZoomBody");
  const btnCloseImageZoom = document.getElementById("btnCloseImageZoom");

  if (sc && btn) {
    const toBottom = () => { sc.scrollTop = sc.scrollHeight; };
    requestAnimationFrame(() => requestAnimationFrame(toBottom));
    btn.addEventListener("click", toBottom);
  }

  const openDrawer = () => {
    if (!participantsDrawer || !drawerBackdrop) return;
    participantsDrawer.classList.add("open");
    drawerBackdrop.classList.add("open");
    participantsDrawer.setAttribute("aria-hidden", "false");
  };

  const closeDrawer = () => {
    if (!participantsDrawer || !drawerBackdrop) return;
    participantsDrawer.classList.remove("open");
    drawerBackdrop.classList.remove("open");
    participantsDrawer.setAttribute("aria-hidden", "true");
  };

  btnParticipants?.addEventListener("click", openDrawer);
  btnCloseParticipants?.addEventListener("click", closeDrawer);
  drawerBackdrop?.addEventListener("click", closeDrawer);

  function getExt(fileName) {
    const m = String(fileName || "").toLowerCase().match(/\.([a-z0-9]{1,12})$/);
    return m ? m[1] : "";
  }

  function isTextLikeMime(mime) {
    const m = String(mime || "").toLowerCase();
    return (
      m.startsWith("text/") ||
      m.includes("json") ||
      m.includes("xml") ||
      m.includes("javascript") ||
      m.includes("ecmascript") ||
      m.includes("yaml")
    );
  }

    function getPreviewKind(fileName, mime) {
    const ext = getExt(fileName);
    const m = String(mime || "").toLowerCase();

    const officeLikeExt = [
      "doc", "docx", "docm",
      "xls", "xlsx", "xlsm", "xlsb",
      "ppt", "pptx", "pptm",
      "msg", "rtf", "odt", "ods", "odp"
    ];

    const spreadsheetTextExt = ["csv", "tsv"];

    const officeLikeMimeHints = [
      "application/vnd.openxmlformats-officedocument",
      "application/vnd.ms-excel",
      "application/vnd.ms-powerpoint",
      "application/msword",
      "application/vnd.oasis.opendocument",
      "application/vnd.ms-outlook",
      "officedocument",
      "spreadsheetml",
      "wordprocessingml",
      "presentationml"
    ];

    if (m.startsWith("image/")) return "image";
    if (m === "application/pdf" || ext === "pdf") return "pdf";
    if (m.startsWith("audio/")) return "audio";
    if (m.startsWith("video/")) return "video";

    if (spreadsheetTextExt.includes(ext)) return "text";

    if (officeLikeExt.includes(ext)) return "unsupported";
    if (officeLikeMimeHints.some((hint) => m.includes(hint))) return "unsupported";

    if (
      isTextLikeMime(m) ||
      ["txt","md","json","log","xml","html","htm","css","js","ts","yaml","yml","ini","conf","config","sql"].includes(ext)
    ) {
      return "text";
    }

    if (["svg"].includes(ext)) return "image";
    if (["mht","mhtml"].includes(ext)) return "iframe";

    if (m.startsWith("application/")) return "iframe";
    return "unsupported";
  }

    function base64ToBytes(base64) {
    const bin = atob(base64);
    const bytes = new Uint8Array(bin.length);

    for (let i = 0; i < bin.length; i++) {
      bytes[i] = bin.charCodeAt(i);
    }

    return bytes;
  }

  function latin1FallbackFromBytes(bytes) {
    let out = "";
    const chunk = 0x8000;

    for (let i = 0; i < bytes.length; i += chunk) {
      out += String.fromCharCode(...bytes.slice(i, i + chunk));
    }

    return out;
  }

  function decodeBytesToText(bytes) {
    if (!bytes || !bytes.length) return "";

    try {
      return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
    } catch {}

    try {
      return new TextDecoder("utf-8").decode(bytes);
    } catch {}

    return latin1FallbackFromBytes(bytes);
  }

  function dataUrlToText(dataUrl) {
    const s = String(dataUrl || "");
    const commaIndex = s.indexOf(",");
    if (commaIndex < 0) return "";

    const header = s.slice(0, commaIndex);
    const headerLower = header.toLowerCase();
    const payload = s.slice(commaIndex + 1);

    try {
      if (headerLower.includes(";base64")) {
        const bytes = base64ToBytes(payload);
        return decodeBytesToText(bytes);
      }

      const decoded = decodeURIComponent(payload);

      // Convert JS string to UTF-8 bytes and normalize through TextDecoder
      const bytes = new TextEncoder().encode(decoded);
      return decodeBytesToText(bytes);
    } catch {
      try {
        return decodeURIComponent(payload);
      } catch {
        return "Preview decode failed.";
      }
    }
  }

  function clearFilePreview() {
    if (!filePreviewBody) return;
    filePreviewBody.innerHTML = "";
  }

  function openFilePreview(name, mime, dataUrl) {
    if (!filePreviewModal || !filePreviewBackdrop || !filePreviewBody) return;

    const kind = getPreviewKind(name, mime);

    filePreviewTitle.textContent = name || TXT.filePreview;
    filePreviewSub.textContent = (mime || "application/octet-stream") + " • " + kind;
    clearFilePreview();

    if (kind === "image") {
      const img = document.createElement("img");
      img.className = "filePreviewImage";
      img.alt = name || "image preview";
      img.src = dataUrl;
      filePreviewBody.appendChild(img);
    } else if (kind === "pdf") {
      const frame = document.createElement("iframe");
      frame.className = "filePreviewFrame";
      frame.src = dataUrl;
      frame.setAttribute("title", name || "PDF preview");
      filePreviewBody.appendChild(frame);
    } else if (kind === "audio") {
      const audio = document.createElement("audio");
      audio.className = "filePreviewMedia";
      audio.controls = true;
      audio.src = dataUrl;
      filePreviewBody.appendChild(audio);
    } else if (kind === "video") {
      const video = document.createElement("video");
      video.className = "filePreviewMedia";
      video.controls = true;
      video.src = dataUrl;
      filePreviewBody.appendChild(video);
    } else if (kind === "text") {
      const pre = document.createElement("pre");
      pre.className = "filePreviewText";
      pre.textContent = dataUrlToText(dataUrl);
      filePreviewBody.appendChild(pre);
    } else if (kind === "iframe") {
      const frame = document.createElement("iframe");
      frame.className = "filePreviewFrame";
      frame.src = dataUrl;
      frame.setAttribute("title", name || "File preview");
      filePreviewBody.appendChild(frame);
    } else {
      const wrap = document.createElement("div");
      wrap.className = "filePreviewUnsupported";

      const t = document.createElement("div");
      t.className = "filePreviewUnsupportedTitle";
      t.textContent = TXT.unsupportedPreviewTitle;

      const p = document.createElement("div");
      p.className = "filePreviewUnsupportedText";
            const typeInfo = document.createElement("div");
      typeInfo.className = "filePreviewUnsupportedText";
      typeInfo.textContent = TXT.fileType + ": " + (mime || "application/octet-stream");
            p.textContent = TXT.unsupportedPreviewText;

      const actions = document.createElement("div");
      actions.className = "filePreviewActions";

      const dl = document.createElement("a");
      dl.className = "attDownload";
      dl.href = dataUrl;
      dl.download = name || "file";
      dl.textContent = TXT.downloadOffline;

      actions.appendChild(dl);
      wrap.appendChild(t);
      wrap.appendChild(p);
      wrap.appendChild(typeInfo);
      wrap.appendChild(actions);
      filePreviewBody.appendChild(wrap);
    }

    filePreviewModal.classList.add("open");
    filePreviewBackdrop.classList.add("open");
    filePreviewModal.setAttribute("aria-hidden", "false");
  }

  function closeFilePreview() {
    if (!filePreviewModal || !filePreviewBackdrop) return;
    filePreviewModal.classList.remove("open");
    filePreviewBackdrop.classList.remove("open");
    filePreviewModal.setAttribute("aria-hidden", "true");
    clearFilePreview();
  }

    function clearImageZoom() {
    if (!imageZoomBody) return;
    imageZoomBody.innerHTML = "";
  }

  function openImageZoom(src, alt) {
    if (!imageZoomModal || !imageZoomBackdrop || !imageZoomBody) return;
    if (!src) return;

    clearImageZoom();

    const img = document.createElement("img");
    img.className = "imageZoomImg";
    img.src = src;
    img.alt = alt || "image preview";

    imageZoomTitle.textContent = alt || TXT.imagePreview;
    imageZoomSub.textContent = TXT.closeHint;

    imageZoomBody.appendChild(img);
    imageZoomModal.classList.add("open");
    imageZoomBackdrop.classList.add("open");
    imageZoomModal.setAttribute("aria-hidden", "false");
  }

  function closeImageZoom() {
    if (!imageZoomModal || !imageZoomBackdrop) return;
    imageZoomModal.classList.remove("open");
    imageZoomBackdrop.classList.remove("open");
    imageZoomModal.setAttribute("aria-hidden", "true");
    clearImageZoom();
  }

  btnCloseFilePreview?.addEventListener("click", closeFilePreview);
  filePreviewBackdrop?.addEventListener("click", closeFilePreview);

  document.addEventListener("click", (e) => {
    const btn = e.target?.closest?.(".attPreviewBtn");
    if (!btn) return;

    const dataUrl = btn.getAttribute("data-file-url") || "";
    const fileName = btn.getAttribute("data-file-name") || "File";
    const mime = btn.getAttribute("data-file-mime") || "application/octet-stream";

    if (!dataUrl.startsWith("data:")) return;
    openFilePreview(fileName, mime, dataUrl);
  });

    btnCloseImageZoom?.addEventListener("click", closeImageZoom);
  imageZoomBackdrop?.addEventListener("click", closeImageZoom);

  document.addEventListener("click", (e) => {
    const img = e.target?.closest?.(".body img, .fwdInner img, .quotePreview img");
    if (!img) return;

    const src = img.getAttribute("src") || "";
    const alt = img.getAttribute("alt") || "Image preview";

    if (!src) return;
    openImageZoom(src, alt);
  });

  if (pop && list && title) {
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
      title.textContent = fmt(TXT.reactionListTitle, {
      emoji,
      count,
      s: count === "1" ? "" : "s",
      reactionType,
    });
      list.innerHTML = "";

      if (!names.length) {
        const row = document.createElement("div");
        row.className = "tItem";

        const av = document.createElement("div");
        av.className = "pAvatar";
        av.textContent = emoji;

        const nm = document.createElement("div");
        nm.className = "pName";
        nm.textContent = TXT.noParticipantNames;

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
  }

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      if (pop) pop.style.display = "none";
      closeDrawer();
      closeFilePreview();
      closeImageZoom();
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

function supportsFolderExport() {
  return typeof window !== "undefined" && "showDirectoryPicker" in window;
}

function buildBatchFolderName() {
  return `Teams_Export_${formatFileTimestampYYYYMMDDHHMMSS(new Date())}`;
}

async function ensureExportDirectoryPicked() {
  if (!supportsFolderExport()) return null;

  if (_exportDirectoryHandle) {
    return _exportDirectoryHandle;
  }

  const dirHandle = await window.showDirectoryPicker({
    mode: "readwrite",
    startIn: "downloads",
  });

  _exportDirectoryHandle = dirHandle;
  return dirHandle;
}

async function ensureBatchFolderHandle(totalQueuedCount) {
  if (!supportsFolderExport()) return null;

  const rootDir = await ensureExportDirectoryPicked();
  if (!rootDir) return null;

  if (totalQueuedCount <= 1) {
    _currentBatchFolderHandle = null;
    _currentBatchFolderName = "";
    return rootDir;
  }

  if (_currentBatchFolderHandle) {
    return _currentBatchFolderHandle;
  }

  const folderName = buildBatchFolderName();
  const subdir = await rootDir.getDirectoryHandle(folderName, { create: true });

  _currentBatchFolderHandle = subdir;
  _currentBatchFolderName = folderName;

  return subdir;
}

async function writeTextFileToDirectory(dirHandle, fileName, text, mime = "text/html;charset=utf-8") {
  const fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();

  await writable.write(new Blob([text], { type: mime }));
  await writable.close();
}

async function saveExportHtmlFile(fileName, text, totalQueuedCount = 1) {
  if (!supportsFolderExport()) {
    downloadTextFile(fileName, text, "text/html;charset=utf-8");
    return {
      mode: "download",
      folderName: "",
    };
  }

  const targetDir = await ensureBatchFolderHandle(totalQueuedCount);
  await writeTextFileToDirectory(targetDir, fileName, text, "text/html;charset=utf-8");

  return {
    mode: "filesystem",
    folderName: _currentBatchFolderName || "",
  };
}

function resetCurrentBatchFolder() {
  _currentBatchFolderHandle = null;
  _currentBatchFolderName = "";
}

function clearLastExportSummary() {
  _lastExportSummary = null;
}

function beginFreshExportRun() {
  _queueItems = [];
  clearLastExportSummary();
  resetCurrentBatchFolder();
}

function formatBytesHumanCompact(bytes) {
  const n = Number(bytes || 0);
  if (!Number.isFinite(n) || n <= 0) return "0 B";

  const gb = 1024 * 1024 * 1024;
  const mb = 1024 * 1024;

  if (n >= gb) {
    const value = n / gb;
    return `${value % 1 === 0 ? value.toFixed(0) : value.toFixed(1)} GB`;
  }

  if (n >= mb) {
    const value = n / mb;
    return `${value % 1 === 0 ? value.toFixed(0) : value.toFixed(1)} MB`;
  }

  return formatBytesHuman(n);
}

function buildSkippedLargeFilesReportText(items, lang) {
  const list = Array.isArray(items) ? items.filter(Boolean) : [];
  if (!list.length) return "";

  const limitLabel = formatBytesHumanCompact(LIMITS.MAX_ATTACH_BYTES_EACH);

  const groups = new Map();

  for (const item of list) {
    const chatTitle =
      String(item?.chatTitle || "").trim() ||
      tFor(lang, "skippedLargeFilesReportUnknownChat");

    if (!groups.has(chatTitle)) {
      groups.set(chatTitle, []);
    }

    groups.get(chatTitle).push(item);
  }

  const lines = [];
  lines.push(tFor(lang, "skippedLargeFilesReportTitle"));
  lines.push("");
  lines.push(
    tFor(lang, "skippedLargeFilesReportIntro", {
      limit: limitLabel,
    })
  );
  lines.push("");

  for (const [chatTitle, groupItems] of groups.entries()) {
    lines.push(
      tFor(lang, "skippedLargeFilesReportChat", {
        title: chatTitle,
      })
    );

    groupItems.forEach((item, idx) => {
      lines.push(
        tFor(lang, "skippedLargeFilesReportItem", {
          index: idx + 1,
          file: item?.fileName || tFor(lang, "attachmentFileFallback"),
          size: formatBytesHumanCompact(item?.sizeBytes || 0),
          sentAt:
            String(item?.sentAtLabel || "").trim() ||
            tFor(lang, "skippedLargeFilesReportUnknownSentAt"),
        })
      );
    });

    lines.push("");
  }

  return lines.join("\n").trim() + "\n";
}

function buildSkippedLargeFilesReportFileName(lang) {
  const base = safeFileName(
    tFor(lang, "skippedLargeFilesReportFileName") || "Skipped_large_files"
  );
  const ts = formatFileTimestampYYYYMMDDHHMMSS(new Date());
  return `${base}_${ts}.txt`;
}

async function savePlainTextFile(fileName, text, totalQueuedCount = 1) {
  if (!supportsFolderExport()) {
    downloadTextFile(fileName, text, "text/plain;charset=utf-8");
    return {
      mode: "download",
      folderName: "",
    };
  }

  const targetDir = await ensureBatchFolderHandle(totalQueuedCount);
  await writeTextFileToDirectory(targetDir, fileName, text, "text/plain;charset=utf-8");

  return {
    mode: "filesystem",
    folderName: _currentBatchFolderName || "",
  };
}

/** =========================
 * EXPORT PIPELINE
 * ========================= */
async function exportChatToOfflineHtml(chat, me, usedFileNames = null, totalQueuedCount = 1) {
  const exportLang = getCurrentLanguage();
  let members = Array.isArray(chat?.members) ? [...chat.members] : [];

  try {
    const fullMembers = await loadChatMembersCached(chat.id);
    members = mergeMembers(fullMembers, members);
  } catch {
    members = mergeMembers(members, []);
  }

  const exportTitle = chatDisplayName(chat, members);
  const initialName = buildExportFileName(chat, me, members);
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

      attachTotalCount: 0,
      attachTotalBytes: 0,

      skippedLargeFiles: [],
      failures: [],
    };

  const fileTokenToMeta = new Map();
  const dtos = [];

  for (const msg of messages) {
    const created = new Date(msg.createdDateTime);
    const dtMs = Date.parse(msg.createdDateTime) || 0;
    const dtStr = isNaN(created.getTime()) ? "(няма дата/час)" : formatDDMMYYYY_HHMMSS(created);

    const sender = extractSender(msg) || "Unknown";
    let bodyHtml = normalizeTeamsHtml(extractBodyHtml(msg));

    const isDeletedNotice = isDeletedMessage(msg, bodyHtml);
    const isSystemEvent = !isDeletedNotice && isSystemEventMessage(msg, bodyHtml);
    const isEmojiOnly = !isSystemEvent && !isDeletedNotice && isEmojiOnlyHtml(bodyHtml);

    if (isDeletedNotice) {
      bodyHtml = `<div class="deletedNoticeText">${esc(tFor(exportLang, "exportDeletedMessage"))}</div>`;
    }

    if (isSystemEvent) {
      const systemText = extractSystemEventText(msg, bodyHtml, members, me);
      bodyHtml = normalizeSystemEventHtml(systemText, exportLang);
    }

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
    forwardedAttachments = await enrichForwardedAuthors(forwardedAttachments, members, me, stats);

    const { quotes: quotesFromAttachments, fileAttachments } =
      splitQuotesFromAttachments(msg.attachments || []);

    stats.attachTotalCount += fileAttachments.length;
    stats.attachTotalBytes += computeAttachmentsTotalSizeBytes(fileAttachments);

        bodyHtml = tokenizeFileAttachments(
          bodyHtml,
          fileAttachments,
          fileTokenToMeta,
          {
            chatTitle: exportTitle,
            messageCreatedDateTime: msg.createdDateTime || "",
            messageCreatedLabel: dtStr,
          }
        );

    dtos.push({
      kind: "message",
      mid: msg.id,
      sender,
      dtStr,
      dtMs,
      isMe: isSystemEvent || isDeletedNotice ? false : isFromMe(msg, me?.id),
      isSystemEvent,
      isDeletedNotice,
      isEmojiOnly,
      bodyHtml,
      replyToId: isSystemEvent || isDeletedNotice ? "" : (msg.replyToId || ""),
      reactionsHtml: isSystemEvent || isDeletedNotice ? "" : buildReactionsFromGraph(msg, members, me),
      forwardHtml: "",
      quoteHtml: "",
      _quotesFromAttachments: isSystemEvent || isDeletedNotice ? [] : quotesFromAttachments,
      _forwardedAttachments: isSystemEvent || isDeletedNotice ? [] : forwardedAttachments,
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

    dto.quoteHtml = renderQuoteStack(quotes, exportLang);

    const forwarded =
      (dto._forwardedAttachments && dto._forwardedAttachments[0]) ||
      extractForwardFromBody(dto._rawBodyHtml);

    dto.forwardHtml = renderForwardBlock(forwarded, exportLang);
  }

  const tokenToDataUrl = await embedFileAttachments(fileTokenToMeta, stats);

  for (const dto of dtos) {
    dto.bodyHtml = replaceFileTokensOffline(
      dto.bodyHtml,
      tokenToDataUrl,
      fileTokenToMeta,
      stats,
      exportLang
    );

  }

  const archiveRange = getArchiveRange(dtos);
  const html = buildHtml(dtos, stats, exportTitle, archiveRange, members, me, exportLang);

  const saveInfo = await saveExportHtmlFile(outName, html, totalQueuedCount);

  return {
    outName,
    stats,
    items: dtos.length,
    archiveRange,
    saveMode: saveInfo?.mode || "download",
    batchFolderName: saveInfo?.folderName || "",
  };
}

/** =========================
 * QUEUE HELPERS
 * ========================= */
function getQueueStatusLabel(status) {
  if (status === "queued") return t("statusQueued");
  if (status === "running") return t("statusRunning");
  if (status === "done") return t("statusDone");
  if (status === "failed") return t("statusFailed");
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
    log(t("pleaseLoginFirst"));
    return;
  }

  const queued = _queueItems.filter((x) => x.status === "queued");
  if (!queued.length) {
    log(t("noQueuedChats"));
    return;
  }

    const batchSkippedLargeFiles = [];

  _queueRunning = true;
  resetCurrentBatchFolder();
  renderChats();
  setCounts();
  setBusy(true);

  appendLogLine(t("queueStart", { count: queued.length }));

  if (supportsFolderExport()) {
    try {
      await ensureBatchFolderHandle(queued.length);

      if (queued.length > 1 && _currentBatchFolderName) {
        appendLogLine(t("exportFolderSubfolder", { folder: _currentBatchFolderName }));
      } else {
        appendLogLine(t("exportFolderSelected"));
      }
    } catch (e) {
      _queueRunning = false;
      resetCurrentBatchFolder();
      removeQueuedItemsThatNeverStarted();
      renderChats();
      setCounts();
      setBusy(false);

      const msg = normalizeErrorMessage(e).toLowerCase();

      if (
        msg.includes("aborterror") ||
        msg.includes("the user aborted a request") ||
        msg.includes("user cancelled") ||
        msg.includes("user canceled")
      ) {
        appendLogLine(t("exportFolderCancelled"));
      } else {
        appendLogLine(t("exportFolderFailed", { message: normalizeErrorMessage(e) }));
      }

      return;
    }
  } else {
        appendLogLine(t("fsApiFallback"));
        appendLogLine(t("allowMultipleDownloads"));
  }

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
        appendLogLine(
          t("queueItemFailed", {
            title: item.title,
            details: formatFailureForLog(item),
          })
        );
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
      appendLogLine(
        t("queueItemRunning", {
          title: item.title,
        })
      );

      try {
        const res = await exportChatToOfflineHtml(chat, _me, usedFileNames, queued.length);
        item.status = "done";
        item.result = res;
        item.finishedAt = new Date().toISOString();

        if (Array.isArray(res?.stats?.skippedLargeFiles) && res.stats.skippedLargeFiles.length) {
          batchSkippedLargeFiles.push(...res.stats.skippedLargeFiles);
        }

        appendLogLine(
          t("queueItemDone", {
            title: item.title,
            file: res.outName,
            items: res.items,
            period: res.archiveRange?.label || "—",
          })
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

        appendLogLine(
          t("queueItemFailed", {
            title: item.title,
            details: formatFailureForLog(item),
          })
        );
        setBusy(true);
      }

      renderChats();
      setCounts();
      await sleep(180);
    }

        const qc = getQueueSummaryCounts();
    appendLogLine(
      t("queueFinishedSummary", {
        done: qc.done,
        failed: qc.failed,
        total: _queueItems.length,
      })
    );

    if (batchSkippedLargeFiles.length) {
      const reportLang = getCurrentLanguage();
      const reportText = buildSkippedLargeFilesReportText(batchSkippedLargeFiles, reportLang);
      const reportFileName = buildSkippedLargeFilesReportFileName(reportLang);

      await savePlainTextFile(reportFileName, reportText, queued.length);

      appendLogLine(
        `INFO: skipped large files report generated -> ${reportFileName} (${batchSkippedLargeFiles.length} item(s))`
      );
    }
  } finally {
    _queueRunning = false;
    resetCurrentBatchFolder();
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
      `${t("chatStarted")}: ${started} • ${t("lastMessage")}: ${lastWritten}${queueItem ? ` • ${t("queueLabel")}: ${getQueueStatusLabel(queueItem.status)}` : ""}`;

    main.appendChild(title);
    main.appendChild(meta);

    if (queueItem?.status === "failed") {
      const err1 = document.createElement("div");
      err1.className = "chatError";
      err1.textContent = `${t("reason")}: ${queueItem.errorHuman || queueItem.error || t("unknownError")}`;

      const err2 = document.createElement("div");
      err2.className = "chatErrorMeta";
      err2.textContent = `${t("stage")}: ${queueItem.errorStage || "unknown"} • ${t("code")}: ${queueItem.errorCode || "unknown"}`;

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
    log(t("pleaseLoginFirst"));
    return;
  }

  if (_selectedChatIds.size === 0) {
    log(t("selectAtLeastOneChat"));
    return;
  }

  markExportStartedForSession();
  beginFreshExportRun();

  const info = buildQueueFromSelection();
  renderChats();
  setCounts();

  appendLogLine(t("queuePrepared", {
    added: info.added,
    totalSelected: info.totalSelected,
  }));

  await runExportQueue();
}

async function onQueueAllClick() {
  if (!_me) {
    log(t("pleaseLoginFirst"));
    return;
  }

  if (_chats.length === 0) {
    log(t("loadChatsFirst"));
    return;
  }

  markExportStartedForSession();
  beginFreshExportRun();

  const info = buildQueueFromAllChats();
  renderChats();
  setCounts();

  appendLogLine(t("queueAllPrepared", {
    added: info.added,
    total: _chats.length,
  }));

  await runExportQueue();
}

async function onRetryFailedOnlyClick() {
  if (!_me) {
    log(t("pleaseLoginFirst"));
    return;
  }

  const failedCount = getQueueSummaryCounts().failed;
  if (!failedCount) {
    log(t("noFailedChats"));
    return;
  }

  markExportStartedForSession();

  const changed = requeueFailedOnly();
  renderChats();
  setCounts();

  appendLogLine(t("retryFailedPrepared", {
    count: changed,
  }));

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
      appendLogLine(t("loginButtonClicked"));
      await login();
    } catch (e) {
      console.error(e);
      const msg = normalizeErrorMessage(e);
      appendLogLine(`${t("loginError")}: ${msg}`);
      alert(`${t("loginError")}:\n${msg}`);
    } finally {
      setBusy(false);
      setCounts();
    }
  });

  ui.btnChats?.addEventListener("click", async () => {
    try {
      if (!_me) {
        log(t("pleaseLoginFirst"));
        return;
      }

      setBusy(true);
      log(t("loadingChats"));

      _chats = await loadMyChats();

      const validIds = new Set(_chats.map((c) => c.id));
      _selectedChatIds = new Set([..._selectedChatIds].filter((id) => validIds.has(id)));
      _queueItems = _queueItems.filter((x) => validIds.has(x.chatId));

      renderChats();
      setCounts();

      log({ chats: _chats.length, sample: _chats.slice(0, 3) });
      appendLogLine(t("chatsLoaded", { count: _chats.length }));
    } catch (e) {
      console.error(e);
      log(`${t("loadChatsFailed")}:\n${normalizeErrorMessage(e)}`);
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
      log(`${t("exportQueueFailed")}:\n${normalizeErrorMessage(e)}`);
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
        log(`${t("queueAllFailed")}:\n${normalizeErrorMessage(e)}`);
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
        log(`${t("retryFailedError")}:\n${normalizeErrorMessage(e)}`);
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
  refreshUiRefs();
  populateLanguageSelector();

  if (!assertUiIds()) return;

  ensureExtraButtons();
  refreshUiRefs();
  bindLanguageSelector();
  bindUiEvents();
  applyLanguageToUi();
  showLoginView();

  try {
    setBusy(true);
    log(t("initializingAuth"));

    const acc = await initAuth();

    if (acc) {
      log("Signed in. Loading /me…");
      const data = await loadMe();
      _me = {
        id: data.id,
        displayName: data.displayName,
        userPrincipalName: data.userPrincipalName,
        mail: data.mail || "",
      };

      if (!_loginAtMs) recordAuthenticatedSession({ reset: true });
      else recordAuthenticatedSession({ reset: false });

      showAppView();
      appendLogLine(t("readySignedInAs", {
        name: _me.displayName || _me.userPrincipalName || _me.id,
      }));
    } else {
      _me = null;
      clearAppSessionMeta();
      showLoginView();
      log(t("notSignedInUseButton"));
    }

    setCounts();
    applyLanguageToUi();
  } catch (e) {
    console.error(e);

    if (isMsalNoTokenRequestCacheError(e)) {
      console.warn("[MSAL] Ignored no_token_request_cache_error during app init:", e);
      _me = null;
      clearAppSessionMeta();
      showLoginView();
      setCounts();
      applyLanguageToUi();
      return;
    }

    showLoginView();
    log(`${t("initFailed")}:\n${normalizeErrorMessage(e)}`);
    alert(`${t("initFailed")}:\n${normalizeErrorMessage(e)}`);
  } finally {
    setBusy(false);
    setCounts();
    applyLanguageToUi();
  }
}

doInit();