require("dotenv").config();
const express = require("express");
const fs = require("fs");
const path = require("path");
const {
  createHmac,
  createPublicKey,
  timingSafeEqual,
  verify: verifySignature,
} = require("crypto");
const msal = require("@azure/msal-node");
const { Client } = require("@microsoft/microsoft-graph-client");

const app = express();
app.use(express.static("public"));
app.use(express.json({ limit: "256kb" }));

const AUTH_AUDIENCE = process.env.AUTH_AUDIENCE || process.env.CLIENT_ID;
const AUTH_ALLOWED_TENANTS = String(
  process.env.AUTH_ALLOWED_TENANTS || process.env.TENANT_ID || ""
)
  .split(",")
  .map((value) => value.trim())
  .filter(Boolean);
const AUTH_ALLOW_ANY_TENANT = AUTH_ALLOWED_TENANTS.includes("*");
const DEV_ACCESS_KEY = process.env.DEV_ACCESS_KEY || "";
const ALLOW_LOCALHOST_BROWSER = process.env.ALLOW_LOCALHOST_BROWSER === "true";
const AUTH_DEBUG = process.env.AUTH_DEBUG === "true";
const FEEDBACK_ENABLED = process.env.FEEDBACK_ENABLED === "true";
const FEEDBACK_DEV_KEY = process.env.FEEDBACK_DEV_KEY || "";
const ADMIN_PIN = String(process.env.ADMIN_PIN || "").trim();
const ADMIN_PIN_ENABLED = /^\d{6}$/.test(ADMIN_PIN);
const ADMIN_SESSION_TTL_MIN = Math.max(
  15,
  Number.parseInt(process.env.ADMIN_SESSION_TTL_MIN || "240", 10) || 240
);
const ADMIN_SESSION_SECRET =
  process.env.ADMIN_SESSION_SECRET || process.env.CLIENT_SECRET || "local-admin-session-secret";
const ADMIN_COOKIE_NAME = "planner_admin_session";
const FEEDBACK_EXPORT_COOKIE_NAME = "planner_feedback_export_session";
const FEEDBACK_EXPORT_SESSION_TTL_MIN = Math.max(
  15,
  Number.parseInt(process.env.FEEDBACK_EXPORT_SESSION_TTL_MIN || "240", 10) || 240
);
const JWKS_CACHE_TTL_MS = 60 * 60 * 1000;
const jwksCache = new Map();
const DATA_DIR = path.join(__dirname, "data");
const FEEDBACK_FILE = path.join(DATA_DIR, "feedback.json");
const DASHBOARD_SELECTION_FILE = path.join(DATA_DIR, "dashboard-selection.json");
const CHANNEL_TEMPLATE_MAP_FILE = path.join(DATA_DIR, "channel-template-map.json");
const DEFAULT_SCOPE_TEMPLATE_ID = "default";

if (!ADMIN_PIN && AUTH_DEBUG) {
  console.warn("[admin] ADMIN_PIN is not configured; admin panel unlock will be disabled.");
}

function parseCookies(req) {
  const raw = String(req.headers.cookie || "");
  if (!raw) return {};
  return raw.split(";").reduce((acc, part) => {
    const [name, ...rest] = part.trim().split("=");
    if (!name) return acc;
    acc[name] = decodeURIComponent(rest.join("="));
    return acc;
  }, {});
}

function getAdminCookiePolicy(req) {
  const forwardedProto = String(req.headers["x-forwarded-proto"] || "")
    .split(",")
    .map((value) => value.trim().toLowerCase());
  const host = String(req.headers.host || "").toLowerCase();
  const isLocalHost =
    host.startsWith("localhost") ||
    host.startsWith("127.0.0.1") ||
    host.startsWith("[::1]");

  if (isLocalHost && !req.secure && !forwardedProto.includes("https")) {
    return {
      sameSite: "Lax",
      secure: false,
    };
  }

  return {
    sameSite: "None",
    secure: true,
  };
}

function signValue(value) {
  return createHmac("sha256", ADMIN_SESSION_SECRET).update(value).digest("base64url");
}

function safeEquals(a, b) {
  const left = Buffer.from(String(a || ""));
  const right = Buffer.from(String(b || ""));
  if (left.length !== right.length) return false;
  return timingSafeEqual(left, right);
}

function createAdminSessionToken(payload) {
  const encoded = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const signature = signValue(encoded);
  return `${encoded}.${signature}`;
}

function verifyAdminSessionToken(token) {
  const [encoded, signature] = String(token || "").split(".");
  if (!encoded || !signature) return null;
  if (!safeEquals(signature, signValue(encoded))) return null;

  try {
    const payload = JSON.parse(Buffer.from(encoded, "base64url").toString("utf8"));
    if (!payload?.exp || payload.exp <= Date.now()) return null;
    return payload;
  } catch {
    return null;
  }
}

function requireAdminSession(req, res, next) {
  const cookies = parseCookies(req);
  const token = cookies[ADMIN_COOKIE_NAME];
  const session = verifyAdminSessionToken(token);
  if (!session) {
    return res.status(403).json({ error: "Admin session required" });
  }
  req.adminSession = session;
  return next();
}

function requireFeedbackExportSession(req, res, next) {
  const cookies = parseCookies(req);
  const token = cookies[FEEDBACK_EXPORT_COOKIE_NAME];
  const session = verifyAdminSessionToken(token);
  if (!session) {
    return res.status(403).json({ error: "Feedback export session required" });
  }
  req.feedbackExportSession = session;
  return next();
}

function requireFeedbackDevAccess(req, res, next) {
  if (!FEEDBACK_DEV_KEY) {
    return res.status(403).json({ error: "Feedback dev access is disabled" });
  }
  const provided = String(req.headers["x-feedback-dev-key"] || "");
  if (!safeEquals(provided, FEEDBACK_DEV_KEY)) {
    return res.status(403).json({ error: "Forbidden" });
  }
  return next();
}

async function readJsonFile(filePath, fallbackValue) {
  try {
    const raw = await fs.promises.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch {
    return fallbackValue;
  }
}

async function writeJsonFile(filePath, value) {
  await fs.promises.mkdir(path.dirname(filePath), { recursive: true });
  await fs.promises.writeFile(filePath, JSON.stringify(value, null, 2));
}

async function readFeedbackEntries() {
  const data = await readJsonFile(FEEDBACK_FILE, []);
  return Array.isArray(data) ? data : [];
}

async function writeFeedbackEntries(entries) {
  await writeJsonFile(FEEDBACK_FILE, entries);
}

function normalizeSelection(raw) {
  const selectedPlanIds = Array.isArray(raw?.selectedPlanIds)
    ? [...new Set(raw.selectedPlanIds.map((id) => String(id).trim()).filter(Boolean))]
    : [];
  return {
    selectedPlanIds,
    includeAll: Boolean(raw?.includeAll),
    updatedAt: raw?.updatedAt || null,
  };
}

function normalizeTemplateName(value, fallback = "Default Scope") {
  const name = String(value || "").trim();
  return name || fallback;
}

function createScopeTemplateId() {
  return `scope-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function normalizeScopeTemplate(raw, fallbackName = "Scope Template") {
  const normalizedSelection = normalizeSelection(raw);
  const id = String(raw?.id || createScopeTemplateId()).trim();
  return {
    id: id || createScopeTemplateId(),
    name: normalizeTemplateName(raw?.name, fallbackName),
    includeAll: normalizedSelection.includeAll,
    selectedPlanIds: normalizedSelection.selectedPlanIds,
    updatedAt: normalizedSelection.updatedAt,
  };
}

function normalizeTemplateStore(raw) {
  const templates = Array.isArray(raw?.templates)
    ? raw.templates.map((item, idx) =>
        normalizeScopeTemplate(item, idx === 0 ? "Default Scope" : `Template ${idx + 1}`)
      )
    : [];

  if (!templates.length) {
    templates.push(
      normalizeScopeTemplate(
        {
          id: DEFAULT_SCOPE_TEMPLATE_ID,
          name: "Default Scope",
          includeAll: true,
          selectedPlanIds: [],
          updatedAt: null,
        },
        "Default Scope"
      )
    );
  }

  const activeTemplateId = String(raw?.activeTemplateId || "").trim();
  const activeExists = templates.some((template) => template.id === activeTemplateId);

  return {
    activeTemplateId: activeExists ? activeTemplateId : templates[0].id,
    templates,
  };
}

async function readDashboardTemplateStore() {
  const data = await readJsonFile(DASHBOARD_SELECTION_FILE, null);
  if (!data) {
    return normalizeTemplateStore(null);
  }

  if (Array.isArray(data?.templates)) {
    return normalizeTemplateStore(data);
  }

  // Legacy single-selection format migration.
  const legacy = normalizeSelection(data);
  return normalizeTemplateStore({
    activeTemplateId: DEFAULT_SCOPE_TEMPLATE_ID,
    templates: [
      {
        id: DEFAULT_SCOPE_TEMPLATE_ID,
        name: "Default Scope",
        includeAll: legacy.includeAll,
        selectedPlanIds: legacy.selectedPlanIds,
        updatedAt: legacy.updatedAt,
      },
    ],
  });
}

async function writeDashboardTemplateStore(store) {
  const normalized = normalizeTemplateStore(store);
  await writeJsonFile(DASHBOARD_SELECTION_FILE, normalized);
  return normalized;
}

function getActiveScopeTemplate(store) {
  return (
    store.templates.find((template) => template.id === store.activeTemplateId) || store.templates[0]
  );
}

async function readDashboardSelection() {
  const store = await readDashboardTemplateStore();
  const active = getActiveScopeTemplate(store);
  return {
    templateId: active.id,
    templateName: active.name,
    selectedPlanIds: active.selectedPlanIds,
    includeAll: active.includeAll,
    updatedAt: active.updatedAt,
  };
}

async function writeDashboardSelection(selection) {
  const store = await readDashboardTemplateStore();
  const activeIndex = store.templates.findIndex((template) => template.id === store.activeTemplateId);
  const targetIndex = activeIndex >= 0 ? activeIndex : 0;
  const current = store.templates[targetIndex] || normalizeScopeTemplate({}, "Default Scope");
  const normalized = normalizeSelection(selection);
  const updatedTemplate = {
    ...current,
    includeAll: normalized.includeAll,
    selectedPlanIds: normalized.selectedPlanIds,
    updatedAt: new Date().toISOString(),
  };
  store.templates[targetIndex] = updatedTemplate;
  store.activeTemplateId = updatedTemplate.id;
  await writeDashboardTemplateStore(store);
  return normalized;
}

async function listDashboardTemplates() {
  const store = await readDashboardTemplateStore();
  return {
    activeTemplateId: store.activeTemplateId,
    templates: store.templates,
  };
}

async function createDashboardTemplate(name, sourceSelection) {
  const store = await readDashboardTemplateStore();
  const baseSelection = normalizeSelection(sourceSelection);
  const template = normalizeScopeTemplate(
    {
      id: createScopeTemplateId(),
      name,
      includeAll: baseSelection.includeAll,
      selectedPlanIds: baseSelection.selectedPlanIds,
      updatedAt: new Date().toISOString(),
    },
    "Scope Template"
  );
  store.templates.push(template);
  store.activeTemplateId = template.id;
  const saved = await writeDashboardTemplateStore(store);
  return {
    activeTemplateId: saved.activeTemplateId,
    template,
  };
}

async function activateDashboardTemplate(templateId) {
  const store = await readDashboardTemplateStore();
  const id = String(templateId || "").trim();
  const exists = store.templates.some((template) => template.id === id);
  if (!exists) return null;
  store.activeTemplateId = id;
  const saved = await writeDashboardTemplateStore(store);
  return getActiveScopeTemplate(saved);
}

async function deleteDashboardTemplate(templateId) {
  const store = await readDashboardTemplateStore();
  if (store.templates.length <= 1) {
    return { error: "At least one template is required" };
  }

  const id = String(templateId || "").trim();
  const nextTemplates = store.templates.filter((template) => template.id !== id);
  if (nextTemplates.length === store.templates.length) {
    return { error: "Template not found" };
  }

  store.templates = nextTemplates;
  if (!store.templates.some((template) => template.id === store.activeTemplateId)) {
    store.activeTemplateId = store.templates[0].id;
  }
  const saved = await writeDashboardTemplateStore(store);
  return {
    activeTemplateId: saved.activeTemplateId,
    templates: saved.templates,
  };
}

async function fetchGroupPlans(client) {
  const groupId = process.env.GROUP_ID;
  const plansRes = await client.api(`/groups/${groupId}/planner/plans`).get();
  return plansRes.value || [];
}

function normalizeChannelTemplateMap(raw) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return {};
  return Object.entries(raw).reduce((acc, [channelId, templateId]) => {
    const normalizedChannelId = String(channelId || "").trim();
    const normalizedTemplateId = String(templateId || "").trim();
    if (!normalizedChannelId || !normalizedTemplateId) return acc;
    acc[normalizedChannelId] = normalizedTemplateId;
    return acc;
  }, {});
}

async function readChannelTemplateMap() {
  const data = await readJsonFile(CHANNEL_TEMPLATE_MAP_FILE, {});
  return normalizeChannelTemplateMap(data);
}

async function writeChannelTemplateMap(value) {
  const normalized = normalizeChannelTemplateMap(value);
  await writeJsonFile(CHANNEL_TEMPLATE_MAP_FILE, normalized);
  return normalized;
}

function getRequestChannelId(req) {
  return String(req?.headers?.["x-teams-channel-id"] || "").trim();
}

async function resolveDashboardScope(req) {
  const store = await readDashboardTemplateStore();
  const channelId = getRequestChannelId(req);
  const channelTemplateMap = await readChannelTemplateMap();

  let activeTemplate = getActiveScopeTemplate(store);
  let resolvedBy = "active-template";

  if (channelId) {
    const mappedTemplateId = channelTemplateMap[channelId];
    if (mappedTemplateId) {
      const mappedTemplate = store.templates.find((template) => template.id === mappedTemplateId);
      if (mappedTemplate) {
        activeTemplate = mappedTemplate;
        resolvedBy = "channel-binding";
      }
    }
  }

  return {
    includeAll: activeTemplate.includeAll,
    selectedPlanIds: activeTemplate.selectedPlanIds,
    selectedCount: activeTemplate.selectedPlanIds.length,
    updatedAt: activeTemplate.updatedAt,
    templateId: activeTemplate.id,
    templateName: activeTemplate.name,
    channelId: channelId || null,
    resolvedBy,
  };
}

async function fetchTeamChannels(client) {
  const groupId = process.env.GROUP_ID;
  const response = await client.api(`/teams/${groupId}/channels`).get();
  return Array.isArray(response?.value) ? response.value : [];
}

function applyScopeToPlans(plans, scope) {
  if (scope.includeAll) return plans;
  const selected = new Set(scope.selectedPlanIds);
  return plans.filter((plan) => selected.has(plan.id));
}

function getScopeCacheKey(prefix, scope) {
  if (scope.includeAll) return `${prefix}:all`;
  const ids = [...scope.selectedPlanIds].sort().join("|") || "none";
  return `${prefix}:selected:${ids}`;
}

function toCsvRow(cells) {
  return cells
    .map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`)
    .join(",");
}

function createFeedbackCsv(entries) {
  const header = toCsvRow([
    "id",
    "type",
    "message",
    "contact",
    "createdAt",
    "sourceHost",
    "userHint",
  ]);
  const rows = entries.map((entry) =>
    toCsvRow([
      entry.id,
      entry.type,
      entry.message,
      entry.contact,
      entry.createdAt,
      entry.sourceHost,
      entry.userHint,
    ])
  );
  return [header, ...rows].join("\n");
}

class AuthError extends Error {
  constructor(code, message) {
    super(message);
    this.name = "AuthError";
    this.code = code;
  }
}

function authLog(level, event, details = {}) {
  const payload = Object.entries(details)
    .map(([key, value]) => `${key}=${value}`)
    .join(" ");
  const line = payload ? `[auth] ${event} ${payload}` : `[auth] ${event}`;

  if (level === "error") console.error(line);
  else if (level === "warn") console.warn(line);
  else if (AUTH_DEBUG) console.log(line);
}

function createRequestTraceId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function base64UrlToBuffer(value) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized + "=".repeat((4 - (normalized.length % 4)) % 4);
  return Buffer.from(padded, "base64");
}

function parseJwt(token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3) throw new AuthError("malformed_jwt", "Malformed JWT");

  const [encodedHeader, encodedPayload, encodedSignature] = parts;
  const header = JSON.parse(base64UrlToBuffer(encodedHeader).toString("utf8"));
  const payload = JSON.parse(base64UrlToBuffer(encodedPayload).toString("utf8"));

  return {
    header,
    payload,
    signingInput: `${encodedHeader}.${encodedPayload}`,
    signature: base64UrlToBuffer(encodedSignature),
  };
}

function isTenantAllowed(tid) {
  if (!tid) return false;
  if (AUTH_ALLOW_ANY_TENANT) return true;
  return AUTH_ALLOWED_TENANTS.includes(tid);
}

function hasExpectedAudience(payload) {
  if (!AUTH_AUDIENCE) return false;
  if (typeof payload?.aud === "string") return payload.aud === AUTH_AUDIENCE;
  if (Array.isArray(payload?.aud)) return payload.aud.includes(AUTH_AUDIENCE);
  return false;
}

function hasAllowedIssuer(payload) {
  const tid = payload?.tid;
  const issuer = String(payload?.iss || "").toLowerCase();
  if (!tid || !issuer) return false;

  return (
    issuer === `https://login.microsoftonline.com/${tid}/v2.0` ||
    issuer === `https://sts.windows.net/${tid}/`
  );
}

async function getSigningJwk(tid, kid) {
  const cacheKey = `${tid}:${kid}`;
  const cached = jwksCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now()) return cached.jwk;

  const keysUrl = `https://login.microsoftonline.com/${encodeURIComponent(
    tid
  )}/discovery/v2.0/keys`;
  const response = await fetch(keysUrl);
  if (!response.ok) {
    throw new Error(`JWKS fetch failed: ${response.status}`);
  }

  const body = await response.json();
  (body.keys || []).forEach((jwk) => {
    if (!jwk?.kid) return;
    jwksCache.set(`${tid}:${jwk.kid}`, {
      jwk,
      expiresAt: Date.now() + JWKS_CACHE_TTL_MS,
    });
  });

  const fresh = jwksCache.get(cacheKey);
  if (!fresh) throw new Error("Signing key not found");
  return fresh.jwk;
}

async function validateBearerToken(token) {
  const { header, payload, signingInput, signature } = parseJwt(token);

  if (header?.alg !== "RS256") {
    throw new AuthError("invalid_alg", "Unsupported token algorithm");
  }
  if (!header?.kid) throw new AuthError("missing_kid", "Token missing kid");
  if (!payload?.tid) throw new AuthError("missing_tid", "Token missing tenant id");
  if (!isTenantAllowed(payload.tid)) throw new AuthError("tenant_blocked", "Tenant not allowed");
  if (!hasExpectedAudience(payload)) throw new AuthError("invalid_audience", "Invalid audience");
  if (!hasAllowedIssuer(payload)) throw new AuthError("invalid_issuer", "Invalid issuer");

  const nowSec = Math.floor(Date.now() / 1000);
  if (typeof payload.exp !== "number" || payload.exp <= nowSec) {
    throw new AuthError("token_expired", "Token expired");
  }
  if (typeof payload.nbf === "number" && payload.nbf > nowSec) {
    throw new AuthError("token_not_yet_valid", "Token not active yet");
  }

  let jwk;
  try {
    jwk = await getSigningJwk(payload.tid, header.kid);
  } catch (err) {
    throw new AuthError("jwks_fetch_failed", err.message);
  }

  const publicKey = createPublicKey({ key: jwk, format: "jwk" });
  const isSignatureValid = verifySignature(
    "RSA-SHA256",
    Buffer.from(signingInput),
    publicKey,
    signature
  );

  if (!isSignatureValid) throw new AuthError("invalid_signature", "Invalid token signature");
  return payload;
}

function requestIsLocal(req) {
  const ip = req.ip || req.connection?.remoteAddress || "";
  return (
    ip === "127.0.0.1" ||
    ip === "::1" ||
    ip.startsWith("::ffff:127.") ||
    ip.startsWith("::ffff:10.")
  );
}

async function requireApiAuth(req, res, next) {
  const traceId = createRequestTraceId();
  req.authTraceId = traceId;
  res.setHeader("x-auth-trace-id", traceId);
  const authHeader = String(req.headers.authorization || "");
  const path = req.originalUrl || req.url;
  authLog("debug", "request_received", {
    traceId,
    path,
    hasBearer: authHeader.toLowerCase().startsWith("bearer "),
    hasDevKey: Boolean(req.headers["x-dev-access-key"]),
  });

  if (authHeader.toLowerCase().startsWith("bearer ")) {
    const token = authHeader.slice(7).trim();
    try {
      const claims = await validateBearerToken(token);
      req.auth = {
        mode: "bearer",
        tenantId: claims.tid,
        userObjectId: claims.oid || null,
        username: claims.preferred_username || claims.upn || null,
      };
      authLog("debug", "bearer_validation_pass", {
        traceId,
        tid: claims.tid,
        oid: claims.oid || "none",
      });
      return next();
    } catch (err) {
      authLog("warn", "bearer_validation_failed", {
        traceId,
        code: err.code || "unknown",
        reason: err.message,
      });
      return res.status(401).json({ error: "Unauthorized" });
    }
  }

  if (DEV_ACCESS_KEY && req.headers["x-dev-access-key"] === DEV_ACCESS_KEY) {
    req.auth = { mode: "dev-key" };
    authLog("debug", "dev_key_pass", { traceId, path });
    return next();
  }

  if (ALLOW_LOCALHOST_BROWSER && requestIsLocal(req)) {
    req.auth = { mode: "localhost" };
    authLog("debug", "localhost_bypass_pass", { traceId, path });
    return next();
  }

  authLog("warn", "request_unauthorized", { traceId, path, reason: "no_valid_auth" });
  return res.status(401).json({ error: "Unauthorized" });
}

const API_CACHE_TTL_MS = 60 * 1000;
const apiCache = new Map();

function getCachedValue(key) {
  const cached = apiCache.get(key);
  if (!cached) return null;
  if (Date.now() - cached.createdAt > API_CACHE_TTL_MS) {
    apiCache.delete(key);
    return null;
  }
  return cached.value;
}

function setCachedValue(key, value) {
  apiCache.set(key, {
    value,
    createdAt: Date.now(),
  });
}

async function withCache(cacheKey, producer) {
  const cached = getCachedValue(cacheKey);
  if (cached) return cached;
  const value = await producer();
  setCachedValue(cacheKey, value);
  return value;
}

// MSAL client credentials config
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getGraphClient() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });

  return Client.init({
    authProvider: (done) => done(null, result.accessToken),
  });
}

// Fetch all pages from a paginated Graph API endpoint
async function fetchAllPages(client, url) {
  let results = [];
  let nextUrl = url;
  while (nextUrl) {
    const res = await client.api(nextUrl).get();
    results.push(...res.value);
    nextUrl = res["@odata.nextLink"] || null;
  }
  return results;
}

// Classify a due date relative to today
function getDueStatus(dueDateTime, percentComplete) {
  if (!dueDateTime) return "noDueDate";
  if (percentComplete === 100) return "onTrack";
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const due = new Date(dueDateTime);
  if (due < now) return "overdue";
  const weekFromNow = new Date(now);
  weekFromNow.setDate(weekFromNow.getDate() + 7);
  if (due <= weekFromNow) return "dueSoon";
  return "onTrack";
}

// Resolve user display names from IDs (graceful fallback)
async function resolveUsers(client, userIds) {
  const userMap = new Map();
  const results = await Promise.allSettled(
    userIds.map((id) =>
      client.api(`/users/${id}`).select("displayName").get()
    )
  );
  userIds.forEach((id, i) => {
    if (results[i].status === "fulfilled") {
      userMap.set(id, results[i].value.displayName);
    } else {
      userMap.set(id, id);
    }
  });
  return userMap;
}

// --- Snapshot helpers for sparkline trends ---
const SNAPSHOT_DIR = path.join(__dirname, "data");
const SNAPSHOT_FILE = path.join(SNAPSHOT_DIR, "snapshots.json");
const SNAPSHOT_INTERVAL_MS = 24 * 60 * 60 * 1000;
const SNAPSHOT_TIME_UTC = process.env.SNAPSHOT_TIME_UTC || "";
let snapshotTimer = null;

function parseSnapshotTimeUtc(value) {
  if (!value) return null;
  const match = String(value).trim().match(/^([01]\d|2[0-3]):([0-5]\d)$/);
  if (!match) return null;
  return {
    hours: Number.parseInt(match[1], 10),
    minutes: Number.parseInt(match[2], 10),
  };
}

function getSnapshotCapturedAtMs(dateKey, dayData) {
  const fromMeta = dayData?.__meta?.capturedAt;
  if (typeof fromMeta === "string") {
    const parsed = Date.parse(fromMeta);
    if (!Number.isNaN(parsed)) return parsed;
  }

  const legacy = Date.parse(`${dateKey}T00:00:00.000Z`);
  return Number.isNaN(legacy) ? null : legacy;
}

function getSnapshotPlanEntries(dayData) {
  if (!dayData || typeof dayData !== "object") return [];
  return Object.entries(dayData).filter(
    ([key, stats]) => !key.startsWith("__") && stats && typeof stats === "object"
  );
}

function getLatestSnapshotInfo(snapshots) {
  let latest = null;

  Object.entries(snapshots || {}).forEach(([dateKey, dayData]) => {
    const capturedAtMs = getSnapshotCapturedAtMs(dateKey, dayData);
    if (capturedAtMs === null) return;
    if (!latest || capturedAtMs > latest.capturedAtMs) {
      latest = { dateKey, capturedAtMs };
    }
  });

  return latest;
}

function getNextConfiguredSnapshotUtcDate(hours, minutes, fromDate = new Date()) {
  const next = new Date(fromDate);
  next.setUTCHours(hours, minutes, 0, 0);
  if (next <= fromDate) next.setUTCDate(next.getUTCDate() + 1);
  return next;
}

function scheduleSnapshotTimeout(runAt, modeLabel) {
  if (snapshotTimer) clearTimeout(snapshotTimer);
  const delayMs = Math.max(0, runAt.getTime() - Date.now());

  console.log(`[snapshot] Scheduler mode: ${modeLabel}. Next run at ${runAt.toISOString()}`);

  snapshotTimer = setTimeout(async () => {
    await captureSnapshot();
    await scheduleNextSnapshot();
  }, delayMs);
}

async function scheduleRollingSnapshot() {
  const snapshots = await readSnapshots();
  const latest = getLatestSnapshotInfo(snapshots);

  if (!latest) {
    console.log("[snapshot] No previous snapshot found; capturing on startup.");
    await captureSnapshot();
    const refreshed = await readSnapshots();
    const latestAfterCapture = getLatestSnapshotInfo(refreshed);
    const nextRunAt = new Date((latestAfterCapture?.capturedAtMs || Date.now()) + SNAPSHOT_INTERVAL_MS);
    scheduleSnapshotTimeout(nextRunAt, "rolling-24h");
    return;
  }

  const nextRunAt = new Date(latest.capturedAtMs + SNAPSHOT_INTERVAL_MS);
  if (nextRunAt <= new Date()) {
    console.log("[snapshot] Next rolling snapshot is due now; capturing immediately.");
    await captureSnapshot();
    const refreshed = await readSnapshots();
    const latestAfterCapture = getLatestSnapshotInfo(refreshed);
    const delayedRunAt = new Date((latestAfterCapture?.capturedAtMs || Date.now()) + SNAPSHOT_INTERVAL_MS);
    scheduleSnapshotTimeout(delayedRunAt, "rolling-24h");
    return;
  }

  scheduleSnapshotTimeout(nextRunAt, "rolling-24h");
}

async function scheduleConfiguredSnapshot() {
  const parsed = parseSnapshotTimeUtc(SNAPSHOT_TIME_UTC);
  if (!parsed) {
    console.warn(
      `[snapshot] Invalid SNAPSHOT_TIME_UTC="${SNAPSHOT_TIME_UTC}" (expected HH:mm). Falling back to rolling 24h mode.`
    );
    await scheduleRollingSnapshot();
    return;
  }

  const nextRunAt = getNextConfiguredSnapshotUtcDate(parsed.hours, parsed.minutes);
  scheduleSnapshotTimeout(nextRunAt, `configured-utc (${SNAPSHOT_TIME_UTC})`);
}

async function scheduleNextSnapshot() {
  if (SNAPSHOT_TIME_UTC) {
    await scheduleConfiguredSnapshot();
    return;
  }

  await scheduleRollingSnapshot();
}

async function startSnapshotScheduler() {
  await scheduleNextSnapshot();
}

async function readSnapshots() {
  try {
    const raw = await fs.promises.readFile(SNAPSHOT_FILE, "utf8");
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

async function writeSnapshots(data) {
  await fs.promises.mkdir(SNAPSHOT_DIR, { recursive: true });
  await fs.promises.writeFile(SNAPSHOT_FILE, JSON.stringify(data, null, 2));
}

async function captureSnapshot() {
  const today = new Date().toISOString().split("T")[0];
  const snapshots = await readSnapshots();
  if (snapshots[today]) return; // already captured today

  try {
    const client = await getGraphClient();
    const groupId = process.env.GROUP_ID;
    const plansRes = await client.api(`/groups/${groupId}/planner/plans`).get();

    const capturedAt = new Date().toISOString();
    const dayData = {
      __meta: {
        capturedAt,
      },
    };
    for (const plan of plansRes.value) {
      const tasks = await fetchAllPages(client, `/planner/plans/${plan.id}/tasks`);
      const total = tasks.length;
      const completed = tasks.filter((t) => t.percentComplete === 100).length;
      const overdue = tasks.filter(
        (t) => getDueStatus(t.dueDateTime, t.percentComplete) === "overdue"
      ).length;
      const dueSoon = tasks.filter(
        (t) => getDueStatus(t.dueDateTime, t.percentComplete) === "dueSoon"
      ).length;
      const unassigned = tasks.filter(
        (t) => !t.assignments || Object.keys(t.assignments).length === 0
      ).length;
      dayData[plan.id] = {
        total,
        completed,
        overdue,
        dueSoon,
        unassigned,
        percentComplete: total > 0 ? Math.round((completed / total) * 100) : 0,
      };
    }

    snapshots[today] = dayData;
    await writeSnapshots(snapshots);
    console.log(`Snapshot captured for ${today} at ${capturedAt}`);
  } catch (err) {
    console.error("Snapshot capture failed:", err.message);
  }
}

// Start snapshot scheduler
startSnapshotScheduler().catch((err) => {
  console.error("Snapshot scheduler failed:", err.message);
});

app.get("/auth-config", (_req, res) => {
  res.json({ authDebug: AUTH_DEBUG });
});

app.get("/ui-config", async (req, res) => {
  const scope = await resolveDashboardScope(req);
  res.json({
    feedbackEnabled: FEEDBACK_ENABLED,
    feedbackExportEnabled: FEEDBACK_ENABLED && Boolean(FEEDBACK_DEV_KEY),
    adminEnabled: ADMIN_PIN_ENABLED,
    scope,
  });
});

app.use("/api", requireApiAuth);

app.get("/api/dashboard-scope", async (req, res) => {
  try {
    const scope = await resolveDashboardScope(req);
    res.json(scope);
  } catch (err) {
    console.error("Error fetching dashboard scope:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/admin/unlock", async (req, res) => {
  const pin = String(req.body?.pin || "").trim();
  if (!ADMIN_PIN_ENABLED) {
    return res.status(403).json({ error: "Admin panel is disabled" });
  }
  if (!/^\d{6}$/.test(pin)) {
    return res.status(400).json({ error: "PIN must be 6 digits" });
  }
  if (!safeEquals(pin, ADMIN_PIN)) {
    return res.status(403).json({ error: "Invalid PIN" });
  }

  const now = Date.now();
  const exp = now + ADMIN_SESSION_TTL_MIN * 60 * 1000;
  const token = createAdminSessionToken({
    iat: now,
    exp,
    sub: req.auth?.username || req.auth?.userObjectId || "admin",
  });

  const cookiePolicy = getAdminCookiePolicy(req);
  const cookieParts = [
    `${ADMIN_COOKIE_NAME}=${encodeURIComponent(token)}`,
    "Path=/",
    `Max-Age=${ADMIN_SESSION_TTL_MIN * 60}`,
    "HttpOnly",
    `SameSite=${cookiePolicy.sameSite}`,
  ];
  if (cookiePolicy.secure) cookieParts.push("Secure");
  res.setHeader("Set-Cookie", cookieParts.join("; "));
  return res.json({ success: true, expiresAt: new Date(exp).toISOString() });
});

app.post("/api/admin/logout", (req, res) => {
  const cookiePolicy = getAdminCookiePolicy(req);
  const cookieParts = [
    `${ADMIN_COOKIE_NAME}=`,
    "Path=/",
    "Max-Age=0",
    "HttpOnly",
    `SameSite=${cookiePolicy.sameSite}`,
  ];
  if (cookiePolicy.secure) cookieParts.push("Secure");
  res.setHeader("Set-Cookie", cookieParts.join("; "));
  res.json({ success: true });
});

app.get("/api/admin/session", requireAdminSession, (req, res) => {
  res.json({
    authenticated: true,
    expiresAt: new Date(req.adminSession.exp).toISOString(),
  });
});

app.get("/api/admin/plans", requireAdminSession, async (req, res) => {
  try {
    const client = await getGraphClient();
    const plans = await fetchGroupPlans(client);
    res.json(plans);
  } catch (err) {
    console.error("Error fetching admin plans:", err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/admin/selection", requireAdminSession, async (req, res) => {
  try {
    const selection = await readDashboardSelection();
    res.json(selection);
  } catch (err) {
    console.error("Error fetching dashboard selection:", err);
    res.status(500).json({ error: err.message });
  }
});

app.put("/api/admin/selection", requireAdminSession, async (req, res) => {
  try {
    const includeAll = Boolean(req.body?.includeAll);
    const selectedPlanIds = Array.isArray(req.body?.selectedPlanIds)
      ? req.body.selectedPlanIds
      : [];
    const selection = await writeDashboardSelection({ includeAll, selectedPlanIds });
    apiCache.clear();
    res.json(selection);
  } catch (err) {
    console.error("Error updating dashboard selection:", err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/admin/templates", requireAdminSession, async (_req, res) => {
  try {
    const templates = await listDashboardTemplates();
    res.json(templates);
  } catch (err) {
    console.error("Error fetching dashboard templates:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/admin/templates", requireAdminSession, async (req, res) => {
  try {
    const name = normalizeTemplateName(req.body?.name, "Scope Template");
    const includeAll = Boolean(req.body?.includeAll);
    const selectedPlanIds = Array.isArray(req.body?.selectedPlanIds)
      ? req.body.selectedPlanIds
      : [];
    const created = await createDashboardTemplate(name, {
      includeAll,
      selectedPlanIds,
    });
    apiCache.clear();
    res.json(created);
  } catch (err) {
    console.error("Error creating dashboard template:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/admin/templates/:id/activate", requireAdminSession, async (req, res) => {
  try {
    const template = await activateDashboardTemplate(req.params.id);
    if (!template) {
      return res.status(404).json({ error: "Template not found" });
    }
    apiCache.clear();
    return res.json({
      activeTemplateId: template.id,
      template,
    });
  } catch (err) {
    console.error("Error activating dashboard template:", err);
    return res.status(500).json({ error: err.message });
  }
});

app.delete("/api/admin/templates/:id", requireAdminSession, async (req, res) => {
  try {
    const deletedTemplateId = String(req.params.id || "").trim();
    const result = await deleteDashboardTemplate(req.params.id);
    if (result.error === "Template not found") {
      return res.status(404).json({ error: result.error });
    }
    if (result.error) {
      return res.status(400).json({ error: result.error });
    }

    const channelTemplateMap = await readChannelTemplateMap();
    const nextMap = Object.entries(channelTemplateMap).reduce((acc, [channelId, templateId]) => {
      if (templateId !== deletedTemplateId) acc[channelId] = templateId;
      return acc;
    }, {});
    await writeChannelTemplateMap(nextMap);

    apiCache.clear();
    return res.json(result);
  } catch (err) {
    console.error("Error deleting dashboard template:", err);
    return res.status(500).json({ error: err.message });
  }
});

app.get("/api/admin/channels", requireAdminSession, async (_req, res) => {
  try {
    const client = await getGraphClient();
    const [channels, channelTemplateMap] = await Promise.all([
      fetchTeamChannels(client),
      readChannelTemplateMap(),
    ]);

    const payload = channels.map((channel) => ({
      id: String(channel.id || ""),
      displayName: String(channel.displayName || "").trim() || String(channel.id || ""),
      membershipType: String(channel.membershipType || "standard"),
      templateId: channelTemplateMap[String(channel.id || "")] || null,
    }));

    res.json({ channels: payload });
  } catch (err) {
    console.error("Error fetching team channels:", err);
    res.status(500).json({ error: err.message });
  }
});

app.put("/api/admin/channel-bindings/:channelId", requireAdminSession, async (req, res) => {
  try {
    const channelId = String(req.params.channelId || "").trim();
    if (!channelId) {
      return res.status(400).json({ error: "Channel id is required" });
    }

    const templateId = String(req.body?.templateId || "").trim();
    const channelTemplateMap = await readChannelTemplateMap();

    if (!templateId) {
      delete channelTemplateMap[channelId];
      const saved = await writeChannelTemplateMap(channelTemplateMap);
      apiCache.clear();
      return res.json({ success: true, channelId, templateId: null, bindings: saved });
    }

    const templates = await listDashboardTemplates();
    const exists = templates.templates.some((template) => template.id === templateId);
    if (!exists) {
      return res.status(404).json({ error: "Template not found" });
    }

    channelTemplateMap[channelId] = templateId;
    const saved = await writeChannelTemplateMap(channelTemplateMap);
    apiCache.clear();
    return res.json({ success: true, channelId, templateId, bindings: saved });
  } catch (err) {
    console.error("Error updating channel binding:", err);
    return res.status(500).json({ error: err.message });
  }
});

app.post("/api/feedback/export/unlock", async (req, res) => {
  if (!FEEDBACK_ENABLED) {
    return res.status(404).json({ error: "Feedback is disabled" });
  }
  if (!FEEDBACK_DEV_KEY) {
    return res.status(403).json({ error: "Feedback export is disabled" });
  }

  const key = String(req.body?.key || "").trim();
  if (!key) {
    return res.status(400).json({ error: "Developer key is required" });
  }
  if (!safeEquals(key, FEEDBACK_DEV_KEY)) {
    return res.status(403).json({ error: "Invalid developer key" });
  }

  const now = Date.now();
  const exp = now + FEEDBACK_EXPORT_SESSION_TTL_MIN * 60 * 1000;
  const token = createAdminSessionToken({
    iat: now,
    exp,
    sub: req.auth?.username || req.auth?.userObjectId || "feedback-export",
  });

  const cookiePolicy = getAdminCookiePolicy(req);
  const cookieParts = [
    `${FEEDBACK_EXPORT_COOKIE_NAME}=${encodeURIComponent(token)}`,
    "Path=/",
    `Max-Age=${FEEDBACK_EXPORT_SESSION_TTL_MIN * 60}`,
    "HttpOnly",
    `SameSite=${cookiePolicy.sameSite}`,
  ];
  if (cookiePolicy.secure) cookieParts.push("Secure");
  res.setHeader("Set-Cookie", cookieParts.join("; "));
  return res.json({ success: true, expiresAt: new Date(exp).toISOString() });
});

app.post("/api/feedback/export/logout", (req, res) => {
  const cookiePolicy = getAdminCookiePolicy(req);
  const cookieParts = [
    `${FEEDBACK_EXPORT_COOKIE_NAME}=`,
    "Path=/",
    "Max-Age=0",
    "HttpOnly",
    `SameSite=${cookiePolicy.sameSite}`,
  ];
  if (cookiePolicy.secure) cookieParts.push("Secure");
  res.setHeader("Set-Cookie", cookieParts.join("; "));
  res.json({ success: true });
});

app.get("/api/feedback/export/session", requireFeedbackExportSession, (req, res) => {
  res.json({
    authenticated: true,
    expiresAt: new Date(req.feedbackExportSession.exp).toISOString(),
  });
});

app.get("/api/feedback/export.csv", requireFeedbackExportSession, async (req, res) => {
  try {
    const entries = await readFeedbackEntries();
    const csv = createFeedbackCsv(entries);
    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", 'attachment; filename="feedback.csv"');
    return res.send(csv);
  } catch (err) {
    console.error("Error exporting feedback CSV:", err);
    return res.status(500).json({ error: err.message });
  }
});

app.post("/api/feedback", async (req, res) => {
  if (!FEEDBACK_ENABLED) {
    return res.status(404).json({ error: "Feedback is disabled" });
  }

  const type = req.body?.type === "bug" ? "bug" : "general";
  const message = String(req.body?.message || "").trim();
  const contact = String(req.body?.contact || "").trim();

  if (!message) {
    return res.status(400).json({ error: "Message is required" });
  }

  if (message.length > 4000 || contact.length > 320) {
    return res.status(400).json({ error: "Feedback is too long" });
  }

  try {
    const entries = await readFeedbackEntries();
    const item = {
      id: `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
      type,
      message,
      contact,
      createdAt: new Date().toISOString(),
      sourceHost: req.headers.host || "unknown",
      userHint: req.auth?.username || req.auth?.userObjectId || null,
    };
    entries.push(item);
    await writeFeedbackEntries(entries);
    res.json({ success: true });
  } catch (err) {
    console.error("Error saving feedback:", err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/feedback", requireFeedbackDevAccess, async (req, res) => {
  try {
    const entries = await readFeedbackEntries();
    if (req.query.format === "csv") {
      res.setHeader("Content-Type", "text/csv; charset=utf-8");
      res.setHeader("Content-Disposition", 'attachment; filename="feedback.csv"');
      return res.send(createFeedbackCsv(entries));
    }

    return res.json(entries);
  } catch (err) {
    console.error("Error reading feedback:", err);
    return res.status(500).json({ error: err.message });
  }
});

app.delete("/api/feedback/:id", requireFeedbackDevAccess, async (req, res) => {
  try {
    const entries = await readFeedbackEntries();
    const next = entries.filter((entry) => entry.id !== req.params.id);
    if (next.length === entries.length) {
      return res.status(404).json({ error: "Feedback entry not found" });
    }
    await writeFeedbackEntries(next);
    return res.json({ success: true });
  } catch (err) {
    console.error("Error deleting feedback:", err);
    return res.status(500).json({ error: err.message });
  }
});

// Get all plans for the group
app.get("/api/plans", async (req, res) => {
  try {
    const client = await getGraphClient();
    const allPlans = await fetchGroupPlans(client);
    const scope = await resolveDashboardScope(req);
    const scopedPlans = applyScopeToPlans(allPlans, scope);
    res.json(scopedPlans);
  } catch (err) {
    console.error("Error fetching plans:", err);
    res.status(500).json({ error: err.message });
  }
});

// All-plans overview — now includes unassigned count
app.get("/api/overview", async (req, res) => {
  try {
    const scope = await resolveDashboardScope(req);
    const overview = await withCache(getScopeCacheKey("overview", scope), async () => {
      const client = await getGraphClient();
      const allPlans = await fetchGroupPlans(client);
      const plans = applyScopeToPlans(allPlans, scope);

      return Promise.all(
        plans.map(async (plan) => {
          const tasks = await fetchAllPages(client, `/planner/plans/${plan.id}/tasks`);

          const total = tasks.length;
          const completed = tasks.filter((t) => t.percentComplete === 100).length;
          const inProgress = tasks.filter(
            (t) => t.percentComplete > 0 && t.percentComplete < 100
          ).length;
          const notStarted = tasks.filter((t) => t.percentComplete === 0).length;
          const overdue = tasks.filter(
            (t) => getDueStatus(t.dueDateTime, t.percentComplete) === "overdue"
          ).length;
          const dueSoon = tasks.filter(
            (t) => getDueStatus(t.dueDateTime, t.percentComplete) === "dueSoon"
          ).length;
          const unassigned = tasks.filter(
            (t) => !t.assignments || Object.keys(t.assignments).length === 0
          ).length;

          return {
            planId: plan.id,
            title: plan.title,
            total,
            completed,
            inProgress,
            notStarted,
            overdue,
            dueSoon,
            unassigned,
            percentComplete: total > 0 ? Math.round((completed / total) * 100) : 0,
          };
        })
      );
    });

    res.json(overview);
  } catch (err) {
    console.error("Error fetching overview:", err);
    res.status(500).json({ error: err.message });
  }
});

// Unassigned tasks across all plans
app.get("/api/unassigned", async (req, res) => {
  try {
    const scope = await resolveDashboardScope(req);
    const data = await withCache(getScopeCacheKey("unassigned", scope), async () => {
      const client = await getGraphClient();
      const allPlans = await fetchGroupPlans(client);
      const plans = applyScopeToPlans(allPlans, scope);

      const allUnassigned = [];
      for (const plan of plans) {
        const tasks = await fetchAllPages(client, `/planner/plans/${plan.id}/tasks`);
        tasks
          .filter((t) => !t.assignments || Object.keys(t.assignments).length === 0)
          .forEach((t) =>
            allUnassigned.push({
              id: t.id,
              title: t.title,
              planTitle: plan.title,
              dueDateTime: t.dueDateTime,
              dueStatus: getDueStatus(t.dueDateTime, t.percentComplete),
              status:
                t.percentComplete === 100
                  ? "completed"
                  : t.percentComplete > 0
                  ? "inProgress"
                  : "notStarted",
            })
          );
      }

      return { count: allUnassigned.length, tasks: allUnassigned };
    });

    res.json(data);
  } catch (err) {
    console.error("Error fetching unassigned:", err);
    res.status(500).json({ error: err.message });
  }
});

// Team workload — assignee stats across all plans
app.get("/api/workload", async (req, res) => {
  try {
    const scope = await resolveDashboardScope(req);
    const result = await withCache(getScopeCacheKey("workload", scope), async () => {
      const client = await getGraphClient();
      const allPlans = await fetchGroupPlans(client);
      const plans = applyScopeToPlans(allPlans, scope);

      const allTasks = [];
      for (const plan of plans) {
        const tasks = await fetchAllPages(client, `/planner/plans/${plan.id}/tasks`);
        allTasks.push(...tasks);
      }

      // Collect unique user IDs
      const allUserIds = new Set();
      allTasks.forEach((t) => {
        if (t.assignments) {
          Object.keys(t.assignments).forEach((uid) => allUserIds.add(uid));
        }
      });

      const userMap = await resolveUsers(client, [...allUserIds]);

      // Build workload map
      const workload = {};
      allTasks.forEach((t) => {
        const assigneeIds = t.assignments ? Object.keys(t.assignments) : [];
        const status =
          t.percentComplete === 100
            ? "completed"
            : t.percentComplete > 0
            ? "inProgress"
            : "notStarted";
        const dueStatus = getDueStatus(t.dueDateTime, t.percentComplete);
        const isOverdue = dueStatus === "overdue";

        const names =
          assigneeIds.length > 0
            ? assigneeIds.map((uid) => userMap.get(uid) || uid)
            : ["Unassigned"];

        names.forEach((name) => {
          if (!workload[name]) {
            workload[name] = {
              displayName: name,
              completed: 0,
              inProgress: 0,
              notStarted: 0,
              total: 0,
              overdue: 0,
            };
          }
          workload[name][status]++;
          workload[name].total++;
          if (isOverdue) workload[name].overdue++;
        });
      });

      return Object.values(workload).sort((a, b) => b.total - a.total);
    });

    res.json(result);
  } catch (err) {
    console.error("Error fetching workload:", err);
    res.status(500).json({ error: err.message });
  }
});

// Task details (description/notes) — lazy loaded
app.get("/api/tasks/:taskId/details", async (req, res) => {
  try {
    const client = await getGraphClient();
    const details = await client
      .api(`/planner/tasks/${req.params.taskId}/details`)
      .get();
    res.json({ description: details.description || "" });
  } catch (err) {
    console.error("Error fetching task details:", err);
    res.status(500).json({ error: err.message });
  }
});

// Trend data — last 14 days of snapshots
app.get("/api/trends", async (req, res) => {
  try {
    const scope = await resolveDashboardScope(req);
    const selectedSet = new Set(scope.selectedPlanIds);
    const snapshots = await readSnapshots();
    const dates = Object.keys(snapshots).sort().slice(-14);

    const trends = {};
    dates.forEach((date) => {
      const dayData = snapshots[date];
      getSnapshotPlanEntries(dayData).forEach(([planId, stats]) => {
        if (!scope.includeAll && !selectedSet.has(planId)) return;
        if (!trends[planId]) trends[planId] = [];
        trends[planId].push({ date, percentComplete: stats.percentComplete || 0 });
      });
    });

    res.json(trends);
  } catch (err) {
    console.error("Error fetching trends:", err);
    res.status(500).json({ error: err.message });
  }
});

// Snapshot deltas — latest day vs previous captured day
app.get("/api/deltas", async (req, res) => {
  try {
    const scope = await resolveDashboardScope(req);
    const selectedSet = new Set(scope.selectedPlanIds);
    const snapshots = await readSnapshots();
    const dates = Object.keys(snapshots).sort();

    if (dates.length < 2) {
      res.json({
        currentDate: dates.length === 1 ? dates[0] : null,
        previousDate: null,
        portfolio: null,
        byPlan: {},
      });
      return;
    }

    const currentDate = dates[dates.length - 1];
    const previousDate = dates[dates.length - 2];
    const current = snapshots[currentDate] || {};
    const previous = snapshots[previousDate] || {};
    const planIds = new Set([
      ...getSnapshotPlanEntries(current).map(([planId]) => planId),
      ...getSnapshotPlanEntries(previous).map(([planId]) => planId),
    ]);

    const byPlan = {};
    for (const planId of planIds) {
      if (!scope.includeAll && !selectedSet.has(planId)) continue;
      const now = current[planId] || {};
      const prev = previous[planId] || {};
      byPlan[planId] = {
        completedDelta: (now.completed || 0) - (prev.completed || 0),
        overdueDelta: (now.overdue || 0) - (prev.overdue || 0),
        dueSoonDelta: (now.dueSoon || 0) - (prev.dueSoon || 0),
        unassignedDelta: (now.unassigned || 0) - (prev.unassigned || 0),
      };
    }

    const portfolio = Object.values(byPlan).reduce(
      (acc, planDelta) => ({
        completedDelta: acc.completedDelta + planDelta.completedDelta,
        overdueDelta: acc.overdueDelta + planDelta.overdueDelta,
        dueSoonDelta: acc.dueSoonDelta + planDelta.dueSoonDelta,
        unassignedDelta: acc.unassignedDelta + planDelta.unassignedDelta,
      }),
      { completedDelta: 0, overdueDelta: 0, dueSoonDelta: 0, unassignedDelta: 0 }
    );

    res.json({
      currentDate,
      previousDate,
      portfolio,
      byPlan,
    });
  } catch (err) {
    console.error("Error fetching deltas:", err);
    res.status(500).json({ error: err.message });
  }
});

// Get buckets and tasks for a specific plan
app.get("/api/plans/:planId/stats", async (req, res) => {
  try {
    const client = await getGraphClient();
    const { planId } = req.params;

    const [buckets, tasks] = await Promise.all([
      fetchAllPages(client, `/planner/plans/${planId}/buckets`),
      fetchAllPages(client, `/planner/plans/${planId}/tasks`),
    ]);

    const allUserIds = new Set();
    tasks.forEach((t) => {
      if (t.assignments) {
        Object.keys(t.assignments).forEach((uid) => allUserIds.add(uid));
      }
    });

    const userMap = await resolveUsers(client, [...allUserIds]);

    const stats = buckets.map((bucket) => {
      const bucketTasks = tasks.filter((t) => t.bucketId === bucket.id);

      const enrichedTasks = bucketTasks.map((t) => {
        const assigneeIds = t.assignments ? Object.keys(t.assignments) : [];
        const status =
          t.percentComplete === 100
            ? "completed"
            : t.percentComplete > 0
            ? "inProgress"
            : "notStarted";

        return {
          id: t.id,
          title: t.title,
          bucketName: bucket.name,
          status,
          percentComplete: t.percentComplete,
          hasDescription: t.hasDescription !== false,
          assignees: assigneeIds.map((uid) => ({
            userId: uid,
            displayName: userMap.get(uid) || uid,
          })),
          dueDateTime: t.dueDateTime || null,
          dueStatus: getDueStatus(t.dueDateTime, t.percentComplete),
        };
      });

      const assigneeSummary = {};
      enrichedTasks.forEach((t) => {
        t.assignees.forEach((a) => {
          assigneeSummary[a.displayName] =
            (assigneeSummary[a.displayName] || 0) + 1;
        });
      });

      return {
        bucketId: bucket.id,
        bucketName: bucket.name,
        notStarted: enrichedTasks.filter((t) => t.status === "notStarted").length,
        inProgress: enrichedTasks.filter((t) => t.status === "inProgress").length,
        completed: enrichedTasks.filter((t) => t.status === "completed").length,
        total: enrichedTasks.length,
        overdue: enrichedTasks.filter((t) => t.dueStatus === "overdue").length,
        dueSoon: enrichedTasks.filter((t) => t.dueStatus === "dueSoon").length,
        onTrack: enrichedTasks.filter((t) => t.dueStatus === "onTrack").length,
        noDueDate: enrichedTasks.filter((t) => t.dueStatus === "noDueDate").length,
        assigneeSummary,
        tasks: enrichedTasks,
      };
    });

    res.json(stats);
  } catch (err) {
    console.error("Error fetching plan stats:", err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Planner Stats running on http://localhost:${PORT}`);
});
