const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const HOST = "0.0.0.0";
const PORT = Number(process.env.PORT) || 4173;
const ROOT = __dirname;
const INDEX_FILE = path.join(ROOT, "index.html");
const MAX_BODY_BYTES = 1024 * 1024;
const OAUTH_STATE_TTL_MS = 10 * 60 * 1000;

const SHOPIFY_API_KEY = String(process.env.SHOPIFY_API_KEY || "").trim();
const SHOPIFY_API_SECRET = String(process.env.SHOPIFY_API_SECRET || "").trim();
const SHOPIFY_API_VERSION = String(process.env.SHOPIFY_API_VERSION || "2025-10").trim();
const SHOPIFY_SCOPES = String(
  process.env.SHOPIFY_SCOPES || "read_orders,read_locations"
)
  .split(",")
  .map((scope) => scope.trim())
  .filter(Boolean);
const SUPABASE_URL = String(
  process.env.SUPABASE_URL || "https://pxcqxubehvnyaubqjcrf.supabase.co"
).replace(/\/+$/, "");
const SUPABASE_PUBLISHABLE_KEY = String(
  process.env.SUPABASE_PUBLISHABLE_KEY || "sb_publishable_MfL9s44GmmR9peD1jouetw_aZOa8xVa"
).trim();
const SUPABASE_SERVICE_ROLE_KEY = String(process.env.SUPABASE_SERVICE_ROLE_KEY || "").trim();
const SHOPIFY_TOKEN_ENCRYPTION_KEY = String(process.env.SHOPIFY_TOKEN_ENCRYPTION_KEY || "").trim();
const REGISTRATION_INVITES_TABLE = "registration_invites";
const ADMIN_SETTINGS_TABLE = "admin_settings";
const HISTORY_TABLE = "label_generations";
const PASSWORD_MIN_LENGTH = 10;
const DEFAULT_INVITE_EXPIRY_DAYS = 14;
const MAX_INVITE_EXPIRY_DAYS = 90;
const ADMIN_SETTINGS_SCOPE = "global";
const DEFAULT_ADMIN_SETTINGS = {
  carrier_discount_pct: 25,
  client_discount_pct: 20,
};
const RETAIL_BASE_RATE = {
  Economy: 7.1,
  Priority: 12.2,
  "International Express": 21.8,
};
const INVITE_ADMIN_EMAILS = String(process.env.INVITE_ADMIN_EMAILS || "")
  .split(",")
  .map((email) => email.trim().toLowerCase())
  .filter(Boolean);

const oauthStateStore = new Map();

const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".geojson": "application/geo+json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".ico": "image/x-icon",
  ".pdf": "application/pdf",
  ".txt": "text/plain; charset=utf-8",
  ".html": "text/html; charset=utf-8",
};

function send(res, status, headers, body) {
  res.writeHead(status, headers);
  res.end(body);
}

function sendJson(res, status, payload) {
  send(
    res,
    status,
    {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
    },
    JSON.stringify(payload)
  );
}

function sendRedirect(res, location) {
  send(
    res,
    302,
    {
      Location: location,
      "Cache-Control": "no-store",
    },
    ""
  );
}

function safeJoin(root, requestPath) {
  const fullPath = path.normalize(path.join(root, requestPath));
  if (!fullPath.startsWith(root)) return null;
  return fullPath;
}

function sendFile(res, filePath) {
  fs.readFile(filePath, (readError, data) => {
    if (readError) {
      send(res, 500, { "Content-Type": "text/plain; charset=utf-8" }, "Internal Server Error");
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    const isDev = process.env.NODE_ENV !== "production";
    const cacheControl = isDev
      ? "no-store"
      : ext === ".html"
        ? "no-cache"
        : "public, max-age=3600";
    send(
      res,
      200,
      {
        "Content-Type": MIME_TYPES[ext] || "application/octet-stream",
        "Cache-Control": cacheControl,
      },
      data
    );
  });
}

function buildPublicBaseUrl(req) {
  const forwardedProto = String(req.headers["x-forwarded-proto"] || "").split(",")[0].trim();
  const forwardedHost = String(req.headers["x-forwarded-host"] || "").split(",")[0].trim();
  const protocol = forwardedProto || (req.socket.encrypted ? "https" : "http");
  const host = forwardedHost || req.headers.host || `localhost:${PORT}`;
  return `${protocol}://${host}`;
}

function normalizeShopDomain(value) {
  let normalized = String(value || "")
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, "")
    .replace(/\/.*$/, "");
  if (!normalized) return "";
  if (!normalized.endsWith(".myshopify.com") && /^[a-z0-9][a-z0-9-]*$/.test(normalized)) {
    normalized = `${normalized}.myshopify.com`;
  }
  if (!/^[a-z0-9][a-z0-9-]*\.myshopify\.com$/.test(normalized)) {
    return "";
  }
  return normalized;
}

function getBearerToken(req) {
  const auth = String(req.headers.authorization || "");
  if (!auth.toLowerCase().startsWith("bearer ")) return "";
  return auth.slice(7).trim();
}

async function readJsonBody(req) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > MAX_BODY_BYTES) {
      throw new Error("Request body too large.");
    }
    chunks.push(chunk);
  }
  if (!chunks.length) return {};
  const body = Buffer.concat(chunks).toString("utf8").trim();
  if (!body) return {};
  try {
    return JSON.parse(body);
  } catch (_error) {
    throw new Error("Invalid JSON request body.");
  }
}

async function fetchSupabaseUser(accessToken) {
  if (!SUPABASE_URL || !SUPABASE_PUBLISHABLE_KEY) {
    throw new Error("Supabase auth environment is missing.");
  }
  if (!accessToken) {
    return null;
  }

  const response = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
    method: "GET",
    headers: {
      apikey: SUPABASE_PUBLISHABLE_KEY,
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    return null;
  }

  const payload = await response.json().catch(() => null);
  if (!payload || typeof payload !== "object" || !payload.id) {
    return null;
  }
  return payload;
}

async function getAuthenticatedUser(req) {
  const token = getBearerToken(req);
  if (!token) return null;
  return fetchSupabaseUser(token);
}

function verifyShopifyHmac(requestUrl) {
  if (!SHOPIFY_API_SECRET) return false;
  const receivedHmac = String(requestUrl.searchParams.get("hmac") || "").trim().toLowerCase();
  if (!receivedHmac || !/^[a-f0-9]+$/.test(receivedHmac)) return false;

  const params = new URLSearchParams(requestUrl.searchParams);
  params.delete("hmac");
  params.delete("signature");

  const message = Array.from(params.entries())
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, value]) => `${key}=${value}`)
    .join("&");

  const expectedHmac = crypto
    .createHmac("sha256", SHOPIFY_API_SECRET)
    .update(message)
    .digest("hex")
    .toLowerCase();

  const receivedBuffer = Buffer.from(receivedHmac, "utf8");
  const expectedBuffer = Buffer.from(expectedHmac, "utf8");
  if (receivedBuffer.length !== expectedBuffer.length) {
    return false;
  }
  return crypto.timingSafeEqual(receivedBuffer, expectedBuffer);
}

function pruneOauthStates() {
  const now = Date.now();
  for (const [state, entry] of oauthStateStore.entries()) {
    if (now - entry.createdAt > OAUTH_STATE_TTL_MS) {
      oauthStateStore.delete(state);
    }
  }
}

function createOauthState(payload) {
  pruneOauthStates();
  const state = crypto.randomBytes(24).toString("hex");
  oauthStateStore.set(state, {
    ...payload,
    createdAt: Date.now(),
  });
  return state;
}

function consumeOauthState(state, shop) {
  if (!state) return null;
  const stored = oauthStateStore.get(state);
  if (!stored) return null;
  oauthStateStore.delete(state);
  if (Date.now() - stored.createdAt > OAUTH_STATE_TTL_MS) return null;
  if (stored.shop !== shop) return null;
  return stored;
}

function getTokenCipherKey() {
  const secretSource = SHOPIFY_TOKEN_ENCRYPTION_KEY || SHOPIFY_API_SECRET;
  if (!secretSource) {
    throw new Error("Token encryption key is not configured.");
  }
  return crypto.createHash("sha256").update(secretSource).digest();
}

function encryptToken(plainToken) {
  const key = getTokenCipherKey();
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const encrypted = Buffer.concat([cipher.update(String(plainToken || ""), "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `v1:${iv.toString("base64url")}:${tag.toString("base64url")}:${encrypted.toString(
    "base64url"
  )}`;
}

function decryptToken(token) {
  const raw = String(token || "");
  if (!raw.startsWith("v1:")) {
    return raw;
  }
  const parts = raw.split(":");
  if (parts.length !== 4) {
    throw new Error("Invalid encrypted token format.");
  }
  const [, ivEncoded, tagEncoded, payloadEncoded] = parts;
  const key = getTokenCipherKey();
  const iv = Buffer.from(ivEncoded, "base64url");
  const tag = Buffer.from(tagEncoded, "base64url");
  const payload = Buffer.from(payloadEncoded, "base64url");
  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAuthTag(tag);
  const decrypted = Buffer.concat([decipher.update(payload), decipher.final()]);
  return decrypted.toString("utf8");
}

async function supabaseServiceRequest(pathnameWithQuery, options = {}) {
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    throw new Error("Supabase service role key is required for provider integration.");
  }
  const response = await fetch(`${SUPABASE_URL}${pathnameWithQuery}`, {
    ...options,
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
      ...(options.headers || {}),
    },
  });
  return response;
}

function normalizeInviteToken(value) {
  return String(value || "").trim();
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function isValidEmailFormat(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function hashInviteToken(value) {
  return crypto.createHash("sha256").update(String(value || "")).digest("hex");
}

function parseInviteExpiryDays(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return DEFAULT_INVITE_EXPIRY_DAYS;
  return Math.max(1, Math.min(MAX_INVITE_EXPIRY_DAYS, Math.round(numeric)));
}

function canManageRegistrationInvites(user) {
  if (!user || !user.id) return false;
  if (user?.user_metadata?.app_admin === true) return true;
  if (!INVITE_ADMIN_EMAILS.length) return true;
  const email = normalizeEmail(user.email || "");
  return Boolean(email && INVITE_ADMIN_EMAILS.includes(email));
}

function normalizePercent(value, fallback = 0) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.max(0, Math.min(100, Number(numeric.toFixed(2))));
}

function normalizeAdminSettings(rawSettings = {}) {
  return {
    carrier_discount_pct: normalizePercent(
      rawSettings?.carrier_discount_pct,
      DEFAULT_ADMIN_SETTINGS.carrier_discount_pct
    ),
    client_discount_pct: normalizePercent(
      rawSettings?.client_discount_pct,
      DEFAULT_ADMIN_SETTINGS.client_discount_pct
    ),
    updated_at: String(rawSettings?.updated_at || "").trim(),
    updated_by: rawSettings?.updated_by || null,
    storage_ready: rawSettings?.storage_ready !== false,
  };
}

function inviteIsRevoked(invite) {
  return Boolean(String(invite?.revoked_at || "").trim());
}

function getInviteStatus(invite) {
  if (invite?.claimed_at) return "claimed";
  if (inviteIsRevoked(invite)) return "revoked";
  if (inviteIsExpired(invite)) return "expired";
  return "open";
}

function monthsBetweenInclusive(startValue, endValue = Date.now()) {
  const start = new Date(startValue);
  const end = new Date(endValue);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return 1;
  const years = end.getFullYear() - start.getFullYear();
  const months = end.getMonth() - start.getMonth();
  return Math.max(1, years * 12 + months + 1);
}

function getEstimatedRetailTotal(serviceType, quantity, fallbackTotal = 0) {
  const retailRate = Number(RETAIL_BASE_RATE[String(serviceType || "").trim()] || 0);
  if (retailRate > 0) {
    return retailRate * Math.max(0, Number(quantity) || 0);
  }
  return Math.max(0, Number(fallbackTotal) || 0);
}

function buildAdminClientMetrics(user, historyRows, settings) {
  const rows = Array.isArray(historyRows) ? historyRows : [];
  const totalLabels = rows.reduce((sum, row) => sum + Math.max(0, Number(row?.quantity) || 0), 0);
  const totalRevenue = rows.reduce(
    (sum, row) => sum + Math.max(0, Number(row?.total_price) || 0),
    0
  );
  const estimatedRetail = rows.reduce(
    (sum, row) =>
      sum +
      getEstimatedRetailTotal(row?.service_type, row?.quantity, Number(row?.total_price) || 0),
    0
  );
  const carrierDiscountPct = normalizePercent(settings?.carrier_discount_pct, 0);
  const estimatedCarrierCost = Number(
    (estimatedRetail * (1 - carrierDiscountPct / 100)).toFixed(2)
  );
  const totalProfit = Number((totalRevenue - estimatedCarrierCost).toFixed(2));
  const activeMonths = monthsBetweenInclusive(user?.created_at || Date.now(), Date.now());
  const mrr = Number((totalRevenue / activeMonths).toFixed(2));
  const mrp = Number((totalProfit / activeMonths).toFixed(2));
  const avgParcelsPerMonth = Number((totalLabels / activeMonths).toFixed(1));
  const lastGenerationAt = rows.reduce((latest, row) => {
    const createdAt = String(row?.created_at || "").trim();
    if (!createdAt) return latest;
    if (!latest) return createdAt;
    return Date.parse(createdAt) > Date.parse(latest) ? createdAt : latest;
  }, "");

  let activityStatus = "dormant";
  if (lastGenerationAt) {
    const ageDays = (Date.now() - Date.parse(lastGenerationAt)) / (1000 * 60 * 60 * 24);
    if (ageDays <= 30) {
      activityStatus = "active";
    } else if (ageDays <= 90) {
      activityStatus = "quiet";
    }
  } else if (totalLabels > 0) {
    activityStatus = "quiet";
  }

  return {
    total_labels: totalLabels,
    total_revenue_ex_vat: Number(totalRevenue.toFixed(2)),
    estimated_carrier_cost_ex_vat: estimatedCarrierCost,
    total_profit_ex_vat: totalProfit,
    mrr_ex_vat: mrr,
    mrp_ex_vat: mrp,
    avg_parcels_per_month: avgParcelsPerMonth,
    last_generation_at: lastGenerationAt || null,
    avg_payment_days: null,
    last_invoice_tracking: totalRevenue > 0 ? "Billing not live" : "No billable activity",
    activity_status: activityStatus,
  };
}

function buildAdminClientSummary(clients = [], invites = []) {
  const activeClients = clients.filter((client) => client?.metrics?.activity_status === "active").length;
  const openInvites = invites.filter((invite) => getInviteStatus(invite) === "open").length;
  const totalRevenue = clients.reduce(
    (sum, client) => sum + Math.max(0, Number(client?.metrics?.total_revenue_ex_vat) || 0),
    0
  );
  const totalProfit = clients.reduce(
    (sum, client) => sum + Number(client?.metrics?.total_profit_ex_vat || 0),
    0
  );
  return {
    total_clients: clients.length,
    active_clients: activeClients,
    open_invites: openInvites,
    total_revenue_ex_vat: Number(totalRevenue.toFixed(2)),
    total_profit_ex_vat: Number(totalProfit.toFixed(2)),
  };
}

function normalizeRegistrationProfile(rawProfile) {
  const profile = rawProfile && typeof rawProfile === "object" ? rawProfile : {};
  return {
    companyName: String(profile.companyName || "").trim(),
    contactName: String(profile.contactName || "").trim(),
    contactEmail: normalizeEmail(profile.contactEmail || ""),
    contactPhone: String(profile.contactPhone || "").trim(),
    billingAddress: String(profile.billingAddress || "").trim(),
    taxId: String(profile.taxId || "").trim(),
    customerId: String(profile.customerId || "").trim(),
  };
}

function inviteIsExpired(invite) {
  const expiresAt = String(invite?.expires_at || "").trim();
  if (!expiresAt) return true;
  const expiresMs = Date.parse(expiresAt);
  if (!Number.isFinite(expiresMs)) return true;
  return Date.now() > expiresMs;
}

function mapInviteToPublic(invite) {
  return {
    email: normalizeEmail(invite?.invited_email || ""),
    companyName: String(invite?.company_name || "").trim(),
    contactName: String(invite?.contact_name || "").trim(),
    contactEmail: normalizeEmail(invite?.contact_email || ""),
    contactPhone: String(invite?.contact_phone || "").trim(),
    billingAddress: String(invite?.billing_address || "").trim(),
    taxId: String(invite?.tax_id || "").trim(),
    customerId: String(invite?.customer_id || "").trim(),
    planName: String(invite?.plan_name || "Growth").trim() || "Growth",
  };
}

function buildRegistrationMetadata(invite, submittedProfile, email, preferredLanguage = "en") {
  const inviteProfile = mapInviteToPublic(invite);
  return {
    company_name: submittedProfile.companyName || inviteProfile.companyName || "",
    contact_name: submittedProfile.contactName || inviteProfile.contactName || "",
    contact_email:
      submittedProfile.contactEmail || inviteProfile.contactEmail || normalizeEmail(email),
    contact_phone: submittedProfile.contactPhone || inviteProfile.contactPhone || "",
    billing_address: submittedProfile.billingAddress || inviteProfile.billingAddress || "",
    tax_id: submittedProfile.taxId || inviteProfile.taxId || "",
    customer_id: submittedProfile.customerId || inviteProfile.customerId || "",
    plan_name: inviteProfile.planName || "Growth",
    registration_source: "invite",
    preferred_language: String(preferredLanguage || "en").trim().toLowerCase() || "en",
  };
}

async function getRegistrationInviteByToken(token) {
  const inviteToken = normalizeInviteToken(token);
  if (!inviteToken) return null;
  const tokenHash = hashInviteToken(inviteToken);
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,invited_email,company_name,contact_name,contact_email,contact_phone,billing_address,tax_id,customer_id,plan_name,expires_at,claimed_at,revoked_at"
  );
  params.set("token_hash", `eq.${tokenHash}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${REGISTRATION_INVITES_TABLE}?${params.toString()}`,
    {
      method: "GET",
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Invite lookup failed (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) return null;
  return rows[0];
}

async function insertRegistrationInvite({
  tokenHash,
  tokenEncrypted,
  invitedEmail,
  expiresAt,
  createdBy,
}) {
  const response = await supabaseServiceRequest(`/rest/v1/${REGISTRATION_INVITES_TABLE}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify([
      {
        token_hash: tokenHash,
        token_encrypted: tokenEncrypted || null,
        invited_email: invitedEmail || null,
        expires_at: expiresAt,
        created_by: createdBy || null,
      },
    ]),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not create invite (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function createRegistrationInvite({ invitedEmail, expiresInDays, createdBy }) {
  const safeInvitedEmail = normalizeEmail(invitedEmail || "");
  const expiryDays = parseInviteExpiryDays(expiresInDays);
  const expiresAt = new Date(Date.now() + expiryDays * 24 * 60 * 60 * 1000).toISOString();

  for (let attempt = 0; attempt < 3; attempt += 1) {
    const token = crypto.randomBytes(32).toString("hex");
    const tokenHash = hashInviteToken(token);
    try {
      await insertRegistrationInvite({
        tokenHash,
        tokenEncrypted: encryptToken(token),
        invitedEmail: safeInvitedEmail,
        expiresAt,
        createdBy,
      });
      return {
        token,
        expiresAt,
      };
    } catch (error) {
      const message = String(error?.message || "");
      if (message.includes("23505") || message.includes("duplicate key")) {
        continue;
      }
      throw error;
    }
  }
  throw new Error("Could not create invite. Please try again.");
}

async function listRegistrationInvites({ createdBy, limit = 20 }) {
  const safeLimit = Math.max(1, Math.min(100, Number(limit) || 20));
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,invited_email,expires_at,created_at,claimed_at,claimed_email,created_by,token_encrypted,revoked_at,revoked_by"
  );
  params.set("order", "created_at.desc");
  params.set("limit", String(safeLimit));
  if (createdBy) {
    params.set("created_by", `eq.${createdBy}`);
  }
  const response = await supabaseServiceRequest(
    `/rest/v1/${REGISTRATION_INVITES_TABLE}?${params.toString()}`,
    {
      method: "GET",
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Invite history lookup failed (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function revokeRegistrationInvite(inviteId, revokedBy) {
  const safeInviteId = String(inviteId || "").trim();
  if (!safeInviteId) {
    throw new Error("Invite id is required.");
  }

  const selectParams = new URLSearchParams();
  selectParams.set("select", "id,claimed_at,revoked_at");
  selectParams.set("id", `eq.${safeInviteId}`);
  selectParams.set("limit", "1");

  const existingResponse = await supabaseServiceRequest(
    `/rest/v1/${REGISTRATION_INVITES_TABLE}?${selectParams.toString()}`,
    { method: "GET" }
  );
  if (!existingResponse.ok) {
    const details = await existingResponse.text().catch(() => "");
    throw new Error(`Could not load invite (${existingResponse.status}) ${details}`.trim());
  }
  const existingRows = await existingResponse.json().catch(() => []);
  const existingInvite = Array.isArray(existingRows) ? existingRows[0] : null;
  if (!existingInvite?.id) {
    throw new Error("Invite not found.");
  }
  if (existingInvite.claimed_at) {
    throw new Error("Claimed invites cannot be revoked.");
  }
  if (existingInvite.revoked_at) {
    throw new Error("Invite has already been revoked.");
  }

  const updateParams = new URLSearchParams();
  updateParams.set("id", `eq.${safeInviteId}`);
  const response = await supabaseServiceRequest(
    `/rest/v1/${REGISTRATION_INVITES_TABLE}?${updateParams.toString()}`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=minimal",
      },
      body: JSON.stringify({
        revoked_at: new Date().toISOString(),
        revoked_by: revokedBy || null,
      }),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not revoke invite (${response.status}) ${details}`.trim());
  }
}

async function listSupabaseUsers(limit = 200) {
  const safeLimit = Math.max(1, Math.min(1000, Number(limit) || 200));
  const pageSize = Math.min(100, safeLimit);
  const users = [];
  for (let page = 1; users.length < safeLimit && page <= 20; page += 1) {
    const response = await fetch(
      `${SUPABASE_URL}/auth/v1/admin/users?page=${page}&per_page=${pageSize}`,
      {
        method: "GET",
        headers: {
          apikey: SUPABASE_SERVICE_ROLE_KEY,
          Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
        },
      }
    );
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message =
        payload?.msg ||
        payload?.error_description ||
        payload?.error ||
        `Could not load clients (${response.status}).`;
      throw new Error(String(message).trim());
    }
    const pageUsers = Array.isArray(payload?.users) ? payload.users : [];
    users.push(...pageUsers);
    if (pageUsers.length < pageSize) break;
  }
  return users.slice(0, safeLimit);
}

async function listGenerationHistoryRows(limit = 5000) {
  const safeLimit = Math.max(1, Math.min(20000, Number(limit) || 5000));
  const pageSize = 1000;
  const rows = [];
  for (let start = 0; start < safeLimit; start += pageSize) {
    const end = Math.min(safeLimit - 1, start + pageSize - 1);
    const params = new URLSearchParams();
    params.set("select", "user_id,created_at,service_type,quantity,total_price,payload");
    params.set("order", "created_at.desc");
    const response = await supabaseServiceRequest(
      `/rest/v1/${HISTORY_TABLE}?${params.toString()}`,
      {
        method: "GET",
        headers: {
          Range: `${start}-${end}`,
          "Range-Unit": "items",
        },
      }
    );
    if (!response.ok) {
      const details = await response.text().catch(() => "");
      throw new Error(`Could not load label history (${response.status}) ${details}`.trim());
    }
    const pageRows = await response.json().catch(() => []);
    if (!Array.isArray(pageRows) || !pageRows.length) break;
    rows.push(...pageRows);
    if (pageRows.length < pageSize) break;
  }
  return rows;
}

async function getAdminSettings() {
  const params = new URLSearchParams();
  params.set("select", "scope,carrier_discount_pct,client_discount_pct,updated_at,updated_by");
  params.set("scope", `eq.${ADMIN_SETTINGS_SCOPE}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${ADMIN_SETTINGS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (/relation .*admin_settings/i.test(details)) {
      return {
        ...DEFAULT_ADMIN_SETTINGS,
        updated_at: "",
        updated_by: null,
        storage_ready: false,
      };
    }
    throw new Error(`Could not load admin settings (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  const settingsRow = Array.isArray(rows) && rows.length ? rows[0] : DEFAULT_ADMIN_SETTINGS;
  return normalizeAdminSettings(settingsRow);
}

async function saveAdminSettings(userId, payload) {
  const normalized = normalizeAdminSettings(payload);
  const response = await supabaseServiceRequest(
    `/rest/v1/${ADMIN_SETTINGS_TABLE}?on_conflict=scope`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=representation",
      },
      body: JSON.stringify([
        {
          scope: ADMIN_SETTINGS_SCOPE,
          carrier_discount_pct: normalized.carrier_discount_pct,
          client_discount_pct: normalized.client_discount_pct,
          updated_at: new Date().toISOString(),
          updated_by: userId || null,
        },
      ]),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not save admin settings (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return normalizeAdminSettings(Array.isArray(rows) && rows.length ? rows[0] : normalized);
}

function buildInviteUrl(baseUrl, invite) {
  const encrypted = String(invite?.token_encrypted || "").trim();
  if (!encrypted) return "";
  try {
    const token = decryptToken(encrypted);
    const inviteUrl = new URL("/register", baseUrl);
    inviteUrl.searchParams.set("invite", token);
    return inviteUrl.toString();
  } catch (_error) {
    return "";
  }
}

function mapInviteHistoryRow(invite, baseUrl) {
  return {
    id: invite?.id || null,
    invited_email: normalizeEmail(invite?.invited_email || ""),
    expires_at: invite?.expires_at || null,
    created_at: invite?.created_at || null,
    claimed_at: invite?.claimed_at || null,
    claimed_email: normalizeEmail(invite?.claimed_email || ""),
    revoked_at: invite?.revoked_at || null,
    invite_url: buildInviteUrl(baseUrl, invite),
  };
}

function shouldIncludeAdminClient(user) {
  if (!user?.id) return false;
  if (user?.user_metadata?.app_admin === true) return false;
  const email = normalizeEmail(user.email || "");
  if (email && INVITE_ADMIN_EMAILS.includes(email)) return false;
  return true;
}

async function buildAdminDashboard(baseUrl) {
  const [settings, invites, users, historyRows] = await Promise.all([
    getAdminSettings(),
    listRegistrationInvites({ limit: 50 }),
    listSupabaseUsers(250),
    listGenerationHistoryRows(10000),
  ]);

  const historyByUserId = new Map();
  historyRows.forEach((row) => {
    const userId = String(row?.user_id || "").trim();
    if (!userId) return;
    if (!historyByUserId.has(userId)) {
      historyByUserId.set(userId, []);
    }
    historyByUserId.get(userId).push(row);
  });

  const clients = users
    .filter(shouldIncludeAdminClient)
    .map((user) => ({
      user: {
        id: user.id,
        email: normalizeEmail(user.email || ""),
        created_at: user.created_at || null,
        user_metadata: user.user_metadata && typeof user.user_metadata === "object" ? user.user_metadata : {},
      },
      metrics: buildAdminClientMetrics(user, historyByUserId.get(user.id) || [], settings),
    }))
    .sort((left, right) => {
      const leftValue = Date.parse(left?.metrics?.last_generation_at || left?.user?.created_at || 0);
      const rightValue = Date.parse(right?.metrics?.last_generation_at || right?.user?.created_at || 0);
      return rightValue - leftValue;
    });

  const inviteHistory = invites.map((invite) => mapInviteHistoryRow(invite, baseUrl));
  return {
    settings,
    invites: inviteHistory,
    clients,
    summary: buildAdminClientSummary(clients, invites),
  };
}

async function markRegistrationInviteClaimed(inviteId, userId, email) {
  if (!inviteId || !userId) return;
  const params = new URLSearchParams();
  params.set("id", `eq.${inviteId}`);
  const response = await supabaseServiceRequest(
    `/rest/v1/${REGISTRATION_INVITES_TABLE}?${params.toString()}`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=minimal",
      },
      body: JSON.stringify({
        claimed_at: new Date().toISOString(),
        claimed_user_id: userId,
        claimed_email: normalizeEmail(email),
      }),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed to mark invite as claimed (${response.status}) ${details}`.trim());
  }
}

async function ensureAccountSettingsRow(userId) {
  if (!userId) return;
  const response = await supabaseServiceRequest(
    "/rest/v1/account_settings?on_conflict=user_id",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify([
        {
          user_id: userId,
          warehouse_origins: [],
          updated_at: new Date().toISOString(),
        },
      ]),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed creating account settings (${response.status}) ${details}`.trim());
  }
}

async function createSupabaseUserWithMetadata(email, password, userMetadata) {
  const response = await fetch(`${SUPABASE_URL}/auth/v1/admin/users`, {
    method: "POST",
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      email: normalizeEmail(email),
      password: String(password || ""),
      email_confirm: true,
      user_metadata: userMetadata,
    }),
  });
  const payload = await response.json().catch(() => null);
  if (!response.ok) {
    const message =
      payload?.msg || payload?.error_description || payload?.error || "Could not create account.";
    throw new Error(`${message}`.trim());
  }
  if (!payload || typeof payload !== "object" || !payload.id) {
    throw new Error("Account creation response was invalid.");
  }
  return payload;
}

async function exchangeShopifyAccessToken(shop, code) {
  const response = await fetch(`https://${shop}/admin/oauth/access_token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      client_id: SHOPIFY_API_KEY,
      client_secret: SHOPIFY_API_SECRET,
      code,
    }),
  });

  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Shopify token exchange failed (${response.status}) ${details}`.trim());
  }

  const payload = await response.json().catch(() => null);
  if (!payload || !payload.access_token) {
    throw new Error("Shopify token response was invalid.");
  }
  return payload;
}

async function upsertShopifyConnection(userId, shop, accessToken, scopes) {
  const encryptedToken = encryptToken(accessToken);
  const nowIso = new Date().toISOString();
  const payload = [
    {
      user_id: userId,
      provider: "shopify",
      shop_domain: shop,
      status: "connected",
      scopes: String(scopes || ""),
      access_token: encryptedToken,
      updated_at: nowIso,
      connected_at: nowIso,
    },
  ];

  const response = await supabaseServiceRequest(
    "/rest/v1/provider_connections?on_conflict=user_id,provider,shop_domain",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify(payload),
    }
  );

  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed saving Shopify connection (${response.status}) ${details}`.trim());
  }
}

async function getShopifyConnection(userId, requestedShop, options = {}) {
  const { includeSettings = false } = options;
  const params = new URLSearchParams();
  params.set(
    "select",
    includeSettings
      ? "shop_domain,status,scopes,connected_at,updated_at,access_token,import_settings"
      : "shop_domain,status,scopes,connected_at,updated_at,access_token"
  );
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.shopify");
  params.set("status", "eq.connected");
  if (requestedShop) {
    params.set("shop_domain", `eq.${requestedShop}`);
  }
  params.set("order", "updated_at.desc");
  params.set("limit", "1");

  const response = await supabaseServiceRequest(`/rest/v1/provider_connections?${params.toString()}`, {
    method: "GET",
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (
      includeSettings &&
      /import_settings/.test(String(details || "").toLowerCase()) &&
      /column/.test(String(details || "").toLowerCase())
    ) {
      const error = new Error(
        "Shopify settings storage is not enabled yet. Run the latest supabase_provider_connections.sql migration."
      );
      error.code = "missing_import_settings_column";
      throw error;
    }
    throw new Error(`Failed reading Shopify connection (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) {
    return null;
  }
  return rows[0];
}

function getSelectedLocationIdsFromImportSettings(importSettings) {
  if (!importSettings || typeof importSettings !== "object") return [];
  const selectedRaw = Array.isArray(importSettings.selected_location_ids)
    ? importSettings.selected_location_ids
    : Array.isArray(importSettings.selectedLocationIds)
      ? importSettings.selectedLocationIds
      : [];
  return sanitizeSelectedLocationIds(selectedRaw);
}

async function saveShopifyImportSettings(userId, shop, selectedLocationIds) {
  const params = new URLSearchParams();
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.shopify");
  params.set("shop_domain", `eq.${shop}`);
  params.set("status", "eq.connected");

  const response = await supabaseServiceRequest(
    `/rest/v1/provider_connections?${params.toString()}`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=minimal",
      },
      body: JSON.stringify({
        import_settings: {
          selected_location_ids: sanitizeSelectedLocationIds(selectedLocationIds),
        },
        updated_at: new Date().toISOString(),
      }),
    }
  );

  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (
      /import_settings/.test(String(details || "").toLowerCase()) &&
      /column/.test(String(details || "").toLowerCase())
    ) {
      const error = new Error(
        "Shopify settings storage is not enabled yet. Run the latest supabase_provider_connections.sql migration."
      );
      error.code = "missing_import_settings_column";
      throw error;
    }
    throw new Error(`Failed saving Shopify settings (${response.status}) ${details}`.trim());
  }
}

async function setShopifyConnectionStatus(userId, shop, status) {
  const params = new URLSearchParams();
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.shopify");
  params.set("shop_domain", `eq.${shop}`);

  const response = await supabaseServiceRequest(
    `/rest/v1/provider_connections?${params.toString()}`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=minimal",
      },
      body: JSON.stringify({
        status: String(status || "disconnected"),
        updated_at: new Date().toISOString(),
      }),
    }
  );

  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed updating Shopify connection (${response.status}) ${details}`.trim());
  }
}

function normalizeLocationId(value) {
  const asString = String(value ?? "").trim();
  if (!asString) return "";
  const numeric = Number(asString);
  if (Number.isFinite(numeric) && numeric > 0) {
    return String(Math.trunc(numeric));
  }
  return asString;
}

function sanitizeSelectedLocationIds(values) {
  if (!Array.isArray(values)) return [];
  const ids = new Set();
  values.forEach((value) => {
    const id = normalizeLocationId(value);
    if (!id) return;
    ids.add(id);
  });
  return Array.from(ids);
}

function mapLocationToSender(location) {
  if (!location || typeof location !== "object") {
    return {
      senderName: "",
      senderStreet: "",
      senderCity: "",
      senderState: "",
      senderZip: "",
    };
  }
  const street = [location.address1, location.address2]
    .map((part) => String(part || "").trim())
    .filter(Boolean)
    .join(", ");
  return {
    senderName: String(location.name || "").trim(),
    senderStreet: street,
    senderCity: String(location.city || "").trim(),
    senderState: String(location.province_code || location.province || "").trim(),
    senderZip: String(location.zip || "").trim(),
  };
}

function getOrderLocationId(order) {
  const candidates = [];
  candidates.push(order?.location_id);
  candidates.push(order?.origin_location?.id);

  if (Array.isArray(order?.fulfillments)) {
    order.fulfillments.forEach((fulfillment) => {
      candidates.push(fulfillment?.location_id);
    });
  }

  if (Array.isArray(order?.line_items)) {
    order.line_items.forEach((item) => {
      candidates.push(item?.origin_location?.id);
    });
  }

  for (const value of candidates) {
    const normalized = normalizeLocationId(value);
    if (normalized) return normalized;
  }
  return "";
}

function indexLocationsById(locations) {
  const byId = {};
  if (!Array.isArray(locations)) return byId;
  locations.forEach((location) => {
    const id = normalizeLocationId(location?.id);
    if (!id) return;
    byId[id] = location;
  });
  return byId;
}

function mapShopifyOrdersToCsvRows(orders, options = {}) {
  const { locationById = {}, selectedLocationIds = [] } = options;
  if (!Array.isArray(orders)) return [];

  const selectedSet = new Set(sanitizeSelectedLocationIds(selectedLocationIds));
  const singleOverrideId = selectedSet.size === 1 ? selectedSet.values().next().value : "";
  const mirrorSelection = selectedSet.size > 1;

  return orders.map((order) => {
    const shipping = order?.shipping_address || {};
    const recipientName =
      String(shipping?.name || "").trim() ||
      [shipping?.first_name, shipping?.last_name].filter(Boolean).join(" ").trim();
    const totalWeightGrams = Number(order?.total_weight);
    const weightKg =
      Number.isFinite(totalWeightGrams) && totalWeightGrams > 0
        ? (totalWeightGrams / 1000).toFixed(2)
        : "0.50";

    let senderLocationId = "";
    if (singleOverrideId) {
      senderLocationId = singleOverrideId;
    } else if (mirrorSelection) {
      const orderLocationId = getOrderLocationId(order);
      if (orderLocationId && selectedSet.has(orderLocationId)) {
        senderLocationId = orderLocationId;
      }
    }

    const sender = mapLocationToSender(locationById[senderLocationId]);

    return {
      senderName: sender.senderName,
      senderStreet: sender.senderStreet,
      senderCity: sender.senderCity,
      senderState: sender.senderState,
      senderZip: sender.senderZip,
      recipientName,
      recipientStreet: String(shipping?.address1 || "").trim(),
      recipientCity: String(shipping?.city || "").trim(),
      recipientState: String(shipping?.province_code || shipping?.province || "").trim(),
      recipientZip: String(shipping?.zip || "").trim(),
      recipientCountry: String(shipping?.country_code || shipping?.country || "").trim(),
      packageWeight: weightKg,
      packageDims: "25 x 20 x 10",
    };
  });
}

async function fetchShopifyLocations(shop, accessToken) {
  const url = new URL(`https://${shop}/admin/api/${SHOPIFY_API_VERSION}/locations.json`);
  url.searchParams.set("limit", "250");
  url.searchParams.set(
    "fields",
    "id,name,address1,address2,city,province,province_code,zip,country,country_code"
  );

  const response = await fetch(url.toString(), {
    method: "GET",
    headers: {
      "X-Shopify-Access-Token": accessToken,
      "Content-Type": "application/json",
    },
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const error = new Error(`Shopify location fetch failed (${response.status}) ${details}`.trim());
    error.status = response.status;
    throw error;
  }
  const payload = await response.json().catch(() => null);
  const locations = Array.isArray(payload?.locations) ? payload.locations : [];
  return locations
    .map((location) => ({
      id: normalizeLocationId(location?.id),
      name: String(location?.name || "").trim(),
      address1: String(location?.address1 || "").trim(),
      address2: String(location?.address2 || "").trim(),
      city: String(location?.city || "").trim(),
      province: String(location?.province || "").trim(),
      province_code: String(location?.province_code || "").trim(),
      zip: String(location?.zip || "").trim(),
      country: String(location?.country_code || location?.country || "").trim(),
    }))
    .filter((location) => location.id && location.name);
}

async function fetchShopifyOrders(shop, accessToken, limit) {
  const safeLimit = Number.isFinite(limit) ? Math.max(1, Math.min(250, Math.trunc(limit))) : 50;
  const url = new URL(`https://${shop}/admin/api/${SHOPIFY_API_VERSION}/orders.json`);
  url.searchParams.set("status", "open");
  url.searchParams.set("limit", String(safeLimit));
  url.searchParams.set(
    "fields",
    "id,name,created_at,total_weight,shipping_address,currency,current_total_price,total_price,location_id,origin_location,line_items,fulfillments"
  );
  const response = await fetch(url.toString(), {
    method: "GET",
    headers: {
      "X-Shopify-Access-Token": accessToken,
      "Content-Type": "application/json",
    },
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const error = new Error(`Shopify order import failed (${response.status}) ${details}`.trim());
    error.status = response.status;
    throw error;
  }
  const payload = await response.json().catch(() => null);
  const orders = Array.isArray(payload?.orders) ? payload.orders : [];
  return orders;
}

function getCallbackRedirect(req, params) {
  const url = new URL("/label-info", buildPublicBaseUrl(req));
  Object.entries(params).forEach(([key, value]) => {
    if (value === undefined || value === null || value === "") return;
    url.searchParams.set(key, String(value));
  });
  return url.toString();
}

function assertShopifyAppConfig() {
  if (!SHOPIFY_API_KEY || !SHOPIFY_API_SECRET) {
    throw new Error("SHOPIFY_API_KEY and SHOPIFY_API_SECRET must be configured.");
  }
  if (!SHOPIFY_SCOPES.length) {
    throw new Error("SHOPIFY_SCOPES must include at least one scope.");
  }
  if (!SUPABASE_URL || !SUPABASE_PUBLISHABLE_KEY || !SUPABASE_SERVICE_ROLE_KEY) {
    throw new Error(
      "SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, and SUPABASE_SERVICE_ROLE_KEY are required."
    );
  }
}

async function handleCreateRegistrationInvite(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to create client invites." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const invitedEmail = normalizeEmail(body?.invitedEmail || "");
  if (!invitedEmail) {
    sendJson(res, 400, { error: "Client email is required." });
    return;
  }
  if (!isValidEmailFormat(invitedEmail)) {
    sendJson(res, 400, { error: "Invalid client email format." });
    return;
  }

  try {
    const created = await createRegistrationInvite({
      invitedEmail,
      expiresInDays: body?.expiresInDays,
      createdBy: user.id,
    });
    const inviteUrl = new URL("/register", buildPublicBaseUrl(req));
    inviteUrl.searchParams.set("invite", created.token);
    sendJson(res, 200, {
      ok: true,
      inviteUrl: inviteUrl.toString(),
      expiresAt: created.expiresAt,
      invitedEmail: invitedEmail || null,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not create invite." });
  }
}

async function handleListRegistrationInvites(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to create client invites." });
    return;
  }

  const limit = Number(requestUrl.searchParams.get("limit") || "20");
  try {
    const invites = await listRegistrationInvites({
      createdBy: null,
      limit,
    });
    sendJson(res, 200, {
      invites: invites.map((invite) => mapInviteHistoryRow(invite, buildPublicBaseUrl(req))),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load invite history." });
  }
}

async function handleAdminStatus(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  sendJson(res, 200, { allowed: canManageRegistrationInvites(user) });
}

async function handleAdminDashboard(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to access the admin panel." });
    return;
  }
  try {
    const dashboard = await buildAdminDashboard(buildPublicBaseUrl(req));
    sendJson(res, 200, dashboard);
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load admin dashboard." });
  }
}

async function handleAdminSettingsSave(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to update admin settings." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  try {
    const settings = await saveAdminSettings(user.id, body || {});
    sendJson(res, 200, { ok: true, settings });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not save admin settings." });
  }
}

async function handleAdminInviteRevoke(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to revoke client invites." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const inviteId = String(body?.inviteId || "").trim();
  if (!inviteId) {
    sendJson(res, 400, { error: "Invite id is required." });
    return;
  }

  try {
    await revokeRegistrationInvite(inviteId, user.id);
    sendJson(res, 200, { ok: true });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not revoke invite." });
  }
}

async function handleRegisterInviteValidate(_req, res, requestUrl) {
  const token = normalizeInviteToken(requestUrl.searchParams.get("token"));
  if (!token) {
    sendJson(res, 400, { error: "Registration link required." });
    return;
  }
  try {
    const invite = await getRegistrationInviteByToken(token);
    if (!invite || invite.claimed_at || inviteIsRevoked(invite) || inviteIsExpired(invite)) {
      sendJson(res, 410, { error: "This registration link is invalid or expired." });
      return;
    }
    sendJson(res, 200, {
      valid: true,
      invite: mapInviteToPublic(invite),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not validate registration link." });
  }
}

async function handleRegisterWithInvite(req, res) {
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const token = normalizeInviteToken(body?.token);
  const email = normalizeEmail(body?.email);
  const password = String(body?.password || "");
  const profile = normalizeRegistrationProfile(body?.profile);
  const preferredLanguage = String(body?.preferredLanguage || "en").trim().toLowerCase();

  if (!token) {
    sendJson(res, 400, { error: "Registration link required." });
    return;
  }
  if (!email || !password) {
    sendJson(res, 400, { error: "Email and password are required." });
    return;
  }
  if (password.length < PASSWORD_MIN_LENGTH) {
    sendJson(res, 400, {
      error: `Password must be at least ${PASSWORD_MIN_LENGTH} characters.`,
    });
    return;
  }

  const requiredProfileValues = [
    profile.companyName,
    profile.contactName,
    profile.contactEmail,
    profile.contactPhone,
    profile.billingAddress,
    profile.taxId,
  ];
  if (requiredProfileValues.some((value) => !String(value || "").trim())) {
    sendJson(res, 400, { error: "All registration fields are required except Customer ID." });
    return;
  }

  try {
    const invite = await getRegistrationInviteByToken(token);
    if (!invite || invite.claimed_at || inviteIsRevoked(invite) || inviteIsExpired(invite)) {
      sendJson(res, 410, { error: "This registration link is invalid or expired." });
      return;
    }

    const metadata = buildRegistrationMetadata(invite, profile, email, preferredLanguage);
    const createdUser = await createSupabaseUserWithMetadata(email, password, metadata);

    await ensureAccountSettingsRow(createdUser.id);
    await markRegistrationInviteClaimed(invite.id, createdUser.id, email);

    sendJson(res, 200, {
      ok: true,
      email,
      userId: createdUser.id,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not complete registration." });
  }
}

async function handleShopifyInstallLink(req, res, requestUrl) {
  assertShopifyAppConfig();
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  const body = await readJsonBody(req);
  const shop = normalizeShopDomain(body?.shop);
  if (!shop) {
    sendJson(res, 400, { error: "Enter a valid Shopify domain (store.myshopify.com)." });
    return;
  }

  const state = createOauthState({
    userId: user.id,
    shop,
  });
  const callbackUrl = `${buildPublicBaseUrl(req)}/api/shopify/callback`;
  const installUrl = new URL(`https://${shop}/admin/oauth/authorize`);
  installUrl.searchParams.set("client_id", SHOPIFY_API_KEY);
  installUrl.searchParams.set("scope", SHOPIFY_SCOPES.join(","));
  installUrl.searchParams.set("redirect_uri", callbackUrl);
  installUrl.searchParams.set("state", state);

  sendJson(res, 200, { url: installUrl.toString(), shop });
}

async function handleShopifyCallback(req, res, requestUrl) {
  assertShopifyAppConfig();
  const shop = normalizeShopDomain(requestUrl.searchParams.get("shop"));
  const code = String(requestUrl.searchParams.get("code") || "");
  const state = String(requestUrl.searchParams.get("state") || "");
  if (!shop || !code || !state) {
    sendRedirect(
      res,
      getCallbackRedirect(req, {
        provider: "shopify",
        shopify: "error",
        message: "Missing callback parameters.",
      })
    );
    return;
  }
  if (!verifyShopifyHmac(requestUrl)) {
    sendRedirect(
      res,
      getCallbackRedirect(req, {
        provider: "shopify",
        shopify: "error",
        message: "Shopify callback validation failed.",
      })
    );
    return;
  }
  const statePayload = consumeOauthState(state, shop);
  if (!statePayload?.userId) {
    sendRedirect(
      res,
      getCallbackRedirect(req, {
        provider: "shopify",
        shopify: "error",
        message: "OAuth session expired. Start Shopify connect again.",
      })
    );
    return;
  }

  try {
    const tokenData = await exchangeShopifyAccessToken(shop, code);
    await upsertShopifyConnection(
      statePayload.userId,
      shop,
      tokenData.access_token,
      tokenData.scope || SHOPIFY_SCOPES.join(",")
    );
    sendRedirect(
      res,
      getCallbackRedirect(req, {
        provider: "shopify",
        shopify: "connected",
        shop,
      })
    );
  } catch (error) {
    sendRedirect(
      res,
      getCallbackRedirect(req, {
        provider: "shopify",
        shopify: "error",
        message: error.message || "Shopify connection failed.",
      })
    );
  }
}

async function handleShopifyConnection(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  try {
    const row = await getShopifyConnection(user.id, "");
    if (!row) {
      sendJson(res, 200, { connected: false, connection: null });
      return;
    }
    sendJson(res, 200, {
      connected: true,
      connection: {
        shop: row.shop_domain,
        status: row.status,
        scopes: row.scopes || "",
        connectedAt: row.connected_at || "",
        updatedAt: row.updated_at || "",
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Failed to load Shopify connection." });
  }
}

async function handleShopifyLocations(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }

  const requestedShop = normalizeShopDomain(requestUrl.searchParams.get("shop"));
  let resolvedShop = requestedShop;
  try {
    const connection = await getShopifyConnection(user.id, requestedShop);
    if (!connection) {
      sendJson(res, 404, { error: "Shopify is not connected for this account." });
      return;
    }
    resolvedShop = String(connection.shop_domain || requestedShop || "").trim();
    const accessToken = decryptToken(connection.access_token);
    const locations = await fetchShopifyLocations(connection.shop_domain, accessToken);
    sendJson(res, 200, {
      shop: connection.shop_domain,
      count: locations.length,
      locations,
    });
  } catch (error) {
    if (error?.status === 401 && resolvedShop) {
      await setShopifyConnectionStatus(user.id, resolvedShop, "token_invalid").catch(() => {});
      sendJson(res, 409, {
        error: "Shopify connection expired or was revoked. Reconnect Shopify and try again.",
      });
      return;
    }
    sendJson(res, 500, { error: error.message || "Failed to load Shopify locations." });
  }
}

async function handleShopifySettingsGet(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }

  const requestedShop = normalizeShopDomain(requestUrl.searchParams.get("shop"));
  try {
    const connection = await getShopifyConnection(user.id, requestedShop, {
      includeSettings: true,
    });
    if (!connection) {
      sendJson(res, 404, { error: "Shopify is not connected for this account." });
      return;
    }
    const selectedLocationIds = getSelectedLocationIdsFromImportSettings(
      connection.import_settings
    );
    sendJson(res, 200, {
      shop: connection.shop_domain,
      settings: {
        selectedLocationIds,
      },
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Failed to load Shopify settings." });
  }
}

async function handleShopifySettingsPost(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const requestedShop = normalizeShopDomain(body?.shop);
  const selectedLocationIds = sanitizeSelectedLocationIds(body?.selectedLocationIds);
  if (!selectedLocationIds.length) {
    sendJson(res, 400, { error: "Select at least one fulfillment location." });
    return;
  }

  try {
    const connection = await getShopifyConnection(user.id, requestedShop, {
      includeSettings: true,
    });
    if (!connection) {
      sendJson(res, 404, { error: "Shopify is not connected for this account." });
      return;
    }
    await saveShopifyImportSettings(user.id, connection.shop_domain, selectedLocationIds);
    sendJson(res, 200, {
      shop: connection.shop_domain,
      settings: {
        selectedLocationIds,
      },
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Failed to save Shopify settings." });
  }
}

async function handleShopifyImportOrders(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const requestedShop = normalizeShopDomain(body?.shop);
  const selectedLocationIds = sanitizeSelectedLocationIds(body?.selectedLocationIds);
  const limit = Number(body?.limit);
  let resolvedShop = requestedShop;

  try {
    const connection = await getShopifyConnection(user.id, requestedShop, {
      includeSettings: true,
    });
    if (!connection) {
      sendJson(res, 404, { error: "Shopify is not connected for this account." });
      return;
    }
    resolvedShop = String(connection.shop_domain || resolvedShop || "").trim();
    const accessToken = decryptToken(connection.access_token);
    const locations = await fetchShopifyLocations(connection.shop_domain, accessToken);
    const savedSelectedLocationIds = getSelectedLocationIdsFromImportSettings(
      connection.import_settings
    );
    let resolvedSelectedLocationIds = selectedLocationIds.length
      ? selectedLocationIds
      : savedSelectedLocationIds;
    if (!resolvedSelectedLocationIds.length) {
      resolvedSelectedLocationIds = locations.map((location) => location.id);
    }
    const locationById = indexLocationsById(locations);
    const orders = await fetchShopifyOrders(connection.shop_domain, accessToken, limit);
    const rows = mapShopifyOrdersToCsvRows(orders, {
      locationById,
      selectedLocationIds: resolvedSelectedLocationIds,
    });
    sendJson(res, 200, {
      shop: connection.shop_domain,
      count: rows.length,
      rows,
    });
  } catch (error) {
    if (error?.status === 401 && resolvedShop) {
      await setShopifyConnectionStatus(user.id, resolvedShop, "token_invalid").catch(() => {});
      sendJson(res, 409, {
        error: "Shopify connection expired or was revoked. Reconnect Shopify and try again.",
      });
      return;
    }
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Shopify import failed." });
  }
}

async function handleApi(req, res, requestUrl) {
  const pathname = requestUrl.pathname.replace(/\/+$/, "") || "/";
  if (pathname === "/api/admin/status" && req.method === "GET") {
    await handleAdminStatus(req, res);
    return true;
  }
  if (pathname === "/api/admin/dashboard" && req.method === "GET") {
    await handleAdminDashboard(req, res);
    return true;
  }
  if (pathname === "/api/admin/settings" && req.method === "POST") {
    await handleAdminSettingsSave(req, res);
    return true;
  }
  if (pathname === "/api/admin/invites/revoke" && req.method === "POST") {
    await handleAdminInviteRevoke(req, res);
    return true;
  }
  if (pathname === "/api/auth/invites" && req.method === "GET") {
    await handleListRegistrationInvites(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/auth/invites" && req.method === "POST") {
    await handleCreateRegistrationInvite(req, res);
    return true;
  }
  if (pathname === "/api/auth/register-invite" && req.method === "GET") {
    await handleRegisterInviteValidate(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/auth/register" && req.method === "POST") {
    await handleRegisterWithInvite(req, res);
    return true;
  }
  if (pathname === "/api/shopify/install-link" && req.method === "POST") {
    await handleShopifyInstallLink(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/shopify/callback" && req.method === "GET") {
    await handleShopifyCallback(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/shopify/connection" && req.method === "GET") {
    await handleShopifyConnection(req, res);
    return true;
  }
  if (
    (pathname === "/api/shopify/locations" || pathname === "/api/shopify/location") &&
    req.method === "GET"
  ) {
    await handleShopifyLocations(req, res, requestUrl);
    return true;
  }
  if (
    (pathname === "/api/shopify/settings" || pathname === "/api/shopify/setting") &&
    req.method === "GET"
  ) {
    await handleShopifySettingsGet(req, res, requestUrl);
    return true;
  }
  if (
    (pathname === "/api/shopify/settings" || pathname === "/api/shopify/setting") &&
    req.method === "POST"
  ) {
    await handleShopifySettingsPost(req, res);
    return true;
  }
  if (pathname === "/api/shopify/import-orders" && req.method === "POST") {
    await handleShopifyImportOrders(req, res);
    return true;
  }
  if (pathname.startsWith("/api/")) {
    sendJson(res, 404, { error: "API route not found." });
    return true;
  }
  return false;
}

const server = http.createServer(async (req, res) => {
  const requestUrl = new URL(req.url || "/", "http://localhost");

  try {
    const apiHandled = await handleApi(req, res, requestUrl);
    if (apiHandled) {
      return;
    }
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Unexpected server error." });
    return;
  }

  let pathname = decodeURIComponent(requestUrl.pathname);

  if (pathname === "/") {
    sendFile(res, INDEX_FILE);
    return;
  }

  const requested = safeJoin(ROOT, pathname.replace(/^\/+/, ""));
  if (!requested) {
    send(res, 403, { "Content-Type": "text/plain; charset=utf-8" }, "Forbidden");
    return;
  }

  fs.stat(requested, (statError, stats) => {
    if (!statError && stats.isFile()) {
      sendFile(res, requested);
      return;
    }

    const normalizedPathname = pathname.replace(/\/+$/, "");
    const hasFileExtension = path.extname(normalizedPathname) !== "";
    if (hasFileExtension) {
      send(res, 404, { "Content-Type": "text/plain; charset=utf-8" }, "Not Found");
      return;
    }

    sendFile(res, INDEX_FILE);
  });
});

server.listen(PORT, HOST, () => {
  console.log(`Shipr app running at http://${HOST}:${PORT}`);
});
