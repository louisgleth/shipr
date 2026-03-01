const MAX_BODY_BYTES = 1024 * 1024;
const OAUTH_STATE_TTL_MS = 10 * 60 * 1000;
const DEFAULT_SHOPIFY_SCOPES = "read_orders,read_locations";
const DEFAULT_SHOPIFY_API_VERSION = "2025-10";
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

const textEncoder = new TextEncoder();

export default {
  async fetch(request, env) {
    try {
      const url = new URL(request.url);
      const pathname = url.pathname.replace(/\/+$/, "") || "/";

      if (request.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: noStoreHeaders(),
        });
      }

      if (pathname === "/api/admin/status" && request.method === "GET") {
        return handleAdminStatus(request, env);
      }
      if (pathname === "/api/admin/dashboard" && request.method === "GET") {
        return handleAdminDashboard(request, env);
      }
      if (pathname === "/api/admin/settings" && request.method === "POST") {
        return handleAdminSettingsSave(request, env);
      }
      if (pathname === "/api/admin/invites/revoke" && request.method === "POST") {
        return handleAdminInviteRevoke(request, env);
      }
      if (pathname === "/api/auth/invites" && request.method === "GET") {
        return handleListRegistrationInvites(request, env, url);
      }
      if (pathname === "/api/auth/invites" && request.method === "POST") {
        return handleCreateRegistrationInvite(request, env);
      }
      if (pathname === "/api/auth/register-invite" && request.method === "GET") {
        return handleRegisterInviteValidate(request, env, url);
      }
      if (pathname === "/api/auth/register" && request.method === "POST") {
        return handleRegisterWithInvite(request, env);
      }

      if (pathname === "/api/shopify/install-link" && request.method === "POST") {
        return handleInstallLink(request, env);
      }
      if (pathname === "/api/shopify/callback" && request.method === "GET") {
        return handleShopifyCallback(request, env, url);
      }
      if (pathname === "/api/shopify/connection" && request.method === "GET") {
        return handleConnection(request, env);
      }
      if (
        (pathname === "/api/shopify/locations" || pathname === "/api/shopify/location") &&
        request.method === "GET"
      ) {
        return handleLocations(request, env, url);
      }
      if (
        (pathname === "/api/shopify/settings" || pathname === "/api/shopify/setting") &&
        request.method === "GET"
      ) {
        return handleSettingsGet(request, env, url);
      }
      if (
        (pathname === "/api/shopify/settings" || pathname === "/api/shopify/setting") &&
        request.method === "POST"
      ) {
        return handleSettingsPost(request, env);
      }
      if (pathname === "/api/shopify/import-orders" && request.method === "POST") {
        return handleImportOrders(request, env);
      }

      return jsonResponse({ error: "API route not found." }, 404);
    } catch (error) {
      return jsonResponse({ error: error?.message || "Unexpected server error." }, 500);
    }
  },
};

function noStoreHeaders(extra = {}) {
  return {
    "Cache-Control": "no-store",
    ...extra,
  };
}

function jsonResponse(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: noStoreHeaders({
      "Content-Type": "application/json; charset=utf-8",
    }),
  });
}

function redirectResponse(location) {
  return new Response(null, {
    status: 302,
    headers: noStoreHeaders({
      Location: location,
    }),
  });
}

function getPublicOrigin(request) {
  const url = new URL(request.url);
  return `${url.protocol}//${url.host}`;
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

function getBearerToken(request) {
  const auth = String(request.headers.get("authorization") || "");
  if (!auth.toLowerCase().startsWith("bearer ")) return "";
  return auth.slice(7).trim();
}

async function readJsonBody(request) {
  const text = await request.text();
  if (!text.trim()) return {};
  if (text.length > MAX_BODY_BYTES) {
    throw new Error("Request body too large.");
  }
  try {
    return JSON.parse(text);
  } catch (_error) {
    throw new Error("Invalid JSON request body.");
  }
}

function getShopifyScopes(env) {
  return String(env.SHOPIFY_SCOPES || DEFAULT_SHOPIFY_SCOPES)
    .split(",")
    .map((scope) => scope.trim())
    .filter(Boolean);
}

function assertConfig(env) {
  if (!env.SHOPIFY_API_KEY || !env.SHOPIFY_API_SECRET) {
    throw new Error("SHOPIFY_API_KEY and SHOPIFY_API_SECRET are required.");
  }
  if (!env.SUPABASE_URL || !env.SUPABASE_PUBLISHABLE_KEY || !env.SUPABASE_SERVICE_ROLE_KEY) {
    throw new Error(
      "SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, and SUPABASE_SERVICE_ROLE_KEY are required."
    );
  }
}

function getStateSecret(env) {
  const secret = String(env.OAUTH_STATE_SECRET || env.SHOPIFY_API_SECRET || "").trim();
  if (!secret) {
    throw new Error("OAUTH_STATE_SECRET is required.");
  }
  return secret;
}

function getTokenSecret(env) {
  const secret = String(env.SHOPIFY_TOKEN_ENCRYPTION_KEY || env.SHOPIFY_API_SECRET || "").trim();
  if (!secret) {
    throw new Error("SHOPIFY_TOKEN_ENCRYPTION_KEY is required.");
  }
  return secret;
}

function toBase64Url(bytes) {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function fromBase64Url(value) {
  const normalized = String(value || "").replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized + "===".slice((normalized.length + 3) % 4);
  const binary = atob(padded);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function timingSafeEqual(a, b) {
  const left = String(a || "");
  const right = String(b || "");
  if (!left || !right || left.length !== right.length) return false;
  let out = 0;
  for (let i = 0; i < left.length; i += 1) {
    out |= left.charCodeAt(i) ^ right.charCodeAt(i);
  }
  return out === 0;
}

async function hmacSha256Hex(secret, message) {
  const key = await crypto.subtle.importKey(
    "raw",
    textEncoder.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const signature = await crypto.subtle.sign("HMAC", key, textEncoder.encode(message));
  const bytes = new Uint8Array(signature);
  return Array.from(bytes)
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

async function makeStateToken(env, payload) {
  const stateSecret = getStateSecret(env);
  const bytes = new Uint8Array(12);
  crypto.getRandomValues(bytes);
  const tokenPayload = {
    ...payload,
    nonce: toBase64Url(bytes),
    exp: Date.now() + OAUTH_STATE_TTL_MS,
  };
  const encoded = toBase64Url(textEncoder.encode(JSON.stringify(tokenPayload)));
  const signature = await hmacSha256Hex(stateSecret, encoded);
  return `${encoded}.${signature}`;
}

async function parseStateToken(env, token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 2) return null;
  const [encoded, signature] = parts;
  if (!encoded || !signature) return null;
  const stateSecret = getStateSecret(env);
  const expected = await hmacSha256Hex(stateSecret, encoded);
  if (!timingSafeEqual(signature, expected)) return null;
  let payload = null;
  try {
    payload = JSON.parse(new TextDecoder().decode(fromBase64Url(encoded)));
  } catch (_error) {
    return null;
  }
  if (!payload || typeof payload !== "object") return null;
  if (!payload.exp || Date.now() > Number(payload.exp)) return null;
  return payload;
}

async function getAesKey(secret) {
  const digest = await crypto.subtle.digest("SHA-256", textEncoder.encode(secret));
  return crypto.subtle.importKey("raw", digest, "AES-GCM", false, ["encrypt", "decrypt"]);
}

async function encryptToken(env, token) {
  const key = await getAesKey(getTokenSecret(env));
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);
  const encrypted = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv },
    key,
    textEncoder.encode(String(token || ""))
  );
  return `v1:${toBase64Url(iv)}:${toBase64Url(new Uint8Array(encrypted))}`;
}

async function decryptToken(env, token) {
  const raw = String(token || "");
  if (!raw.startsWith("v1:")) return raw;
  const parts = raw.split(":");
  if (parts.length !== 3) {
    throw new Error("Invalid encrypted Shopify token.");
  }
  const [, ivEncoded, encryptedEncoded] = parts;
  const key = await getAesKey(getTokenSecret(env));
  const decrypted = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: fromBase64Url(ivEncoded) },
    key,
    fromBase64Url(encryptedEncoded)
  );
  return new TextDecoder().decode(decrypted);
}

async function fetchSupabaseUser(env, accessToken) {
  if (!accessToken) return null;
  const response = await fetch(`${String(env.SUPABASE_URL).replace(/\/+$/, "")}/auth/v1/user`, {
    method: "GET",
    headers: {
      apikey: env.SUPABASE_PUBLISHABLE_KEY,
      Authorization: `Bearer ${accessToken}`,
    },
  });
  if (!response.ok) return null;
  const payload = await response.json().catch(() => null);
  if (!payload || typeof payload !== "object" || !payload.id) return null;
  return payload;
}

async function getAuthenticatedUser(request, env) {
  const token = getBearerToken(request);
  if (!token) return null;
  return fetchSupabaseUser(env, token);
}

async function supabaseServiceRequest(env, pathnameWithQuery, options = {}) {
  const response = await fetch(
    `${String(env.SUPABASE_URL).replace(/\/+$/, "")}${pathnameWithQuery}`,
    {
      ...options,
      headers: {
        apikey: env.SUPABASE_SERVICE_ROLE_KEY,
        Authorization: `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}`,
        ...(options.headers || {}),
      },
    }
  );
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

function parseInviteExpiryDays(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return DEFAULT_INVITE_EXPIRY_DAYS;
  return Math.max(1, Math.min(MAX_INVITE_EXPIRY_DAYS, Math.round(numeric)));
}

function getInviteAdminEmails(env) {
  return String(env.INVITE_ADMIN_EMAILS || "")
    .split(",")
    .map((email) => email.trim().toLowerCase())
    .filter(Boolean);
}

function canManageRegistrationInvites(user, env) {
  if (!user?.id) return false;
  if (user?.user_metadata?.app_admin === true) return true;
  const allowlist = getInviteAdminEmails(env);
  if (!allowlist.length) return true;
  const email = normalizeEmail(user.email || "");
  return Boolean(email && allowlist.includes(email));
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

async function sha256Hex(value) {
  const digest = await crypto.subtle.digest("SHA-256", textEncoder.encode(String(value || "")));
  const bytes = new Uint8Array(digest);
  return Array.from(bytes)
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
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

async function getRegistrationInviteByToken(env, token) {
  const inviteToken = normalizeInviteToken(token);
  if (!inviteToken) return null;
  const tokenHash = await sha256Hex(inviteToken);
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,invited_email,company_name,contact_name,contact_email,contact_phone,billing_address,tax_id,customer_id,plan_name,expires_at,claimed_at,revoked_at"
  );
  params.set("token_hash", `eq.${tokenHash}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    env,
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

function createInviteToken() {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  return Array.from(bytes)
    .map((value) => value.toString(16).padStart(2, "0"))
    .join("");
}

async function insertRegistrationInvite(
  env,
  { tokenHash, tokenEncrypted, invitedEmail, expiresAt, createdBy }
) {
  const response = await supabaseServiceRequest(env, `/rest/v1/${REGISTRATION_INVITES_TABLE}`, {
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

async function createRegistrationInvite(env, { invitedEmail, expiresInDays, createdBy }) {
  const safeInvitedEmail = normalizeEmail(invitedEmail || "");
  const expiryDays = parseInviteExpiryDays(expiresInDays);
  const expiresAt = new Date(Date.now() + expiryDays * 24 * 60 * 60 * 1000).toISOString();

  for (let attempt = 0; attempt < 3; attempt += 1) {
    const token = createInviteToken();
    const tokenHash = await sha256Hex(token);
    try {
      await insertRegistrationInvite(env, {
        tokenHash,
        tokenEncrypted: await encryptToken(env, token),
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

async function listRegistrationInvites(env, { createdBy, limit = 20 }) {
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
    env,
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

async function revokeRegistrationInvite(env, inviteId, revokedBy) {
  const safeInviteId = String(inviteId || "").trim();
  if (!safeInviteId) {
    throw new Error("Invite id is required.");
  }

  const selectParams = new URLSearchParams();
  selectParams.set("select", "id,claimed_at,revoked_at");
  selectParams.set("id", `eq.${safeInviteId}`);
  selectParams.set("limit", "1");

  const existingResponse = await supabaseServiceRequest(
    env,
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
    env,
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

async function listSupabaseUsers(env, limit = 200) {
  const safeLimit = Math.max(1, Math.min(1000, Number(limit) || 200));
  const pageSize = Math.min(100, safeLimit);
  const users = [];
  for (let page = 1; users.length < safeLimit && page <= 20; page += 1) {
    const response = await fetch(
      `${String(env.SUPABASE_URL).replace(/\/+$/, "")}/auth/v1/admin/users?page=${page}&per_page=${pageSize}`,
      {
        method: "GET",
        headers: {
          apikey: env.SUPABASE_SERVICE_ROLE_KEY,
          Authorization: `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}`,
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

async function listGenerationHistoryRows(env, limit = 5000) {
  const safeLimit = Math.max(1, Math.min(20000, Number(limit) || 5000));
  const pageSize = 1000;
  const rows = [];
  for (let start = 0; start < safeLimit; start += pageSize) {
    const end = Math.min(safeLimit - 1, start + pageSize - 1);
    const params = new URLSearchParams();
    params.set("select", "user_id,created_at,service_type,quantity,total_price,payload");
    params.set("order", "created_at.desc");
    const response = await supabaseServiceRequest(
      env,
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

async function getAdminSettings(env) {
  const params = new URLSearchParams();
  params.set("select", "scope,carrier_discount_pct,client_discount_pct,updated_at,updated_by");
  params.set("scope", `eq.${ADMIN_SETTINGS_SCOPE}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    env,
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

async function saveAdminSettings(env, userId, payload) {
  const normalized = normalizeAdminSettings(payload);
  const response = await supabaseServiceRequest(
    env,
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

async function buildInviteUrl(request, env, invite) {
  const encrypted = String(invite?.token_encrypted || "").trim();
  if (!encrypted) return "";
  try {
    const token = await decryptToken(env, encrypted);
    const inviteUrl = new URL("/register", getPublicOrigin(request));
    inviteUrl.searchParams.set("invite", token);
    return inviteUrl.toString();
  } catch (_error) {
    return "";
  }
}

async function mapInviteHistoryRow(request, env, invite) {
  return {
    id: invite?.id || null,
    invited_email: normalizeEmail(invite?.invited_email || ""),
    expires_at: invite?.expires_at || null,
    created_at: invite?.created_at || null,
    claimed_at: invite?.claimed_at || null,
    claimed_email: normalizeEmail(invite?.claimed_email || ""),
    revoked_at: invite?.revoked_at || null,
    invite_url: await buildInviteUrl(request, env, invite),
  };
}

function shouldIncludeAdminClient(user, env) {
  if (!user?.id) return false;
  if (user?.user_metadata?.app_admin === true) return false;
  const email = normalizeEmail(user.email || "");
  if (email && getInviteAdminEmails(env).includes(email)) return false;
  return true;
}

async function buildAdminDashboard(request, env) {
  const [settings, invites, users, historyRows] = await Promise.all([
    getAdminSettings(env),
    listRegistrationInvites(env, { limit: 50 }),
    listSupabaseUsers(env, 250),
    listGenerationHistoryRows(env, 10000),
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
    .filter((user) => shouldIncludeAdminClient(user, env))
    .map((user) => ({
      user: {
        id: user.id,
        email: normalizeEmail(user.email || ""),
        created_at: user.created_at || null,
        user_metadata:
          user.user_metadata && typeof user.user_metadata === "object" ? user.user_metadata : {},
      },
      metrics: buildAdminClientMetrics(user, historyByUserId.get(user.id) || [], settings),
    }))
    .sort((left, right) => {
      const leftValue = Date.parse(left?.metrics?.last_generation_at || left?.user?.created_at || 0);
      const rightValue = Date.parse(right?.metrics?.last_generation_at || right?.user?.created_at || 0);
      return rightValue - leftValue;
    });

  const inviteHistory = await Promise.all(
    invites.map((invite) => mapInviteHistoryRow(request, env, invite))
  );

  return {
    settings,
    invites: inviteHistory,
    clients,
    summary: buildAdminClientSummary(clients, invites),
  };
}

async function markRegistrationInviteClaimed(env, inviteId, userId, email) {
  if (!inviteId || !userId) return;
  const params = new URLSearchParams();
  params.set("id", `eq.${inviteId}`);
  const response = await supabaseServiceRequest(
    env,
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

async function ensureAccountSettingsRow(env, userId) {
  if (!userId) return;
  const response = await supabaseServiceRequest(env, "/rest/v1/account_settings?on_conflict=user_id", {
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
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed creating account settings (${response.status}) ${details}`.trim());
  }
}

async function createSupabaseUserWithMetadata(env, email, password, userMetadata) {
  const response = await fetch(`${String(env.SUPABASE_URL).replace(/\/+$/, "")}/auth/v1/admin/users`, {
    method: "POST",
    headers: {
      apikey: env.SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}`,
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

async function upsertShopifyConnection(env, payload) {
  const response = await supabaseServiceRequest(
    env,
    "/rest/v1/provider_connections?on_conflict=user_id,provider,shop_domain",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify([payload]),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed saving Shopify connection (${response.status}) ${details}`.trim());
  }
}

async function getShopifyConnection(env, userId, shop, options = {}) {
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
  if (shop) {
    params.set("shop_domain", `eq.${shop}`);
  }
  params.set("order", "updated_at.desc");
  params.set("limit", "1");

  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/provider_connections?${params.toString()}`,
    {
      method: "GET",
    }
  );
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
  if (!Array.isArray(rows) || !rows.length) return null;
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

async function saveShopifyImportSettings(env, userId, shop, selectedLocationIds) {
  const params = new URLSearchParams();
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.shopify");
  params.set("shop_domain", `eq.${shop}`);
  params.set("status", "eq.connected");
  const response = await supabaseServiceRequest(
    env,
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

async function setShopifyConnectionStatus(env, userId, shop, status) {
  const params = new URLSearchParams();
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.shopify");
  params.set("shop_domain", `eq.${shop}`);
  const response = await supabaseServiceRequest(
    env,
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
    const packageWeight =
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
      packageWeight,
      packageDims: "25 x 20 x 10",
    };
  });
}

async function fetchShopifyLocations(env, shop, accessToken) {
  const apiVersion = String(env.SHOPIFY_API_VERSION || DEFAULT_SHOPIFY_API_VERSION);
  const locationsUrl = new URL(`https://${shop}/admin/api/${apiVersion}/locations.json`);
  locationsUrl.searchParams.set("limit", "250");
  locationsUrl.searchParams.set(
    "fields",
    "id,name,address1,address2,city,province,province_code,zip,country,country_code"
  );
  const response = await fetch(locationsUrl.toString(), {
    method: "GET",
    headers: {
      "X-Shopify-Access-Token": accessToken,
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

async function verifyShopifyCallbackHmac(env, callbackUrl) {
  const received = String(callbackUrl.searchParams.get("hmac") || "")
    .trim()
    .toLowerCase();
  if (!received || !/^[a-f0-9]+$/.test(received)) return false;

  const entries = [];
  callbackUrl.searchParams.forEach((value, key) => {
    if (key === "hmac" || key === "signature") return;
    entries.push([key, value]);
  });
  entries.sort(([left], [right]) => left.localeCompare(right));
  const message = entries.map(([key, value]) => `${key}=${value}`).join("&");
  const expected = await hmacSha256Hex(env.SHOPIFY_API_SECRET, message);
  return timingSafeEqual(received, expected);
}

function getCallbackRedirect(request, params) {
  const url = new URL("/label-info", getPublicOrigin(request));
  Object.entries(params).forEach(([key, value]) => {
    if (value === undefined || value === null || value === "") return;
    url.searchParams.set(key, String(value));
  });
  return url.toString();
}

async function handleCreateRegistrationInvite(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to create client invites." }, 403);
  }

  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error.message || "Invalid request body." }, 400);
  }

  const invitedEmail = normalizeEmail(body?.invitedEmail || "");
  if (!invitedEmail) {
    return jsonResponse({ error: "Client email is required." }, 400);
  }
  if (!isValidEmailFormat(invitedEmail)) {
    return jsonResponse({ error: "Invalid client email format." }, 400);
  }

  try {
    const created = await createRegistrationInvite(env, {
      invitedEmail,
      expiresInDays: body?.expiresInDays,
      createdBy: user.id,
    });
    const inviteUrl = new URL("/register", getPublicOrigin(request));
    inviteUrl.searchParams.set("invite", created.token);
    return jsonResponse({
      ok: true,
      inviteUrl: inviteUrl.toString(),
      expiresAt: created.expiresAt,
      invitedEmail: invitedEmail || null,
    });
  } catch (error) {
    return jsonResponse({ error: error.message || "Could not create invite." }, 500);
  }
}

async function handleListRegistrationInvites(request, env, url) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to create client invites." }, 403);
  }

  const limit = Number(url.searchParams.get("limit") || "20");
  try {
    const invites = await listRegistrationInvites(env, {
      createdBy: null,
      limit,
    });
    const payload = await Promise.all(
      invites.map((invite) => mapInviteHistoryRow(request, env, invite))
    );
    return jsonResponse({ invites: payload });
  } catch (error) {
    return jsonResponse({ error: error.message || "Could not load invite history." }, 500);
  }
}

async function handleAdminStatus(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  return jsonResponse({ allowed: canManageRegistrationInvites(user, env) });
}

async function handleAdminDashboard(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to access the admin panel." }, 403);
  }
  try {
    const dashboard = await buildAdminDashboard(request, env);
    return jsonResponse(dashboard);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not load admin dashboard." }, 500);
  }
}

async function handleAdminSettingsSave(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to update admin settings." }, 403);
  }

  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error.message || "Invalid request body." }, 400);
  }

  try {
    const settings = await saveAdminSettings(env, user.id, body || {});
    return jsonResponse({ ok: true, settings });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not save admin settings." }, 500);
  }
}

async function handleAdminInviteRevoke(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to revoke client invites." }, 403);
  }

  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error.message || "Invalid request body." }, 400);
  }

  const inviteId = String(body?.inviteId || "").trim();
  if (!inviteId) {
    return jsonResponse({ error: "Invite id is required." }, 400);
  }

  try {
    await revokeRegistrationInvite(env, inviteId, user.id);
    return jsonResponse({ ok: true });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not revoke invite." }, 500);
  }
}

async function handleRegisterInviteValidate(_request, env, url) {
  const token = normalizeInviteToken(url.searchParams.get("token"));
  if (!token) {
    return jsonResponse({ error: "Registration link required." }, 400);
  }
  try {
    const invite = await getRegistrationInviteByToken(env, token);
    if (!invite || invite.claimed_at || inviteIsRevoked(invite) || inviteIsExpired(invite)) {
      return jsonResponse({ error: "This registration link is invalid or expired." }, 410);
    }
    return jsonResponse({
      valid: true,
      invite: mapInviteToPublic(invite),
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not validate registration link." }, 500);
  }
}

async function handleRegisterWithInvite(request, env) {
  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Invalid request body." }, 400);
  }

  const token = normalizeInviteToken(body?.token);
  const email = normalizeEmail(body?.email);
  const password = String(body?.password || "");
  const profile = normalizeRegistrationProfile(body?.profile);
  const preferredLanguage = String(body?.preferredLanguage || "en").trim().toLowerCase();

  if (!token) {
    return jsonResponse({ error: "Registration link required." }, 400);
  }
  if (!email || !password) {
    return jsonResponse({ error: "Email and password are required." }, 400);
  }
  if (password.length < PASSWORD_MIN_LENGTH) {
    return jsonResponse(
      { error: `Password must be at least ${PASSWORD_MIN_LENGTH} characters.` },
      400
    );
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
    return jsonResponse(
      { error: "All registration fields are required except Customer ID." },
      400
    );
  }

  try {
    const invite = await getRegistrationInviteByToken(env, token);
    if (!invite || invite.claimed_at || inviteIsRevoked(invite) || inviteIsExpired(invite)) {
      return jsonResponse({ error: "This registration link is invalid or expired." }, 410);
    }

    const metadata = buildRegistrationMetadata(invite, profile, email, preferredLanguage);
    const createdUser = await createSupabaseUserWithMetadata(env, email, password, metadata);
    await ensureAccountSettingsRow(env, createdUser.id);
    await markRegistrationInviteClaimed(env, invite.id, createdUser.id, email);

    return jsonResponse({
      ok: true,
      email,
      userId: createdUser.id,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not complete registration." }, 500);
  }
}

async function handleInstallLink(request, env) {
  try {
    assertConfig(env);
    const user = await getAuthenticatedUser(request, env);
    if (!user?.id) {
      return jsonResponse({ error: "Authentication required." }, 401);
    }
    const body = await readJsonBody(request);
    const shop = normalizeShopDomain(body?.shop);
    if (!shop) {
      return jsonResponse(
        { error: "Enter a valid Shopify domain (store.myshopify.com)." },
        400
      );
    }

    const state = await makeStateToken(env, {
      userId: user.id,
      shop,
    });
    const callbackUrl = `${getPublicOrigin(request)}/api/shopify/callback`;
    const installUrl = new URL(`https://${shop}/admin/oauth/authorize`);
    installUrl.searchParams.set("client_id", env.SHOPIFY_API_KEY);
    installUrl.searchParams.set("scope", getShopifyScopes(env).join(","));
    installUrl.searchParams.set("redirect_uri", callbackUrl);
    installUrl.searchParams.set("state", state);

    return jsonResponse({
      url: installUrl.toString(),
      shop,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not start Shopify install flow." }, 500);
  }
}

async function handleShopifyCallback(request, env, callbackUrl) {
  try {
    assertConfig(env);
    const shop = normalizeShopDomain(callbackUrl.searchParams.get("shop"));
    const code = String(callbackUrl.searchParams.get("code") || "");
    const stateToken = String(callbackUrl.searchParams.get("state") || "");
    if (!shop || !code || !stateToken) {
      return redirectResponse(
        getCallbackRedirect(request, {
          provider: "shopify",
          shopify: "error",
          message: "Missing callback parameters.",
        })
      );
    }

    const hmacValid = await verifyShopifyCallbackHmac(env, callbackUrl);
    if (!hmacValid) {
      return redirectResponse(
        getCallbackRedirect(request, {
          provider: "shopify",
          shopify: "error",
          message: "Shopify callback validation failed.",
        })
      );
    }

    const statePayload = await parseStateToken(env, stateToken);
    if (!statePayload?.userId || statePayload.shop !== shop) {
      return redirectResponse(
        getCallbackRedirect(request, {
          provider: "shopify",
          shopify: "error",
          message: "OAuth session expired. Start Shopify connect again.",
        })
      );
    }

    const tokenResponse = await fetch(`https://${shop}/admin/oauth/access_token`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        client_id: env.SHOPIFY_API_KEY,
        client_secret: env.SHOPIFY_API_SECRET,
        code,
      }),
    });
    if (!tokenResponse.ok) {
      const details = await tokenResponse.text().catch(() => "");
      throw new Error(`Shopify token exchange failed (${tokenResponse.status}) ${details}`.trim());
    }
    const tokenPayload = await tokenResponse.json().catch(() => null);
    if (!tokenPayload || !tokenPayload.access_token) {
      throw new Error("Shopify token response was invalid.");
    }

    const nowIso = new Date().toISOString();
    await upsertShopifyConnection(env, {
      user_id: statePayload.userId,
      provider: "shopify",
      shop_domain: shop,
      status: "connected",
      scopes: String(tokenPayload.scope || getShopifyScopes(env).join(",")),
      access_token: await encryptToken(env, tokenPayload.access_token),
      updated_at: nowIso,
      connected_at: nowIso,
    });

    return redirectResponse(
      getCallbackRedirect(request, {
        provider: "shopify",
        shopify: "connected",
        shop,
      })
    );
  } catch (error) {
    return redirectResponse(
      getCallbackRedirect(request, {
        provider: "shopify",
        shopify: "error",
        message: error?.message || "Shopify connection failed.",
      })
    );
  }
}

async function handleConnection(request, env) {
  try {
    assertConfig(env);
    const user = await getAuthenticatedUser(request, env);
    if (!user?.id) {
      return jsonResponse({ error: "Authentication required." }, 401);
    }
    const row = await getShopifyConnection(env, user.id, "");
    if (!row) {
      return jsonResponse({ connected: false, connection: null });
    }
    return jsonResponse({
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
    return jsonResponse({ error: error?.message || "Could not load Shopify connection." }, 500);
  }
}

async function handleLocations(request, env, requestUrl) {
  let authUserId = "";
  let connectionShop = "";
  try {
    assertConfig(env);
    const user = await getAuthenticatedUser(request, env);
    if (!user?.id) {
      return jsonResponse({ error: "Authentication required." }, 401);
    }
    authUserId = user.id;

    const requestedShop = normalizeShopDomain(requestUrl.searchParams.get("shop"));
    const connection = await getShopifyConnection(env, user.id, requestedShop);
    if (!connection) {
      return jsonResponse({ error: "Shopify is not connected for this account." }, 404);
    }
    connectionShop = String(connection.shop_domain || requestedShop || "").trim();
    const accessToken = await decryptToken(env, connection.access_token);
    const locations = await fetchShopifyLocations(env, connectionShop, accessToken);
    return jsonResponse({
      shop: connectionShop,
      count: locations.length,
      locations,
    });
  } catch (error) {
    if (error?.status === 401) {
      if (authUserId && connectionShop) {
        await setShopifyConnectionStatus(env, authUserId, connectionShop, "token_invalid").catch(
          () => {}
        );
      }
      return jsonResponse(
        {
          error: "Shopify connection expired or was revoked. Reconnect Shopify and try again.",
        },
        409
      );
    }
    return jsonResponse(
      { error: error?.message || "Failed to load Shopify locations." },
      500
    );
  }
}

async function handleSettingsGet(request, env, requestUrl) {
  try {
    assertConfig(env);
    const user = await getAuthenticatedUser(request, env);
    if (!user?.id) {
      return jsonResponse({ error: "Authentication required." }, 401);
    }

    const requestedShop = normalizeShopDomain(requestUrl.searchParams.get("shop"));
    const connection = await getShopifyConnection(env, user.id, requestedShop, {
      includeSettings: true,
    });
    if (!connection) {
      return jsonResponse({ error: "Shopify is not connected for this account." }, 404);
    }
    return jsonResponse({
      shop: connection.shop_domain,
      settings: {
        selectedLocationIds: getSelectedLocationIdsFromImportSettings(
          connection.import_settings
        ),
      },
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      return jsonResponse({ error: error.message }, 500);
    }
    return jsonResponse({ error: error?.message || "Failed to load Shopify settings." }, 500);
  }
}

async function handleSettingsPost(request, env) {
  try {
    assertConfig(env);
    const user = await getAuthenticatedUser(request, env);
    if (!user?.id) {
      return jsonResponse({ error: "Authentication required." }, 401);
    }

    const body = await readJsonBody(request);
    const requestedShop = normalizeShopDomain(body?.shop);
    const selectedLocationIds = sanitizeSelectedLocationIds(body?.selectedLocationIds);
    if (!selectedLocationIds.length) {
      return jsonResponse({ error: "Select at least one fulfillment location." }, 400);
    }

    const connection = await getShopifyConnection(env, user.id, requestedShop, {
      includeSettings: true,
    });
    if (!connection) {
      return jsonResponse({ error: "Shopify is not connected for this account." }, 404);
    }
    await saveShopifyImportSettings(env, user.id, connection.shop_domain, selectedLocationIds);
    return jsonResponse({
      shop: connection.shop_domain,
      settings: {
        selectedLocationIds,
      },
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      return jsonResponse({ error: error.message }, 500);
    }
    return jsonResponse({ error: error?.message || "Failed to save Shopify settings." }, 500);
  }
}

async function handleImportOrders(request, env) {
  try {
    assertConfig(env);
    const user = await getAuthenticatedUser(request, env);
    if (!user?.id) {
      return jsonResponse({ error: "Authentication required." }, 401);
    }
    const body = await readJsonBody(request);
    const shop = normalizeShopDomain(body?.shop);
    const selectedLocationIds = sanitizeSelectedLocationIds(body?.selectedLocationIds);
    const limitRaw = Number(body?.limit);
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(250, Math.trunc(limitRaw))) : 50;

    const connection = await getShopifyConnection(env, user.id, shop, {
      includeSettings: true,
    });
    if (!connection) {
      return jsonResponse({ error: "Shopify is not connected for this account." }, 404);
    }

    const accessToken = await decryptToken(env, connection.access_token);
    const locations = await fetchShopifyLocations(env, connection.shop_domain, accessToken);
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
    const apiVersion = String(env.SHOPIFY_API_VERSION || DEFAULT_SHOPIFY_API_VERSION);
    const ordersUrl = new URL(`https://${connection.shop_domain}/admin/api/${apiVersion}/orders.json`);
    ordersUrl.searchParams.set("status", "open");
    ordersUrl.searchParams.set("limit", String(limit));
    ordersUrl.searchParams.set(
      "fields",
      "id,name,created_at,total_weight,shipping_address,currency,current_total_price,total_price,location_id,origin_location,line_items,fulfillments"
    );

    const ordersResponse = await fetch(ordersUrl.toString(), {
      method: "GET",
      headers: {
        "X-Shopify-Access-Token": accessToken,
      },
    });
    if (ordersResponse.status === 401) {
      await setShopifyConnectionStatus(env, user.id, connection.shop_domain, "token_invalid").catch(
        () => {}
      );
      return jsonResponse(
        {
          error: "Shopify connection expired or was revoked. Reconnect Shopify and try again.",
        },
        409
      );
    }
    if (!ordersResponse.ok) {
      const details = await ordersResponse.text().catch(() => "");
      throw new Error(`Shopify order import failed (${ordersResponse.status}) ${details}`.trim());
    }
    const payload = await ordersResponse.json().catch(() => null);
    const rows = mapShopifyOrdersToCsvRows(
      Array.isArray(payload?.orders) ? payload.orders : [],
      {
        locationById,
        selectedLocationIds: resolvedSelectedLocationIds,
      }
    );
    return jsonResponse({
      shop: connection.shop_domain,
      count: rows.length,
      rows,
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      return jsonResponse({ error: error.message }, 500);
    }
    return jsonResponse({ error: error?.message || "Shopify import failed." }, 500);
  }
}
