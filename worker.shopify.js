const MAX_BODY_BYTES = 1024 * 1024;
const OAUTH_STATE_TTL_MS = 10 * 60 * 1000;
const DEFAULT_SHOPIFY_SCOPES = "read_orders,read_locations";
const DEFAULT_SHOPIFY_API_VERSION = "2025-10";
const REGISTRATION_INVITES_TABLE = "registration_invites";
const CLICKWRAP_CONTRACTS_TABLE = "clickwrap_contract_versions";
const CLICKWRAP_ACCEPTANCES_TABLE = "clickwrap_acceptances";
const ADMIN_SETTINGS_TABLE = "admin_settings";
const CLIENT_BILLING_PREF_TABLE = "client_billing_preferences";
const HISTORY_TABLE = "label_generations";
const BILLING_INVOICES_TABLE = "billing_invoices";
const BILLING_INVOICE_ITEMS_TABLE = "billing_invoice_items";
const BILLING_INVOICE_EVENTS_TABLE = "billing_invoice_events";
const BILLING_WALLET_TABLE = "billing_wallets";
const BILLING_WALLET_TOPUPS_TABLE = "billing_wallet_topups";
const BILLING_WALLET_TRANSACTIONS_TABLE = "billing_wallet_transactions";
const PASSWORD_MIN_LENGTH = 10;
const DEFAULT_INVITE_EXPIRY_DAYS = 14;
const MAX_INVITE_EXPIRY_DAYS = 90;
const CLICKWRAP_ACCEPTANCE_MAX_AGE_MS = 24 * 60 * 60 * 1000;
const CLICKWRAP_CLOCK_SKEW_MS = 5 * 60 * 1000;
const DEFAULT_CLICKWRAP_CONTRACT_PDF_URL = "/assets/contracts/commAgreement-v1.pdf";
const ADMIN_SETTINGS_SCOPE = "global";
const DEFAULT_BILLING_TERMS_DAYS = 30;
const DEFAULT_VAT_RATE = 0.21;
const DEFAULT_BILLING_CURRENCY = "EUR";
const DEFAULT_IBAN_BENEFICIARY = "Shipide Logistics SRL";
const DEFAULT_IBAN = "BE68 5390 0754 7034";
const DEFAULT_IBAN_BIC = "KREDBEBB";
const DEFAULT_IBAN_TRANSFER_NOTE =
  "Transfers are credited once received (typically 1-2 business days).";
const DEFAULT_ADMIN_SETTINGS = {
  carrier_discount_pct: 25,
  client_discount_pct: 20,
};
const DEFAULT_CLIENT_BILLING_PREF = {
  invoice_enabled: false,
  card_enabled: true,
};
const DEFAULT_CLICKWRAP_CONTRACT_VERSION = "2026-03-13.1";
const DEFAULT_CLICKWRAP_CONTRACT_TITLE = "Shipide Portal Service Agreement";
const DEFAULT_CLICKWRAP_CONTRACT_BODY = `Electronic Acceptance Notice

By clicking "I agree", you consent to use an electronic signature and electronic records for this agreement.

1. Scope of Services
Shipide provides software to prepare shipment data, request transport labels, and manage related billing and reporting operations.

2. Account Security
You must keep credentials confidential and promptly report unauthorized access. Activity under your account is treated as authorized unless proven otherwise.

3. Billing and Payment
Charges may apply per label, shipment, subscription, or wallet usage. Applicable taxes, including VAT, are charged where required by law.

4. Wallet and Top Ups
Wallet balances are pre-funded. Processing time may vary by payment rail. You are responsible for using the exact transfer reference provided.

5. Acceptable Use
You will not use the platform for illegal goods, sanctioned parties, fraudulent shipments, or any prohibited postal content.

6. Data and Retention
Operational shipment and billing records are retained for compliance, service delivery, and dispute resolution.

7. Third-Party Integrations
Provider integrations such as ecommerce platforms and carriers are subject to each provider terms and service availability.

8. Service Availability
Shipide targets high availability but does not guarantee uninterrupted service. Planned maintenance and provider downtime may occur.

9. Limitation of Liability
To the maximum extent permitted by law, liability is limited to direct damages and capped to the fees paid for the affected billing period.

10. Governing Law and Forum
Unless otherwise agreed in writing, this agreement is governed by Belgian law. Disputes may be brought before competent courts in Belgium.

11. Updates to Terms
Updated terms may be published with a new version number and effective date. Continued use after notice constitutes acceptance.

12. Contact
Questions related to this agreement can be sent to legal@shipide.com.

By selecting "I agree", you confirm that you had the opportunity to review the full text and agree to be legally bound.`;
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
      if (pathname === "/api/admin/client-billing" && request.method === "POST") {
        return handleAdminClientBillingSave(request, env);
      }
      if (pathname === "/api/admin/invoices" && request.method === "GET") {
        return handleAdminInvoicesList(request, env, url);
      }
      if (pathname === "/api/admin/invoices/run" && request.method === "POST") {
        return handleAdminInvoicesRun(request, env);
      }
      if (pathname === "/api/admin/invoices/send" && request.method === "POST") {
        return handleAdminInvoiceSend(request, env);
      }
      if (pathname === "/api/admin/invoices/send-test" && request.method === "POST") {
        return handleAdminInvoiceSendTest(request, env);
      }
      if (pathname === "/api/admin/invoices/send-test-sequence" && request.method === "POST") {
        return handleAdminInvoiceSendTestSequence(request, env);
      }
      if (pathname === "/api/admin/reports/send-test" && request.method === "POST") {
        return handleAdminReportsSendTest(request, env);
      }
      if (pathname === "/api/admin/invoices/mark-paid" && request.method === "POST") {
        return handleAdminInvoiceMarkPaid(request, env);
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
      if (pathname === "/api/billing/overview" && request.method === "GET") {
        return handleBillingOverview(request, env);
      }
      if (pathname === "/api/billing/topups" && request.method === "GET") {
        return handleBillingTopups(request, env, url);
      }
      if (pathname === "/api/billing/topups/request" && request.method === "POST") {
        return handleBillingTopupRequest(request, env);
      }
      if (pathname === "/api/billing/checkout" && request.method === "POST") {
        return handleBillingCheckout(request, env);
      }

      return jsonResponse({ error: "API route not found." }, 404);
    } catch (error) {
      return jsonResponse({ error: error?.message || "Unexpected server error." }, 500);
    }
  },
  async scheduled(_event, env) {
    try {
      await runInvoiceReminders(env, { limit: 500 });
    } catch (_error) {
      // Scheduled runs should not throw unhandled errors.
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

function parseNumeric(value) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
}

function parseIsoTimestamp(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  const time = Date.parse(text);
  if (!Number.isFinite(time)) return null;
  return new Date(time).toISOString();
}

function extractSupabaseError(details) {
  const raw = String(details || "").trim();
  if (!raw) return { raw: "", code: "", message: "" };
  try {
    const parsed = JSON.parse(raw);
    return {
      raw,
      code: String(parsed?.code || "").trim(),
      message: String(parsed?.message || parsed?.error || parsed?.details || "").trim(),
    };
  } catch (_error) {
    return { raw, code: "", message: raw };
  }
}

function isMissingClickwrapSchemaError(details) {
  const parsed = extractSupabaseError(details);
  const haystack = `${parsed.raw} ${parsed.code} ${parsed.message}`.toLowerCase();
  return (
    haystack.includes("42p01") &&
      (haystack.includes(CLICKWRAP_CONTRACTS_TABLE) ||
        haystack.includes(CLICKWRAP_ACCEPTANCES_TABLE)) ||
    (haystack.includes("does not exist") &&
      (haystack.includes(CLICKWRAP_CONTRACTS_TABLE) ||
        haystack.includes(CLICKWRAP_ACCEPTANCES_TABLE)))
  );
}

function getEmbeddedClickwrapContract() {
  return {
    id: null,
    version: DEFAULT_CLICKWRAP_CONTRACT_VERSION,
    title: DEFAULT_CLICKWRAP_CONTRACT_TITLE,
    body_text: DEFAULT_CLICKWRAP_CONTRACT_BODY,
    hash_sha256: "",
    pdf_url: DEFAULT_CLICKWRAP_CONTRACT_PDF_URL,
    effective_at: null,
    source: "embedded",
  };
}

async function getActiveClickwrapContract(env) {
  const params = new URLSearchParams();
  params.set("select", "id,version,title,body_text,hash_sha256,effective_at,is_active,created_at");
  params.set("is_active", "eq.true");
  params.set("order", "effective_at.desc.nullslast,created_at.desc");
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${CLICKWRAP_CONTRACTS_TABLE}?${params.toString()}`,
    {
      method: "GET",
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (isMissingClickwrapSchemaError(details)) {
      const embedded = getEmbeddedClickwrapContract();
      return {
        ...embedded,
        hash_sha256: await sha256Hex(embedded.body_text),
      };
    }
    throw new Error(`Could not load click-wrap agreement (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  const contract = Array.isArray(rows) && rows.length ? rows[0] : getEmbeddedClickwrapContract();
  const bodyText = String(contract?.body_text || DEFAULT_CLICKWRAP_CONTRACT_BODY).trim();
  const hash = String(contract?.hash_sha256 || "").trim().toLowerCase() || (await sha256Hex(bodyText));
  return {
    id: contract?.id || null,
    version: String(contract?.version || DEFAULT_CLICKWRAP_CONTRACT_VERSION).trim(),
    title: String(contract?.title || DEFAULT_CLICKWRAP_CONTRACT_TITLE).trim(),
    body_text: bodyText,
    hash_sha256: hash,
    pdf_url: String(contract?.pdf_url || DEFAULT_CLICKWRAP_CONTRACT_PDF_URL).trim(),
    effective_at: contract?.effective_at || null,
    source: contract?.source || "supabase",
  };
}

function mapClickwrapContractToPublic(contract) {
  return {
    id: contract?.id || null,
    version: String(contract?.version || DEFAULT_CLICKWRAP_CONTRACT_VERSION).trim(),
    title: String(contract?.title || DEFAULT_CLICKWRAP_CONTRACT_TITLE).trim(),
    bodyText: String(contract?.body_text || DEFAULT_CLICKWRAP_CONTRACT_BODY).trim(),
    hash: String(contract?.hash_sha256 || "").trim().toLowerCase(),
    pdfUrl: String(contract?.pdf_url || DEFAULT_CLICKWRAP_CONTRACT_PDF_URL).trim(),
    effectiveAt: contract?.effective_at || null,
  };
}

function normalizeClickwrapAgreement(rawAgreement) {
  const agreement =
    rawAgreement && typeof rawAgreement === "object" ? rawAgreement : {};
  return {
    contractId: String(agreement?.contractId || "").trim() || null,
    contractVersion: String(agreement?.contractVersion || "").trim(),
    contractHash: String(agreement?.contractHash || "").trim().toLowerCase(),
    scrolledToEnd: agreement?.scrolledToEnd === true,
    scrolledToEndAt: parseIsoTimestamp(agreement?.scrolledToEndAt),
    agreed: agreement?.agreed === true,
    agreedAt: parseIsoTimestamp(agreement?.agreedAt),
    maxScrollProgress: parseNumeric(agreement?.maxScrollProgress),
    clientTimezone: String(agreement?.clientTimezone || "").trim(),
    clientLocale: String(agreement?.clientLocale || "").trim(),
  };
}

function validateClickwrapAgreement(agreement, contract) {
  if (!agreement?.agreed) {
    return "Agreement acceptance is required.";
  }
  if (!agreement?.scrolledToEnd) {
    return "You must scroll through the full agreement before accepting.";
  }
  if (!agreement?.contractVersion || agreement.contractVersion !== String(contract?.version || "")) {
    return "Agreement version mismatch. Refresh and try again.";
  }
  if (!agreement?.contractHash || agreement.contractHash !== String(contract?.hash_sha256 || "")) {
    return "Agreement integrity check failed. Refresh and try again.";
  }
  if (agreement.contractId && contract?.id && agreement.contractId !== String(contract.id)) {
    return "Agreement identifier mismatch. Refresh and try again.";
  }
  if (!agreement?.agreedAt || !agreement?.scrolledToEndAt) {
    return "Agreement timestamps are missing.";
  }
  const agreedAtMs = Date.parse(agreement.agreedAt);
  const scrolledAtMs = Date.parse(agreement.scrolledToEndAt);
  if (!Number.isFinite(agreedAtMs) || !Number.isFinite(scrolledAtMs)) {
    return "Agreement timestamps are invalid.";
  }
  if (scrolledAtMs - agreedAtMs > CLICKWRAP_CLOCK_SKEW_MS) {
    return "Agreement scroll checkpoint is invalid.";
  }
  if (Date.now() - agreedAtMs > CLICKWRAP_ACCEPTANCE_MAX_AGE_MS) {
    return "Agreement confirmation expired. Please review and accept again.";
  }
  if (agreedAtMs - Date.now() > CLICKWRAP_CLOCK_SKEW_MS) {
    return "Agreement timestamp is in the future.";
  }
  if (Number.isFinite(agreement?.maxScrollProgress) && agreement.maxScrollProgress < 0.985) {
    return "Agreement scroll requirement was not completed.";
  }
  return "";
}

function getRequestIpAddress(headers) {
  if (!headers) return "";
  const candidates = [
    headers.get("cf-connecting-ip"),
    headers.get("x-forwarded-for"),
    headers.get("x-real-ip"),
    headers.get("x-client-ip"),
  ];
  for (const candidate of candidates) {
    const value = String(candidate || "")
      .split(",")[0]
      .trim();
    if (value) return value;
  }
  return "";
}

async function buildClickwrapEvidence(env, request, invite, contract, agreement, createdUserId, email) {
  const acceptedAt = agreement?.agreedAt || new Date().toISOString();
  const payload = {
    schema: "shipide.clickwrap.v1",
    accepted_at: acceptedAt,
    contract_version: String(contract?.version || ""),
    contract_hash_sha256: String(contract?.hash_sha256 || ""),
    contract_id: contract?.id || null,
    invite_id: invite?.id || null,
    accepted_email: normalizeEmail(email || ""),
    invited_email: normalizeEmail(invite?.invited_email || ""),
    user_id: createdUserId || null,
    scrolled_to_end_at: agreement?.scrolledToEndAt || null,
    max_scroll_progress: Number.isFinite(agreement?.maxScrollProgress)
      ? Number(agreement.maxScrollProgress.toFixed(4))
      : null,
    client_timezone: String(agreement?.clientTimezone || "").trim(),
    client_locale: String(agreement?.clientLocale || "").trim(),
    ip_address: getRequestIpAddress(request?.headers),
    user_agent: String(request?.headers?.get("user-agent") || "").trim(),
    captured_at: new Date().toISOString(),
  };
  const payloadCanonical = JSON.stringify(payload);
  const digest = await sha256Hex(payloadCanonical);
  const signature = await hmacSha256Hex(getStateSecret(env), payloadCanonical);
  return {
    payload,
    digest,
    signature,
  };
}

async function insertClickwrapAcceptance(env, record) {
  const response = await supabaseServiceRequest(env, `/rest/v1/${CLICKWRAP_ACCEPTANCES_TABLE}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify([record]),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (isMissingClickwrapSchemaError(details)) {
      throw new Error("Click-wrap schema missing. Run supabase_clickwrap.sql in Supabase SQL editor.");
    }
    throw new Error(`Could not store agreement proof (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
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

function toCents(value) {
  return Math.round((Number(value) || 0) * 100);
}

function fromCents(value) {
  return Number(((Number(value) || 0) / 100).toFixed(2));
}

function allocateCents(totalCents, count) {
  const safeCount = Math.max(1, Number(count) || 1);
  const safeTotal = Math.max(0, Number(totalCents) || 0);
  const base = Math.floor(safeTotal / safeCount);
  const remainder = safeTotal - base * safeCount;
  const out = new Array(safeCount).fill(base);
  for (let index = 0; index < remainder; index += 1) {
    out[index] += 1;
  }
  return out;
}

function toIsoDate(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseDateInput(value) {
  const raw = String(value || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return null;
  const date = new Date(`${raw}T00:00:00.000Z`);
  if (Number.isNaN(date.getTime())) return null;
  return date;
}

function startOfMonthUtc(date = new Date()) {
  const source = date instanceof Date ? date : new Date(date);
  return new Date(Date.UTC(source.getUTCFullYear(), source.getUTCMonth(), 1));
}

function addUtcDays(date, days) {
  const source = date instanceof Date ? date : new Date(date);
  return new Date(source.getTime() + Number(days || 0) * 24 * 60 * 60 * 1000);
}

function getCurrentMonthWindow(now = new Date()) {
  const start = startOfMonthUtc(now);
  const endExclusive = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth() + 1, 1));
  const end = addUtcDays(endExclusive, -1);
  return {
    start,
    end,
    endExclusive,
    startIsoDate: toIsoDate(start),
    endIsoDate: toIsoDate(end),
    startIsoDateTime: start.toISOString(),
    endExclusiveIsoDateTime: endExclusive.toISOString(),
  };
}

function getPreviousMonthWindow(now = new Date()) {
  const currentMonth = startOfMonthUtc(now);
  const start = new Date(Date.UTC(currentMonth.getUTCFullYear(), currentMonth.getUTCMonth() - 1, 1));
  const endExclusive = currentMonth;
  const end = addUtcDays(endExclusive, -1);
  return {
    start,
    end,
    endExclusive,
    startIsoDate: toIsoDate(start),
    endIsoDate: toIsoDate(end),
    startIsoDateTime: start.toISOString(),
    endExclusiveIsoDateTime: endExclusive.toISOString(),
  };
}

function getBillingWindow(rawStart, rawEnd, mode = "previous") {
  const startInput = parseDateInput(rawStart);
  const endInput = parseDateInput(rawEnd);
  if (startInput && endInput && startInput <= endInput) {
    const endExclusive = addUtcDays(endInput, 1);
    return {
      start: startInput,
      end: endInput,
      endExclusive,
      startIsoDate: toIsoDate(startInput),
      endIsoDate: toIsoDate(endInput),
      startIsoDateTime: startInput.toISOString(),
      endExclusiveIsoDateTime: endExclusive.toISOString(),
    };
  }
  return mode === "current" ? getCurrentMonthWindow() : getPreviousMonthWindow();
}

function getInvoiceTermsDays(env) {
  const value = Number(env.INVOICE_TERMS_DAYS);
  if (!Number.isFinite(value)) return DEFAULT_BILLING_TERMS_DAYS;
  return Math.max(1, Math.min(120, Math.round(value)));
}

function getReminderThresholdsDays() {
  return [-15, -7, -1, 0, 3, 10, 15, 30];
}

function getReminderStageForDueDate(dueAt, now = Date.now()) {
  const dueMs = Date.parse(String(dueAt || ""));
  if (!Number.isFinite(dueMs)) return 0;
  const nowMs = now instanceof Date ? now.getTime() : Number(now);
  if (!Number.isFinite(nowMs)) return 0;
  const thresholds = getReminderThresholdsDays();
  let stage = 0;
  thresholds.forEach((offset, index) => {
    const thresholdMs = dueMs + offset * 24 * 60 * 60 * 1000;
    if (nowMs >= thresholdMs) {
      stage = index + 1;
    }
  });
  return stage;
}

function isUuid(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    String(value || "").trim()
  );
}

function toInvoiceReference(invoiceId) {
  const safe = String(invoiceId || "").trim();
  if (!safe) return "INV-UNKNOWN";
  return `INV-${safe.slice(0, 8).toUpperCase()}`;
}

function formatInvoiceSubjectMonthYear(invoice = {}) {
  const candidates = [invoice?.period_start, invoice?.issued_at, invoice?.created_at, invoice?.due_at];
  const formatter = new Intl.DateTimeFormat("en-US", {
    month: "long",
    year: "numeric",
    timeZone: "UTC",
  });
  for (const candidate of candidates) {
    const raw = String(candidate || "").trim();
    if (!raw) continue;
    const dateOnly = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    let date = null;
    if (dateOnly) {
      date = new Date(Date.UTC(Number(dateOnly[1]), Number(dateOnly[2]) - 1, Number(dateOnly[3]), 12));
    } else {
      const parsed = Date.parse(raw);
      if (Number.isFinite(parsed)) date = new Date(parsed);
    }
    if (date && Number.isFinite(date.getTime())) {
      return formatter.format(date);
    }
  }
  return formatter.format(new Date());
}

function buildInvoiceEmailSubject(invoice = {}, options = {}) {
  const reference = toInvoiceReference(invoice?.id);
  const monthYear = formatInvoiceSubjectMonthYear(invoice);
  const suffix = `${monthYear} [${reference}]`;
  const isReminder = Boolean(options?.isReminder);
  const reminderStage = Number(options?.reminderStage) || 0;
  if (isReminder && reminderStage >= 5) {
    return `URGENT ! Invoice Overdue — ${suffix}`;
  }
  if (isReminder) {
    return `Payment Reminder — ${suffix}`;
  }
  return `Your Invoice — ${suffix}`;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatInvoiceTrackingLabel(invoice) {
  if (!invoice) return "No invoices yet";
  const status = String(invoice.status || "").trim().toLowerCase();
  if (status === "paid") return "Paid";
  if (status === "overdue") return "Overdue";
  if (status === "sent") {
    const stage = Number(invoice.reminder_stage) || 0;
    return stage > 0 ? `Reminder stage ${stage}` : "Sent";
  }
  if (status === "cancelled") return "Cancelled";
  return "Draft ready";
}

function getInvoiceRecipientFromSnapshot(invoice, fallbackEmail = "") {
  const recipient = String(invoice?.contact_email || "").trim().toLowerCase();
  if (recipient) return recipient;
  return String(fallbackEmail || "").trim().toLowerCase();
}

function assertResendConfig(env) {
  if (!env.RESEND_API_KEY) {
    throw new Error("RESEND_API_KEY is required.");
  }
}

async function sendResendEmail(env, payload) {
  assertResendConfig(env);
  const fromName = String(payload?.fromName || env.RESEND_FROM_NAME || "Shipide Billing").trim();
  const fromEmail = String(payload?.fromEmail || env.RESEND_FROM_EMAIL || "").trim();
  if (!fromEmail) {
    throw new Error("RESEND_FROM_EMAIL (or payload.fromEmail) is required.");
  }
  const requestBody = {
    from: `${fromName} <${fromEmail}>`,
    to: Array.isArray(payload?.to) ? payload.to : [String(payload?.to || "").trim()],
    subject: String(payload?.subject || "").trim(),
    html: String(payload?.html || ""),
    text: String(payload?.text || ""),
  };
  const replyTo = String(payload?.replyTo || env.RESEND_REPLY_TO || "").trim();
  if (replyTo) {
    requestBody.reply_to = replyTo;
  }
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${env.RESEND_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(requestBody),
  });
  const payloadJson = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message =
      payloadJson?.message ||
      payloadJson?.error ||
      `Resend request failed (${response.status}).`;
    throw new Error(String(message).trim());
  }
  return payloadJson;
}

function normalizeClientBillingPreference(value) {
  const raw = value && typeof value === "object" ? value : {};
  let invoiceEnabled = raw.invoice_enabled === true;
  let cardEnabled = raw.card_enabled === false ? false : true;
  if (!invoiceEnabled && !cardEnabled) {
    cardEnabled = true;
  }
  return {
    invoice_enabled: invoiceEnabled,
    card_enabled: cardEnabled,
    updated_at: raw.updated_at || null,
    updated_by: raw.updated_by || null,
  };
}

function getClientPaymentMode(pref) {
  const normalized = normalizeClientBillingPreference(pref);
  if (normalized.invoice_enabled && normalized.card_enabled) return "hybrid";
  if (normalized.card_enabled) return "card";
  return "invoice";
}

function buildAdminClientMetrics(user, historyRows, settings, billingPreference, invoiceStats) {
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

  const normalizedBilling = normalizeClientBillingPreference(billingPreference);
  const paymentMode = getClientPaymentMode(normalizedBilling);
  const normalizedInvoiceStats =
    invoiceStats && typeof invoiceStats === "object" ? invoiceStats : null;
  let avgPaymentDays = normalizedInvoiceStats?.avg_payment_days ?? null;
  let lastInvoiceTracking =
    normalizedInvoiceStats?.last_invoice_tracking ||
    (totalRevenue > 0 ? "Billing not live" : "No billable activity");
  if (paymentMode === "card") {
    avgPaymentDays = totalRevenue > 0 ? 0 : null;
    lastInvoiceTracking = totalRevenue > 0 ? "Card auto-charge" : "No billable activity";
  } else if (paymentMode === "hybrid") {
    lastInvoiceTracking =
      normalizedInvoiceStats?.last_invoice_tracking ||
      (totalRevenue > 0 ? "Hybrid billing" : "No billable activity");
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
    avg_payment_days: avgPaymentDays,
    last_invoice_tracking: lastInvoiceTracking,
    outstanding_balance: normalizedInvoiceStats?.outstanding_balance || 0,
    payment_mode: paymentMode,
    invoice_enabled: normalizedBilling.invoice_enabled,
    card_enabled: normalizedBilling.card_enabled,
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
  const totalOutstanding = clients.reduce(
    (sum, client) => sum + Math.max(0, Number(client?.metrics?.outstanding_balance) || 0),
    0
  );
  return {
    total_clients: clients.length,
    active_clients: activeClients,
    open_invites: openInvites,
    total_revenue_ex_vat: Number(totalRevenue.toFixed(2)),
    total_profit_ex_vat: Number(totalProfit.toFixed(2)),
    total_outstanding_balance: Number(totalOutstanding.toFixed(2)),
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

async function listUnbilledGenerationRows(env, { window, userId = "" } = {}) {
  const safeWindow = window || getPreviousMonthWindow();
  const safeUserId = String(userId || "").trim();
  const rows = [];
  const pageSize = 1000;
  for (let start = 0; start < 20000; start += pageSize) {
    const end = start + pageSize - 1;
    const params = new URLSearchParams();
    params.set("select", "id,user_id,created_at,service_type,quantity,total_price,payload,billed_invoice_id");
    params.set("order", "created_at.asc");
    params.set("created_at", `gte.${safeWindow.startIsoDateTime}`);
    params.append("created_at", `lt.${safeWindow.endExclusiveIsoDateTime}`);
    params.set("billed_invoice_id", "is.null");
    if (safeUserId) {
      params.set("user_id", `eq.${safeUserId}`);
    }
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
      if (/billed_invoice_id/i.test(details) || /billing_invoices/i.test(details)) {
        throw new Error(
          "Billing schema missing. Run supabase_invoicing.sql in Supabase SQL editor."
        );
      }
      throw new Error(`Could not load unbilled label history (${response.status}) ${details}`.trim());
    }
    const pageRows = await response.json().catch(() => []);
    if (!Array.isArray(pageRows) || !pageRows.length) break;
    rows.push(...pageRows);
    if (pageRows.length < pageSize) break;
  }
  return rows;
}

async function getClientBillingPreferenceForUser(env, userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return { ...DEFAULT_CLIENT_BILLING_PREF };
  const params = new URLSearchParams();
  params.set("select", "user_id,invoice_enabled,card_enabled,updated_at,updated_by");
  params.set("user_id", `eq.${safeUserId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${CLIENT_BILLING_PREF_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (/relation .*client_billing_preferences/i.test(details)) {
      return { ...DEFAULT_CLIENT_BILLING_PREF };
    }
    throw new Error(`Could not load client billing settings (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) return { ...DEFAULT_CLIENT_BILLING_PREF };
  return normalizeClientBillingPreference(rows[0]);
}

function getBillingIbanConfig(env) {
  return {
    beneficiary: String(env.BILLING_IBAN_BENEFICIARY || DEFAULT_IBAN_BENEFICIARY).trim(),
    iban: String(env.BILLING_IBAN || DEFAULT_IBAN).trim(),
    bic: String(env.BILLING_IBAN_BIC || DEFAULT_IBAN_BIC).trim(),
    note: String(env.BILLING_IBAN_NOTE || DEFAULT_IBAN_TRANSFER_NOTE).trim(),
  };
}

function buildTopupReference(userId = "") {
  const userChunk = String(userId || "")
    .replace(/[^a-z0-9]/gi, "")
    .slice(0, 6)
    .toUpperCase() || "SHIPID";
  const timeChunk = Date.now().toString(36).toUpperCase();
  const randChunk = Math.floor(Math.random() * 0xffffff)
    .toString(36)
    .padStart(4, "0")
    .slice(0, 4)
    .toUpperCase();
  return `SHIP-${userChunk}-${timeChunk}-${randChunk}`;
}

function resolveWalletAccessIssue(details, tableName) {
  const raw = String(details || "");
  const lower = raw.toLowerCase();
  const table = String(tableName || "").toLowerCase();
  if (!lower || !table || !lower.includes(table)) return null;

  if (lower.includes("schema cache")) {
    return {
      kind: "cache",
      message: "Wallet API schema cache is stale. Run: NOTIFY pgrst, 'reload schema'; in Supabase SQL editor.",
    };
  }
  if (
    lower.includes("does not exist") ||
    lower.includes("undefined table") ||
    lower.includes("\"code\":\"42p01\"")
  ) {
    return {
      kind: "missing",
      message: "Wallet schema missing. Run supabase_wallet_billing.sql in Supabase SQL editor.",
    };
  }
  if (
    lower.includes("permission denied") ||
    lower.includes("insufficient_privilege") ||
    lower.includes("\"code\":\"42501\"")
  ) {
    return {
      kind: "permission",
      message: `Wallet access denied for ${tableName}. Grant table privileges (service_role/authenticated), then run: NOTIFY pgrst, 'reload schema';`,
    };
  }
  return null;
}

function normalizeTopupStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["pending", "received", "credited", "cancelled", "failed"].includes(normalized)) {
    return normalized;
  }
  return "pending";
}

function mapBillingTopupRow(row) {
  const status = normalizeTopupStatus(row?.status);
  const amountCents = Math.max(0, toCents(fromCents(row?.amount_cents)));
  return {
    id: String(row?.id || "").trim(),
    amount_eur: fromCents(amountCents),
    currency: String(row?.currency || DEFAULT_BILLING_CURRENCY).trim() || DEFAULT_BILLING_CURRENCY,
    reference: String(row?.reference || "").trim(),
    status,
    requested_at: row?.requested_at || row?.created_at || null,
    received_at: row?.received_at || null,
    credited_at: row?.credited_at || null,
    metadata: row?.metadata && typeof row.metadata === "object" ? row.metadata : {},
  };
}

function mapBillingWalletTransactionRow(row) {
  const amountCents = Number(row?.amount_cents) || 0;
  return {
    id: row?.id ?? null,
    source: String(row?.source || "").trim() || "manual",
    direction: amountCents >= 0 ? "credit" : "debit",
    amount_eur: fromCents(Math.abs(amountCents)),
    amount_signed_eur: fromCents(amountCents),
    balance_after_eur: fromCents(Number(row?.balance_after_cents) || 0),
    reference: String(row?.reference || "").trim(),
    created_at: row?.created_at || null,
    metadata: row?.metadata && typeof row.metadata === "object" ? row.metadata : {},
  };
}

async function getOrCreateBillingWallet(env, userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) {
    return {
      user_id: "",
      balance_cents: 0,
      currency: DEFAULT_BILLING_CURRENCY,
      updated_at: null,
    };
  }

  const params = new URLSearchParams();
  params.set("select", "user_id,balance_cents,currency,updated_at");
  params.set("user_id", `eq.${safeUserId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_WALLET_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const walletIssue = resolveWalletAccessIssue(details, BILLING_WALLET_TABLE);
    if (walletIssue) throw new Error(walletIssue.message);
    throw new Error(`Could not load wallet balance (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (Array.isArray(rows) && rows.length) {
    return rows[0];
  }

  const insertResponse = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_WALLET_TABLE}`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=representation,resolution=merge-duplicates",
      },
      body: JSON.stringify([
        {
          user_id: safeUserId,
          balance_cents: 0,
          currency: DEFAULT_BILLING_CURRENCY,
        },
      ]),
    }
  );
  if (!insertResponse.ok) {
    const details = await insertResponse.text().catch(() => "");
    const walletIssue = resolveWalletAccessIssue(details, BILLING_WALLET_TABLE);
    if (walletIssue) throw new Error(walletIssue.message);
    throw new Error(`Could not create wallet (${insertResponse.status}) ${details}`.trim());
  }
  const insertedRows = await insertResponse.json().catch(() => []);
  return Array.isArray(insertedRows) && insertedRows.length
    ? insertedRows[0]
    : {
        user_id: safeUserId,
        balance_cents: 0,
        currency: DEFAULT_BILLING_CURRENCY,
        updated_at: new Date().toISOString(),
      };
}

async function listBillingTopups(
  env,
  { userId = "", statuses = [], limit = 30, allowMissing = false } = {}
) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return [];
  const safeLimit = Math.max(1, Math.min(200, Number(limit) || 30));
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,user_id,amount_cents,currency,status,reference,requested_at,received_at,credited_at,metadata,created_at"
  );
  params.set("user_id", `eq.${safeUserId}`);
  params.set("order", "requested_at.desc");
  params.set("limit", String(safeLimit));
  if (Array.isArray(statuses) && statuses.length) {
    const normalized = statuses
      .map((entry) => normalizeTopupStatus(entry))
      .filter(Boolean);
    if (normalized.length) {
      params.set("status", `in.(${normalized.join(",")})`);
    }
  }
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_WALLET_TOPUPS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const walletIssue = resolveWalletAccessIssue(details, BILLING_WALLET_TOPUPS_TABLE);
    if (walletIssue) {
      if (allowMissing && walletIssue.kind === "missing") return [];
      throw new Error(walletIssue.message);
    }
    throw new Error(`Could not load bank top-ups (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function listBillingWalletTransactions(
  env,
  { userId = "", limit = 30, allowMissing = false } = {}
) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return [];
  const safeLimit = Math.max(1, Math.min(200, Number(limit) || 30));
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,user_id,source,amount_cents,balance_after_cents,reference,metadata,created_at"
  );
  params.set("user_id", `eq.${safeUserId}`);
  params.set("order", "created_at.desc,id.desc");
  params.set("limit", String(safeLimit));
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_WALLET_TRANSACTIONS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const walletIssue = resolveWalletAccessIssue(details, BILLING_WALLET_TRANSACTIONS_TABLE);
    if (walletIssue) {
      if (allowMissing && walletIssue.kind === "missing") return [];
      throw new Error(walletIssue.message);
    }
    throw new Error(`Could not load wallet transactions (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function createBillingTopupRequest(env, user, amountEur = null) {
  const userId = String(user?.id || "").trim();
  if (!userId) {
    throw new Error("Authentication required.");
  }
  const numericAmount = Number(amountEur);
  const amountCents =
    Number.isFinite(numericAmount) && numericAmount > 0 ? Math.max(0, toCents(numericAmount)) : 0;
  await getOrCreateBillingWallet(env, userId);
  const ibanConfig = getBillingIbanConfig(env);
  const reference = buildTopupReference(userId);
  const payload = {
    user_id: userId,
    amount_cents: amountCents,
    currency: DEFAULT_BILLING_CURRENCY,
    status: "pending",
    reference,
    requested_at: new Date().toISOString(),
    metadata: {
      requested_by: normalizeEmail(user?.email || "") || userId,
      iban_beneficiary: ibanConfig.beneficiary,
      iban: ibanConfig.iban,
      bic: ibanConfig.bic,
    },
  };
  const response = await supabaseServiceRequest(env, `/rest/v1/${BILLING_WALLET_TOPUPS_TABLE}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify([payload]),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const walletIssue = resolveWalletAccessIssue(details, BILLING_WALLET_TOPUPS_TABLE);
    if (walletIssue) throw new Error(walletIssue.message);
    throw new Error(`Could not create top-up request (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  const created = Array.isArray(rows) && rows.length ? rows[0] : payload;
  return {
    topup: mapBillingTopupRow(created),
    instructions: {
      ...ibanConfig,
      amount_eur: amountCents > 0 ? fromCents(amountCents) : null,
      currency: DEFAULT_BILLING_CURRENCY,
      reference,
    },
  };
}

async function saveWalletTransaction(env, payload) {
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_WALLET_TRANSACTIONS_TABLE}`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=representation",
      },
      body: JSON.stringify([payload]),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const walletIssue = resolveWalletAccessIssue(details, BILLING_WALLET_TRANSACTIONS_TABLE);
    if (walletIssue) throw new Error(walletIssue.message);
    throw new Error(`Could not save wallet transaction (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : payload;
}

async function debitWalletForCheckout(env, user, amountEur, metadata = {}) {
  const userId = String(user?.id || "").trim();
  if (!userId) {
    throw new Error("Authentication required.");
  }
  const debitCents = toCents(amountEur);
  if (!Number.isFinite(debitCents) || debitCents <= 0) {
    throw new Error("Invalid checkout amount.");
  }
  const wallet = await getOrCreateBillingWallet(env, userId);
  const currentBalanceCents = Number(wallet?.balance_cents) || 0;
  if (currentBalanceCents < debitCents) {
    const missing = fromCents(debitCents - currentBalanceCents).toFixed(2);
    throw new Error(`Insufficient wallet balance. Add at least €${missing} by bank transfer.`);
  }
  const nextBalanceCents = Math.max(0, currentBalanceCents - debitCents);
  const updateResponse = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_WALLET_TABLE}?user_id=eq.${encodeURIComponent(userId)}`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=representation",
      },
      body: JSON.stringify({
        balance_cents: nextBalanceCents,
        updated_at: new Date().toISOString(),
      }),
    }
  );
  if (!updateResponse.ok) {
    const details = await updateResponse.text().catch(() => "");
    throw new Error(`Could not update wallet balance (${updateResponse.status}) ${details}`.trim());
  }
  const updatedRows = await updateResponse.json().catch(() => []);
  const updatedWallet =
    Array.isArray(updatedRows) && updatedRows.length
      ? updatedRows[0]
      : {
          user_id: userId,
          balance_cents: nextBalanceCents,
          currency: wallet?.currency || DEFAULT_BILLING_CURRENCY,
        };
  const reference = buildTopupReference(userId);
  await saveWalletTransaction(env, {
    user_id: userId,
    source: "label_checkout",
    amount_cents: -debitCents,
    balance_after_cents: nextBalanceCents,
    reference,
    metadata: {
      checkout: metadata && typeof metadata === "object" ? metadata : {},
      actor: normalizeEmail(user?.email || "") || userId,
    },
  });
  return {
    wallet: updatedWallet,
    transaction_reference: reference,
  };
}

function mapBillingOverviewInvoiceRow(invoice) {
  if (!invoice || typeof invoice !== "object") return null;
  return {
    id: String(invoice.id || "").trim(),
    status: normalizeInvoiceStatus(invoice.status),
    reference: toInvoiceReference(invoice.id),
    issued_at: invoice.issued_at || invoice.created_at || null,
    due_at: invoice.due_at || null,
    total_inc_vat: fromCents(toCents(invoice.total_inc_vat)),
    total_ex_vat: fromCents(toCents(invoice.subtotal_ex_vat)),
    tracking: formatInvoiceTrackingLabel(invoice),
  };
}

async function listBillingInvoices(env, { limit = 100, userId = "", statuses = [], allowMissing = false } = {}) {
  const safeUserId = String(userId || "").trim();
  const safeLimit = Math.max(1, Math.min(1000, Number(limit) || 100));
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,user_id,period_start,period_end,due_at,issued_at,sent_at,paid_at,status,payment_mode,currency,vat_rate,subtotal_ex_vat,vat_amount,total_inc_vat,labels_count,line_count,reminder_stage,last_reminder_sent_at,email_message_id,company_name,contact_name,contact_email,customer_id,account_manager,payment_reference,payment_received_amount,metadata,created_at,updated_at"
  );
  params.set("order", "updated_at.desc");
  params.set("limit", String(safeLimit));
  if (safeUserId) {
    params.set("user_id", `eq.${safeUserId}`);
  }
  if (Array.isArray(statuses) && statuses.length) {
    const normalized = statuses
      .map((entry) => String(entry || "").trim().toLowerCase())
      .filter(Boolean);
    if (normalized.length) {
      params.set("status", `in.(${normalized.join(",")})`);
    }
  }
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_INVOICES_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (/relation .*billing_invoices/i.test(details)) {
      if (allowMissing) return [];
      throw new Error("Billing schema missing. Run supabase_invoicing.sql in Supabase SQL editor.");
    }
    throw new Error(`Could not load invoices (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function getBillingInvoiceByPeriod(env, userId, periodStart, periodEnd) {
  const safeUserId = String(userId || "").trim();
  const safeStart = String(periodStart || "").trim();
  const safeEnd = String(periodEnd || "").trim();
  if (!safeUserId || !safeStart || !safeEnd) return null;
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,user_id,period_start,period_end,due_at,issued_at,sent_at,paid_at,status,payment_mode,currency,vat_rate,subtotal_ex_vat,vat_amount,total_inc_vat,labels_count,line_count,reminder_stage,last_reminder_sent_at,email_message_id,company_name,contact_name,contact_email,customer_id,account_manager,payment_reference,payment_received_amount,metadata,created_at,updated_at"
  );
  params.set("user_id", `eq.${safeUserId}`);
  params.set("period_start", `eq.${safeStart}`);
  params.set("period_end", `eq.${safeEnd}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_INVOICES_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (/relation .*billing_invoices/i.test(details)) {
      throw new Error("Billing schema missing. Run supabase_invoicing.sql in Supabase SQL editor.");
    }
    throw new Error(`Could not load invoice period (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function getBillingInvoiceById(env, invoiceId, { withItems = false } = {}) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId) return null;
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,user_id,period_start,period_end,due_at,issued_at,sent_at,paid_at,status,payment_mode,currency,vat_rate,subtotal_ex_vat,vat_amount,total_inc_vat,labels_count,line_count,reminder_stage,last_reminder_sent_at,email_message_id,company_name,contact_name,contact_email,customer_id,account_manager,payment_reference,payment_received_amount,metadata,created_at,updated_at"
  );
  params.set("id", `eq.${safeInvoiceId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_INVOICES_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (/relation .*billing_invoices/i.test(details)) {
      throw new Error("Billing schema missing. Run supabase_invoicing.sql in Supabase SQL editor.");
    }
    throw new Error(`Could not load invoice (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  const invoice = Array.isArray(rows) && rows.length ? rows[0] : null;
  if (!invoice || !withItems) return invoice;

  const itemParams = new URLSearchParams();
  itemParams.set(
    "select",
    "id,invoice_id,generation_id,generated_at,service_type,quantity,recipient_name,recipient_city,recipient_country,tracking_id,label_id,unit_price_ex_vat,amount_ex_vat,vat_amount,amount_inc_vat,sort_index,metadata,created_at"
  );
  itemParams.set("invoice_id", `eq.${safeInvoiceId}`);
  itemParams.set("order", "sort_index.asc,id.asc");
  itemParams.set("limit", "10000");
  const itemsResponse = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_INVOICE_ITEMS_TABLE}?${itemParams.toString()}`,
    { method: "GET" }
  );
  if (!itemsResponse.ok) {
    const details = await itemsResponse.text().catch(() => "");
    throw new Error(`Could not load invoice items (${itemsResponse.status}) ${details}`.trim());
  }
  const items = await itemsResponse.json().catch(() => []);
  return {
    ...invoice,
    items: Array.isArray(items) ? items : [],
  };
}

async function upsertBillingInvoice(env, invoicePayload) {
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_INVOICES_TABLE}?on_conflict=user_id,period_start,period_end`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=representation",
      },
      body: JSON.stringify([invoicePayload]),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (/relation .*billing_invoices/i.test(details)) {
      throw new Error("Billing schema missing. Run supabase_invoicing.sql in Supabase SQL editor.");
    }
    throw new Error(`Could not save invoice (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function replaceBillingInvoiceItems(env, invoiceId, items = []) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId) return;
  const deleteParams = new URLSearchParams();
  deleteParams.set("invoice_id", `eq.${safeInvoiceId}`);
  const deleteResponse = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_INVOICE_ITEMS_TABLE}?${deleteParams.toString()}`,
    {
      method: "DELETE",
      headers: {
        Prefer: "return=minimal",
      },
    }
  );
  if (!deleteResponse.ok) {
    const details = await deleteResponse.text().catch(() => "");
    throw new Error(`Could not clear invoice items (${deleteResponse.status}) ${details}`.trim());
  }
  if (!Array.isArray(items) || !items.length) return;

  const response = await supabaseServiceRequest(env, `/rest/v1/${BILLING_INVOICE_ITEMS_TABLE}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=minimal",
    },
    body: JSON.stringify(
      items.map((item, index) => ({
        invoice_id: safeInvoiceId,
        generation_id: isUuid(item?.generation_id) ? item.generation_id : null,
        generated_at: item?.generated_at || null,
        service_type: String(item?.service_type || "").trim() || null,
        quantity: Math.max(1, Number(item?.quantity) || 1),
        recipient_name: String(item?.recipient_name || "").trim() || null,
        recipient_city: String(item?.recipient_city || "").trim() || null,
        recipient_country: String(item?.recipient_country || "").trim() || null,
        tracking_id: String(item?.tracking_id || "").trim() || null,
        label_id: String(item?.label_id || "").trim() || null,
        unit_price_ex_vat: fromCents(toCents(item?.unit_price_ex_vat)),
        amount_ex_vat: fromCents(toCents(item?.amount_ex_vat)),
        vat_amount: fromCents(toCents(item?.vat_amount)),
        amount_inc_vat: fromCents(toCents(item?.amount_inc_vat)),
        sort_index: Number.isFinite(Number(item?.sort_index)) ? Number(item.sort_index) : index,
        metadata:
          item?.metadata && typeof item.metadata === "object" ? item.metadata : {},
      }))
    ),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not save invoice items (${response.status}) ${details}`.trim());
  }
}

async function markGenerationRowsBilled(env, generationIds = [], invoiceId) {
  const validIds = generationIds.filter((value) => isUuid(value));
  if (!validIds.length || !isUuid(invoiceId)) return;
  const billedAt = new Date().toISOString();
  const chunkSize = 80;
  for (let index = 0; index < validIds.length; index += chunkSize) {
    const chunk = validIds.slice(index, index + chunkSize);
    const params = new URLSearchParams();
    params.set("id", `in.(${chunk.join(",")})`);
    params.set("billed_invoice_id", "is.null");
    const response = await supabaseServiceRequest(env, `/rest/v1/${HISTORY_TABLE}?${params.toString()}`, {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=minimal",
      },
      body: JSON.stringify({
        billed_invoice_id: invoiceId,
        billed_at: billedAt,
      }),
    });
    if (!response.ok) {
      const details = await response.text().catch(() => "");
      throw new Error(`Could not mark billed rows (${response.status}) ${details}`.trim());
    }
  }
}

async function insertBillingInvoiceEvent(env, invoiceId, event) {
  if (!isUuid(invoiceId)) return;
  const payload = {
    invoice_id: invoiceId,
    event_type: String(event?.event_type || "updated").trim().toLowerCase(),
    event_at: event?.event_at || new Date().toISOString(),
    actor: String(event?.actor || "system").trim(),
    channel: String(event?.channel || "").trim() || null,
    message: String(event?.message || "").trim() || null,
    metadata: event?.metadata && typeof event.metadata === "object" ? event.metadata : {},
  };
  const response = await supabaseServiceRequest(env, `/rest/v1/${BILLING_INVOICE_EVENTS_TABLE}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=minimal",
    },
    body: JSON.stringify([payload]),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (/relation .*billing_invoice_events/i.test(details)) {
      throw new Error("Billing schema missing. Run supabase_invoicing.sql in Supabase SQL editor.");
    }
    throw new Error(`Could not save invoice event (${response.status}) ${details}`.trim());
  }
}

function getRecipientFromLabelPayload(payload = {}) {
  const data = payload && typeof payload === "object" ? payload : {};
  const recipientName = String(data?.recipientName || "").trim();
  const recipientCity = String(data?.recipientCity || "").trim();
  const recipientCountry = String(data?.recipientCountry || "").trim();
  return {
    recipient_name: recipientName || null,
    recipient_city: recipientCity || null,
    recipient_country: recipientCountry || null,
  };
}

function buildInvoiceItemsFromHistoryRows(rows = [], vatRate = DEFAULT_VAT_RATE) {
  const safeVatRate = Math.max(0, Number(vatRate) || DEFAULT_VAT_RATE);
  const items = [];
  const sortedRows = Array.isArray(rows)
    ? rows.slice().sort((left, right) => Date.parse(left?.created_at || 0) - Date.parse(right?.created_at || 0))
    : [];

  sortedRows.forEach((row) => {
    const payload = row?.payload && typeof row.payload === "object" ? row.payload : {};
    const labels = Array.isArray(payload?.labels) ? payload.labels : [];
    const totalExCents = Math.max(0, toCents(row?.total_price));
    const quantityFromRow = Math.max(1, Number(row?.quantity) || labels.length || 1);

    if (labels.length) {
      const perLabelEx = allocateCents(totalExCents, labels.length);
      labels.forEach((label, labelIndex) => {
        const labelData = label?.data && typeof label.data === "object" ? label.data : {};
        const recipient = getRecipientFromLabelPayload(labelData);
        const amountEx = perLabelEx[labelIndex] || 0;
        const vat = Math.round(amountEx * safeVatRate);
        items.push({
          generation_id: row?.id || null,
          generated_at: row?.created_at || null,
          service_type: row?.service_type || null,
          quantity: 1,
          recipient_name: recipient.recipient_name,
          recipient_city: recipient.recipient_city,
          recipient_country: recipient.recipient_country,
          tracking_id: String(label?.trackingId || "").trim() || null,
          label_id: String(label?.labelId || "").trim() || null,
          unit_price_ex_vat: fromCents(amountEx),
          amount_ex_vat: fromCents(amountEx),
          vat_amount: fromCents(vat),
          amount_inc_vat: fromCents(amountEx + vat),
          sort_index: items.length,
          metadata: {
            source: "label_generation",
            generation_quantity: quantityFromRow,
            label_index: Number(label?.index) || labelIndex + 1,
          },
        });
      });
      return;
    }

    const unitExCents = Math.round(totalExCents / quantityFromRow);
    const vatCents = Math.round(totalExCents * safeVatRate);
    const recipient = getRecipientFromLabelPayload(payload?.selection || payload || {});
    items.push({
      generation_id: row?.id || null,
      generated_at: row?.created_at || null,
      service_type: row?.service_type || null,
      quantity: quantityFromRow,
      recipient_name: recipient.recipient_name,
      recipient_city: recipient.recipient_city,
      recipient_country: recipient.recipient_country,
      tracking_id: null,
      label_id: null,
      unit_price_ex_vat: fromCents(unitExCents),
      amount_ex_vat: fromCents(totalExCents),
      vat_amount: fromCents(vatCents),
      amount_inc_vat: fromCents(totalExCents + vatCents),
      sort_index: items.length,
      metadata: {
        source: "label_generation",
        generation_quantity: quantityFromRow,
      },
    });
  });

  const subtotalCents = items.reduce((sum, item) => sum + toCents(item.amount_ex_vat), 0);
  const vatCents = items.reduce((sum, item) => sum + toCents(item.vat_amount), 0);
  const totalCents = items.reduce((sum, item) => sum + toCents(item.amount_inc_vat), 0);
  const labelsCount = items.reduce((sum, item) => sum + Math.max(1, Number(item.quantity) || 1), 0);
  return {
    items,
    totals: {
      subtotal_ex_vat: fromCents(subtotalCents),
      vat_amount: fromCents(vatCents),
      total_inc_vat: fromCents(totalCents),
      labels_count: labelsCount,
      line_count: items.length,
      vat_rate: safeVatRate,
    },
  };
}

function buildInvoiceEmailHtml(invoice, items = [], options = {}) {
  const isReminder = Boolean(options?.isReminder);
  const reminderStage = Math.max(0, Number(options?.reminderStage ?? invoice?.reminder_stage) || 0);
  const viewUrl = String(options?.viewUrl || "#").trim() || "#";
  const isOverdueReminder = isReminder && reminderStage >= 5;
  const logoUrl = "https://portal.shipide.com/shipide_logo.png";
  const titleLine1 = isOverdueReminder
    ? "Urgent Invoice"
    : isReminder
      ? "Payment Reminder"
      : "Your Monthly Invoice";
  const titleLine2 = isOverdueReminder ? "Is Overdue" : isReminder ? "Invoice Due Soon" : "Is Ready";
  const subtitleLine1 = "Your invoice PDF is attached to this email.";
  const subtitleLine2 = "You can also view it instantly with the button below.";

  let stageLabel = "Due in 30 days";
  let stageBg = "rgba(143, 226, 178, 0.16)";
  let stageBorder = "#8fe2b2";
  let stageText = "#8fe2b2";
  if (reminderStage === 1) {
    stageLabel = "Due in 15 days";
    stageBg = "#2f2515";
    stageBorder = "#b57d23";
    stageText = "#ffd48c";
  } else if (reminderStage === 2) {
    stageLabel = "Due in 7 days";
    stageBg = "#322614";
    stageBorder = "#bf8122";
    stageText = "#ffd194";
  } else if (reminderStage === 3) {
    stageLabel = "Due tomorrow";
    stageBg = "#342413";
    stageBorder = "#c9821f";
    stageText = "#ffc889";
  } else if (reminderStage === 4) {
    stageLabel = "Due today";
    stageBg = "#3a171d";
    stageBorder = "#c24e62";
    stageText = "#ffbcc8";
  } else if (reminderStage >= 5) {
    stageLabel = "Overdue";
    stageBg = "#351a1e";
    stageBorder = "#b84b5e";
    stageText = "#ffb8c4";
  }
  return `
    <style>
      @media only screen and (max-width: 640px) {
        .shipide-email-hero-title {
          font-size: 42px !important;
          line-height: 1.04 !important;
          letter-spacing: -0.03em !important;
        }
        .shipide-email-hero-sub {
          font-size: 14px !important;
          line-height: 1.45 !important;
          max-width: 340px !important;
        }
        .shipide-email-hero-wrap {
          padding: 38px 8px 20px !important;
        }
        .shipide-email-hero-spacer {
          height: 56px !important;
          line-height: 56px !important;
        }
        .shipide-email-logo {
          width: 96px !important;
          margin-bottom: 24px !important;
        }
      }
    </style>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0;padding:0;background:#00060f;font-family:Helvetica,Arial,sans-serif;color:#f3f6ff;border-top:4px solid #7747e3;">
      <tr>
        <td align="center" style="padding:0 16px 24px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:980px;background:#00060f;">
            <tr>
              <td align="center" class="shipide-email-hero-wrap" style="padding:70px 8px 24px;">
                <img src="${escapeHtml(
                  logoUrl
                )}" alt="Shipide" class="shipide-email-logo" width="124" style="display:block;width:124px;height:auto;margin:0 auto 30px;" />
                <div class="shipide-email-hero-title" style="font-size:62px;line-height:1;letter-spacing:-0.035em;color:#f3f6ff;font-weight:300;">
                  ${escapeHtml(titleLine1)}<br/>${escapeHtml(titleLine2)}
                </div>
                <div class="shipide-email-hero-sub" style="max-width:700px;margin:18px auto 0;font-size:16px;line-height:1.5;color:#9aa3b2;">
                  ${escapeHtml(subtitleLine1)}<br/>${escapeHtml(subtitleLine2)}
                </div>
                <a href="${escapeHtml(viewUrl)}" style="display:inline-block;margin-top:24px;padding:12px 20px;border-radius:4px;border:1px solid rgb(46,46,46);background:#1c2026;color:#f3f6ff;text-decoration:none;font-size:14px;line-height:1;">
                  View Invoice
                </a>
              </td>
            </tr>
            <tr>
              <td class="shipide-email-hero-spacer" style="height:72px;font-size:0;line-height:72px;">&nbsp;</td>
            </tr>
            <tr>
              <td align="center" style="padding:0 8px 8px;">
                <span style="display:inline-block;padding:6px 12px;border-radius:999px;border:1px solid ${stageBorder};background:${stageBg};color:${stageText};font-size:11px;letter-spacing:0.01em;white-space:nowrap;">
                  ${escapeHtml(stageLabel)}
                </span>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  `;
}

function buildInvoiceEmailText(invoice, options = {}) {
  const reference = toInvoiceReference(invoice?.id);
  const dueLabel = String(invoice?.due_at || "").trim()
    ? new Date(invoice.due_at).toISOString().slice(0, 10)
    : "--";
  const prefix = options?.isReminder ? "Payment reminder" : "Invoice";
  return [
    `${prefix}: ${reference}`,
    `Client: ${invoice?.company_name || "Client account"}`,
    `Due date: ${dueLabel}`,
    `Subtotal (EX. VAT): €${fromCents(toCents(invoice?.subtotal_ex_vat)).toFixed(2)}`,
    `VAT: €${fromCents(toCents(invoice?.vat_amount)).toFixed(2)}`,
    `Total (INCL. VAT): €${fromCents(toCents(invoice?.total_inc_vat)).toFixed(2)}`,
  ].join("\n");
}

function formatEmailEuroAmount(value) {
  const numeric = Number(value);
  const safe = Number.isFinite(numeric) ? numeric : 0;
  return `€${safe.toFixed(2)}`;
}

function buildMonthlyReportsEmailSubject(options = {}) {
  const monthYear = formatInvoiceSubjectMonthYear({
    period_start: options?.monthDate || new Date().toISOString().slice(0, 10),
  });
  return `Monthly Reports — ${monthYear}`;
}

function getReportsPortalUrl(request, env) {
  const configured = String(env.REPORTS_PORTAL_URL || "").trim();
  if (configured) return configured;
  return `${getPublicOrigin(request)}/reports?range=monthly`;
}

function buildMonthlyReportsEmailHtml(options = {}) {
  const profitLabel = formatEmailEuroAmount(options?.profitAmount);
  const reportsUrl =
    String(options?.reportsUrl || "").trim() || "https://portal.shipide.com/reports?range=monthly";
  return `
    <style>
      @media only screen and (max-width: 640px) {
        .shipide-email-hero-title {
          font-size: 40px !important;
          line-height: 1.06 !important;
          letter-spacing: -0.03em !important;
        }
        .shipide-email-hero-sub {
          font-size: 14px !important;
          line-height: 1.45 !important;
          max-width: 340px !important;
        }
        .shipide-email-hero-wrap {
          padding: 38px 8px 20px !important;
        }
        .shipide-email-hero-spacer {
          height: 56px !important;
          line-height: 56px !important;
        }
        .shipide-email-logo {
          width: 96px !important;
          margin-bottom: 24px !important;
        }
      }
    </style>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0;padding:0;background:#00060f;font-family:Helvetica,Arial,sans-serif;color:#f3f6ff;border-top:4px solid #7747e3;">
      <tr>
        <td align="center" style="padding:0 16px 24px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:980px;background:#00060f;">
            <tr>
              <td align="center" class="shipide-email-hero-wrap" style="padding:70px 8px 24px;">
                <img src="https://portal.shipide.com/shipide_logo.png" alt="Shipide" class="shipide-email-logo" width="124" style="display:block;width:124px;height:auto;margin:0 auto 30px;" />
                <div class="shipide-email-hero-title" style="font-size:58px;line-height:1.03;letter-spacing:-0.035em;color:#f3f6ff;font-weight:300;">
                  This month, you profited<br/>an extra <span style="color:#8fe2b2;">${escapeHtml(
                    profitLabel
                  )}</span> with us.
                </div>
                <div class="shipide-email-hero-sub" style="max-width:700px;margin:18px auto 0;font-size:16px;line-height:1.5;color:#9aa3b2;">
                  Your monthly performance report is ready.<br/>Open it instantly with the button below.
                </div>
                <a href="${escapeHtml(reportsUrl)}" style="display:inline-block;margin-top:24px;padding:12px 20px;border-radius:4px;border:1px solid rgb(46,46,46);background:#1c2026;color:#f3f6ff;text-decoration:none;font-size:14px;line-height:1;">
                  View Report
                </a>
              </td>
            </tr>
            <tr>
              <td class="shipide-email-hero-spacer" style="height:72px;font-size:0;line-height:72px;">&nbsp;</td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  `;
}

function buildMonthlyReportsEmailText(options = {}) {
  const profitLabel = formatEmailEuroAmount(options?.profitAmount);
  const reportsUrl =
    String(options?.reportsUrl || "").trim() || "https://portal.shipide.com/reports?range=monthly";
  return [
    `This month, you profited an extra ${profitLabel} with us.`,
    "Your monthly performance report is ready.",
    "Open it instantly with the button below.",
    `View report: ${reportsUrl}`,
  ].join("\n");
}

async function updateBillingInvoiceFields(env, invoiceId, patch = {}) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId) throw new Error("Invoice id is required.");
  const params = new URLSearchParams();
  params.set("id", `eq.${safeInvoiceId}`);
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${BILLING_INVOICES_TABLE}?${params.toString()}`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=representation",
      },
      body: JSON.stringify({
        ...patch,
        updated_at: new Date().toISOString(),
      }),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not update invoice (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function getSupabaseUserById(env, userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return null;
  const response = await fetch(
    `${String(env.SUPABASE_URL).replace(/\/+$/, "")}/auth/v1/admin/users/${safeUserId}`,
    {
      method: "GET",
      headers: {
        apikey: env.SUPABASE_SERVICE_ROLE_KEY,
        Authorization: `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}`,
      },
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not load user (${response.status}) ${details}`.trim());
  }
  const payload = await response.json().catch(() => null);
  if (!payload || typeof payload !== "object") return null;
  return payload?.user && typeof payload.user === "object" ? payload.user : payload;
}

function mapInvoiceProfileFromUser(user = {}) {
  const metadata =
    user?.user_metadata && typeof user.user_metadata === "object" ? user.user_metadata : {};
  return {
    company_name: String(metadata.company_name || metadata.companyName || "").trim() || "Client account",
    contact_name: String(metadata.contact_name || metadata.contactName || "").trim() || "",
    contact_email:
      normalizeEmail(metadata.contact_email || metadata.contactEmail || user?.email || "") || "",
    customer_id: String(metadata.customer_id || metadata.customerId || "").trim() || "",
    account_manager: String(metadata.account_manager || metadata.accountManager || "").trim() || "",
  };
}

function buildInvoiceStatsByUser(invoices = []) {
  const grouped = new Map();
  invoices.forEach((invoice) => {
    const userId = String(invoice?.user_id || "").trim();
    if (!userId) return;
    if (!grouped.has(userId)) grouped.set(userId, []);
    grouped.get(userId).push(invoice);
  });

  const out = new Map();
  grouped.forEach((rows, userId) => {
    const sorted = rows
      .slice()
      .sort(
        (left, right) =>
          Date.parse(right?.updated_at || right?.sent_at || right?.created_at || 0) -
          Date.parse(left?.updated_at || left?.sent_at || left?.created_at || 0)
      );
    const paidDurations = sorted
      .filter((invoice) => invoice?.paid_at && (invoice?.issued_at || invoice?.sent_at))
      .map((invoice) => {
        const startMs = Date.parse(invoice?.issued_at || invoice?.sent_at || 0);
        const endMs = Date.parse(invoice?.paid_at || 0);
        if (!Number.isFinite(startMs) || !Number.isFinite(endMs) || endMs < startMs) return null;
        return (endMs - startMs) / (1000 * 60 * 60 * 24);
      })
      .filter((value) => Number.isFinite(value));
    const avgPaymentDays = paidDurations.length
      ? Number((paidDurations.reduce((sum, value) => sum + value, 0) / paidDurations.length).toFixed(1))
      : null;
    const latest = sorted[0] || null;
    const outstandingCents = sorted
      .filter((invoice) => ["sent", "overdue"].includes(String(invoice?.status || "").toLowerCase()))
      .reduce((sum, invoice) => sum + toCents(invoice?.total_inc_vat), 0);
    out.set(userId, {
      avg_payment_days: avgPaymentDays,
      last_invoice_tracking: formatInvoiceTrackingLabel(latest),
      last_invoice_status: String(latest?.status || "").trim().toLowerCase() || null,
      last_invoice_at: latest?.updated_at || latest?.sent_at || latest?.created_at || null,
      outstanding_balance: fromCents(outstandingCents),
    });
  });
  return out;
}

function buildInvoiceListResponseRows(invoices = []) {
  return (Array.isArray(invoices) ? invoices : []).map((invoice) => ({
    id: invoice?.id || null,
    user_id: invoice?.user_id || null,
    reference: toInvoiceReference(invoice?.id),
    company_name: String(invoice?.company_name || "").trim() || "Client account",
    contact_email: normalizeEmail(invoice?.contact_email || ""),
    period_start: invoice?.period_start || null,
    period_end: invoice?.period_end || null,
    due_at: invoice?.due_at || null,
    issued_at: invoice?.issued_at || null,
    sent_at: invoice?.sent_at || null,
    paid_at: invoice?.paid_at || null,
    status: String(invoice?.status || "draft").trim().toLowerCase(),
    payment_mode: String(invoice?.payment_mode || "invoice").trim().toLowerCase(),
    subtotal_ex_vat: fromCents(toCents(invoice?.subtotal_ex_vat)),
    vat_amount: fromCents(toCents(invoice?.vat_amount)),
    total_inc_vat: fromCents(toCents(invoice?.total_inc_vat)),
    labels_count: Number(invoice?.labels_count) || 0,
    line_count: Number(invoice?.line_count) || 0,
    reminder_stage: Number(invoice?.reminder_stage) || 0,
    last_reminder_sent_at: invoice?.last_reminder_sent_at || null,
    tracking_label: formatInvoiceTrackingLabel(invoice),
    updated_at: invoice?.updated_at || null,
  }));
}

async function listClientBillingPreferences(env, limit = 2000) {
  const safeLimit = Math.max(1, Math.min(5000, Number(limit) || 2000));
  const params = new URLSearchParams();
  params.set("select", "user_id,invoice_enabled,card_enabled,updated_at,updated_by");
  params.set("order", "updated_at.desc.nullslast");
  params.set("limit", String(safeLimit));
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${CLIENT_BILLING_PREF_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (/relation .*client_billing_preferences/i.test(details)) {
      return [];
    }
    throw new Error(`Could not load client billing settings (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function saveClientBillingPreference(env, adminUserId, userId, payload) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) {
    throw new Error("Client id is required.");
  }
  const normalized = normalizeClientBillingPreference(payload);
  const response = await supabaseServiceRequest(
    env,
    `/rest/v1/${CLIENT_BILLING_PREF_TABLE}?on_conflict=user_id`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=representation",
      },
      body: JSON.stringify([
        {
          user_id: safeUserId,
          invoice_enabled: normalized.invoice_enabled,
          card_enabled: normalized.card_enabled,
          updated_at: new Date().toISOString(),
          updated_by: adminUserId || null,
        },
      ]),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not save billing settings (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return normalizeClientBillingPreference(Array.isArray(rows) && rows.length ? rows[0] : normalized);
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
  const [settings, invites, users, historyRows, billingPreferences, invoices] = await Promise.all([
    getAdminSettings(env),
    listRegistrationInvites(env, { limit: 50 }),
    listSupabaseUsers(env, 250),
    listGenerationHistoryRows(env, 10000),
    listClientBillingPreferences(env, 2000),
    listBillingInvoices(env, { limit: 500, allowMissing: true }),
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

  const billingByUserId = new Map();
  billingPreferences.forEach((row) => {
    const userId = String(row?.user_id || "").trim();
    if (!userId) return;
    billingByUserId.set(userId, normalizeClientBillingPreference(row));
  });
  const invoiceStatsByUserId = buildInvoiceStatsByUser(invoices);

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
      billing: billingByUserId.get(user.id) || { ...DEFAULT_CLIENT_BILLING_PREF },
      metrics: buildAdminClientMetrics(
        user,
        historyByUserId.get(user.id) || [],
        settings,
        billingByUserId.get(user.id) || DEFAULT_CLIENT_BILLING_PREF,
        invoiceStatsByUserId.get(user.id) || null
      ),
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
    billing: {
      invoices: buildInvoiceListResponseRows(invoices.slice(0, 120)),
    },
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

async function deleteSupabaseUserById(env, userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return;
  const response = await fetch(
    `${String(env.SUPABASE_URL).replace(/\/+$/, "")}/auth/v1/admin/users/${safeUserId}`,
    {
      method: "DELETE",
      headers: {
        apikey: env.SUPABASE_SERVICE_ROLE_KEY,
        Authorization: `Bearer ${env.SUPABASE_SERVICE_ROLE_KEY}`,
      },
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Rollback failed: could not delete user (${response.status}) ${details}`.trim());
  }
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

function normalizeInvoiceRunMode(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["preview", "create", "send"].includes(normalized)) return normalized;
  return "create";
}

function normalizeInvoiceStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["draft", "sent", "overdue", "paid", "cancelled"].includes(normalized)) {
    return normalized;
  }
  return "draft";
}

function normalizeInvoiceRunRequest(body = {}) {
  const periodMode = String(body?.periodMode || body?.period || "previous").trim().toLowerCase();
  const mode = normalizeInvoiceRunMode(body?.mode);
  const window = getBillingWindow(
    body?.startDate,
    body?.endDate,
    periodMode === "current" ? "current" : "previous"
  );
  return {
    mode,
    includeReminders: body?.includeReminders !== false,
    userId: String(body?.userId || "").trim(),
    periodMode,
    window,
  };
}

function getReminderTitleByStage(stage, dueAt) {
  const dueLabel = String(dueAt || "").trim()
    ? new Date(dueAt).toISOString().slice(0, 10)
    : "the due date";
  const mapping = {
    1: `Invoice due in 15 days (${dueLabel})`,
    2: `Invoice due in 7 days (${dueLabel})`,
    3: `Invoice due tomorrow (${dueLabel})`,
    4: `Invoice due today (${dueLabel})`,
    5: `Invoice is 3 days overdue`,
    6: `Invoice is 10 days overdue`,
    7: `Invoice is 15 days overdue`,
    8: `Invoice is 30 days overdue`,
  };
  return mapping[Number(stage)] || "Invoice payment reminder";
}

async function sendBillingInvoiceById(env, invoiceId, options = {}) {
  const invoiceWithItems = await getBillingInvoiceById(env, invoiceId, { withItems: true });
  if (!invoiceWithItems?.id) {
    throw new Error("Invoice not found.");
  }
  const currentStatus = normalizeInvoiceStatus(invoiceWithItems.status);
  if (currentStatus === "paid" || currentStatus === "cancelled") {
    return {
      skipped: true,
      reason: `Invoice is already ${currentStatus}.`,
      invoice: invoiceWithItems,
    };
  }

  let toEmail = getInvoiceRecipientFromSnapshot(invoiceWithItems, "");
  if (!toEmail) {
    const user = await getSupabaseUserById(env, invoiceWithItems.user_id);
    toEmail = normalizeEmail(user?.email || "");
  }
  if (!toEmail) {
    throw new Error("Invoice contact email is missing.");
  }

  const isReminder = Boolean(options?.isReminder);
  const reminderStage = Number(options?.reminderStage) || 0;
  const reminderTitle = String(options?.reminderTitle || "").trim();
  const subject = buildInvoiceEmailSubject(invoiceWithItems, { isReminder, reminderStage });
  const emailPayload = await sendResendEmail(env, {
    to: toEmail,
    subject,
    html: buildInvoiceEmailHtml(invoiceWithItems, invoiceWithItems.items || [], {
      isReminder,
      reminderTitle,
      reminderStage,
    }),
    text: buildInvoiceEmailText(invoiceWithItems, { isReminder }),
  });

  const nowIso = new Date().toISOString();
  const dueMs = Date.parse(invoiceWithItems.due_at || 0);
  const nextStatus =
    Number.isFinite(dueMs) && dueMs < Date.now() && currentStatus !== "paid" ? "overdue" : "sent";
  const patch = {
    status: nextStatus,
    email_message_id: String(emailPayload?.id || "").trim() || null,
  };
  if (!invoiceWithItems.issued_at) {
    patch.issued_at = nowIso;
  }
  if (!invoiceWithItems.sent_at) {
    patch.sent_at = nowIso;
  }
  if (isReminder) {
    patch.reminder_stage = Math.max(Number(invoiceWithItems.reminder_stage) || 0, reminderStage);
    patch.last_reminder_sent_at = nowIso;
  }
  const updated = await updateBillingInvoiceFields(env, invoiceWithItems.id, patch);
  await insertBillingInvoiceEvent(env, invoiceWithItems.id, {
    event_type: isReminder ? "reminder" : "sent",
    actor: "system",
    channel: "email",
    message: isReminder ? reminderTitle || "Reminder sent." : "Invoice sent.",
    metadata: {
      resend_id: emailPayload?.id || null,
      to: toEmail,
      reminder_stage: isReminder ? reminderStage : null,
    },
  });
  return {
    skipped: false,
    invoice: {
      ...(updated || invoiceWithItems),
      items: invoiceWithItems.items || [],
    },
    resend: emailPayload,
  };
}

async function runInvoiceReminders(env, options = {}) {
  const safeLimit = Math.max(1, Math.min(500, Number(options?.limit) || 200));
  const candidateInvoices = await listBillingInvoices(env, {
    limit: safeLimit,
    statuses: ["sent", "overdue"],
  });
  const nowMs = Date.now();
  const out = [];
  for (const invoice of candidateInvoices) {
    const dueAt = invoice?.due_at;
    const desiredStage = getReminderStageForDueDate(dueAt, nowMs);
    const currentStage = Math.max(0, Number(invoice?.reminder_stage) || 0);
    if (desiredStage <= currentStage) continue;
    const reminderTitle = getReminderTitleByStage(desiredStage, dueAt);
    try {
      const result = await sendBillingInvoiceById(env, invoice.id, {
        isReminder: true,
        reminderStage: desiredStage,
        reminderTitle,
      });
      out.push({
        invoice_id: invoice.id,
        reminder_stage: desiredStage,
        status: result?.invoice?.status || invoice.status,
      });
    } catch (error) {
      await insertBillingInvoiceEvent(env, invoice.id, {
        event_type: "failed",
        actor: "system",
        channel: "email",
        message: `Reminder failed: ${error?.message || "Unknown error."}`,
        metadata: {
          reminder_stage: desiredStage,
        },
      }).catch(() => {});
      out.push({
        invoice_id: invoice.id,
        reminder_stage: desiredStage,
        error: error?.message || "Reminder failed.",
      });
    }
  }
  return out;
}

async function buildInvoicesFromLabelHistory(env, options = {}) {
  const mode = normalizeInvoiceRunMode(options?.mode);
  const billingWindow = options?.window || getPreviousMonthWindow();
  const onlyUserId = String(options?.userId || "").trim();
  const runTimestamp = new Date().toISOString();
  const termsDays = getInvoiceTermsDays(env);
  const settings = await getAdminSettings(env);

  const [historyRows, users, billingPrefs] = await Promise.all([
    listUnbilledGenerationRows(env, { window: billingWindow, userId: onlyUserId }),
    listSupabaseUsers(env, 1000),
    listClientBillingPreferences(env, 5000),
  ]);

  const userById = new Map();
  users.forEach((user) => {
    const userId = String(user?.id || "").trim();
    if (userId) userById.set(userId, user);
  });

  const billingByUserId = new Map();
  billingPrefs.forEach((pref) => {
    const userId = String(pref?.user_id || "").trim();
    if (userId) billingByUserId.set(userId, normalizeClientBillingPreference(pref));
  });

  const rowsByUser = new Map();
  historyRows.forEach((row) => {
    const userId = String(row?.user_id || "").trim();
    if (!userId) return;
    if (!rowsByUser.has(userId)) rowsByUser.set(userId, []);
    rowsByUser.get(userId).push(row);
  });

  const result = {
    mode,
    period_start: billingWindow.startIsoDate,
    period_end: billingWindow.endIsoDate,
    users_scanned: rowsByUser.size,
    rows_scanned: historyRows.length,
    invoices_created: 0,
    invoices_updated: 0,
    invoices_sent: 0,
    rows_marked_billed: 0,
    skipped: [],
    invoices: [],
    totals: {
      subtotal_ex_vat: 0,
      vat_amount: 0,
      total_inc_vat: 0,
    },
    settings: {
      client_discount_pct: normalizePercent(settings?.client_discount_pct, DEFAULT_ADMIN_SETTINGS.client_discount_pct),
      carrier_discount_pct: normalizePercent(settings?.carrier_discount_pct, DEFAULT_ADMIN_SETTINGS.carrier_discount_pct),
    },
  };

  for (const [userId, userRows] of rowsByUser.entries()) {
    const billingPref = billingByUserId.get(userId) || { ...DEFAULT_CLIENT_BILLING_PREF };
    const paymentMode = getClientPaymentMode(billingPref);
    if (!billingPref.invoice_enabled) {
      result.skipped.push({
        user_id: userId,
        reason: "Invoice billing disabled for this client.",
      });
      continue;
    }
    const invoiceData = buildInvoiceItemsFromHistoryRows(userRows, DEFAULT_VAT_RATE);
    const totals = invoiceData?.totals || {
      subtotal_ex_vat: 0,
      vat_amount: 0,
      total_inc_vat: 0,
      labels_count: 0,
      line_count: 0,
      vat_rate: DEFAULT_VAT_RATE,
    };
    if (!invoiceData?.items?.length) {
      result.skipped.push({
        user_id: userId,
        reason: "No invoiceable rows found.",
      });
      continue;
    }

    const user = userById.get(userId) || null;
    const profile = mapInvoiceProfileFromUser(user);
    const existing = await getBillingInvoiceByPeriod(
      env,
      userId,
      billingWindow.startIsoDate,
      billingWindow.endIsoDate
    );
    if (existing && ["paid", "cancelled"].includes(normalizeInvoiceStatus(existing.status))) {
      result.skipped.push({
        user_id: userId,
        invoice_id: existing.id,
        reason: `Invoice already ${normalizeInvoiceStatus(existing.status)}.`,
      });
      continue;
    }

    const dueAt =
      String(existing?.due_at || "").trim() ||
      addUtcDays(new Date(`${billingWindow.endIsoDate}T00:00:00.000Z`), termsDays).toISOString();
    const baseStatus = normalizeInvoiceStatus(existing?.status || "draft");
    const nextStatus = ["sent", "overdue", "draft"].includes(baseStatus) ? baseStatus : "draft";
    const invoicePayload = {
      user_id: userId,
      period_start: billingWindow.startIsoDate,
      period_end: billingWindow.endIsoDate,
      due_at: dueAt,
      issued_at: existing?.issued_at || null,
      sent_at: existing?.sent_at || null,
      paid_at: existing?.paid_at || null,
      status: nextStatus,
      payment_mode: paymentMode,
      currency: "EUR",
      vat_rate: totals.vat_rate,
      subtotal_ex_vat: totals.subtotal_ex_vat,
      vat_amount: totals.vat_amount,
      total_inc_vat: totals.total_inc_vat,
      labels_count: totals.labels_count,
      line_count: totals.line_count,
      reminder_stage: Number(existing?.reminder_stage) || 0,
      last_reminder_sent_at: existing?.last_reminder_sent_at || null,
      email_message_id: existing?.email_message_id || null,
      company_name: profile.company_name,
      contact_name: profile.contact_name,
      contact_email: profile.contact_email,
      customer_id: profile.customer_id,
      account_manager: profile.account_manager,
      payment_reference: existing?.payment_reference || null,
      payment_received_amount: existing?.payment_received_amount ?? null,
      metadata:
        existing?.metadata && typeof existing.metadata === "object"
          ? existing.metadata
          : {
              source: "portal_label_generations",
            },
      updated_at: runTimestamp,
    };

    const previewInvoice = {
      id: existing?.id || null,
      user_id: userId,
      company_name: profile.company_name,
      contact_email: profile.contact_email,
      period_start: billingWindow.startIsoDate,
      period_end: billingWindow.endIsoDate,
      due_at: dueAt,
      status: nextStatus,
      payment_mode: paymentMode,
      subtotal_ex_vat: totals.subtotal_ex_vat,
      vat_amount: totals.vat_amount,
      total_inc_vat: totals.total_inc_vat,
      labels_count: totals.labels_count,
      line_count: totals.line_count,
      reminder_stage: Number(existing?.reminder_stage) || 0,
    };

    if (mode === "preview") {
      result.invoices.push({
        ...previewInvoice,
        reference: previewInvoice.id ? toInvoiceReference(previewInvoice.id) : "INV-PREVIEW",
      });
      result.totals.subtotal_ex_vat = fromCents(
        toCents(result.totals.subtotal_ex_vat) + toCents(totals.subtotal_ex_vat)
      );
      result.totals.vat_amount = fromCents(
        toCents(result.totals.vat_amount) + toCents(totals.vat_amount)
      );
      result.totals.total_inc_vat = fromCents(
        toCents(result.totals.total_inc_vat) + toCents(totals.total_inc_vat)
      );
      continue;
    }

    const persisted = await upsertBillingInvoice(env, invoicePayload);
    if (!persisted?.id) {
      throw new Error("Invoice could not be persisted.");
    }
    await replaceBillingInvoiceItems(env, persisted.id, invoiceData.items);
    const generationIds = userRows.map((row) => row?.id).filter(Boolean);
    await markGenerationRowsBilled(env, generationIds, persisted.id);
    await insertBillingInvoiceEvent(env, persisted.id, {
      event_type: existing?.id ? "updated" : "created",
      actor: "system",
      channel: "portal",
      message: existing?.id ? "Invoice updated from label history." : "Invoice created from label history.",
      metadata: {
        rows_count: userRows.length,
        labels_count: totals.labels_count,
      },
    });

    if (existing?.id) {
      result.invoices_updated += 1;
    } else {
      result.invoices_created += 1;
    }
    result.rows_marked_billed += generationIds.filter((id) => isUuid(id)).length;
    result.totals.subtotal_ex_vat = fromCents(
      toCents(result.totals.subtotal_ex_vat) + toCents(totals.subtotal_ex_vat)
    );
    result.totals.vat_amount = fromCents(
      toCents(result.totals.vat_amount) + toCents(totals.vat_amount)
    );
    result.totals.total_inc_vat = fromCents(
      toCents(result.totals.total_inc_vat) + toCents(totals.total_inc_vat)
    );

    let finalInvoice = persisted;
    if (mode === "send") {
      try {
        const sendResult = await sendBillingInvoiceById(env, persisted.id, {
          isReminder: false,
        });
        if (!sendResult?.skipped) {
          result.invoices_sent += 1;
        }
        finalInvoice = sendResult?.invoice || persisted;
      } catch (error) {
        await insertBillingInvoiceEvent(env, persisted.id, {
          event_type: "failed",
          actor: "system",
          channel: "email",
          message: `Invoice send failed: ${error?.message || "Unknown error."}`,
          metadata: {},
        }).catch(() => {});
        result.skipped.push({
          user_id: userId,
          invoice_id: persisted.id,
          reason: error?.message || "Invoice send failed.",
        });
      }
    }

    result.invoices.push({
      ...buildInvoiceListResponseRows([finalInvoice])[0],
    });
  }

  return result;
}

async function computeCurrentMonthUnbilledTotal(env, userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return 0;
  const monthWindow = getCurrentMonthWindow();
  const rows = await listUnbilledGenerationRows(env, {
    window: monthWindow,
    userId: safeUserId,
  });
  const totalCents = rows.reduce((sum, row) => sum + toCents(row?.total_price), 0);
  return fromCents(totalCents);
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

async function handleAdminClientBillingSave(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to update client billing settings." }, 403);
  }

  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error.message || "Invalid request body." }, 400);
  }

  const targetUserId = String(body?.userId || "").trim();
  if (!targetUserId) {
    return jsonResponse({ error: "Client id is required." }, 400);
  }
  const invoiceEnabled = body?.invoiceEnabled !== false;
  const cardEnabled = body?.cardEnabled === true;
  if (!invoiceEnabled && !cardEnabled) {
    return jsonResponse({ error: "At least one payment method must remain enabled." }, 400);
  }

  try {
    const billing = await saveClientBillingPreference(env, user.id, targetUserId, {
      invoice_enabled: invoiceEnabled,
      card_enabled: cardEnabled,
    });
    return jsonResponse({
      ok: true,
      userId: targetUserId,
      billing,
    });
  } catch (error) {
    return jsonResponse(
      { error: error?.message || "Could not update client billing settings." },
      500
    );
  }
}

async function handleAdminInvoicesList(request, env, url) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to access billing invoices." }, 403);
  }
  const limit = Math.max(1, Math.min(300, Number(url.searchParams.get("limit")) || 120));
  const userId = String(url.searchParams.get("userId") || "").trim();
  const statusFilter = String(url.searchParams.get("status") || "").trim().toLowerCase();
  const statuses = statusFilter
    ? statusFilter
        .split(",")
        .map((entry) => entry.trim())
        .filter(Boolean)
    : [];
  try {
    const invoices = await listBillingInvoices(env, {
      limit,
      userId,
      statuses,
    });
    return jsonResponse({
      invoices: buildInvoiceListResponseRows(invoices),
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not load invoices." }, 500);
  }
}

async function handleAdminInvoicesRun(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to run billing." }, 403);
  }
  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Invalid request body." }, 400);
  }
  const runRequest = normalizeInvoiceRunRequest(body || {});
  try {
    const result = await buildInvoicesFromLabelHistory(env, runRequest);
    let reminders = [];
    if (runRequest.mode !== "preview" && runRequest.includeReminders) {
      reminders = await runInvoiceReminders(env, { limit: 300 });
    }
    return jsonResponse({
      ok: true,
      run: result,
      reminders,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Billing run failed." }, 500);
  }
}

async function handleAdminInvoiceSend(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to send invoices." }, 403);
  }
  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Invalid request body." }, 400);
  }
  const invoiceId = String(body?.invoiceId || "").trim();
  if (!invoiceId) {
    return jsonResponse({ error: "Invoice id is required." }, 400);
  }
  try {
    const result = await sendBillingInvoiceById(env, invoiceId, {
      isReminder: false,
    });
    return jsonResponse({
      ok: true,
      skipped: Boolean(result?.skipped),
      invoice: buildInvoiceListResponseRows([result?.invoice || {}])[0] || null,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not send invoice." }, 500);
  }
}

function buildAdminBillingTestInvoice(env, toEmail) {
  const now = new Date();
  const nowIso = now.toISOString();
  return {
    invoice: {
      id: crypto.randomUUID(),
      company_name: "Test Client",
      contact_email: toEmail,
      period_start: nowIso.slice(0, 10),
      period_end: nowIso.slice(0, 10),
      due_at: addUtcDays(now, getInvoiceTermsDays(env)).toISOString(),
      subtotal_ex_vat: 120,
      vat_amount: 25.2,
      total_inc_vat: 145.2,
      vat_rate: DEFAULT_VAT_RATE,
    },
    items: [
      {
        service_type: "Economy",
        recipient_name: "Sample Recipient",
        recipient_country: "Belgium",
        amount_ex_vat: 120,
      },
    ],
  };
}

async function sendAdminBillingTestSequenceEmails(env, toEmail) {
  const { invoice, items } = buildAdminBillingTestInvoice(env, toEmail);
  const steps = [
    { reminderStage: 0, isReminder: false, key: "initial" },
    { reminderStage: 1, isReminder: true, key: "due_15_days" },
    { reminderStage: 2, isReminder: true, key: "due_7_days" },
    { reminderStage: 3, isReminder: true, key: "due_tomorrow" },
    { reminderStage: 4, isReminder: true, key: "due_today" },
    { reminderStage: 5, isReminder: true, key: "overdue_3_days" },
    { reminderStage: 6, isReminder: true, key: "overdue_10_days" },
    { reminderStage: 7, isReminder: true, key: "overdue_15_days" },
    { reminderStage: 8, isReminder: true, key: "overdue_30_days" },
  ];

  const results = [];
  for (let index = 0; index < steps.length; index += 1) {
    const step = steps[index];
    const subject = buildInvoiceEmailSubject(invoice, {
      isReminder: step.isReminder,
      reminderStage: step.reminderStage,
    });
    try {
      const response = await sendResendEmail(env, {
        to: toEmail,
        subject,
        html: buildInvoiceEmailHtml(invoice, items, {
          isReminder: step.isReminder,
          reminderStage: step.reminderStage,
        }),
        text: buildInvoiceEmailText(invoice, {
          isReminder: step.isReminder,
        }),
      });
      results.push({
        step: step.key,
        reminder_stage: step.reminderStage,
        subject,
        resend_id: response?.id || null,
        ok: true,
      });
    } catch (error) {
      results.push({
        step: step.key,
        reminder_stage: step.reminderStage,
        subject,
        error: error?.message || "Could not send email.",
        ok: false,
      });
    }
    if (index < steps.length - 1) {
      await new Promise((resolve) => setTimeout(resolve, 120));
    }
  }
  return {
    to: toEmail,
    total_count: steps.length,
    sent_count: results.filter((row) => row.ok).length,
    failed_count: results.filter((row) => !row.ok).length,
    results,
  };
}

async function handleAdminInvoiceSendTest(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to send billing tests." }, 403);
  }
  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Invalid request body." }, 400);
  }
  const toEmail = normalizeEmail(body?.toEmail || user.email || "");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    return jsonResponse({ error: "A valid test email is required." }, 400);
  }
  try {
    const { invoice: testInvoice, items: testItems } = buildAdminBillingTestInvoice(env, toEmail);
    const resendResponse = await sendResendEmail(env, {
      to: toEmail,
      subject: buildInvoiceEmailSubject(testInvoice, { isReminder: false, reminderStage: 0 }),
      html: buildInvoiceEmailHtml(testInvoice, testItems, {
        isReminder: false,
        reminderStage: 0,
      }),
      text: buildInvoiceEmailText(testInvoice, {
        isReminder: false,
      }),
    });
    return jsonResponse({
      ok: true,
      to: toEmail,
      resendId: resendResponse?.id || null,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not send test invoice email." }, 500);
  }
}

async function handleAdminInvoiceSendTestSequence(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to send billing tests." }, 403);
  }
  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Invalid request body." }, 400);
  }
  const toEmail = normalizeEmail(body?.toEmail || user.email || "");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    return jsonResponse({ error: "A valid test email is required." }, 400);
  }
  try {
    const sequence = await sendAdminBillingTestSequenceEmails(env, toEmail);
    return jsonResponse({
      ok: sequence.failed_count === 0,
      ...sequence,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not send follow-up sequence." }, 500);
  }
}

async function handleAdminReportsSendTest(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to send report tests." }, 403);
  }
  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Invalid request body." }, 400);
  }
  const toEmail = normalizeEmail(body?.toEmail || user.email || "");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    return jsonResponse({ error: "A valid test email is required." }, 400);
  }
  const profitAmountRaw = Number(body?.profitAmount);
  const profitAmount = Number.isFinite(profitAmountRaw) ? Math.max(0, profitAmountRaw) : 1248.4;
  const reportsUrl =
    String(body?.reportsUrl || getReportsPortalUrl(request, env) || "").trim() ||
    "https://portal.shipide.com/reports?range=monthly";
  const reportsFromEmail = String(env.REPORTS_FROM_EMAIL || "reports@shipide.com").trim();
  const reportsFromName = String(env.REPORTS_FROM_NAME || "Shipide Reports").trim();
  try {
    const resendResponse = await sendResendEmail(env, {
      to: toEmail,
      fromName: reportsFromName,
      fromEmail: reportsFromEmail,
      subject: buildMonthlyReportsEmailSubject(),
      html: buildMonthlyReportsEmailHtml({
        profitAmount,
        reportsUrl,
      }),
      text: buildMonthlyReportsEmailText({
        profitAmount,
        reportsUrl,
      }),
    });
    return jsonResponse({
      ok: true,
      to: toEmail,
      resendId: resendResponse?.id || null,
      profitAmount,
      reportsUrl,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not send test reports email." }, 500);
  }
}

async function handleAdminInvoiceMarkPaid(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  if (!canManageRegistrationInvites(user, env)) {
    return jsonResponse({ error: "You are not allowed to update invoice payment." }, 403);
  }
  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Invalid request body." }, 400);
  }
  const invoiceId = String(body?.invoiceId || "").trim();
  if (!invoiceId) {
    return jsonResponse({ error: "Invoice id is required." }, 400);
  }
  const paymentReference = String(body?.paymentReference || "").trim();
  const paidAtRaw = String(body?.paidAt || "").trim();
  const paidAt = paidAtRaw ? new Date(paidAtRaw).toISOString() : new Date().toISOString();
  try {
    const invoice = await getBillingInvoiceById(env, invoiceId);
    if (!invoice?.id) {
      return jsonResponse({ error: "Invoice not found." }, 404);
    }
    const updated = await updateBillingInvoiceFields(env, invoice.id, {
      status: "paid",
      paid_at: paidAt,
      payment_reference: paymentReference || null,
      payment_received_amount: fromCents(toCents(invoice.total_inc_vat)),
    });
    await insertBillingInvoiceEvent(env, invoice.id, {
      event_type: "paid",
      actor: normalizeEmail(user.email || "") || user.id,
      channel: "manual",
      message: "Marked as paid by admin.",
      metadata: {
        payment_reference: paymentReference || null,
      },
    });
    return jsonResponse({
      ok: true,
      invoice: buildInvoiceListResponseRows([updated || invoice])[0] || null,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not mark invoice as paid." }, 500);
  }
}

async function handleBillingOverview(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  try {
    const billingPref = await getClientBillingPreferenceForUser(env, user.id);
    const paymentMode = getClientPaymentMode(billingPref);
    const settings = await getAdminSettings(env);
    const currentMonth = getCurrentMonthWindow();
    const invoices = await listBillingInvoices(env, {
      userId: user.id,
      limit: 120,
    });
    let wallet = {
      user_id: user.id,
      balance_cents: 0,
      currency: DEFAULT_BILLING_CURRENCY,
      updated_at: null,
    };
    let pendingTopups = [];
    let walletTransactions = [];
    try {
      [wallet, pendingTopups, walletTransactions] = await Promise.all([
        getOrCreateBillingWallet(env, user.id),
        listBillingTopups(env, {
          userId: user.id,
          statuses: ["pending", "received"],
          limit: 80,
          allowMissing: true,
        }),
        listBillingWalletTransactions(env, {
          userId: user.id,
          limit: 24,
          allowMissing: true,
        }),
      ]);
    } catch (walletError) {
      const message = String(walletError?.message || "");
      if (!/Wallet schema missing/i.test(message)) {
        throw walletError;
      }
    }
    const activeInvoices = invoices.filter((invoice) =>
      ["draft", "sent", "overdue"].includes(normalizeInvoiceStatus(invoice?.status))
    );
    const paidInvoices = invoices.filter(
      (invoice) => normalizeInvoiceStatus(invoice?.status) === "paid"
    );
    const nextDraft = activeInvoices.find(
      (invoice) => normalizeInvoiceStatus(invoice?.status) === "draft"
    );
    const outstandingCents = activeInvoices
      .filter((invoice) => ["sent", "overdue"].includes(normalizeInvoiceStatus(invoice?.status)))
      .reduce((sum, invoice) => sum + toCents(invoice?.total_inc_vat), 0);
    const projectedExVat = await computeCurrentMonthUnbilledTotal(env, user.id);
    const projectedVat = fromCents(Math.round(toCents(projectedExVat) * DEFAULT_VAT_RATE));
    const projectedIncl = fromCents(toCents(projectedExVat) + toCents(projectedVat));

    const nextInvoiceDate = nextDraft?.period_end
      ? `${nextDraft.period_end}T00:00:00.000Z`
      : `${currentMonth.endIsoDate}T00:00:00.000Z`;
    const nextExVat = nextDraft
      ? fromCents(toCents(nextDraft.subtotal_ex_vat))
      : projectedExVat;
    const nextVat = nextDraft ? fromCents(toCents(nextDraft.vat_amount)) : projectedVat;
    const nextIncl = nextDraft ? fromCents(toCents(nextDraft.total_inc_vat)) : projectedIncl;
    const pendingTopupCents = pendingTopups.reduce(
      (sum, topup) => sum + Math.max(0, Number(topup?.amount_cents) || 0),
      0
    );
    const walletBalanceCents = Math.max(0, Number(wallet?.balance_cents) || 0);

    return jsonResponse({
      payment_mode: paymentMode,
      invoice_enabled: Boolean(billingPref.invoice_enabled),
      card_enabled: Boolean(billingPref.card_enabled),
      wallet_enabled: true,
      client_discount_pct: normalizePercent(
        settings?.client_discount_pct,
        DEFAULT_ADMIN_SETTINGS.client_discount_pct
      ),
      carrier_discount_pct: normalizePercent(
        settings?.carrier_discount_pct,
        DEFAULT_ADMIN_SETTINGS.carrier_discount_pct
      ),
      next_invoice_amount_ex_vat: billingPref.invoice_enabled ? nextExVat : 0,
      next_invoice_vat_amount: billingPref.invoice_enabled ? nextVat : 0,
      next_invoice_amount_incl_vat: billingPref.invoice_enabled ? nextIncl : 0,
      next_invoice_date: billingPref.invoice_enabled ? nextInvoiceDate : null,
      current_period_start: currentMonth.startIsoDate,
      current_period_end: currentMonth.endIsoDate,
      outstanding_balance_incl_vat: fromCents(outstandingCents),
      open_invoice_count: activeInvoices.length,
      paid_invoice_count: paidInvoices.length,
      last_invoice_tracking: formatInvoiceTrackingLabel(activeInvoices[0] || null),
      open_invoices: activeInvoices.slice(0, 8).map(mapBillingOverviewInvoiceRow).filter(Boolean),
      paid_invoices: paidInvoices.slice(0, 8).map(mapBillingOverviewInvoiceRow).filter(Boolean),
      wallet_balance_eur: fromCents(walletBalanceCents),
      wallet_pending_topups_eur: fromCents(pendingTopupCents),
      wallet_currency: String(wallet?.currency || DEFAULT_BILLING_CURRENCY).trim() || DEFAULT_BILLING_CURRENCY,
      wallet_transactions: walletTransactions.slice(0, 10).map(mapBillingWalletTransactionRow),
      recent_topups: pendingTopups.slice(0, 8).map(mapBillingTopupRow),
      iban_instructions: getBillingIbanConfig(env),
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not load billing overview." }, 500);
  }
}

async function handleBillingTopups(request, env, url) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  const limit = Math.max(1, Math.min(100, Number(url.searchParams.get("limit")) || 20));
  try {
    const topups = await listBillingTopups(env, {
      userId: user.id,
      limit,
      allowMissing: true,
    });
    return jsonResponse({
      topups: topups.map(mapBillingTopupRow),
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not load top-up history." }, 500);
  }
}

async function handleBillingTopupRequest(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Invalid request body." }, 400);
  }
  const rawAmount = body?.amount ?? body?.amountEur ?? null;
  try {
    const result = await createBillingTopupRequest(env, user, rawAmount);
    return jsonResponse({
      ok: true,
      ...result,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not create top-up request." }, 500);
  }
}

async function handleBillingCheckout(request, env) {
  const user = await getAuthenticatedUser(request, env);
  if (!user?.id) {
    return jsonResponse({ error: "Authentication required." }, 401);
  }
  let body = {};
  try {
    body = await readJsonBody(request);
  } catch (error) {
    return jsonResponse({ error: error?.message || "Invalid request body." }, 400);
  }
  const method = String(body?.method || "").trim().toLowerCase();
  const amount = Number(body?.amount || body?.amountEur || body?.amountInclVat);
  if (!Number.isFinite(amount) || amount <= 0) {
    return jsonResponse({ error: "A valid checkout amount is required." }, 400);
  }
  if (!["invoice", "card", "wallet"].includes(method)) {
    return jsonResponse({ error: "Invalid payment method." }, 400);
  }
  try {
    const billingPref = await getClientBillingPreferenceForUser(env, user.id);
    if (method === "invoice" && !billingPref.invoice_enabled) {
      return jsonResponse({ error: "Invoice payment is not enabled for this account." }, 403);
    }
    if (method === "card" && !billingPref.card_enabled) {
      return jsonResponse({ error: "Card payment is not enabled for this account." }, 403);
    }
    if (method === "wallet") {
      const result = await debitWalletForCheckout(env, user, amount, {
        labels_count: Number(body?.labelsCount) || 0,
        service: String(body?.service || "").trim(),
      });
      return jsonResponse({
        ok: true,
        method: "wallet",
        charged_amount_eur: fromCents(toCents(amount)),
        wallet_balance_eur: fromCents(Math.max(0, Number(result.wallet?.balance_cents) || 0)),
        transaction_reference: result.transaction_reference,
      });
    }
    if (method === "card") {
      const authId = `card_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
      return jsonResponse({
        ok: true,
        method: "card",
        charged_amount_eur: fromCents(toCents(amount)),
        authorization_id: authId,
      });
    }
    return jsonResponse({
      ok: true,
      method: "invoice",
      charged_amount_eur: fromCents(toCents(amount)),
      note: "Queued for month-end invoice.",
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Could not process checkout." }, 500);
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
    const contract = await getActiveClickwrapContract(env);
    const invite = await getRegistrationInviteByToken(env, token);
    if (!invite || invite.claimed_at || inviteIsRevoked(invite) || inviteIsExpired(invite)) {
      return jsonResponse({ error: "This registration link is invalid or expired." }, 410);
    }
    return jsonResponse({
      valid: true,
      invite: mapInviteToPublic(invite),
      contract: mapClickwrapContractToPublic(contract),
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
  const agreement = normalizeClickwrapAgreement(body?.agreement);
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

  let createdUser = null;
  try {
    const contract = await getActiveClickwrapContract(env);
    const agreementError = validateClickwrapAgreement(agreement, contract);
    if (agreementError) {
      return jsonResponse({ error: agreementError }, 400);
    }

    const invite = await getRegistrationInviteByToken(env, token);
    if (!invite || invite.claimed_at || inviteIsRevoked(invite) || inviteIsExpired(invite)) {
      return jsonResponse({ error: "This registration link is invalid or expired." }, 410);
    }

    const metadata = {
      ...buildRegistrationMetadata(invite, profile, email, preferredLanguage),
      registration_terms_version: contract.version,
      registration_terms_hash: contract.hash_sha256,
      registration_terms_accepted_at: agreement.agreedAt,
    };
    createdUser = await createSupabaseUserWithMetadata(env, email, password, metadata);
    await ensureAccountSettingsRow(env, createdUser.id);

    const proof = await buildClickwrapEvidence(
      env,
      request,
      invite,
      contract,
      agreement,
      createdUser.id,
      email
    );
    const acceptance = await insertClickwrapAcceptance(env, {
      user_id: createdUser.id,
      invite_id: invite.id || null,
      accepted_email: email,
      invited_email: normalizeEmail(invite?.invited_email || ""),
      contract_id: contract.id || null,
      contract_version: contract.version,
      contract_hash_sha256: contract.hash_sha256,
      agreement_method: "scroll_clickwrap",
      scrolled_to_end: true,
      scrolled_to_end_at: agreement.scrolledToEndAt,
      agreed_at: agreement.agreedAt,
      ip_address: proof.payload.ip_address || null,
      user_agent: proof.payload.user_agent || null,
      client_timezone: agreement.clientTimezone || null,
      client_locale: agreement.clientLocale || null,
      proof_digest_sha256: proof.digest,
      proof_signature_hmac: proof.signature,
      proof_payload: proof.payload,
    });

    await markRegistrationInviteClaimed(env, invite.id, createdUser.id, email);

    return jsonResponse({
      ok: true,
      email,
      userId: createdUser.id,
      agreementReceiptId: acceptance?.id || null,
    });
  } catch (error) {
    if (createdUser?.id) {
      await deleteSupabaseUserById(env, createdUser.id).catch(() => {});
    }
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
