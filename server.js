const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const dns = require("node:dns").promises;
const net = require("node:net");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");
const fontkit = require("@pdf-lib/fontkit");
const { createClient } = require("@supabase/supabase-js");

const HOST = "0.0.0.0";
const PORT = Number(process.env.PORT) || 4173;
const ROOT = __dirname;
const INDEX_FILE = path.join(ROOT, "index.html");
const INVOICE_PREVIEW_FILE = path.join(ROOT, "invoice-preview.html");
const DOCUMENTS_PREVIEW_FILE = path.join(ROOT, "documents-preview.html");
const IBAN_TOPUP_PREVIEW_FILE = path.join(ROOT, "iban-topup-preview.html");
const LANDING_PLATFORM_MOCK_FILE = path.join(ROOT, "landing-platform-mock.html");
const SHIPPING_DATA_CLEANER_FILE = path.join(ROOT, "shipping-data-cleaner.html");
const MAX_BODY_BYTES = 32 * 1024 * 1024;
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
const SHOPIFY_COMPLIANCE_WEBHOOK_TOPICS = Object.freeze([
  "customers/data_request",
  "customers/redact",
  "shop/redact",
]);
const WIX_APP_ID = String(process.env.WIX_APP_ID || "").trim();
const WIX_APP_SECRET = String(process.env.WIX_APP_SECRET || "").trim();
const WIX_APP_INSTALL_URL = String(process.env.WIX_APP_INSTALL_URL || "").trim();
const LINKEDIN_CLIENT_ID = String(process.env.LINKEDIN_CLIENT_ID || "").trim();
const LINKEDIN_CLIENT_SECRET = String(process.env.LINKEDIN_CLIENT_SECRET || "").trim();
const LINKEDIN_REDIRECT_URI = String(process.env.LINKEDIN_REDIRECT_URI || "").trim();
const LINKEDIN_SCOPES = String(
  process.env.LINKEDIN_SCOPES || "rw_organization_admin w_organization_social"
)
  .split(/\s+/)
  .map((scope) => scope.trim())
  .filter(Boolean);
const LINKEDIN_API_VERSION = String(process.env.LINKEDIN_API_VERSION || "202602").trim() || "202602";
const LINKEDIN_PROVIDER = "linkedin";
const LINKEDIN_CONNECTION_KEY = "linkedin";
const LINKEDIN_ALLOWED_POST_ROLES = new Set([
  "ADMINISTRATOR",
  "DIRECT_SPONSORED_CONTENT_POSTER",
  "CONTENT_ADMINISTRATOR",
  "CONTENT_ADMIN",
  "RECRUITING_POSTER",
]);
const LINKEDIN_CONNECTION_STATUS_CONNECTED = "connected";
const SHOPIFY_FINANCIAL_STATUS_OPTIONS = Object.freeze([
  "paid",
  "pending",
  "authorized",
  "partially_paid",
  "partially_refunded",
  "refunded",
  "voided",
  "unpaid",
]);
const DEFAULT_SHOPIFY_FINANCIAL_STATUS = "paid";
const DEFAULT_SHOPIFY_FINANCIAL_STATUSES = Object.freeze([DEFAULT_SHOPIFY_FINANCIAL_STATUS]);
const WIX_IMPORT_STATUS_OPTIONS = Object.freeze(["PENDING", "APPROVED", "CANCELED", "REJECTED"]);
const DEFAULT_WIX_IMPORT_STATUSES = Object.freeze(["APPROVED"]);
const WOOCOMMERCE_APP_NAME = String(process.env.WOOCOMMERCE_APP_NAME || "Shipide").trim() || "Shipide";
const WOOCOMMERCE_IMPORT_STATUS_OPTIONS = Object.freeze([
  "pending",
  "processing",
  "on-hold",
  "completed",
  "cancelled",
  "refunded",
  "failed",
]);
const DEFAULT_WOOCOMMERCE_IMPORT_STATUSES = Object.freeze(["pending", "processing", "on-hold"]);
const SUPABASE_URL = String(
  process.env.SUPABASE_URL || "https://pxcqxubehvnyaubqjcrf.supabase.co"
).replace(/\/+$/, "");
const SUPABASE_PUBLISHABLE_KEY = String(
  process.env.SUPABASE_PUBLISHABLE_KEY || "sb_publishable_MfL9s44GmmR9peD1jouetw_aZOa8xVa"
).trim();
const SUPABASE_SERVICE_ROLE_KEY = String(process.env.SUPABASE_SERVICE_ROLE_KEY || "").trim();
const SHOPIFY_TOKEN_ENCRYPTION_KEY = String(process.env.SHOPIFY_TOKEN_ENCRYPTION_KEY || "").trim();
const INVOICE_VIEW_LINK_SECRET = String(process.env.INVOICE_VIEW_LINK_SECRET || "").trim();
const OAUTH_STATE_SECRET = String(process.env.OAUTH_STATE_SECRET || "").trim();
const CLICKWRAP_PROOF_SECRET = String(process.env.CLICKWRAP_PROOF_SECRET || "").trim();
const REGISTRATION_INVITES_TABLE = "registration_invites";
const SHIPMENT_EXTRACT_REQUESTS_TABLE = "shipment_extract_requests";
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
const BILLING_INVOICE_SELECT_FIELDS = "id,user_id,invoice_kind,source_topup_id,invoice_number,period_start,period_end,due_at,issued_at,sent_at,paid_at,status,payment_mode,currency,vat_rate,subtotal_ex_vat,vat_amount,total_inc_vat,labels_count,line_count,reminder_stage,last_reminder_sent_at,email_message_id,company_name,contact_name,contact_email,customer_id,account_manager,payment_reference,payment_received_amount,metadata,created_at,updated_at";
const BILLING_INVOICE_ITEM_SELECT_FIELDS = "id,invoice_id,generation_id,generated_at,service_type,quantity,recipient_name,recipient_city,recipient_country,tracking_id,label_id,unit_price_ex_vat,amount_ex_vat,vat_amount,amount_inc_vat,sort_index,metadata,created_at";
const BILLING_TOPUP_SELECT_FIELDS = "id,user_id,amount_cents,currency,status,reference,requested_at,received_at,credited_at,metadata,created_at,updated_at";
const BILLING_BANK_RECEIPTS_TABLE = "billing_bank_receipts";
const BILLING_INVOICE_PAYMENTS_TABLE = "billing_invoice_payments";
const WISE_WEBHOOK_EVENTS_TABLE = "wise_webhook_events";
const PRIVACY_REQUESTS_TABLE = "privacy_requests";
const PRIVACY_DELETE_CONFIRMATION = "DELETE MY DATA";
const PRIVACY_RETENTION_YEARS = 5;
const PRIVACY_FINANCIAL_RETENTION_YEARS = 10;
const PRIVACY_INVITE_RETENTION_DAYS = 180;
const PRIVACY_STALE_CONNECTION_DAYS = 30;
const PASSWORD_MIN_LENGTH = 10;
const DEFAULT_INVITE_EXPIRY_DAYS = 14;
const MAX_INVITE_EXPIRY_DAYS = 90;
const CLICKWRAP_ACCEPTANCE_MAX_AGE_MS = 24 * 60 * 60 * 1000;
const CLICKWRAP_CLOCK_SKEW_MS = 5 * 60 * 1000;
const DEFAULT_CLICKWRAP_CONTRACT_PDF_URL = "/assets/contracts/commAgreement-v1.pdf";
const CLICKWRAP_CONTRACT_TEMPLATE_PDF_URL = "/assets/contracts/commAgreement-v1-template.pdf";
const CLICKWRAP_CONTRACT_TEMPLATE_FILE = path.join(
  ROOT,
  "assets",
  "contracts",
  "commAgreement-v1-template.pdf"
);
const CLICKWRAP_CONTRACT_TEMPLATE_FONT_FILE = path.join(
  ROOT,
  "assets",
  "fonts",
  "PTMono-Regular.ttf"
);
const CLICKWRAP_STORAGE_BUCKET =
  String(process.env.CLICKWRAP_STORAGE_BUCKET || "clickwrap-contracts").trim()
  || "clickwrap-contracts";
const BILLING_INVOICE_STORAGE_BUCKET =
  String(process.env.BILLING_INVOICE_STORAGE_BUCKET || "billing-invoices").trim()
  || "billing-invoices";
const CLICKWRAP_PREVIEW_TTL_MS = 60 * 60 * 1000;
const CLICKWRAP_PREVIEW_MAX_ENTRIES = 128;
const ADMIN_SETTINGS_SCOPE = "global";
const DEFAULT_BILLING_TERMS_DAYS = 30;
const DEFAULT_VAT_RATE = 0;
const DEFAULT_BILLING_CURRENCY = "EUR";
const APPROVED_INVOICE_RENDERER_VERSION = "native-html-v1";
const BILLING_TOPUP_EXPIRY_MS = 7 * 24 * 60 * 60 * 1000;
const DEFAULT_IBAN_BENEFICIARY = "Cryvelin LLC";
const DEFAULT_IBAN = "BE68 5390 0754 7034";
const DEFAULT_IBAN_BIC = "KREDBEBB";
const DEFAULT_IBAN_ADDRESS = "";
const DEFAULT_IBAN_TRANSFER_NOTE =
  "Transfers are credited once received (typically 1-2 business days).";
const DEFAULT_INVOICE_PDF_IBAN = String(
  process.env.INVOICE_PDF_IBAN || "BE71 0000 1111 2222"
).trim();
const DEFAULT_INVOICE_ISSUER = Object.freeze({
  legalName: "Cryvelin LLC",
  brandName: "Shipide",
  descriptor: "Operates Shipide",
  addressLines: [
    "8 The Green STE B,",
    "Dover 19901,",
    "Delaware, USA",
  ],
  ein: "38-4390983",
});
const RESEND_API_KEY = String(process.env.RESEND_API_KEY || "").trim();
const RESEND_FROM_EMAIL = String(process.env.RESEND_FROM_EMAIL || "billing@shipide.com").trim();
const RESEND_FROM_NAME = String(process.env.RESEND_FROM_NAME || "Shipide Billing").trim();
const RESEND_REPLY_TO = String(process.env.RESEND_REPLY_TO || "").trim();
const DATA_CLEANER_SUBMISSION_EMAIL = String(
  process.env.DATA_CLEANER_SUBMISSION_EMAIL || "shipextract@shipide.com"
).trim();
const BILLING_IBAN_BENEFICIARY = String(
  process.env.BILLING_IBAN_BENEFICIARY || DEFAULT_IBAN_BENEFICIARY
).trim();
const BILLING_IBAN = String(process.env.BILLING_IBAN || DEFAULT_IBAN).trim();
const BILLING_IBAN_BIC = String(process.env.BILLING_IBAN_BIC || DEFAULT_IBAN_BIC).trim();
const BILLING_IBAN_ADDRESS = String(process.env.BILLING_IBAN_ADDRESS || DEFAULT_IBAN_ADDRESS).trim();
const BILLING_IBAN_NOTE = String(
  process.env.BILLING_IBAN_NOTE || DEFAULT_IBAN_TRANSFER_NOTE
).trim();
const REPORTS_FROM_EMAIL = String(process.env.REPORTS_FROM_EMAIL || "reports@shipide.com").trim();
const REPORTS_FROM_NAME = String(process.env.REPORTS_FROM_NAME || "Shipide Reports").trim();
const WELCOME_FROM_EMAIL = String(process.env.WELCOME_FROM_EMAIL || "welcome@shipide.com").trim();
const WELCOME_FROM_NAME = String(process.env.WELCOME_FROM_NAME || "Shipide").trim();
const WELCOME_PORTAL_URL = String(
  process.env.WELCOME_PORTAL_URL || "https://portal.shipide.com/login"
).trim();
const REPORTS_PORTAL_URL = String(
  process.env.REPORTS_PORTAL_URL || "https://portal.shipide.com/reports?range=monthly"
).trim();
const PUBLIC_APP_URL = String(process.env.PUBLIC_APP_URL || "").trim();
const DEFAULT_PUBLIC_APP_URL = "https://portal.shipide.com";
const DEFAULT_INVOICE_VIEW_LINK_TTL_MS = 30 * 24 * 60 * 60 * 1000;
const WISE_DEFAULT_LOOKBACK_DAYS = 45;
const WISE_WEBHOOK_SYNC_LOOKBACK_DAYS = 7;
const WISE_DEFAULT_API_BASE_URL = "https://api.wise.com";
const WISE_SANDBOX_API_BASE_URL = "https://api.sandbox.transferwise.tech";
const WISE_APPLY_RECEIPT_RPC = "apply_billing_bank_receipt_resolution";
const WISE_PRODUCTION_PUBLIC_KEY = `-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAvO8vXV+JksBzZAY6GhSO
XdoTCfhXaaiZ+qAbtaDBiu2AGkGVpmEygFmWP4Li9m5+Ni85BhVvZOodM9epgW3F
bA5Q1SexvAF1PPjX4JpMstak/QhAgl1qMSqEevL8cmUeTgcMuVWCJmlge9h7B1CS
D4rtlimGZozG39rUBDg6Qt2K+P4wBfLblL0k4C4YUdLnpGYEDIth+i8XsRpFlogx
CAFyH9+knYsDbR43UJ9shtc42Ybd40Afihj8KnYKXzchyQ42aC8aZ/h5hyZ28yVy
Oj3Vos0VdBIs/gAyJ/4yyQFCXYte64I7ssrlbGRaco4nKF3HmaNhxwyKyJafz19e
HwIDAQAB
-----END PUBLIC KEY-----`;
const WISE_SANDBOX_PUBLIC_KEY = `-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAwpb91cEYuyJNQepZAVfP
ZIlPZfNUefH+n6w9SW3fykqKu938cR7WadQv87oF2VuT+fDt7kqeRziTmPSUhqPU
ys/V2Q1rlfJuXbE+Gga37t7zwd0egQ+KyOEHQOpcTwKmtZ81ieGHynAQzsn1We3j
wt760MsCPJ7GMT141ByQM+yW1Bx+4SG3IGjXWyqOWrcXsxAvIXkpUD/jK/L958Cg
nZEgz0BSEh0QxYLITnW1lLokSx/dTianWPFEhMC9BgijempgNXHNfcVirg1lPSyg
z7KqoKUN0oHqWLr2U1A+7kqrl6O2nx3CKs1bj1hToT1+p4kcMoHXA7kA+VBLUpEs
VwIDAQAB
-----END PUBLIC KEY-----`;
const DEFAULT_ADMIN_SETTINGS = {
  carrier_discount_pct: 25,
  client_discount_pct: 20,
};
const DEFAULT_CLIENT_BILLING_PREF = {
  invoice_enabled: false,
  card_enabled: false,
};
const DEFAULT_CLICKWRAP_CONTRACT_VERSION = "v2.0";
const DEFAULT_CLICKWRAP_CONTRACT_TITLE = "Commercial Agreement";
const DEFAULT_CLICKWRAP_CONTRACT_BODY = `Electronic Acceptance Notice

By clicking "I agree", you consent to use an electronic signature and electronic records for this agreement.

1. Scope of Services
Shipide provides software to prepare shipment data, request transport labels, and manage related billing and reporting operations.

2. Account Security
You must keep credentials confidential and promptly report unauthorized access. Activity under your account is treated as authorized unless proven otherwise.

3. Billing and Payment
Charges may apply per label, shipment, subscription, or wallet usage.

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
  Standard: 12.2,
  Priority: 12.2,
  "International Express": 21.8,
};
const INVITE_ADMIN_EMAILS = String(process.env.INVITE_ADMIN_EMAILS || "")
  .split(",")
  .map((email) => email.trim().toLowerCase())
  .filter(Boolean);
const LOCAL_SIGNUP_PREVIEW_TOKEN = "local-signup-preview";
const LOCAL_SIGNUP_PREVIEW_INVITE = {
  id: "local-signup-preview-invite",
  invited_email: "claire@ateliermeridian.com",
  company_name: "Atelier Meridian",
  contact_name: "Claire Dupont",
  contact_email: "operations@ateliermeridian.com",
  contact_phone: "+32 2 555 01 29",
  billing_address: "Avenue Louise 120, 1050 Brussels, Belgium",
  tax_id: "BE0123456789",
  customer_id: "SHIP-24018",
  plan_name: "Growth",
  expires_at: "2099-12-31T23:59:59.000Z",
  claimed_at: null,
  revoked_at: null,
};

function getRuntimeEnvironmentName() {
  return String(process.env.NODE_ENV || process.env.APP_ENV || "").trim().toLowerCase();
}

function isProductionRuntime() {
  const environment = getRuntimeEnvironmentName();
  return environment === "production" || environment === "prod";
}

function isShopifyDevSeedOrdersEnabled() {
  if (isProductionRuntime()) return false;
  return /^(1|true|yes)$/i.test(String(process.env.ENABLE_SHOPIFY_DEV_SEED_ORDERS || "").trim());
}
const LOCAL_SIGNUP_PREVIEW_CONTRACT = {
  id: "local-signup-preview-contract",
  version: "v2.0",
  title: "Shipide Service Agreement",
  body_text:
    "Preview mode only. This local route generates a personalized agreement preview from the project template without creating a live account.",
  hash_sha256: sha256Hex("local-signup-preview-contract-v2.0"),
  pdf_url: CLICKWRAP_CONTRACT_TEMPLATE_PDF_URL,
  effective_at: null,
  source: "local-preview",
};

const oauthStateStore = new Map();
const clickwrapPreviewStore = new Map();
let clickwrapPreviewAssetsPromise = null;
let supabaseAdminClient = null;
let clickwrapStorageBucketPromise = null;
let billingInvoiceStorageBucketPromise = null;

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
  ".csv": "text/csv; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".html": "text/html; charset=utf-8",
};

const PREVIEW_VERSION_FILES = Object.freeze([
  path.join(ROOT, "invoice-preview.html"),
  path.join(ROOT, "documents-preview.html"),
  path.join(ROOT, "iban-topup-preview.html"),
  path.join(ROOT, "app.js"),
  path.join(ROOT, "styles.css"),
]);

function send(res, status, headers, body) {
  res.writeHead(status, headers);
  res.end(body);
}

function sendJson(res, status, payload, extraHeaders = {}) {
  send(
    res,
    status,
    {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      ...extraHeaders,
    },
    JSON.stringify(payload)
  );
}

function buildPreviewVersionPayload() {
  const files = PREVIEW_VERSION_FILES.map((filePath) => {
    try {
      const stats = fs.statSync(filePath);
      return {
        file: path.relative(ROOT, filePath),
        size: stats.size,
        mtimeMs: Math.round(stats.mtimeMs),
      };
    } catch (error) {
      return {
        file: path.relative(ROOT, filePath),
        missing: true,
      };
    }
  });
  const version = crypto
    .createHash("sha1")
    .update(JSON.stringify(files))
    .digest("hex");
  return {
    scope: "invoice-preview",
    version,
    files,
    generatedAt: new Date().toISOString(),
  };
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

function getPublicAppUrl(req = null) {
  const explicit = PUBLIC_APP_URL.replace(/\/+$/, "");
  if (explicit) return explicit;
  if (req) {
    return buildPublicBaseUrl(req).replace(/\/+$/, "");
  }
  return DEFAULT_PUBLIC_APP_URL;
}

function getRequiredSecret(name, value) {
  const secret = String(value || "").trim();
  if (!secret) {
    throw new Error(`${name} is required.`);
  }
  return secret;
}

function getInvoiceViewLinkSecret() {
  return getRequiredSecret("INVOICE_VIEW_LINK_SECRET", INVOICE_VIEW_LINK_SECRET);
}

function hmacSha256Hex(secret, value) {
  return crypto.createHmac("sha256", secret).update(String(value || ""), "utf8").digest("hex");
}

function makeInvoiceViewToken(payload, ttlMs = DEFAULT_INVOICE_VIEW_LINK_TTL_MS) {
  const tokenPayload = {
    ...payload,
    nonce: crypto.randomBytes(12).toString("base64url"),
    exp: Date.now() + Math.max(60 * 1000, Number(ttlMs) || DEFAULT_INVOICE_VIEW_LINK_TTL_MS),
  };
  const encoded = Buffer.from(JSON.stringify(tokenPayload), "utf8").toString("base64url");
  const signature = hmacSha256Hex(getInvoiceViewLinkSecret(), encoded);
  return `${encoded}.${signature}`;
}

function parseInvoiceViewToken(token) {
  const [encoded = "", signature = ""] = String(token || "").split(".");
  if (!encoded || !signature) return null;
  const expected = hmacSha256Hex(getInvoiceViewLinkSecret(), encoded);
  const receivedBuffer = Buffer.from(signature, "utf8");
  const expectedBuffer = Buffer.from(expected, "utf8");
  if (receivedBuffer.length !== expectedBuffer.length) return null;
  if (!crypto.timingSafeEqual(receivedBuffer, expectedBuffer)) return null;
  try {
    const payload = JSON.parse(Buffer.from(encoded, "base64url").toString("utf8"));
    if (!payload || typeof payload !== "object") return null;
    if (!payload.exp || Date.now() > Number(payload.exp)) return null;
    return payload;
  } catch (_error) {
    return null;
  }
}

function buildInvoiceViewUrl(invoice, reminderStage, options = {}) {
  const invoiceId = String(invoice?.id || "").trim();
  if (!invoiceId) return "#";
  const stage = normalizeInvoicePdfVariantStage(reminderStage);
  const publicAppUrl =
    String(options?.publicAppUrl || "").trim().replace(/\/+$/, "")
    || getPublicAppUrl(options?.request || null);
  const token = makeInvoiceViewToken({
    type: "invoice_view",
    invoiceId,
    stage,
  });
  return `${publicAppUrl}/invoice-view?token=${encodeURIComponent(token)}`;
}

function sanitizeDownloadFilename(filename, fallback) {
  const sanitized = String(filename || "").replace(/[^A-Za-z0-9._-]+/g, "_").trim();
  return sanitized || String(fallback || "invoice.pdf").replace(/[^A-Za-z0-9._-]+/g, "_");
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

function createWooCommerceStoreUrlError(
  message = "Enter a valid public HTTPS WooCommerce store URL."
) {
  const error = new Error(message);
  error.code = "woocommerce_store_url_disallowed";
  return error;
}

function isLiteralIpv4Address(value) {
  if (!/^\d{1,3}(?:\.\d{1,3}){3}$/.test(String(value || "").trim())) return false;
  return String(value || "")
    .trim()
    .split(".")
    .every((part) => {
      const numeric = Number(part);
      return Number.isInteger(numeric) && numeric >= 0 && numeric <= 255;
    });
}

function isDisallowedWooCommerceHostname(hostname) {
  const normalized = String(hostname || "")
    .trim()
    .toLowerCase()
    .replace(/\.$/, "");
  if (!normalized) return true;
  if (!normalized.includes(".")) return true;
  if (
    normalized === "localhost"
    || normalized.endsWith(".localhost")
    || normalized.endsWith(".local")
    || normalized.endsWith(".localdomain")
    || normalized.endsWith(".internal")
    || normalized.endsWith(".home.arpa")
    || normalized.endsWith(".test")
    || normalized.endsWith(".example")
    || normalized.endsWith(".invalid")
  ) {
    return true;
  }
  return net.isIP(normalized) !== 0 || isLiteralIpv4Address(normalized);
}

function isPrivateOrReservedIpv4Address(address) {
  const parts = String(address || "")
    .trim()
    .split(".")
    .map((part) => Number(part));
  if (parts.length !== 4 || parts.some((part) => !Number.isInteger(part) || part < 0 || part > 255)) {
    return true;
  }
  const [first, second, third] = parts;
  if (first === 0 || first === 10 || first === 127) return true;
  if (first === 100 && second >= 64 && second <= 127) return true;
  if (first === 169 && second === 254) return true;
  if (first === 172 && second >= 16 && second <= 31) return true;
  if (first === 192 && second === 0 && third === 0) return true;
  if (first === 192 && second === 0 && third === 2) return true;
  if (first === 192 && second === 168) return true;
  if (first === 198 && (second === 18 || second === 19)) return true;
  if (first === 198 && second === 51 && third === 100) return true;
  if (first === 203 && second === 0 && third === 113) return true;
  if (first >= 224) return true;
  return false;
}

function isPrivateOrReservedIpAddress(address) {
  const normalized = String(address || "")
    .trim()
    .toLowerCase()
    .replace(/^\[|\]$/g, "")
    .split("%")[0];
  if (!normalized) return true;
  if (normalized.startsWith("::ffff:")) {
    return isPrivateOrReservedIpAddress(normalized.slice(7));
  }
  const family = net.isIP(normalized);
  if (family === 4) {
    return isPrivateOrReservedIpv4Address(normalized);
  }
  if (family === 6) {
    if (normalized === "::" || normalized === "::1") return true;
    if (normalized.startsWith("fc") || normalized.startsWith("fd")) return true;
    if (
      normalized.startsWith("fe8")
      || normalized.startsWith("fe9")
      || normalized.startsWith("fea")
      || normalized.startsWith("feb")
    ) {
      return true;
    }
    if (normalized.startsWith("2001:db8")) return true;
    return false;
  }
  return true;
}

function normalizeWooCommerceStoreUrl(value) {
  const raw = String(value || "").trim();
  if (!raw || /\s/.test(raw)) return "";
  const candidate = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`;
  let parsed;
  try {
    parsed = new URL(candidate);
  } catch (_error) {
    return "";
  }
  if (parsed.protocol !== "https:") return "";
  if (!parsed.hostname || parsed.username || parsed.password) return "";
  if (isDisallowedWooCommerceHostname(parsed.hostname)) return "";
  const pathname = parsed.pathname.replace(/\/+$/, "");
  const storePath = pathname
    .replace(/\/wp-json(?:\/wc\/v3)?$/i, "")
    .replace(/\/wc-auth\/v1\/authorize$/i, "");
  return `${parsed.origin}${storePath === "/" ? "" : storePath}`;
}

async function assertSafeWooCommerceStoreUrlForFetch(storeUrl) {
  const normalizedStoreUrl = normalizeWooCommerceStoreUrl(storeUrl);
  if (!normalizedStoreUrl) {
    throw createWooCommerceStoreUrlError();
  }
  const parsed = new URL(normalizedStoreUrl);
  let records = [];
  try {
    records = await dns.lookup(parsed.hostname, { all: true, verbatim: true });
  } catch (_error) {
    throw createWooCommerceStoreUrlError(
      "WooCommerce store URL could not be resolved to a public HTTPS storefront."
    );
  }
  if (!Array.isArray(records) || !records.length) {
    throw createWooCommerceStoreUrlError(
      "WooCommerce store URL could not be resolved to a public HTTPS storefront."
    );
  }
  if (records.some((record) => isPrivateOrReservedIpAddress(record?.address))) {
    throw createWooCommerceStoreUrlError(
      "WooCommerce store URL must resolve to a public HTTPS storefront."
    );
  }
  return normalizedStoreUrl;
}

function getBearerToken(req) {
  const auth = String(req.headers.authorization || "");
  if (!auth.toLowerCase().startsWith("bearer ")) return "";
  return auth.slice(7).trim();
}

async function readRequestBodyBuffer(req) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > MAX_BODY_BYTES) {
      throw new Error("Request body too large.");
    }
    chunks.push(chunk);
  }
  return chunks.length ? Buffer.concat(chunks) : Buffer.alloc(0);
}

async function readTextBody(req) {
  const bodyBuffer = await readRequestBodyBuffer(req);
  return bodyBuffer.toString("utf8");
}

async function readJsonBody(req) {
  const body = (await readTextBody(req)).trim();
  if (!body) return {};
  try {
    return JSON.parse(body);
  } catch (_error) {
    throw new Error("Invalid JSON request body.");
  }
}

function safeParseJson(value) {
  try {
    return JSON.parse(String(value || ""));
  } catch (_error) {
    return null;
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

function fromBase64Url(value) {
  const normalized = String(value || "").replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized + "=".repeat((4 - (normalized.length % 4)) % 4);
  return Buffer.from(padded, "base64");
}

function createPublicLinkToken(byteLength = 16) {
  return crypto.randomBytes(byteLength).toString("base64url");
}

function normalizeShopifySessionShop(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) {
    try {
      return normalizeShopDomain(new URL(raw).hostname);
    } catch (_error) {
      return "";
    }
  }
  return normalizeShopDomain(raw.replace(/\/.*$/, ""));
}

function verifyShopifyWebhookHmac(rawBody, receivedHmac) {
  if (!SHOPIFY_API_SECRET) return false;
  const incoming = String(receivedHmac || "").trim();
  if (!incoming) return false;
  const expected = crypto
    .createHmac("sha256", SHOPIFY_API_SECRET)
    .update(Buffer.isBuffer(rawBody) ? rawBody : Buffer.from(String(rawBody || ""), "utf8"))
    .digest("base64");
  const incomingBuffer = Buffer.from(incoming, "utf8");
  const expectedBuffer = Buffer.from(expected, "utf8");
  if (incomingBuffer.length !== expectedBuffer.length) {
    return false;
  }
  return crypto.timingSafeEqual(incomingBuffer, expectedBuffer);
}

function verifyShopifySessionToken(token) {
  if (!SHOPIFY_API_SECRET || !SHOPIFY_API_KEY) return null;
  const parts = String(token || "").trim().split(".");
  if (parts.length !== 3) return null;
  const [encodedHeader, encodedPayload, encodedSignature] = parts;
  if (!encodedHeader || !encodedPayload || !encodedSignature) return null;

  const header = safeParseJson(fromBase64Url(encodedHeader).toString("utf8"));
  const payload = safeParseJson(fromBase64Url(encodedPayload).toString("utf8"));
  if (!header || typeof header !== "object" || !payload || typeof payload !== "object") {
    return null;
  }
  if (String(header.alg || "").trim().toUpperCase() !== "HS256") {
    return null;
  }

  const expectedSignature = crypto
    .createHmac("sha256", SHOPIFY_API_SECRET)
    .update(`${encodedHeader}.${encodedPayload}`)
    .digest();
  const receivedSignature = fromBase64Url(encodedSignature);
  if (receivedSignature.length !== expectedSignature.length) {
    return null;
  }
  if (!crypto.timingSafeEqual(receivedSignature, expectedSignature)) {
    return null;
  }

  const audience = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
  const audienceValid = audience.some((value) => String(value || "").trim() === SHOPIFY_API_KEY);
  if (!audienceValid) {
    return null;
  }

  const nowSec = Math.floor(Date.now() / 1000);
  const exp = Number(payload.exp);
  const nbf = Number(payload.nbf);
  if (Number.isFinite(exp) && nowSec > exp + 10) {
    return null;
  }
  if (Number.isFinite(nbf) && nowSec + 10 < nbf) {
    return null;
  }

  const shop = normalizeShopifySessionShop(payload.dest || payload.destination || payload.iss);
  if (!shop) {
    return null;
  }

  return {
    ...payload,
    shop,
  };
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

function getOauthStateSecret() {
  return getRequiredSecret("OAUTH_STATE_SECRET", OAUTH_STATE_SECRET);
}

function makeSignedStateToken(payload, ttlMs = OAUTH_STATE_TTL_MS) {
  const secret = getOauthStateSecret();
  if (!secret) {
    throw new Error("OAUTH_STATE_SECRET is not configured.");
  }
  const tokenPayload = {
    ...payload,
    nonce: crypto.randomBytes(12).toString("base64url"),
    exp: Date.now() + Math.max(60 * 1000, Number(ttlMs) || OAUTH_STATE_TTL_MS),
  };
  const encoded = Buffer.from(JSON.stringify(tokenPayload), "utf8").toString("base64url");
  const signature = hmacSha256Hex(secret, encoded);
  return `${encoded}.${signature}`;
}

function parseSignedStateToken(token) {
  const secret = getOauthStateSecret();
  if (!secret) {
    throw new Error("OAUTH_STATE_SECRET is not configured.");
  }
  const [encoded = "", signature = ""] = String(token || "").split(".");
  if (!encoded || !signature) return null;
  const expected = hmacSha256Hex(secret, encoded);
  const receivedBuffer = Buffer.from(signature, "utf8");
  const expectedBuffer = Buffer.from(expected, "utf8");
  if (receivedBuffer.length !== expectedBuffer.length) return null;
  if (!crypto.timingSafeEqual(receivedBuffer, expectedBuffer)) return null;
  try {
    const payload = JSON.parse(Buffer.from(encoded, "base64url").toString("utf8"));
    if (!payload || typeof payload !== "object") return null;
    if (!payload.exp || Date.now() > Number(payload.exp)) return null;
    return payload;
  } catch (_error) {
    return null;
  }
}

function getTokenSecretCandidates() {
  const candidates = [SHOPIFY_TOKEN_ENCRYPTION_KEY]
    .map((value) => String(value || "").trim())
    .filter(Boolean);
  const unique = [];
  candidates.forEach((candidate) => {
    if (!unique.includes(candidate)) {
      unique.push(candidate);
    }
  });
  if (!unique.length) {
    throw new Error("Token encryption key is not configured.");
  }
  return unique;
}

function getPrimaryTokenCipherKey() {
  return crypto.createHash("sha256").update(getTokenSecretCandidates()[0]).digest();
}

function createTokenDecryptError() {
  const error = new Error(
    "Stored Shopify connection could not be decrypted. Reconnect Shopify and try again."
  );
  error.code = "token_decrypt_failed";
  error.status = 409;
  return error;
}

function createWooCommerceDecryptError() {
  const error = new Error(
    "Stored WooCommerce connection could not be decrypted. Reconnect WooCommerce and try again."
  );
  error.code = "token_decrypt_failed";
  error.status = 409;
  return error;
}

function encodeBasicAuth(username, password) {
  return Buffer.from(`${String(username || "")}:${String(password || "")}`, "utf8").toString(
    "base64"
  );
}

function encryptToken(plainToken) {
  const key = getPrimaryTokenCipherKey();
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
    throw createTokenDecryptError();
  }
  const parts = raw.split(":");
  const candidateKeys = getTokenSecretCandidates().map((secret) =>
    crypto.createHash("sha256").update(secret).digest()
  );

  if (parts.length === 4) {
    const [, ivEncoded, tagEncoded, payloadEncoded] = parts;
    const iv = Buffer.from(ivEncoded, "base64url");
    const tag = Buffer.from(tagEncoded, "base64url");
    const payload = Buffer.from(payloadEncoded, "base64url");

    for (const key of candidateKeys) {
      try {
        const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
        decipher.setAuthTag(tag);
        const decrypted = Buffer.concat([decipher.update(payload), decipher.final()]);
        return decrypted.toString("utf8");
      } catch {}
    }
    throw createTokenDecryptError();
  }

  if (parts.length === 3) {
    const [, ivEncoded, encryptedEncoded] = parts;
    const iv = Buffer.from(ivEncoded, "base64url");
    const encrypted = Buffer.from(encryptedEncoded, "base64url");
    if (encrypted.length < 17) {
      throw createTokenDecryptError();
    }
    const payload = encrypted.subarray(0, encrypted.length - 16);
    const tag = encrypted.subarray(encrypted.length - 16);

    for (const key of candidateKeys) {
      try {
        const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
        decipher.setAuthTag(tag);
        const decrypted = Buffer.concat([decipher.update(payload), decipher.final()]);
        return decrypted.toString("utf8");
      } catch {}
    }
    throw createTokenDecryptError();
  }

  throw createTokenDecryptError();
}

async function supabaseServiceRequest(pathnameWithQuery, options = {}) {
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    throw new Error(
      "SUPABASE_SERVICE_ROLE_KEY is required on the server for invite validation, registration, and Supabase admin requests."
    );
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

function getSupabaseAdminClient() {
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    throw new Error(
      "SUPABASE_SERVICE_ROLE_KEY is required on the server for Supabase Storage uploads."
    );
  }
  if (!supabaseAdminClient) {
    supabaseAdminClient = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
      auth: {
        autoRefreshToken: false,
        persistSession: false,
      },
    });
  }
  return supabaseAdminClient;
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

function sha256Hex(value) {
  return crypto.createHash("sha256").update(String(value || "")).digest("hex");
}

function hmacSha256Hex(secret, message) {
  return crypto
    .createHmac("sha256", String(secret || ""))
    .update(String(message || ""))
    .digest("hex");
}

function getClickwrapProofSecret() {
  return getRequiredSecret("CLICKWRAP_PROOF_SECRET", CLICKWRAP_PROOF_SECRET);
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
    (haystack.includes("42p01") &&
      (haystack.includes(CLICKWRAP_CONTRACTS_TABLE) ||
        haystack.includes(CLICKWRAP_ACCEPTANCES_TABLE))) ||
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

function normalizeClickwrapContractSnapshot(contract) {
  const bodyText = String(contract?.body_text || contract?.bodyText || DEFAULT_CLICKWRAP_CONTRACT_BODY)
    .trim();
  const hash =
    String(contract?.hash_sha256 || contract?.hash || "")
      .trim()
      .toLowerCase()
    || sha256Hex(bodyText);
  return {
    id: contract?.id || null,
    version: String(contract?.version || DEFAULT_CLICKWRAP_CONTRACT_VERSION).trim(),
    title: String(contract?.title || DEFAULT_CLICKWRAP_CONTRACT_TITLE).trim(),
    body_text: bodyText,
    hash_sha256: hash,
    pdf_url: String(contract?.pdf_url || contract?.pdfUrl || DEFAULT_CLICKWRAP_CONTRACT_PDF_URL).trim(),
    effective_at: contract?.effective_at || contract?.effectiveAt || null,
    source: contract?.source || "snapshot",
  };
}

function getLocalSignupPreviewInvite() {
  return {
    ...LOCAL_SIGNUP_PREVIEW_INVITE,
  };
}

function getLocalSignupPreviewContract() {
  return normalizeClickwrapContractSnapshot(LOCAL_SIGNUP_PREVIEW_CONTRACT);
}

function pruneClickwrapPreviewStore() {
  const now = Date.now();
  for (const [recordId, entry] of clickwrapPreviewStore.entries()) {
    if (!entry || now >= Number(entry.expiresAt || 0)) {
      clickwrapPreviewStore.delete(recordId);
    }
  }
  if (clickwrapPreviewStore.size <= CLICKWRAP_PREVIEW_MAX_ENTRIES) {
    return;
  }
  const entries = Array.from(clickwrapPreviewStore.entries()).sort(
    (left, right) => Number(left[1]?.createdAt || 0) - Number(right[1]?.createdAt || 0)
  );
  while (entries.length && clickwrapPreviewStore.size > CLICKWRAP_PREVIEW_MAX_ENTRIES) {
    const [recordId] = entries.shift();
    clickwrapPreviewStore.delete(recordId);
  }
}

function createClickwrapPreviewRecordId() {
  return `CW-${Date.now().toString(36).toUpperCase()}-${crypto
    .randomBytes(2)
    .toString("hex")
    .toUpperCase()}`;
}

function normalizeClickwrapPreviewRecordId(value) {
  const normalized = String(value || "").trim();
  return /^[a-z0-9_-]{8,80}$/i.test(normalized) ? normalized : "";
}

function formatClickwrapAcceptanceTimestamp(value) {
  const iso = parseIsoTimestamp(value);
  if (!iso) return "";
  return iso.replace("T", " ").replace(/\.\d{3}Z$/, " UTC");
}

function compactClickwrapPreviewAddress(value) {
  const segments = String(value || "")
    .trim()
    .split(",")
    .map((segment) => String(segment || "").trim())
    .filter(Boolean);
  if (!segments.length) return "";
  let compact = segments.slice(0, 2).join(", ") || segments[0];
  compact = compact
    .replace(/\bAvenue\b/gi, "Ave.")
    .replace(/\bBoulevard\b/gi, "Blvd.")
    .replace(/\bStreet\b/gi, "St.")
    .replace(/\bRoad\b/gi, "Rd.")
    .replace(/\bSquare\b/gi, "Sq.");
  if (compact.length > 32 && segments[0]) {
    compact = segments[0];
  }
  return compact;
}

function getClickwrapPreviewFieldRectWidth(textField) {
  try {
    const widgets = textField?.acroField?.getWidgets?.() || [];
    const rect = widgets[0]?.getRectangle?.();
    const width = Number(rect?.width || 0);
    return Number.isFinite(width) ? width : 0;
  } catch (_error) {
    return 0;
  }
}

function getClickwrapPreviewFieldBaseFontSize(fieldName) {
  return fieldName === "merchant_legal_name_cover" ? 20 : 10;
}

function getClickwrapPreviewFieldMinFontSize(fieldName) {
  if (fieldName === "merchant_legal_name_cover") return 10;
  if (fieldName === "merchant_address" || fieldName === "merchant_email") return 8.5;
  return 8;
}

function resolveClickwrapPreviewFieldFontSize(textField, font, fieldName, value) {
  const text = String(value || "").trim();
  const baseFontSize = getClickwrapPreviewFieldBaseFontSize(fieldName);
  if (!text || !font) {
    return baseFontSize;
  }

  const fieldWidth = getClickwrapPreviewFieldRectWidth(textField);
  if (!fieldWidth) {
    return baseFontSize;
  }

  const minFontSize = getClickwrapPreviewFieldMinFontSize(fieldName);
  const horizontalPadding = fieldName === "merchant_legal_name_cover" ? 10 : 6;
  const usableWidth = Math.max(12, fieldWidth - horizontalPadding);
  let fontSize = baseFontSize;

  while (fontSize > minFontSize && font.widthOfTextAtSize(text, fontSize) > usableWidth) {
    fontSize -= 0.25;
  }

  return Number(fontSize.toFixed(2));
}

function buildClickwrapPreviewProfileSnapshot(email, profile) {
  const normalizedProfile = normalizeRegistrationProfile(profile);
  return {
    email: normalizeEmail(email),
    companyName: normalizedProfile.companyName,
    contactName: normalizedProfile.contactName,
    contactEmail: normalizedProfile.contactEmail,
    contactPhone: normalizedProfile.contactPhone,
    billingAddress: normalizedProfile.billingAddress,
    taxId: normalizedProfile.taxId,
    customerId: normalizedProfile.customerId,
  };
}

function clickwrapPreviewSnapshotsMatch(left, right) {
  return JSON.stringify(left || {}) === JSON.stringify(right || {});
}

function buildClickwrapPreviewFieldValues({ contract, email, profile, ipAddress, recordId, acceptedAt }) {
  const normalizedProfile = normalizeRegistrationProfile(profile);
  const acceptedEmail = normalizeEmail(email || normalizedProfile.contactEmail || "");
  return {
    merchant_legal_name: normalizedProfile.companyName,
    merchant_tax_id: normalizedProfile.taxId,
    merchant_address: compactClickwrapPreviewAddress(normalizedProfile.billingAddress),
    merchant_contact_name: normalizedProfile.contactName,
    merchant_email: acceptedEmail,
    merchant_ip_address: String(ipAddress || "").trim(),
    merchant_customer_id: normalizedProfile.customerId,
    record_id: String(recordId || "").trim(),
    acceptance_timestamp_utc: formatClickwrapAcceptanceTimestamp(acceptedAt),
    agreement_version: String(contract?.version || "").trim(),
    merchant_legal_name_cover: normalizedProfile.companyName,
  };
}

async function loadClickwrapPreviewAssets() {
  if (!clickwrapPreviewAssetsPromise) {
    clickwrapPreviewAssetsPromise = Promise.all([
      fs.promises.readFile(CLICKWRAP_CONTRACT_TEMPLATE_FILE),
      fs.promises.readFile(CLICKWRAP_CONTRACT_TEMPLATE_FONT_FILE),
    ])
      .then(([templateBytes, fontBytes]) => ({ templateBytes, fontBytes }))
      .catch((error) => {
        clickwrapPreviewAssetsPromise = null;
        throw error;
      });
  }
  return clickwrapPreviewAssetsPromise;
}

async function buildClickwrapPreviewPdf({ contract, email, profile, ipAddress, recordId, acceptedAt }) {
  const { templateBytes, fontBytes } = await loadClickwrapPreviewAssets();
  const pdfDocument = await PDFDocument.load(templateBytes);
  pdfDocument.registerFontkit(fontkit);
  const ptMonoFont = await pdfDocument.embedFont(fontBytes);
  const form = pdfDocument.getForm();
  const fieldValues = buildClickwrapPreviewFieldValues({
    contract,
    email,
    profile,
    ipAddress,
    recordId,
    acceptedAt,
  });
  let coverFallback = null;

  for (const [fieldName, fieldValue] of Object.entries(fieldValues)) {
    try {
      const textField = form.getTextField(fieldName);
      const fontSize = resolveClickwrapPreviewFieldFontSize(textField, ptMonoFont, fieldName, fieldValue);
      textField.setFontSize(fontSize);
      textField.setText(String(fieldValue || ""));
      if (fieldName === "merchant_legal_name_cover") {
        const widget = textField?.acroField?.getWidgets?.()?.[0] || null;
        const rect = widget?.getRectangle?.() || null;
        if (rect && String(fieldValue || "").trim()) {
          coverFallback = {
            text: String(fieldValue || "").trim(),
            rect,
            fontSize,
          };
        }
      }
    } catch (_error) {
      // Allow template iterations to omit fields without breaking preview generation.
    }
  }

  form.updateFieldAppearances(ptMonoFont);
  form.flatten();
  if (coverFallback) {
    const firstPage = pdfDocument.getPages()[0];
    if (firstPage) {
      firstPage.drawText(coverFallback.text, {
        x: coverFallback.rect.x + 5,
        y: coverFallback.rect.y + Math.max(2, (coverFallback.rect.height - coverFallback.fontSize) / 2),
        font: ptMonoFont,
        size: coverFallback.fontSize,
        color: rgb(1, 1, 1),
      });
    }
  }
  return Buffer.from(await pdfDocument.save());
}

function buildClickwrapPreviewDownloadName(contract, recordId) {
  const version = String(contract?.version || "agreement")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
  const safeRecordId = String(recordId || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]+/g, "");
  return `shipide-agreement-${version || "agreement"}-${safeRecordId || "preview"}.pdf`;
}

function buildClickwrapStoragePath(contract, recordId, filename) {
  const versionSegment = String(contract?.version || "agreement")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
  const recordSegment = String(recordId || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]+/g, "");
  return `accepted/${versionSegment || "agreement"}/${recordSegment || "record"}/${filename}`;
}

async function ensureClickwrapStorageBucket() {
  if (!clickwrapStorageBucketPromise) {
    clickwrapStorageBucketPromise = (async () => {
      const client = getSupabaseAdminClient();
      const { data, error } = await client.storage.getBucket(CLICKWRAP_STORAGE_BUCKET);
      if (!error && data) {
        return data;
      }
      const message = String(error?.message || "").toLowerCase();
      const missingBucket =
        !message
        || message.includes("not found")
        || message.includes("does not exist")
        || message.includes("not exists");
      if (!missingBucket && error) {
        throw new Error(error.message || "Could not access clickwrap storage bucket.");
      }
      const { data: createdBucket, error: createError } = await client.storage.createBucket(
        CLICKWRAP_STORAGE_BUCKET,
        {
          public: false,
          allowedMimeTypes: ["application/pdf"],
          fileSizeLimit: "10MB",
        }
      );
      if (createError) {
        const createMessage = String(createError?.message || "").toLowerCase();
        if (!createMessage.includes("already exists")) {
          throw new Error(createError.message || "Could not create clickwrap storage bucket.");
        }
      }
      return createdBucket || { id: CLICKWRAP_STORAGE_BUCKET, name: CLICKWRAP_STORAGE_BUCKET };
    })().catch((error) => {
      clickwrapStorageBucketPromise = null;
      throw error;
    });
  }
  return clickwrapStorageBucketPromise;
}

async function persistAcceptedClickwrapContract({ contract, recordId, pdfBytes, acceptedAt }) {
  const filename = buildClickwrapPreviewDownloadName(contract, recordId);
  const storagePath = buildClickwrapStoragePath(contract, recordId, filename);
  await ensureClickwrapStorageBucket();
  const client = getSupabaseAdminClient();
  const { data, error } = await client.storage
    .from(CLICKWRAP_STORAGE_BUCKET)
    .upload(storagePath, pdfBytes, {
      contentType: "application/pdf",
      cacheControl: "3600",
      upsert: true,
    });
  if (error) {
    throw new Error(error.message || "Could not store accepted agreement PDF.");
  }
  return {
    recordId: String(recordId || "").trim(),
    filename,
    bucket: CLICKWRAP_STORAGE_BUCKET,
    objectPath: data?.path || storagePath,
    fullPath: data?.fullPath || null,
    sha256: crypto.createHash("sha256").update(pdfBytes).digest("hex"),
    sizeBytes: Buffer.byteLength(pdfBytes),
    acceptedAt: parseIsoTimestamp(acceptedAt),
    storedAt: new Date().toISOString(),
  };
}

function normalizeInvoicePdfVariantStage(stage) {
  return String(Math.max(0, Number(stage) || 0));
}

function buildBillingInvoiceStoragePath(invoice, reminderStage, filename) {
  const referenceSegment = getBillingInvoiceReference(invoice)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
  const stageSegment = normalizeInvoicePdfVariantStage(reminderStage);
  return `pdf/${referenceSegment || "invoice"}/stage-${stageSegment}/${filename}`;
}

async function ensureBillingInvoiceStorageBucket() {
  if (!billingInvoiceStorageBucketPromise) {
    billingInvoiceStorageBucketPromise = (async () => {
      const client = getSupabaseAdminClient();
      const { data, error } = await client.storage.getBucket(BILLING_INVOICE_STORAGE_BUCKET);
      if (!error && data) {
        return data;
      }
      const message = String(error?.message || "").toLowerCase();
      const missingBucket =
        !message
        || message.includes("not found")
        || message.includes("does not exist")
        || message.includes("not exists");
      if (!missingBucket && error) {
        throw new Error(error.message || "Could not access billing invoice storage bucket.");
      }
      const { data: createdBucket, error: createError } = await client.storage.createBucket(
        BILLING_INVOICE_STORAGE_BUCKET,
        {
          public: false,
          allowedMimeTypes: ["application/pdf"],
          fileSizeLimit: "25MB",
        }
      );
      if (createError) {
        const createMessage = String(createError?.message || "").toLowerCase();
        if (!createMessage.includes("already exists")) {
          throw new Error(createError.message || "Could not create billing invoice storage bucket.");
        }
      }
      return createdBucket || { id: BILLING_INVOICE_STORAGE_BUCKET, name: BILLING_INVOICE_STORAGE_BUCKET };
    })().catch((error) => {
      billingInvoiceStorageBucketPromise = null;
      throw error;
    });
  }
  return billingInvoiceStorageBucketPromise;
}

async function persistBillingInvoicePdfVariant(invoice, reminderStage, pdfBytes, filename) {
  const safeFilename =
    String(filename || buildInvoiceVariantPdfFilename(invoice, reminderStage)).trim()
    || buildInvoiceVariantPdfFilename(invoice, reminderStage);
  const storagePath = buildBillingInvoiceStoragePath(invoice, reminderStage, safeFilename);
  await ensureBillingInvoiceStorageBucket();
  const client = getSupabaseAdminClient();
  const { data, error } = await client.storage
    .from(BILLING_INVOICE_STORAGE_BUCKET)
    .upload(storagePath, pdfBytes, {
      contentType: "application/pdf",
      cacheControl: "3600",
      upsert: true,
    });
  if (error) {
    throw new Error(error.message || "Could not store invoice PDF.");
  }
  return {
    stage: normalizeInvoicePdfVariantStage(reminderStage),
    rendererVersion: APPROVED_INVOICE_RENDERER_VERSION,
    filename: safeFilename,
    bucket: BILLING_INVOICE_STORAGE_BUCKET,
    objectPath: data?.path || storagePath,
    fullPath: data?.fullPath || null,
    sha256: crypto.createHash("sha256").update(pdfBytes).digest("hex"),
    sizeBytes: Buffer.byteLength(pdfBytes),
    storedAt: new Date().toISOString(),
  };
}

function getBillingInvoicePdfVariantMap(invoice) {
  const metadata = invoice?.metadata && typeof invoice.metadata === "object" ? invoice.metadata : {};
  const variants =
    metadata?.invoice_pdf_variants && typeof metadata.invoice_pdf_variants === "object"
      ? metadata.invoice_pdf_variants
      : {};
  return variants;
}

function getBillingInvoicePdfVariantRecord(invoice, reminderStage, options = {}) {
  const variants = getBillingInvoicePdfVariantMap(invoice);
  const exactKey = normalizeInvoicePdfVariantStage(reminderStage);
  const exact = variants[exactKey] && typeof variants[exactKey] === "object" ? variants[exactKey] : null;
  if (exact && exact.rendererVersion === APPROVED_INVOICE_RENDERER_VERSION) return exact;
  if (options?.allowFallback === true) {
    const initial = variants["0"] && typeof variants["0"] === "object" ? variants["0"] : null;
    if (initial && initial.rendererVersion === APPROVED_INVOICE_RENDERER_VERSION) return initial;
  }
  return null;
}

function mergeBillingInvoicePdfVariantMetadata(invoice, variants = []) {
  const metadata = invoice?.metadata && typeof invoice.metadata === "object" ? invoice.metadata : {};
  const nextVariants = {
    ...getBillingInvoicePdfVariantMap(invoice),
  };
  (Array.isArray(variants) ? variants : []).forEach((variant) => {
    const stageKey = normalizeInvoicePdfVariantStage(variant?.stage);
    nextVariants[stageKey] = {
      stage: stageKey,
      rendererVersion:
        String(variant?.rendererVersion || APPROVED_INVOICE_RENDERER_VERSION).trim()
        || APPROVED_INVOICE_RENDERER_VERSION,
      filename: String(variant?.filename || "").trim() || buildInvoiceVariantPdfFilename(invoice, stageKey),
      bucket: String(variant?.bucket || BILLING_INVOICE_STORAGE_BUCKET).trim() || BILLING_INVOICE_STORAGE_BUCKET,
      objectPath: String(variant?.objectPath || variant?.fullPath || "").trim() || null,
      fullPath: String(variant?.fullPath || variant?.objectPath || "").trim() || null,
      sha256: String(variant?.sha256 || "").trim().toLowerCase() || null,
      sizeBytes: Number(variant?.sizeBytes) || 0,
      storedAt: parseIsoTimestamp(variant?.storedAt) || new Date().toISOString(),
    };
  });
  return {
    ...metadata,
    invoice_pdf_variants: nextVariants,
  };
}

async function loadStoredBillingInvoicePdfVariant(invoice, reminderStage, options = {}) {
  const variant = getBillingInvoicePdfVariantRecord(invoice, reminderStage, options);
  if (!variant?.objectPath && !variant?.fullPath) return null;
  const objectPath = String(variant.objectPath || variant.fullPath || "").trim();
  if (!objectPath) return null;
  const client = getSupabaseAdminClient();
  const { data, error } = await client.storage
    .from(String(variant.bucket || BILLING_INVOICE_STORAGE_BUCKET).trim() || BILLING_INVOICE_STORAGE_BUCKET)
    .download(objectPath);
  if (error) {
    throw new Error(error.message || "Could not load stored invoice PDF.");
  }
  const bytes = Buffer.from(await data.arrayBuffer());
  return {
    bytes,
    filename:
      String(variant.filename || buildInvoiceVariantPdfFilename(invoice, reminderStage)).trim()
      || buildInvoiceVariantPdfFilename(invoice, reminderStage),
    stage: normalizeInvoicePdfVariantStage(variant?.stage || reminderStage),
    exact: normalizeInvoicePdfVariantStage(variant?.stage || reminderStage) === normalizeInvoicePdfVariantStage(reminderStage),
  };
}

async function handleInvoicePdfView(req, res, requestUrl) {
  const payload = parseInvoiceViewToken(requestUrl.searchParams.get("token"));
  if (!payload || payload.type !== "invoice_view") {
    send(res, 404, { "Content-Type": "text/plain; charset=utf-8" }, "Invoice not found.");
    return;
  }
  const invoiceId = String(payload.invoiceId || "").trim();
  const reminderStage = normalizeInvoicePdfVariantStage(payload.stage);
  if (!invoiceId) {
    send(res, 404, { "Content-Type": "text/plain; charset=utf-8" }, "Invoice not found.");
    return;
  }
  try {
    const invoice = await getBillingInvoiceById(invoiceId);
    if (!invoice?.id) {
      send(res, 404, { "Content-Type": "text/plain; charset=utf-8" }, "Invoice not found.");
      return;
    }
    const variant = await loadStoredBillingInvoicePdfVariant(invoice, reminderStage, {
      allowFallback: false,
    }).catch(() => null);
    if (!variant?.bytes) {
      send(res, 404, { "Content-Type": "text/plain; charset=utf-8" }, "Invoice not found.");
      return;
    }
    const filename = sanitizeDownloadFilename(
      variant.filename,
      buildInvoiceVariantPdfFilename(invoice, reminderStage)
    );
    send(
      res,
      200,
      {
        "Content-Type": "application/pdf",
        "Content-Disposition": `inline; filename="${filename}"`,
        "Cache-Control": "private, no-store, max-age=0",
        "X-Robots-Tag": "noindex, nofollow",
      },
      variant.bytes
    );
  } catch (_error) {
    send(res, 404, { "Content-Type": "text/plain; charset=utf-8" }, "Invoice not found.");
  }
}

function buildAcceptedAgreementEmailSubject(contract) {
  return "Welcome to Shipide!";
}

function buildAcceptedAgreementEmailHtml(options = {}) {
  const welcomeUrl =
    String(options?.welcomeUrl || WELCOME_PORTAL_URL || "").trim()
    || "https://portal.shipide.com/login";
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
          height: 86px !important;
          line-height: 86px !important;
        }
        .shipide-email-logo {
          width: 96px !important;
          margin-bottom: 24px !important;
        }
      }
    </style>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0;padding:0;background:#00060f;background-image:radial-gradient(196% 108% at 50% 120%, rgba(186, 138, 255, 0.98) 0%, rgba(119, 71, 227, 0.9) 16%, rgba(78, 42, 165, 0.58) 31%, rgba(35, 18, 92, 0.28) 45%, rgba(0, 6, 15, 0) 64%),radial-gradient(78% 54% at 10% 100%, rgba(119, 71, 227, 0.24) 0%, rgba(0, 6, 15, 0) 76%),radial-gradient(78% 54% at 90% 100%, rgba(119, 71, 227, 0.24) 0%, rgba(0, 6, 15, 0) 76%),linear-gradient(180deg, rgba(0, 6, 15, 0) 0%, rgba(0, 6, 15, 0) 50%, rgba(20, 11, 54, 0.14) 68%, rgba(119, 71, 227, 0.14) 84%, rgba(186, 138, 255, 0.11) 100%);background-repeat:no-repeat;background-size:100% 100%;font-family:Helvetica,Arial,sans-serif;color:#f3f6ff;border-top:4px solid #7747e3;">
      <tr>
        <td align="center" style="padding:0 16px 24px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:980px;background:transparent;">
            <tr>
              <td align="center" class="shipide-email-hero-wrap" style="padding:70px 8px 24px;">
                <img src="https://portal.shipide.com/shipide_logo.png" alt="Shipide" class="shipide-email-logo" width="124" style="display:block;width:124px;height:auto;margin:0 auto 30px;" />
                <div class="shipide-email-hero-title" style="font-size:58px;line-height:1.03;letter-spacing:-0.035em;color:#f3f6ff;font-weight:300;">
                  Welcome to <span style="color:#ffffff;">Shipide!</span>
                </div>
                <div class="shipide-email-hero-sub" style="max-width:700px;margin:18px auto 0;font-size:16px;line-height:1.5;color:#9aa3b2;">
                  Your account is now active. The PDF copy of your accepted<br/>service agreement is attached to this email.
                </div>
                <a href="${escapeHtml(welcomeUrl)}" style="display:inline-block;margin-top:24px;padding:12px 20px;border-radius:4px;border:1px solid rgb(46,46,46);background:#1c2026;color:#f3f6ff;text-decoration:none;font-size:14px;line-height:1;">
                  Head to Shipide
                </a>
              </td>
            </tr>
            <tr>
              <td class="shipide-email-hero-spacer" style="height:122px;font-size:0;line-height:122px;">&nbsp;</td>
            </tr>
          </table>
        </td>
      </tr>
    </table>`;
}

function buildAcceptedAgreementEmailText(options = {}) {
  const welcomeUrl =
    String(options?.welcomeUrl || WELCOME_PORTAL_URL || "").trim()
    || "https://portal.shipide.com/login";
  return [
    "Welcome to Shipide!",
    "",
    "Your account is now active. The PDF copy of your accepted Shipide service agreement is attached to this email.",
    "",
    `Head to Shipide: ${welcomeUrl}`,
  ].join("\n");
}

async function sendAcceptedAgreementEmail({ email, contract, artifact, pdfBytes, recipientName }) {
  return sendResendEmail({
    to: normalizeEmail(email),
    fromName: WELCOME_FROM_NAME,
    fromEmail: WELCOME_FROM_EMAIL,
    subject: buildAcceptedAgreementEmailSubject(contract),
    idempotencyKey: `agreement-${String(artifact?.recordId || "").trim()}`,
    html: buildAcceptedAgreementEmailHtml({
      contract,
      artifact,
      recipientName,
    }),
    text: buildAcceptedAgreementEmailText({
      contract,
      artifact,
      recipientName,
    }),
    attachments: [
      {
        filename: artifact.filename,
        content: pdfBytes,
        contentType: "application/pdf",
      },
    ],
  });
}

function buildAcceptedAgreementPreviewDocument(subject, emailHtml) {
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${escapeHtml(subject)}</title>
    <style>
      body {
        margin: 0;
        background: #050913;
      }
    </style>
  </head>
  <body>
    ${emailHtml}
  </body>
</html>`;
}

function buildAcceptedAgreementTestProfile(toEmail) {
  return {
    companyName: "Atelier Meridian",
    contactName: "Claire Dupont",
    contactEmail: normalizeEmail(toEmail),
    contactPhone: "+32 2 555 01 29",
    billingAddress: "Avenue Louise 120, 1050 Brussels, Belgium",
    taxId: "BE0123456789",
    customerId: "SHIPIDE-TEST",
  };
}

async function buildAcceptedAgreementTestPayload(req, toEmail, options = {}) {
  const includePdf = options?.includePdf === true;
  const contract = await getActiveClickwrapContract();
  const acceptedAt = new Date().toISOString();
  const recordId = createClickwrapPreviewRecordId();
  const profile = buildAcceptedAgreementTestProfile(toEmail);
  const artifact = {
    recordId,
    filename: buildClickwrapPreviewDownloadName(contract, recordId),
    acceptedAt,
  };
  let pdfBytes = null;
  if (includePdf) {
    pdfBytes = await buildClickwrapPreviewPdf({
      contract,
      email: toEmail,
      profile,
      ipAddress: getRequestIpAddress(req),
      recordId,
      acceptedAt,
    });
  }
  return {
    contract,
    artifact,
    pdfBytes,
    recipientName: profile.contactName,
  };
}

function buildClickwrapPreviewPublicContract(contract, previewEntry) {
  const publicContract = mapClickwrapContractToPublic(contract);
  return {
    ...publicContract,
    pdfUrl: `/api/auth/register-contract-preview-file?recordId=${encodeURIComponent(
      previewEntry.recordId
    )}`,
    downloadName: previewEntry.downloadName,
    recordId: previewEntry.recordId,
    previewPdfHash: previewEntry.pdfSha256,
  };
}

function getClickwrapPreviewRecord(recordId) {
  pruneClickwrapPreviewStore();
  const normalized = normalizeClickwrapPreviewRecordId(recordId);
  if (!normalized) return null;
  const entry = clickwrapPreviewStore.get(normalized);
  if (!entry) return null;
  if (Date.now() >= Number(entry.expiresAt || 0)) {
    clickwrapPreviewStore.delete(normalized);
    return null;
  }
  return entry;
}

async function getActiveClickwrapContract() {
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,version,title,body_text,hash_sha256,pdf_url,effective_at,is_active,created_at"
  );
  params.set("is_active", "eq.true");
  params.set("order", "effective_at.desc.nullslast,created_at.desc");
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${CLICKWRAP_CONTRACTS_TABLE}?${params.toString()}`,
    {
      method: "GET",
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (isMissingClickwrapSchemaError(details)) {
      return normalizeClickwrapContractSnapshot(getEmbeddedClickwrapContract());
    }
    throw new Error(`Could not load click-wrap agreement (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  const contract = Array.isArray(rows) && rows.length ? rows[0] : getEmbeddedClickwrapContract();
  return normalizeClickwrapContractSnapshot({
    ...contract,
    source: contract?.source || "supabase",
  });
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
  const agreement = rawAgreement && typeof rawAgreement === "object" ? rawAgreement : {};
  return {
    contractId: String(agreement?.contractId || "").trim() || null,
    contractVersion: String(agreement?.contractVersion || "").trim(),
    contractHash: String(agreement?.contractHash || "").trim().toLowerCase(),
    recordId:
      normalizeClickwrapPreviewRecordId(agreement?.recordId || agreement?.previewRecordId) || null,
    scrolledToEnd: agreement?.scrolledToEnd === true,
    scrolledToEndAt: parseIsoTimestamp(agreement?.scrolledToEndAt),
    agreed: agreement?.agreed === true,
    agreedAt: parseIsoTimestamp(agreement?.agreedAt),
    maxScrollProgress: parseNumeric(agreement?.maxScrollProgress),
    clientTimezone: String(agreement?.clientTimezone || "").trim(),
    clientLocale: String(agreement?.clientLocale || "").trim(),
  };
}

function validateClickwrapAgreement(agreement, contract, options = {}) {
  const { previewRecord = null, email = "", profile = null } = options;
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
  if (agreement?.recordId) {
    if (!previewRecord) {
      return "Agreement preview expired. Review the agreement again.";
    }
    if (previewRecord.contractVersion !== String(contract?.version || "")) {
      return "Agreement preview version mismatch. Review the agreement again.";
    }
    if (previewRecord.contractHash !== String(contract?.hash_sha256 || "")) {
      return "Agreement preview integrity check failed. Review the agreement again.";
    }
    const incomingSnapshot = buildClickwrapPreviewProfileSnapshot(email, profile);
    if (!clickwrapPreviewSnapshotsMatch(previewRecord.profileSnapshot, incomingSnapshot)) {
      return "Agreement preview no longer matches your registration details. Review again.";
    }
  }
  return "";
}

function getRequestIpAddress(req) {
  const headers = req?.headers || {};
  const candidates = [
    headers["cf-connecting-ip"],
    headers["x-forwarded-for"],
    headers["x-real-ip"],
    headers["x-client-ip"],
  ];
  for (const candidate of candidates) {
    const value = String(candidate || "")
      .split(",")[0]
      .trim();
    if (value) return value;
  }
  return "";
}

function getRemoteRequestAddress(req) {
  const forwarded = getRequestIpAddress(req);
  if (forwarded) return forwarded;
  return String(req?.socket?.remoteAddress || "").trim();
}

function isLocalPreviewRequest(req) {
  const normalized = String(getRemoteRequestAddress(req) || "")
    .trim()
    .toLowerCase()
    .replace(/^\[|\]$/g, "")
    .split("%")[0];
  if (!normalized) return false;
  if (normalized === "::1" || normalized === "::ffff:127.0.0.1" || normalized === "127.0.0.1") {
    return true;
  }
  return false;
}

function buildClickwrapEvidence(req, invite, contract, agreement, createdUserId, email, options = {}) {
  const { previewRecord = null, acceptedContract = null } = options;
  const acceptedAt = agreement?.agreedAt || new Date().toISOString();
  const payload = {
    schema: "shipide.clickwrap.v1",
    accepted_at: acceptedAt,
    contract_version: String(contract?.version || ""),
    contract_hash_sha256: String(contract?.hash_sha256 || ""),
    contract_id: contract?.id || null,
    preview_record_id: agreement?.recordId || null,
    preview_pdf_sha256: String(previewRecord?.pdfSha256 || "").trim(),
    accepted_contract_pdf: acceptedContract
      ? {
          record_id: acceptedContract.recordId,
          filename: acceptedContract.filename,
          sha256: acceptedContract.sha256,
          size_bytes: acceptedContract.sizeBytes,
          storage_bucket: acceptedContract.bucket,
          storage_object_path: acceptedContract.objectPath,
          storage_full_path: acceptedContract.fullPath,
          stored_at: acceptedContract.storedAt,
        }
      : null,
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
    ip_address: getRequestIpAddress(req),
    user_agent: String(req?.headers?.["user-agent"] || "").trim(),
    captured_at: new Date().toISOString(),
  };
  const payloadCanonical = JSON.stringify(payload);
  const digest = sha256Hex(payloadCanonical);
  const secret = getClickwrapProofSecret();
  const signature = hmacSha256Hex(secret, payloadCanonical);
  return {
    payload,
    digest,
    signature,
  };
}

async function insertClickwrapAcceptance(record) {
  const response = await supabaseServiceRequest(`/rest/v1/${CLICKWRAP_ACCEPTANCES_TABLE}`, {
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

function hasServerManagedInviteAdminAccess(user) {
  return user?.app_metadata?.app_admin === true;
}

function hasInviteAdminEmailAllowlistAccess(user) {
  if (!INVITE_ADMIN_EMAILS.length) return false;
  const email = normalizeEmail(user?.email || "");
  return Boolean(email && INVITE_ADMIN_EMAILS.includes(email));
}

function canManageRegistrationInvites(user) {
  if (!user || !user.id) return false;
  // Only trust admin roles from server-managed app metadata or the explicit
  // email allowlist. Client-writable user_metadata must never grant admin access.
  return hasServerManagedInviteAdminAccess(user) || hasInviteAdminEmailAllowlistAccess(user);
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

function getInvoiceTermsDays() {
  const value = Number(process.env.INVOICE_TERMS_DAYS);
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

const PUBLIC_DOCUMENT_CODE_ALPHABET = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ";
const PUBLIC_DOCUMENT_CODE_LENGTH = 6;

function getBillingDocumentPublicCodePrefix(invoiceKind = "monthly", documentType = "invoice") {
  const normalizedKind = String(invoiceKind || "").trim().toLowerCase();
  const normalizedType = String(documentType || "").trim().toLowerCase();
  if (normalizedType === "receipt") return "R-REC";
  if (normalizedKind === "topup") return "T-INV";
  return "M-INV";
}

function generatePublicDocumentCodeSuffix(length = PUBLIC_DOCUMENT_CODE_LENGTH) {
  const safeLength = Math.max(4, Math.min(12, Number(length) || PUBLIC_DOCUMENT_CODE_LENGTH));
  const bytes = crypto.randomBytes(safeLength);
  let result = "";
  for (let index = 0; index < bytes.length; index += 1) {
    result += PUBLIC_DOCUMENT_CODE_ALPHABET[bytes[index] % PUBLIC_DOCUMENT_CODE_ALPHABET.length];
  }
  return result;
}

function buildBillingDocumentPublicCode(invoiceKind = "monthly", documentType = "invoice") {
  return `${getBillingDocumentPublicCodePrefix(invoiceKind, documentType)}-${generatePublicDocumentCodeSuffix()}`;
}

function buildDocumentPdfFilenameFromReference(reference, fallback = "document") {
  const safe = String(reference || "")
    .trim()
    .replace(/[^A-Za-z0-9-]+/g, "-")
    .replace(/^-+|-+$/g, "");
  const stem = safe || String(fallback || "document").trim() || "document";
  return `${stem}.pdf`;
}

function isPublicDocumentCodeCollisionError(error) {
  const message = String(error?.message || "");
  return /(23505|duplicate key)/i.test(message) && /invoice_number|public_code|document code/i.test(message);
}

function getBillingInvoiceKind(invoice = {}) {
  const normalized = String(
    invoice?.invoice_kind
      || invoice?.metadata?.invoice_kind
      || "monthly"
  )
    .trim()
    .toLowerCase();
  return normalized === "topup" ? "topup" : "monthly";
}

function isTopupBillingInvoice(invoice = {}) {
  return getBillingInvoiceKind(invoice) === "topup";
}

function getBillingInvoiceReference(invoice = {}) {
  const customReference = String(
    invoice?.reference
      || invoice?.invoice_number
      || invoice?.metadata?.invoice_number
      || ""
  ).trim();
  return customReference || toInvoiceReference(invoice?.id);
}

function buildTopupInvoiceToken(topup = {}) {
  const rawId = String(topup?.id || "").trim();
  if (isUuid(rawId)) {
    return rawId.replace(/-/g, "").slice(0, 8).toUpperCase();
  }
  const normalizedId = rawId.toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (normalizedId) {
    return normalizedId.slice(0, 8).padEnd(8, "0");
  }
  const normalizedReference = String(topup?.reference || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
  if (normalizedReference) {
    return normalizedReference.slice(-8).padStart(8, "0");
  }
  return "TOPUP000";
}

function buildTopupInvoiceNumber(topup = {}) {
  return buildBillingDocumentPublicCode("topup", "invoice");
}

function getCanonicalInvoiceTimestamp(invoice = {}) {
  const candidates = [
    invoice?.created_at,
    invoice?.issued_at,
    invoice?.sent_at,
    invoice?.paid_at,
    invoice?.updated_at,
  ];
  for (const candidate of candidates) {
    const parsed = Date.parse(String(candidate || "").trim());
    if (Number.isFinite(parsed)) return parsed;
  }
  return Number.MAX_SAFE_INTEGER;
}

function compareCanonicalBillingInvoices(left = {}, right = {}) {
  const leftIssued = Boolean(String(left?.sent_at || left?.email_message_id || "").trim());
  const rightIssued = Boolean(String(right?.sent_at || right?.email_message_id || "").trim());
  if (leftIssued !== rightIssued) {
    return leftIssued ? -1 : 1;
  }
  const leftTime = getCanonicalInvoiceTimestamp(left);
  const rightTime = getCanonicalInvoiceTimestamp(right);
  if (leftTime !== rightTime) {
    return leftTime - rightTime;
  }
  return String(left?.id || "").localeCompare(String(right?.id || ""));
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
  const reference = getBillingInvoiceReference(invoice);
  const isTopup = isTopupBillingInvoice(invoice);
  const monthYear = formatInvoiceSubjectMonthYear(invoice);
  const suffix = `${monthYear} [${reference}]`;
  const isReminder = Boolean(options?.isReminder);
  const reminderStage = Number(options?.reminderStage) || 0;
  if (isTopup && !isReminder) {
    return `Your Invoice — Top-up [${reference}]`;
  }
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

function normalizeInvoiceStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["draft", "sent", "overdue", "paid", "cancelled"].includes(normalized)) {
    return normalized;
  }
  return "draft";
}

function normalizeInvoiceRunMode(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["preview", "create", "send"].includes(normalized)) return normalized;
  return "create";
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

function formatInvoiceTrackingLabel(invoice) {
  if (!invoice) return "No invoices yet";
  if (isTopupBillingInvoice(invoice)) {
    const topupStatus = String(invoice.status || "").trim().toLowerCase();
    if (topupStatus === "paid") return "Credited";
    if (topupStatus === "sent") return "Sent";
  }
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

function getReminderTitleByStage(stage, dueAt) {
  const dueLabel = String(dueAt || "").trim()
    ? new Date(dueAt).toISOString().slice(0, 10)
    : "the due date";
  const mapping = {
    1: `Invoice due in 15 days (${dueLabel})`,
    2: `Invoice due in 7 days (${dueLabel})`,
    3: `Invoice due tomorrow (${dueLabel})`,
    4: `Invoice due today (${dueLabel})`,
    5: "Invoice is 3 days overdue",
    6: "Invoice is 10 days overdue",
    7: "Invoice is 15 days overdue",
    8: "Invoice is 30 days overdue",
  };
  return mapping[Number(stage)] || "Invoice payment reminder";
}

function getInvoiceRecipientFromSnapshot(invoice, fallbackEmail = "") {
  const recipient = String(invoice?.contact_email || "").trim().toLowerCase();
  if (recipient) return recipient;
  return String(fallbackEmail || "").trim().toLowerCase();
}

function assertResendConfig() {
  if (!RESEND_API_KEY) {
    throw new Error("RESEND_API_KEY is required.");
  }
}

async function sendResendEmail(payload) {
  assertResendConfig();
  const fromName = String(payload?.fromName || RESEND_FROM_NAME || "Shipide Billing").trim();
  const fromEmail = String(payload?.fromEmail || RESEND_FROM_EMAIL || "").trim();
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
  const replyTo = String(payload?.replyTo || RESEND_REPLY_TO || "").trim();
  if (replyTo) {
    requestBody.reply_to = replyTo;
  }
  const attachments = Array.isArray(payload?.attachments)
    ? payload.attachments
        .map((attachment) => {
          const filename = String(attachment?.filename || "").trim();
          if (!filename) return null;
          const content =
            typeof attachment?.content === "string"
              ? attachment.content
              : Buffer.isBuffer(attachment?.content)
                ? attachment.content.toString("base64")
                : "";
          if (!content) return null;
          const normalized = {
            filename,
            content,
          };
          const pathValue = String(attachment?.path || "").trim();
          if (pathValue) {
            normalized.path = pathValue;
          }
          const contentType = String(
            attachment?.contentType || attachment?.content_type || ""
          ).trim();
          if (contentType) {
            normalized.content_type = contentType;
          }
          const disposition = String(attachment?.disposition || "").trim();
          if (disposition) {
            normalized.disposition = disposition;
          }
          return normalized;
        })
        .filter(Boolean)
    : [];
  if (attachments.length) {
    requestBody.attachments = attachments;
  }
  const requestHeaders = {
    Authorization: `Bearer ${RESEND_API_KEY}`,
    "Content-Type": "application/json",
  };
  const idempotencyKey = String(payload?.idempotencyKey || "").trim();
  if (idempotencyKey) {
    requestHeaders["Idempotency-Key"] = idempotencyKey;
  }
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: requestHeaders,
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

function sanitizeSubmissionFilename(value, fallback = "shipide-cleaned-shipping-data.csv") {
  const sanitized = String(value || "")
    .replace(/[^A-Za-z0-9._-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 96);
  return sanitized || fallback;
}

function normalizeClientCompanyName(value) {
  return String(value || "").trim().replace(/\s+/g, " ").slice(0, 120);
}

function getShipmentExtractCompanyName(row) {
  return normalizeClientCompanyName(
    row?.client_company_name || row?.clientCompanyName || row?.metadata?.client_company_name || ""
  );
}

function escapeCsvCell(value) {
  const text = String(value ?? "");
  if (/[",\n\r;]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function buildCsvFromMatrix(headers = [], rows = []) {
  return [headers, ...rows]
    .map((row) => (Array.isArray(row) ? row : []).map(escapeCsvCell).join(","))
    .join("\n");
}

async function handleShippingDataCleanerSubmission(req, res) {
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const headers = Array.isArray(body?.headers)
    ? body.headers.map((header) => String(header || "").trim()).filter(Boolean)
    : [];
  const rows = Array.isArray(body?.rows)
    ? body.rows
        .filter((row) => Array.isArray(row))
        .slice(0, 5000)
        .map((row) => row.slice(0, headers.length).map((cell) => String(cell || "").trim()))
    : [];
  if (!headers.length || !rows.length) {
    sendJson(res, 400, { error: "Cleaned shipment data is required." });
    return;
  }
  if (headers.length > 32) {
    sendJson(res, 400, { error: "Too many cleaned columns." });
    return;
  }

  const token = normalizeInviteToken(body?.token || "");
  if (!token) {
    sendJson(res, 400, { error: "Shipment extract link token is required." });
    return;
  }

  let extractRequest = null;
  try {
    extractRequest = await getShipmentExtractRequestByToken(token);
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not validate shipment extract link." });
    return;
  }
  if (!extractRequest?.id) {
    sendJson(res, 404, { error: "Shipment extract link not found." });
    return;
  }
  if (shipmentExtractRequestIsExpired(extractRequest)) {
    sendJson(res, 410, { error: "Shipment extract link has expired." });
    return;
  }

  const sourceFilename = sanitizeSubmissionFilename(body?.fileName || "shipping-export.csv");
  const contactEmail = normalizeEmail(extractRequest?.client_email || "");
  const companyName = getShipmentExtractCompanyName(extractRequest);
  const outputFilename = sanitizeSubmissionFilename(
    `${companyName || "client"}_${contactEmail || "no-email"}_${sourceFilename.replace(/\.[^.]+$/, "")}-shipide-cleaned.csv`
  );
  const removedColumns = Array.isArray(body?.removedColumns)
    ? body.removedColumns.map((column) => String(column || "").trim()).filter(Boolean).slice(0, 80)
    : [];
  const originAddresses = Array.isArray(body?.originAddresses)
    ? body.originAddresses.map((address) => String(address || "").trim()).filter(Boolean).slice(0, 12)
    : [];
  const csv = buildCsvFromMatrix(headers, rows);
  const submittedAt = new Date().toISOString();
  const html = `
    <p>A sanitized shipping data extract was submitted from the public cleaner.</p>
    <ul>
      <li><strong>Source file:</strong> ${escapeHtml(sourceFilename)}</li>
      <li><strong>Company:</strong> ${companyName ? escapeHtml(companyName) : "Not assigned"}</li>
      <li><strong>Client email:</strong> ${contactEmail ? escapeHtml(contactEmail) : "Not assigned"}</li>
      <li><strong>Request id:</strong> ${escapeHtml(extractRequest.id)}</li>
      <li><strong>Rows:</strong> ${rows.length}</li>
      <li><strong>Columns kept:</strong> ${headers.map(escapeHtml).join(", ")}</li>
      <li><strong>Columns removed:</strong> ${removedColumns.length ? removedColumns.map(escapeHtml).join(", ") : "None reported"}</li>
      <li><strong>Manual ship-from addresses:</strong> ${originAddresses.length ? originAddresses.map(escapeHtml).join(" | ") : "None provided"}</li>
      <li><strong>Submitted at:</strong> ${escapeHtml(submittedAt)}</li>
    </ul>
  `;
  const text = [
    "A sanitized shipping data extract was submitted from the public cleaner.",
    `Source file: ${sourceFilename}`,
    `Company: ${companyName || "Not assigned"}`,
    `Client email: ${contactEmail || "Not assigned"}`,
    `Request id: ${extractRequest.id}`,
    `Rows: ${rows.length}`,
    `Columns kept: ${headers.join(", ")}`,
    `Columns removed: ${removedColumns.length ? removedColumns.join(", ") : "None reported"}`,
    `Manual ship-from addresses: ${originAddresses.length ? originAddresses.join(" | ") : "None provided"}`,
    `Submitted at: ${submittedAt}`,
  ].join("\n");

  try {
    const response = await sendResendEmail({
      to: DATA_CLEANER_SUBMISSION_EMAIL,
      fromName: "Shipide Data Cleaner",
      fromEmail: REPORTS_FROM_EMAIL || RESEND_FROM_EMAIL,
      replyTo: contactEmail || RESEND_REPLY_TO,
      subject: `Cleaned shipping data: ${companyName || contactEmail || sourceFilename}`,
      html,
      text,
      attachments: [
        {
          filename: outputFilename,
          content: Buffer.from(csv, "utf8"),
          contentType: "text/csv",
        },
      ],
    });
    await updateShipmentExtractRequestSubmitted(extractRequest.id, {
      submittedAt,
      sourceFilename,
      rows: rows.length,
      columns: headers.length,
      keptColumns: headers,
      removedColumns,
      metadata: {
        resend_email_id: response?.id || null,
        output_filename: outputFilename,
        client_company_name: companyName || null,
        manual_origin_addresses: originAddresses,
      },
    });
    sendJson(res, 200, {
      ok: true,
      id: response?.id || null,
      requestId: extractRequest.id,
      clientEmail: contactEmail || null,
      rows: rows.length,
      columns: headers.length,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not submit cleaned data." });
  }
}

async function handleDocumentsPreviewSendTestEmail(req, res) {
  if (!isLocalPreviewRequest(req)) {
    sendJson(res, 403, { error: "Documents preview email sending is only available from localhost." });
    return;
  }
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to send document previews." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const toEmail = normalizeEmail(body?.toEmail || "louis@gleth.com");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    sendJson(res, 400, { error: "A valid destination email is required." });
    return;
  }
  const previewDocuments = Array.isArray(body?.previewDocuments)
    ? body.previewDocuments
        .map((document) => {
          const kind = String(document?.kind || "").trim().toLowerCase();
          const filename = String(document?.filename || "").trim();
          const viewModel = document?.viewModel && typeof document.viewModel === "object"
            ? document.viewModel
            : null;
          if (!filename || !viewModel || !["receipt", "invoice"].includes(kind)) {
            return null;
          }
          return {
            kind,
            filename,
            viewModel,
          };
        })
        .filter(Boolean)
    : [];
  const attachments = Array.isArray(body?.attachments)
    ? body.attachments
        .map((attachment) => {
          const filename = String(attachment?.filename || "").trim();
          const content = String(attachment?.content || "").trim();
          if (!filename || !content) return null;
          const contentType = String(
            attachment?.contentType || attachment?.content_type || "application/pdf"
          ).trim() || "application/pdf";
          return {
            filename,
            content,
            contentType,
          };
        })
        .filter(Boolean)
    : [];
  if (previewDocuments.length && previewDocuments.length !== 3) {
    sendJson(res, 400, { error: "Exactly three preview documents are required." });
    return;
  }
  if (!RESEND_API_KEY || previewDocuments.length) {
    const bearerToken = getBearerToken(req);
    if (!bearerToken) {
      sendJson(res, 401, { error: "Authentication required." });
      return;
    }
    try {
      const response = await fetch(`${getPublicAppUrl()}/api/documents-preview/send-test-email`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${bearerToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          toEmail,
          attachments,
          previewDocuments,
        }),
      });
      const payload = await response.json().catch(() => ({}));
      sendJson(
        res,
        response.status,
        payload && typeof payload === "object"
          ? payload
          : { error: "Could not send preview PDFs through production mailer." }
      );
      return;
    } catch (error) {
      sendJson(
        res,
        502,
        { error: error?.message || "Could not reach the production preview mailer." }
      );
      return;
    }
  }
  if (attachments.length !== 3) {
    sendJson(res, 400, { error: "Exactly three PDF attachments are required." });
    return;
  }
  try {
    const resendResponse = await sendResendEmail({
      to: toEmail,
      subject: "Shipide document preview PDFs",
      html: `
        <div style="font-family:Arial,sans-serif;color:#111827;line-height:1.55">
          <p>Attached are the three local document preview PDFs generated from the localhost preview page.</p>
          <p>They include long sample rows so pagination can be reviewed across multiple pages.</p>
        </div>
      `,
      text: [
        "Attached are the three local document preview PDFs generated from the localhost preview page.",
        "They include long sample rows so pagination can be reviewed across multiple pages.",
      ].join("\n\n"),
      attachments,
    });
    sendJson(res, 200, {
      ok: true,
      to: toEmail,
      resendId: resendResponse?.id || null,
      attachments: attachments.map((attachment) => attachment.filename),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not send preview PDFs." });
  }
}

async function proxyAuthenticatedPdfRender(req, res, targetPath) {
  const bearerToken = getBearerToken(req);
  if (!bearerToken) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  let rawBody = "";
  try {
    rawBody = await readTextBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error?.message || "Invalid request body." });
    return;
  }
  try {
    const response = await fetch(`${getPublicAppUrl(req)}${targetPath}`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${bearerToken}`,
        "Content-Type": "application/json",
      },
      body: rawBody || "{}",
    });
    const contentType = String(response.headers.get("content-type") || "").trim();
    if (!response.ok || !/application\/pdf/i.test(contentType)) {
      const payload = await response.json().catch(() => ({}));
      sendJson(
        res,
        response.status,
        payload && typeof payload === "object"
          ? payload
          : { error: "Could not render PDF through production renderer." }
      );
      return;
    }
    const arrayBuffer = await response.arrayBuffer();
    send(
      res,
      200,
      {
        "Content-Type": contentType || "application/pdf",
        "Content-Disposition":
          String(response.headers.get("content-disposition") || "").trim()
          || 'inline; filename="document.pdf"',
        "Cache-Control": "no-store",
      },
      Buffer.from(arrayBuffer)
    );
  } catch (error) {
    sendJson(res, 502, {
      error: error?.message || "Could not reach the production PDF renderer.",
    });
  }
}

async function proxyAuthenticatedApi(req, res, targetPathWithQuery) {
  const bearerToken = getBearerToken(req);
  if (!bearerToken) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  let rawBody = "";
  if (!["GET", "HEAD"].includes(String(req.method || "").toUpperCase())) {
    try {
      rawBody = await readTextBody(req);
    } catch (error) {
      sendJson(res, 400, { error: error?.message || "Invalid request body." });
      return;
    }
  }
  try {
    const response = await fetch(`${getPublicAppUrl(req)}${targetPathWithQuery}`, {
      method: req.method || "GET",
      headers: {
        Authorization: `Bearer ${bearerToken}`,
        ...(rawBody
          ? {
              "Content-Type": "application/json",
            }
          : {}),
      },
      ...(rawBody
        ? {
            body: rawBody,
          }
        : {}),
    });
    const contentType = String(response.headers.get("content-type") || "").trim().toLowerCase();
    if (contentType.includes("application/json")) {
      const payload = await response.json().catch(() => ({}));
      sendJson(
        res,
        response.status,
        payload && typeof payload === "object"
          ? payload
          : { error: "Could not proxy API response." }
      );
      return;
    }
    const arrayBuffer = await response.arrayBuffer();
    send(
      res,
      response.status,
      {
        "Content-Type": contentType || "application/octet-stream",
        "Content-Disposition":
          String(response.headers.get("content-disposition") || "").trim()
          || 'inline; filename="document.pdf"',
        "Cache-Control":
          String(response.headers.get("cache-control") || "").trim()
          || "no-store",
      },
      Buffer.from(arrayBuffer)
    );
  } catch (error) {
    sendJson(res, 502, {
      error: error?.message || "Could not reach the production API.",
    });
  }
}

function normalizeClientBillingPreference(value) {
  const raw = value && typeof value === "object" ? value : {};
  const invoiceEnabled = raw.invoice_enabled === true;
  return {
    invoice_enabled: invoiceEnabled,
    card_enabled: false,
    updated_at: raw.updated_at || null,
    updated_by: raw.updated_by || null,
  };
}

function isDefaultClientBillingPreference(value) {
  const normalized = normalizeClientBillingPreference(value);
  return normalized.invoice_enabled !== true;
}

function getClientPaymentMode(pref) {
  const normalized = normalizeClientBillingPreference(pref);
  return normalized.invoice_enabled ? "invoice" : "wallet";
}

function buildBillingMetricsByUser(settings, invoices = [], invoiceItems = [], walletTransactions = []) {
  const carrierDiscountPct = normalizePercent(settings?.carrier_discount_pct, 0);
  const out = new Map();
  const ensureEntry = (userId) => {
    const safeUserId = String(userId || "").trim();
    if (!safeUserId) return null;
    if (!out.has(safeUserId)) {
      out.set(safeUserId, {
        total_revenue_ex_vat: 0,
        estimated_carrier_cost_ex_vat: 0,
        total_profit_ex_vat: 0,
      });
    }
    return out.get(safeUserId);
  };
  const invoiceItemsByInvoiceId = new Map();
  (Array.isArray(invoiceItems) ? invoiceItems : []).forEach((item) => {
    const invoiceId = String(item?.invoice_id || "").trim();
    if (!invoiceId) return;
    if (!invoiceItemsByInvoiceId.has(invoiceId)) invoiceItemsByInvoiceId.set(invoiceId, []);
    invoiceItemsByInvoiceId.get(invoiceId).push(item);
  });

  (Array.isArray(invoices) ? invoices : []).forEach((invoice) => {
    if (isTopupBillingInvoice(invoice)) return;
    const userId = String(invoice?.user_id || "").trim();
    const entry = ensureEntry(userId);
    if (!entry) return;
    entry.total_revenue_ex_vat += Math.max(0, toCents(invoice?.subtotal_ex_vat));
    const items = invoiceItemsByInvoiceId.get(String(invoice?.id || "").trim()) || [];
    items.forEach((item) => {
      const estimatedRetail = getEstimatedRetailTotal(
        item?.service_type,
        item?.quantity,
        Number(item?.amount_ex_vat) || 0
      );
      entry.estimated_carrier_cost_ex_vat += toCents(
        estimatedRetail * (1 - carrierDiscountPct / 100)
      );
    });
  });

  (Array.isArray(walletTransactions) ? walletTransactions : []).forEach((row) => {
    if (String(row?.source || "").trim().toLowerCase() !== "label_checkout") return;
    const amountCents = Number(row?.amount_cents) || 0;
    if (amountCents >= 0) return;
    const userId = String(row?.user_id || "").trim();
    const entry = ensureEntry(userId);
    if (!entry) return;
    const revenueExVat = fromCents(Math.abs(amountCents));
    const checkoutMetadata =
      row?.metadata?.checkout && typeof row.metadata.checkout === "object" ? row.metadata.checkout : {};
    const labelsCount = Math.max(0, Number(checkoutMetadata?.labels_count) || 0);
    const serviceType = String(checkoutMetadata?.service || row?.metadata?.service || "").trim();
    const estimatedRetail = getEstimatedRetailTotal(serviceType, labelsCount, revenueExVat);
    entry.total_revenue_ex_vat += toCents(revenueExVat);
    entry.estimated_carrier_cost_ex_vat += toCents(
      estimatedRetail * (1 - carrierDiscountPct / 100)
    );
  });

  out.forEach((entry) => {
    const revenue = fromCents(entry.total_revenue_ex_vat);
    const carrierCost = fromCents(entry.estimated_carrier_cost_ex_vat);
    entry.total_revenue_ex_vat = revenue;
    entry.estimated_carrier_cost_ex_vat = carrierCost;
    entry.total_profit_ex_vat = Number((revenue - carrierCost).toFixed(2));
  });
  return out;
}

function buildAdminClientMetrics(user, historyRows, billingMetrics, billingPreference, invoiceStats) {
  const rows = Array.isArray(historyRows) ? historyRows : [];
  const totalLabels = rows.reduce((sum, row) => sum + Math.max(0, Number(row?.quantity) || 0), 0);
  const normalizedBillingMetrics =
    billingMetrics && typeof billingMetrics === "object" ? billingMetrics : null;
  const totalRevenue = Number(normalizedBillingMetrics?.total_revenue_ex_vat || 0);
  const estimatedCarrierCost = Number(
    normalizedBillingMetrics?.estimated_carrier_cost_ex_vat || 0
  );
  const totalProfit = Number(normalizedBillingMetrics?.total_profit_ex_vat || 0);
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
  if (paymentMode === "wallet") {
    avgPaymentDays = totalRevenue > 0 ? 0 : null;
    lastInvoiceTracking = totalRevenue > 0 ? "Account balance debit" : "No billable activity";
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

function getInviteBoundEmail(invite) {
  return normalizeEmail(invite?.invited_email || "");
}

function inviteEmailMatches(invite, email) {
  const inviteEmail = getInviteBoundEmail(invite);
  return Boolean(inviteEmail && inviteEmail === normalizeEmail(email || ""));
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
  companyName,
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
        company_name: normalizeClientCompanyName(companyName) || null,
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

async function createRegistrationInvite({ invitedEmail, companyName, expiresInDays, createdBy }) {
  const safeInvitedEmail = normalizeEmail(invitedEmail || "");
  const safeCompanyName = normalizeClientCompanyName(companyName);
  const expiryDays = parseInviteExpiryDays(expiresInDays);
  const expiresAt = new Date(Date.now() + expiryDays * 24 * 60 * 60 * 1000).toISOString();

  for (let attempt = 0; attempt < 3; attempt += 1) {
    const token = createPublicLinkToken();
    const tokenHash = hashInviteToken(token);
    try {
      const row = await insertRegistrationInvite({
        tokenHash,
        tokenEncrypted: encryptToken(token),
        invitedEmail: safeInvitedEmail,
        companyName: safeCompanyName,
        expiresAt,
        createdBy,
      });
      return {
        id: row?.id || null,
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
    "id,invited_email,company_name,expires_at,created_at,claimed_at,claimed_email,created_by,token_encrypted,revoked_at,revoked_by"
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

function shipmentExtractRequestIsExpired(request) {
  const expiresAt = String(request?.expires_at || "").trim();
  if (!expiresAt) return true;
  const expiresMs = Date.parse(expiresAt);
  if (!Number.isFinite(expiresMs)) return true;
  return Date.now() > expiresMs;
}

async function getShipmentExtractRequestByToken(token) {
  const requestToken = normalizeInviteToken(token);
  if (!requestToken) return null;
  const tokenHash = hashInviteToken(requestToken);
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,client_email,expires_at,created_at,created_by,submitted_at,submitted_filename,submitted_rows,submitted_columns,kept_columns,removed_columns,registration_invite_id,token_encrypted,metadata"
  );
  params.set("token_hash", `eq.${tokenHash}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${SHIPMENT_EXTRACT_REQUESTS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Shipment extract request lookup failed (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function getShipmentExtractRequestById(requestId) {
  const safeRequestId = String(requestId || "").trim();
  if (!safeRequestId) return null;
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,client_email,expires_at,created_at,created_by,submitted_at,submitted_filename,submitted_rows,submitted_columns,kept_columns,removed_columns,registration_invite_id,token_encrypted,metadata"
  );
  params.set("id", `eq.${safeRequestId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${SHIPMENT_EXTRACT_REQUESTS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Shipment extract request lookup failed (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function insertShipmentExtractRequest({
  tokenHash,
  tokenEncrypted,
  clientEmail,
  companyName,
  expiresAt,
  createdBy,
}) {
  const safeCompanyName = normalizeClientCompanyName(companyName);
  const response = await supabaseServiceRequest(`/rest/v1/${SHIPMENT_EXTRACT_REQUESTS_TABLE}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify([
      {
        token_hash: tokenHash,
        token_encrypted: tokenEncrypted || null,
        client_email: clientEmail,
        expires_at: expiresAt,
        created_by: createdBy || null,
        metadata: safeCompanyName ? { client_company_name: safeCompanyName } : {},
      },
    ]),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not create shipment extract link (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function createShipmentExtractRequest({ clientEmail, companyName, expiresInDays, createdBy }) {
  const safeClientEmail = normalizeEmail(clientEmail || "");
  const safeCompanyName = normalizeClientCompanyName(companyName);
  const expiryDays = parseInviteExpiryDays(expiresInDays);
  const expiresAt = new Date(Date.now() + expiryDays * 24 * 60 * 60 * 1000).toISOString();

  for (let attempt = 0; attempt < 3; attempt += 1) {
    const token = createPublicLinkToken();
    const tokenHash = hashInviteToken(token);
    try {
      const row = await insertShipmentExtractRequest({
        tokenHash,
        tokenEncrypted: encryptToken(token),
        clientEmail: safeClientEmail,
        companyName: safeCompanyName,
        expiresAt,
        createdBy,
      });
      return { row, token, expiresAt };
    } catch (error) {
      const message = String(error?.message || "");
      if (message.includes("23505") || message.includes("duplicate key")) {
        continue;
      }
      throw error;
    }
  }
  throw new Error("Could not create shipment extract link. Please try again.");
}

async function listShipmentExtractRequests({ limit = 50 }) {
  const safeLimit = Math.max(1, Math.min(100, Number(limit) || 50));
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,client_email,expires_at,created_at,created_by,submitted_at,submitted_filename,submitted_rows,submitted_columns,kept_columns,removed_columns,registration_invite_id,token_encrypted,metadata"
  );
  params.set("order", "created_at.desc");
  params.set("limit", String(safeLimit));
  const response = await supabaseServiceRequest(
    `/rest/v1/${SHIPMENT_EXTRACT_REQUESTS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Shipment extract history lookup failed (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function updateShipmentExtractRequestSubmitted(requestId, payload = {}) {
  const safeRequestId = String(requestId || "").trim();
  if (!safeRequestId) return null;
  const params = new URLSearchParams();
  params.set("id", `eq.${safeRequestId}`);
  const response = await supabaseServiceRequest(
    `/rest/v1/${SHIPMENT_EXTRACT_REQUESTS_TABLE}?${params.toString()}`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=representation",
      },
      body: JSON.stringify({
        submitted_at: payload.submittedAt || new Date().toISOString(),
        submitted_filename: payload.sourceFilename || null,
        submitted_rows: Number(payload.rows) || 0,
        submitted_columns: Number(payload.columns) || 0,
        kept_columns: Array.isArray(payload.keptColumns) ? payload.keptColumns : [],
        removed_columns: Array.isArray(payload.removedColumns) ? payload.removedColumns : [],
        metadata: payload.metadata && typeof payload.metadata === "object" ? payload.metadata : {},
      }),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not update shipment extract request (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function linkShipmentExtractRequestToRegistrationInvite(requestId, inviteId) {
  const safeRequestId = String(requestId || "").trim();
  const safeInviteId = String(inviteId || "").trim();
  if (!safeRequestId || !safeInviteId) return;
  const params = new URLSearchParams();
  params.set("id", `eq.${safeRequestId}`);
  const response = await supabaseServiceRequest(
    `/rest/v1/${SHIPMENT_EXTRACT_REQUESTS_TABLE}?${params.toString()}`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=minimal",
      },
      body: JSON.stringify({ registration_invite_id: safeInviteId }),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not link registration invite (${response.status}) ${details}`.trim());
  }
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

async function listUnbilledGenerationRows({ window, userId = "" } = {}) {
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

async function listBillingInvoices({ limit = 100, userId = "", statuses = [], allowMissing = false } = {}) {
  const safeUserId = String(userId || "").trim();
  const safeLimit = Math.max(1, Math.min(1000, Number(limit) || 100));
  const params = new URLSearchParams();
  params.set("select", BILLING_INVOICE_SELECT_FIELDS);
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
    `/rest/v1/${BILLING_INVOICES_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (isBillingInvoicesSchemaMissingDetails(details)) {
      if (allowMissing) return [];
      throw new Error("Billing schema missing. Run supabase_invoicing.sql and supabase_topup_invoice_automation.sql in Supabase SQL editor.");
    }
    throw new Error(`Could not load invoices (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function getBillingInvoiceByPeriod(userId, periodStart, periodEnd) {
  const safeUserId = String(userId || "").trim();
  const safeStart = String(periodStart || "").trim();
  const safeEnd = String(periodEnd || "").trim();
  if (!safeUserId || !safeStart || !safeEnd) return null;
  const params = new URLSearchParams();
  params.set("select", BILLING_INVOICE_SELECT_FIELDS);
  params.set("user_id", `eq.${safeUserId}`);
  params.set("invoice_kind", "eq.monthly");
  params.set("period_start", `eq.${safeStart}`);
  params.set("period_end", `eq.${safeEnd}`);
  params.set("order", "created_at.asc,updated_at.asc,id.asc");
  params.set("limit", "5");
  const response = await supabaseServiceRequest(
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
  if (!Array.isArray(rows) || !rows.length) return null;
  const sorted = rows.slice().sort(compareCanonicalBillingInvoices);
  if (sorted.length > 1) {
    console.error("[billing invoice duplicate period select]", `${safeUserId}:${safeStart}:${safeEnd}`, sorted.map((invoice) => ({
      id: invoice?.id || null,
      status: invoice?.status || null,
      issued_at: invoice?.issued_at || null,
      sent_at: invoice?.sent_at || null,
      created_at: invoice?.created_at || null,
    })));
  }
  return sorted[0];
}

async function getBillingInvoiceById(invoiceId, { withItems = false } = {}) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId) return null;
  const params = new URLSearchParams();
  params.set("select", BILLING_INVOICE_SELECT_FIELDS);
  params.set("id", `eq.${safeInvoiceId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
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
  itemParams.set("select", BILLING_INVOICE_ITEM_SELECT_FIELDS);
  itemParams.set("invoice_id", `eq.${safeInvoiceId}`);
  itemParams.set("order", "sort_index.asc,id.asc");
  itemParams.set("limit", "10000");
  const itemsResponse = await supabaseServiceRequest(
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

async function upsertBillingInvoice(invoicePayload) {
  const response = await supabaseServiceRequest(
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

async function ensureBillingInvoicePublicCode(invoice = {}) {
  if (!invoice?.id) return invoice || null;
  const existingCode = String(
    invoice?.invoice_number || invoice?.metadata?.invoice_number || ""
  ).trim();
  if (existingCode) return invoice;
  const invoiceKind = getBillingInvoiceKind(invoice);
  let lastError = null;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const nextCode = buildBillingDocumentPublicCode(invoiceKind, "invoice");
    try {
      const updated = await updateBillingInvoiceFields(invoice.id, {
        invoice_number: nextCode,
        metadata: {
          ...(invoice?.metadata && typeof invoice.metadata === "object" ? invoice.metadata : {}),
          invoice_kind: invoiceKind,
          invoice_number: nextCode,
        },
      });
      return updated || {
        ...invoice,
        invoice_number: nextCode,
        metadata: {
          ...(invoice?.metadata && typeof invoice.metadata === "object" ? invoice.metadata : {}),
          invoice_kind: invoiceKind,
          invoice_number: nextCode,
        },
      };
    } catch (error) {
      if (!isPublicDocumentCodeCollisionError(error)) {
        throw error;
      }
      lastError = error;
    }
  }
  throw lastError || new Error("Could not allocate a unique invoice document code.");
}

async function upsertBillingInvoiceWithPublicCode(invoicePayload, existing = null) {
  const fixedCode = String(
    existing?.invoice_number || existing?.metadata?.invoice_number || invoicePayload?.invoice_number || ""
  ).trim();
  let lastError = null;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const invoiceNumber = fixedCode || buildBillingDocumentPublicCode("monthly", "invoice");
    const payload = {
      ...invoicePayload,
      invoice_number: invoiceNumber,
      metadata: {
        ...(invoicePayload?.metadata && typeof invoicePayload.metadata === "object"
          ? invoicePayload.metadata
          : {}),
        invoice_kind: "monthly",
        invoice_number: invoiceNumber,
      },
    };
    try {
      return await upsertBillingInvoice(payload);
    } catch (error) {
      if (fixedCode || !isPublicDocumentCodeCollisionError(error)) {
        throw error;
      }
      lastError = error;
    }
  }
  throw lastError || new Error("Could not allocate a unique monthly invoice document code.");
}

async function getBillingInvoiceBySourceTopupId(topupId, { withItems = false, allowMissing = false } = {}) {
  const safeTopupId = String(topupId || "").trim();
  if (!safeTopupId) return null;
  const params = new URLSearchParams();
  params.set("select", BILLING_INVOICE_SELECT_FIELDS);
  params.set("source_topup_id", `eq.${safeTopupId}`);
  params.set("invoice_kind", "eq.topup");
  params.set("order", "created_at.asc,updated_at.asc,id.asc");
  params.set("limit", "5");
  const response = await supabaseServiceRequest(
    `/rest/v1/${BILLING_INVOICES_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (/relation .*billing_invoices/i.test(details) || /invoice_kind|source_topup_id|invoice_number/i.test(details)) {
      if (allowMissing) return null;
      throw new Error("Billing schema missing. Run supabase_invoicing.sql and supabase_topup_invoice_automation.sql in Supabase SQL editor.");
    }
    throw new Error(`Could not load top-up invoice (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) return null;
  const sorted = rows.slice().sort(compareCanonicalBillingInvoices);
  if (sorted.length > 1) {
    console.error("[billing topup duplicate invoice]", safeTopupId, sorted.map((invoice) => ({
      id: invoice?.id || null,
      invoice_number: invoice?.invoice_number || null,
      status: invoice?.status || null,
      sent_at: invoice?.sent_at || null,
      created_at: invoice?.created_at || null,
    })));
  }
  const invoice = sorted[0];
  if (!invoice || !withItems) return invoice;
  return getBillingInvoiceById(invoice.id, { withItems: true });
}

async function createBillingTopupInvoice(invoicePayload) {
  const response = await supabaseServiceRequest(`/rest/v1/${BILLING_INVOICES_TABLE}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify([invoicePayload]),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (
      /relation .*billing_invoices/i.test(details)
      || /invoice_kind|source_topup_id|invoice_number/i.test(details)
    ) {
      throw new Error("Billing schema missing. Run supabase_invoicing.sql and supabase_topup_invoice_automation.sql in Supabase SQL editor.");
    }
    if ((details.includes("23505") || /duplicate key/i.test(details)) && invoicePayload?.source_topup_id) {
      const existing = await getBillingInvoiceBySourceTopupId(invoicePayload.source_topup_id, {
        withItems: false,
        allowMissing: true,
      });
      if (existing?.id) return existing;
    }
    throw new Error(`Could not create top-up invoice (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

function getStoredGenerationReceiptPublicCode(row = {}) {
  return String(row?.payload?.document?.public_code || "").trim();
}

async function getGenerationHistoryRowById(generationId, { allowMissing = false } = {}) {
  const safeGenerationId = String(generationId || "").trim();
  if (!safeGenerationId) return null;
  const params = new URLSearchParams();
  params.set("select", "id,user_id,created_at,service_type,quantity,total_price,payload,billed_invoice_id,billed_at");
  params.set("id", `eq.${safeGenerationId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${HISTORY_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (allowMissing && /relation .*label_generations/i.test(details)) {
      return null;
    }
    throw new Error(`Could not load generation row (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function updateGenerationHistoryPayload(generationId, payload = {}) {
  const safeGenerationId = String(generationId || "").trim();
  if (!safeGenerationId) throw new Error("Generation id is required.");
  const params = new URLSearchParams();
  params.set("id", `eq.${safeGenerationId}`);
  const response = await supabaseServiceRequest(
    `/rest/v1/${HISTORY_TABLE}?${params.toString()}`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Prefer: "return=representation",
      },
      body: JSON.stringify({
        payload,
      }),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not update generation row (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function ensureGenerationReceiptPublicCode(row = {}) {
  if (!row?.id) return { row: row || null, publicCode: "" };
  const existingCode = getStoredGenerationReceiptPublicCode(row);
  if (existingCode) {
    return { row, publicCode: existingCode };
  }
  const currentPayload = row?.payload && typeof row.payload === "object" ? row.payload : {};
  let lastError = null;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const nextCode = buildBillingDocumentPublicCode("receipt", "receipt");
    const nextPayload = {
      ...currentPayload,
      document: {
        ...(currentPayload?.document && typeof currentPayload.document === "object"
          ? currentPayload.document
          : {}),
        kind: "receipt",
        public_code: nextCode,
        filename: buildDocumentPdfFilenameFromReference(nextCode, "R-REC-UNKNOWN"),
      },
    };
    try {
      const updatedRow = await updateGenerationHistoryPayload(row.id, nextPayload);
      return {
        row: updatedRow || {
          ...row,
          payload: nextPayload,
        },
        publicCode: nextCode,
      };
    } catch (error) {
      if (!isPublicDocumentCodeCollisionError(error) && !/(23505|duplicate key)/i.test(String(error?.message || ""))) {
        throw error;
      }
      lastError = error;
    }
  }
  throw lastError || new Error("Could not allocate a unique receipt document code.");
}

async function replaceBillingInvoiceItems(invoiceId, items = []) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId) return;
  const deleteParams = new URLSearchParams();
  deleteParams.set("invoice_id", `eq.${safeInvoiceId}`);
  const deleteResponse = await supabaseServiceRequest(
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

  const response = await supabaseServiceRequest(`/rest/v1/${BILLING_INVOICE_ITEMS_TABLE}`, {
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

async function markGenerationRowsBilled(generationIds = [], invoiceId) {
  const validIds = generationIds.filter((value) => isUuid(value));
  if (!validIds.length || !isUuid(invoiceId)) return;
  const billedAt = new Date().toISOString();
  const chunkSize = 80;
  for (let index = 0; index < validIds.length; index += chunkSize) {
    const chunk = validIds.slice(index, index + chunkSize);
    const params = new URLSearchParams();
    params.set("id", `in.(${chunk.join(",")})`);
    params.set("billed_invoice_id", "is.null");
    const response = await supabaseServiceRequest(`/rest/v1/${HISTORY_TABLE}?${params.toString()}`, {
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

async function insertBillingInvoiceEvent(invoiceId, event) {
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
  const response = await supabaseServiceRequest(`/rest/v1/${BILLING_INVOICE_EVENTS_TABLE}`, {
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

function getInvoiceItemShipideSourceMetadata(payload = {}) {
  const source =
    payload?.shipideSource && typeof payload.shipideSource === "object"
      ? payload.shipideSource
      : null;
  if (!source) return null;

  const provider = String(source?.provider || "").trim();
  if (!provider) return null;

  return {
    provider,
    shop: normalizeShopDomain(source?.shop || "") || null,
    orderId: String(source?.orderId || "").trim() || null,
    orderName: String(source?.orderName || "").trim() || null,
    redacted: Boolean(source?.redacted),
    redactedAt: String(source?.redactedAt || "").trim() || null,
  };
}

function buildInvoiceItemsFromHistoryRows(rows = [], vatRate = DEFAULT_VAT_RATE) {
  const safeVatRate = Math.max(0, Number(vatRate) || DEFAULT_VAT_RATE);
  const items = [];
  const sortedRows = Array.isArray(rows)
    ? rows
        .slice()
        .sort((left, right) => Date.parse(left?.created_at || 0) - Date.parse(right?.created_at || 0))
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
        const shipideSourceMetadata = getInvoiceItemShipideSourceMetadata(labelData);
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
            ...(shipideSourceMetadata ? { shipide_source: shipideSourceMetadata } : {}),
          },
        });
      });
      return;
    }

    const unitExCents = Math.round(totalExCents / quantityFromRow);
    const vatCents = Math.round(totalExCents * safeVatRate);
    const selectionPayload = payload?.selection || payload || {};
    const recipient = getRecipientFromLabelPayload(selectionPayload);
    const shipideSourceMetadata = getInvoiceItemShipideSourceMetadata(selectionPayload);
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
        ...(shipideSourceMetadata ? { shipide_source: shipideSourceMetadata } : {}),
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

function formatInvoicePdfDate(value) {
  const iso = parseIsoTimestamp(value);
  return iso ? iso.slice(0, 10) : "--";
}

function formatInvoicePdfMoney(value) {
  return `EUR ${fromCents(toCents(value)).toFixed(2)}`;
}

function invoiceRequiresManualSettlement(invoice = {}) {
  return String(invoice?.payment_mode || "invoice").trim().toLowerCase() === "invoice";
}

function formatInvoicePaymentModeLabel(invoice = {}) {
  if (isTopupBillingInvoice(invoice)) return "Prepaid Credit for Shipping Labels";
  const mode = String(invoice?.payment_mode || "invoice").trim().toLowerCase();
  if (mode === "invoice") return "Monthly billing";
  return "Account balance";
}

function buildInvoicePdfFilename(invoice = {}) {
  return buildDocumentPdfFilenameFromReference(
    getBillingInvoiceReference(invoice),
    "INV-UNKNOWN"
  );
}

function buildInvoiceVariantPdfFilename(invoice = {}, reminderStage = 0) {
  const base = buildInvoicePdfFilename(invoice).replace(/\.pdf$/i, "");
  const stage = Math.max(0, Number(reminderStage) || 0);
  if (stage <= 0) return `${base}.pdf`;
  return `${base}-stage-${stage}.pdf`;
}

function splitInvoiceAddressLines(value) {
  const raw = String(value || "").trim();
  if (!raw) return [];
  const directLines = raw
    .split(/\r?\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (directLines.length > 1) return directLines;
  const parts = raw
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length <= 2) return parts;
  return [
    parts[0],
    parts.slice(1, -1).join(", "),
    parts.at(-1),
  ].filter(Boolean);
}

function fitPdfText(text, font, size, maxWidth) {
  const raw = String(text || "").trim();
  if (!raw) return "";
  if (font.widthOfTextAtSize(raw, size) <= maxWidth) return raw;
  let output = raw;
  while (output.length > 1) {
    output = output.slice(0, -1).trimEnd();
    const candidate = `${output}...`;
    if (font.widthOfTextAtSize(candidate, size) <= maxWidth) {
      return candidate;
    }
  }
  return "...";
}

function wrapPdfText(text, font, size, maxWidth, maxLines = 4) {
  const source = String(text || "")
    .replace(/\s+/g, " ")
    .trim();
  if (!source) return [];
  const words = source.split(" ");
  const lines = [];
  let current = "";
  words.forEach((word) => {
    const candidate = current ? `${current} ${word}` : word;
    if (!current || font.widthOfTextAtSize(candidate, size) <= maxWidth) {
      current = candidate;
      return;
    }
    lines.push(current);
    current = word;
  });
  if (current) lines.push(current);
  if (lines.length <= maxLines) return lines;
  const clipped = lines.slice(0, maxLines);
  clipped[maxLines - 1] = fitPdfText(clipped[maxLines - 1], font, size, maxWidth);
  return clipped;
}

function normalizeInvoiceProfileSnapshot(rawProfile = {}) {
  const profile = rawProfile && typeof rawProfile === "object" ? rawProfile : {};
  return {
    companyName: String(profile.companyName || profile.company_name || "").trim(),
    contactName: String(profile.contactName || profile.contact_name || "").trim(),
    contactEmail: normalizeEmail(profile.contactEmail || profile.contact_email || ""),
    contactPhone: String(profile.contactPhone || profile.contact_phone || "").trim(),
    billingAddress: String(profile.billingAddress || profile.billing_address || "").trim(),
    taxId: String(profile.taxId || profile.tax_id || "").trim(),
    customerId: String(profile.customerId || profile.customer_id || "").trim(),
    accountManager: String(profile.accountManager || profile.account_manager || "").trim(),
  };
}

function resolveInvoiceProfileSnapshot(invoice = {}, user = null) {
  const metadata =
    invoice?.metadata && typeof invoice.metadata === "object" ? invoice.metadata : {};
  const snapshot = normalizeInvoiceProfileSnapshot(metadata.invoice_profile);
  const userProfile = user ? mapInvoiceProfileFromUser(user) : {};
  return {
    companyName:
      snapshot.companyName
      || String(invoice?.company_name || "").trim()
      || userProfile.company_name
      || "Client account",
    contactName:
      snapshot.contactName
      || String(invoice?.contact_name || "").trim()
      || userProfile.contact_name
      || "",
    contactEmail:
      snapshot.contactEmail
      || normalizeEmail(invoice?.contact_email || "")
      || userProfile.contact_email
      || "",
    contactPhone: snapshot.contactPhone || userProfile.contact_phone || "",
    billingAddress: snapshot.billingAddress || userProfile.billing_address || "",
    taxId: snapshot.taxId || userProfile.tax_id || "",
    customerId:
      snapshot.customerId
      || String(invoice?.customer_id || "").trim()
      || userProfile.customer_id
      || "",
    accountManager:
      snapshot.accountManager
      || String(invoice?.account_manager || "").trim()
      || userProfile.account_manager
      || "",
  };
}

function startOfUtcDay(dateLike) {
  const date = new Date(dateLike);
  if (Number.isNaN(date.getTime())) return null;
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

function getInvoiceSettlementBadgeMeta(invoice = {}, options = {}) {
  const manual = invoiceRequiresManualSettlement(invoice);
  const status = normalizeInvoiceStatus(invoice?.status);
  const reminderStage = Number(options?.reminderStage) || 0;
  if (isTopupBillingInvoice(invoice)) {
    return { label: "Account balance credited", tone: "success" };
  }
  if (manual && reminderStage > 0) {
    const mapping = {
      1: { label: "Due in 15 days", tone: "warning" },
      2: { label: "Due in 7 days", tone: "warning" },
      3: { label: "Due tomorrow", tone: "attention" },
      4: { label: "Due today", tone: "danger" },
      5: { label: "3 days overdue", tone: "danger" },
      6: { label: "10 days overdue", tone: "danger" },
      7: { label: "15 days overdue", tone: "danger" },
      8: { label: "30 days overdue", tone: "danger" },
    };
    if (mapping[reminderStage]) return mapping[reminderStage];
  }
  if (!manual) {
    if (status === "paid") return { label: "Collected automatically", tone: "success" };
    if (status === "overdue") return { label: "Collection overdue", tone: "danger" };
    return { label: "Pending collection", tone: "success" };
  }
  if (status === "paid") {
    return { label: "Paid", tone: "success" };
  }
  const dueAt = options?.dueAt || invoice?.due_at || null;
  const today = startOfUtcDay(new Date());
  const dueDay = startOfUtcDay(dueAt);
  if (!today || !dueDay) return { label: `Due in ${getInvoiceTermsDays()} days`, tone: "success" };
  const diffDays = Math.round((dueDay.getTime() - today.getTime()) / 86400000);
  if (diffDays > 15) return { label: `Due in ${diffDays} days`, tone: "success" };
  if (diffDays > 1) return { label: `Due in ${diffDays} days`, tone: "warning" };
  if (diffDays === 1) return { label: "Due tomorrow", tone: "attention" };
  if (diffDays === 0) return { label: "Due today", tone: "danger" };
  if (diffDays === -1) return { label: "1 day overdue", tone: "danger" };
  return { label: `${Math.abs(diffDays)} days overdue`, tone: "danger" };
}

function getAutomatedInvoiceStatusLabel(invoice = {}) {
  return isTopupBillingInvoice(invoice) ? "Account balance credited" : "Paid automatically";
}

function buildInvoicePdfSettlementModel(invoice = {}, options = {}) {
  const manual = invoiceRequiresManualSettlement(invoice);
  const issuedAt = options?.issuedAt || invoice?.issued_at || null;
  const dueAt = options?.dueAt || invoice?.due_at || null;
  const paidAt = options?.paidAt || invoice?.paid_at || issuedAt || null;
  const reference = String(
    invoice?.payment_reference
      || invoice?.metadata?.topup_reference
      || getBillingInvoiceReference(invoice)
  ).trim();
  const badgeMeta = getInvoiceSettlementBadgeMeta(invoice, {
    dueAt,
    paidAt,
    reminderStage: options?.reminderStage,
  });
  if (manual) {
    return {
      badge: badgeMeta.label,
      badgeTone: badgeMeta.tone,
      transferRows: [
        { label: "IBAN", value: DEFAULT_INVOICE_PDF_IBAN, mono: true },
        { label: "Communication", value: reference, mono: true },
      ],
      lines: [],
      note: "",
    };
  }
  if (isTopupBillingInvoice(invoice)) {
    return {
      badge: badgeMeta.label,
      badgeTone: badgeMeta.tone,
      transferRows: [
        { label: "Transfer reference", value: reference || "--", mono: true },
      ],
      lines: [],
      note: "",
    };
  }
  return {
    badge: badgeMeta.label,
    badgeTone: badgeMeta.tone,
    transferRows: [],
    lines: [],
    note: "",
  };
}

function formatInvoicePdfServiceName(value) {
  const raw = String(value || "").trim();
  if (!raw) return "Shipping labels";
  return raw
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function buildInvoicePdfServiceRows(items = [], invoice = {}) {
  const grouped = new Map();
  const safeItems = Array.isArray(items) && items.length
    ? items
    : [{
        service_type: "Shipping labels",
        quantity: Math.max(1, Number(invoice?.labels_count) || 1),
        amount_inc_vat: invoice?.total_inc_vat ?? invoice?.subtotal_ex_vat ?? 0,
        sort_index: 0,
      }];

  safeItems.forEach((item, index) => {
    const service = formatInvoicePdfServiceName(item?.service_type || "Shipping labels");
    const key = service.toLowerCase();
    const sortIndex = Number.isFinite(Number(item?.sort_index)) ? Number(item.sort_index) : index;
    const quantity = Math.max(1, Number(item?.quantity) || 1);
    const totalCents = toCents(item?.amount_inc_vat ?? item?.amount_ex_vat);
    const existing = grouped.get(key) || {
      service,
      quantity: 0,
      totalCents: 0,
      sortIndex,
    };
    existing.quantity += quantity;
    existing.totalCents += totalCents;
    existing.sortIndex = Math.min(existing.sortIndex, sortIndex);
    grouped.set(key, existing);
  });

  return Array.from(grouped.values())
    .sort((left, right) => left.sortIndex - right.sortIndex)
    .map((entry) => {
      const unitCents = entry.quantity ? Math.round(entry.totalCents / entry.quantity) : entry.totalCents;
      return {
        service: entry.service,
        quantity: String(entry.quantity),
        rate: formatInvoicePdfMoney(fromCents(unitCents)),
        total: formatInvoicePdfMoney(fromCents(entry.totalCents)),
      };
    });
}

function drawInvoicePdfBrand(page, fonts, colors, x, yTop) {
  const mark = 10;
  const gap = 3;
  const markBottom = yTop - mark;
  page.drawRectangle({
    x,
    y: markBottom,
    width: mark,
    height: mark,
    borderColor: colors.text,
    borderWidth: 1,
  });
  page.drawRectangle({
    x: x + gap + 4,
    y: markBottom + gap + 4,
    width: mark,
    height: mark,
    borderColor: colors.text,
    borderWidth: 1,
  });
  page.drawText("Shipide", {
    x: x + 26,
    y: yTop - 15,
    font: fonts.bold,
    size: 18,
    color: colors.text,
  });
}

function drawInvoicePdfCard(page, colors, x, yTop, width, height) {
  page.drawRectangle({
    x,
    y: yTop - height,
    width,
    height,
    color: colors.panel,
    borderColor: colors.stroke,
    borderWidth: 1,
  });
}

function drawInvoicePdfMetaList(page, fonts, colors, entries, x, yTop, width, options = {}) {
  const labelWidth = Number(options?.labelWidth) || 82;
  const rowGap = Number(options?.rowGap) || 18;
  const labelSize = Number(options?.labelSize) || 8;
  const valueSize = Number(options?.valueSize) || 10;
  let cursor = yTop;
  entries.forEach((entry) => {
    const label = String(entry?.label || "").trim();
    const value = String(entry?.value || "--").trim() || "--";
    const valueFont = entry?.mono ? fonts.mono : fonts.regular;
    page.drawText(label, {
      x,
      y: cursor - labelSize,
      font: fonts.regular,
      size: labelSize,
      color: colors.muted,
    });
    page.drawText(fitPdfText(value, valueFont, valueSize, Math.max(20, width - labelWidth)), {
      x: x + labelWidth,
      y: cursor - valueSize,
      font: valueFont,
      size: valueSize,
      color: colors.text,
    });
    cursor -= rowGap;
  });
}

async function buildInvoicePdf(invoice = {}, items = [], options = {}) {
  const pdf = await PDFDocument.create();
  const fonts = {
    regular: await pdf.embedFont(StandardFonts.Helvetica),
    bold: await pdf.embedFont(StandardFonts.HelveticaBold),
    mono: await pdf.embedFont(StandardFonts.Courier),
  };
  const colors = {
    page: rgb(11 / 255, 16 / 255, 24 / 255),
    panel: rgb(28 / 255, 34 / 255, 44 / 255),
    stroke: rgb(53 / 255, 60 / 255, 75 / 255),
    text: rgb(243 / 255, 246 / 255, 255 / 255),
    muted: rgb(154 / 255, 163 / 255, 178 / 255),
    accent: rgb(119 / 255, 71 / 255, 227 / 255),
    accentSoft: rgb(58 / 255, 37 / 255, 112 / 255),
    success: rgb(143 / 255, 226 / 255, 178 / 255),
  };
  const reference = getBillingInvoiceReference(invoice);
  const profile = resolveInvoiceProfileSnapshot(invoice, options?.user || null);
  const settlement = buildInvoicePdfSettlementModel(invoice, options);
  const issueDate = formatInvoicePdfDate(options?.issuedAt || invoice?.issued_at);
  const issuer = DEFAULT_INVOICE_ISSUER;
  const lineRows = buildInvoicePdfServiceRows(items, invoice);
  const labelCount = Math.max(
    1,
    Number(invoice?.labels_count)
      || lineRows.reduce((sum, row) => sum + Math.max(1, Number(row.quantity) || 1), 0)
  );
  const settlementHeight = 128;
  const totalAmount = formatInvoicePdfMoney(invoice?.total_inc_vat ?? invoice?.subtotal_ex_vat ?? 0);
  const dueOrStatusLabel = invoiceRequiresManualSettlement(invoice) ? "Due date" : "Status";
  const dueOrStatusValue = invoiceRequiresManualSettlement(invoice)
    ? formatInvoicePdfDate(options?.dueAt || invoice?.due_at)
    : getAutomatedInvoiceStatusLabel(invoice);
  const taxNote =
    "VAT not charged. Any reverse-charge VAT is due by the recipient.";
  const badgeStyles = {
    success: {
      color: colors.success,
      fill: rgb(28 / 255, 49 / 255, 38 / 255),
    },
    warning: {
      color: rgb(241 / 255, 194 / 255, 120 / 255),
      fill: rgb(47 / 255, 37 / 255, 21 / 255),
    },
    attention: {
      color: rgb(255 / 255, 200 / 255, 137 / 255),
      fill: rgb(52 / 255, 36 / 255, 19 / 255),
    },
    danger: {
      color: rgb(255 / 255, 188 / 255, 200 / 255),
      fill: rgb(53 / 255, 26 / 255, 30 / 255),
    },
    accent: {
      color: colors.accent,
      fill: colors.accentSoft,
    },
  };
  const selectedBadgeStyle = badgeStyles[settlement.badgeTone] || badgeStyles.accent;
  const paymentBadgeColor = selectedBadgeStyle.color;
  const paymentBadgeFill = selectedBadgeStyle.fill;

  const pageWidth = 595.28;
  const pageHeight = 841.89;
  const marginX = 26;
  const marginTop = 28;
  const marginBottom = 24;
  const footerHeight = 20;
  const contentWidth = pageWidth - marginX * 2;
  const tableGap = 16;
  const rowHeight = 28;
  const tableHeadHeight = 24;
  const tableTitleHeight = 42;
  const tablePadding = 14;
  const taxNoteLines = wrapPdfText(taxNote, fonts.regular, 8.5, contentWidth, 3);
  const taxNoteHeight = taxNoteLines.length * 11;
  const settlementNoteGap = 10;
  const settlementBlockHeight = settlementHeight + settlementNoteGap + taxNoteHeight;
  let rowIndex = 0;
  while (rowIndex < Math.max(1, lineRows.length)) {
    const firstPage = rowIndex === 0;
    const page = pdf.addPage([pageWidth, pageHeight]);
    page.drawRectangle({ x: 0, y: 0, width: pageWidth, height: pageHeight, color: colors.page });
    page.drawRectangle({ x: 0, y: pageHeight - 4, width: pageWidth, height: 4, color: colors.accent });

    const footerY = marginBottom;
    page.drawLine({
      start: { x: marginX, y: footerY + footerHeight + 8 },
      end: { x: pageWidth - marginX, y: footerY + footerHeight + 8 },
      thickness: 1,
      color: colors.stroke,
      opacity: 0.65,
    });
    page.drawText("Cryvelin LLC", {
      x: marginX,
      y: footerY,
      font: fonts.regular,
      size: 9,
      color: colors.muted,
    });
    page.drawText(reference, {
      x: pageWidth - marginX - fonts.mono.widthOfTextAtSize(reference, 9),
      y: footerY,
      font: fonts.mono,
      size: 9,
      color: colors.muted,
    });

    let cursorY = pageHeight - marginTop;
    if (firstPage) {
      drawInvoicePdfBrand(page, fonts, colors, marginX, cursorY - 2);
      const tagLabel = "Invoice";
      const tagLabelWidth = fonts.mono.widthOfTextAtSize(tagLabel, 11);
      const tagLabelX = pageWidth - marginX - tagLabelWidth;
      page.drawRectangle({
        x: tagLabelX - 14,
        y: cursorY - 10,
        width: 6,
        height: 6,
        color: colors.accent,
      });
      page.drawText(tagLabel, {
        x: tagLabelX,
        y: cursorY - 12,
        font: fonts.mono,
        size: 11,
        color: colors.text,
      });
      page.drawText(issueDate, {
        x: pageWidth - marginX - fonts.mono.widthOfTextAtSize(issueDate, 9),
        y: cursorY - 30,
        font: fonts.mono,
        size: 9,
        color: colors.muted,
      });
      page.drawText(reference, {
        x: pageWidth - marginX - fonts.mono.widthOfTextAtSize(reference, 9),
        y: cursorY - 44,
        font: fonts.mono,
        size: 9,
        color: colors.muted,
      });
      cursorY -= 58;

      page.drawLine({
        start: { x: marginX, y: cursorY },
        end: { x: pageWidth - marginX, y: cursorY },
        thickness: 1,
        color: colors.stroke,
        opacity: 0.7,
      });
      cursorY -= 18;

      const partyGap = 18;
      const partyWidth = (contentWidth - partyGap) / 2;
      const billedToTop = cursorY;
      const billedToLines = [
        ...splitInvoiceAddressLines(profile.billingAddress),
        profile.taxId ? `VAT ${profile.taxId}` : "",
      ].filter(Boolean);
      const billedByLines = [
        ...(Array.isArray(issuer.addressLines)
          ? issuer.addressLines
          : splitInvoiceAddressLines(issuer.jurisdiction)
        ).map((line) => ({ text: line, font: fonts.regular, size: 9 })),
        { text: issuer.ein, font: fonts.regular, size: 9 },
      ].filter((line) => String(line?.text || "").trim());
      const partyLineCount = Math.max(billedToLines.length, billedByLines.length, 1);
      const partyBlockHeight = 52 + partyLineCount * 16;
      page.drawText("Billed to", {
        x: marginX,
        y: billedToTop - 9,
        font: fonts.regular,
        size: 9,
        color: colors.muted,
      });
      page.drawText(fitPdfText(profile.companyName || "Client account", fonts.bold, 13, partyWidth), {
        x: marginX,
        y: billedToTop - 31,
        font: fonts.bold,
        size: 13,
        color: colors.text,
      });
      let billedToLineY = billedToTop - 50;
      billedToLines.forEach((line, index) => {
        const valueFont = index === billedToLines.length - 1 ? fonts.mono : fonts.regular;
        const valueLines = wrapPdfText(line, valueFont, 9, partyWidth, 1);
        if (valueLines[0]) {
          page.drawText(valueLines[0], {
            x: marginX,
            y: billedToLineY,
            font: valueFont,
            size: 9,
            color: colors.muted,
          });
          billedToLineY -= 16;
        }
      });

      const billedByX = marginX + partyWidth + partyGap;
      page.drawText("Billed by", {
        x: billedByX,
        y: billedToTop - 9,
        font: fonts.regular,
        size: 9,
        color: colors.muted,
      });
      const issuerName = String(issuer.legalName || "Cryvelin LLC");
      page.drawText(fitPdfText(issuerName, fonts.bold, 13, partyWidth), {
        x: billedByX,
        y: billedToTop - 31,
        font: fonts.bold,
        size: 13,
        color: colors.text,
      });
      const issuerNameWidth = fonts.bold.widthOfTextAtSize(issuerName, 13);
      const descriptor = String(issuer.descriptor || "").trim();
      if (descriptor) {
        page.drawText(fitPdfText(descriptor, fonts.regular, 8.5, Math.max(20, partyWidth - issuerNameWidth - 8)), {
          x: billedByX + issuerNameWidth + 8,
          y: billedToTop - 28.5,
          font: fonts.regular,
          size: 8.5,
          color: colors.muted,
        });
      }
      let billedByLineY = billedToTop - 50;
      billedByLines.forEach((line) => {
        const valueFont = line.font || fonts.regular;
        const valueSize = Number(line.size) || 9;
        const valueLines = wrapPdfText(line.text, valueFont, valueSize, partyWidth, 1);
        if (valueLines[0]) {
          page.drawText(valueLines[0], {
            x: billedByX,
            y: billedByLineY,
            font: valueFont,
            size: valueSize,
            color: colors.muted,
          });
          billedByLineY -= 16;
        }
      });

      cursorY -= partyBlockHeight;
      page.drawLine({
        start: { x: marginX, y: cursorY },
        end: { x: pageWidth - marginX, y: cursorY },
        thickness: 1,
        color: colors.stroke,
        opacity: 0.7,
      });
      cursorY -= 16;

      const metaGap = 18;
      const metaWidth = (contentWidth - metaGap * 3) / 4;
      const metaEntries = [
        { label: "Invoice no.", value: reference, mono: true },
        { label: "Invoice date", value: issueDate, mono: true },
        { label: dueOrStatusLabel, value: dueOrStatusValue, mono: invoiceRequiresManualSettlement(invoice) },
        { label: "Payment", value: formatInvoicePaymentModeLabel(invoice), mono: false },
      ];
      metaEntries.forEach((entry, index) => {
        const x = marginX + index * (metaWidth + metaGap);
        const valueFont = entry.mono ? fonts.mono : fonts.regular;
        page.drawText(entry.label, {
          x,
          y: cursorY - 8,
          font: fonts.regular,
          size: 8,
          color: colors.muted,
        });
        page.drawText(fitPdfText(entry.value, valueFont, 10.5, metaWidth), {
          x,
          y: cursorY - 28,
          font: valueFont,
          size: 10.5,
          color: colors.text,
        });
      });
      cursorY -= 44 + tableGap;
    } else {
      cursorY -= 6;
    }

    const tableTop = cursorY;
    const titleHeight = firstPage ? tableTitleHeight : 0;
    const availableHeight = tableTop - (marginBottom + footerHeight + 16);
    const remainingRows = lineRows.length - rowIndex;
    const maxRowsWithSettlement = Math.max(
      1,
      Math.floor(
        (availableHeight - settlementBlockHeight - tableGap - titleHeight - tableHeadHeight - tablePadding * 2)
          / rowHeight
      )
    );
    const maxRowsNoSettlement = Math.max(
      1,
      Math.floor((availableHeight - titleHeight - tableHeadHeight - tablePadding * 2) / rowHeight)
    );
    const showSettlement = remainingRows <= maxRowsWithSettlement;
    const rowsForPage = lineRows.slice(
      rowIndex,
      rowIndex + Math.min(remainingRows, showSettlement ? maxRowsWithSettlement : maxRowsNoSettlement)
    );
    rowIndex += rowsForPage.length;

    const tableHeight =
      titleHeight + tableHeadHeight + tablePadding * 2 + rowsForPage.length * rowHeight;
    drawInvoicePdfCard(page, colors, marginX, tableTop, contentWidth, tableHeight);

    let tableCursor = tableTop - tablePadding;
    if (firstPage) {
      page.drawText("Services", {
        x: marginX + 16,
        y: tableCursor - 12,
        font: fonts.bold,
        size: 13,
        color: colors.text,
      });
      page.drawText("Condensed by billed service.", {
        x: marginX + 16,
        y: tableCursor - 28,
        font: fonts.regular,
        size: 9,
        color: colors.muted,
      });
      const badge = `${labelCount} labels`;
      page.drawText(badge, {
        x: pageWidth - marginX - 16 - fonts.mono.widthOfTextAtSize(badge, 9),
        y: tableCursor - 20,
        font: fonts.mono,
        size: 9,
        color: colors.text,
      });
      tableCursor -= titleHeight;
    }

    const colX = {
      service: marginX + 16,
      qty: marginX + 356,
      rateRight: pageWidth - marginX - 118,
      totalRight: pageWidth - marginX - 16,
    };
    const colWidths = {
      service: 300,
    };

    page.drawLine({
      start: { x: marginX + 14, y: tableCursor - tableHeadHeight + 4 },
      end: { x: pageWidth - marginX - 14, y: tableCursor - tableHeadHeight + 4 },
      thickness: 1,
      color: colors.stroke,
      opacity: 0.8,
    });
    [
      ["Service", colX.service],
      ["Qty", colX.qty],
    ].forEach(([label, x]) => {
      page.drawText(label, {
        x,
        y: tableCursor - 18,
        font: fonts.regular,
        size: 8,
        color: colors.muted,
      });
    });
    const rateLabel = "Rate";
    const totalLabel = "Line total";
    page.drawText(rateLabel, {
      x: colX.rateRight - fonts.regular.widthOfTextAtSize(rateLabel, 8),
      y: tableCursor - 18,
      font: fonts.regular,
      size: 8,
      color: colors.muted,
    });
    page.drawText(totalLabel, {
      x: colX.totalRight - fonts.regular.widthOfTextAtSize(totalLabel, 8),
      y: tableCursor - 18,
      font: fonts.regular,
      size: 8,
      color: colors.muted,
    });
    tableCursor -= tableHeadHeight;

    rowsForPage.forEach((row, index) => {
      const baseY = tableCursor - index * rowHeight;
      if (index > 0) {
        page.drawLine({
          start: { x: marginX + 14, y: baseY + 4 },
          end: { x: pageWidth - marginX - 14, y: baseY + 4 },
          thickness: 1,
          color: colors.stroke,
          opacity: 0.45,
        });
      }
      page.drawText(
        fitPdfText(row.service, fonts.regular, 9, colWidths.service),
        {
          x: colX.service,
          y: baseY - 14,
          font: fonts.regular,
          size: 9,
          color: colors.text,
        }
      );
      page.drawText(row.quantity, {
        x: colX.qty,
        y: baseY - 14,
        font: fonts.mono,
        size: 9,
        color: colors.text,
      });
      const rateWidth = fonts.mono.widthOfTextAtSize(row.rate, 9);
      page.drawText(row.rate, {
        x: colX.rateRight - rateWidth,
        y: baseY - 14,
        font: fonts.mono,
        size: 9,
        color: colors.text,
      });
      const totalWidth = fonts.mono.widthOfTextAtSize(row.total, 9);
      page.drawText(row.total, {
        x: colX.totalRight - totalWidth,
        y: baseY - 14,
        font: fonts.mono,
        size: 9,
        color: colors.text,
      });
    });

    if (showSettlement) {
      const settlementTop = tableTop - tableHeight - tableGap;
      const summaryWidth = 148;
      const summaryX = pageWidth - marginX - summaryWidth;
      const paymentX = marginX + 16;
      const transferRows = Array.isArray(settlement.transferRows) ? settlement.transferRows : [];
      const hasTransferRows = transferRows.length > 0;
      const paymentWidth = hasTransferRows ? 118 : summaryX - marginX - 28;
      const transferX = paymentX + paymentWidth + 26;
      const transferWidth = hasTransferRows ? Math.max(92, summaryX - transferX - 18) : 0;
      drawInvoicePdfCard(page, colors, marginX, settlementTop, contentWidth, settlementHeight);
      page.drawText("Payment", {
        x: paymentX,
        y: settlementTop - 18,
        font: fonts.regular,
        size: 9,
        color: colors.muted,
      });
      const badgeWidth = Math.min(paymentWidth, fonts.regular.widthOfTextAtSize(settlement.badge, 8) + 18);
      page.drawRectangle({
        x: paymentX,
        y: settlementTop - 42,
        width: badgeWidth,
        height: 16,
        color: paymentBadgeFill,
        borderColor: paymentBadgeColor,
        borderWidth: 1,
      });
      const badgeTextWidth = fonts.regular.widthOfTextAtSize(settlement.badge, 8);
      page.drawText(settlement.badge, {
        x: paymentX + (badgeWidth - badgeTextWidth) / 2,
        y: settlementTop - 36.5,
        font: fonts.regular,
        size: 8,
        color: paymentBadgeColor,
      });

      if (hasTransferRows) {
        page.drawLine({
          start: { x: transferX - 18, y: settlementTop - settlementHeight + 16 },
          end: { x: transferX - 18, y: settlementTop - 16 },
          thickness: 1,
          color: colors.stroke,
          opacity: 0.7,
        });
        transferRows.forEach((row, index) => {
          const rowY = index === 0 ? settlementTop - 48 : settlementTop - 92;
          const valueFont = row.mono ? fonts.mono : fonts.regular;
          const labelWidth = 78;
          const valueX = transferX + labelWidth;
          const valueWidth = Math.max(36, transferWidth - labelWidth);
          page.drawText(row.label, {
            x: transferX,
            y: rowY,
            font: fonts.regular,
            size: 9,
            color: colors.muted,
          });
          const fittedValue = fitPdfText(row.value, valueFont, 9.5, valueWidth);
          page.drawText(fittedValue, {
            x: valueX,
            y: rowY,
            font: valueFont,
            size: 9.5,
            color: colors.text,
          });
        });
      }

      page.drawLine({
        start: { x: summaryX - 18, y: settlementTop - settlementHeight + 16 },
        end: { x: summaryX - 18, y: settlementTop - 16 },
        thickness: 1,
        color: colors.stroke,
        opacity: 0.7,
      });
      page.drawText("Summary", {
        x: summaryX,
        y: settlementTop - 18,
        font: fonts.regular,
        size: 9,
        color: colors.muted,
      });
      page.drawText("Subtotal", {
        x: summaryX,
        y: settlementTop - 48,
        font: fonts.regular,
        size: 9,
        color: colors.muted,
      });
      page.drawText(totalAmount, {
        x: pageWidth - marginX - 16 - fonts.mono.widthOfTextAtSize(totalAmount, 9),
        y: settlementTop - 48,
        font: fonts.mono,
        size: 9,
        color: colors.text,
      });
      page.drawLine({
        start: { x: summaryX, y: settlementTop - 64 },
        end: { x: pageWidth - marginX - 16, y: settlementTop - 64 },
        thickness: 1,
        color: colors.stroke,
        opacity: 0.7,
      });
      page.drawText("Total", {
        x: summaryX,
        y: settlementTop - 92,
        font: fonts.bold,
        size: 11,
        color: colors.text,
      });
      page.drawText(totalAmount, {
        x: pageWidth - marginX - 16 - fonts.mono.widthOfTextAtSize(totalAmount, 11),
        y: settlementTop - 92,
        font: fonts.mono,
        size: 11,
        color: colors.text,
      });

      const taxNoteTop = settlementTop - settlementHeight - settlementNoteGap;
      taxNoteLines.forEach((line, index) => {
        page.drawText(line, {
          x: marginX,
          y: taxNoteTop - index * 11,
          font: fonts.regular,
          size: 8.5,
          color: colors.muted,
        });
      });
    }
  }

  return Buffer.from(await pdf.save());
}

function buildInvoiceEmailHtml(invoice, items = [], options = {}) {
  const isReminder = Boolean(options?.isReminder);
  const reminderStage = Math.max(0, Number(options?.reminderStage ?? invoice?.reminder_stage) || 0);
  const viewUrl = String(options?.viewUrl || "#").trim() || "#";
  const hasViewUrl = viewUrl !== "#";
  const isOverdueReminder = isReminder && reminderStage >= 5;
  const manualSettlement = invoiceRequiresManualSettlement(invoice);
  const isTopup = isTopupBillingInvoice(invoice);
  const logoUrl = "https://portal.shipide.com/shipide_logo.png";
  const titleLine1 = isTopup
    ? "Here's your"
    : isOverdueReminder
      ? "Urgent Invoice"
      : isReminder
        ? "Payment Reminder"
        : "Your Monthly Invoice";
  const titleLine2 = isTopup
    ? "latest invoice"
    : isOverdueReminder
      ? "Is Overdue"
      : isReminder
        ? "Invoice Due Soon"
        : "Is Ready";
  const subtitleLine1 = isTopup
    ? "Your top-up invoice PDF is attached to this email."
    : "Your invoice PDF is attached to this email.";
  const subtitleLine2 = hasViewUrl ? "You can also view it instantly with the button below." : "";

  let stageLabel = "Due in 30 days";
  let stageBg = "rgba(143, 226, 178, 0.16)";
  let stageBorder = "#8fe2b2";
  let stageText = "#8fe2b2";
  if (isTopup) {
    stageLabel = "Account balance credited";
  } else if (!manualSettlement) {
    stageLabel = "Paid automatically";
    stageBg = "rgba(143, 226, 178, 0.16)";
    stageBorder = "#8fe2b2";
    stageText = "#8fe2b2";
  } else if (reminderStage === 1) {
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
                  ${escapeHtml(subtitleLine1)}${subtitleLine2 ? `<br/>${escapeHtml(subtitleLine2)}` : ""}
                </div>
                ${hasViewUrl
                  ? `<a href="${escapeHtml(viewUrl)}" style="display:inline-block;margin-top:24px;padding:12px 20px;border-radius:4px;border:1px solid rgb(46,46,46);background:#1c2026;color:#f3f6ff;text-decoration:none;font-size:14px;line-height:1;">
                  View Invoice
                </a>`
                  : ""}
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
  const reference = getBillingInvoiceReference(invoice);
  const dueLabel = String(invoice?.due_at || "").trim()
    ? new Date(invoice.due_at).toISOString().slice(0, 10)
    : "--";
  const isTopup = isTopupBillingInvoice(invoice);
  const prefix = options?.isReminder ? "Payment reminder" : isTopup ? "Top-up invoice" : "Invoice";
  const viewUrl = String(options?.viewUrl || "").trim();
  const lines = [
    `${prefix}: ${reference}`,
    `Client: ${invoice?.company_name || "Client account"}`,
    `Subtotal: €${fromCents(toCents(invoice?.subtotal_ex_vat)).toFixed(2)}`,
    `Total: €${fromCents(toCents(invoice?.total_inc_vat)).toFixed(2)}`,
  ];
  if (invoiceRequiresManualSettlement(invoice)) {
    lines.splice(2, 0, `Due date: ${dueLabel}`);
  } else {
    lines.splice(2, 0, `Status: ${getAutomatedInvoiceStatusLabel(invoice)}`);
  }
  if (isTopup) {
    lines.splice(3, 0, `Transfer reference: ${String(invoice?.payment_reference || invoice?.metadata?.topup_reference || "--").trim() || "--"}`);
  }
  if (viewUrl && viewUrl !== "#") {
    lines.push(`View invoice: ${viewUrl}`);
  }
  return lines.join("\n");
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

function buildMonthlyReportsEmailHtml(options = {}) {
  const profitLabel = formatEmailEuroAmount(options?.profitAmount);
  const reportsUrl =
    String(options?.reportsUrl || REPORTS_PORTAL_URL || "").trim() ||
    "https://portal.shipide.com/reports?range=monthly";
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
    String(options?.reportsUrl || REPORTS_PORTAL_URL || "").trim() ||
    "https://portal.shipide.com/reports?range=monthly";
  return [
    `This month, you profited an extra ${profitLabel} with us.`,
    "Your monthly performance report is ready.",
    "Open it instantly with the button below.",
    `View report: ${reportsUrl}`,
  ].join("\n");
}

async function updateBillingInvoiceFields(invoiceId, patch = {}) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId) throw new Error("Invoice id is required.");
  const params = new URLSearchParams();
  params.set("id", `eq.${safeInvoiceId}`);
  const response = await supabaseServiceRequest(
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

async function getSupabaseUserById(userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return null;
  const response = await fetch(`${SUPABASE_URL}/auth/v1/admin/users/${safeUserId}`, {
    method: "GET",
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    },
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Could not load user (${response.status}) ${details}`.trim());
  }
  const payload = await response.json().catch(() => null);
  if (!payload || typeof payload !== "object") return null;
  return payload?.user && typeof payload.user === "object" ? payload.user : payload;
}

async function isTrustedClientBillingPreferenceUpdater(userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return false;
  const user = await getSupabaseUserById(safeUserId);
  return canManageRegistrationInvites(user);
}

async function filterTrustedClientBillingPreferenceRows(rows = []) {
  const trustedUpdaterCache = new Map();
  const trustedRows = [];
  for (const row of Array.isArray(rows) ? rows : []) {
    if (isDefaultClientBillingPreference(row)) {
      trustedRows.push(row);
      continue;
    }
    const updaterId = String(row?.updated_by || "").trim();
    if (!updaterId) continue;
    let trustedUpdater = trustedUpdaterCache.get(updaterId);
    if (trustedUpdater == null) {
      try {
        trustedUpdater = await isTrustedClientBillingPreferenceUpdater(updaterId);
      } catch (_error) {
        trustedUpdater = false;
      }
      trustedUpdaterCache.set(updaterId, trustedUpdater);
    }
    if (trustedUpdater) {
      trustedRows.push(row);
    }
  }
  return trustedRows;
}

function mapInvoiceProfileFromUser(user = {}) {
  const metadata =
    user?.user_metadata && typeof user.user_metadata === "object" ? user.user_metadata : {};
  return {
    company_name: String(metadata.company_name || metadata.companyName || "").trim() || "Client account",
    contact_name: String(metadata.contact_name || metadata.contactName || "").trim() || "",
    contact_email:
      normalizeEmail(metadata.contact_email || metadata.contactEmail || user?.email || "") || "",
    contact_phone: String(metadata.contact_phone || metadata.contactPhone || "").trim() || "",
    billing_address: String(metadata.billing_address || metadata.billingAddress || "").trim() || "",
    tax_id: String(metadata.tax_id || metadata.taxId || "").trim() || "",
    customer_id: String(metadata.customer_id || metadata.customerId || "").trim() || "",
    account_manager: String(metadata.account_manager || metadata.accountManager || "").trim() || "",
  };
}

function buildInvoiceStatsByUser(invoices = []) {
  const grouped = new Map();
  invoices.forEach((invoice) => {
    if (isTopupBillingInvoice(invoice)) return;
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
    ...(function getRowAccountingExport() {
      const metadata =
        invoice?.metadata && typeof invoice.metadata === "object" ? invoice.metadata : {};
      const accountingExport =
        metadata?.accounting_export && typeof metadata.accounting_export === "object"
          ? metadata.accounting_export
          : {};
      const exportedAt = String(accountingExport?.exported_at || "").trim() || null;
      const exportedBy = String(accountingExport?.exported_by || "").trim() || null;
      const exportBatchId = String(accountingExport?.export_batch_id || "").trim() || null;
      const exportFormat =
        String(accountingExport?.export_format || "").trim().toLowerCase() || null;
      return {
        created_at: invoice?.created_at || null,
        accounting_exported: Boolean(exportedAt),
        accounting_exported_at: exportedAt,
        accounting_exported_by: exportedBy,
        accounting_export_batch_id: exportBatchId,
        accounting_export_format: exportFormat,
      };
    })(),
    id: invoice?.id || null,
    user_id: invoice?.user_id || null,
    reference: getBillingInvoiceReference(invoice),
    invoice_kind: getBillingInvoiceKind(invoice),
    source_topup_id: invoice?.source_topup_id || null,
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

function isBillingInvoicesSchemaMissingDetails(details = "") {
  const message = String(details || "");
  return (
    /relation .*billing_invoices/i.test(message)
    || /invoice_kind|source_topup_id|invoice_number/i.test(message)
    || /column .*billing_invoices/i.test(message)
  );
}

async function getClientBillingPreferenceForUser(userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return { ...DEFAULT_CLIENT_BILLING_PREF };
  const params = new URLSearchParams();
  params.set("select", "user_id,invoice_enabled,card_enabled,updated_at,updated_by");
  params.set("user_id", `eq.${safeUserId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
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
  const trustedRows = await filterTrustedClientBillingPreferenceRows(rows);
  if (!trustedRows.length) return { ...DEFAULT_CLIENT_BILLING_PREF };
  return normalizeClientBillingPreference(trustedRows[0]);
}

function getBillingIbanConfig() {
  return {
    beneficiary: BILLING_IBAN_BENEFICIARY || DEFAULT_IBAN_BENEFICIARY,
    iban: BILLING_IBAN || DEFAULT_IBAN,
    bic: BILLING_IBAN_BIC || DEFAULT_IBAN_BIC,
    address: BILLING_IBAN_ADDRESS || DEFAULT_IBAN_ADDRESS,
    note: BILLING_IBAN_NOTE || DEFAULT_IBAN_TRANSFER_NOTE,
  };
}

function parseTopupRequestedAt(row) {
  const requestedAt = Date.parse(String(row?.requested_at || row?.created_at || ""));
  return Number.isFinite(requestedAt) ? requestedAt : null;
}

function getEffectiveTopupStatus(row) {
  const status = normalizeTopupStatus(row?.status);
  if (status !== "pending") return status;
  const requestedAt = parseTopupRequestedAt(row);
  if (requestedAt !== null && Date.now() - requestedAt > BILLING_TOPUP_EXPIRY_MS) {
    return "expired";
  }
  return status;
}

function buildTopupReference(userId = "") {
  const userChunk = String(userId || "")
    .replace(/[^a-z0-9]/gi, "")
    .slice(0, 6)
    .toUpperCase() || "SHIPID";
  const timeChunk = Date.now().toString(36).toUpperCase();
  const randChunk = crypto.randomBytes(2).toString("hex").toUpperCase();
  return `SHIP-${userChunk}-${timeChunk}-${randChunk}`;
}

function buildTopupInstructions(ibanConfig, row, amountCents = null) {
  const rowAmountCents = Number(row?.amount_cents);
  const safeAmountCents = Number.isFinite(rowAmountCents)
    ? Math.max(0, rowAmountCents)
    : Math.max(0, Number(amountCents) || 0);
  return {
    ...ibanConfig,
    amount_eur: safeAmountCents > 0 ? fromCents(safeAmountCents) : null,
    currency: String(row?.currency || DEFAULT_BILLING_CURRENCY).trim() || DEFAULT_BILLING_CURRENCY,
    reference: String(row?.reference || "").trim(),
  };
}

function isReusablePendingTopup(row) {
  if (getEffectiveTopupStatus(row) !== "pending") return false;
  const reference = String(row?.reference || "").trim();
  if (!reference) return false;
  return true;
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

function extractSupabaseErrorMessage(details) {
  const raw = String(details || "").trim();
  if (!raw) return "";
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") {
      return String(parsed.message || parsed.details || parsed.hint || raw).trim();
    }
  } catch (_error) {}
  return raw;
}

function resolveWalletRpcIssue(details) {
  const raw = String(details || "");
  const lower = raw.toLowerCase();
  if (
    lower.includes("apply_billing_wallet_transaction") ||
    lower.includes("schema cache") ||
    lower.includes("\"code\":\"pgrst202\"") ||
    lower.includes("\"code\":\"42883\"")
  ) {
    if (lower.includes("schema cache") || lower.includes("could not find the function") || lower.includes("42883")) {
      return {
        kind: "missing",
        message:
          "Wallet transaction RPC missing or stale. Run supabase_wallet_billing.sql, then run: NOTIFY pgrst, 'reload schema';",
      };
    }
  }
  if (lower.includes("permission denied") || lower.includes("\"code\":\"42501\"")) {
    return {
      kind: "permission",
      message:
        "Wallet transaction RPC access denied. Grant execute to service_role, then run: NOTIFY pgrst, 'reload schema';",
    };
  }
  return null;
}

function resolveWiseStorageIssue(details, tableName) {
  const raw = String(details || "");
  const lower = raw.toLowerCase();
  const table = String(tableName || "").toLowerCase();
  if (!lower || !table || !lower.includes(table)) return null;
  if (lower.includes("schema cache")) {
    return {
      kind: "cache",
      message:
        "Wise reconciliation schema cache is stale. Run: NOTIFY pgrst, 'reload schema'; in Supabase SQL editor.",
    };
  }
  if (
    lower.includes("does not exist") ||
    lower.includes("undefined table") ||
    lower.includes("\"code\":\"42p01\"")
  ) {
    return {
      kind: "missing",
      message:
        "Wise reconciliation schema missing. Run supabase_wise_reconciliation.sql in Supabase SQL editor.",
    };
  }
  if (
    lower.includes("permission denied") ||
    lower.includes("insufficient_privilege") ||
    lower.includes("\"code\":\"42501\"")
  ) {
    return {
      kind: "permission",
      message: `Wise reconciliation access denied for ${tableName}. Grant table privileges (service_role/authenticated), then run: NOTIFY pgrst, 'reload schema';`,
    };
  }
  return null;
}

function getWiseEnvironmentMode() {
  const explicit = String(process.env.WISE_ENVIRONMENT || "").trim().toLowerCase();
  if (explicit === "sandbox" || explicit === "test") return "sandbox";
  if (explicit === "production" || explicit === "prod" || explicit === "live") return "production";
  const apiBase = String(process.env.WISE_API_BASE_URL || "").trim().toLowerCase();
  if (apiBase.includes("sandbox")) return "sandbox";
  return "production";
}

function getWiseApiBaseUrl() {
  const explicit = String(process.env.WISE_API_BASE_URL || "").trim();
  if (explicit) return explicit.replace(/\/+$/, "");
  return getWiseEnvironmentMode() === "sandbox" ? WISE_SANDBOX_API_BASE_URL : WISE_DEFAULT_API_BASE_URL;
}

function getWiseWebhookPublicKeyPem() {
  const explicit = String(process.env.WISE_WEBHOOK_PUBLIC_KEY || "").trim();
  if (explicit) return explicit;
  return getWiseEnvironmentMode() === "sandbox"
    ? WISE_SANDBOX_PUBLIC_KEY
    : WISE_PRODUCTION_PUBLIC_KEY;
}

function getWiseApiToken() {
  return String(process.env.WISE_API_TOKEN || "").trim();
}

function getWiseProfileId() {
  return String(process.env.WISE_PROFILE_ID || "").trim();
}

function getWiseBalanceId() {
  return String(process.env.WISE_BALANCE_ID || "").trim();
}

function getWiseBalanceCurrency() {
  return String(process.env.WISE_BALANCE_CURRENCY || DEFAULT_BILLING_CURRENCY).trim()
    || DEFAULT_BILLING_CURRENCY;
}

function getWiseStatementLookbackDays(defaultValue = WISE_DEFAULT_LOOKBACK_DAYS) {
  const numeric = Number(process.env.WISE_STATEMENT_LOOKBACK_DAYS);
  if (!Number.isFinite(numeric)) return defaultValue;
  return Math.max(1, Math.min(469, Math.round(numeric)));
}

function getWiseConfigSummary() {
  return {
    configured: Boolean(getWiseApiToken() && getWiseProfileId() && getWiseBalanceId()),
    environment: getWiseEnvironmentMode(),
    api_base_url: getWiseApiBaseUrl(),
    profile_id: getWiseProfileId() || null,
    balance_id: getWiseBalanceId() || null,
    currency: getWiseBalanceCurrency(),
  };
}

function assertWiseStatementConfig() {
  if (!getWiseApiToken()) {
    throw new Error("WISE_API_TOKEN is required for Wise reconciliation.");
  }
  if (!getWiseProfileId() || !getWiseBalanceId()) {
    throw new Error("WISE_PROFILE_ID and WISE_BALANCE_ID are required for Wise reconciliation.");
  }
}

function verifyWiseWebhookSignature(rawBody, receivedSignature) {
  const signature = String(receivedSignature || "").trim();
  const publicKey = getWiseWebhookPublicKeyPem();
  if (!signature || !publicKey) return false;
  try {
    const verifier = crypto.createVerify("RSA-SHA256");
    verifier.update(Buffer.isBuffer(rawBody) ? rawBody : Buffer.from(String(rawBody || ""), "utf8"));
    verifier.end();
    return verifier.verify(publicKey, signature, "base64");
  } catch (_error) {
    return false;
  }
}

function normalizeTopupStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["pending", "received", "credited", "cancelled", "failed"].includes(normalized)) {
    return normalized;
  }
  return "pending";
}

function mapBillingTopupRow(row) {
  const status = getEffectiveTopupStatus(row);
  const amountCents = Math.max(0, toCents(fromCents(row?.amount_cents)));
  const metadata = row?.metadata && typeof row.metadata === "object" ? row.metadata : {};
  return {
    id: String(row?.id || "").trim(),
    amount_eur: fromCents(amountCents),
    currency: String(row?.currency || DEFAULT_BILLING_CURRENCY).trim() || DEFAULT_BILLING_CURRENCY,
    reference: String(row?.reference || "").trim(),
    status,
    requested_at: row?.requested_at || row?.created_at || null,
    received_at: row?.received_at || null,
    credited_at: row?.credited_at || null,
    invoice_id: String(metadata?.invoice_id || "").trim() || null,
    invoice_reference:
      String(metadata?.invoice_reference || metadata?.invoice_number || "").trim() || null,
    metadata,
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

async function getOrCreateBillingWallet(userId) {
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

  const insertResponse = await supabaseServiceRequest(`/rest/v1/${BILLING_WALLET_TABLE}`, {
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
  });
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
  { userId = "", statuses = [], limit = 30, allowMissing = false } = {}
) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return [];
  const safeLimit = Math.max(1, Math.min(200, Number(limit) || 30));
  const params = new URLSearchParams();
  params.set("select", BILLING_TOPUP_SELECT_FIELDS);
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

async function listRecentCreditedTopupsForInvoiceBackfill(
  { limit = 40, allowMissing = false } = {}
) {
  const safeLimit = Math.max(1, Math.min(200, Number(limit) || 40));
  const params = new URLSearchParams();
  params.set("select", BILLING_TOPUP_SELECT_FIELDS);
  params.set("status", "eq.credited");
  params.set("order", "credited_at.desc.nullslast,requested_at.desc.nullslast,created_at.desc");
  params.set("limit", String(safeLimit));
  const response = await supabaseServiceRequest(
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
    throw new Error(`Could not load credited bank top-ups (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function runCreditedTopupInvoiceBackfill(
  { limit = 1, publicAppUrl = "", logger = console } = {}
) {
  const topups = await listRecentCreditedTopupsForInvoiceBackfill({
    limit,
    allowMissing: true,
  }).then((rows) => (Array.isArray(rows) ? rows : []));
  const summary = {
    scanned: topups.length,
    attempted: 0,
    sent: 0,
    skipped: 0,
    failed: 0,
    failures: [],
  };
  for (const topup of topups) {
    const topupId = String(topup?.id || "").trim();
    if (!topupId) continue;
    summary.attempted += 1;
    try {
      const result = await ensureBillingTopupInvoiceAndSend(topupId, {
        publicAppUrl,
        ensurePdf: true,
      });
      if (result?.skipped) {
        summary.skipped += 1;
      } else {
        summary.sent += 1;
      }
    } catch (error) {
      summary.failed += 1;
      summary.failures.push({
        topup_id: topupId,
        reference: String(topup?.reference || "").trim() || null,
        error: error?.message || "Top-up invoice backfill failed.",
      });
      logger?.error?.("[topup invoice backfill]", topupId, error);
    }
  }
  return summary;
}

async function listBillingWalletTransactions({ userId = "", limit = 30, allowMissing = false } = {}) {
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

async function createBillingTopupRequest(user, amountEur = null) {
  const userId = String(user?.id || "").trim();
  if (!userId) {
    throw new Error("Authentication required.");
  }
  const numericAmount = Number(amountEur);
  const amountCents =
    Number.isFinite(numericAmount) && numericAmount > 0 ? Math.max(0, toCents(numericAmount)) : 0;
  await getOrCreateBillingWallet(userId);
  const ibanConfig = getBillingIbanConfig();
  const recentPending = await listBillingTopups({
    userId,
    statuses: ["pending"],
    limit: 1,
    allowMissing: true,
  });
  const reusable = Array.isArray(recentPending) ? recentPending.find(isReusablePendingTopup) : null;
  if (reusable) {
    return {
      topup: mapBillingTopupRow(reusable),
      instructions: buildTopupInstructions(ibanConfig, reusable, amountCents),
    };
  }

  let created = null;
  for (let attempt = 0; attempt < 4; attempt += 1) {
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
        address: ibanConfig.address,
      },
    };
    const response = await supabaseServiceRequest(`/rest/v1/${BILLING_WALLET_TOPUPS_TABLE}`, {
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
      const message = `Could not create top-up request (${response.status}) ${details}`.trim();
      if (message.includes("23505") || message.includes("duplicate key")) {
        continue;
      }
      throw new Error(message);
    }
    const rows = await response.json().catch(() => []);
    created = Array.isArray(rows) && rows.length ? rows[0] : payload;
    break;
  }
  if (!created) {
    throw new Error("Could not create top-up request. Please try again.");
  }
  return {
    topup: mapBillingTopupRow(created),
    instructions: buildTopupInstructions(ibanConfig, created, amountCents),
  };
}

async function updateBillingTopupFields(topupId, patch = {}) {
  const safeTopupId = String(topupId || "").trim();
  if (!safeTopupId) throw new Error("Top-up id is required.");
  const params = new URLSearchParams();
  params.set("id", `eq.${safeTopupId}`);
  const response = await supabaseServiceRequest(
    `/rest/v1/${BILLING_WALLET_TOPUPS_TABLE}?${params.toString()}`,
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
    const walletIssue = resolveWalletAccessIssue(details, BILLING_WALLET_TOPUPS_TABLE);
    if (walletIssue) throw new Error(walletIssue.message);
    throw new Error(`Could not update top-up (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

function mapBillingWalletRpcResult(row, fallback = {}) {
  const safeUserId = String(row?.wallet_user_id || fallback.userId || "").trim();
  const balanceCents = Number(row?.wallet_balance_cents);
  const amountCents = Number(fallback.amountCents) || 0;
  const transactionReference = String(row?.transaction_reference || fallback.reference || "").trim();
  return {
    wallet: {
      user_id: safeUserId,
      balance_cents: Number.isFinite(balanceCents) ? balanceCents : 0,
      currency: String(row?.wallet_currency || fallback.currency || DEFAULT_BILLING_CURRENCY).trim() || DEFAULT_BILLING_CURRENCY,
      updated_at: row?.wallet_updated_at || null,
    },
    transaction: {
      id: row?.transaction_id || null,
      user_id: safeUserId,
      source: String(fallback.source || "").trim(),
      amount_cents: amountCents,
      balance_after_cents: Number.isFinite(balanceCents) ? balanceCents : 0,
      reference: transactionReference || null,
      metadata: fallback.metadata && typeof fallback.metadata === "object" ? fallback.metadata : {},
    },
    transaction_reference: transactionReference,
  };
}

async function applyBillingWalletTransaction({
  userId,
  amountCents,
  source,
  reference = null,
  metadata = {},
  currency = DEFAULT_BILLING_CURRENCY,
} = {}) {
  const safeUserId = String(userId || "").trim();
  const safeSource = String(source || "").trim();
  const safeAmountCents = Number(amountCents);
  if (!safeUserId) {
    throw new Error("Client id is required.");
  }
  if (!Number.isFinite(safeAmountCents) || safeAmountCents === 0) {
    throw new Error("A non-zero wallet transaction amount is required.");
  }
  if (!safeSource) {
    throw new Error("Wallet transaction source is required.");
  }

  const response = await supabaseServiceRequest("/rest/v1/rpc/apply_billing_wallet_transaction", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      p_user_id: safeUserId,
      p_amount_cents: Math.trunc(safeAmountCents),
      p_source: safeSource,
      p_reference: String(reference || "").trim() || null,
      p_metadata: metadata && typeof metadata === "object" ? metadata : {},
      p_currency: String(currency || DEFAULT_BILLING_CURRENCY).trim() || DEFAULT_BILLING_CURRENCY,
    }),
  });

  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const rpcIssue = resolveWalletRpcIssue(details);
    if (rpcIssue) throw new Error(rpcIssue.message);
    throw new Error(extractSupabaseErrorMessage(details) || `Could not apply wallet transaction (${response.status}).`);
  }

  const rows = await response.json().catch(() => []);
  const row = Array.isArray(rows) && rows.length ? rows[0] : rows && typeof rows === "object" ? rows : null;
  if (!row) {
    throw new Error("Wallet transaction did not return a result.");
  }
  return mapBillingWalletRpcResult(row, {
    userId: safeUserId,
    amountCents: Math.trunc(safeAmountCents),
    source: safeSource,
    reference,
    metadata,
    currency,
  });
}

async function debitWalletForCheckout(user, amountEur, metadata = {}) {
  const userId = String(user?.id || "").trim();
  if (!userId) {
    throw new Error("Authentication required.");
  }
  const debitCents = toCents(amountEur);
  if (!Number.isFinite(debitCents) || debitCents <= 0) {
    throw new Error("Invalid checkout amount.");
  }
  const reference = buildTopupReference(userId);
  try {
    return await applyBillingWalletTransaction({
      userId,
      amountCents: -debitCents,
      source: "label_checkout",
      reference,
      metadata: {
        checkout: metadata && typeof metadata === "object" ? metadata : {},
        actor: normalizeEmail(user?.email || "") || userId,
      },
      currency: DEFAULT_BILLING_CURRENCY,
    });
  } catch (error) {
    if (/insufficient wallet balance/i.test(String(error?.message || ""))) {
      const wallet = await getOrCreateBillingWallet(userId).catch(() => null);
      const currentBalanceCents = Number(wallet?.balance_cents) || 0;
      const missing = fromCents(Math.max(0, debitCents - currentBalanceCents)).toFixed(2);
      throw new Error(`Insufficient wallet balance. Add at least €${missing} by bank transfer.`);
    }
    throw error;
  }
}

function normalizeGenerationHistoryInput(user, body = {}) {
  const userId = String(user?.id || "").trim();
  if (!userId) {
    throw new Error("Authentication required.");
  }
  const serviceType = String(body?.service_type || body?.serviceType || "").trim();
  if (!serviceType || serviceType.length > 120) {
    throw new Error("A valid service type is required.");
  }
  const quantity = Number(body?.quantity);
  if (!Number.isInteger(quantity) || quantity <= 0 || quantity > 10000) {
    throw new Error("A valid generation quantity is required.");
  }
  const totalPrice = Number(body?.total_price ?? body?.totalPrice);
  if (!Number.isFinite(totalPrice) || totalPrice < 0 || totalPrice > 1000000) {
    throw new Error("A valid generation total is required.");
  }
  const payload = body?.payload;
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new Error("A valid generation payload is required.");
  }
  return {
    user_id: userId,
    service_type: serviceType,
    quantity,
    total_price: Math.round(totalPrice * 100) / 100,
    payload,
  };
}

async function createGenerationHistoryRecord(user, body = {}) {
  const record = normalizeGenerationHistoryInput(user, body);
  const response = await supabaseServiceRequest(`/rest/v1/${HISTORY_TABLE}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify([record]),
  });

  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(
      extractSupabaseErrorMessage(details) || `Could not save generation history (${response.status}).`
    );
  }

  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length
    ? rows[0]
    : {
        id: null,
        created_at: new Date().toISOString(),
        service_type: record.service_type,
        quantity: record.quantity,
        total_price: record.total_price,
        payload: record.payload,
      };
}

async function handleLabelGenerationCreate(req, res) {
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
  try {
    const generation = await createGenerationHistoryRecord(user, body);
    sendJson(res, 200, {
      ok: true,
      generation,
    });
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Could not save generation history." });
  }
}

function mapBillingOverviewInvoiceRow(invoice) {
  if (!invoice || typeof invoice !== "object") return null;
  return {
    id: String(invoice.id || "").trim(),
    status: normalizeInvoiceStatus(invoice.status),
    reference: getBillingInvoiceReference(invoice),
    invoice_kind: getBillingInvoiceKind(invoice),
    issued_at: invoice.issued_at || invoice.created_at || null,
    due_at: invoice.due_at || null,
    total_inc_vat: fromCents(toCents(invoice.total_inc_vat)),
    total_ex_vat: fromCents(toCents(invoice.subtotal_ex_vat)),
    tracking: formatInvoiceTrackingLabel(invoice),
  };
}

function normalizeReferenceForMatch(value) {
  return String(value || "").trim().toUpperCase().replace(/[^A-Z0-9]+/g, "");
}

function normalizeInvoiceReferenceCandidate(value) {
  const match = String(value || "").toUpperCase().match(/\bINV[\s-]*([A-Z0-9]{8})\b/);
  if (!match) return "";
  return `INV-${match[1]}`;
}

function normalizeTopupReferenceCandidate(value) {
  const match = String(value || "").toUpperCase().match(/\bSHIP(?:[\s-]*[A-Z0-9]{2,}){2,}\b/);
  if (!match) return "";
  return match[0].replace(/[^A-Z0-9]+/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
}

function collectNestedStringValues(value, bucket = [], depth = 0) {
  if (bucket.length >= 120 || depth > 6 || value == null) return bucket;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed) bucket.push(trimmed);
    return bucket;
  }
  if (Array.isArray(value)) {
    value.forEach((entry) => {
      if (bucket.length < 120) collectNestedStringValues(entry, bucket, depth + 1);
    });
    return bucket;
  }
  if (typeof value === "object") {
    Object.values(value).forEach((entry) => {
      if (bucket.length < 120) collectNestedStringValues(entry, bucket, depth + 1);
    });
  }
  return bucket;
}

function extractStructuredReferenceCandidates(value) {
  const candidates = new Set();
  collectNestedStringValues(value).forEach((entry) => {
    const invoiceReference = normalizeInvoiceReferenceCandidate(entry);
    if (invoiceReference) candidates.add(invoiceReference);
    const topupReference = normalizeTopupReferenceCandidate(entry);
    if (topupReference) candidates.add(topupReference);
  });
  return Array.from(candidates);
}

function parseWiseMoneyPayload(value, fallbackCurrency = DEFAULT_BILLING_CURRENCY) {
  if (value == null) {
    return {
      amount_cents: 0,
      currency: fallbackCurrency,
    };
  }
  if (typeof value === "number" || typeof value === "string") {
    const numeric = Number(value);
    return {
      amount_cents: Number.isFinite(numeric) ? Math.max(0, Math.round(numeric * 100)) : 0,
      currency: fallbackCurrency,
    };
  }
  if (typeof value === "object") {
    const currency = String(
      value.currency || value.currencyCode || value.currency_code || fallbackCurrency
    ).trim().toUpperCase() || fallbackCurrency;
    const numeric =
      Number(value.value)
      || Number(value.amount)
      || Number(value.total)
      || Number(value.amountValue)
      || 0;
    return {
      amount_cents: Number.isFinite(numeric) ? Math.max(0, Math.round(numeric * 100)) : 0,
      currency,
    };
  }
  return {
    amount_cents: 0,
    currency: fallbackCurrency,
  };
}

function pickFirstNonEmpty(...values) {
  for (const value of values) {
    const text = String(value || "").trim();
    if (text) return text;
  }
  return "";
}

function parseWiseTimestamp(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const parsed = Date.parse(raw);
  if (!Number.isFinite(parsed)) return null;
  return new Date(parsed).toISOString();
}

function buildWiseReceiptExternalKey(prefix, payload = {}) {
  const rawKey = [
    String(prefix || "wise").trim().toLowerCase() || "wise",
    String(payload?.reference_number || "").trim(),
    String(payload?.payment_reference || "").trim(),
    String(payload?.occurred_at || "").trim(),
    String(payload?.amount_cents || "").trim(),
    String(payload?.currency || "").trim(),
    String(payload?.event_type || "").trim(),
  ]
    .filter(Boolean)
    .join("|")
    || JSON.stringify(payload || {});
  return `wise:${crypto.createHash("sha256").update(rawKey).digest("hex")}`;
}

function extractWiseReceiptFromStatementTransaction(transaction = {}) {
  if (!transaction || typeof transaction !== "object") return null;
  const money = parseWiseMoneyPayload(
    transaction.amount
      || transaction.transactionAmount
      || transaction.totalAmount
      || transaction.details?.amount
      || transaction.value,
    getWiseBalanceCurrency()
  );
  if (!money.amount_cents) return null;
  const referenceCandidates = extractStructuredReferenceCandidates(transaction);
  const paymentReference = pickFirstNonEmpty(
    transaction?.details?.paymentReference,
    transaction?.details?.reference,
    transaction?.details?.publicReference,
    transaction?.details?.message,
    transaction?.reference,
    referenceCandidates[0]
  );
  const referenceNumber = pickFirstNonEmpty(
    transaction?.referenceNumber,
    transaction?.transactionNumber,
    transaction?.transferId,
    transaction?.id
  );
  const occurredAt = parseWiseTimestamp(
    transaction?.date
      || transaction?.bookingDate
      || transaction?.createdOn
      || transaction?.created_at
      || transaction?.occurredAt
      || transaction?.details?.createdOn
  );
  const senderName = pickFirstNonEmpty(
    transaction?.details?.senderName,
    transaction?.details?.counterpartyName,
    transaction?.details?.merchantName,
    transaction?.senderName,
    transaction?.counterpartyName
  );
  const senderAccount = pickFirstNonEmpty(
    transaction?.details?.senderAccount,
    transaction?.details?.senderIban,
    transaction?.details?.counterpartyAccount,
    transaction?.senderAccount
  );
  const senderBank = pickFirstNonEmpty(
    transaction?.details?.senderBankName,
    transaction?.details?.bankName,
    transaction?.details?.counterpartyBankName,
    transaction?.senderBank
  );
  return {
    provider: "wise",
    source: "statement",
    external_key: buildWiseReceiptExternalKey("statement", {
      reference_number: referenceNumber,
      payment_reference: paymentReference,
      occurred_at: occurredAt,
      amount_cents: money.amount_cents,
      currency: money.currency,
      event_type: pickFirstNonEmpty(transaction?.type, transaction?.details?.type),
    }),
    event_type: pickFirstNonEmpty(transaction?.type, transaction?.details?.type) || null,
    amount_cents: money.amount_cents,
    currency: money.currency,
    payment_reference: paymentReference || null,
    reference_number: referenceNumber || null,
    sender_name: senderName || null,
    sender_account: senderAccount || null,
    sender_bank: senderBank || null,
    occurred_at: occurredAt,
    metadata: {
      extractor: "wise_statement",
      reference_candidates: referenceCandidates,
    },
    raw_payload: transaction,
  };
}

function extractWiseReceiptCandidatesFromWebhookPayload(payload = {}) {
  const eventType = pickFirstNonEmpty(
    payload?.event_type,
    payload?.eventType,
    payload?.type,
    payload?.name
  );
  const candidateObjects = [
    payload?.data,
    payload?.data?.resource,
    payload?.data?.action,
    payload?.resource,
    payload,
  ].filter((entry) => entry && typeof entry === "object");
  const receipts = [];
  candidateObjects.forEach((entry) => {
    const money = parseWiseMoneyPayload(
      entry?.amount || entry?.money || entry?.value || entry?.details?.amount,
      getWiseBalanceCurrency()
    );
    if (!money.amount_cents) return;
    const referenceCandidates = extractStructuredReferenceCandidates(entry);
    const paymentReference = pickFirstNonEmpty(
      entry?.paymentReference,
      entry?.reference,
      entry?.publicReference,
      entry?.details?.paymentReference,
      entry?.details?.reference,
      referenceCandidates[0]
    );
    const referenceNumber = pickFirstNonEmpty(
      entry?.referenceNumber,
      entry?.transactionNumber,
      entry?.transferId,
      entry?.id,
      entry?.details?.id
    );
    const occurredAt = parseWiseTimestamp(
      entry?.occurredAt
        || entry?.createdOn
        || entry?.created_at
        || entry?.completedAt
        || payload?.sent_at
    );
    receipts.push({
      provider: "wise",
      source: "webhook",
      external_key: buildWiseReceiptExternalKey("webhook", {
        reference_number: referenceNumber,
        payment_reference: paymentReference,
        occurred_at: occurredAt,
        amount_cents: money.amount_cents,
        currency: money.currency,
        event_type: eventType,
      }),
      event_type: eventType || null,
      amount_cents: money.amount_cents,
      currency: money.currency,
      payment_reference: paymentReference || null,
      reference_number: referenceNumber || null,
      sender_name: pickFirstNonEmpty(
        entry?.senderName,
        entry?.details?.senderName,
        entry?.counterpartyName
      ) || null,
      sender_account: pickFirstNonEmpty(
        entry?.senderAccount,
        entry?.details?.senderAccount,
        entry?.details?.senderIban
      ) || null,
      sender_bank: pickFirstNonEmpty(
        entry?.senderBank,
        entry?.details?.senderBankName,
        entry?.details?.bankName
      ) || null,
      occurred_at: occurredAt,
      metadata: {
        extractor: "wise_webhook",
        reference_candidates: referenceCandidates,
      },
      raw_payload: payload,
    });
  });
  const unique = new Map();
  receipts.forEach((entry) => {
    if (!unique.has(entry.external_key)) {
      unique.set(entry.external_key, entry);
    }
  });
  return Array.from(unique.values());
}

function getBillingBankReceiptSelectColumns() {
  return [
    "id",
    "provider",
    "source",
    "external_key",
    "delivery_id",
    "event_type",
    "status",
    "amount_cents",
    "currency",
    "payment_reference",
    "reference_number",
    "sender_name",
    "sender_account",
    "sender_bank",
    "occurred_at",
    "matched_user_id",
    "matched_topup_id",
    "matched_invoice_id",
    "match_reason",
    "review_note",
    "credited_at",
    "metadata",
    "raw_payload",
    "created_at",
    "updated_at",
  ].join(",");
}

function mapBillingBankReceiptRow(row) {
  if (!row || typeof row !== "object") return null;
  const metadata = row?.metadata && typeof row.metadata === "object" ? row.metadata : {};
  return {
    id: String(row?.id || "").trim(),
    provider: String(row?.provider || "wise").trim() || "wise",
    source: String(row?.source || "statement").trim() || "statement",
    status: String(row?.status || "pending").trim().toLowerCase() || "pending",
    amount_cents: Math.max(0, Number(row?.amount_cents) || 0),
    amount_eur: fromCents(Math.max(0, Number(row?.amount_cents) || 0)),
    currency: String(row?.currency || DEFAULT_BILLING_CURRENCY).trim() || DEFAULT_BILLING_CURRENCY,
    payment_reference: String(row?.payment_reference || "").trim(),
    reference_number: String(row?.reference_number || "").trim(),
    sender_name: String(row?.sender_name || "").trim(),
    sender_account: String(row?.sender_account || "").trim(),
    sender_bank: String(row?.sender_bank || "").trim(),
    occurred_at: row?.occurred_at || null,
    matched_user_id: String(row?.matched_user_id || "").trim() || null,
    matched_topup_id: String(row?.matched_topup_id || "").trim() || null,
    matched_invoice_id: String(row?.matched_invoice_id || "").trim() || null,
    match_reason: String(row?.match_reason || "").trim(),
    review_note: String(row?.review_note || "").trim(),
    credited_at: row?.credited_at || null,
    created_at: row?.created_at || null,
    updated_at: row?.updated_at || null,
    suggested_topup_id: String(metadata?.suggested_topup_id || "").trim() || null,
    suggested_topup_reference: String(metadata?.suggested_topup_reference || "").trim() || null,
    suggested_invoice_id: String(metadata?.suggested_invoice_id || "").trim() || null,
    suggested_invoice_reference: String(metadata?.suggested_invoice_reference || "").trim() || null,
  };
}

async function getBillingBankReceiptById(receiptId, { allowMissing = false } = {}) {
  const safeReceiptId = String(receiptId || "").trim();
  if (!safeReceiptId) return null;
  const params = new URLSearchParams();
  params.set("select", getBillingBankReceiptSelectColumns());
  params.set("id", `eq.${safeReceiptId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${BILLING_BANK_RECEIPTS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const issue = resolveWiseStorageIssue(details, BILLING_BANK_RECEIPTS_TABLE);
    if (issue) {
      if (allowMissing && issue.kind === "missing") return null;
      throw new Error(issue.message);
    }
    throw new Error(`Could not load bank receipt (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function getBillingBankReceiptByExternalKey(externalKey, { allowMissing = false } = {}) {
  const safeExternalKey = String(externalKey || "").trim();
  if (!safeExternalKey) return null;
  const params = new URLSearchParams();
  params.set("select", getBillingBankReceiptSelectColumns());
  params.set("external_key", `eq.${safeExternalKey}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${BILLING_BANK_RECEIPTS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const issue = resolveWiseStorageIssue(details, BILLING_BANK_RECEIPTS_TABLE);
    if (issue) {
      if (allowMissing && issue.kind === "missing") return null;
      throw new Error(issue.message);
    }
    throw new Error(`Could not load bank receipt (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function listBillingBankReceipts({ limit = 80, statuses = [], allowMissing = false } = {}) {
  const safeLimit = Math.max(1, Math.min(300, Number(limit) || 80));
  const params = new URLSearchParams();
  params.set("select", getBillingBankReceiptSelectColumns());
  params.set("order", "occurred_at.desc.nullslast,created_at.desc");
  params.set("limit", String(safeLimit));
  if (Array.isArray(statuses) && statuses.length) {
    const normalized = statuses
      .map((entry) => String(entry || "").trim().toLowerCase())
      .filter(Boolean);
    if (normalized.length) {
      params.set("status", `in.(${normalized.join(",")})`);
    }
  }
  const response = await supabaseServiceRequest(
    `/rest/v1/${BILLING_BANK_RECEIPTS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const issue = resolveWiseStorageIssue(details, BILLING_BANK_RECEIPTS_TABLE);
    if (issue) {
      if (allowMissing && issue.kind === "missing") return [];
      throw new Error(issue.message);
    }
    throw new Error(`Could not load bank receipts (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function upsertBillingBankReceipt(candidate = {}) {
  const existing = await getBillingBankReceiptByExternalKey(candidate?.external_key, {
    allowMissing: true,
  });
  const existingMetadata =
    existing?.metadata && typeof existing.metadata === "object" ? existing.metadata : {};
  const nextMetadata =
    candidate?.metadata && typeof candidate.metadata === "object" ? candidate.metadata : {};
  const payload = {
    provider: "wise",
    source: String(candidate?.source || existing?.source || "statement").trim() || "statement",
    external_key: String(candidate?.external_key || existing?.external_key || "").trim(),
    delivery_id: String(candidate?.delivery_id || existing?.delivery_id || "").trim() || null,
    event_type: String(candidate?.event_type || existing?.event_type || "").trim() || null,
    status: String(existing?.status || candidate?.status || "pending").trim().toLowerCase() || "pending",
    amount_cents: Math.max(
      0,
      Number(candidate?.amount_cents ?? existing?.amount_cents) || 0
    ),
    currency:
      String(candidate?.currency || existing?.currency || DEFAULT_BILLING_CURRENCY).trim()
      || DEFAULT_BILLING_CURRENCY,
    payment_reference:
      String(candidate?.payment_reference || existing?.payment_reference || "").trim() || null,
    reference_number:
      String(candidate?.reference_number || existing?.reference_number || "").trim() || null,
    sender_name: String(candidate?.sender_name || existing?.sender_name || "").trim() || null,
    sender_account:
      String(candidate?.sender_account || existing?.sender_account || "").trim() || null,
    sender_bank: String(candidate?.sender_bank || existing?.sender_bank || "").trim() || null,
    occurred_at: candidate?.occurred_at || existing?.occurred_at || null,
    matched_user_id: existing?.matched_user_id || null,
    matched_topup_id: existing?.matched_topup_id || null,
    matched_invoice_id: existing?.matched_invoice_id || null,
    match_reason: String(existing?.match_reason || candidate?.match_reason || "").trim() || null,
    review_note: String(existing?.review_note || candidate?.review_note || "").trim() || null,
    credited_at: existing?.credited_at || candidate?.credited_at || null,
    metadata: {
      ...existingMetadata,
      ...nextMetadata,
    },
    raw_payload:
      candidate?.raw_payload && typeof candidate.raw_payload === "object"
        ? candidate.raw_payload
        : existing?.raw_payload && typeof existing.raw_payload === "object"
          ? existing.raw_payload
          : {},
    updated_at: new Date().toISOString(),
  };
  if (existing?.id) {
    payload.id = existing.id;
    payload.created_at = existing.created_at || undefined;
  }
  const response = await supabaseServiceRequest(
    `/rest/v1/${BILLING_BANK_RECEIPTS_TABLE}?on_conflict=external_key`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=representation",
      },
      body: JSON.stringify([payload]),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const issue = resolveWiseStorageIssue(details, BILLING_BANK_RECEIPTS_TABLE);
    if (issue) throw new Error(issue.message);
    throw new Error(`Could not save bank receipt (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : payload;
}

async function updateBillingBankReceiptRow(receipt, patch = {}) {
  const current = receipt?.id ? receipt : await getBillingBankReceiptById(receipt, { allowMissing: false });
  if (!current?.id) {
    throw new Error("Bank receipt not found.");
  }
  return upsertBillingBankReceipt({
    ...current,
    ...patch,
    external_key: current.external_key,
    metadata: {
      ...(
        current?.metadata && typeof current.metadata === "object" ? current.metadata : {}
      ),
      ...(
        patch?.metadata && typeof patch.metadata === "object" ? patch.metadata : {}
      ),
    },
    raw_payload:
      patch?.raw_payload && typeof patch.raw_payload === "object"
        ? patch.raw_payload
        : current.raw_payload,
  });
}

async function getWiseWebhookEventByDeliveryId(deliveryId, { allowMissing = false } = {}) {
  const safeDeliveryId = String(deliveryId || "").trim();
  if (!safeDeliveryId) return null;
  const params = new URLSearchParams();
  params.set(
    "select",
    "id,delivery_id,event_type,subscription_id,schema_version,signature_valid,is_test,processing_status,processing_error,payload,created_at,processed_at"
  );
  params.set("delivery_id", `eq.${safeDeliveryId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${WISE_WEBHOOK_EVENTS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const issue = resolveWiseStorageIssue(details, WISE_WEBHOOK_EVENTS_TABLE);
    if (issue) {
      if (allowMissing && issue.kind === "missing") return null;
      throw new Error(issue.message);
    }
    throw new Error(`Could not load Wise webhook event (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function upsertWiseWebhookEvent(payload = {}) {
  const existing = await getWiseWebhookEventByDeliveryId(payload?.delivery_id, { allowMissing: true });
  const record = {
    delivery_id: String(payload?.delivery_id || existing?.delivery_id || "").trim(),
    event_type: String(payload?.event_type || existing?.event_type || "unknown").trim() || "unknown",
    subscription_id: String(payload?.subscription_id || "").trim() || null,
    schema_version: String(payload?.schema_version || "").trim() || null,
    signature_valid: payload?.signature_valid !== false,
    is_test: payload?.is_test === true,
    processing_status:
      String(existing?.processing_status || payload?.processing_status || "received").trim()
      || "received",
    processing_error: String(payload?.processing_error || existing?.processing_error || "").trim() || null,
    payload: payload?.payload && typeof payload.payload === "object" ? payload.payload : {},
    processed_at: payload?.processed_at || existing?.processed_at || null,
  };
  if (existing?.id) {
    record.id = existing.id;
    record.created_at = existing.created_at || undefined;
  }
  const response = await supabaseServiceRequest(
    `/rest/v1/${WISE_WEBHOOK_EVENTS_TABLE}?on_conflict=delivery_id`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=representation",
      },
      body: JSON.stringify([record]),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const issue = resolveWiseStorageIssue(details, WISE_WEBHOOK_EVENTS_TABLE);
    if (issue) throw new Error(issue.message);
    throw new Error(`Could not save Wise webhook event (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : record;
}

async function updateWiseWebhookEventStatus(deliveryId, status, processingError = "") {
  const existing = await getWiseWebhookEventByDeliveryId(deliveryId, { allowMissing: false });
  if (!existing?.delivery_id) {
    throw new Error("Wise webhook delivery not found.");
  }
  return upsertWiseWebhookEvent({
    ...existing,
    delivery_id: existing.delivery_id,
    event_type: existing.event_type,
    processing_status: status,
    processing_error: String(processingError || "").trim() || null,
    processed_at:
      status === "processed" || status === "ignored" || status === "failed"
        ? new Date().toISOString()
        : existing.processed_at || null,
  });
}

async function getBillingTopupByReference(reference, { allowMissing = false } = {}) {
  const safeReference = String(reference || "").trim().toUpperCase();
  if (!safeReference) return null;
  const params = new URLSearchParams();
  params.set("select", BILLING_TOPUP_SELECT_FIELDS);
  params.set("reference", `eq.${safeReference}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${BILLING_WALLET_TOPUPS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const issue = resolveWalletAccessIssue(details, BILLING_WALLET_TOPUPS_TABLE);
    if (issue) {
      if (allowMissing && issue.kind === "missing") return null;
      throw new Error(issue.message);
    }
    throw new Error(`Could not load wallet top-up (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function getBillingTopupById(topupId, { allowMissing = false } = {}) {
  const safeTopupId = String(topupId || "").trim();
  if (!safeTopupId) return null;
  const params = new URLSearchParams();
  params.set("select", BILLING_TOPUP_SELECT_FIELDS);
  params.set("id", `eq.${safeTopupId}`);
  params.set("limit", "1");
  const response = await supabaseServiceRequest(
    `/rest/v1/${BILLING_WALLET_TOPUPS_TABLE}?${params.toString()}`,
    { method: "GET" }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const issue = resolveWalletAccessIssue(details, BILLING_WALLET_TOPUPS_TABLE);
    if (issue) {
      if (allowMissing && issue.kind === "missing") return null;
      throw new Error(issue.message);
    }
    throw new Error(`Could not load wallet top-up (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

function buildInvoiceProfileSnapshotMetadata(profile = {}) {
  return {
    companyName: String(profile?.company_name || "").trim(),
    contactName: String(profile?.contact_name || "").trim(),
    contactEmail: normalizeEmail(profile?.contact_email || ""),
    contactPhone: String(profile?.contact_phone || "").trim(),
    billingAddress: String(profile?.billing_address || "").trim(),
    taxId: String(profile?.tax_id || "").trim(),
    customerId: String(profile?.customer_id || "").trim(),
    accountManager: String(profile?.account_manager || "").trim(),
  };
}

function buildTopupInvoiceItems(topup = {}) {
  const amount = fromCents(Math.max(0, Number(topup?.amount_cents) || 0));
  return [
    {
      service_type: "Prepaid Credit for Shipping Labels",
      quantity: 1,
      amount_ex_vat: amount,
      vat_amount: 0,
      amount_inc_vat: amount,
      sort_index: 0,
      metadata: {
        invoice_kind: "topup",
        topup_reference: String(topup?.reference || "").trim() || null,
      },
    },
  ];
}

function buildTopupInvoicePayload(topup = {}, user = null, options = {}) {
  const issuedAt =
    parseIsoTimestamp(
      topup?.credited_at || topup?.received_at || topup?.requested_at || topup?.created_at || ""
    ) || new Date().toISOString();
  const amount = fromCents(Math.max(0, Number(topup?.amount_cents) || 0));
  const profile = mapInvoiceProfileFromUser(user);
  const invoiceNumber =
    String(options?.invoiceNumber || "").trim()
    || buildTopupInvoiceNumber(topup);
  return {
    user_id: String(topup?.user_id || "").trim() || null,
    invoice_kind: "topup",
    source_topup_id: String(topup?.id || "").trim() || null,
    invoice_number: invoiceNumber,
    period_start: issuedAt.slice(0, 10),
    period_end: issuedAt.slice(0, 10),
    due_at: issuedAt,
    issued_at: issuedAt,
    sent_at: null,
    paid_at: issuedAt,
    status: "paid",
    payment_mode: "wallet",
    currency: String(topup?.currency || DEFAULT_BILLING_CURRENCY).trim() || DEFAULT_BILLING_CURRENCY,
    vat_rate: 0,
    subtotal_ex_vat: amount,
    vat_amount: 0,
    total_inc_vat: amount,
    labels_count: 1,
    line_count: 1,
    company_name: profile.company_name,
    contact_name: profile.contact_name,
    contact_email: profile.contact_email,
    customer_id: profile.customer_id,
    account_manager: profile.account_manager,
    payment_reference: String(topup?.reference || "").trim() || null,
    payment_received_amount: amount,
    metadata: {
      invoice_kind: "topup",
      invoice_number: invoiceNumber,
      source_topup_id: String(topup?.id || "").trim() || null,
      topup_reference: String(topup?.reference || "").trim() || null,
      invoice_profile: buildInvoiceProfileSnapshotMetadata(profile),
    },
  };
}

async function upsertBillingTopupInvoice(topup, user = null) {
  if (!topup?.id) return null;
  const existing = await getBillingInvoiceBySourceTopupId(topup.id, {
    withItems: false,
    allowMissing: true,
  });
  const fixedInvoiceNumber = String(
    existing?.invoice_number || existing?.metadata?.invoice_number || ""
  ).trim();
  let invoice = existing;
  let createdInvoice = false;
  let lastError = null;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const payload = buildTopupInvoicePayload(topup, user, {
      invoiceNumber: fixedInvoiceNumber || buildBillingDocumentPublicCode("topup", "invoice"),
    });
    try {
      if (!invoice?.id) {
        invoice = await createBillingTopupInvoice(payload);
        createdInvoice = Boolean(invoice?.id);
      } else {
        invoice = await updateBillingInvoiceFields(invoice.id, {
          invoice_kind: "topup",
          source_topup_id: topup.id,
          invoice_number:
            String(invoice?.invoice_number || invoice?.metadata?.invoice_number || payload.invoice_number || "").trim()
            || null,
          issued_at: invoice?.issued_at || payload.issued_at,
          due_at: invoice?.due_at || payload.due_at,
          paid_at: invoice?.paid_at || payload.paid_at,
          status: invoice?.status || payload.status,
          payment_mode: "wallet",
          currency: payload.currency,
          vat_rate: 0,
          subtotal_ex_vat: payload.subtotal_ex_vat,
          vat_amount: 0,
          total_inc_vat: payload.total_inc_vat,
          labels_count: 1,
          line_count: 1,
          company_name: payload.company_name,
          contact_name: payload.contact_name,
          contact_email: payload.contact_email,
          customer_id: payload.customer_id,
          account_manager: payload.account_manager,
          payment_reference: payload.payment_reference,
          payment_received_amount: payload.payment_received_amount,
          metadata: {
            ...(invoice?.metadata && typeof invoice.metadata === "object" ? invoice.metadata : {}),
            ...(payload.metadata && typeof payload.metadata === "object" ? payload.metadata : {}),
          },
        });
      }
      break;
    } catch (error) {
      if (fixedInvoiceNumber || !isPublicDocumentCodeCollisionError(error)) {
        throw error;
      }
      lastError = error;
      invoice = existing;
    }
  }
  if (!invoice?.id && lastError) {
    throw lastError;
  }
  if (createdInvoice && invoice?.id) {
    await insertBillingInvoiceEvent(invoice.id, {
      event_type: "created",
      actor: "system",
      channel: "wallet",
      message: "Top-up invoice created from credited wallet top-up.",
      metadata: {
        topup_id: topup.id,
        topup_reference: topup.reference || null,
      },
    }).catch(() => {});
  }
  if (!invoice?.id) {
    invoice = await getBillingInvoiceBySourceTopupId(topup.id, {
      withItems: false,
      allowMissing: true,
    });
  }
  if (!invoice?.id) return null;
  await replaceBillingInvoiceItems(invoice.id, buildTopupInvoiceItems(topup));
  return getBillingInvoiceById(invoice.id, { withItems: true });
}

async function linkBillingTopupInvoice(topup, invoice) {
  if (!topup?.id || !invoice?.id) return topup || null;
  const currentMetadata = topup?.metadata && typeof topup.metadata === "object" ? topup.metadata : {};
  const desiredMetadata = {
    ...currentMetadata,
    invoice_id: invoice.id,
    invoice_reference: getBillingInvoiceReference(invoice),
    invoice_number: String(invoice?.invoice_number || currentMetadata?.invoice_number || "").trim() || null,
    invoice_kind: "topup",
  };
  const metadataChanged =
    String(currentMetadata?.invoice_id || "").trim() !== String(desiredMetadata.invoice_id || "").trim()
    || String(currentMetadata?.invoice_reference || "").trim() !== String(desiredMetadata.invoice_reference || "").trim()
    || String(currentMetadata?.invoice_number || "").trim() !== String(desiredMetadata.invoice_number || "").trim()
    || String(currentMetadata?.invoice_kind || "").trim() !== "topup";
  if (!metadataChanged) return topup;
  return updateBillingTopupFields(topup.id, { metadata: desiredMetadata });
}

async function ensureBillingTopupInvoiceAndSend(topupId, options = {}) {
  const topup = await getBillingTopupById(topupId, { allowMissing: true });
  if (!topup?.id) {
    return { skipped: true, reason: "Top-up not found.", invoice: null, topup: null };
  }
  if (getEffectiveTopupStatus(topup) !== "credited") {
    return { skipped: true, reason: "Top-up is not credited yet.", invoice: null, topup };
  }

  let user = null;
  if (topup?.user_id) {
    try {
      user = await getSupabaseUserById(topup.user_id);
    } catch (_error) {
      user = null;
    }
  }

  const invoice = await upsertBillingTopupInvoice(topup, user);
  if (!invoice?.id) {
    throw new Error("Could not create top-up invoice.");
  }
  const linkedTopup = await linkBillingTopupInvoice(topup, invoice).catch(() => topup);
  const shouldEnsurePdf = options?.ensurePdf !== false;
  const alreadySent =
    Boolean(String(invoice?.sent_at || "").trim())
    || Boolean(String(invoice?.email_message_id || "").trim());
  if (alreadySent) {
    let finalInvoice =
      await getBillingInvoiceById(invoice.id, { withItems: true }).catch(() => invoice)
      || invoice;
    if (shouldEnsurePdf) {
      await getApprovedBillingInvoicePdfExport(finalInvoice.id, { reminderStage: 0 });
      finalInvoice =
        await getBillingInvoiceById(invoice.id, { withItems: true }).catch(() => finalInvoice)
        || finalInvoice;
    }
    const finalTopup = await linkBillingTopupInvoice(linkedTopup || topup, finalInvoice).catch(
      () => linkedTopup || topup
    );
    return { skipped: true, reason: "Top-up invoice already sent.", invoice: finalInvoice, topup: finalTopup };
  }
  const sendResult = await sendBillingInvoiceById(invoice.id, {
    publicAppUrl: options?.publicAppUrl,
  });
  const finalInvoice = sendResult?.invoice || invoice;
  if (shouldEnsurePdf) {
    await getApprovedBillingInvoicePdfExport(finalInvoice.id, { reminderStage: 0 });
  }
  const readyInvoice =
    await getBillingInvoiceById(invoice.id, { withItems: true }).catch(() => finalInvoice)
    || finalInvoice;
  const finalTopup = await linkBillingTopupInvoice(linkedTopup || topup, readyInvoice).catch(
    () => linkedTopup || topup
  );
  return {
    skipped: Boolean(sendResult?.skipped),
    reason: sendResult?.reason || "",
    invoice: readyInvoice,
    topup: finalTopup,
  };
}

async function findBillingInvoiceByReference(reference, { allowMissing = false } = {}) {
  const safeReference = normalizeInvoiceReferenceCandidate(reference);
  if (!safeReference) return null;
  const invoices = await listBillingInvoices({
    limit: 2000,
    statuses: ["draft", "sent", "overdue", "paid"],
    allowMissing,
  });
  const matches = invoices.filter((invoice) => {
    const canonicalInvoiceReference = normalizeReferenceForMatch(toInvoiceReference(invoice?.id));
    const canonicalPaymentReference = normalizeReferenceForMatch(invoice?.payment_reference || "");
    const canonicalTarget = normalizeReferenceForMatch(safeReference);
    return canonicalTarget && (
      canonicalTarget === canonicalInvoiceReference
      || canonicalTarget === canonicalPaymentReference
    );
  });
  if (matches.length !== 1) return null;
  return matches[0];
}

async function applyBillingBankReceiptResolution({
  receiptId = "",
  resolution = "",
  actor = "system",
  topupId = null,
  invoiceId = null,
  note = "",
} = {}) {
  const response = await supabaseServiceRequest(`/rest/v1/rpc/${WISE_APPLY_RECEIPT_RPC}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      p_receipt_id: String(receiptId || "").trim(),
      p_resolution: String(resolution || "").trim().toLowerCase(),
      p_actor: String(actor || "system").trim() || "system",
      p_topup_id: topupId ? String(topupId || "").trim() : null,
      p_invoice_id: invoiceId ? String(invoiceId || "").trim() : null,
      p_note: String(note || "").trim() || null,
    }),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (
      details.includes(WISE_APPLY_RECEIPT_RPC)
      || details.includes("\"code\":\"42883\"")
      || /function .*apply_billing_bank_receipt_resolution/i.test(details)
    ) {
      throw new Error(
        "Wise reconciliation schema missing. Run supabase_wise_reconciliation.sql in Supabase SQL editor."
      );
    }
    throw new Error(`Could not resolve bank receipt (${response.status}) ${details}`.trim());
  }
  const payload = await response.json().catch(() => null);
  if (Array.isArray(payload)) {
    return payload[0] || null;
  }
  return payload && typeof payload === "object" ? payload : null;
}

async function resolveBillingBankReceiptTarget(receipt, resolution, target = "") {
  const action = String(resolution || "").trim().toLowerCase();
  const metadata = receipt?.metadata && typeof receipt.metadata === "object" ? receipt.metadata : {};
  const explicitTarget = String(target || "").trim();
  const resolvedTarget = explicitTarget
    || (
      action === "topup"
        ? String(metadata?.suggested_topup_id || metadata?.suggested_topup_reference || "").trim()
        : String(metadata?.suggested_invoice_id || metadata?.suggested_invoice_reference || "").trim()
    );
  if (!resolvedTarget) return null;
  if (isUuid(resolvedTarget)) {
    return action === "topup"
      ? getBillingTopupById(resolvedTarget, { allowMissing: true })
      : getBillingInvoiceById(resolvedTarget, { withItems: false });
  }
  if (action === "topup") {
    const topupReference = normalizeTopupReferenceCandidate(resolvedTarget);
    return topupReference ? getBillingTopupByReference(topupReference, { allowMissing: true }) : null;
  }
  const invoiceReference = normalizeInvoiceReferenceCandidate(resolvedTarget);
  return invoiceReference ? findBillingInvoiceByReference(invoiceReference, { allowMissing: true }) : null;
}

async function tryAutoResolveBankReceipt(receipt, actor = "wise-auto") {
  const currentReceipt = receipt?.id
    ? await getBillingBankReceiptById(receipt.id, { allowMissing: false })
    : null;
  if (!currentReceipt?.id) {
    return { action: "skipped", receipt: null };
  }
  const currentStatus = String(currentReceipt.status || "").trim().toLowerCase();
  if (["credited", "matched", "ignored"].includes(currentStatus)) {
    return { action: "skipped", receipt: currentReceipt };
  }

  const referenceCandidates = Array.from(
    new Set(
      [
        ...extractStructuredReferenceCandidates(currentReceipt?.payment_reference || ""),
        ...extractStructuredReferenceCandidates(currentReceipt?.reference_number || ""),
        ...extractStructuredReferenceCandidates(currentReceipt?.raw_payload || {}),
        ...(
          currentReceipt?.metadata?.reference_candidates
          && Array.isArray(currentReceipt.metadata.reference_candidates)
            ? currentReceipt.metadata.reference_candidates
            : []
        ),
      ].filter(Boolean)
    )
  );

  const topupMatches = [];
  const invoiceMatches = [];
  for (const candidate of referenceCandidates) {
    const topupReference = normalizeTopupReferenceCandidate(candidate);
    if (topupReference) {
      const topup = await getBillingTopupByReference(topupReference, { allowMissing: true });
      if (topup?.id && !topupMatches.some((entry) => entry?.row?.id === topup.id)) {
        topupMatches.push({ reference: topupReference, row: topup });
      }
    }
    const invoiceReference = normalizeInvoiceReferenceCandidate(candidate);
    if (invoiceReference) {
      const invoice = await findBillingInvoiceByReference(invoiceReference, { allowMissing: true });
      if (invoice?.id && !invoiceMatches.some((entry) => entry?.row?.id === invoice.id)) {
        invoiceMatches.push({ reference: invoiceReference, row: invoice });
      }
    }
  }

  const metadataPatch = {
    reference_candidates: referenceCandidates,
    suggested_topup_id: topupMatches.length === 1 ? topupMatches[0].row.id : null,
    suggested_topup_reference: topupMatches.length === 1 ? topupMatches[0].reference : null,
    suggested_invoice_id: invoiceMatches.length === 1 ? invoiceMatches[0].row.id : null,
    suggested_invoice_reference: invoiceMatches.length === 1 ? invoiceMatches[0].reference : null,
  };

  if (topupMatches.length === 1 && invoiceMatches.length === 0) {
    await updateBillingBankReceiptRow(currentReceipt, {
      status: "pending",
      match_reason: "topup_reference_exact",
      metadata: metadataPatch,
    });
    const resolved = await applyBillingBankReceiptResolution({
      receiptId: currentReceipt.id,
      resolution: "topup",
      actor,
      topupId: topupMatches[0].row.id,
    });
    await ensureBillingTopupInvoiceAndSend(topupMatches[0].row.id, {
      publicAppUrl: getPublicAppUrl(),
    }).catch((error) => {
      console.error("[topup invoice auto-send]", topupMatches[0].row.id, error);
    });
    return { action: "topup", receipt: resolved || currentReceipt };
  }

  if (invoiceMatches.length === 1 && topupMatches.length === 0) {
    const invoice = invoiceMatches[0].row;
    const outstandingCents = Math.max(
      0,
      toCents(invoice?.total_inc_vat) - toCents(invoice?.payment_received_amount)
    );
    if (outstandingCents <= 0) {
      const updated = await updateBillingBankReceiptRow(currentReceipt, {
        status: "manual_review",
        match_reason: "invoice_already_paid_review",
        review_note: "Reference matched an invoice that is already fully paid.",
        metadata: metadataPatch,
      });
      return { action: "manual_review", receipt: updated };
    }
    if (outstandingCents > 0 && Number(currentReceipt?.amount_cents || 0) > outstandingCents) {
      const updated = await updateBillingBankReceiptRow(currentReceipt, {
        status: "manual_review",
        match_reason: "invoice_overpayment_review",
        review_note:
          "Reference matched an invoice, but the received amount is larger than the outstanding balance.",
        metadata: metadataPatch,
      });
      return { action: "manual_review", receipt: updated };
    }
    await updateBillingBankReceiptRow(currentReceipt, {
      status: "pending",
      match_reason:
        outstandingCents > 0 && Number(currentReceipt?.amount_cents || 0) < outstandingCents
          ? "invoice_reference_partial"
          : "invoice_reference_exact",
      metadata: metadataPatch,
    });
    const resolved = await applyBillingBankReceiptResolution({
      receiptId: currentReceipt.id,
      resolution: "invoice",
      actor,
      invoiceId: invoice.id,
    });
    return { action: "invoice", receipt: resolved || currentReceipt };
  }

  const updated = await updateBillingBankReceiptRow(currentReceipt, {
    status: "manual_review",
    match_reason:
      topupMatches.length && invoiceMatches.length
        ? "ambiguous_reference"
        : referenceCandidates.length
          ? "no_exact_match"
          : "missing_reference",
    metadata: metadataPatch,
  });
  return { action: "manual_review", receipt: updated };
}

function extractWiseStatementTransactions(payload) {
  if (Array.isArray(payload?.transactions)) return payload.transactions;
  if (Array.isArray(payload?.statement?.transactions)) return payload.statement.transactions;
  if (Array.isArray(payload?.data?.transactions)) return payload.data.transactions;
  return [];
}

async function fetchWiseBalanceStatement({ lookbackDays = WISE_DEFAULT_LOOKBACK_DAYS } = {}) {
  assertWiseStatementConfig();
  const safeLookbackDays = Math.max(1, Math.min(469, Number(lookbackDays) || WISE_DEFAULT_LOOKBACK_DAYS));
  const intervalEnd = new Date();
  const intervalStart = new Date(Date.now() - safeLookbackDays * 24 * 60 * 60 * 1000);
  const url = new URL(
    `/v1/profiles/${encodeURIComponent(getWiseProfileId())}/balance-statements/${encodeURIComponent(
      getWiseBalanceId()
    )}/statement.json`,
    `${getWiseApiBaseUrl()}/`
  );
  url.searchParams.set("currency", getWiseBalanceCurrency());
  url.searchParams.set("intervalStart", intervalStart.toISOString());
  url.searchParams.set("intervalEnd", intervalEnd.toISOString());
  url.searchParams.set("type", "COMPACT");
  url.searchParams.set("statementLocale", "en");

  const response = await fetch(url.toString(), {
    method: "GET",
    headers: {
      Authorization: `Bearer ${getWiseApiToken()}`,
      "Content-Type": "application/json",
    },
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    if (response.status === 401 || response.status === 403) {
      throw new Error(
        "Wise statement access was denied. Refresh Wise API access or complete the required Strong Customer Authentication in Wise."
      );
    }
    throw new Error(`Could not load Wise balance statement (${response.status}) ${details}`.trim());
  }
  return response.json().catch(() => ({}));
}

async function syncWiseBalanceStatement({
  lookbackDays = getWiseStatementLookbackDays(),
  actor = "wise-sync",
} = {}) {
  const statement = await fetchWiseBalanceStatement({ lookbackDays });
  const transactions = extractWiseStatementTransactions(statement);
  const summary = {
    scanned: transactions.length,
    created: 0,
    updated: 0,
    auto_topups: 0,
    auto_invoices: 0,
    manual_review: 0,
    ignored: 0,
  };
  for (const transaction of transactions) {
    const candidate = extractWiseReceiptFromStatementTransaction(transaction);
    if (!candidate) {
      summary.ignored += 1;
      continue;
    }
    const existing = await getBillingBankReceiptByExternalKey(candidate.external_key, {
      allowMissing: true,
    });
    const receipt = await upsertBillingBankReceipt(candidate);
    if (existing?.id) {
      summary.updated += 1;
    } else {
      summary.created += 1;
    }
    const result = await tryAutoResolveBankReceipt(receipt, actor);
    if (result.action === "topup") summary.auto_topups += 1;
    if (result.action === "invoice") summary.auto_invoices += 1;
    if (result.action === "manual_review") summary.manual_review += 1;
  }
  return summary;
}

async function acceptWiseWebhookDelivery(rawBody, headers = {}) {
  const deliveryId = String(headers?.deliveryId || "").trim()
    || `wise-${crypto.createHash("sha256").update(rawBody).digest("hex").slice(0, 32)}`;
  let payload = null;
  try {
    payload = JSON.parse(Buffer.isBuffer(rawBody) ? rawBody.toString("utf8") : String(rawBody || ""));
  } catch (_error) {
    throw new Error("Invalid Wise webhook payload.");
  }
  const signatureValid = verifyWiseWebhookSignature(rawBody, headers?.signature);
  if (!signatureValid) {
    const error = new Error("Wise webhook signature verification failed.");
    error.statusCode = 401;
    throw error;
  }
  const eventType = pickFirstNonEmpty(
    payload?.event_type,
    payload?.eventType,
    payload?.type,
    payload?.name
  ) || "unknown";
  const eventRecord = await upsertWiseWebhookEvent({
    delivery_id: deliveryId,
    event_type: eventType,
    subscription_id: String(headers?.subscriptionId || "").trim() || null,
    schema_version: String(headers?.schemaVersion || "").trim() || null,
    signature_valid: true,
    is_test: headers?.isTest === true,
    processing_status: "received",
    payload,
  });
  return {
    deliveryId,
    eventType,
    payload,
    isTest: headers?.isTest === true,
    eventRecord,
  };
}

async function processWiseWebhookAcceptedEvent(accepted) {
  const deliveryId = String(accepted?.deliveryId || "").trim();
  if (!deliveryId) return { scanned: 0, created: 0, updated: 0, auto_topups: 0, auto_invoices: 0, manual_review: 0, ignored: 0 };
  try {
    if (accepted?.isTest) {
      await updateWiseWebhookEventStatus(deliveryId, "ignored", "");
      return {
        scanned: 0,
        created: 0,
        updated: 0,
        auto_topups: 0,
        auto_invoices: 0,
        manual_review: 0,
        ignored: 1,
      };
    }

    let summary = null;
    try {
      if (getWiseConfigSummary().configured) {
        summary = await syncWiseBalanceStatement({
          lookbackDays: WISE_WEBHOOK_SYNC_LOOKBACK_DAYS,
          actor: "wise-webhook",
        });
      }
    } catch (_error) {
      summary = null;
    }

    if (!summary || Number(summary.scanned || 0) === 0) {
      const fallbackSummary = {
        scanned: 0,
        created: 0,
        updated: 0,
        auto_topups: 0,
        auto_invoices: 0,
        manual_review: 0,
        ignored: 0,
      };
      const candidates = extractWiseReceiptCandidatesFromWebhookPayload(accepted?.payload || {});
      for (const candidate of candidates) {
        fallbackSummary.scanned += 1;
        const existing = await getBillingBankReceiptByExternalKey(candidate.external_key, {
          allowMissing: true,
        });
        const receipt = await upsertBillingBankReceipt({
          ...candidate,
          delivery_id: deliveryId,
        });
        if (existing?.id) {
          fallbackSummary.updated += 1;
        } else {
          fallbackSummary.created += 1;
        }
        const result = await tryAutoResolveBankReceipt(receipt, "wise-webhook");
        if (result.action === "topup") fallbackSummary.auto_topups += 1;
        if (result.action === "invoice") fallbackSummary.auto_invoices += 1;
        if (result.action === "manual_review") fallbackSummary.manual_review += 1;
      }
      summary = fallbackSummary;
    }

    await updateWiseWebhookEventStatus(deliveryId, "processed", "");
    return summary;
  } catch (error) {
    await updateWiseWebhookEventStatus(
      deliveryId,
      "failed",
      error?.message || "Wise webhook processing failed."
    ).catch(() => {});
    throw error;
  }
}

async function sendBillingInvoiceById(invoiceId, options = {}) {
  const invoiceWithItems = await getBillingInvoiceById(invoiceId, { withItems: true });
  if (!invoiceWithItems?.id) {
    throw new Error("Invoice not found.");
  }
  const currentStatus = normalizeInvoiceStatus(invoiceWithItems.status);
  const manualSettlement = invoiceRequiresManualSettlement(invoiceWithItems);
  const topupPaidButUnsent =
    isTopupBillingInvoice(invoiceWithItems)
    && currentStatus === "paid"
    && !String(invoiceWithItems?.sent_at || "").trim();
  const monthlyPaidButUnsent =
    !isTopupBillingInvoice(invoiceWithItems)
    && manualSettlement
    && currentStatus === "paid"
    && !String(invoiceWithItems?.sent_at || "").trim();
  if (
    (currentStatus === "paid" && !topupPaidButUnsent && !monthlyPaidButUnsent)
    || currentStatus === "cancelled"
  ) {
    return {
      skipped: true,
      reason: `Invoice is already ${currentStatus}.`,
      invoice: invoiceWithItems,
    };
  }

  let user = null;
  if (invoiceWithItems.user_id) {
    try {
      user = await getSupabaseUserById(invoiceWithItems.user_id);
    } catch (_error) {
      user = null;
    }
  }

  let toEmail = getInvoiceRecipientFromSnapshot(invoiceWithItems, "");
  if (!toEmail) {
    toEmail = normalizeEmail(user?.email || "");
  }
  if (!toEmail) {
    throw new Error("Invoice contact email is missing.");
  }

  const isReminder = Boolean(options?.isReminder);
  const reminderStage = Number(options?.reminderStage) || 0;
  const reminderTitle = String(options?.reminderTitle || "").trim();
  const nowIso = new Date().toISOString();
  const effectiveIssuedAt = parseIsoTimestamp(invoiceWithItems.issued_at || nowIso) || nowIso;
  const shouldResetDueDate = manualSettlement && (!invoiceWithItems.sent_at || !invoiceWithItems.issued_at);
  const effectiveDueAt = manualSettlement
    ? (
        shouldResetDueDate
          ? addUtcDays(new Date(effectiveIssuedAt), getInvoiceTermsDays()).toISOString()
          : parseIsoTimestamp(invoiceWithItems.due_at)
            || addUtcDays(new Date(effectiveIssuedAt), getInvoiceTermsDays()).toISOString()
      )
    : null;
  const effectivePaidAt = manualSettlement
    ? parseIsoTimestamp(invoiceWithItems.paid_at || "")
    : parseIsoTimestamp(invoiceWithItems.paid_at || invoiceWithItems.sent_at || effectiveIssuedAt)
      || effectiveIssuedAt;
  const pdfInvoice = {
    ...invoiceWithItems,
    issued_at: effectiveIssuedAt,
    due_at: effectiveDueAt,
    paid_at: effectivePaidAt,
  };
  const desiredStage = isReminder ? reminderStage : 0;
  const providedVariants = Array.isArray(options?.pdfVariants) ? options.pdfVariants : [];
  let persistedPdfVariants = [];
  if (providedVariants.length) {
    persistedPdfVariants = await Promise.all(
      providedVariants.map((variant) =>
        persistBillingInvoicePdfVariant(
          pdfInvoice,
          variant?.stage,
          variant?.bytes,
          variant?.filename || buildInvoiceVariantPdfFilename(invoiceWithItems, variant?.stage)
        )
      )
    );
    const mergedMetadata = mergeBillingInvoicePdfVariantMetadata(invoiceWithItems, persistedPdfVariants);
    const metadataUpdate = await updateBillingInvoiceFields(invoiceWithItems.id, {
      metadata: mergedMetadata,
    });
    if (metadataUpdate?.metadata) {
      invoiceWithItems.metadata = metadataUpdate.metadata;
      pdfInvoice.metadata = metadataUpdate.metadata;
    }
  }
  const providedVariant =
    providedVariants.find((entry) => entry?.stage === normalizeInvoicePdfVariantStage(desiredStage))
    || (!isReminder
      ? providedVariants.find((entry) => entry?.stage === "0")
      : null)
    || null;
  const storedVariant = providedVariant
    ? null
    : await loadStoredBillingInvoicePdfVariant(pdfInvoice, desiredStage, {
        allowFallback: false,
      }).catch(() => null);
  const pdfBytes = providedVariant?.bytes
    || storedVariant?.bytes
    || null;
  if (!pdfBytes) {
    throw new Error("Approved invoice PDF is unavailable. Regenerate the invoice and try again.");
  }
  const pdfFilename =
    String(providedVariant?.filename || storedVariant?.filename || "").trim()
    || buildInvoiceVariantPdfFilename(invoiceWithItems, desiredStage);
  const viewUrl = buildInvoiceViewUrl(invoiceWithItems, desiredStage, {
    publicAppUrl: options?.publicAppUrl,
  });
  const subject = buildInvoiceEmailSubject(pdfInvoice, { isReminder, reminderStage });
  const emailPayload = await sendResendEmail({
    to: toEmail,
    subject,
    html: buildInvoiceEmailHtml(pdfInvoice, invoiceWithItems.items || [], {
      isReminder,
      reminderTitle,
      reminderStage,
      viewUrl,
    }),
    text: buildInvoiceEmailText(pdfInvoice, { isReminder, viewUrl }),
    attachments: [
      {
        filename: pdfFilename,
        content: pdfBytes,
        contentType: "application/pdf",
      },
    ],
  });

  const dueMs = Date.parse(effectiveDueAt || 0);
  const nextStatus = currentStatus === "paid"
    ? "paid"
    : manualSettlement
      ? Number.isFinite(dueMs) && dueMs < Date.now()
        ? "overdue"
        : "sent"
      : "paid";
  const patch = {
    status: nextStatus,
    email_message_id: String(emailPayload?.id || "").trim() || null,
  };
  if (!invoiceWithItems.issued_at) {
    patch.issued_at = effectiveIssuedAt;
  }
  if (!invoiceWithItems.sent_at) {
    patch.sent_at = nowIso;
  }
  if (manualSettlement) {
    if (!invoiceWithItems.due_at || shouldResetDueDate) {
      patch.due_at = effectiveDueAt;
    }
  } else if (!invoiceWithItems.paid_at) {
    patch.paid_at = effectivePaidAt;
  }
  if (isReminder) {
    patch.reminder_stage = Math.max(Number(invoiceWithItems.reminder_stage) || 0, reminderStage);
    patch.last_reminder_sent_at = nowIso;
  }
  const updated = await updateBillingInvoiceFields(invoiceWithItems.id, patch);
  await insertBillingInvoiceEvent(invoiceWithItems.id, {
    event_type: isReminder ? "reminder" : "sent",
    actor: "system",
    channel: "email",
    message: isReminder ? reminderTitle || "Reminder sent." : "Invoice sent.",
    metadata: {
      resend_id: emailPayload?.id || null,
      to: toEmail,
      reminder_stage: isReminder ? reminderStage : null,
      attachment_source: providedVariant ? "browser-upload" : storedVariant ? "stored-browser-pdf" : "server-fallback",
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

async function runInvoiceReminders(options = {}) {
  const safeLimit = Math.max(1, Math.min(500, Number(options?.limit) || 200));
  const candidateInvoices = await listBillingInvoices({
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
      const result = await sendBillingInvoiceById(invoice.id, {
        isReminder: true,
        reminderStage: desiredStage,
        reminderTitle,
        requireApprovedPdf: true,
        publicAppUrl: options?.publicAppUrl,
      });
      out.push({
        invoice_id: invoice.id,
        reminder_stage: desiredStage,
        status: result?.invoice?.status || invoice.status,
      });
    } catch (error) {
      await insertBillingInvoiceEvent(invoice.id, {
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

async function buildInvoicesFromLabelHistory(options = {}) {
  const mode = normalizeInvoiceRunMode(options?.mode);
  const billingWindow = options?.window || getPreviousMonthWindow();
  const onlyUserId = String(options?.userId || "").trim();
  const runTimestamp = new Date().toISOString();
  const termsDays = getInvoiceTermsDays();
  const settings = await getAdminSettings();

  const [historyRows, users, billingPrefs] = await Promise.all([
    listUnbilledGenerationRows({ window: billingWindow, userId: onlyUserId }),
    listSupabaseUsers(1000),
    listClientBillingPreferences(5000),
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
      userId,
      billingWindow.startIsoDate,
      billingWindow.endIsoDate
    );
    const existingStatus = normalizeInvoiceStatus(existing?.status);
    const existingAlreadyIssued =
      Boolean(String(existing?.sent_at || "").trim())
      || ["sent", "overdue", "paid", "cancelled"].includes(existingStatus);
    if (existing && existingAlreadyIssued) {
      result.skipped.push({
        user_id: userId,
        invoice_id: existing.id,
        reason:
          ["paid", "cancelled"].includes(existingStatus)
            ? `Invoice already ${existingStatus}.`
            : `Invoice already issued (${existingStatus}).`,
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
      metadata: {
        ...(
          existing?.metadata && typeof existing.metadata === "object"
            ? existing.metadata
            : {
                source: "portal_label_generations",
              }
        ),
        invoice_profile: {
          companyName: profile.company_name,
          contactName: profile.contact_name,
          contactEmail: profile.contact_email,
          contactPhone: profile.contact_phone,
          billingAddress: profile.billing_address,
          taxId: profile.tax_id,
          customerId: profile.customer_id,
          accountManager: profile.account_manager,
        },
      },
      updated_at: runTimestamp,
    };

    const previewInvoice = {
      id: existing?.id || null,
      user_id: userId,
      invoice_number: String(existing?.invoice_number || existing?.metadata?.invoice_number || "").trim() || null,
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
        reference:
          previewInvoice.invoice_number
          || (previewInvoice.id ? toInvoiceReference(previewInvoice.id) : "M-INV-PREVIEW"),
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

    const persisted = await upsertBillingInvoiceWithPublicCode(invoicePayload, existing);
    if (!persisted?.id) {
      throw new Error("Invoice could not be persisted.");
    }
    await replaceBillingInvoiceItems(persisted.id, invoiceData.items);
    const generationIds = userRows.map((row) => row?.id).filter(Boolean);
    await markGenerationRowsBilled(generationIds, persisted.id);
    await insertBillingInvoiceEvent(persisted.id, {
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
        const sendResult = await sendBillingInvoiceById(persisted.id, {
          isReminder: false,
          requireApprovedPdf: true,
          publicAppUrl: options?.publicAppUrl,
        });
        if (!sendResult?.skipped) {
          result.invoices_sent += 1;
        }
        finalInvoice = sendResult?.invoice || persisted;
      } catch (error) {
        await insertBillingInvoiceEvent(persisted.id, {
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

async function computeCurrentMonthUnbilledTotal(userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return 0;
  const monthWindow = getCurrentMonthWindow();
  const rows = await listUnbilledGenerationRows({
    window: monthWindow,
    userId: safeUserId,
  });
  const totalCents = rows.reduce((sum, row) => sum + toCents(row?.total_price), 0);
  return fromCents(totalCents);
}

async function listClientBillingPreferences(limit = 2000) {
  const safeLimit = Math.max(1, Math.min(5000, Number(limit) || 2000));
  const params = new URLSearchParams();
  params.set("select", "user_id,invoice_enabled,card_enabled,updated_at,updated_by");
  params.set("order", "updated_at.desc.nullslast");
  params.set("limit", String(safeLimit));
  const response = await supabaseServiceRequest(
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
  return filterTrustedClientBillingPreferenceRows(rows);
}

async function saveClientBillingPreference(adminUserId, userId, payload) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) {
    throw new Error("Client id is required.");
  }
  const normalized = normalizeClientBillingPreference(payload);
  const response = await supabaseServiceRequest(
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
    company_name: normalizeClientCompanyName(invite?.company_name || ""),
    expires_at: invite?.expires_at || null,
    created_at: invite?.created_at || null,
    claimed_at: invite?.claimed_at || null,
    claimed_email: normalizeEmail(invite?.claimed_email || ""),
    revoked_at: invite?.revoked_at || null,
    invite_url: buildInviteUrl(baseUrl, invite),
  };
}

function buildShipmentExtractRequestUrl(baseUrl, request) {
  const encrypted = String(request?.token_encrypted || "").trim();
  if (!encrypted) return "";
  try {
    const token = decryptToken(encrypted);
    const requestUrl = new URL("/clean-data", baseUrl);
    requestUrl.searchParams.set("token", token);
    return requestUrl.toString();
  } catch (_error) {
    return "";
  }
}

function getShipmentExtractRequestStatus(row) {
  if (row?.submitted_at) return "submitted";
  if (shipmentExtractRequestIsExpired(row)) return "expired";
  return "open";
}

function mapShipmentExtractRequestRow(row, baseUrl) {
  return {
    id: row?.id || null,
    client_email: normalizeEmail(row?.client_email || ""),
    client_company_name: getShipmentExtractCompanyName(row),
    expires_at: row?.expires_at || null,
    created_at: row?.created_at || null,
    submitted_at: row?.submitted_at || null,
    submitted_filename: row?.submitted_filename || "",
    submitted_rows: Number(row?.submitted_rows) || 0,
    submitted_columns: Number(row?.submitted_columns) || 0,
    kept_columns: Array.isArray(row?.kept_columns) ? row.kept_columns : [],
    removed_columns: Array.isArray(row?.removed_columns) ? row.removed_columns : [],
    registration_invite_id: row?.registration_invite_id || null,
    request_url: buildShipmentExtractRequestUrl(baseUrl, row),
    status: getShipmentExtractRequestStatus(row),
  };
}

function shouldIncludeAdminClient(user) {
  if (!user?.id) return false;
  return !canManageRegistrationInvites(user);
}

async function buildAdminDashboard(baseUrl) {
  const [settings, invites, shipmentExtractRequests, users, historyRows, billingPreferences, invoices, wiseReceipts] = await Promise.all([
    getAdminSettings(),
    listRegistrationInvites({ limit: 50 }),
    listShipmentExtractRequests({ limit: 50 }).catch(() => []),
    listSupabaseUsers(1000),
    listGenerationHistoryRows(10000),
    listClientBillingPreferences(2000),
    fetchAllSupabaseRows(
      BILLING_INVOICES_TABLE,
      {
        select: BILLING_INVOICE_SELECT_FIELDS,
        order: "updated_at.desc,created_at.desc",
      },
      { pageSize: 500, allowMissingSchema: true }
    ).catch(() => []),
    listBillingBankReceipts({
      limit: 80,
      statuses: ["manual_review", "pending"],
      allowMissing: true,
    }),
  ]);
  const nonTopupInvoiceIds = invoices
    .filter((invoice) => !isTopupBillingInvoice(invoice))
    .map((invoice) => String(invoice?.id || "").trim())
    .filter(Boolean);
  const [invoiceItems, walletCheckoutTransactions] = await Promise.all([
    (async () => {
      const items = [];
      for (const invoiceIdChunk of chunkArray(nonTopupInvoiceIds, 100)) {
        const inFilter = buildSupabaseInFilter(invoiceIdChunk);
        if (!inFilter) continue;
        const rows = await fetchAllSupabaseRows(
          BILLING_INVOICE_ITEMS_TABLE,
          {
            select:
              "invoice_id,service_type,quantity,amount_ex_vat",
            invoice_id: inFilter,
          },
          { pageSize: 500, allowMissingSchema: true }
        ).catch(() => []);
        items.push(...rows);
      }
      return items;
    })(),
    fetchAllSupabaseRows(
      BILLING_WALLET_TRANSACTIONS_TABLE,
      {
        select: "user_id,source,amount_cents,metadata",
        source: "eq.label_checkout",
        amount_cents: "lt.0",
      },
      { pageSize: 500, allowMissingSchema: true }
    ).catch(() => []),
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
  const billingMetricsByUserId = buildBillingMetricsByUser(
    settings,
    invoices,
    invoiceItems,
    walletCheckoutTransactions
  );

  const clients = users
    .filter(shouldIncludeAdminClient)
    .map((user) => ({
      user: {
        id: user.id,
        email: normalizeEmail(user.email || ""),
        created_at: user.created_at || null,
        user_metadata: user.user_metadata && typeof user.user_metadata === "object" ? user.user_metadata : {},
      },
      billing: billingByUserId.get(user.id) || { ...DEFAULT_CLIENT_BILLING_PREF },
      metrics: buildAdminClientMetrics(
        user,
        historyByUserId.get(user.id) || [],
        billingMetricsByUserId.get(user.id) || null,
        billingByUserId.get(user.id) || DEFAULT_CLIENT_BILLING_PREF,
        invoiceStatsByUserId.get(user.id) || null
      ),
    }))
    .sort((left, right) => {
      const leftValue = Date.parse(left?.metrics?.last_generation_at || left?.user?.created_at || 0);
      const rightValue = Date.parse(right?.metrics?.last_generation_at || right?.user?.created_at || 0);
      return rightValue - leftValue;
    });

  const inviteHistory = invites.map((invite) => mapInviteHistoryRow(invite, baseUrl));
  const shipmentExtractHistory = shipmentExtractRequests.map((request) =>
    mapShipmentExtractRequestRow(request, baseUrl)
  );
  return {
    settings,
    invites: inviteHistory,
    shipment_extract_requests: shipmentExtractHistory,
    clients,
    summary: buildAdminClientSummary(clients, invites),
    billing: {
      invoices: buildInvoiceListResponseRows(invoices.slice(0, 120)),
      wise: {
        ...getWiseConfigSummary(),
        receipts: wiseReceipts.map(mapBillingBankReceiptRow).filter(Boolean),
      },
    },
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

async function deleteSupabaseUserById(userId) {
  const safeUserId = String(userId || "").trim();
  if (!safeUserId) return;
  const response = await fetch(`${SUPABASE_URL}/auth/v1/admin/users/${safeUserId}`, {
    method: "DELETE",
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    },
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Rollback failed: could not delete user (${response.status}) ${details}`.trim());
  }
}

function buildSupabaseInFilter(values) {
  const safeValues = Array.isArray(values)
    ? values.map((value) => String(value || '').trim()).filter(Boolean)
    : [];
  if (!safeValues.length) return '';
  return `in.(${safeValues.map((value) => JSON.stringify(value)).join(',')})`;
}

function chunkArray(values = [], size = 100) {
  const safeValues = Array.isArray(values) ? values : [];
  const chunkSize = Math.max(1, Number(size) || 100);
  const chunks = [];
  for (let index = 0; index < safeValues.length; index += chunkSize) {
    chunks.push(safeValues.slice(index, index + chunkSize));
  }
  return chunks;
}

function isMissingSupabaseSchemaError(tableName, details) {
  const table = String(tableName || '').toLowerCase();
  const text = String(details || '').toLowerCase();
  return Boolean(
    table
    && text.includes(table)
    && (text.includes('relation') || text.includes('column') || text.includes('schema'))
  );
}

function isMissingPrivacyRequestsSchemaError(details) {
  return isMissingSupabaseSchemaError(PRIVACY_REQUESTS_TABLE, details);
}

async function fetchAllSupabaseRows(tableName, paramsInit = {}, options = {}) {
  const pageSize = Math.max(1, Math.min(1000, Number(options.pageSize) || 500));
  const allowMissingSchema = options.allowMissingSchema === true;
  const rows = [];
  let offset = 0;

  while (true) {
    const params = new URLSearchParams();
    Object.entries(paramsInit || {}).forEach(([key, value]) => {
      if (value === undefined || value === null || value === '') return;
      params.set(key, String(value));
    });
    params.set('limit', String(pageSize));
    params.set('offset', String(offset));

    const response = await supabaseServiceRequest(`/rest/v1/${tableName}?${params.toString()}`, {
      method: 'GET',
    });
    if (!response.ok) {
      const details = await response.text().catch(() => '');
      if (allowMissingSchema && isMissingSupabaseSchemaError(tableName, details)) {
        return [];
      }
      throw new Error(`Failed reading ${tableName} (${response.status}) ${details}`.trim());
    }

    const pageRows = await response.json().catch(() => []);
    if (!Array.isArray(pageRows) || !pageRows.length) break;
    rows.push(...pageRows);
    if (pageRows.length < pageSize) break;
    offset += pageRows.length;
  }

  return rows;
}

async function deleteSupabaseRows(tableName, paramsInit = {}, options = {}) {
  const params = new URLSearchParams();
  Object.entries(paramsInit || {}).forEach(([key, value]) => {
    if (value === undefined || value === null || value === '') return;
    params.set(key, String(value));
  });

  const response = await supabaseServiceRequest(`/rest/v1/${tableName}?${params.toString()}`, {
    method: 'DELETE',
    headers: {
      Prefer: options.returnRepresentation === true ? 'return=representation' : 'return=minimal',
    },
  });
  if (!response.ok) {
    const details = await response.text().catch(() => '');
    throw new Error(`Failed deleting ${tableName} (${response.status}) ${details}`.trim());
  }
  if (options.returnRepresentation === true) {
    const rows = await response.json().catch(() => []);
    return Array.isArray(rows) ? rows : [];
  }
  return [];
}

async function patchSupabaseRows(tableName, paramsInit = {}, patch = {}, options = {}) {
  const params = new URLSearchParams();
  Object.entries(paramsInit || {}).forEach(([key, value]) => {
    if (value === undefined || value === null || value === '') return;
    params.set(key, String(value));
  });

  const response = await supabaseServiceRequest(`/rest/v1/${tableName}?${params.toString()}`, {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json',
      Prefer: options.returnRepresentation === true ? 'return=representation' : 'return=minimal',
    },
    body: JSON.stringify(patch && typeof patch === 'object' ? patch : {}),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => '');
    throw new Error(`Failed updating ${tableName} (${response.status}) ${details}`.trim());
  }
  if (options.returnRepresentation === true) {
    const rows = await response.json().catch(() => []);
    return Array.isArray(rows) ? rows : [];
  }
  return [];
}

async function insertPrivacyRequest(record, options = {}) {
  const nowIso = new Date().toISOString();
  const response = await supabaseServiceRequest(`/rest/v1/${PRIVACY_REQUESTS_TABLE}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Prefer: 'return=representation',
    },
    body: JSON.stringify([
      {
        source: String(record?.source || 'portal').trim() || 'portal',
        request_type: String(record?.requestType || 'export').trim() || 'export',
        status: String(record?.status || 'received').trim() || 'received',
        subject_user_id: String(record?.subjectUserId || '').trim() || null,
        requested_by_user_id: String(record?.requestedByUserId || '').trim() || null,
        provider: String(record?.provider || '').trim() || null,
        shop_domain: String(record?.shopDomain || '').trim() || null,
        external_customer_id: String(record?.externalCustomerId || '').trim() || null,
        external_customer_email: normalizeEmail(record?.externalCustomerEmail || '') || null,
        request_payload:
          record?.requestPayload && typeof record.requestPayload === 'object'
            ? record.requestPayload
            : {},
        result_payload:
          record?.resultPayload && typeof record.resultPayload === 'object'
            ? record.resultPayload
            : {},
        requested_at: record?.requestedAt || nowIso,
        processed_at: record?.processedAt || null,
        updated_at: nowIso,
      },
    ]),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => '');
    if (options.allowMissingSchema === true && isMissingPrivacyRequestsSchemaError(details)) {
      return null;
    }
    throw new Error(`Failed creating privacy request (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function updatePrivacyRequest(requestId, patch = {}, options = {}) {
  const safeRequestId = String(requestId || '').trim();
  if (!safeRequestId) return null;
  const response = await supabaseServiceRequest(
    `/rest/v1/${PRIVACY_REQUESTS_TABLE}?id=eq.${safeRequestId}`,
    {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        Prefer: 'return=representation',
      },
      body: JSON.stringify({
        ...(patch && typeof patch === 'object' ? patch : {}),
        updated_at: new Date().toISOString(),
      }),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => '');
    if (options.allowMissingSchema === true && isMissingPrivacyRequestsSchemaError(details)) {
      return null;
    }
    throw new Error(`Failed updating privacy request (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

function buildPrivacyExportFilename(date = new Date()) {
  const stamp = date.toISOString().slice(0, 10);
  return `shipide-privacy-export-${stamp}.json`;
}

function buildPrivacyExportSummary(bundle = {}) {
  return {
    provider_connections: Array.isArray(bundle.provider_connections) ? bundle.provider_connections.length : 0,
    history_rows: Array.isArray(bundle.history) ? bundle.history.length : 0,
    invoices: Array.isArray(bundle.billing?.invoices) ? bundle.billing.invoices.length : 0,
    invoice_items: Array.isArray(bundle.billing?.invoice_items) ? bundle.billing.invoice_items.length : 0,
    topups: Array.isArray(bundle.billing?.topups) ? bundle.billing.topups.length : 0,
    wallet_transactions: Array.isArray(bundle.billing?.wallet_transactions)
      ? bundle.billing.wallet_transactions.length
      : 0,
    clickwrap_acceptances: Array.isArray(bundle.clickwrap_acceptances)
      ? bundle.clickwrap_acceptances.length
      : 0,
    privacy_requests: Array.isArray(bundle.privacy_requests) ? bundle.privacy_requests.length : 0,
  };
}

async function getSingleOptionalSupabaseRow(tableName, paramsInit = {}) {
  const rows = await fetchAllSupabaseRows(tableName, paramsInit, {
    pageSize: 1,
    allowMissingSchema: true,
  }).catch(() => []);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function listProviderConnectionsForPrivacyExport(userId) {
  return fetchAllSupabaseRows(
    'provider_connections',
    {
      select: 'provider,shop_domain,status,scopes,connected_at,updated_at',
      user_id: `eq.${userId}`,
      order: 'updated_at.desc',
    },
    { pageSize: 200, allowMissingSchema: true }
  ).catch(() => []);
}

async function listHistoryRowsForPrivacyExport(userId) {
  return fetchAllSupabaseRows(
    HISTORY_TABLE,
    {
      select: '*',
      user_id: `eq.${userId}`,
      order: 'created_at.desc',
    },
    { pageSize: 500, allowMissingSchema: true }
  ).catch(() => []);
}

async function listBillingInvoicesForPrivacyExport(userId) {
  return fetchAllSupabaseRows(
    BILLING_INVOICES_TABLE,
    {
      select: BILLING_INVOICE_SELECT_FIELDS,
      user_id: `eq.${userId}`,
      order: 'updated_at.desc',
    },
    { pageSize: 200, allowMissingSchema: true }
  ).catch(() => []);
}

async function listBillingInvoiceItemsForPrivacyExport(invoiceIds = []) {
  const items = [];
  for (const invoiceIdChunk of chunkArray(invoiceIds, 100)) {
    const inFilter = buildSupabaseInFilter(invoiceIdChunk);
    if (!inFilter) continue;
    const rows = await fetchAllSupabaseRows(
      BILLING_INVOICE_ITEMS_TABLE,
      {
        select:
          'id,invoice_id,generation_id,generated_at,service_type,quantity,recipient_name,recipient_city,recipient_country,tracking_id,label_id,unit_price_ex_vat,amount_ex_vat,vat_amount,amount_inc_vat,sort_index,metadata,created_at',
        invoice_id: inFilter,
        order: 'sort_index.asc,id.asc',
      },
      { pageSize: 500, allowMissingSchema: true }
    ).catch(() => []);
    items.push(...rows);
  }
  return items;
}

async function listBillingTopupsForPrivacyExport(userId) {
  return fetchAllSupabaseRows(
    BILLING_WALLET_TOPUPS_TABLE,
    {
      select: BILLING_TOPUP_SELECT_FIELDS,
      user_id: `eq.${userId}`,
      order: 'requested_at.desc',
    },
    { pageSize: 200, allowMissingSchema: true }
  ).catch(() => []);
}

async function listBillingWalletTransactionsForPrivacyExport(userId) {
  return fetchAllSupabaseRows(
    BILLING_WALLET_TRANSACTIONS_TABLE,
    {
      select: 'id,user_id,source,amount_cents,balance_after_cents,reference,metadata,created_at',
      user_id: `eq.${userId}`,
      order: 'created_at.desc,id.desc',
    },
    { pageSize: 200, allowMissingSchema: true }
  ).catch(() => []);
}

async function listClickwrapAcceptancesForPrivacyExport(userId) {
  return fetchAllSupabaseRows(
    CLICKWRAP_ACCEPTANCES_TABLE,
    {
      select:
        'id,invite_id,accepted_email,invited_email,contract_id,contract_version,contract_hash_sha256,agreement_method,scrolled_to_end,scrolled_to_end_at,agreed_at,client_timezone,client_locale,proof_payload,created_at',
      user_id: `eq.${userId}`,
      order: 'agreed_at.desc,created_at.desc',
    },
    { pageSize: 200, allowMissingSchema: true }
  ).catch(() => []);
}

async function listPrivacyRequestsForPrivacyExport(userId) {
  return fetchAllSupabaseRows(
    PRIVACY_REQUESTS_TABLE,
    {
      select:
        'id,source,request_type,status,provider,shop_domain,external_customer_id,external_customer_email,request_payload,result_payload,requested_at,processed_at,created_at,updated_at',
      requested_by_user_id: `eq.${userId}`,
      order: 'requested_at.desc,created_at.desc',
    },
    { pageSize: 200, allowMissingSchema: true }
  ).catch(() => []);
}

async function buildPrivacyExportBundle(user) {
  const userId = String(user?.id || '').trim();
  if (!userId) {
    throw new Error('Authentication required.');
  }
  const [
    accountSettings,
    providerConnections,
    historyRows,
    invoices,
    topups,
    walletTransactions,
    clickwrapAcceptances,
    privacyRequests,
    billingWallet,
    billingPreference,
  ] = await Promise.all([
    getSingleOptionalSupabaseRow('account_settings', {
      select: '*',
      user_id: `eq.${userId}`,
    }),
    listProviderConnectionsForPrivacyExport(userId),
    listHistoryRowsForPrivacyExport(userId),
    listBillingInvoicesForPrivacyExport(userId),
    listBillingTopupsForPrivacyExport(userId),
    listBillingWalletTransactionsForPrivacyExport(userId),
    listClickwrapAcceptancesForPrivacyExport(userId),
    listPrivacyRequestsForPrivacyExport(userId),
    getSingleOptionalSupabaseRow(BILLING_WALLET_TABLE, {
      select: 'user_id,balance_cents,currency,created_at,updated_at',
      user_id: `eq.${userId}`,
    }),
    getClientBillingPreferenceForUser(userId).catch(() => ({ ...DEFAULT_CLIENT_BILLING_PREF })),
  ]);
  const invoiceIds = invoices.map((invoice) => String(invoice?.id || '').trim()).filter(Boolean);
  const invoiceItems = await listBillingInvoiceItemsForPrivacyExport(invoiceIds);
  const bundle = {
    schema: 'shipide.privacy-export.v1',
    generated_at: new Date().toISOString(),
    user: {
      id: userId,
      email: normalizeEmail(user?.email || '') || null,
      created_at: user?.created_at || null,
      last_sign_in_at: user?.last_sign_in_at || null,
      user_metadata: user?.user_metadata && typeof user.user_metadata === 'object' ? user.user_metadata : {},
    },
    account_settings: accountSettings,
    provider_connections: providerConnections,
    history: historyRows,
    billing: {
      preference: normalizeClientBillingPreference(billingPreference),
      wallet: billingWallet,
      invoices,
      invoice_items: invoiceItems,
      topups,
      wallet_transactions: walletTransactions,
    },
    clickwrap_acceptances: clickwrapAcceptances,
    privacy_requests: privacyRequests,
  };
  return {
    bundle,
    summary: buildPrivacyExportSummary(bundle),
  };
}

async function handlePrivacyExport(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: 'Authentication required.' });
    return;
  }
  const requestRecord = await insertPrivacyRequest(
    {
      source: 'portal',
      requestType: 'export',
      status: 'processing',
      subjectUserId: user.id,
      requestedByUserId: user.id,
      requestPayload: {
        requested_via: '/api/privacy/export',
        requested_email: normalizeEmail(user.email || '') || null,
      },
    },
    { allowMissingSchema: true }
  ).catch(() => null);
  try {
    const { bundle, summary } = await buildPrivacyExportBundle(user);
    if (requestRecord?.id) {
      await updatePrivacyRequest(
        requestRecord.id,
        {
          status: 'completed',
          processed_at: new Date().toISOString(),
          result_payload: {
            exported_at: bundle.generated_at,
            summary,
          },
        },
        { allowMissingSchema: true }
      ).catch(() => null);
    }
    send(
      res,
      200,
      {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store',
        'Content-Disposition': `attachment; filename="${buildPrivacyExportFilename()}"`,
      },
      Buffer.from(JSON.stringify(bundle, null, 2))
    );
  } catch (error) {
    if (requestRecord?.id) {
      await updatePrivacyRequest(
        requestRecord.id,
        {
          status: 'failed',
          processed_at: new Date().toISOString(),
          result_payload: {
            error: error?.message || 'Privacy export failed.',
          },
        },
        { allowMissingSchema: true }
      ).catch(() => null);
    }
    sendJson(res, 500, { error: error.message || 'Could not build privacy export.' });
  }
}

async function handlePrivacyDelete(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: 'Authentication required.' });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || 'Invalid request body.' });
    return;
  }
  if (body?.confirm !== true) {
    sendJson(res, 400, {
      error: 'Account deletion requests require an explicit confirm=true flag.',
    });
    return;
  }
  const requestRecord = await insertPrivacyRequest(
    {
      source: 'portal',
      requestType: 'delete',
      status: 'manual_review_required',
      subjectUserId: user.id,
      requestedByUserId: user.id,
      requestPayload: {
        requested_via: '/api/privacy/delete',
        requested_email: normalizeEmail(user.email || '') || null,
        reason: String(body?.reason || '').trim() || null,
      },
      resultPayload: {
        queued_for_manual_review: true,
      },
    },
    { allowMissingSchema: true }
  ).catch(() => null);
  if (!requestRecord?.id) {
    sendJson(res, 503, { error: 'Privacy request storage is unavailable.' });
    return;
  }
  sendJson(res, 202, {
    ok: true,
    queued: true,
    requestId: requestRecord.id,
    status: requestRecord.status,
    message: 'Deletion request recorded and queued for manual review.',
  });
}

async function deleteShopifyConnection(userId, shop) {
  const normalizedShop = normalizeShopDomain(shop);
  if (!userId || !normalizedShop) return;
  await deleteSupabaseRows('provider_connections', {
    user_id: `eq.${userId}`,
    provider: 'eq.shopify',
    shop_domain: `eq.${normalizedShop}`,
  });
}

async function deleteWixConnection(userId, instanceId) {
  const normalizedInstanceId = String(instanceId || '').trim();
  if (!userId || !normalizedInstanceId) return;
  await deleteSupabaseRows('provider_connections', {
    user_id: `eq.${userId}`,
    provider: 'eq.wix',
    shop_domain: `eq.${normalizedInstanceId}`,
  });
}

async function deleteWooCommerceConnection(userId, storeUrl) {
  const normalizedStoreUrl = normalizeWooCommerceStoreUrl(storeUrl);
  if (!userId || !normalizedStoreUrl) return;
  await deleteSupabaseRows('provider_connections', {
    user_id: `eq.${userId}`,
    provider: 'eq.woocommerce',
    shop_domain: `eq.${normalizedStoreUrl}`,
  });
}

async function deleteShopifyConnectionsByShop(shop) {
  const normalizedShop = normalizeShopDomain(shop);
  if (!normalizedShop) return [];
  return deleteSupabaseRows(
    'provider_connections',
    {
      provider: 'eq.shopify',
      shop_domain: `eq.${normalizedShop}`,
    },
    { returnRepresentation: true }
  );
}

async function deleteShopifySessionsByShop(shop) {
  const normalizedShop = normalizeShopDomain(shop);
  if (!normalizedShop) return;
  await deleteSupabaseRows('Session', {
    shop: `eq.${normalizedShop}`,
  });
}

function normalizePrivacyMatchText(value) {
  return String(value || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function buildRecipientFingerprint(parts = {}) {
  const pieces = [
    parts?.name,
    parts?.street,
    parts?.city,
    parts?.zip,
    parts?.country,
  ].map((value) => normalizePrivacyMatchText(value));
  if (pieces.filter(Boolean).length < 3) {
    return "";
  }
  return pieces.join("|");
}

function buildShopifyOrderRecipientFingerprint(order) {
  const shipping =
    order?.shipping_address && typeof order.shipping_address === "object"
      ? order.shipping_address
      : order?.billing_address && typeof order.billing_address === "object"
        ? order.billing_address
        : {};
  const name =
    String(shipping?.name || "").trim() ||
    [shipping?.first_name, shipping?.last_name].filter(Boolean).join(" ").trim();
  return buildRecipientFingerprint({
    name,
    street: shipping?.address1,
    city: shipping?.city,
    zip: shipping?.zip,
    country: shipping?.country_code || shipping?.country,
  });
}

function buildLabelRecipientFingerprint(labelData = {}) {
  return buildRecipientFingerprint({
    name: labelData?.recipientName,
    street: labelData?.recipientStreet,
    city: labelData?.recipientCity,
    zip: labelData?.recipientZip,
    country: labelData?.recipientCountry,
  });
}

function normalizeShopifyComplianceOrderIds(payload = {}) {
  const candidates = [];
  const requested = Array.isArray(payload?.orders_requested) ? payload.orders_requested : [];
  const toRedact = Array.isArray(payload?.orders_to_redact) ? payload.orders_to_redact : [];
  requested.concat(toRedact).forEach((value) => {
    const normalized = String(value || "").trim();
    if (normalized) {
      candidates.push(normalized);
    }
  });
  return Array.from(new Set(candidates));
}

function serializeShopifyPrivacyOrder(order = {}) {
  const shipping =
    order?.shipping_address && typeof order.shipping_address === "object"
      ? order.shipping_address
      : order?.billing_address && typeof order.billing_address === "object"
        ? order.billing_address
        : {};
  const customer = order?.customer && typeof order.customer === "object" ? order.customer : {};
  return {
    id: String(order?.id || "").trim() || null,
    name: String(order?.name || "").trim() || null,
    created_at: order?.created_at || null,
    customer_id: String(customer?.id || "").trim() || null,
    customer_email: normalizeEmail(order?.email || customer?.email || "") || null,
    shipping_address: {
      name:
        String(shipping?.name || "").trim() ||
        [shipping?.first_name, shipping?.last_name].filter(Boolean).join(" ").trim() ||
        null,
      address1: String(shipping?.address1 || "").trim() || null,
      city: String(shipping?.city || "").trim() || null,
      zip: String(shipping?.zip || "").trim() || null,
      country: String(shipping?.country_code || shipping?.country || "").trim() || null,
    },
  };
}

async function listShopifyConnectionsByShop(shop) {
  const normalizedShop = normalizeShopDomain(shop);
  if (!normalizedShop) return [];
  return fetchAllSupabaseRows(
    "provider_connections",
    {
      select: "user_id,shop_domain,access_token,status,connected_at,updated_at",
      provider: "eq.shopify",
      shop_domain: `eq.${normalizedShop}`,
      order: "updated_at.desc",
    },
    { pageSize: 200 }
  );
}

async function resolveShopifyPrivacyAccessToken(connections = []) {
  for (const connection of Array.isArray(connections) ? connections : []) {
    try {
      const accessToken = decryptToken(connection?.access_token || "");
      if (accessToken) {
        return accessToken;
      }
    } catch (_error) {}
  }
  return "";
}

async function fetchShopifyOrdersByIds(shop, accessToken, orderIds = []) {
  const normalizedShop = normalizeShopDomain(shop);
  const safeOrderIds = Array.from(
    new Set(
      Array.isArray(orderIds)
        ? orderIds.map((value) => String(value || "").trim()).filter(Boolean)
        : []
    )
  );
  if (!normalizedShop || !accessToken || !safeOrderIds.length) {
    return [];
  }

  const ordersById = new Map();
  for (const idChunk of chunkArray(safeOrderIds, 100)) {
    const url = new URL(`https://${normalizedShop}/admin/api/${SHOPIFY_API_VERSION}/orders.json`);
    url.searchParams.set("ids", idChunk.join(","));
    url.searchParams.set("status", "any");
    url.searchParams.set("limit", String(idChunk.length));
    url.searchParams.set(
      "fields",
      "id,name,created_at,email,phone,shipping_address,billing_address,customer"
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
      const error = new Error(
        `Shopify privacy order fetch failed (${response.status}) ${details}`.trim()
      );
      error.status = response.status;
      throw error;
    }
    const payload = await response.json().catch(() => null);
    const orders = Array.isArray(payload?.orders) ? payload.orders : [];
    orders.forEach((order) => {
      const orderId = String(order?.id || "").trim();
      if (!orderId || ordersById.has(orderId)) return;
      ordersById.set(orderId, order);
    });
  }

  return Array.from(ordersById.values()).sort((left, right) => {
    const leftTime = Date.parse(left?.created_at || "") || 0;
    const rightTime = Date.parse(right?.created_at || "") || 0;
    return rightTime - leftTime;
  });
}

function buildShopifyPrivacyTargets(shop, payload = {}, orders = []) {
  const normalizedShop = normalizeShopDomain(shop);
  const orderIds = new Set(normalizeShopifyComplianceOrderIds(payload));
  const customerIds = new Set();
  const customerEmails = new Set();
  const recipientFingerprints = new Set();
  const customer = payload?.customer && typeof payload.customer === "object" ? payload.customer : {};

  const payloadCustomerId = String(customer?.id || "").trim();
  const payloadCustomerEmail = normalizeEmail(customer?.email || payload?.email || "");
  if (payloadCustomerId) customerIds.add(payloadCustomerId);
  if (payloadCustomerEmail) customerEmails.add(payloadCustomerEmail);

  orders.forEach((order) => {
    const orderId = String(order?.id || "").trim();
    const customerId = String(order?.customer?.id || "").trim();
    const customerEmail = normalizeEmail(order?.email || order?.customer?.email || "");
    const fingerprint = buildShopifyOrderRecipientFingerprint(order);
    if (orderId) orderIds.add(orderId);
    if (customerId) customerIds.add(customerId);
    if (customerEmail) customerEmails.add(customerEmail);
    if (fingerprint) recipientFingerprints.add(fingerprint);
  });

  return {
    shop: normalizedShop,
    orderIds,
    customerIds,
    customerEmails,
    recipientFingerprints,
    orderExports: orders.map(serializeShopifyPrivacyOrder),
  };
}

function getShopifySourceFromLabelData(labelData = {}) {
  if (!labelData || typeof labelData !== "object") return null;
  const source = labelData?.shipideSource;
  if (!source || typeof source !== "object") return null;
  return {
    provider: String(source?.provider || "").trim().toLowerCase(),
    shop: normalizeShopDomain(source?.shop || ""),
    orderId: String(source?.orderId || "").trim(),
    orderName: String(source?.orderName || "").trim(),
    customerId: String(source?.customerId || "").trim(),
    customerEmail: normalizeEmail(source?.customerEmail || ""),
    importedAt: String(source?.importedAt || "").trim(),
    redacted: Boolean(source?.redacted),
    redactedAt: String(source?.redactedAt || "").trim(),
  };
}

function matchShopifyLabelToTargets(labelData, shop, targets) {
  const normalizedShop = normalizeShopDomain(shop);
  const source = getShopifySourceFromLabelData(labelData);

  if (source?.provider && source.provider !== "shopify") {
    return { matched: false, strategy: "", source };
  }
  if (source?.shop && normalizedShop && source.shop !== normalizedShop) {
    return { matched: false, strategy: "", source };
  }
  if (source?.orderId && targets.orderIds.has(source.orderId)) {
    return { matched: true, strategy: "source_order_id", source };
  }
  if (source?.customerId && targets.customerIds.has(source.customerId)) {
    return { matched: true, strategy: "source_customer_id", source };
  }
  if (source?.customerEmail && targets.customerEmails.has(source.customerEmail)) {
    return { matched: true, strategy: "source_customer_email", source };
  }

  const fingerprint = buildLabelRecipientFingerprint(labelData);
  if (fingerprint && targets.recipientFingerprints.has(fingerprint)) {
    return {
      matched: true,
      strategy: source?.provider === "shopify" ? "recipient_fingerprint" : "recipient_fingerprint_legacy",
      source,
    };
  }

  return { matched: false, strategy: "", source };
}

function buildShopifyRedactedLabelData(labelData = {}, shop, nowIso) {
  const redacted = labelData && typeof labelData === "object" ? { ...labelData } : {};
  redacted.recipientName = "";
  redacted.recipientStreet = "";
  redacted.recipientCity = "";
  redacted.recipientState = "";
  redacted.recipientZip = "";
  redacted.recipientCountry = "";
  redacted.shipideSource = {
    provider: "shopify",
    shop: normalizeShopDomain(shop),
    redacted: true,
    redactedAt: nowIso,
  };
  return redacted;
}

function buildShopifyPrivacyExportMatch(row, label, labelData, match) {
  const source = getShopifySourceFromLabelData(labelData);
  return {
    generation_id: row?.id || null,
    user_id: row?.user_id || null,
    generated_at: row?.created_at || null,
    service_type: row?.service_type || null,
    label_index: Number(label?.index) || null,
    label_id: String(label?.labelId || "").trim() || null,
    tracking_id: String(label?.trackingId || "").trim() || null,
    matching_strategy: match?.strategy || null,
    recipient: {
      name: String(labelData?.recipientName || "").trim() || null,
      street: String(labelData?.recipientStreet || "").trim() || null,
      city: String(labelData?.recipientCity || "").trim() || null,
      state: String(labelData?.recipientState || "").trim() || null,
      zip: String(labelData?.recipientZip || "").trim() || null,
      country: String(labelData?.recipientCountry || "").trim() || null,
    },
    shipide_source: source
      ? {
          provider: source.provider || null,
          shop: source.shop || null,
          orderId: source.orderId || null,
          orderName: source.orderName || null,
          customerId: source.customerId || null,
          customerEmail: source.customerEmail || null,
          importedAt: source.importedAt || null,
          redacted: Boolean(source.redacted),
          redactedAt: source.redactedAt || null,
        }
      : null,
  };
}

function getShopifySourceFromInvoiceItem(item = {}) {
  const metadata = item?.metadata && typeof item.metadata === "object" ? item.metadata : {};
  const source =
    metadata?.shipide_source && typeof metadata.shipide_source === "object"
      ? metadata.shipide_source
      : null;
  if (!source) return null;
  const provider = String(source?.provider || "").trim().toLowerCase();
  if (!provider) return null;
  return {
    provider,
    shop: normalizeShopDomain(source?.shop || ""),
    orderId: String(source?.orderId || "").trim(),
    orderName: String(source?.orderName || "").trim(),
    redacted: Boolean(source?.redacted),
    redactedAt: String(source?.redactedAt || "").trim(),
  };
}

async function processShopifyCustomerPrivacyOperation(topic, shop, payload = {}) {
  const normalizedShop = normalizeShopDomain(shop);
  const nowIso = new Date().toISOString();
  const connections = await listShopifyConnectionsByShop(normalizedShop).catch(() => []);
  const userIds = Array.from(
    new Set(
      connections
        .map((connection) => String(connection?.user_id || "").trim())
        .filter(Boolean)
    )
  );
  const accessToken = await resolveShopifyPrivacyAccessToken(connections);
  const requestedOrderIds = normalizeShopifyComplianceOrderIds(payload);
  let fetchedOrders = [];
  if (requestedOrderIds.length && accessToken) {
    fetchedOrders = await fetchShopifyOrdersByIds(normalizedShop, accessToken, requestedOrderIds).catch(
      () => []
    );
  }
  const targets = buildShopifyPrivacyTargets(normalizedShop, payload, fetchedOrders);
  const historyFilters = {
    select: "id,user_id,created_at,service_type,quantity,total_price,payload,billed_invoice_id,billed_at",
    order: "created_at.desc",
  };
  if (userIds.length) {
    historyFilters.user_id = buildSupabaseInFilter(userIds);
  }
  const historyRows = await fetchAllSupabaseRows(HISTORY_TABLE, historyFilters, { pageSize: 500 });

  const exportMatches = [];
  const strategies = new Set();
  const matchedGenerationIds = new Set();
  const matchedLabelIdsByGeneration = new Map();
  const matchedTrackingIdsByGeneration = new Map();
  let legacyFingerprintMatched = false;
  let redactedHistoryRows = 0;
  let redactedHistoryLabels = 0;

  for (const row of historyRows) {
    const payloadObject = row?.payload && typeof row.payload === "object" ? row.payload : {};
    const labels = Array.isArray(payloadObject?.labels) ? payloadObject.labels : [];

    if (!labels.length) {
      const selectionPayload =
        payloadObject?.selection && typeof payloadObject.selection === "object"
          ? payloadObject.selection
          : payloadObject;
      const match = matchShopifyLabelToTargets(selectionPayload, normalizedShop, targets);
      if (!match.matched) {
        continue;
      }

      strategies.add(match.strategy);
      if (match.strategy === "recipient_fingerprint_legacy") {
        legacyFingerprintMatched = true;
      }

      exportMatches.push(
        buildShopifyPrivacyExportMatch(
          row,
          { index: 1, labelId: null, trackingId: null },
          selectionPayload,
          match
        )
      );

      if (topic === "customers/redact") {
        matchedGenerationIds.add(String(row?.id || "").trim());
        const nextSelectionPayload = buildShopifyRedactedLabelData(selectionPayload, normalizedShop, nowIso);
        const nextPayload =
          payloadObject?.selection && typeof payloadObject.selection === "object"
            ? {
                ...payloadObject,
                selection: nextSelectionPayload,
              }
            : {
                ...payloadObject,
                ...nextSelectionPayload,
              };
        await patchSupabaseRows(
          HISTORY_TABLE,
          {
            id: `eq.${row.id}`,
          },
          {
            payload: nextPayload,
          }
        );
        redactedHistoryRows += 1;
        redactedHistoryLabels += 1;
      }
      continue;
    }

    let rowChanged = false;
    const nextLabels = labels.map((label) => {
      const labelData = label?.data && typeof label.data === "object" ? label.data : {};
      const match = matchShopifyLabelToTargets(labelData, normalizedShop, targets);
      if (!match.matched) {
        return label;
      }

      strategies.add(match.strategy);
      if (match.strategy === "recipient_fingerprint_legacy") {
        legacyFingerprintMatched = true;
      }

      exportMatches.push(buildShopifyPrivacyExportMatch(row, label, labelData, match));

      if (topic !== "customers/redact") {
        return label;
      }

      rowChanged = true;
      matchedGenerationIds.add(String(row?.id || "").trim());
      const generationId = String(row?.id || "").trim();
      const labelId = String(label?.labelId || "").trim();
      const trackingId = String(label?.trackingId || "").trim();
      if (generationId) {
        if (!matchedLabelIdsByGeneration.has(generationId)) {
          matchedLabelIdsByGeneration.set(generationId, new Set());
        }
        if (!matchedTrackingIdsByGeneration.has(generationId)) {
          matchedTrackingIdsByGeneration.set(generationId, new Set());
        }
        if (labelId) matchedLabelIdsByGeneration.get(generationId).add(labelId);
        if (trackingId) matchedTrackingIdsByGeneration.get(generationId).add(trackingId);
      }
      redactedHistoryLabels += 1;
      return {
        ...label,
        labelId: null,
        trackingId: null,
        data: buildShopifyRedactedLabelData(labelData, normalizedShop, nowIso),
      };
    });

    if (topic === "customers/redact" && rowChanged) {
      await patchSupabaseRows(
        HISTORY_TABLE,
        {
          id: `eq.${row.id}`,
        },
        {
          payload: {
            ...payloadObject,
            labels: nextLabels,
          },
        }
      );
      redactedHistoryRows += 1;
    }
  }

  let redactedInvoiceItems = 0;
  if (topic === "customers/redact" && matchedGenerationIds.size) {
    const invoiceItems = await fetchAllSupabaseRows(
      BILLING_INVOICE_ITEMS_TABLE,
      {
        select: "id,invoice_id,generation_id,label_id,tracking_id,metadata",
        generation_id: buildSupabaseInFilter(Array.from(matchedGenerationIds)),
        order: "id.asc",
      },
      { pageSize: 1000 }
    );

    for (const item of invoiceItems) {
      const generationId = String(item?.generation_id || "").trim();
      if (!generationId) continue;
      const matchedLabelIds = matchedLabelIdsByGeneration.get(generationId) || new Set();
      const matchedTrackingIds = matchedTrackingIdsByGeneration.get(generationId) || new Set();
      const labelId = String(item?.label_id || "").trim();
      const trackingId = String(item?.tracking_id || "").trim();
      const itemSource = getShopifySourceFromInvoiceItem(item);
      const generationMatched = matchedGenerationIds.has(generationId);
      const shouldRedact =
        (labelId && matchedLabelIds.has(labelId)) ||
        (trackingId && matchedTrackingIds.has(trackingId)) ||
        (
          itemSource?.provider === "shopify" &&
          itemSource.shop === normalizedShop &&
          (
            (itemSource.orderId && targets.orderIds.has(itemSource.orderId)) ||
            generationMatched
          )
        );
      if (!shouldRedact) {
        continue;
      }

      const metadata = item?.metadata && typeof item.metadata === "object" ? item.metadata : {};
      await patchSupabaseRows(
        BILLING_INVOICE_ITEMS_TABLE,
        {
          id: `eq.${item.id}`,
        },
        {
          recipient_name: null,
          recipient_city: null,
          recipient_country: null,
          tracking_id: null,
          label_id: null,
          metadata: {
            ...metadata,
            shipide_source: {
              ...(itemSource || {}),
              provider: "shopify",
              shop: normalizedShop,
              redacted: true,
              redactedAt: nowIso,
            },
            privacy_redacted: true,
            privacy_redacted_at: nowIso,
            privacy_redaction_provider: "shopify",
            privacy_redaction_shop: normalizedShop,
          },
        }
      );
      redactedInvoiceItems += 1;
    }
  }

  return {
    completed_at: nowIso,
    provider_connections_scanned: connections.length,
    user_accounts_scanned: userIds.length,
    history_scan_scope: userIds.length ? "connected_users" : "full_scan",
    requested_order_ids: Array.from(targets.orderIds),
    shopify_orders: targets.orderExports,
    local_matches: exportMatches,
    matched_history_rows: Array.from(new Set(exportMatches.map((match) => match.generation_id))).length,
    matched_labels: exportMatches.length,
    redacted_history_rows: redactedHistoryRows,
    redacted_history_labels: redactedHistoryLabels,
    redacted_invoice_items: redactedInvoiceItems,
    matching_strategies: Array.from(strategies),
    legacy_recipient_fingerprint_matches: legacyFingerprintMatched,
    note: legacyFingerprintMatched
      ? "Legacy records without explicit Shopify provenance were matched using exact recipient fingerprints. New Shopify imports now persist source provenance for deterministic future compliance handling."
      : "Shopify customer privacy processing completed with stored source provenance and exact identifier matching.",
  };
}

async function processShopifyShopRedactOperation(shop) {
  const normalizedShop = normalizeShopDomain(shop);
  const nowIso = new Date().toISOString();
  const connections = await listShopifyConnectionsByShop(normalizedShop).catch(() => []);
  const historyRows = await fetchAllSupabaseRows(
    HISTORY_TABLE,
    {
      select: "id,user_id,created_at,service_type,quantity,total_price,payload,billed_invoice_id,billed_at",
      order: "created_at.desc",
    },
    { pageSize: 500 }
  );

  const matchedGenerationIds = new Set();
  const matchedLabelIdsByGeneration = new Map();
  const matchedTrackingIdsByGeneration = new Map();
  let redactedHistoryRows = 0;
  let redactedHistoryLabels = 0;

  for (const row of historyRows) {
    const payloadObject = row?.payload && typeof row.payload === "object" ? row.payload : {};
    const labels = Array.isArray(payloadObject?.labels) ? payloadObject.labels : [];

    if (!labels.length) {
      const selectionPayload =
        payloadObject?.selection && typeof payloadObject.selection === "object"
          ? payloadObject.selection
          : payloadObject;
      const source = getShopifySourceFromLabelData(selectionPayload);
      if (!(source?.provider === "shopify" && source.shop === normalizedShop)) {
        continue;
      }

      matchedGenerationIds.add(String(row?.id || "").trim());
      const nextSelectionPayload = buildShopifyRedactedLabelData(selectionPayload, normalizedShop, nowIso);
      const nextPayload =
        payloadObject?.selection && typeof payloadObject.selection === "object"
          ? {
              ...payloadObject,
              selection: nextSelectionPayload,
            }
          : {
              ...payloadObject,
              ...nextSelectionPayload,
            };
      await patchSupabaseRows(
        HISTORY_TABLE,
        {
          id: `eq.${row.id}`,
        },
        {
          payload: nextPayload,
        }
      );
      redactedHistoryRows += 1;
      redactedHistoryLabels += 1;
      continue;
    }

    let rowChanged = false;
    const nextLabels = labels.map((label) => {
      const labelData = label?.data && typeof label.data === "object" ? label.data : {};
      const source = getShopifySourceFromLabelData(labelData);
      if (!(source?.provider === "shopify" && source.shop === normalizedShop)) {
        return label;
      }

      rowChanged = true;
      matchedGenerationIds.add(String(row?.id || "").trim());
      const generationId = String(row?.id || "").trim();
      const labelId = String(label?.labelId || "").trim();
      const trackingId = String(label?.trackingId || "").trim();
      if (generationId) {
        if (!matchedLabelIdsByGeneration.has(generationId)) {
          matchedLabelIdsByGeneration.set(generationId, new Set());
        }
        if (!matchedTrackingIdsByGeneration.has(generationId)) {
          matchedTrackingIdsByGeneration.set(generationId, new Set());
        }
        if (labelId) matchedLabelIdsByGeneration.get(generationId).add(labelId);
        if (trackingId) matchedTrackingIdsByGeneration.get(generationId).add(trackingId);
      }
      redactedHistoryLabels += 1;
      return {
        ...label,
        labelId: null,
        trackingId: null,
        data: buildShopifyRedactedLabelData(labelData, normalizedShop, nowIso),
      };
    });

    if (rowChanged) {
      await patchSupabaseRows(
        HISTORY_TABLE,
        {
          id: `eq.${row.id}`,
        },
        {
          payload: {
            ...payloadObject,
            labels: nextLabels,
          },
        }
      );
      redactedHistoryRows += 1;
    }
  }

  const invoiceItems = await fetchAllSupabaseRows(
    BILLING_INVOICE_ITEMS_TABLE,
    {
      select: "id,invoice_id,generation_id,label_id,tracking_id,metadata",
      order: "id.asc",
    },
    { pageSize: 1000 }
  );

  let redactedInvoiceItems = 0;
  for (const item of invoiceItems) {
    const generationId = String(item?.generation_id || "").trim();
    const matchedLabelIds = generationId ? matchedLabelIdsByGeneration.get(generationId) || new Set() : new Set();
    const matchedTrackingIds =
      generationId ? matchedTrackingIdsByGeneration.get(generationId) || new Set() : new Set();
    const labelId = String(item?.label_id || "").trim();
    const trackingId = String(item?.tracking_id || "").trim();
    const itemSource = getShopifySourceFromInvoiceItem(item);
    const shouldRedact =
      (itemSource?.provider === "shopify" && itemSource.shop === normalizedShop) ||
      (generationId && matchedGenerationIds.has(generationId)) ||
      (labelId && matchedLabelIds.has(labelId)) ||
      (trackingId && matchedTrackingIds.has(trackingId));
    if (!shouldRedact) {
      continue;
    }

    const metadata = item?.metadata && typeof item.metadata === "object" ? item.metadata : {};
    await patchSupabaseRows(
      BILLING_INVOICE_ITEMS_TABLE,
      {
        id: `eq.${item.id}`,
      },
      {
        recipient_name: null,
        recipient_city: null,
        recipient_country: null,
        tracking_id: null,
        label_id: null,
        metadata: {
          ...metadata,
          shipide_source: {
            ...(itemSource || {}),
            provider: "shopify",
            shop: normalizedShop,
            redacted: true,
            redactedAt: nowIso,
          },
          privacy_redacted: true,
          privacy_redacted_at: nowIso,
          privacy_redaction_provider: "shopify",
          privacy_redaction_shop: normalizedShop,
        },
      }
    );
    redactedInvoiceItems += 1;
  }

  const deletedConnections = await deleteShopifyConnectionsByShop(shop).catch(() => []);
  await deleteShopifySessionsByShop(shop).catch(() => {});

  return {
    completed_at: nowIso,
    provider_connections_scanned: connections.length,
    history_scan_scope: "full_scan",
    redacted_history_rows: redactedHistoryRows,
    redacted_history_labels: redactedHistoryLabels,
    redacted_invoice_items: redactedInvoiceItems,
    deleted_shopify_connections: Array.isArray(deletedConnections) ? deletedConnections.length : 0,
    deleted_shopify_sessions_for_shop: normalizedShop,
    note: "Shopify shop redact scrubbed locally stored shipment and invoice records linked to this shop before deleting the live connection and sessions.",
  };
}

async function processShopifyComplianceRequest(topic, shop, payload = {}) {
  const customer =
    payload?.customer && typeof payload.customer === 'object' ? payload.customer : {};
  const requestRecord = await insertPrivacyRequest(
    {
      source: 'shopify_webhook',
      requestType:
        topic === 'shop/redact'
          ? 'shop_redact'
          : topic === 'customers/data_request'
            ? 'customer_data_request'
            : 'customer_redact',
      status: 'processing',
      provider: 'shopify',
      shopDomain: shop,
      externalCustomerId: customer?.id,
      externalCustomerEmail: customer?.email,
      requestPayload: payload && typeof payload === 'object' ? payload : {},
    },
    { allowMissingSchema: true }
  ).catch(() => null);

  let status = 'completed';
  let resultPayload = {
    completed_at: new Date().toISOString(),
  };

  if (topic === 'shop/redact') {
    resultPayload = await processShopifyShopRedactOperation(shop);
  } else {
    resultPayload = await processShopifyCustomerPrivacyOperation(topic, shop, payload);
  }

  if (requestRecord?.id) {
    await updatePrivacyRequest(
      requestRecord.id,
      {
        status,
        processed_at: new Date().toISOString(),
        result_payload: resultPayload,
      },
      { allowMissingSchema: true }
    ).catch(() => null);
  }

  return {
    requestId: requestRecord?.id || null,
    status,
    result: resultPayload,
  };
}

async function runPrivacyMaintenance(options = {}) {
  const now = new Date();
  const nowIso = now.toISOString();
  const inviteCutoffIso = new Date(
    now.getTime() - PRIVACY_INVITE_RETENTION_DAYS * 24 * 60 * 60 * 1000
  ).toISOString();
  const staleConnectionIso = new Date(
    now.getTime() - PRIVACY_STALE_CONNECTION_DAYS * 24 * 60 * 60 * 1000
  ).toISOString();
  const operationalRetentionCutoffIso = new Date(
    now.getTime() - PRIVACY_RETENTION_YEARS * 365 * 24 * 60 * 60 * 1000
  ).toISOString();
  const financialRetentionCutoffIso = new Date(
    now.getTime() - PRIVACY_FINANCIAL_RETENTION_YEARS * 365 * 24 * 60 * 60 * 1000
  ).toISOString();
  const financialRetentionCutoffDate = financialRetentionCutoffIso.slice(0, 10);
  const anonymizedPayload = {
    retention_status: 'anonymized',
    anonymized_at: nowIso,
  };

  const deletedInvites = await deleteSupabaseRows(
    REGISTRATION_INVITES_TABLE,
    {
      expires_at: `lt.${inviteCutoffIso}`,
    },
    { returnRepresentation: true }
  ).catch(() => []);

  const deletedStaleConnections = await deleteSupabaseRows(
    'provider_connections',
    {
      status: 'in.(disconnected,token_invalid)',
      updated_at: `lt.${staleConnectionIso}`,
    },
    { returnRepresentation: true }
  ).catch(() => []);

  const deletedExpiredSessions = await deleteSupabaseRows(
    'Session',
    {
      expires: `lt.${nowIso}`,
    },
    { returnRepresentation: true }
  ).catch(() => []);

  await patchSupabaseRows(
    HISTORY_TABLE,
    {
      created_at: `lt.${operationalRetentionCutoffIso}`,
      billed_invoice_id: 'is.null',
    },
    {
      payload: anonymizedPayload,
    }
  ).catch(() => null);

  await patchSupabaseRows(
    HISTORY_TABLE,
    {
      created_at: `lt.${financialRetentionCutoffIso}`,
      billed_invoice_id: 'not.is.null',
    },
    {
      payload: anonymizedPayload,
    }
  ).catch(() => null);

  const oldInvoices = await fetchAllSupabaseRows(
    BILLING_INVOICES_TABLE,
    {
      select: 'id',
      period_end: `lt.${financialRetentionCutoffDate}`,
    },
    { pageSize: 200 }
  ).catch(() => []);

  const invoiceIds = oldInvoices.map((row) => String(row?.id || '').trim()).filter(Boolean);
  for (const invoiceIdChunk of chunkArray(invoiceIds, 100)) {
    const inFilter = buildSupabaseInFilter(invoiceIdChunk);
    if (!inFilter) continue;
    await patchSupabaseRows(
      BILLING_INVOICES_TABLE,
      {
        id: inFilter,
      },
      {
        company_name: null,
        contact_name: null,
        contact_email: null,
        customer_id: null,
        account_manager: null,
        metadata: anonymizedPayload,
      }
    ).catch(() => null);
    await patchSupabaseRows(
      BILLING_INVOICE_ITEMS_TABLE,
      {
        invoice_id: inFilter,
      },
      {
        recipient_name: null,
        recipient_city: null,
        recipient_country: null,
        tracking_id: null,
        label_id: null,
        metadata: anonymizedPayload,
      }
    ).catch(() => null);
  }

  return {
    ranAt: nowIso,
    legalFinancialRetentionYears: PRIVACY_FINANCIAL_RETENTION_YEARS,
    deletedExpiredInvites: Array.isArray(deletedInvites) ? deletedInvites.length : 0,
    deletedStaleConnections: Array.isArray(deletedStaleConnections)
      ? deletedStaleConnections.length
      : 0,
    deletedExpiredShopifySessions: Array.isArray(deletedExpiredSessions)
      ? deletedExpiredSessions.length
      : 0,
    anonymizedInvoices: invoiceIds.length,
  };
}

async function handleAdminPrivacyMaintenanceRun(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: 'Authentication required.' });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: 'You are not allowed to run privacy maintenance.' });
    return;
  }
  try {
    const result = await runPrivacyMaintenance({
      triggeredBy: user.id,
    });
    sendJson(res, 200, {
      ok: true,
      result,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || 'Privacy maintenance failed.' });
  }
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

async function getShopifyConnectionByShop(shop, options = {}) {
  const { includeSettings = false } = options;
  const normalizedShop = normalizeShopDomain(shop);
  if (!normalizedShop) return null;
  const params = new URLSearchParams();
  params.set(
    "select",
    includeSettings
      ? "user_id,shop_domain,status,scopes,connected_at,updated_at,access_token,import_settings"
      : "user_id,shop_domain,status,scopes,connected_at,updated_at,access_token"
  );
  params.set("provider", "eq.shopify");
  params.set("shop_domain", `eq.${normalizedShop}`);
  params.set("status", "eq.connected");
  params.set("order", "updated_at.desc");
  params.set("limit", "1");

  const response = await supabaseServiceRequest(`/rest/v1/provider_connections?${params.toString()}`, {
    method: "GET",
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed reading Shopify connection (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) return null;
  return rows[0];
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

async function resolveShopifyEmbeddedSessionContext(token) {
  const session = verifyShopifySessionToken(token);
  if (!session?.shop) {
    return null;
  }

  const connection = await getShopifyConnectionByShop(session.shop).catch(() => null);
  if (!connection?.user_id) {
    return {
      session,
      connection: null,
      portalUser: null,
    };
  }

  let portalUser = null;
  try {
    portalUser = await getSupabaseUserById(connection.user_id);
  } catch (_error) {
    portalUser = {
      id: connection.user_id,
      email: "",
    };
  }

  return {
    session,
    connection,
    portalUser,
  };
}

async function upsertWooCommerceConnection(userId, storeUrl, consumerKey, consumerSecret) {
  const encryptedToken = encryptToken(
    JSON.stringify({
      consumerKey: String(consumerKey || "").trim(),
      consumerSecret: String(consumerSecret || "").trim(),
    })
  );
  const nowIso = new Date().toISOString();
  const payload = [
    {
      user_id: userId,
      provider: "woocommerce",
      shop_domain: storeUrl,
      status: "connected",
      scopes: "read_orders",
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
    throw new Error(`Failed saving WooCommerce connection (${response.status}) ${details}`.trim());
  }
}

async function getWooCommerceConnection(userId, requestedStoreUrl, options = {}) {
  const { includeSettings = false } = options;
  const params = new URLSearchParams();
  params.set(
    "select",
    includeSettings
      ? "shop_domain,status,connected_at,updated_at,access_token,import_settings"
      : "shop_domain,status,connected_at,updated_at,access_token"
  );
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.woocommerce");
  params.set("status", "eq.connected");
  if (requestedStoreUrl) {
    params.set("shop_domain", `eq.${requestedStoreUrl}`);
  }
  params.set("order", "updated_at.desc");
  params.set("limit", "1");

  const response = await supabaseServiceRequest(
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
        "WooCommerce settings storage is not enabled yet. Run the latest supabase_provider_connections.sql migration."
      );
      error.code = "missing_import_settings_column";
      throw error;
    }
    throw new Error(`Failed reading WooCommerce connection (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) {
    return null;
  }
  return rows[0];
}

async function upsertWixConnection(userId, instanceId, payload = {}) {
  const encryptedToken = encryptToken(
    JSON.stringify({
      instanceId: String(payload.instanceId || instanceId || "").trim(),
      siteId: String(payload.siteId || "").trim(),
      siteUrl: String(payload.siteUrl || "").trim(),
      permissions: Array.isArray(payload.permissions) ? payload.permissions : [],
      raw: payload.raw && typeof payload.raw === "object" ? payload.raw : null,
    })
  );
  const permissions = Array.isArray(payload.permissions)
    ? payload.permissions.map((value) => String(value || "").trim()).filter(Boolean)
    : [];
  const nowIso = new Date().toISOString();
  const response = await supabaseServiceRequest(
    "/rest/v1/provider_connections?on_conflict=user_id,provider,shop_domain",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify([
        {
          user_id: userId,
          provider: "wix",
          shop_domain: String(instanceId || "").trim(),
          status: "connected",
          scopes: permissions.join(","),
          access_token: encryptedToken,
          updated_at: nowIso,
          connected_at: nowIso,
        },
      ]),
    }
  );

  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed saving Wix connection (${response.status}) ${details}`.trim());
  }
}

async function getWixConnection(userId, requestedInstanceId, options = {}) {
  const { includeAccessToken = false, includeSettings = false } = options;
  const params = new URLSearchParams();
  params.set(
    "select",
    includeAccessToken
      ? includeSettings
        ? "shop_domain,status,scopes,connected_at,updated_at,access_token,import_settings"
        : "shop_domain,status,scopes,connected_at,updated_at,access_token"
      : includeSettings
        ? "shop_domain,status,scopes,connected_at,updated_at,import_settings"
        : "shop_domain,status,scopes,connected_at,updated_at"
  );
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.wix");
  params.set("status", "eq.connected");
  if (requestedInstanceId) {
    params.set("shop_domain", `eq.${requestedInstanceId}`);
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
        "Wix settings storage is not enabled yet. Run the latest supabase_provider_connections.sql migration."
      );
      error.code = "missing_import_settings_column";
      throw error;
    }
    throw new Error(`Failed reading Wix connection (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) {
    return null;
  }
  return rows[0];
}

function getLinkedInConnectionRowKey() {
  return LINKEDIN_CONNECTION_KEY;
}

function normalizeLinkedInOrganizationId(value) {
  const match = /(\d+)$/.exec(String(value || "").trim());
  return match ? match[1] : "";
}

function getLinkedInApiHeaders(accessToken, extraHeaders = {}) {
  return {
    Authorization: `Bearer ${accessToken}`,
    "Linkedin-Version": getLinkedInApiVersion(),
    "X-Restli-Protocol-Version": "2.0.0",
    ...extraHeaders,
  };
}

async function fetchLinkedInJson(accessToken, path, options = {}) {
  const response = await fetch(`https://api.linkedin.com${path}`, {
    ...options,
    headers: getLinkedInApiHeaders(accessToken, {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    }),
  });
  const payload = await response.json().catch(() => null);
  if (!response.ok) {
    const message =
      String(payload?.message || payload?.error_description || payload?.error || "").trim()
      || `LinkedIn request failed (${response.status}).`;
    const error = new Error(message);
    error.status = response.status;
    error.payload = payload;
    throw error;
  }
  return payload;
}

function normalizeLinkedInStoredPayload(payload = {}) {
  const organizations = Array.isArray(payload?.organizations)
    ? payload.organizations
        .map((entry) => {
          const id = normalizeLinkedInOrganizationId(entry?.id || entry?.urn || "");
          const urn = String(entry?.urn || (id ? `urn:li:organization:${id}` : "")).trim();
          if (!id || !urn) return null;
          return {
            id,
            urn,
            name: String(entry?.name || entry?.localizedName || "").trim() || `Organization ${id}`,
            vanityName: String(entry?.vanityName || "").trim(),
            roles: Array.isArray(entry?.roles)
              ? entry.roles.map((value) => String(value || "").trim()).filter(Boolean)
              : [],
          };
        })
        .filter(Boolean)
    : [];
  const selectedOrganizationUrn = String(
    payload?.selectedOrganizationUrn || organizations[0]?.urn || ""
  ).trim();
  return {
    accessToken: String(payload?.accessToken || "").trim(),
    refreshToken: String(payload?.refreshToken || "").trim(),
    expiresAt: String(payload?.expiresAt || "").trim(),
    refreshExpiresAt: String(payload?.refreshExpiresAt || "").trim(),
    scopes: Array.isArray(payload?.scopes)
      ? payload.scopes.map((value) => String(value || "").trim()).filter(Boolean)
      : String(payload?.scope || "")
          .split(/\s+/)
          .map((value) => String(value || "").trim())
          .filter(Boolean),
    organizations,
    selectedOrganizationUrn,
  };
}

function parseStoredLinkedInConnection(encryptedToken) {
  if (!String(encryptedToken || "").trim()) return null;
  const decrypted = decryptToken(encryptedToken);
  if (!decrypted) return null;
  try {
    return normalizeLinkedInStoredPayload(JSON.parse(decrypted));
  } catch (_error) {
    return null;
  }
}

async function upsertLinkedInConnection(userId, payload = {}) {
  const nowIso = new Date().toISOString();
  const normalizedPayload = normalizeLinkedInStoredPayload(payload);
  const connectedAt = String(payload?.connectedAt || payload?.connected_at || nowIso).trim() || nowIso;
  const response = await supabaseServiceRequest(
    "/rest/v1/provider_connections?on_conflict=user_id,provider,shop_domain",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify([
        {
          user_id: userId,
          provider: LINKEDIN_PROVIDER,
          shop_domain: getLinkedInConnectionRowKey(),
          status: LINKEDIN_CONNECTION_STATUS_CONNECTED,
          scopes: normalizedPayload.scopes.join(" "),
          access_token: encryptToken(JSON.stringify(normalizedPayload)),
          updated_at: nowIso,
          connected_at: connectedAt,
        },
      ]),
    }
  );
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed saving LinkedIn connection (${response.status}) ${details}`.trim());
  }
}

async function getLinkedInConnection(userId) {
  const params = new URLSearchParams();
  params.set("select", "shop_domain,status,scopes,connected_at,updated_at,access_token");
  params.set("user_id", `eq.${userId}`);
  params.set("provider", `eq.${LINKEDIN_PROVIDER}`);
  params.set("shop_domain", `eq.${getLinkedInConnectionRowKey()}`);
  params.set("status", `eq.${LINKEDIN_CONNECTION_STATUS_CONNECTED}`);
  params.set("order", "updated_at.desc");
  params.set("limit", "1");
  const response = await supabaseServiceRequest(`/rest/v1/provider_connections?${params.toString()}`, {
    method: "GET",
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed reading LinkedIn connection (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) return null;
  return rows[0];
}

async function listLinkedInOrganizations(accessToken) {
  const collected = new Map();
  let start = 0;
  for (let iteration = 0; iteration < 10; iteration += 1) {
    const aclPayload = await fetchLinkedInJson(
      accessToken,
      `/rest/organizationAcls?q=roleAssignee&state=APPROVED&count=100&start=${start}`,
      { method: "GET", headers: { "Content-Type": "application/json" } }
    );
    const elements = Array.isArray(aclPayload?.elements) ? aclPayload.elements : [];
    elements.forEach((entry) => {
      const role = String(entry?.role || "").trim().toUpperCase();
      if (!LINKEDIN_ALLOWED_POST_ROLES.has(role)) return;
      const urn = String(entry?.organization || entry?.organizationTarget || "").trim();
      const id = normalizeLinkedInOrganizationId(urn);
      if (!urn || !id) return;
      const current = collected.get(id) || {
        id,
        urn,
        name: "",
        vanityName: "",
        roles: [],
      };
      if (!current.roles.includes(role)) {
        current.roles.push(role);
      }
      collected.set(id, current);
    });
    const paging = aclPayload?.paging && typeof aclPayload.paging === "object" ? aclPayload.paging : {};
    const count = Number(paging.count) || elements.length || 0;
    const total = Number(paging.total);
    if (!elements.length || !Number.isFinite(total) || start + count >= total) {
      break;
    }
    start += count;
  }
  const ids = Array.from(collected.keys());
  for (const idChunk of chunkArray(ids, 50)) {
    if (!idChunk.length) continue;
    const lookupPayload = await fetchLinkedInJson(
      accessToken,
      `/rest/organizationsLookup?ids=List(${idChunk.map((value) => encodeURIComponent(value)).join(",")})`,
      { method: "GET", headers: { "Content-Type": "application/json" } }
    );
    const results = lookupPayload?.results && typeof lookupPayload.results === "object" ? lookupPayload.results : {};
    Object.entries(results).forEach(([id, entry]) => {
      const current = collected.get(id);
      if (!current) return;
      current.name =
        String(entry?.localizedName || "").trim()
        || String(entry?.name?.localized?.en_US || "").trim()
        || current.name
        || `Organization ${id}`;
      current.vanityName = String(entry?.vanityName || "").trim() || current.vanityName;
      collected.set(id, current);
    });
  }
  return Array.from(collected.values()).sort((left, right) => left.name.localeCompare(right.name));
}

async function refreshLinkedInConnectionOrganizations(row, stored) {
  const organizations = await listLinkedInOrganizations(stored.accessToken);
  const nextPayload = normalizeLinkedInStoredPayload({
    ...stored,
    organizations,
    selectedOrganizationUrn: organizations.some((entry) => entry.urn === stored.selectedOrganizationUrn)
      ? stored.selectedOrganizationUrn
      : organizations[0]?.urn || "",
  });
  await upsertLinkedInConnection(row.user_id, {
    ...nextPayload,
    connectedAt: row?.connected_at || "",
  });
  return nextPayload;
}

function normalizeWixInstallUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  try {
    const parsed = new URL(raw);
    if (!/^https?:$/i.test(parsed.protocol)) return "";
    return parsed.toString();
  } catch (_error) {
    return "";
  }
}

function normalizeWixPermissions(value) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item || "").trim()).filter(Boolean);
  }
  if (typeof value === "string") {
    return value
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean);
  }
  return [];
}

function decodeWixBase64Url(value) {
  return Buffer.from(String(value || ""), "base64url").toString("utf8");
}

function dataUrlToBytes(value) {
  const match = /^data:([^;,]+)?;base64,(.+)$/i.exec(String(value || "").trim());
  if (!match) {
    throw new Error("Image payload must be a valid base64 data URL.");
  }
  return {
    mimeType: String(match[1] || "application/octet-stream").trim().toLowerCase(),
    bytes: Buffer.from(match[2], "base64"),
  };
}

function getLinkedInApiVersion() {
  return LINKEDIN_API_VERSION || "202602";
}

function getLinkedInRedirectUri(req) {
  if (LINKEDIN_REDIRECT_URI) return LINKEDIN_REDIRECT_URI;
  return `${buildPublicBaseUrl(req)}/api/linkedin/callback`;
}

function getLinkedInScopes() {
  return LINKEDIN_SCOPES.length ? LINKEDIN_SCOPES : ["rw_organization_admin", "w_organization_social"];
}

function assertLinkedInConfig() {
  if (!LINKEDIN_CLIENT_ID) {
    throw new Error("LINKEDIN_CLIENT_ID is required for LinkedIn posting.");
  }
  if (!LINKEDIN_CLIENT_SECRET) {
    throw new Error("LINKEDIN_CLIENT_SECRET is required for LinkedIn posting.");
  }
}

function verifyWixInstanceSignature(signature, encodedData) {
  const secret = String(WIX_APP_SECRET || "").trim();
  if (!secret || !signature || !encodedData) return false;
  const digestBuffer = crypto
    .createHmac("sha256", secret)
    .update(String(encodedData), "utf8")
    .digest();
  const expectedCandidates = [
    digestBuffer.toString("hex"),
    digestBuffer.toString("base64"),
    digestBuffer.toString("base64url"),
  ];
  return expectedCandidates.some((candidate) => {
    const receivedBuffer = Buffer.from(String(signature), "utf8");
    const expectedBuffer = Buffer.from(String(candidate), "utf8");
    return (
      receivedBuffer.length === expectedBuffer.length
      && crypto.timingSafeEqual(receivedBuffer, expectedBuffer)
    );
  });
}

function parseWixInstanceToken(instanceToken) {
  const parts = String(instanceToken || "").trim().split(".");
  if (parts.length !== 2) return null;
  const [signature, encodedData] = parts;
  if (!verifyWixInstanceSignature(signature, encodedData)) {
    return null;
  }
  try {
    const payload = JSON.parse(decodeWixBase64Url(encodedData));
    return payload && typeof payload === "object" ? payload : null;
  } catch (_error) {
    return null;
  }
}

async function fetchWixTokenInfo(instanceToken) {
  const token = String(instanceToken || "").trim();
  if (!token) return null;
  const response = await fetch("https://www.wixapis.com/oauth2/token-info", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ token }),
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    throw new Error(`Failed verifying Wix instance token (${response.status}) ${details}`.trim());
  }
  const payload = await response.json().catch(() => null);
  return payload && typeof payload === "object" ? payload : null;
}

async function resolveWixInstancePayload(instanceToken) {
  const tokenInfo = await fetchWixTokenInfo(instanceToken).catch(() => null);
  const legacyPayload = parseWixInstanceToken(instanceToken);
  const tokenInfoActive = Boolean(tokenInfo && tokenInfo.active);

  if (tokenInfoActive && WIX_APP_ID) {
    const clientId = String(tokenInfo?.clientId || tokenInfo?.client_id || "").trim();
    if (clientId && clientId !== WIX_APP_ID) {
      return null;
    }
  }

  if (!tokenInfoActive && !legacyPayload) {
    return null;
  }

  return {
    ...(legacyPayload && typeof legacyPayload === "object" ? legacyPayload : {}),
    ...(tokenInfo && typeof tokenInfo === "object" ? tokenInfo : {}),
  };
}

function getWixInstanceConnectionPayload(instancePayload) {
  const instanceId = String(
    instancePayload?.instanceId
      || instancePayload?.instance_id
      || instancePayload?.instance?.id
      || instancePayload?.subjectId
      || instancePayload?.subject_id
      || ""
  ).trim();
  const siteId = String(
    instancePayload?.siteId
      || instancePayload?.site_id
      || instancePayload?.metaSiteId
      || instancePayload?.meta_site_id
      || instancePayload?.site?.id
      || instancePayload?.context?.siteId
      || instancePayload?.context?.site_id
      || ""
  ).trim();
  const siteUrl = String(
    instancePayload?.siteUrl
      || instancePayload?.site_url
      || instancePayload?.site?.url
      || instancePayload?.site?.siteUrl
      || instancePayload?.externalBaseUrl
      || instancePayload?.context?.siteUrl
      || instancePayload?.context?.site_url
      || ""
  ).trim();
  const permissions = normalizeWixPermissions(
    instancePayload?.permissions
      || instancePayload?.permission
      || instancePayload?.scope
      || instancePayload?.scopes
  );
  if (!instanceId) {
    return null;
  }
  return {
    instanceId,
    siteId,
    siteUrl,
    permissions,
    raw: instancePayload,
  };
}

function parseStoredWixConnection(token) {
  let decrypted = "";
  try {
    decrypted = decryptToken(token);
  } catch (error) {
    if (error?.code === "token_decrypt_failed") {
      return null;
    }
    throw error;
  }
  try {
    const payload = JSON.parse(decrypted);
    return payload && typeof payload === "object" ? payload : null;
  } catch (_error) {
    return null;
  }
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

function normalizeShopifyFinancialStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return SHOPIFY_FINANCIAL_STATUS_OPTIONS.includes(normalized) ? normalized : "";
}

function normalizeShopifyFinancialStatuses(values) {
  const rawValues = Array.isArray(values)
    ? values
    : values == null || values === ""
      ? []
      : [values];
  const seen = new Set();
  rawValues.forEach((value) => {
    const normalized = normalizeShopifyFinancialStatus(value);
    if (!normalized || seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
  });
  return Array.from(seen);
}

function normalizeWixImportStatuses(values) {
  const allowed = new Set(WIX_IMPORT_STATUS_OPTIONS);
  const seen = new Set();
  if (!Array.isArray(values)) {
    return [];
  }
  values.forEach((value) => {
    const normalized = String(value || "").trim().toUpperCase();
    if (!normalized || !allowed.has(normalized) || seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
  });
  return Array.from(seen);
}

function getWixImportSettings(importSettings) {
  if (!importSettings || typeof importSettings !== "object") {
    return {
      selectedStatuses: [...DEFAULT_WIX_IMPORT_STATUSES],
      autoRefreshEnabled: false,
    };
  }
  const selectedStatuses = normalizeWixImportStatuses(
    Array.isArray(importSettings.selected_statuses)
      ? importSettings.selected_statuses
      : Array.isArray(importSettings.selectedStatuses)
        ? importSettings.selectedStatuses
        : []
  );
  return {
    selectedStatuses: selectedStatuses.length
      ? selectedStatuses
      : [...DEFAULT_WIX_IMPORT_STATUSES],
    autoRefreshEnabled: normalizeWooCommerceAutoRefreshEnabled(
      importSettings.auto_refresh_enabled ?? importSettings.autoRefreshEnabled
    ),
  };
}

function getShopifyImportSettings(importSettings) {
  if (!importSettings || typeof importSettings !== "object") {
    return {
      selectedLocationIds: [],
      selectedFinancialStatuses: [...DEFAULT_SHOPIFY_FINANCIAL_STATUSES],
      autoRefreshEnabled: false,
    };
  }

  const selectedFinancialStatuses = normalizeShopifyFinancialStatuses(
    Array.isArray(importSettings.selected_financial_statuses)
      ? importSettings.selected_financial_statuses
      : Array.isArray(importSettings.selectedFinancialStatuses)
        ? importSettings.selectedFinancialStatuses
        : importSettings.financial_status ?? importSettings.financialStatus
  );

  return {
    selectedLocationIds: getSelectedLocationIdsFromImportSettings(importSettings),
    selectedFinancialStatuses: selectedFinancialStatuses.length
      ? selectedFinancialStatuses
      : [...DEFAULT_SHOPIFY_FINANCIAL_STATUSES],
    autoRefreshEnabled: normalizeWooCommerceAutoRefreshEnabled(
      importSettings.auto_refresh_enabled ?? importSettings.autoRefreshEnabled
    ),
  };
}

function normalizeWooCommerceImportStatuses(values) {
  const allowed = new Set(WOOCOMMERCE_IMPORT_STATUS_OPTIONS);
  const seen = new Set();

  if (!Array.isArray(values)) {
    return [];
  }

  values.forEach((value) => {
    const normalized = String(value || "").trim().toLowerCase();
    if (!normalized || !allowed.has(normalized) || seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
  });

  return Array.from(seen);
}

function normalizeWooCommerceAutoRefreshEnabled(value) {
  if (value === true || value === 1) return true;
  const normalized = String(value || "").trim().toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "on";
}

function getWooCommerceImportSettings(importSettings) {
  if (!importSettings || typeof importSettings !== "object") {
    return {
      selectedStatuses: [...DEFAULT_WOOCOMMERCE_IMPORT_STATUSES],
      autoRefreshEnabled: false,
    };
  }

  const selectedStatuses = normalizeWooCommerceImportStatuses(
    Array.isArray(importSettings.selected_statuses)
      ? importSettings.selected_statuses
      : Array.isArray(importSettings.selectedStatuses)
        ? importSettings.selectedStatuses
        : []
  );

  return {
    selectedStatuses: selectedStatuses.length
      ? selectedStatuses
      : [...DEFAULT_WOOCOMMERCE_IMPORT_STATUSES],
    autoRefreshEnabled: normalizeWooCommerceAutoRefreshEnabled(
      importSettings.auto_refresh_enabled ?? importSettings.autoRefreshEnabled
    ),
  };
}

async function saveShopifyImportSettings(
  userId,
  shop,
  selectedLocationIds,
  selectedFinancialStatuses,
  autoRefreshEnabled
) {
  const normalizedFinancialStatuses = normalizeShopifyFinancialStatuses(selectedFinancialStatuses);
  const resolvedFinancialStatuses = normalizedFinancialStatuses.length
    ? normalizedFinancialStatuses
    : [...DEFAULT_SHOPIFY_FINANCIAL_STATUSES];
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
          selected_financial_statuses: resolvedFinancialStatuses,
          financial_status: resolvedFinancialStatuses[0] || DEFAULT_SHOPIFY_FINANCIAL_STATUS,
          auto_refresh_enabled: normalizeWooCommerceAutoRefreshEnabled(autoRefreshEnabled),
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

async function saveWixImportSettings(
  userId,
  instanceId,
  selectedStatuses,
  autoRefreshEnabled
) {
  const params = new URLSearchParams();
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.wix");
  params.set("shop_domain", `eq.${instanceId}`);
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
          selected_statuses: normalizeWixImportStatuses(selectedStatuses),
          auto_refresh_enabled: Boolean(autoRefreshEnabled),
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
        "Wix settings storage is not enabled yet. Run the latest supabase_provider_connections.sql migration."
      );
      error.code = "missing_import_settings_column";
      throw error;
    }
    throw new Error(`Failed saving Wix settings (${response.status}) ${details}`.trim());
  }
}

async function saveWooCommerceImportSettings(
  userId,
  storeUrl,
  selectedStatuses,
  autoRefreshEnabled
) {
  const params = new URLSearchParams();
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.woocommerce");
  params.set("shop_domain", `eq.${storeUrl}`);
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
          selected_statuses: normalizeWooCommerceImportStatuses(selectedStatuses),
          auto_refresh_enabled: Boolean(autoRefreshEnabled),
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
        "WooCommerce settings storage is not enabled yet. Run the latest supabase_provider_connections.sql migration."
      );
      error.code = "missing_import_settings_column";
      throw error;
    }
    throw new Error(`Failed saving WooCommerce settings (${response.status}) ${details}`.trim());
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

async function setWixConnectionStatus(userId, instanceId, status) {
  const params = new URLSearchParams();
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.wix");
  params.set("shop_domain", `eq.${instanceId}`);

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
    throw new Error(`Failed updating Wix connection (${response.status}) ${details}`.trim());
  }
}

async function setWooCommerceConnectionStatus(userId, storeUrl, status) {
  const params = new URLSearchParams();
  params.set("user_id", `eq.${userId}`);
  params.set("provider", "eq.woocommerce");
  params.set("shop_domain", `eq.${storeUrl}`);

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
    throw new Error(`Failed updating WooCommerce connection (${response.status}) ${details}`.trim());
  }
}

function parseWooCommerceCredentials(token) {
  let decrypted = "";
  try {
    decrypted = decryptToken(token);
  } catch (error) {
    if (error?.code === "token_decrypt_failed") {
      throw createWooCommerceDecryptError();
    }
    throw error;
  }
  let parsed = null;
  try {
    parsed = JSON.parse(decrypted);
  } catch (_error) {
    throw createWooCommerceDecryptError();
  }
  const consumerKey = String(parsed?.consumerKey || "").trim();
  const consumerSecret = String(parsed?.consumerSecret || "").trim();
  if (!consumerKey || !consumerSecret) {
    throw createWooCommerceDecryptError();
  }
  return { consumerKey, consumerSecret };
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

function normalizeImportedRegion(region, country) {
  const normalizedRegion = String(region || "").trim().toUpperCase();
  if (!normalizedRegion) return "";
  const normalizedCountry = String(country || "").trim().toUpperCase();
  const countriesThatNeedRegion = new Set(["US", "CA", "AU", "BR", "MX", "JP"]);
  const shortRegion =
    normalizedCountry && normalizedRegion.startsWith(`${normalizedCountry}-`)
      ? normalizedRegion.slice(normalizedCountry.length + 1)
      : normalizedRegion;
  if (!shortRegion) return "";
  return countriesThatNeedRegion.has(normalizedCountry) ? shortRegion : "";
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
  const { locationById = {}, selectedLocationIds = [], shopDomain = "" } = options;
  if (!Array.isArray(orders)) return [];

  const selectedSet = new Set(sanitizeSelectedLocationIds(selectedLocationIds));
  const singleOverrideId = selectedSet.size === 1 ? selectedSet.values().next().value : "";
  const mirrorSelection = selectedSet.size > 1;
  const normalizedShop = normalizeShopDomain(shopDomain);
  const importedAt = new Date().toISOString();

  return orders.map((order) => {
    const shipping = order?.shipping_address || {};
    const recipientName =
      String(shipping?.name || "").trim() ||
      [shipping?.first_name, shipping?.last_name].filter(Boolean).join(" ").trim();
    const totalWeightGrams = Number(order?.total_weight);
    const weightKg =
      Number.isFinite(totalWeightGrams) && totalWeightGrams > 0
        ? (totalWeightGrams / 1000).toFixed(2)
        : "";

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
    const customerId = String(order?.customer?.id || "").trim();
    const customerEmail = normalizeEmail(order?.email || order?.customer?.email || "");
    const recipientCountry = String(shipping?.country_code || shipping?.country || "").trim();

    return {
      senderName: sender.senderName,
      senderStreet: sender.senderStreet,
      senderCity: sender.senderCity,
      senderState: sender.senderState,
      senderZip: sender.senderZip,
      recipientName,
      recipientStreet: String(shipping?.address1 || "").trim(),
      recipientCity: String(shipping?.city || "").trim(),
      recipientState: normalizeImportedRegion(
        String(shipping?.province_code || shipping?.province || "").trim(),
        recipientCountry
      ),
      recipientZip: String(shipping?.zip || "").trim(),
      recipientCountry,
      packageWeight: weightKg,
      packageDims: "",
      shipideSource: {
        provider: "shopify",
        shop: normalizedShop,
        orderId: String(order?.id || "").trim(),
        orderName: String(order?.name || "").trim(),
        customerId,
        customerEmail,
        importedAt,
      },
    };
  });
}

function getWooCommerceOrderAddress(order) {
  const shipping = order?.shipping && typeof order.shipping === "object" ? order.shipping : {};
  const billing = order?.billing && typeof order.billing === "object" ? order.billing : {};
  const shippingHasAddress = [
    shipping?.first_name,
    shipping?.last_name,
    shipping?.address_1,
    shipping?.city,
    shipping?.country,
  ].some((value) => String(value || "").trim());
  return shippingHasAddress ? shipping : billing;
}

function getWooCommerceSettingValue(settings, settingId) {
  const key = String(settingId || "").trim();
  if (!key) return "";

  const items = Array.isArray(settings)
    ? settings
    : Array.isArray(settings?.settings)
      ? settings.settings
      : [];

  for (const item of items) {
    if (String(item?.id || "").trim() !== key) continue;
    return String(item?.value || "").trim();
  }

  if (settings && typeof settings === "object" && !Array.isArray(settings) && settings[key] != null) {
    return String(settings[key]).trim();
  }

  return "";
}

function parseWooCommerceDefaultCountry(value) {
  const normalized = String(value || "").trim();
  if (!normalized) {
    return { country: "", state: "" };
  }

  const [country = "", state = ""] = normalized.split(":");
  return {
    country: String(country || "").trim().toUpperCase(),
    state: String(state || "").trim(),
  };
}

function getWooCommerceStoreLabel(storeUrl) {
  try {
    return new URL(String(storeUrl || "")).hostname.replace(/^www\./i, "").trim();
  } catch (_error) {
    return "";
  }
}

function buildWooCommerceSenderOrigin(settings, storeUrl) {
  const address1 = getWooCommerceSettingValue(settings, "woocommerce_store_address");
  const address2 = getWooCommerceSettingValue(settings, "woocommerce_store_address_2");
  const city = getWooCommerceSettingValue(settings, "woocommerce_store_city");
  const postcode = getWooCommerceSettingValue(settings, "woocommerce_store_postcode");
  const { state } = parseWooCommerceDefaultCountry(
    getWooCommerceSettingValue(settings, "woocommerce_default_country")
  );
  const street = [address1, address2].filter(Boolean).join(", ");
  const hasOrigin = [street, city, state, postcode].some((value) => String(value || "").trim());

  if (!hasOrigin) {
    return null;
  }

  return {
    senderName: getWooCommerceStoreLabel(storeUrl),
    senderStreet: street,
    senderCity: city,
    senderState: state,
    senderZip: postcode,
  };
}

function normalizeWooCommerceWeightUnit(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["g", "gram", "grams"].includes(normalized)) return "g";
  if (["lbs", "lb", "pound", "pounds"].includes(normalized)) return "lbs";
  if (["oz", "ounce", "ounces"].includes(normalized)) return "oz";
  return "kg";
}

function convertWooCommerceWeightToKg(value, unit = "kg") {
  const normalized = String(value || "").trim().replace(",", ".");
  const amount = Number(normalized);
  if (!Number.isFinite(amount) || amount <= 0) return 0;
  const normalizedUnit = normalizeWooCommerceWeightUnit(unit);
  if (normalizedUnit === "g") return amount / 1000;
  if (normalizedUnit === "lbs") return amount * 0.45359237;
  if (normalizedUnit === "oz") return amount * 0.028349523125;
  return amount;
}

function formatWooCommerceWeightKg(value) {
  const amount = Number(value);
  return Number.isFinite(amount) && amount > 0 ? amount.toFixed(2) : "";
}

function getWooCommerceLineItemWeightKey(item) {
  const productId = String(item?.product_id || "").trim();
  const variationId = String(item?.variation_id || "").trim();
  if (!productId || productId === "0") return "";
  return `${productId}:${variationId === "0" ? "" : variationId}`;
}

function getWooCommerceOrderWeightKg(order, weightByLineItemKey = {}) {
  if (!Array.isArray(order?.line_items)) return 0;
  return order.line_items.reduce((sum, item) => {
    const key = getWooCommerceLineItemWeightKey(item);
    const itemWeightKg = Number(weightByLineItemKey[key] || 0);
    if (!Number.isFinite(itemWeightKg) || itemWeightKg <= 0) return sum;
    const quantity = Math.max(1, Number(item?.quantity || 1) || 1);
    return sum + itemWeightKg * quantity;
  }, 0);
}

async function fetchWooCommerceProductWeightKg(
  storeUrl,
  consumerKey,
  consumerSecret,
  productId,
  variationId,
  weightUnit
) {
  const safeStoreUrl = await assertSafeWooCommerceStoreUrlForFetch(storeUrl);
  const authHeaders = {
    Authorization: `Basic ${encodeBasicAuth(consumerKey, consumerSecret)}`,
    Accept: "application/json",
  };
  const fetchWeight = async (url) => {
    const response = await fetch(url.toString(), {
      method: "GET",
      headers: authHeaders,
    });
    if (!response.ok) return 0;
    const payload = await response.json().catch(() => null);
    return convertWooCommerceWeightToKg(payload?.weight, weightUnit);
  };

  if (variationId) {
    const variationUrl = new URL(
      `${safeStoreUrl.replace(/\/+$/, "")}/wp-json/wc/v3/products/${encodeURIComponent(productId)}/variations/${encodeURIComponent(variationId)}`
    );
    variationUrl.searchParams.set("_fields", "id,weight");
    const variationWeight = await fetchWeight(variationUrl);
    if (variationWeight > 0) return variationWeight;
  }

  const productUrl = new URL(
    `${safeStoreUrl.replace(/\/+$/, "")}/wp-json/wc/v3/products/${encodeURIComponent(productId)}`
  );
  productUrl.searchParams.set("_fields", "id,weight");
  return fetchWeight(productUrl);
}

async function fetchWooCommerceLineItemWeights(
  storeUrl,
  consumerKey,
  consumerSecret,
  orders,
  weightUnit
) {
  const productKeys = new Map();
  (Array.isArray(orders) ? orders : []).forEach((order) => {
    (Array.isArray(order?.line_items) ? order.line_items : []).forEach((item) => {
      const key = getWooCommerceLineItemWeightKey(item);
      if (!key || productKeys.has(key)) return;
      productKeys.set(key, {
        productId: String(item?.product_id || "").trim(),
        variationId:
          String(item?.variation_id || "").trim() === "0"
            ? ""
            : String(item?.variation_id || "").trim(),
      });
    });
  });

  const entries = await Promise.all(
    Array.from(productKeys.entries()).map(async ([key, item]) => {
      const weightKg = await fetchWooCommerceProductWeightKg(
        storeUrl,
        consumerKey,
        consumerSecret,
        item.productId,
        item.variationId,
        weightUnit
      ).catch(() => 0);
      return [key, weightKg];
    })
  );

  return Object.fromEntries(entries.filter(([, weightKg]) => Number(weightKg) > 0));
}

function mapWooCommerceOrdersToCsvRows(
  orders,
  senderOrigin = null,
  selectedStatuses = [],
  weightByLineItemKey = {}
) {
  if (!Array.isArray(orders)) return [];
  const normalizedStatuses = normalizeWooCommerceImportStatuses(selectedStatuses);
  const allowedStatuses = new Set(
    normalizedStatuses.length ? normalizedStatuses : DEFAULT_WOOCOMMERCE_IMPORT_STATUSES
  );
  const sender = {
    senderName: String(senderOrigin?.senderName || "").trim(),
    senderStreet: String(senderOrigin?.senderStreet || "").trim(),
    senderCity: String(senderOrigin?.senderCity || "").trim(),
    senderState: String(senderOrigin?.senderState || "").trim(),
    senderZip: String(senderOrigin?.senderZip || "").trim(),
  };

  return orders
    .filter((order) => {
      const status = String(order?.status || "").trim().toLowerCase();
      return allowedStatuses.has(status);
    })
    .map((order) => {
      const address = getWooCommerceOrderAddress(order);
      const recipientName =
        [address?.first_name, address?.last_name]
          .map((part) => String(part || "").trim())
          .filter(Boolean)
          .join(" ")
          .trim() ||
        String(address?.company || "").trim();
      const recipientStreet = [address?.address_1, address?.address_2]
        .map((part) => String(part || "").trim())
        .filter(Boolean)
        .join(", ");
      const packageWeight = formatWooCommerceWeightKg(
        getWooCommerceOrderWeightKg(order, weightByLineItemKey)
      );
      const recipientCountry = String(address?.country || "").trim();

      return {
        senderName: sender.senderName,
        senderStreet: sender.senderStreet,
        senderCity: sender.senderCity,
        senderState: sender.senderState,
        senderZip: sender.senderZip,
        recipientName,
        recipientStreet,
        recipientCity: String(address?.city || "").trim(),
        recipientState: normalizeImportedRegion(String(address?.state || "").trim(), recipientCountry),
        recipientZip: String(address?.postcode || "").trim(),
        recipientCountry,
        packageWeight,
        packageDims: "",
      };
    })
    .filter((row) => row.recipientName || row.recipientStreet || row.recipientCity);
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

async function fetchShopifyOrders(shop, accessToken, limit, selectedFinancialStatuses) {
  const safeLimit = Number.isFinite(limit) ? Math.max(1, Math.min(250, Math.trunc(limit))) : 50;
  const normalizedFinancialStatuses = normalizeShopifyFinancialStatuses(selectedFinancialStatuses);
  const resolvedFinancialStatuses = normalizedFinancialStatuses.length
    ? normalizedFinancialStatuses
    : [...DEFAULT_SHOPIFY_FINANCIAL_STATUSES];
  const ordersById = new Map();

  await Promise.all(
    resolvedFinancialStatuses.map(async (financialStatus) => {
      const url = new URL(`https://${shop}/admin/api/${SHOPIFY_API_VERSION}/orders.json`);
      url.searchParams.set("status", "open");
      url.searchParams.set("limit", String(safeLimit));
      url.searchParams.set("financial_status", financialStatus);
      url.searchParams.set(
        "fields",
        "id,name,created_at,total_weight,shipping_address,currency,current_total_price,total_price,location_id,origin_location,line_items,fulfillments,email,customer"
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
      orders.forEach((order) => {
        const orderId = String(order?.id || "").trim();
        if (!orderId || ordersById.has(orderId)) {
          return;
        }
        ordersById.set(orderId, order);
      });
    })
  );

  return Array.from(ordersById.values())
    .sort((left, right) => {
      const leftTime = Date.parse(left?.created_at || "") || 0;
      const rightTime = Date.parse(right?.created_at || "") || 0;
      return rightTime - leftTime;
    })
    .slice(0, safeLimit);
}

async function fetchWooCommerceOrders(
  storeUrl,
  consumerKey,
  consumerSecret,
  limit,
  selectedStatuses = []
) {
  const safeLimit = Number.isFinite(limit) ? Math.max(1, Math.min(100, Math.trunc(limit))) : 50;
  const normalizedStatuses = normalizeWooCommerceImportStatuses(selectedStatuses);
  const safeStoreUrl = await assertSafeWooCommerceStoreUrlForFetch(storeUrl);
  const url = new URL(`${safeStoreUrl.replace(/\/+$/, "")}/wp-json/wc/v3/orders`);
  url.searchParams.set("per_page", String(safeLimit));
  url.searchParams.set("orderby", "date");
  url.searchParams.set("order", "desc");
  url.searchParams.set("include_meta", "false");
  url.searchParams.set("_fields", "id,status,billing,shipping,date_created,line_items");
  if (normalizedStatuses.length) {
    url.searchParams.set("status", normalizedStatuses.join(","));
  }

  const response = await fetch(url.toString(), {
    method: "GET",
    headers: {
      Authorization: `Basic ${encodeBasicAuth(consumerKey, consumerSecret)}`,
      Accept: "application/json",
    },
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const error = new Error(`WooCommerce order import failed (${response.status}) ${details}`.trim());
    error.status = response.status;
    if (response.status === 401 || response.status === 403) {
      error.code = "woocommerce_auth_invalid";
    } else if (response.status === 404) {
      error.code = "woocommerce_api_not_found";
    }
    throw error;
  }
  const payload = await response.json().catch(() => null);
  if (!Array.isArray(payload)) {
    throw new Error("WooCommerce order response was invalid.");
  }
  return payload;
}

async function fetchWooCommerceGeneralSettings(storeUrl, consumerKey, consumerSecret) {
  const safeStoreUrl = await assertSafeWooCommerceStoreUrlForFetch(storeUrl);
  const url = new URL(`${safeStoreUrl.replace(/\/+$/, "")}/wp-json/wc/v3/settings/general`);
  url.searchParams.set("_fields", "id,value");

  const response = await fetch(url.toString(), {
    method: "GET",
    headers: {
      Authorization: `Basic ${encodeBasicAuth(consumerKey, consumerSecret)}`,
      Accept: "application/json",
    },
  });
  if (!response.ok) {
    const details = await response.text().catch(() => "");
    const error = new Error(
      `WooCommerce settings fetch failed (${response.status}) ${details}`.trim()
    );
    error.status = response.status;
    if (response.status === 401 || response.status === 403) {
      error.code = "woocommerce_auth_invalid";
    } else if (response.status === 404) {
      error.code = "woocommerce_settings_unavailable";
    }
    throw error;
  }

  const payload = await response.json().catch(() => null);
  if (!Array.isArray(payload) && !Array.isArray(payload?.settings)) {
    throw new Error("WooCommerce settings response was invalid.");
  }
  return payload;
}

async function fetchWooCommerceStoreSenderOrigin(storeUrl, consumerKey, consumerSecret) {
  const settings = await fetchWooCommerceGeneralSettings(storeUrl, consumerKey, consumerSecret);
  return buildWooCommerceSenderOrigin(settings, storeUrl);
}

function normalizeWixOrderStatus(value) {
  const normalized = String(value || "").trim().toUpperCase();
  return WIX_IMPORT_STATUS_OPTIONS.includes(normalized) ? normalized : "";
}

function normalizeWixWeightUnit(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["g", "gram", "grams"].includes(normalized)) return "g";
  if (["lb", "lbs", "pound", "pounds"].includes(normalized)) return "lbs";
  if (["oz", "ounce", "ounces"].includes(normalized)) return "oz";
  return "kg";
}

function convertWixWeightToKg(value, fallbackUnit = "kg") {
  if (value == null || value === "") return 0;
  if (typeof value === "object") {
    const amount = value.value ?? value.amount ?? value.weight ?? value.number;
    const unit = value.unit ?? value.weightUnit ?? value.measurementUnit ?? fallbackUnit;
    return convertWixWeightToKg(amount, unit);
  }
  const amount = Number(String(value).trim().replace(",", "."));
  if (!Number.isFinite(amount) || amount <= 0) return 0;
  const unit = normalizeWixWeightUnit(fallbackUnit);
  if (unit === "g") return amount / 1000;
  if (unit === "lbs") return amount * 0.45359237;
  if (unit === "oz") return amount * 0.028349523125;
  return amount;
}

function formatWixWeightKg(value) {
  const amount = Number(value);
  return Number.isFinite(amount) && amount > 0 ? amount.toFixed(2) : "";
}

function getFirstPositiveWixWeight(candidates, fallbackUnit = "kg") {
  for (const candidate of candidates) {
    const weight = convertWixWeightToKg(candidate, fallbackUnit);
    if (weight > 0) return weight;
  }
  return 0;
}

async function createWixAccessToken(instanceId) {
  const appId = String(WIX_APP_ID || "").trim();
  const appSecret = String(WIX_APP_SECRET || "").trim();
  const normalizedInstanceId = String(instanceId || "").trim();
  if (!appId || !appSecret || !normalizedInstanceId) {
    const error = new Error("Wix OAuth credentials are not configured.");
    error.code = "wix_auth_config_missing";
    throw error;
  }

  const response = await fetch("https://www.wixapis.com/oauth2/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: JSON.stringify({
      grant_type: "client_credentials",
      client_id: appId,
      client_secret: appSecret,
      instance_id: normalizedInstanceId,
    }),
  });
  const payload = await response.json().catch(() => null);
  let resolvedPayload = payload;
  if (payload?.body && typeof payload.body === "string") {
    try {
      resolvedPayload = JSON.parse(payload.body);
    } catch (_error) {
      resolvedPayload = payload;
    }
  }
  const statusCode = Number(payload?.statusCode || response.status);
  if (!response.ok || statusCode >= 400) {
    const error = new Error(
      `Wix access token request failed (${statusCode || response.status})`.trim()
    );
    error.code = "wix_auth_invalid";
    error.status = statusCode || response.status;
    throw error;
  }
  const accessToken = String(resolvedPayload?.access_token || resolvedPayload?.accessToken || "").trim();
  if (!accessToken) {
    const error = new Error("Wix access token response was invalid.");
    error.code = "wix_auth_invalid";
    throw error;
  }
  return accessToken;
}

async function fetchWixJson(accessToken, path, options = {}) {
  const response = await fetch(`https://www.wixapis.com${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...(options.headers || {}),
    },
  });
  const text = await response.text().catch(() => "");
  let payload = null;
  if (text) {
    try {
      payload = JSON.parse(text);
    } catch (_error) {
      payload = null;
    }
  }
  if (!response.ok) {
    const error = new Error(`Wix API request failed (${response.status}) ${text}`.trim());
    error.status = response.status;
    error.payload = payload;
    if (response.status === 401 || response.status === 403) {
      error.code = "wix_auth_invalid";
    }
    throw error;
  }
  return payload;
}

async function fetchWixOrders(accessToken, limit, selectedStatuses = []) {
  const safeLimit = Number.isFinite(limit) ? Math.max(1, Math.min(100, Math.trunc(limit))) : 50;
  const normalizedStatuses = normalizeWixImportStatuses(selectedStatuses);
  const filter = {
    status: {
      $in: normalizedStatuses.length ? normalizedStatuses : [...DEFAULT_WIX_IMPORT_STATUSES],
    },
  };
  const payload = await fetchWixJson(accessToken, "/ecom/v1/orders/search", {
    method: "POST",
    body: JSON.stringify({
      search: {
        filter,
        sort: [{ fieldName: "createdDate", order: "DESC" }],
        cursorPaging: { limit: safeLimit },
      },
    }),
  });
  const orders = Array.isArray(payload?.orders)
    ? payload.orders
    : Array.isArray(payload?.results)
      ? payload.results
      : Array.isArray(payload?.items)
        ? payload.items
        : [];
  return orders.slice(0, safeLimit);
}

async function fetchWixCatalogVersion(accessToken) {
  const payload = await fetchWixJson(accessToken, "/stores/v3/provision/version", {
    method: "GET",
  }).catch(() => null);
  const version = String(payload?.catalogVersion || payload?.version || "").trim().toUpperCase();
  if (version === "V3_CATALOG" || version === "CATALOG_V3") return "V3_CATALOG";
  if (version === "V1_CATALOG" || version === "CATALOG_V1") return "V1_CATALOG";
  return "";
}

function getWixLineItemProductId(item) {
  return String(
    item?.catalogReference?.catalogItemId
      || item?.catalogReference?.productId
      || item?.catalogItemId
      || item?.productId
      || item?.product?.id
      || ""
  ).trim();
}

function getWixLineItemVariantId(item) {
  return String(
    item?.catalogReference?.options?.variantId
      || item?.catalogReference?.options?.variant_id
      || item?.catalogReference?.variantId
      || item?.variantId
      || item?.product?.variantId
      || ""
  ).trim();
}

function getWixLineItemQuantity(item) {
  const quantity = Number(item?.quantity || item?.qty || 1);
  return Number.isFinite(quantity) && quantity > 0 ? quantity : 1;
}

function extractWixLineItemWeightKg(item) {
  return getFirstPositiveWixWeight([
    item?.physicalProperties?.weight,
    item?.physicalProperties?.shippingWeight,
    item?.productDetails?.weight,
    item?.product?.weight,
    item?.weight,
  ]);
}

function extractWixProductWeightKg(product, variantId = "") {
  if (!product || typeof product !== "object") return 0;
  const requestedVariantId = String(variantId || "").trim();
  const variants = [
    ...(Array.isArray(product?.variants) ? product.variants : []),
    ...(Array.isArray(product?.productOptions?.variants) ? product.productOptions.variants : []),
    ...(Array.isArray(product?.variantsInfo?.variants) ? product.variantsInfo.variants : []),
  ];
  if (requestedVariantId && variants.length) {
    const matchedVariant = variants.find((variant) => {
      const id = String(variant?.id || variant?._id || variant?.variantId || "").trim();
      return id === requestedVariantId;
    });
    const variantWeight = getFirstPositiveWixWeight([
      matchedVariant?.physicalProperties?.weight,
      matchedVariant?.physicalProperties?.shippingWeight,
      matchedVariant?.weight,
      matchedVariant?.variant?.weight,
      matchedVariant?.productData?.weight,
    ]);
    if (variantWeight > 0) return variantWeight;
  }
  return getFirstPositiveWixWeight([
    product?.physicalProperties?.weight,
    product?.physicalProperties?.shippingWeight,
    product?.weight,
    product?.shippingWeight,
    product?.product?.physicalProperties?.weight,
    product?.product?.weight,
  ]);
}

function extractWixVariantWeightKg(variant) {
  return getFirstPositiveWixWeight([
    variant?.physicalProperties?.weight,
    variant?.physicalProperties?.shippingWeight,
    variant?.weight,
    variant?.variant?.weight,
    variant?.productData?.weight,
    variant?.productData?.physicalProperties?.weight,
  ]);
}

async function fetchWixV1ProductWeightKg(accessToken, productId, variantId = "") {
  if (!productId) return 0;
  const payload = await fetchWixJson(
    accessToken,
    `/stores-reader/v1/products/${encodeURIComponent(productId)}`,
    { method: "GET" }
  ).catch(() => null);
  const product = payload?.product || payload;
  return extractWixProductWeightKg(product, variantId);
}

async function fetchWixV3ProductWeights(accessToken, productIds = []) {
  const ids = Array.from(new Set(productIds.map((id) => String(id || "").trim()).filter(Boolean)));
  if (!ids.length) return {};
  const result = {};
  const variantPayload = await fetchWixJson(accessToken, "/stores/v3/products/query-variants", {
    method: "POST",
    body: JSON.stringify({
      query: {
        filter: { "productData.productId": { $hasSome: ids } },
        cursorPaging: { limit: Math.min(1000, Math.max(1, ids.length * 20)) },
      },
    }),
  }).catch(() => null);
  const variants = Array.isArray(variantPayload?.variants)
    ? variantPayload.variants
    : Array.isArray(variantPayload?.items)
      ? variantPayload.items
      : [];
  variants.forEach((variant) => {
    const productId = String(
      variant?.productData?.productId || variant?.productId || variant?.product?.id || ""
    ).trim();
    const variantId = String(variant?.id || variant?._id || variant?.variantId || "").trim();
    const weightKg = extractWixVariantWeightKg(variant);
    if (productId && weightKg > 0) {
      result[`${productId}:${variantId}`] = weightKg;
      if (!result[`${productId}:`]) {
        result[`${productId}:`] = weightKg;
      }
    }
  });

  const missingProductIds = ids.filter((id) => !result[`${id}:`]);
  if (missingProductIds.length) {
    const productPayload = await fetchWixJson(accessToken, "/stores/v3/products/query", {
      method: "POST",
      body: JSON.stringify({
        query: {
          filter: { id: { $hasSome: missingProductIds } },
          cursorPaging: { limit: Math.min(100, missingProductIds.length) },
        },
        fields: ["PHYSICAL_PROPERTIES"],
      }),
    }).catch(() => null);
    const products = Array.isArray(productPayload?.products)
      ? productPayload.products
      : Array.isArray(productPayload?.items)
        ? productPayload.items
        : [];
    products.forEach((product) => {
      const productId = String(product?.id || product?._id || product?.productId || "").trim();
      const weightKg = extractWixProductWeightKg(product);
      if (productId && weightKg > 0) {
        result[`${productId}:`] = weightKg;
      }
    });
  }

  return result;
}

async function fetchWixLineItemWeights(accessToken, orders = []) {
  const lineItems = [];
  (Array.isArray(orders) ? orders : []).forEach((order) => {
    (Array.isArray(order?.lineItems) ? order.lineItems : []).forEach((item) => {
      const productId = getWixLineItemProductId(item);
      if (!productId) return;
      lineItems.push({
        productId,
        variantId: getWixLineItemVariantId(item),
      });
    });
  });
  const productIds = Array.from(new Set(lineItems.map((item) => item.productId)));
  if (!productIds.length) return {};
  const catalogVersion = await fetchWixCatalogVersion(accessToken);
  if (catalogVersion === "V3_CATALOG") {
    return fetchWixV3ProductWeights(accessToken, productIds);
  }

  const entries = await Promise.all(
    lineItems.map(async (item) => {
      const key = `${item.productId}:${item.variantId}`;
      const fallbackKey = `${item.productId}:`;
      const weightKg = await fetchWixV1ProductWeightKg(
        accessToken,
        item.productId,
        item.variantId
      ).catch(() => 0);
      return [
        [key, weightKg],
        [fallbackKey, weightKg],
      ];
    })
  );
  const result = {};
  entries.flat().forEach(([key, weightKg]) => {
    if (!result[key] && Number(weightKg) > 0) {
      result[key] = weightKg;
    }
  });
  return result;
}

function getWixOrderWeightKg(order, weightByLineItemKey = {}) {
  const directWeight = getFirstPositiveWixWeight([
    order?.shippingInfo?.logistics?.weight,
    order?.shippingInfo?.logistics?.totalWeight,
    order?.weight,
  ]);
  if (directWeight > 0) return directWeight;
  return (Array.isArray(order?.lineItems) ? order.lineItems : []).reduce((sum, item) => {
    const inlineWeight = extractWixLineItemWeightKg(item);
    const productId = getWixLineItemProductId(item);
    const variantId = getWixLineItemVariantId(item);
    const catalogWeight =
      Number(weightByLineItemKey[`${productId}:${variantId}`] || 0)
      || Number(weightByLineItemKey[`${productId}:`] || 0);
    const itemWeightKg = inlineWeight > 0 ? inlineWeight : catalogWeight;
    if (!Number.isFinite(itemWeightKg) || itemWeightKg <= 0) return sum;
    return sum + itemWeightKg * getWixLineItemQuantity(item);
  }, 0);
}

function getWixAddressValue(address, keys = []) {
  for (const key of keys) {
    const value = key.split(".").reduce((current, part) => current?.[part], address);
    if (String(value || "").trim()) return String(value).trim();
  }
  return "";
}

function getWixOrderShippingAddress(order) {
  return (
    order?.shippingInfo?.logistics?.shippingDestination?.address
    || order?.shippingInfo?.shipmentDetails?.address
    || order?.shippingInfo?.address
    || order?.recipientInfo?.address
    || order?.billingInfo?.address
    || {}
  );
}

function getWixOrderContact(order) {
  return (
    order?.shippingInfo?.logistics?.shippingDestination?.contactDetails
    || order?.shippingInfo?.shipmentDetails?.contactDetails
    || order?.recipientInfo?.contactDetails
    || order?.billingInfo?.contactDetails
    || order?.buyerInfo
    || {}
  );
}

function formatWixStreet(address) {
  const direct = getWixAddressValue(address, [
    "addressLine1",
    "address.line",
    "address.line1",
    "addressLine",
    "line1",
    "street",
    "streetAddress",
    "streetAddress.formattedAddressLine",
    "streetAddress.formattedAddress",
    "streetAddress.addressLine",
    "formattedAddress",
  ]);
  const extra = getWixAddressValue(address, ["addressLine2", "address.line2", "line2"]);
  const streetName = getWixAddressValue(address, [
    "streetAddress.name",
    "streetAddress.street",
    "streetAddress.streetName",
    "streetAddress.route",
  ]);
  const streetNumber = getWixAddressValue(address, [
    "streetAddress.number",
    "streetAddress.streetNumber",
    "streetAddress.houseNumber",
  ]);
  const streetParts = [streetName, streetNumber].filter(Boolean).join(" ");
  return [direct || streetParts, extra].filter(Boolean).join(", ");
}

function mapWixOrdersToCsvRows(orders, options = {}) {
  const { selectedStatuses = [], weightByLineItemKey = {}, instanceId = "", siteUrl = "" } = options;
  const normalizedStatuses = normalizeWixImportStatuses(selectedStatuses);
  const allowedStatuses = new Set(
    normalizedStatuses.length ? normalizedStatuses : DEFAULT_WIX_IMPORT_STATUSES
  );
  const importedAt = new Date().toISOString();
  return (Array.isArray(orders) ? orders : [])
    .filter((order) => {
      const status = normalizeWixOrderStatus(order?.status);
      return !status || allowedStatuses.has(status);
    })
    .map((order) => {
      const address = getWixOrderShippingAddress(order);
      const contact = getWixOrderContact(order);
      const firstName = getWixAddressValue(contact, ["firstName", "first_name"]);
      const lastName = getWixAddressValue(contact, ["lastName", "last_name"]);
      const recipientName =
        [firstName, lastName].filter(Boolean).join(" ").trim()
        || getWixAddressValue(contact, ["fullName", "name", "company"])
        || getWixAddressValue(address, ["contactDetails.fullName", "contactDetails.company"]);
      const recipientCountry = getWixAddressValue(address, ["country", "countryCode"]);
      return {
        senderName: "",
        senderStreet: "",
        senderCity: "",
        senderState: "",
        senderZip: "",
        recipientName,
        recipientStreet: formatWixStreet(address),
        recipientCity: getWixAddressValue(address, ["city", "subdivision", "address.city"]),
        recipientState: normalizeImportedRegion(
          getWixAddressValue(address, ["subdivision", "state", "region"]),
          recipientCountry
        ),
        recipientZip: getWixAddressValue(address, ["postalCode", "zipCode", "postcode", "zip"]),
        recipientCountry,
        packageWeight: formatWixWeightKg(getWixOrderWeightKg(order, weightByLineItemKey)),
        packageDims: "",
        shipideSource: {
          provider: "wix",
          shop: String(siteUrl || instanceId || "").trim(),
          orderId: String(order?.id || order?._id || "").trim(),
          orderName: String(order?.number || order?.displayId || order?.id || "").trim(),
          customerId: String(order?.buyerInfo?.contactId || order?.buyerInfo?.memberId || "").trim(),
          customerEmail: normalizeEmail(order?.buyerInfo?.email || contact?.email || ""),
          importedAt,
        },
      };
    })
    .filter((row) => row.recipientName || row.recipientStreet || row.recipientCity);
}

const DEV_SHOPIFY_ORDER_RECIPIENTS = Object.freeze([
  {
    firstName: "Elise",
    lastName: "Vermeulen",
    company: "Atelier Meridian",
    address1: "1739 Rue du Port",
    city: "Antwerp",
    province: "Antwerp",
    provinceCode: "VAN",
    zip: "2000",
    country: "Belgium",
    countryCode: "BE",
    phone: "+32 470 11 22 33",
  },
  {
    firstName: "Marta",
    lastName: "Dubois",
    company: "Rivage Essentials",
    address1: "17 Herengracht",
    city: "Amsterdam",
    province: "Noord-Holland",
    provinceCode: "NH",
    zip: "1015",
    country: "Netherlands",
    countryCode: "NL",
    phone: "+31 612 34 56 78",
  },
  {
    firstName: "Clara",
    lastName: "Moreau",
    company: "Maison Carmin",
    address1: "44 Rue Sainte-Catherine",
    city: "Bordeaux",
    province: "Gironde",
    provinceCode: "NAQ",
    zip: "33000",
    country: "France",
    countryCode: "FR",
    phone: "+33 6 45 78 11 92",
  },
  {
    firstName: "Lucia",
    lastName: "Rossi",
    company: "Studio Vento",
    address1: "28 Via Torino",
    city: "Milan",
    province: "Lombardy",
    provinceCode: "MI",
    zip: "20123",
    country: "Italy",
    countryCode: "IT",
    phone: "+39 347 55 21 908",
  },
  {
    firstName: "Noah",
    lastName: "Svensson",
    company: "Nordform Goods",
    address1: "9 Stora Nygatan",
    city: "Stockholm",
    province: "Stockholm County",
    provinceCode: "AB",
    zip: "11127",
    country: "Sweden",
    countryCode: "SE",
    phone: "+46 70 000 10 38",
  },
]);

const DEV_SHOPIFY_ORDER_PRODUCTS = Object.freeze([
  { title: "Big Brown Bear Boots", price: 74.99, grams: 820 },
  { title: "Canvas Utility Tote", price: 39.5, grams: 410 },
  { title: "Meridian Cotton Hoodie", price: 62.0, grams: 690 },
  { title: "Studio Glass Water Bottle", price: 28.9, grams: 540 },
  { title: "Nordform Desk Lamp", price: 91.25, grams: 1180 },
]);

function buildBogusShopifySeedOrder(seedIndex) {
  const recipient =
    DEV_SHOPIFY_ORDER_RECIPIENTS[
      Math.abs(Number(seedIndex) || 0) % DEV_SHOPIFY_ORDER_RECIPIENTS.length
    ];
  const product =
    DEV_SHOPIFY_ORDER_PRODUCTS[
      Math.abs(Number(seedIndex) || 0) % DEV_SHOPIFY_ORDER_PRODUCTS.length
    ];
  const quantity = (Math.abs(Number(seedIndex) || 0) % 3) + 1;
  const subtotal = product.price * quantity;
  const taxRate = 0.21;
  const total = subtotal * (1 + taxRate);
  const emailSlug = `${recipient.firstName}.${recipient.lastName}.${seedIndex + 1}`
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ".")
    .replace(/(^\.|\.$)/g, "");
  const address = {
    first_name: recipient.firstName,
    last_name: recipient.lastName,
    company: recipient.company,
    address1: recipient.address1,
    city: recipient.city,
    province: recipient.province,
    province_code: recipient.provinceCode,
    zip: recipient.zip,
    country: recipient.country,
    country_code: recipient.countryCode,
    phone: recipient.phone,
  };

  return {
    email: `${emailSlug}@example.com`,
    currency: "EUR",
    financial_status: "paid",
    send_receipt: false,
    send_fulfillment_receipt: false,
    tags: "shipide-dev-seed,shipide-import-test",
    note: `Shipide dev seed order ${seedIndex + 1}`,
    line_items: [
      {
        title: product.title,
        price: product.price.toFixed(2),
        quantity,
        grams: product.grams,
        taxable: true,
        requires_shipping: true,
      },
    ],
    billing_address: address,
    shipping_address: address,
    transactions: [
      {
        kind: "sale",
        status: "success",
        amount: total.toFixed(2),
        currency: "EUR",
      },
    ],
  };
}

async function createBogusShopifyOrders(shop, accessToken, count) {
  const safeCount = Number.isFinite(count) ? Math.max(1, Math.min(25, Math.trunc(count))) : 10;
  const url = `https://${shop}/admin/api/${SHOPIFY_API_VERSION}/orders.json`;
  const createdOrders = [];

  for (let index = 0; index < safeCount; index += 1) {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "X-Shopify-Access-Token": accessToken,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        order: buildBogusShopifySeedOrder(index),
      }),
    });

    const payloadText = await response.text().catch(() => "");
    let payload = null;
    if (payloadText) {
      try {
        payload = JSON.parse(payloadText);
      } catch {
        payload = null;
      }
    }

    if (!response.ok) {
      const details =
        payload?.errors && typeof payload.errors === "object"
          ? JSON.stringify(payload.errors)
          : payloadText;
      const error = new Error(`Shopify bogus order creation failed (${response.status}) ${details}`.trim());
      error.status = response.status;
      if (/write_orders|scope|permission|access denied/i.test(String(details || ""))) {
        error.code = "missing_write_orders";
      }
      throw error;
    }

    const order = payload?.order;
    if (!order?.id) {
      throw new Error("Shopify bogus order creation succeeded without returning an order id.");
    }
    createdOrders.push({
      id: order.id,
      name: String(order.name || "").trim(),
      email: String(order.email || "").trim(),
    });
  }

  return createdOrders;
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
  if (!OAUTH_STATE_SECRET || !SHOPIFY_TOKEN_ENCRYPTION_KEY) {
    throw new Error("OAUTH_STATE_SECRET and SHOPIFY_TOKEN_ENCRYPTION_KEY are required.");
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
  const companyName = normalizeClientCompanyName(body?.companyName || "");
  if (!companyName) {
    sendJson(res, 400, { error: "Company name is required." });
    return;
  }
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
      companyName,
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
      companyName,
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

async function handleCreateShipmentExtractRequest(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to create shipment extract links." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const clientEmail = normalizeEmail(body?.clientEmail || "");
  const companyName = normalizeClientCompanyName(body?.companyName || "");
  if (!companyName) {
    sendJson(res, 400, { error: "Company name is required." });
    return;
  }
  if (!clientEmail) {
    sendJson(res, 400, { error: "Client email is required." });
    return;
  }
  if (!isValidEmailFormat(clientEmail)) {
    sendJson(res, 400, { error: "Invalid client email format." });
    return;
  }

  try {
    const created = await createShipmentExtractRequest({
      clientEmail,
      companyName,
      expiresInDays: body?.expiresInDays,
      createdBy: user.id,
    });
    const requestUrl = new URL("/clean-data", buildPublicBaseUrl(req));
    requestUrl.searchParams.set("token", created.token);
    sendJson(res, 200, {
      ok: true,
      requestUrl: requestUrl.toString(),
      expiresAt: created.expiresAt,
      clientEmail,
      companyName,
      request: mapShipmentExtractRequestRow(created.row, buildPublicBaseUrl(req)),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not create shipment extract link." });
  }
}

async function handleListShipmentExtractRequests(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to view shipment extract links." });
    return;
  }

  const limit = Number(requestUrl.searchParams.get("limit") || "50");
  try {
    const rows = await listShipmentExtractRequests({ limit });
    sendJson(res, 200, {
      requests: rows.map((row) => mapShipmentExtractRequestRow(row, buildPublicBaseUrl(req))),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load shipment extract links." });
  }
}

async function handleCreateRegistrationInviteFromShipmentExtract(req, res) {
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

  try {
    const extractRequest = await getShipmentExtractRequestById(body?.requestId);
    const invitedEmail = normalizeEmail(extractRequest?.client_email || "");
    if (!extractRequest?.id || !invitedEmail) {
      sendJson(res, 404, { error: "Shipment extract request not found." });
      return;
    }
    if (!isValidEmailFormat(invitedEmail)) {
      sendJson(res, 400, { error: "Invalid client email format." });
      return;
    }
    const companyName = getShipmentExtractCompanyName(extractRequest);
    const created = await createRegistrationInvite({
      invitedEmail,
      companyName,
      expiresInDays: body?.expiresInDays,
      createdBy: user.id,
    });
    if (created?.id) {
      await linkShipmentExtractRequestToRegistrationInvite(extractRequest.id, created.id);
    }
    const inviteUrl = new URL("/register", buildPublicBaseUrl(req));
    inviteUrl.searchParams.set("invite", created.token);
    sendJson(res, 200, {
      ok: true,
      inviteUrl: inviteUrl.toString(),
      expiresAt: created.expiresAt,
      invitedEmail,
      companyName,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not create registration invite." });
  }
}

async function handleShippingDataCleanerRequestValidate(req, res, requestUrl) {
  const token = normalizeInviteToken(requestUrl.searchParams.get("token"));
  if (!token) {
    sendJson(res, 400, { error: "Shipment extract link token is required." });
    return;
  }

  try {
    const request = await getShipmentExtractRequestByToken(token);
    if (!request?.id) {
      sendJson(res, 404, { error: "Shipment extract link not found." });
      return;
    }
    if (shipmentExtractRequestIsExpired(request)) {
      sendJson(res, 410, { error: "Shipment extract link has expired." });
      return;
    }
    sendJson(res, 200, {
      ok: true,
      request: {
        id: request.id,
        clientEmail: normalizeEmail(request.client_email || ""),
        companyName: getShipmentExtractCompanyName(request),
        expiresAt: request.expires_at || null,
        submittedAt: request.submitted_at || null,
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not validate shipment extract link." });
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

async function handleAdminClientBillingSave(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to update client billing settings." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const targetUserId = String(body?.userId || "").trim();
  if (!targetUserId) {
    sendJson(res, 400, { error: "Client id is required." });
    return;
  }

  const rawPaymentMode = String(
    body?.paymentMode || body?.method || (body?.invoiceEnabled === true ? "invoice" : "wallet")
  )
    .trim()
    .toLowerCase();
  const paymentMode =
    rawPaymentMode === "invoice"
      ? "invoice"
      : ["wallet", "account_balance", "account-balance", "balance"].includes(rawPaymentMode)
        ? "wallet"
        : "";
  if (!paymentMode) {
    sendJson(res, 400, { error: "A valid client payment mode is required." });
    return;
  }

  try {
    const billing = await saveClientBillingPreference(user.id, targetUserId, {
      invoice_enabled: paymentMode === "invoice",
      card_enabled: false,
    });
    sendJson(res, 200, {
      ok: true,
      userId: targetUserId,
      billing,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not update client billing settings." });
  }
}

async function handleAdminInvoicesList(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to access billing invoices." });
    return;
  }
  const limit = Math.max(1, Math.min(300, Number(requestUrl.searchParams.get("limit")) || 120));
  const userId = String(requestUrl.searchParams.get("userId") || "").trim();
  const statusFilter = String(requestUrl.searchParams.get("status") || "").trim().toLowerCase();
  const statuses = statusFilter
    ? statusFilter
        .split(",")
        .map((entry) => entry.trim())
        .filter(Boolean)
    : [];
  try {
    const invoices = await listBillingInvoices({
      limit,
      userId,
      statuses,
    });
    sendJson(res, 200, {
      invoices: buildInvoiceListResponseRows(invoices),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load invoices." });
  }
}

function buildAccountingExportBatchId(now = new Date()) {
  const iso = new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString();
  return `acct-${iso.slice(0, 10).replace(/-/g, "")}-${iso.slice(11, 19).replace(/:/g, "")}`;
}

async function handleAdminInvoiceAccountingExport(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to export invoices for accounting." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const invoiceIds = Array.from(
    new Set(
      (Array.isArray(body?.invoiceIds) ? body.invoiceIds : [])
        .map((value) => String(value || "").trim())
        .filter(Boolean)
    )
  );
  if (!invoiceIds.length) {
    sendJson(res, 400, { error: "Select at least one invoice to export." });
    return;
  }
  const requestedFormat = String(body?.format || "csv").trim().toLowerCase() || "csv";
  const exportedAtRaw = String(body?.exportedAt || "").trim();
  const exportedAtDate = exportedAtRaw ? new Date(exportedAtRaw) : new Date();
  if (Number.isNaN(exportedAtDate.getTime())) {
    sendJson(res, 400, { error: "Export date is invalid." });
    return;
  }
  const exportedAt = exportedAtDate.toISOString();
  const actor = normalizeEmail(user.email || "") || user.id;
  const exportBatchId =
    String(body?.batchId || "").trim() || buildAccountingExportBatchId(exportedAtDate);
  try {
    const updatedInvoices = [];
    for (const invoiceId of invoiceIds) {
      const invoice = await getBillingInvoiceById(invoiceId);
      if (!invoice?.id) continue;
      const metadata = invoice?.metadata && typeof invoice.metadata === "object" ? invoice.metadata : {};
      const currentExport =
        metadata?.accounting_export && typeof metadata.accounting_export === "object"
          ? metadata.accounting_export
          : {};
      const nextExport = {
        exported_at: String(currentExport?.exported_at || "").trim() || exportedAt,
        exported_by: String(currentExport?.exported_by || "").trim() || actor,
        export_batch_id: String(currentExport?.export_batch_id || "").trim() || exportBatchId,
        export_format: String(currentExport?.export_format || "").trim().toLowerCase() || requestedFormat,
      };
      const updated = await updateBillingInvoiceFields(invoice.id, {
        metadata: {
          ...metadata,
          accounting_export: nextExport,
        },
      });
      await insertBillingInvoiceEvent(invoice.id, {
        event_type: "updated",
        actor,
        channel: "accounting",
        message: "Exported to accounting ledger.",
        metadata: {
          accounting_export: nextExport,
        },
      });
      updatedInvoices.push(updated || invoice);
    }
    if (!updatedInvoices.length) {
      sendJson(res, 404, { error: "No invoices were updated." });
      return;
    }
    sendJson(res, 200, {
      ok: true,
      batchId: exportBatchId,
      invoices: buildInvoiceListResponseRows(updatedInvoices),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not export invoices for accounting." });
  }
}

async function getApprovedBillingInvoicePdfExport(invoiceId, options = {}) {
  const safeInvoiceId = String(invoiceId || "").trim();
  if (!safeInvoiceId) {
    throw new Error("Invoice id is required.");
  }
  const invoiceWithItems = await getBillingInvoiceById(safeInvoiceId, { withItems: true });
  if (!invoiceWithItems?.id) {
    throw new Error("Invoice not found.");
  }
  const desiredStage = normalizeInvoicePdfVariantStage(options?.reminderStage || 0);
  const storedVariant = await loadStoredBillingInvoicePdfVariant(invoiceWithItems, desiredStage, {
    allowFallback: false,
  }).catch(() => null);
  if (!storedVariant?.bytes) {
    throw new Error("Approved invoice PDF is unavailable in local mode. Generate or send the invoice first.");
  }
  return {
    invoice: invoiceWithItems,
    bytes: storedVariant.bytes,
    filename:
      String(storedVariant.filename || "").trim()
      || buildInvoiceVariantPdfFilename(invoiceWithItems, desiredStage),
  };
}

async function handleAdminInvoiceDetail(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to view invoices." });
    return;
  }
  const invoiceId = String(requestUrl.searchParams.get("invoiceId") || "").trim();
  if (!invoiceId) {
    sendJson(res, 400, { error: "Invoice id is required." });
    return;
  }
  try {
    const invoice = await getBillingInvoiceById(invoiceId, { withItems: true });
    if (!invoice?.id) {
      sendJson(res, 404, { error: "Invoice not found." });
      return;
    }
    sendJson(res, 200, {
      invoice: {
        ...invoice,
        reference: getBillingInvoiceReference(invoice),
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load invoice." });
  }
}

async function handleAdminInvoicePdf(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to download invoice PDFs." });
    return;
  }
  const invoiceId = String(requestUrl.searchParams.get("invoiceId") || "").trim();
  if (!invoiceId) {
    sendJson(res, 400, { error: "Invoice id is required." });
    return;
  }
  try {
    const pdfExport = await getApprovedBillingInvoicePdfExport(invoiceId, { reminderStage: 0 });
    send(res, 200, {
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="${String(pdfExport.filename || "invoice-shipide.pdf").replace(/"/g, "")}"`,
      "Cache-Control": "no-store",
    }, Buffer.from(pdfExport.bytes));
  } catch (error) {
    const message = error?.message || "Could not load invoice PDF.";
    const status = /not found/i.test(message) ? 404 : /unavailable/i.test(message) ? 409 : 500;
    sendJson(res, status, { error: message });
  }
}

async function handleBillingInvoiceDetail(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  const invoiceId = String(requestUrl.searchParams.get("invoiceId") || "").trim();
  if (!invoiceId) {
    sendJson(res, 400, { error: "Invoice id is required." });
    return;
  }
  try {
    let invoice = await getBillingInvoiceById(invoiceId, { withItems: true });
    if (!invoice?.id || String(invoice?.user_id || "").trim() !== String(user.id || "").trim()) {
      sendJson(res, 404, { error: "Invoice not found." });
      return;
    }
    invoice = await ensureBillingInvoicePublicCode(invoice);
    sendJson(res, 200, {
      invoice: {
        ...invoice,
        reference: getBillingInvoiceReference(invoice),
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load invoice." });
  }
}

async function handleAdminInvoicesRun(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to run billing." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const runRequest = normalizeInvoiceRunRequest(body || {});
  try {
    const publicAppUrl = getPublicAppUrl(req);
    const result = await buildInvoicesFromLabelHistory({
      ...runRequest,
      publicAppUrl,
    });
    let reminders = [];
    if (runRequest.mode !== "preview" && runRequest.includeReminders) {
      reminders = await runInvoiceReminders({ limit: 300, publicAppUrl });
    }
    sendJson(res, 200, {
      ok: true,
      run: result,
      reminders,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Billing run failed." });
  }
}

async function handleAdminInvoiceSend(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to send invoices." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const invoiceId = String(body?.invoiceId || "").trim();
  if (!invoiceId) {
    sendJson(res, 400, { error: "Invoice id is required." });
    return;
  }
  try {
    const providedVariants = normalizeProvidedInvoicePdfVariantsFromBody(body);
    const result = await sendBillingInvoiceById(invoiceId, {
      isReminder: false,
      pdfVariants: providedVariants,
      requireApprovedPdf: true,
      publicAppUrl: getPublicAppUrl(req),
    });
    sendJson(res, 200, {
      ok: true,
      skipped: Boolean(result?.skipped),
      invoice: buildInvoiceListResponseRows([result?.invoice || {}])[0] || null,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not send invoice." });
  }
}

function buildAdminBillingTestInvoice(toEmail) {
  const now = new Date();
  const nowIso = now.toISOString();
  return {
    invoice: {
      id: crypto.randomUUID(),
      invoice_kind: "monthly",
      company_name: "Test Client",
      contact_name: "Claire Dupont",
      contact_email: toEmail,
      period_start: nowIso.slice(0, 10),
      period_end: nowIso.slice(0, 10),
      due_at: addUtcDays(now, getInvoiceTermsDays()).toISOString(),
      issued_at: nowIso,
      subtotal_ex_vat: 120,
      vat_amount: 0,
      total_inc_vat: 120,
      vat_rate: DEFAULT_VAT_RATE,
      labels_count: 1,
      line_count: 1,
      payment_mode: "invoice",
      metadata: {
        invoice_profile: {
          companyName: "Atelier Meridian",
          contactName: "Claire Dupont",
          contactEmail: toEmail,
          contactPhone: "+32 2 555 01 29",
          billingAddress: "Avenue Louise 120, 1050 Brussels, Belgium",
          taxId: "BE0123456789",
          customerId: "SHIPIDE-TEST",
          accountManager: "Shipide Billing",
        },
      },
    },
    items: [
      {
        service_type: "Economy",
        recipient_name: "Sample Recipient",
        recipient_city: "Brussels",
        recipient_country: "Belgium",
        tracking_id: "TRK-50482912",
        label_id: "LBL-201948",
        amount_ex_vat: 120,
        amount_inc_vat: 120,
      },
    ],
  };
}

function buildAdminTopupBillingTestInvoice(toEmail) {
  const now = new Date();
  const nowIso = now.toISOString();
  const topupSeed = {
    id: crypto.randomUUID(),
    reference: "SHIP-TEST-TOPUP-2604",
    credited_at: nowIso,
    amount_cents: 500,
    currency: "EUR",
  };
  const invoiceNumber = buildTopupInvoiceNumber(topupSeed);
  return {
    invoice: {
      id: crypto.randomUUID(),
      invoice_kind: "topup",
      source_topup_id: topupSeed.id,
      invoice_number: invoiceNumber,
      company_name: "Atelier Meridian",
      contact_name: "Claire Dupont",
      contact_email: toEmail,
      period_start: nowIso.slice(0, 10),
      period_end: nowIso.slice(0, 10),
      due_at: nowIso,
      issued_at: nowIso,
      subtotal_ex_vat: 5,
      vat_amount: 0,
      total_inc_vat: 5,
      vat_rate: 0,
      labels_count: 1,
      line_count: 1,
      payment_mode: "wallet",
      payment_reference: topupSeed.reference,
      payment_received_amount: 5,
      metadata: {
        invoice_kind: "topup",
        invoice_number: invoiceNumber,
        source_topup_id: topupSeed.id,
        topup_reference: topupSeed.reference,
        invoice_profile: {
          companyName: "Atelier Meridian",
          contactName: "Claire Dupont",
          contactEmail: toEmail,
          contactPhone: "+32 2 555 01 29",
          billingAddress: "Avenue Louise 120, 1050 Brussels, Belgium",
          taxId: "BE0123456789",
          customerId: "SHIPIDE-TEST",
          accountManager: "Shipide Billing",
        },
      },
    },
    items: [
      {
        service_type: "Prepaid Credit for Shipping Labels",
        quantity: 1,
        amount_ex_vat: 5,
        amount_inc_vat: 5,
        metadata: {
          topup_reference: topupSeed.reference,
        },
      },
    ],
  };
}

function normalizeProvidedInvoicePdfVariantsFromBody(body = {}) {
  const rows = Array.isArray(body?.pdfVariants) ? body.pdfVariants : [];
  return rows
    .map((entry) => {
      const pdfBase64 = String(entry?.pdfBase64 || "").trim();
      if (!pdfBase64) return null;
      return {
        stage: normalizeInvoicePdfVariantStage(entry?.reminderStage),
        filename: String(entry?.filename || "").trim() || null,
        bytes: Buffer.from(pdfBase64, "base64"),
      };
    })
    .filter(Boolean);
}

async function sendAdminBillingTestSequenceEmails(toEmail, options = {}) {
  const { invoice, items } = buildAdminBillingTestInvoice(toEmail);
  const providedVariants = Array.isArray(options?.pdfVariants) ? options.pdfVariants : [];
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
    const manualSettlement = invoiceRequiresManualSettlement(invoice);
    const issuedAt = parseIsoTimestamp(invoice.issued_at || new Date().toISOString()) || new Date().toISOString();
    const dueAt = manualSettlement
      ? (
          step.isReminder
            ? parseIsoTimestamp(invoice.due_at)
            : addUtcDays(new Date(issuedAt), getInvoiceTermsDays()).toISOString()
        )
      : null;
    const paidAt = manualSettlement ? null : issuedAt;
    const emailInvoice = {
      ...invoice,
      issued_at: issuedAt,
      due_at: dueAt,
      paid_at: paidAt,
    };
    const providedVariant =
      providedVariants.find((entry) => entry?.stage === normalizeInvoicePdfVariantStage(step.reminderStage))
      || null;
    const pdfBytes = providedVariant?.bytes || null;
    if (!pdfBytes) {
      throw new Error(`Approved invoice PDF is unavailable for sequence stage ${step.reminderStage}.`);
    }
    try {
      const response = await sendResendEmail({
        to: toEmail,
        subject,
        html: buildInvoiceEmailHtml(emailInvoice, items, {
          isReminder: step.isReminder,
          reminderStage: step.reminderStage,
        }),
        text: buildInvoiceEmailText(emailInvoice, {
          isReminder: step.isReminder,
        }),
        attachments: [
          {
            filename:
              String(providedVariant?.filename || "").trim()
              || buildInvoiceVariantPdfFilename(emailInvoice, step.reminderStage),
            content: pdfBytes,
            contentType: "application/pdf",
          },
        ],
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

async function handleAdminInvoiceSendTest(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to send billing tests." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const toEmail = normalizeEmail(body?.toEmail || user.email || "");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    sendJson(res, 400, { error: "A valid test email is required." });
    return;
  }
  try {
    const { invoice: testInvoice, items: testItems } = buildAdminBillingTestInvoice(toEmail);
    const providedPdfBase64 = String(body?.pdfBase64 || "").trim();
    const providedFilename = String(body?.filename || "").trim();
    let pdfBytes = providedPdfBase64
      ? Buffer.from(providedPdfBase64, "base64")
      : null;
    if (!pdfBytes) {
      pdfBytes = await buildInvoicePdf(testInvoice, testItems, {
        issuedAt: testInvoice.issued_at,
        dueAt: testInvoice.due_at,
      });
    }
    const resendResponse = await sendResendEmail({
      to: toEmail,
      subject: buildInvoiceEmailSubject(testInvoice, { isReminder: false, reminderStage: 0 }),
      html: buildInvoiceEmailHtml(testInvoice, testItems, {
        isReminder: false,
        reminderStage: 0,
      }),
      text: buildInvoiceEmailText(testInvoice, {
        isReminder: false,
      }),
      attachments: [
        {
          filename: providedFilename || buildInvoicePdfFilename(testInvoice),
          content: pdfBytes,
          contentType: "application/pdf",
        },
      ],
    });
    sendJson(res, 200, {
      ok: true,
      to: toEmail,
      resendId: resendResponse?.id || null,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not send test invoice email." });
  }
}

async function handleAdminTopupInvoiceSendTest(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to send billing tests." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const toEmail = normalizeEmail(body?.toEmail || user.email || "");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    sendJson(res, 400, { error: "A valid test email is required." });
    return;
  }
  try {
    const { invoice: testInvoice, items: testItems } = buildAdminTopupBillingTestInvoice(toEmail);
    const providedPdfBase64 = String(body?.pdfBase64 || "").trim();
    const providedFilename = String(body?.filename || "").trim();
    const pdfBytes = providedPdfBase64
      ? Buffer.from(providedPdfBase64, "base64")
      : await buildInvoicePdf(testInvoice, testItems, {
          issuedAt: testInvoice.issued_at,
          paidAt: testInvoice.issued_at,
        });
    const resendResponse = await sendResendEmail({
      to: toEmail,
      subject: buildInvoiceEmailSubject(testInvoice, { isReminder: false, reminderStage: 0 }),
      html: buildInvoiceEmailHtml(testInvoice, testItems, {
        isReminder: false,
        reminderStage: 0,
      }),
      text: buildInvoiceEmailText(testInvoice, {
        isReminder: false,
      }),
      attachments: [
        {
          filename: providedFilename || buildInvoicePdfFilename(testInvoice),
          content: pdfBytes,
          contentType: "application/pdf",
        },
      ],
    });
    sendJson(res, 200, {
      ok: true,
      to: toEmail,
      resendId: resendResponse?.id || null,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not send test top-up invoice email." });
  }
}

async function handleAdminInvoiceSendTestSequence(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to send billing tests." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const toEmail = normalizeEmail(body?.toEmail || user.email || "");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    sendJson(res, 400, { error: "A valid test email is required." });
    return;
  }
  try {
    const sequence = await sendAdminBillingTestSequenceEmails(toEmail, {
      pdfVariants: normalizeProvidedInvoicePdfVariantsFromBody(body),
    });
    sendJson(res, 200, {
      ok: sequence.failed_count === 0,
      ...sequence,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not send follow-up sequence." });
  }
}

async function handleAdminReportsSendTest(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to send report tests." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const toEmail = normalizeEmail(body?.toEmail || user.email || "");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    sendJson(res, 400, { error: "A valid test email is required." });
    return;
  }
  const profitAmountRaw = Number(body?.profitAmount);
  const profitAmount = Number.isFinite(profitAmountRaw) ? Math.max(0, profitAmountRaw) : 1248.4;
  const reportsUrl =
    String(body?.reportsUrl || REPORTS_PORTAL_URL || "").trim() ||
    "https://portal.shipide.com/reports?range=monthly";
  try {
    const resendResponse = await sendResendEmail({
      to: toEmail,
      fromName: REPORTS_FROM_NAME,
      fromEmail: REPORTS_FROM_EMAIL,
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
    sendJson(res, 200, {
      ok: true,
      to: toEmail,
      resendId: resendResponse?.id || null,
      profitAmount,
      reportsUrl,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not send test reports email." });
  }
}

async function handleAdminAgreementPreviewTest(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to preview agreement emails." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const toEmail = normalizeEmail(body?.toEmail || user.email || "");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    sendJson(res, 400, { error: "A valid test email is required." });
    return;
  }
  try {
    const payload = await buildAcceptedAgreementTestPayload(req, toEmail, {
      includePdf: false,
    });
    const subject = buildAcceptedAgreementEmailSubject(payload.contract);
    sendJson(res, 200, {
      ok: true,
      to: toEmail,
      subject,
      html: buildAcceptedAgreementPreviewDocument(
        subject,
        buildAcceptedAgreementEmailHtml({
          contract: payload.contract,
          artifact: payload.artifact,
          recipientName: payload.recipientName,
        })
      ),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not build agreement email preview." });
  }
}

async function handleAdminAgreementSendTest(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to send agreement tests." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const toEmail = normalizeEmail(body?.toEmail || user.email || "");
  if (!toEmail || !isValidEmailFormat(toEmail)) {
    sendJson(res, 400, { error: "A valid test email is required." });
    return;
  }
  try {
    const payload = await buildAcceptedAgreementTestPayload(req, toEmail, {
      includePdf: true,
    });
    const resendResponse = await sendAcceptedAgreementEmail({
      email: toEmail,
      contract: payload.contract,
      artifact: payload.artifact,
      pdfBytes: payload.pdfBytes,
      recipientName: payload.recipientName,
    });
    sendJson(res, 200, {
      ok: true,
      to: toEmail,
      resendId: resendResponse?.id || null,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not send test agreement email." });
  }
}

async function handleAdminInvoiceMarkPaid(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to update invoice payment." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  const invoiceId = String(body?.invoiceId || "").trim();
  if (!invoiceId) {
    sendJson(res, 400, { error: "Invoice id is required." });
    return;
  }
  const paymentReference = String(body?.paymentReference || "").trim();
  const paidAtRaw = String(body?.paidAt || "").trim();
  const paidAt = paidAtRaw ? new Date(paidAtRaw).toISOString() : new Date().toISOString();
  try {
    const invoice = await getBillingInvoiceById(invoiceId);
    if (!invoice?.id) {
      sendJson(res, 404, { error: "Invoice not found." });
      return;
    }
    const updated = await updateBillingInvoiceFields(invoice.id, {
      status: "paid",
      paid_at: paidAt,
      payment_reference: paymentReference || null,
      payment_received_amount: fromCents(toCents(invoice.total_inc_vat)),
    });
    await insertBillingInvoiceEvent(invoice.id, {
      event_type: "paid",
      actor: normalizeEmail(user.email || "") || user.id,
      channel: "manual",
      message: "Marked as paid by admin.",
      metadata: {
        payment_reference: paymentReference || null,
      },
    });
    sendJson(res, 200, {
      ok: true,
      invoice: buildInvoiceListResponseRows([updated || invoice])[0] || null,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not mark invoice as paid." });
  }
}

async function handleWiseWebhook(req, res) {
  let rawBody = Buffer.alloc(0);
  try {
    rawBody = await readRequestBodyBuffer(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }
  try {
    const accepted = await acceptWiseWebhookDelivery(rawBody, {
      signature: req.headers["x-signature-sha256"],
      deliveryId: req.headers["x-delivery-id"],
      subscriptionId: req.headers["x-subscription-id"],
      schemaVersion: req.headers["x-schema-version"],
      isTest: String(req.headers["x-test-notification"] || "").trim().toLowerCase() === "true",
    });
    sendJson(res, 202, {
      ok: true,
      deliveryId: accepted.deliveryId,
      eventType: accepted.eventType,
    });
    setImmediate(() => {
      void processWiseWebhookAcceptedEvent(accepted).catch((error) => {
        console.error("[wise webhook]", error);
      });
    });
  } catch (error) {
    const statusCode = Number(error?.statusCode) || 500;
    sendJson(res, statusCode, { error: error.message || "Could not accept Wise webhook." });
  }
}

async function handleAdminWiseReceiptsList(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to access Wise reconciliation." });
    return;
  }
  const limit = Math.max(1, Math.min(200, Number(requestUrl.searchParams.get("limit")) || 80));
  const statuses = String(requestUrl.searchParams.get("status") || "manual_review,pending")
    .split(",")
    .map((entry) => String(entry || "").trim().toLowerCase())
    .filter(Boolean);
  try {
    const receipts = await listBillingBankReceipts({
      limit,
      statuses,
      allowMissing: true,
    });
    sendJson(res, 200, {
      wise: {
        ...getWiseConfigSummary(),
        receipts: receipts.map(mapBillingBankReceiptRow).filter(Boolean),
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load Wise reconciliation queue." });
  }
}

async function handleAdminWiseSync(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to sync Wise reconciliation." });
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
    const summary = await syncWiseBalanceStatement({
      lookbackDays: Number(body?.lookbackDays) || getWiseStatementLookbackDays(),
      actor: normalizeEmail(user.email || "") || user.id,
    });
    const topupInvoiceBackfill = await runCreditedTopupInvoiceBackfill({
      limit: 1,
      publicAppUrl: getPublicAppUrl(req),
    });
    const receipts = await listBillingBankReceipts({
      limit: 80,
      statuses: ["manual_review", "pending"],
      allowMissing: true,
    });
    sendJson(res, 200, {
      ok: true,
      summary: {
        ...summary,
        topup_invoice_backfill: topupInvoiceBackfill,
      },
      wise: {
        ...getWiseConfigSummary(),
        receipts: receipts.map(mapBillingBankReceiptRow).filter(Boolean),
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not sync Wise statement." });
  }
}

async function handleAdminWiseReceiptResolve(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "You are not allowed to update Wise reconciliation." });
    return;
  }
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const receiptId = String(body?.receiptId || "").trim();
  const resolution = String(body?.action || body?.resolution || "").trim().toLowerCase();
  const target = String(body?.target || "").trim();
  const note = String(body?.note || "").trim();
  if (!receiptId) {
    sendJson(res, 400, { error: "Bank receipt id is required." });
    return;
  }
  if (!["topup", "invoice", "ignore"].includes(resolution)) {
    sendJson(res, 400, { error: "Resolution must be topup, invoice, or ignore." });
    return;
  }

  try {
    const receipt = await getBillingBankReceiptById(receiptId, { allowMissing: false });
    if (!receipt?.id) {
      sendJson(res, 404, { error: "Bank receipt not found." });
      return;
    }
    if (resolution === "ignore") {
      const updated = await applyBillingBankReceiptResolution({
        receiptId,
        resolution,
        actor: normalizeEmail(user.email || "") || user.id,
        note,
      });
      sendJson(res, 200, {
        ok: true,
        receipt: mapBillingBankReceiptRow(updated || receipt),
      });
      return;
    }

    const targetRow = await resolveBillingBankReceiptTarget(receipt, resolution, target);
    if (!targetRow?.id) {
      sendJson(res, 400, {
        error:
          resolution === "topup"
            ? "Provide a valid SHIP- reference or top-up id."
            : "Provide a valid INV- reference or invoice id.",
      });
      return;
    }
    if (resolution === "invoice") {
      const outstandingCents = Math.max(
        0,
        toCents(targetRow?.total_inc_vat) - toCents(targetRow?.payment_received_amount)
      );
      if (outstandingCents <= 0) {
        sendJson(res, 400, {
          error: "This invoice is already fully paid.",
        });
        return;
      }
      if (outstandingCents > 0 && Number(receipt?.amount_cents || 0) > outstandingCents) {
        sendJson(res, 400, {
          error:
            "This bank receipt is larger than the invoice outstanding balance. Use a wallet credit or review it manually.",
        });
        return;
      }
    }

    const updated = await applyBillingBankReceiptResolution({
      receiptId,
      resolution,
      actor: normalizeEmail(user.email || "") || user.id,
      topupId: resolution === "topup" ? targetRow.id : null,
      invoiceId: resolution === "invoice" ? targetRow.id : null,
      note,
    });
    let topupInvoiceWarning = "";
    if (resolution === "topup" && targetRow?.id) {
      try {
        await ensureBillingTopupInvoiceAndSend(targetRow.id, {
          publicAppUrl: getPublicAppUrl(req),
        });
      } catch (topupInvoiceError) {
        topupInvoiceWarning =
          topupInvoiceError?.message || "Wallet was credited, but the top-up invoice email could not be sent.";
      }
    }
    sendJson(res, 200, {
      ok: true,
      receipt: mapBillingBankReceiptRow(updated || receipt),
      warning: topupInvoiceWarning || null,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not resolve bank receipt." });
  }
}

async function handleBillingOverview(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  try {
    const billingPref = await getClientBillingPreferenceForUser(user.id);
    const paymentMode = getClientPaymentMode(billingPref);
    const settings = await getAdminSettings();
    const currentMonth = getCurrentMonthWindow();
    const invoices = await listBillingInvoices({
      userId: user.id,
      limit: 120,
      allowMissing: true,
    });
    let wallet = {
      user_id: user.id,
      balance_cents: 0,
      currency: DEFAULT_BILLING_CURRENCY,
      updated_at: null,
    };
    let pendingTopups = [];
    let recentTopups = [];
    let walletTransactions = [];
    try {
      [wallet, pendingTopups, recentTopups, walletTransactions] = await Promise.all([
        getOrCreateBillingWallet(user.id),
        listBillingTopups({
          userId: user.id,
          statuses: ["pending", "received"],
          limit: 80,
          allowMissing: true,
        }),
        listBillingTopups({
          userId: user.id,
          limit: 30,
          allowMissing: true,
        }),
        listBillingWalletTransactions({
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
    const projectedExVat = await computeCurrentMonthUnbilledTotal(user.id);
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
    const pendingTopupCents = pendingTopups.reduce((sum, topup) => {
      if (getEffectiveTopupStatus(topup) === "expired") return sum;
      return sum + Math.max(0, Number(topup?.amount_cents) || 0);
    }, 0);
    const walletBalanceCents = Math.max(0, Number(wallet?.balance_cents) || 0);

    sendJson(res, 200, {
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
      recent_topups: recentTopups.slice(0, 30).map(mapBillingTopupRow),
      iban_instructions: getBillingIbanConfig(),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load billing overview." });
  }
}

async function handleBillingTopups(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  const limit = Math.max(1, Math.min(100, Number(requestUrl.searchParams.get("limit")) || 20));
  try {
    const topups = await listBillingTopups({
      userId: user.id,
      limit,
      allowMissing: true,
    });
    sendJson(res, 200, {
      topups: topups.map(mapBillingTopupRow),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load top-up history." });
  }
}

async function handleBillingTopupRequest(req, res) {
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
  const rawAmount = body?.amount ?? body?.amountEur ?? null;
  try {
    const result = await createBillingTopupRequest(user, rawAmount);
    sendJson(res, 200, {
      ok: true,
      ...result,
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not create top-up request." });
  }
}

async function handleBillingCheckout(req, res) {
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
  const method = String(body?.method || "").trim().toLowerCase();
  const amount = Number(body?.amount || body?.amountEur || body?.amountInclVat);
  if (!Number.isFinite(amount) || amount <= 0) {
    sendJson(res, 400, { error: "A valid checkout amount is required." });
    return;
  }
  if (!["invoice", "card", "wallet"].includes(method)) {
    sendJson(res, 400, { error: "Invalid payment method." });
    return;
  }
  try {
    const billingPref = await getClientBillingPreferenceForUser(user.id);
    if (method === "invoice" && !billingPref.invoice_enabled) {
      sendJson(res, 403, { error: "Invoice payment is not enabled for this account." });
      return;
    }
    if (method === "card" && !billingPref.card_enabled) {
      sendJson(res, 403, { error: "Card payment is not enabled for this account." });
      return;
    }
    if (method === "wallet") {
      const result = await debitWalletForCheckout(user, amount, {
        labels_count: Number(body?.labelsCount) || 0,
        service: String(body?.service || "").trim(),
      });
      sendJson(res, 200, {
        ok: true,
        method: "wallet",
        charged_amount_eur: fromCents(toCents(amount)),
        wallet_balance_eur: fromCents(Math.max(0, Number(result.wallet?.balance_cents) || 0)),
        transaction_reference: result.transaction_reference,
      });
      return;
    }
    if (method === "card") {
      const authId = `card_${Date.now().toString(36)}_${crypto
        .randomBytes(3)
        .toString("hex")}`;
      sendJson(res, 200, {
        ok: true,
        method: "card",
        charged_amount_eur: fromCents(toCents(amount)),
        authorization_id: authId,
      });
      return;
    }
    sendJson(res, 200, {
      ok: true,
      method: "invoice",
      charged_amount_eur: fromCents(toCents(amount)),
      note: "Queued for month-end invoice.",
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not process checkout." });
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

async function handleRegisterContractPreview(req, res) {
  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid request body." });
    return;
  }

  const token = normalizeInviteToken(body?.token);
  const email = normalizeEmail(body?.email);
  const profile = normalizeRegistrationProfile(body?.profile);
  const requiredProfileValues = [
    profile.companyName,
    profile.contactName,
    profile.contactEmail,
    profile.contactPhone,
    profile.billingAddress,
    profile.taxId,
  ];

  if (!token) {
    sendJson(res, 400, { error: "Registration link required." });
    return;
  }
  if (!email || !isValidEmailFormat(email)) {
    sendJson(res, 400, { error: "A valid email address is required." });
    return;
  }
  if (requiredProfileValues.some((value) => !String(value || "").trim())) {
    sendJson(res, 400, { error: "All registration fields are required except Customer ID." });
    return;
  }

  try {
    let contract = null;
    let invite = null;

    if (token === LOCAL_SIGNUP_PREVIEW_TOKEN) {
      contract = getLocalSignupPreviewContract();
      invite = getLocalSignupPreviewInvite();
    } else {
      contract = await getActiveClickwrapContract();
      invite = await getRegistrationInviteByToken(token);
      if (!invite || invite.claimed_at || inviteIsRevoked(invite) || inviteIsExpired(invite)) {
        sendJson(res, 410, { error: "This registration link is invalid or expired." });
        return;
      }
    }

    pruneClickwrapPreviewStore();
    const recordId = createClickwrapPreviewRecordId();
    const pdfBytes = await buildClickwrapPreviewPdf({
      contract,
      email,
      profile,
      ipAddress: getRequestIpAddress(req),
      recordId,
      acceptedAt: null,
    });
    const previewEntry = {
      recordId,
      createdAt: Date.now(),
      expiresAt: Date.now() + CLICKWRAP_PREVIEW_TTL_MS,
      contractVersion: String(contract?.version || "").trim(),
      contractHash: String(contract?.hash_sha256 || "").trim().toLowerCase(),
      contractSnapshot: normalizeClickwrapContractSnapshot(contract),
      inviteId: invite?.id || null,
      profileSnapshot: buildClickwrapPreviewProfileSnapshot(email, profile),
      pdfBytes,
      pdfSha256: crypto.createHash("sha256").update(pdfBytes).digest("hex"),
      downloadName: buildClickwrapPreviewDownloadName(contract, recordId),
    };
    clickwrapPreviewStore.set(recordId, previewEntry);
    pruneClickwrapPreviewStore();

    sendJson(res, 200, {
      ok: true,
      contract: buildClickwrapPreviewPublicContract(contract, previewEntry),
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not prepare agreement preview." });
  }
}

async function handleRegisterContractPreviewFile(_req, res, requestUrl) {
  const recordId = normalizeClickwrapPreviewRecordId(requestUrl.searchParams.get("recordId"));
  if (!recordId) {
    sendJson(res, 400, { error: "Agreement preview id required." });
    return;
  }
  const previewRecord = getClickwrapPreviewRecord(recordId);
  if (!previewRecord?.pdfBytes) {
    sendJson(res, 404, { error: "Agreement preview not found or expired." });
    return;
  }
  send(
    res,
    200,
    {
      "Content-Type": "application/pdf",
      "Cache-Control": "no-store",
      "Content-Disposition": `inline; filename="${previewRecord.downloadName}"`,
    },
    previewRecord.pdfBytes
  );
}

async function handleRegisterInviteValidate(_req, res, requestUrl) {
  const token = normalizeInviteToken(requestUrl.searchParams.get("token"));
  if (!token) {
    sendJson(res, 400, { error: "Registration link required." });
    return;
  }
  try {
    const contract = await getActiveClickwrapContract();
    const invite = await getRegistrationInviteByToken(token);
    if (!invite || invite.claimed_at || inviteIsRevoked(invite) || inviteIsExpired(invite)) {
      sendJson(res, 410, { error: "This registration link is invalid or expired." });
      return;
    }
    sendJson(res, 200, {
      valid: true,
      invite: mapInviteToPublic(invite),
      contract: mapClickwrapContractToPublic(contract),
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
  const agreement = normalizeClickwrapAgreement(body?.agreement);
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

  let createdUser = null;
  try {
    const previewRecord = agreement?.recordId ? getClickwrapPreviewRecord(agreement.recordId) : null;
    const activeContract = await getActiveClickwrapContract();
    const contract = previewRecord?.contractSnapshot
      ? normalizeClickwrapContractSnapshot(previewRecord.contractSnapshot)
      : activeContract;
    const agreementError = validateClickwrapAgreement(agreement, contract, {
      previewRecord,
      email,
      profile,
    });
    if (agreementError) {
      sendJson(res, 400, { error: agreementError });
      return;
    }
    const finalAgreement = {
      ...agreement,
      recordId: agreement?.recordId || createClickwrapPreviewRecordId(),
    };

    const invite = await getRegistrationInviteByToken(token);
    if (!invite || invite.claimed_at || inviteIsRevoked(invite) || inviteIsExpired(invite)) {
      sendJson(res, 410, { error: "This registration link is invalid or expired." });
      return;
    }
    const inviteEmail = getInviteBoundEmail(invite);
    if (!inviteEmailMatches(invite, email)) {
      sendJson(res, 400, {
        error: inviteEmail
          ? `This registration link is only valid for ${inviteEmail}.`
          : "This registration link is invalid.",
      });
      return;
    }

    const metadata = {
      ...buildRegistrationMetadata(invite, profile, email, preferredLanguage),
      registration_terms_version: contract.version,
      registration_terms_hash: contract.hash_sha256,
      registration_terms_accepted_at: finalAgreement.agreedAt,
    };
    createdUser = await createSupabaseUserWithMetadata(email, password, metadata);

    await ensureAccountSettingsRow(createdUser.id);
    const acceptedContractProfile = {
      ...profile,
      customerId:
        profile.customerId
        || String(invite?.customer_id || "").trim()
        || String(createdUser?.id || "").trim(),
    };
    const acceptedContractPdfBytes = await buildClickwrapPreviewPdf({
      contract,
      email,
      profile: acceptedContractProfile,
      ipAddress: getRequestIpAddress(req),
      recordId: finalAgreement.recordId,
      acceptedAt: finalAgreement.agreedAt,
    });
    const acceptedContractArtifact = await persistAcceptedClickwrapContract({
      contract,
      recordId: finalAgreement.recordId,
      pdfBytes: acceptedContractPdfBytes,
      acceptedAt: finalAgreement.agreedAt,
    });
    const proof = buildClickwrapEvidence(
      req,
      invite,
      contract,
      finalAgreement,
      createdUser.id,
      email,
      {
        previewRecord,
        acceptedContract: acceptedContractArtifact,
      }
    );
    const acceptance = await insertClickwrapAcceptance({
      user_id: createdUser.id,
      invite_id: invite.id || null,
      accepted_email: email,
      invited_email: normalizeEmail(invite?.invited_email || ""),
      contract_id: contract.id || null,
      contract_version: contract.version,
      contract_hash_sha256: contract.hash_sha256,
      agreement_method: "scroll_clickwrap",
      scrolled_to_end: true,
      scrolled_to_end_at: finalAgreement.scrolledToEndAt,
      agreed_at: finalAgreement.agreedAt,
      ip_address: proof.payload.ip_address || null,
      user_agent: proof.payload.user_agent || null,
      client_timezone: finalAgreement.clientTimezone || null,
      client_locale: finalAgreement.clientLocale || null,
      proof_digest_sha256: proof.digest,
      proof_signature_hmac: proof.signature,
      proof_payload: proof.payload,
    });
    await markRegistrationInviteClaimed(invite.id, createdUser.id, email);

    let agreementEmail = {
      attempted: false,
      sent: false,
      resendId: null,
      error: "",
    };
    if (RESEND_API_KEY && RESEND_FROM_EMAIL) {
      agreementEmail.attempted = true;
      try {
        const resendResponse = await sendAcceptedAgreementEmail({
          email,
          contract,
          artifact: acceptedContractArtifact,
          pdfBytes: acceptedContractPdfBytes,
          recipientName: profile.contactName || invite?.contact_name || "",
        });
        agreementEmail.sent = true;
        agreementEmail.resendId = String(resendResponse?.id || "").trim() || null;
      } catch (emailError) {
        agreementEmail.error = String(emailError?.message || "Could not send agreement email.");
        console.warn("Could not send accepted agreement email", agreementEmail.error);
      }
    }

    if (previewRecord?.recordId) {
      clickwrapPreviewStore.delete(previewRecord.recordId);
    }

    sendJson(res, 200, {
      ok: true,
      email,
      userId: createdUser.id,
      agreementReceiptId: acceptance?.id || null,
      agreementRecordId: finalAgreement.recordId,
      agreementEmail,
    });
  } catch (error) {
    if (createdUser?.id) {
      await deleteSupabaseUserById(createdUser.id).catch(() => {});
    }
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

async function handleShopifyPublicConfig(_req, res) {
  sendJson(res, 200, {
    shopifyApiKey: SHOPIFY_API_KEY || "",
    appBridgeCdnUrl: "https://cdn.shopify.com/shopifycloud/app-bridge.js",
    embedded: true,
  });
}

async function handleShopifyEmbeddedSession(req, res, requestUrl) {
  assertShopifyAppConfig();
  const sessionToken = getBearerToken(req);
  if (!sessionToken) {
    sendJson(
      res,
      401,
      { error: "Shopify session token required." },
      { "X-Shopify-Retry-Invalid-Session-Request": "1" }
    );
    return;
  }

  const context = await resolveShopifyEmbeddedSessionContext(sessionToken);
  if (!context?.session?.shop) {
    sendJson(
      res,
      401,
      { error: "Shopify session token was invalid or expired." },
      { "X-Shopify-Retry-Invalid-Session-Request": "1" }
    );
    return;
  }

  const requestedShop = normalizeShopDomain(requestUrl.searchParams.get("shop"));
  if (requestedShop && requestedShop !== context.session.shop) {
    sendJson(
      res,
      401,
      { error: "Shopify session token did not match the requested shop." },
      { "X-Shopify-Retry-Invalid-Session-Request": "1" }
    );
    return;
  }

  sendJson(res, 200, {
    authenticated: true,
    shop: context.session.shop,
    connected: Boolean(context.connection?.user_id),
    connection: context.connection
      ? {
          shop: context.connection.shop_domain,
          scopes: context.connection.scopes || "",
          connectedAt: context.connection.connected_at || "",
        }
      : null,
    portalUser: context.portalUser
      ? {
          id: context.portalUser.id,
          email: normalizeEmail(context.portalUser.email || ""),
        }
      : null,
    portalUrl: `${getPublicAppUrl(req)}/label-info`,
  });
}


async function handleShopifyComplianceWebhook(req, res) {
  assertShopifyAppConfig();
  const rawBody = await readTextBody(req);
  const topic = String(req.headers['x-shopify-topic'] || '').trim().toLowerCase();
  const shop = normalizeShopDomain(req.headers['x-shopify-shop-domain'] || '');
  const receivedHmac = String(req.headers['x-shopify-hmac-sha256'] || '').trim();

  if (!SHOPIFY_COMPLIANCE_WEBHOOK_TOPICS.includes(topic)) {
    sendJson(res, 400, { error: 'Unsupported Shopify compliance webhook topic.' });
    return;
  }
  if (!verifyShopifyWebhookHmac(rawBody, receivedHmac)) {
    sendJson(res, 401, { error: 'Invalid Shopify webhook signature.' });
    return;
  }

  let payload = {};
  try {
    payload = rawBody ? JSON.parse(rawBody) : {};
  } catch (_error) {
    payload = {};
  }

  try {
    const result = await processShopifyComplianceRequest(topic, shop, payload);
    sendJson(res, 200, {
      ok: true,
      topic,
      shop,
      result,
    });
  } catch (error) {
    sendJson(res, 500, {
      error: error.message || 'Failed to process Shopify compliance webhook.',
    });
  }
}

async function handleWixInstallLink(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  const installUrl = normalizeWixInstallUrl(WIX_APP_INSTALL_URL);
  if (!installUrl) {
    sendJson(res, 500, {
      error:
        "WIX_APP_INSTALL_URL is not configured yet. Add the Wix app install link before connecting a site.",
    });
    return;
  }
  sendJson(res, 200, { url: installUrl });
}

async function handleWixConnection(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  try {
    const row = await getWixConnection(user.id, "", { includeAccessToken: true });
    if (!row) {
      sendJson(res, 200, { connected: false, connection: null });
      return;
    }
    const stored = parseStoredWixConnection(row.access_token);
    sendJson(res, 200, {
      connected: true,
      connection: {
        instanceId: row.shop_domain,
        siteId: String(stored?.siteId || "").trim(),
        siteUrl: String(stored?.siteUrl || "").trim(),
        status: row.status,
        scopes: row.scopes || "",
        connectedAt: row.connected_at || "",
        updatedAt: row.updated_at || "",
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Failed to load Wix connection." });
  }
}

async function handleWixLinkInstance(req, res) {
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

  const instanceToken = String(body?.instance || "").trim();
  if (!instanceToken) {
    sendJson(res, 400, { error: "Missing Wix instance token." });
    return;
  }

  const instancePayload = await resolveWixInstancePayload(instanceToken);
  const connectionPayload = getWixInstanceConnectionPayload(instancePayload);
  if (!connectionPayload?.instanceId) {
    sendJson(res, 400, {
      error: "The Wix instance token was invalid or could not be verified.",
    });
    return;
  }

  try {
    await upsertWixConnection(user.id, connectionPayload.instanceId, connectionPayload);
    sendJson(res, 200, {
      connected: true,
      connection: {
        instanceId: connectionPayload.instanceId,
        siteId: connectionPayload.siteId,
        siteUrl: connectionPayload.siteUrl,
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not link Wix connection." });
  }
}

async function handleWixSettingsGet(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }

  const requestedInstanceId = String(requestUrl.searchParams.get("instanceId") || "").trim();
  try {
    const connection = await getWixConnection(user.id, requestedInstanceId, {
      includeSettings: true,
    });
    if (!connection) {
      sendJson(res, 404, { error: "Wix is not connected for this account." });
      return;
    }
    sendJson(res, 200, {
      instanceId: connection.shop_domain,
      settings: getWixImportSettings(connection.import_settings),
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Failed to load Wix settings." });
  }
}

async function handleWixSettingsPost(req, res) {
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

  const requestedInstanceId = String(body?.instanceId || "").trim();
  const selectedStatuses = normalizeWixImportStatuses(body?.selectedStatuses);
  const autoRefreshEnabled = normalizeWooCommerceAutoRefreshEnabled(body?.autoRefreshEnabled);
  if (!selectedStatuses.length) {
    sendJson(res, 400, { error: "Select at least one Wix status to import." });
    return;
  }

  try {
    const connection = await getWixConnection(user.id, requestedInstanceId, {
      includeSettings: true,
    });
    if (!connection) {
      sendJson(res, 404, { error: "Wix is not connected for this account." });
      return;
    }
    await saveWixImportSettings(
      user.id,
      connection.shop_domain,
      selectedStatuses,
      autoRefreshEnabled
    );
    sendJson(res, 200, {
      instanceId: connection.shop_domain,
      settings: {
        selectedStatuses,
        autoRefreshEnabled,
      },
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Failed to save Wix settings." });
  }
}

async function handleWixDisconnect(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (_error) {
    body = {};
  }

  const requestedInstanceId = String(body?.instanceId || "").trim();
  try {
    const connection = await getWixConnection(user.id, requestedInstanceId);
    if (!connection) {
      sendJson(res, 200, { disconnected: false, connection: null });
      return;
    }
    await deleteWixConnection(user.id, connection.shop_domain);
    sendJson(res, 200, {
      disconnected: true,
      connection: {
        instanceId: connection.shop_domain,
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Failed to disconnect Wix." });
  }
}

async function handleWixImportOrders(req, res) {
  let authUserId = "";
  let connectionInstanceId = "";
  try {
    const user = await getAuthenticatedUser(req);
    if (!user?.id) {
      sendJson(res, 401, { error: "Authentication required." });
      return;
    }
    authUserId = user.id;

    const body = await readJsonBody(req);
    const requestedInstanceId = String(body?.instanceId || body?.shop || "").trim();
    const requestedStatuses = normalizeWixImportStatuses(body?.selectedStatuses);
    const limit = Number(body?.limit);
    let connection = null;
    let settings = {
      selectedStatuses: [...DEFAULT_WIX_IMPORT_STATUSES],
      autoRefreshEnabled: false,
    };
    try {
      connection = await getWixConnection(user.id, requestedInstanceId, {
        includeAccessToken: true,
        includeSettings: true,
      });
      if (connection) {
        settings = getWixImportSettings(connection.import_settings);
      }
    } catch (error) {
      if (error?.code !== "missing_import_settings_column") {
        throw error;
      }
      connection = await getWixConnection(user.id, requestedInstanceId, {
        includeAccessToken: true,
      });
    }
    if (!connection) {
      sendJson(res, 404, { error: "Wix is not connected for this account." });
      return;
    }
    connectionInstanceId = String(connection.shop_domain || requestedInstanceId || "").trim();
    const stored = parseStoredWixConnection(connection.access_token);
    const instanceId = String(stored?.instanceId || connectionInstanceId || "").trim();
    const selectedStatuses = requestedStatuses.length ? requestedStatuses : settings.selectedStatuses;
    const accessToken = await createWixAccessToken(instanceId);
    const orders = await fetchWixOrders(accessToken, limit, selectedStatuses);
    const lineItemWeights = await fetchWixLineItemWeights(accessToken, orders);
    const rows = mapWixOrdersToCsvRows(orders, {
      selectedStatuses,
      weightByLineItemKey: lineItemWeights,
      instanceId,
      siteUrl: stored?.siteUrl || "",
    });
    sendJson(res, 200, {
      instanceId,
      siteUrl: String(stored?.siteUrl || "").trim(),
      count: rows.length,
      rows,
      settings: {
        selectedStatuses,
        autoRefreshEnabled: settings.autoRefreshEnabled,
      },
    });
  } catch (error) {
    if (
      (error?.code === "token_decrypt_failed" || error?.code === "wix_auth_invalid") &&
      authUserId &&
      connectionInstanceId
    ) {
      await setWixConnectionStatus(authUserId, connectionInstanceId, "token_invalid").catch(
        () => {}
      );
      sendJson(res, 409, {
        error: "Wix credentials expired or were rejected. Reconnect Wix and try again.",
      });
      return;
    }
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error?.message || "Wix import failed." });
  }
}

async function handleWooCommerceInstallLink(req, res) {
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

  const storeUrl = normalizeWooCommerceStoreUrl(body?.storeUrl);
  const selectedStatuses = normalizeWooCommerceImportStatuses(body?.selectedStatuses);
  const autoRefreshEnabled = normalizeWooCommerceAutoRefreshEnabled(body?.autoRefreshEnabled);
  if (!storeUrl) {
    sendJson(res, 400, { error: "Enter a valid WooCommerce store URL." });
    return;
  }
  if (!selectedStatuses.length) {
    sendJson(res, 400, { error: "Select at least one WooCommerce status to import." });
    return;
  }
  const publicBaseUrl = buildPublicBaseUrl(req);
  if (!/^https:\/\//i.test(publicBaseUrl)) {
    sendJson(res, 400, {
      error:
        "WooCommerce authorization requires Shipide to be served over HTTPS. Use the deployed portal URL for this flow.",
    });
    return;
  }

  try {
    const safeStoreUrl = await assertSafeWooCommerceStoreUrlForFetch(storeUrl);
    const stateToken = makeSignedStateToken({
      type: "woocommerce_auth",
      userId: user.id,
      storeUrl: safeStoreUrl,
      selectedStatuses,
      autoRefreshEnabled,
    });
    const installUrl = new URL(`${safeStoreUrl.replace(/\/+$/, "")}/wc-auth/v1/authorize`);
    const callbackUrl = new URL("/api/woocommerce/callback", publicBaseUrl);
    callbackUrl.searchParams.set("state", stateToken);
    installUrl.searchParams.set("app_name", WOOCOMMERCE_APP_NAME);
    installUrl.searchParams.set("scope", "read");
    installUrl.searchParams.set("user_id", stateToken);
    installUrl.searchParams.set(
      "return_url",
      getCallbackRedirect(req, {
        provider: "woocommerce",
        store: safeStoreUrl,
      })
    );
    installUrl.searchParams.set("callback_url", callbackUrl.toString());

    sendJson(res, 200, {
      url: installUrl.toString(),
      storeUrl: safeStoreUrl,
    });
  } catch (error) {
    if (error?.code === "woocommerce_store_url_disallowed") {
      sendJson(res, 400, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Could not start WooCommerce connect flow." });
  }
}

async function handleWooCommerceCallback(req, res, requestUrl) {
  const stateToken = String(requestUrl.searchParams.get("state") || "").trim();
  const statePayload = parseSignedStateToken(stateToken);
  if (
    !statePayload?.userId ||
    statePayload.type !== "woocommerce_auth" ||
    !normalizeWooCommerceStoreUrl(statePayload.storeUrl)
  ) {
    sendJson(res, 400, { error: "WooCommerce authorization session expired. Start again." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (error) {
    sendJson(res, 400, { error: error.message || "Invalid WooCommerce callback payload." });
    return;
  }

  const callbackUserId = String(body?.user_id || "").trim();
  const consumerKey = String(body?.consumer_key || "").trim();
  const consumerSecret = String(body?.consumer_secret || "").trim();
  const keyPermissions = String(body?.key_permissions || "").trim().toLowerCase();
  let storeUrl = normalizeWooCommerceStoreUrl(statePayload.storeUrl);
  const importSettings = getWooCommerceImportSettings(statePayload);

  if (!consumerKey || !consumerSecret) {
    sendJson(res, 400, { error: "WooCommerce callback did not include API keys." });
    return;
  }
  if (callbackUserId && callbackUserId !== stateToken) {
    sendJson(res, 400, { error: "WooCommerce callback did not match the authorization request." });
    return;
  }
  if (keyPermissions && !/read/.test(keyPermissions)) {
    sendJson(res, 400, {
      error: "WooCommerce granted write-only permissions. Shipide needs read or read_write access.",
    });
    return;
  }

  try {
    storeUrl = await assertSafeWooCommerceStoreUrlForFetch(storeUrl);
    await fetchWooCommerceOrders(storeUrl, consumerKey, consumerSecret, 1);
    await upsertWooCommerceConnection(
      statePayload.userId,
      storeUrl,
      consumerKey,
      consumerSecret
    );
    await saveWooCommerceImportSettings(
      statePayload.userId,
      storeUrl,
      importSettings.selectedStatuses,
      importSettings.autoRefreshEnabled
    );
    sendJson(res, 200, {
      ok: true,
      storeUrl,
    });
  } catch (error) {
    if (error?.code === "woocommerce_auth_invalid") {
      sendJson(res, 400, {
        error:
          "WooCommerce credentials were rejected. Double-check the store URL and try authorizing again.",
      });
      return;
    }
    if (error?.code === "woocommerce_store_url_disallowed") {
      sendJson(res, 400, { error: error.message });
      return;
    }
    if (error?.code === "woocommerce_api_not_found") {
      sendJson(res, 400, {
        error:
          "WooCommerce API endpoint was not found. Check the store URL and confirm the store exposes /wp-json/wc/v3.",
      });
      return;
    }
    sendJson(res, 500, { error: error.message || "Could not finalize WooCommerce connection." });
  }
}

async function handleWooCommerceConnection(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  try {
    const row = await getWooCommerceConnection(user.id, "");
    if (!row) {
      sendJson(res, 200, { connected: false, connection: null });
      return;
    }
    sendJson(res, 200, {
      connected: true,
      connection: {
        storeUrl: row.shop_domain,
        status: row.status,
        connectedAt: row.connected_at || "",
        updatedAt: row.updated_at || "",
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Failed to load WooCommerce connection." });
  }
}

async function handleWooCommerceSettingsGet(req, res, requestUrl) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }

  const requestedStoreUrl = normalizeWooCommerceStoreUrl(requestUrl.searchParams.get("storeUrl"));

  try {
    const connection = await getWooCommerceConnection(user.id, requestedStoreUrl, {
      includeSettings: true,
    });
    if (!connection) {
      sendJson(res, 404, { error: "WooCommerce is not connected for this account." });
      return;
    }
    const settings = getWooCommerceImportSettings(connection.import_settings);
    sendJson(res, 200, {
      storeUrl: connection.shop_domain,
      settings,
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Failed to load WooCommerce settings." });
  }
}

async function handleWooCommerceSettingsPost(req, res) {
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

  const requestedStoreUrl = normalizeWooCommerceStoreUrl(body?.storeUrl);
  const selectedStatuses = normalizeWooCommerceImportStatuses(body?.selectedStatuses);
  const autoRefreshEnabled = normalizeWooCommerceAutoRefreshEnabled(body?.autoRefreshEnabled);

  if (!selectedStatuses.length) {
    sendJson(res, 400, { error: "Select at least one WooCommerce status to import." });
    return;
  }

  try {
    const connection = await getWooCommerceConnection(user.id, requestedStoreUrl, {
      includeSettings: true,
    });
    if (!connection) {
      sendJson(res, 404, { error: "WooCommerce is not connected for this account." });
      return;
    }
    await saveWooCommerceImportSettings(
      user.id,
      connection.shop_domain,
      selectedStatuses,
      autoRefreshEnabled
    );
    sendJson(res, 200, {
      storeUrl: connection.shop_domain,
      settings: {
        selectedStatuses,
        autoRefreshEnabled,
      },
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Failed to save WooCommerce settings." });
  }
}

async function handleWooCommerceDisconnect(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (_error) {
    body = {};
  }

  const requestedStoreUrl = normalizeWooCommerceStoreUrl(body?.storeUrl);
  try {
    const connection = await getWooCommerceConnection(user.id, requestedStoreUrl);
    if (!connection) {
      sendJson(res, 200, { disconnected: false, connection: null });
      return;
    }
    await deleteWooCommerceConnection(user.id, connection.shop_domain);
    sendJson(res, 200, {
      disconnected: true,
      connection: {
        storeUrl: connection.shop_domain,
      },
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Failed to disconnect WooCommerce." });
  }
}

function getLinkedInCallbackRedirect(req, params = {}) {
  const url = new URL("/post", buildPublicBaseUrl(req));
  Object.entries(params).forEach(([key, value]) => {
    if (value === undefined || value === null || value === "") return;
    url.searchParams.set(key, String(value));
  });
  return url.toString();
}

function buildLinkedInConnectionPayload(row, stored, configured = true) {
  const expiresAtMs = Date.parse(String(stored?.expiresAt || ""));
  const needsReconnect = Number.isFinite(expiresAtMs) ? Date.now() >= expiresAtMs - 60 * 1000 : false;
  return {
    configured,
    connected: Boolean(row && stored?.accessToken),
    needsReconnect,
    connection:
      row && stored
        ? {
            connectedAt: row.connected_at || "",
            updatedAt: row.updated_at || "",
            expiresAt: stored.expiresAt || "",
            scopes: Array.isArray(stored.scopes) ? stored.scopes : [],
            selectedOrganizationUrn: stored.selectedOrganizationUrn || "",
            organizations: Array.isArray(stored.organizations) ? stored.organizations : [],
          }
        : null,
  };
}

function getSelectedLinkedInOrganization(stored) {
  if (!stored || !Array.isArray(stored.organizations)) return null;
  return (
    stored.organizations.find((entry) => entry.urn === stored.selectedOrganizationUrn)
    || stored.organizations[0]
    || null
  );
}

async function waitForLinkedInImageAvailable(accessToken, imageUrn) {
  for (let attempt = 0; attempt < 12; attempt += 1) {
    const payload = await fetchLinkedInJson(
      accessToken,
      `/rest/images/${encodeURIComponent(imageUrn)}`,
      { method: "GET", headers: { "Content-Type": "application/json" } }
    );
    const status = String(payload?.status || "").trim().toUpperCase();
    if (status === "AVAILABLE") {
      return payload;
    }
    if (status === "PROCESSING_FAILED") {
      throw new Error("LinkedIn image processing failed.");
    }
    await new Promise((resolve) => setTimeout(resolve, 1500));
  }
  throw new Error("LinkedIn image upload is still processing. Please try again.");
}

async function uploadLinkedInImage(accessToken, ownerUrn, imageDataUrl) {
  const { mimeType, bytes } = dataUrlToBytes(imageDataUrl);
  if (!["image/png", "image/jpeg", "image/jpg", "image/gif"].includes(mimeType)) {
    throw new Error("LinkedIn only accepts PNG, JPG, or GIF images.");
  }
  const initializePayload = await fetchLinkedInJson(
    accessToken,
    "/rest/images?action=initializeUpload",
    {
      method: "POST",
      body: JSON.stringify({
        initializeUploadRequest: {
          owner: ownerUrn,
        },
      }),
    }
  );
  const uploadUrl = String(initializePayload?.value?.uploadUrl || "").trim();
  const imageUrn = String(initializePayload?.value?.image || "").trim();
  if (!uploadUrl || !imageUrn) {
    throw new Error("LinkedIn image upload initialization failed.");
  }
  const uploadResponse = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Type": mimeType,
    },
    body: bytes,
  });
  if (!uploadResponse.ok) {
    const details = await uploadResponse.text().catch(() => "");
    throw new Error(`LinkedIn image upload failed (${uploadResponse.status}) ${details}`.trim());
  }
  await waitForLinkedInImageAvailable(accessToken, imageUrn);
  return imageUrn;
}

async function handleLinkedInInstallLink(req, res) {
  try {
    assertLinkedInConfig();
    const user = await getAuthenticatedUser(req);
    if (!user?.id) {
      sendJson(res, 401, { error: "Authentication required." });
      return;
    }
    const state = makeSignedStateToken({
      userId: user.id,
      provider: LINKEDIN_PROVIDER,
    });
    const installUrl = new URL("https://www.linkedin.com/oauth/v2/authorization");
    installUrl.searchParams.set("response_type", "code");
    installUrl.searchParams.set("client_id", LINKEDIN_CLIENT_ID);
    installUrl.searchParams.set("redirect_uri", getLinkedInRedirectUri(req));
    installUrl.searchParams.set("state", state);
    installUrl.searchParams.set("scope", getLinkedInScopes().join(" "));
    sendJson(res, 200, { url: installUrl.toString() });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not start LinkedIn connect flow." });
  }
}

async function handleLinkedInCallback(req, res, requestUrl) {
  try {
    assertLinkedInConfig();
    const linkedInError = String(requestUrl.searchParams.get("error") || "").trim();
    if (linkedInError) {
      sendRedirect(
        res,
        getLinkedInCallbackRedirect(req, {
          provider: LINKEDIN_PROVIDER,
          linkedin: "error",
          message:
            String(requestUrl.searchParams.get("error_description") || "").trim()
            || "LinkedIn authorization was not approved.",
        })
      );
      return;
    }
    const code = String(requestUrl.searchParams.get("code") || "").trim();
    const stateToken = String(requestUrl.searchParams.get("state") || "").trim();
    if (!code || !stateToken) {
      sendRedirect(
        res,
        getLinkedInCallbackRedirect(req, {
          provider: LINKEDIN_PROVIDER,
          linkedin: "error",
          message: "Missing LinkedIn callback parameters.",
        })
      );
      return;
    }
    const statePayload = parseSignedStateToken(stateToken);
    if (!statePayload?.userId) {
      sendRedirect(
        res,
        getLinkedInCallbackRedirect(req, {
          provider: LINKEDIN_PROVIDER,
          linkedin: "error",
          message: "LinkedIn OAuth session expired. Start the connect flow again.",
        })
      );
      return;
    }
    const tokenBody = new URLSearchParams();
    tokenBody.set("grant_type", "authorization_code");
    tokenBody.set("code", code);
    tokenBody.set("client_id", LINKEDIN_CLIENT_ID);
    tokenBody.set("client_secret", LINKEDIN_CLIENT_SECRET);
    tokenBody.set("redirect_uri", getLinkedInRedirectUri(req));
    const tokenResponse = await fetch("https://www.linkedin.com/oauth/v2/accessToken", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: tokenBody.toString(),
    });
    const tokenPayload = await tokenResponse.json().catch(() => null);
    if (!tokenResponse.ok || !tokenPayload?.access_token) {
      const message =
        String(tokenPayload?.error_description || tokenPayload?.error || "").trim()
        || `LinkedIn token exchange failed (${tokenResponse.status}).`;
      throw new Error(message);
    }
    let organizations = [];
    let callbackMessage = "LinkedIn connected.";
    try {
      organizations = await listLinkedInOrganizations(tokenPayload.access_token);
      if (!organizations.length) {
        callbackMessage =
          "LinkedIn connected, but no eligible company pages were found for this account.";
      }
    } catch (error) {
      callbackMessage =
        error?.message
        || "LinkedIn connected, but company pages could not be loaded yet.";
    }
    const expiresAt = Number(tokenPayload?.expires_in)
      ? new Date(Date.now() + Number(tokenPayload.expires_in) * 1000).toISOString()
      : "";
    const refreshExpiresAt = Number(tokenPayload?.refresh_token_expires_in)
      ? new Date(Date.now() + Number(tokenPayload.refresh_token_expires_in) * 1000).toISOString()
      : "";
    await upsertLinkedInConnection(statePayload.userId, {
      accessToken: tokenPayload.access_token,
      refreshToken: tokenPayload.refresh_token || "",
      expiresAt,
      refreshExpiresAt,
      scope: tokenPayload.scope || getLinkedInScopes().join(" "),
      organizations,
      selectedOrganizationUrn: organizations[0]?.urn || "",
    });
    sendRedirect(
      res,
      getLinkedInCallbackRedirect(req, {
        provider: LINKEDIN_PROVIDER,
        linkedin: "connected",
        message: callbackMessage,
      })
    );
  } catch (error) {
    sendRedirect(
      res,
      getLinkedInCallbackRedirect(req, {
        provider: LINKEDIN_PROVIDER,
        linkedin: "error",
        message: error?.message || "LinkedIn connection failed.",
      })
    );
  }
}

async function handleLinkedInConnection(req, res, requestUrl) {
  try {
    const user = await getAuthenticatedUser(req);
    if (!user?.id) {
      sendJson(res, 401, { error: "Authentication required." });
      return;
    }
    const configured = Boolean(LINKEDIN_CLIENT_ID && LINKEDIN_CLIENT_SECRET);
    if (!configured) {
      sendJson(res, 200, buildLinkedInConnectionPayload(null, null, false));
      return;
    }
    const row = await getLinkedInConnection(user.id);
    if (!row) {
      sendJson(res, 200, buildLinkedInConnectionPayload(null, null, true));
      return;
    }
    let stored = parseStoredLinkedInConnection(row.access_token);
    if (!stored?.accessToken) {
      sendJson(res, 200, buildLinkedInConnectionPayload(null, null, true));
      return;
    }
    const shouldRefresh = String(requestUrl.searchParams.get("refresh") || "").trim() === "1";
    const expiresAtMs = Date.parse(String(stored.expiresAt || ""));
    const expired = Number.isFinite(expiresAtMs) ? Date.now() >= expiresAtMs - 60 * 1000 : false;
    if (shouldRefresh && !expired) {
      stored = await refreshLinkedInConnectionOrganizations(row, stored);
    }
    sendJson(res, 200, buildLinkedInConnectionPayload(row, stored, true));
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not load LinkedIn connection." });
  }
}

async function handleLinkedInSelectOrganization(req, res) {
  try {
    const user = await getAuthenticatedUser(req);
    if (!user?.id) {
      sendJson(res, 401, { error: "Authentication required." });
      return;
    }
    const row = await getLinkedInConnection(user.id);
    if (!row) {
      sendJson(res, 404, { error: "LinkedIn is not connected for this account." });
      return;
    }
    const stored = parseStoredLinkedInConnection(row.access_token);
    if (!stored?.accessToken) {
      sendJson(res, 400, { error: "LinkedIn connection data is invalid. Please reconnect." });
      return;
    }
    const body = await readJsonBody(req);
    const organizationUrn = String(body?.organizationUrn || "").trim();
    if (!organizationUrn) {
      sendJson(res, 400, { error: "Select a LinkedIn company page." });
      return;
    }
    const match = stored.organizations.find((entry) => entry.urn === organizationUrn);
    if (!match) {
      sendJson(res, 400, { error: "The selected LinkedIn company page is not available." });
      return;
    }
    const nextPayload = normalizeLinkedInStoredPayload({
      ...stored,
      selectedOrganizationUrn: organizationUrn,
    });
    await upsertLinkedInConnection(user.id, {
      ...nextPayload,
      connectedAt: row.connected_at || "",
    });
    sendJson(res, 200, buildLinkedInConnectionPayload(row, nextPayload, true));
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not save the LinkedIn page selection." });
  }
}

async function handleLinkedInDisconnect(req, res) {
  try {
    const user = await getAuthenticatedUser(req);
    if (!user?.id) {
      sendJson(res, 401, { error: "Authentication required." });
      return;
    }
    await deleteSupabaseRows(
      "provider_connections",
      {
        user_id: `eq.${user.id}`,
        provider: `eq.${LINKEDIN_PROVIDER}`,
        shop_domain: `eq.${getLinkedInConnectionRowKey()}`,
      },
      { allowAll: false }
    );
    sendJson(res, 200, { disconnected: true });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not disconnect LinkedIn." });
  }
}

async function handleLinkedInPost(req, res) {
  try {
    assertLinkedInConfig();
    const user = await getAuthenticatedUser(req);
    if (!user?.id) {
      sendJson(res, 401, { error: "Authentication required." });
      return;
    }
    const row = await getLinkedInConnection(user.id);
    if (!row) {
      sendJson(res, 404, { error: "Connect LinkedIn before publishing." });
      return;
    }
    const stored = parseStoredLinkedInConnection(row.access_token);
    if (!stored?.accessToken) {
      sendJson(res, 400, { error: "LinkedIn connection data is invalid. Please reconnect." });
      return;
    }
    const expiresAtMs = Date.parse(String(stored.expiresAt || ""));
    if (Number.isFinite(expiresAtMs) && Date.now() >= expiresAtMs - 60 * 1000) {
      sendJson(res, 400, { error: "Your LinkedIn connection has expired. Please reconnect." });
      return;
    }
    const selectedOrganization = getSelectedLinkedInOrganization(stored);
    if (!selectedOrganization?.urn) {
      sendJson(res, 400, { error: "Select a LinkedIn company page before publishing." });
      return;
    }
    const body = await readJsonBody(req);
    const caption = String(body?.caption || "").trim();
    const imageDataUrl = String(body?.imageDataUrl || "").trim();
    const altText = String(body?.altText || "").trim().slice(0, 120);
    if (!caption) {
      sendJson(res, 400, { error: "Add a LinkedIn caption before publishing." });
      return;
    }
    if (!imageDataUrl) {
      sendJson(res, 400, { error: "The LinkedIn image payload is missing." });
      return;
    }
    const imageUrn = await uploadLinkedInImage(
      stored.accessToken,
      selectedOrganization.urn,
      imageDataUrl
    );
    const postResponse = await fetch("https://api.linkedin.com/rest/posts", {
      method: "POST",
      headers: getLinkedInApiHeaders(stored.accessToken, {
        "Content-Type": "application/json",
      }),
      body: JSON.stringify({
        author: selectedOrganization.urn,
        commentary: caption,
        visibility: "PUBLIC",
        distribution: {
          feedDistribution: "MAIN_FEED",
          targetEntities: [],
          thirdPartyDistributionChannels: [],
        },
        content: {
          media: {
            id: imageUrn,
            ...(altText ? { altText } : {}),
          },
        },
        lifecycleState: "PUBLISHED",
        isReshareDisabledByAuthor: false,
      }),
    });
    const postPayload = await postResponse.json().catch(() => null);
    if (!postResponse.ok) {
      const message =
        String(postPayload?.message || postPayload?.error || "").trim()
        || `LinkedIn post publish failed (${postResponse.status}).`;
      throw new Error(message);
    }
    const postUrn = String(postResponse.headers.get("x-restli-id") || postPayload?.id || "").trim();
    const feedUrl = postUrn
      ? `https://www.linkedin.com/feed/update/${encodeURIComponent(postUrn)}/`
      : "";
    sendJson(res, 200, {
      ok: true,
      postUrn,
      feedUrl,
      organization: {
        urn: selectedOrganization.urn,
        name: selectedOrganization.name,
        vanityName: selectedOrganization.vanityName,
      },
    });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Could not publish to LinkedIn." });
  }
}

async function handleWooCommerceImportOrders(req, res) {
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

  const requestedStoreUrl = normalizeWooCommerceStoreUrl(body?.storeUrl);
  const requestedStatuses = normalizeWooCommerceImportStatuses(body?.selectedStatuses);
  const limit = Number(body?.limit);
  let resolvedStoreUrl = requestedStoreUrl;

  try {
    let connection = null;
    let settings = {
      selectedStatuses: [...DEFAULT_WOOCOMMERCE_IMPORT_STATUSES],
      autoRefreshEnabled: false,
    };
    try {
      connection = await getWooCommerceConnection(user.id, requestedStoreUrl, {
        includeSettings: true,
      });
      if (connection) {
        settings = getWooCommerceImportSettings(connection.import_settings);
      }
    } catch (error) {
      if (error?.code !== "missing_import_settings_column") {
        throw error;
      }
      connection = await getWooCommerceConnection(user.id, requestedStoreUrl);
    }
    if (!connection) {
      sendJson(res, 404, { error: "WooCommerce is not connected for this account." });
      return;
    }
    resolvedStoreUrl = String(connection.shop_domain || requestedStoreUrl || "").trim();
    const credentials = parseWooCommerceCredentials(connection.access_token);
    const selectedStatuses = requestedStatuses.length
      ? requestedStatuses
      : settings.selectedStatuses;
    const [orders, storeSettings] = await Promise.all([
      fetchWooCommerceOrders(
        connection.shop_domain,
        credentials.consumerKey,
        credentials.consumerSecret,
        limit,
        selectedStatuses
      ),
      fetchWooCommerceGeneralSettings(
        connection.shop_domain,
        credentials.consumerKey,
        credentials.consumerSecret
      ).catch(() => null),
    ]);
    const senderOrigin = buildWooCommerceSenderOrigin(storeSettings, connection.shop_domain);
    const weightUnit = normalizeWooCommerceWeightUnit(
      getWooCommerceSettingValue(storeSettings, "woocommerce_weight_unit")
    );
    const lineItemWeights = await fetchWooCommerceLineItemWeights(
      connection.shop_domain,
      credentials.consumerKey,
      credentials.consumerSecret,
      orders,
      weightUnit
    );
    const rows = mapWooCommerceOrdersToCsvRows(
      orders,
      senderOrigin,
      selectedStatuses,
      lineItemWeights
    );
    sendJson(res, 200, {
      storeUrl: connection.shop_domain,
      count: rows.length,
      rows,
      settings: {
        selectedStatuses,
        autoRefreshEnabled: settings.autoRefreshEnabled,
      },
    });
  } catch (error) {
    if (error?.code === "token_decrypt_failed") {
      sendJson(res, 409, {
        error:
          "Stored WooCommerce connection could not be decrypted. Reconnect WooCommerce and try again.",
      });
      return;
    }
    if (
      (
        error?.code === "woocommerce_auth_invalid"
        || error?.code === "woocommerce_api_not_found"
        || error?.code === "woocommerce_store_url_disallowed"
      ) &&
      resolvedStoreUrl
    ) {
      await setWooCommerceConnectionStatus(user.id, resolvedStoreUrl, "token_invalid").catch(() => {});
      sendJson(res, 409, {
        error:
          error?.code === "woocommerce_api_not_found"
            ? "WooCommerce API endpoint was not found for this store. Open WooCommerce settings and reconnect with the correct URL."
            : error?.code === "woocommerce_store_url_disallowed"
              ? "WooCommerce store URL is no longer allowed. Reconnect WooCommerce using a public HTTPS store URL."
              : "WooCommerce credentials expired or were rejected. Reconnect WooCommerce and try again.",
      });
      return;
    }
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "WooCommerce import failed." });
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
    if (error?.code === "token_decrypt_failed") {
      sendJson(res, 409, {
        error: "Stored Shopify connection could not be decrypted. Reconnect Shopify and try again.",
      });
      return;
    }
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
    const settings = getShopifyImportSettings(connection.import_settings);
    sendJson(res, 200, {
      shop: connection.shop_domain,
      settings,
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
  const selectedFinancialStatuses = normalizeShopifyFinancialStatuses(
    Array.isArray(body?.selectedFinancialStatuses)
      ? body.selectedFinancialStatuses
      : body?.financialStatus
  );
  const autoRefreshEnabled = normalizeWooCommerceAutoRefreshEnabled(body?.autoRefreshEnabled);
  if (!selectedLocationIds.length) {
    sendJson(res, 400, { error: "Select at least one fulfillment location." });
    return;
  }
  if (!selectedFinancialStatuses.length) {
    sendJson(res, 400, { error: "Select at least one Shopify status to import." });
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
    await saveShopifyImportSettings(
      user.id,
      connection.shop_domain,
      selectedLocationIds,
      selectedFinancialStatuses,
      autoRefreshEnabled
    );
    sendJson(res, 200, {
      shop: connection.shop_domain,
      settings: {
        selectedLocationIds,
        selectedFinancialStatuses,
        autoRefreshEnabled,
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

async function handleShopifyDisconnect(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }

  let body = {};
  try {
    body = await readJsonBody(req);
  } catch (_error) {
    body = {};
  }

  const requestedShop = normalizeShopDomain(body?.shop);
  try {
    const connection = await getShopifyConnection(user.id, requestedShop);
    if (!connection) {
      sendJson(res, 200, { disconnected: false, connection: null });
      return;
    }
    await deleteShopifyConnection(user.id, connection.shop_domain);
    sendJson(res, 200, {
      disconnected: true,
      connection: {
        shop: connection.shop_domain,
      },
    });
  } catch (error) {
    if (error?.code === "missing_import_settings_column") {
      sendJson(res, 500, { error: error.message });
      return;
    }
    sendJson(res, 500, { error: error.message || "Failed to disconnect Shopify." });
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
  const requestedFinancialStatuses = normalizeShopifyFinancialStatuses(
    Array.isArray(body?.selectedFinancialStatuses)
      ? body.selectedFinancialStatuses
      : body?.financialStatus
  );
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
    const savedSettings = getShopifyImportSettings(connection.import_settings);
    const savedSelectedLocationIds = savedSettings.selectedLocationIds;
    let resolvedSelectedLocationIds = selectedLocationIds.length
      ? selectedLocationIds
      : savedSelectedLocationIds;
    if (!resolvedSelectedLocationIds.length) {
      resolvedSelectedLocationIds = locations.map((location) => location.id);
    }
    const locationById = indexLocationsById(locations);
    const resolvedFinancialStatuses = requestedFinancialStatuses.length
      ? requestedFinancialStatuses
      : savedSettings.selectedFinancialStatuses;
    const orders = await fetchShopifyOrders(
      connection.shop_domain,
      accessToken,
      limit,
      resolvedFinancialStatuses
    );
    const rows = mapShopifyOrdersToCsvRows(orders, {
      locationById,
      selectedLocationIds: resolvedSelectedLocationIds,
      shopDomain: connection.shop_domain,
    });
    sendJson(res, 200, {
      shop: connection.shop_domain,
      count: rows.length,
      settings: {
        selectedLocationIds: savedSettings.selectedLocationIds,
        selectedFinancialStatuses: savedSettings.selectedFinancialStatuses,
        autoRefreshEnabled: savedSettings.autoRefreshEnabled,
      },
      rows,
    });
  } catch (error) {
    if (error?.code === "token_decrypt_failed") {
      sendJson(res, 409, {
        error: "Stored Shopify connection could not be decrypted. Reconnect Shopify and try again.",
      });
      return;
    }
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

async function handleShopifyDevSeedOrders(req, res) {
  const user = await getAuthenticatedUser(req);
  if (!user?.id) {
    sendJson(res, 401, { error: "Authentication required." });
    return;
  }
  if (!isShopifyDevSeedOrdersEnabled()) {
    sendJson(res, 404, { error: "API route not found." });
    return;
  }
  if (!canManageRegistrationInvites(user)) {
    sendJson(res, 403, { error: "Admin access required." });
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
  const countRaw = Number(body?.count);
  const count = Number.isFinite(countRaw) ? Math.max(1, Math.min(25, Math.trunc(countRaw))) : 10;
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
    const scopes = String(connection.scopes || "")
      .split(",")
      .map((scope) => scope.trim())
      .filter(Boolean);
    if (!scopes.includes("write_orders")) {
      sendJson(res, 409, {
        error: "Shopify connection is missing write_orders. Reconnect Shopify once, then try seeding again.",
      });
      return;
    }

    const accessToken = decryptToken(connection.access_token);
    const createdOrders = await createBogusShopifyOrders(connection.shop_domain, accessToken, count);
    sendJson(res, 200, {
      shop: connection.shop_domain,
      count: createdOrders.length,
      orders: createdOrders,
    });
  } catch (error) {
    if (error?.code === "token_decrypt_failed") {
      sendJson(res, 409, {
        error: "Stored Shopify connection could not be decrypted. Reconnect Shopify and try again.",
      });
      return;
    }
    if (error?.status === 401 && resolvedShop) {
      await setShopifyConnectionStatus(user.id, resolvedShop, "token_invalid").catch(() => {});
      sendJson(res, 409, {
        error: "Shopify connection expired or was revoked. Reconnect Shopify and try again.",
      });
      return;
    }
    if (error?.code === "missing_write_orders") {
      sendJson(res, 409, {
        error: "Shopify connection is missing write_orders. Reconnect Shopify once, then try seeding again.",
      });
      return;
    }
    sendJson(res, 500, { error: error.message || "Shopify bogus order creation failed." });
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
  if (pathname === "/api/admin/client-billing" && req.method === "POST") {
    await handleAdminClientBillingSave(req, res);
    return true;
  }
  if (pathname === "/api/admin/invoices" && req.method === "GET") {
    await handleAdminInvoicesList(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/admin/invoices/detail" && req.method === "GET") {
    await handleAdminInvoiceDetail(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/admin/invoices/pdf" && req.method === "GET") {
    await proxyAuthenticatedApi(req, res, `/api/admin/invoices/pdf${requestUrl.search}`);
    return true;
  }
  if (pathname === "/api/admin/invoices/run" && req.method === "POST") {
    await handleAdminInvoicesRun(req, res);
    return true;
  }
  if (pathname === "/api/admin/privacy-maintenance/run" && req.method === "POST") {
    await handleAdminPrivacyMaintenanceRun(req, res);
    return true;
  }
  if (pathname === "/api/admin/invoices/send" && req.method === "POST") {
    await proxyAuthenticatedApi(req, res, "/api/admin/invoices/send");
    return true;
  }
  if (pathname === "/api/admin/invoices/accounting-export" && req.method === "POST") {
    await proxyAuthenticatedApi(req, res, "/api/admin/invoices/accounting-export");
    return true;
  }
  if (pathname === "/api/admin/invoices/send-test" && req.method === "POST") {
    await handleAdminInvoiceSendTest(req, res);
    return true;
  }
  if (pathname === "/api/admin/invoices/send-topup-test" && req.method === "POST") {
    await handleAdminTopupInvoiceSendTest(req, res);
    return true;
  }
  if (pathname === "/api/admin/invoices/send-test-sequence" && req.method === "POST") {
    await handleAdminInvoiceSendTestSequence(req, res);
    return true;
  }
  if (pathname === "/api/admin/reports/send-test" && req.method === "POST") {
    await handleAdminReportsSendTest(req, res);
    return true;
  }
  if (pathname === "/api/admin/agreements/preview-test" && req.method === "POST") {
    await handleAdminAgreementPreviewTest(req, res);
    return true;
  }
  if (pathname === "/api/admin/agreements/send-test" && req.method === "POST") {
    await handleAdminAgreementSendTest(req, res);
    return true;
  }
  if (pathname === "/api/admin/invoices/mark-paid" && req.method === "POST") {
    await handleAdminInvoiceMarkPaid(req, res);
    return true;
  }
  if (pathname === "/api/admin/billing/wise/receipts" && req.method === "GET") {
    await handleAdminWiseReceiptsList(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/admin/billing/wise/sync" && req.method === "POST") {
    await handleAdminWiseSync(req, res);
    return true;
  }
  if (pathname === "/api/admin/billing/wise/receipts/resolve" && req.method === "POST") {
    await handleAdminWiseReceiptResolve(req, res);
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
  if (pathname === "/api/admin/shipment-extract-requests" && req.method === "GET") {
    await handleListShipmentExtractRequests(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/admin/shipment-extract-requests" && req.method === "POST") {
    await handleCreateShipmentExtractRequest(req, res);
    return true;
  }
  if (
    pathname === "/api/admin/shipment-extract-requests/create-registration-invite" &&
    req.method === "POST"
  ) {
    await handleCreateRegistrationInviteFromShipmentExtract(req, res);
    return true;
  }
  if (pathname === "/api/auth/register-invite" && req.method === "GET") {
    await handleRegisterInviteValidate(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/auth/register-contract-preview" && req.method === "POST") {
    await handleRegisterContractPreview(req, res);
    return true;
  }
  if (pathname === "/api/auth/register-contract-preview-file" && req.method === "GET") {
    await handleRegisterContractPreviewFile(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/auth/register" && req.method === "POST") {
    await handleRegisterWithInvite(req, res);
    return true;
  }
  if (
    (pathname === "/api/public/clean-data/submit" ||
      pathname === "/api/public/shipping-data-cleaner/submit") &&
    req.method === "POST"
  ) {
    await handleShippingDataCleanerSubmission(req, res);
    return true;
  }
  if (
    (pathname === "/api/public/clean-data/request" ||
      pathname === "/api/public/shipping-data-cleaner/request") &&
    req.method === "GET"
  ) {
    await handleShippingDataCleanerRequestValidate(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/privacy/export" && req.method === "GET") {
    await handlePrivacyExport(req, res);
    return true;
  }
  if (pathname === "/api/privacy/delete" && req.method === "POST") {
    await handlePrivacyDelete(req, res);
    return true;
  }
  if (pathname === "/api/wise/webhooks" && req.method === "POST") {
    await handleWiseWebhook(req, res);
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
  if (pathname === "/api/shopify/public-config" && req.method === "GET") {
    await handleShopifyPublicConfig(req, res);
    return true;
  }
  if (pathname === "/api/shopify/embedded/session" && req.method === "GET") {
    await handleShopifyEmbeddedSession(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/shopify/webhooks/compliance" && req.method === "POST") {
    await handleShopifyComplianceWebhook(req, res);
    return true;
  }
  if (pathname === "/api/shopify/connection" && req.method === "GET") {
    await handleShopifyConnection(req, res);
    return true;
  }
  if (pathname === "/api/wix/install-link" && req.method === "POST") {
    await handleWixInstallLink(req, res);
    return true;
  }
  if (pathname === "/api/wix/connection" && req.method === "GET") {
    await handleWixConnection(req, res);
    return true;
  }
  if (pathname === "/api/wix/link-instance" && req.method === "POST") {
    await handleWixLinkInstance(req, res);
    return true;
  }
  if (pathname === "/api/wix/import-orders" && req.method === "POST") {
    await handleWixImportOrders(req, res);
    return true;
  }
  if (
    (pathname === "/api/wix/settings" || pathname === "/api/wix/setting") &&
    req.method === "GET"
  ) {
    await handleWixSettingsGet(req, res, requestUrl);
    return true;
  }
  if (
    (pathname === "/api/wix/settings" || pathname === "/api/wix/setting") &&
    req.method === "POST"
  ) {
    await handleWixSettingsPost(req, res);
    return true;
  }
  if (pathname === "/api/wix/disconnect" && req.method === "POST") {
    await handleWixDisconnect(req, res);
    return true;
  }
  if (pathname === "/api/woocommerce/install-link" && req.method === "POST") {
    await handleWooCommerceInstallLink(req, res);
    return true;
  }
  if (pathname === "/api/woocommerce/callback" && req.method === "POST") {
    await handleWooCommerceCallback(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/woocommerce/connection" && req.method === "GET") {
    await handleWooCommerceConnection(req, res);
    return true;
  }
  if (
    (pathname === "/api/woocommerce/settings" || pathname === "/api/woocommerce/setting") &&
    req.method === "GET"
  ) {
    await handleWooCommerceSettingsGet(req, res, requestUrl);
    return true;
  }
  if (
    (pathname === "/api/woocommerce/settings" || pathname === "/api/woocommerce/setting") &&
    req.method === "POST"
  ) {
    await handleWooCommerceSettingsPost(req, res);
    return true;
  }
  if (pathname === "/api/woocommerce/import-orders" && req.method === "POST") {
    await handleWooCommerceImportOrders(req, res);
    return true;
  }
  if (pathname === "/api/woocommerce/disconnect" && req.method === "POST") {
    await handleWooCommerceDisconnect(req, res);
    return true;
  }
  if (pathname === "/api/linkedin/install-link" && req.method === "POST") {
    await handleLinkedInInstallLink(req, res);
    return true;
  }
  if (pathname === "/api/linkedin/callback" && req.method === "GET") {
    await handleLinkedInCallback(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/linkedin/connection" && req.method === "GET") {
    await handleLinkedInConnection(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/linkedin/select-organization" && req.method === "POST") {
    await handleLinkedInSelectOrganization(req, res);
    return true;
  }
  if (pathname === "/api/linkedin/post" && req.method === "POST") {
    await handleLinkedInPost(req, res);
    return true;
  }
  if (pathname === "/api/linkedin/disconnect" && req.method === "POST") {
    await handleLinkedInDisconnect(req, res);
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
  if (pathname === "/api/shopify/disconnect" && req.method === "POST") {
    await handleShopifyDisconnect(req, res);
    return true;
  }
  if (pathname === "/api/shopify/dev-seed-orders" && req.method === "POST") {
    await handleShopifyDevSeedOrders(req, res);
    return true;
  }
  if (pathname === "/api/billing/overview" && req.method === "GET") {
    await handleBillingOverview(req, res);
    return true;
  }
  if (pathname === "/api/billing/topups" && req.method === "GET") {
    await handleBillingTopups(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/billing/topups/invoice-pdf" && req.method === "GET") {
    await proxyAuthenticatedApi(req, res, `/api/billing/topups/invoice-pdf${requestUrl.search}`);
    return true;
  }
  if (pathname === "/api/billing/topups/repair-invoice" && req.method === "POST") {
    await proxyAuthenticatedApi(req, res, "/api/billing/topups/repair-invoice");
    return true;
  }
  if (pathname === "/api/billing/invoices/detail" && req.method === "GET") {
    await handleBillingInvoiceDetail(req, res, requestUrl);
    return true;
  }
  if (pathname === "/api/billing/invoices/pdf" && req.method === "GET") {
    await proxyAuthenticatedApi(req, res, `/api/billing/invoices/pdf${requestUrl.search}`);
    return true;
  }
  if (pathname === "/api/billing/render-invoice-pdf" && req.method === "POST") {
    await proxyAuthenticatedPdfRender(req, res, "/api/billing/render-invoice-pdf");
    return true;
  }
  if (pathname === "/api/billing/render-receipt-pdf" && req.method === "POST") {
    await proxyAuthenticatedPdfRender(req, res, "/api/billing/render-receipt-pdf");
    return true;
  }
  if (pathname === "/api/billing/receipts/ensure-code" && req.method === "POST") {
    await proxyAuthenticatedApi(req, res, "/api/billing/receipts/ensure-code");
    return true;
  }
  if (pathname === "/api/document-jobs/status" && req.method === "GET") {
    await proxyAuthenticatedApi(req, res, `/api/document-jobs/status${requestUrl.search}`);
    return true;
  }
  if (pathname === "/api/billing/topups/request" && req.method === "POST") {
    await handleBillingTopupRequest(req, res);
    return true;
  }
  if (pathname === "/api/billing/checkout" && req.method === "POST") {
    await handleBillingCheckout(req, res);
    return true;
  }
  if (pathname === "/api/label-generations" && req.method === "POST") {
    await handleLabelGenerationCreate(req, res);
    return true;
  }
  if (pathname === "/api/documents-preview/send-test-email" && req.method === "POST") {
    await handleDocumentsPreviewSendTestEmail(req, res);
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

  if (pathname === "/__preview-version") {
    sendJson(res, 200, buildPreviewVersionPayload());
    return;
  }

  if (pathname === "/invoice-view" || pathname === "/invoice-view/") {
    await handleInvoicePdfView(req, res, requestUrl);
    return;
  }

  if (pathname === "/invoice-preview" || pathname === "/invoice-preview/") {
    sendFile(res, INVOICE_PREVIEW_FILE);
    return;
  }

  if (pathname === "/documents-preview" || pathname === "/documents-preview/") {
    sendFile(res, DOCUMENTS_PREVIEW_FILE);
    return;
  }

  if (pathname === "/iban-topup-preview" || pathname === "/iban-topup-preview/") {
    sendFile(res, IBAN_TOPUP_PREVIEW_FILE);
    return;
  }

  if (pathname === "/landing-platform-mock" || pathname === "/landing-platform-mock/") {
    sendFile(res, LANDING_PLATFORM_MOCK_FILE);
    return;
  }

  if (
    pathname === "/clean-data" ||
    pathname === "/clean-data/" ||
    pathname === "/shipping-data-cleaner" ||
    pathname === "/shipping-data-cleaner/"
  ) {
    sendFile(res, SHIPPING_DATA_CLEANER_FILE);
    return;
  }

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
