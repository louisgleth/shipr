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

async function getShopifyConnection(userId, requestedShop) {
  const params = new URLSearchParams();
  params.set("select", "shop_domain,status,scopes,connected_at,updated_at,access_token");
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
    throw new Error(`Failed reading Shopify connection (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) {
    return null;
  }
  return rows[0];
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
    const connection = await getShopifyConnection(user.id, requestedShop);
    if (!connection) {
      sendJson(res, 404, { error: "Shopify is not connected for this account." });
      return;
    }
    resolvedShop = String(connection.shop_domain || resolvedShop || "").trim();
    const accessToken = decryptToken(connection.access_token);
    let locationById = {};
    if (selectedLocationIds.length > 0) {
      const locations = await fetchShopifyLocations(connection.shop_domain, accessToken);
      locationById = indexLocationsById(locations);
    }
    const orders = await fetchShopifyOrders(connection.shop_domain, accessToken, limit);
    const rows = mapShopifyOrdersToCsvRows(orders, {
      locationById,
      selectedLocationIds,
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
    sendJson(res, 500, { error: error.message || "Shopify import failed." });
  }
}

async function handleApi(req, res, requestUrl) {
  const pathname = requestUrl.pathname;
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
  if (pathname === "/api/shopify/locations" && req.method === "GET") {
    await handleShopifyLocations(req, res, requestUrl);
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
