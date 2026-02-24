const MAX_BODY_BYTES = 1024 * 1024;
const OAUTH_STATE_TTL_MS = 10 * 60 * 1000;
const DEFAULT_SHOPIFY_SCOPES = "read_orders,read_locations";
const DEFAULT_SHOPIFY_API_VERSION = "2025-10";

const textEncoder = new TextEncoder();

export default {
  async fetch(request, env) {
    try {
      const url = new URL(request.url);

      if (request.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: noStoreHeaders(),
        });
      }

      if (url.pathname === "/api/shopify/install-link" && request.method === "POST") {
        return handleInstallLink(request, env);
      }
      if (url.pathname === "/api/shopify/callback" && request.method === "GET") {
        return handleShopifyCallback(request, env, url);
      }
      if (url.pathname === "/api/shopify/connection" && request.method === "GET") {
        return handleConnection(request, env);
      }
      if (url.pathname === "/api/shopify/locations" && request.method === "GET") {
        return handleLocations(request, env, url);
      }
      if (url.pathname === "/api/shopify/import-orders" && request.method === "POST") {
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

async function getShopifyConnection(env, userId, shop) {
  const params = new URLSearchParams();
  params.set("select", "shop_domain,status,scopes,connected_at,updated_at,access_token");
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
    throw new Error(`Failed reading Shopify connection (${response.status}) ${details}`.trim());
  }
  const rows = await response.json().catch(() => []);
  if (!Array.isArray(rows) || !rows.length) return null;
  return rows[0];
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

    const connection = await getShopifyConnection(env, user.id, shop);
    if (!connection) {
      return jsonResponse({ error: "Shopify is not connected for this account." }, 404);
    }

    const accessToken = await decryptToken(env, connection.access_token);
    let locationById = {};
    if (selectedLocationIds.length > 0) {
      const locations = await fetchShopifyLocations(env, connection.shop_domain, accessToken);
      locationById = indexLocationsById(locations);
    }
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
        selectedLocationIds,
      }
    );
    return jsonResponse({
      shop: connection.shop_domain,
      count: rows.length,
      rows,
    });
  } catch (error) {
    return jsonResponse({ error: error?.message || "Shopify import failed." }, 500);
  }
}
