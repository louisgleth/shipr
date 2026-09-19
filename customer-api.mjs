import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';

export const SCOPES = ['account:read', 'orders:read', 'orders:write', 'shipments:read', 'shipments:write', 'rates:read', 'labels:read', 'labels:write', 'webhooks:read', 'webhooks:write'];
export const EVENTS = ['order.created', 'shipment.created', 'label.created', 'label.voided'];
const encoder = new TextEncoder();
const MAX_BODY = 65536;
export class ApiError extends Error {
  constructor(status, code, message) { super(message); this.status = status; this.code = code; }
}
const fail = (status, code, message) => { throw new ApiError(status, code, message); };
const id = prefix => `${prefix}_${crypto.randomUUID().replaceAll('-', '')}`;
const now = () => new Date().toISOString();
const hex = bytes => Array.from(new Uint8Array(bytes), x => x.toString(16).padStart(2, '0')).join('');
export const hash = async value => hex(await crypto.subtle.digest('SHA-256', encoder.encode(value)));
const randomSecret = prefix => prefix + hex(crypto.getRandomValues(new Uint8Array(32)));
const json = (data, status = 200, extra = {}) => new Response(JSON.stringify(data), { status, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff', ...extra } });
const canonical = value => JSON.stringify(value && typeof value === 'object' ? Array.isArray(value) ? value.map(v => JSON.parse(canonical(v))) : Object.fromEntries(Object.keys(value).sort().map(k => [k, JSON.parse(canonical(value[k]))])) : value);
const stmt = (db, sql, ...values) => db.prepare(sql).bind(...values);
const required = (value, name, max = 200) => {
  if (typeof value !== 'string' || !value.trim() || value.length > max || /[\x00-\x1f]/.test(value)) fail(422, 'invalid_request', `${name} must be a non-empty string of at most ${max} characters.`);
  return value.trim();
};
const only = (body, fields) => {
  const unknown = Object.keys(body).filter(k => !fields.includes(k));
  if (unknown.length) fail(422, 'invalid_request', `Unknown fields: ${unknown.join(', ')}.`);
};
export async function readBody(request) {
  if (!(request.headers.get('content-type') || '').split(';')[0].trim().match(/^application\/json$/i)) fail(415, 'unsupported_media_type', 'Use Content-Type: application/json.');
  if (Number(request.headers.get('content-length')) > MAX_BODY) fail(413, 'body_too_large', 'Request exceeds 64 KiB.');
  const reader = request.body?.getReader();
  let size = 0; const chunks = [];
  if (reader) {
    for (;;) {
      const { done, value } = await reader.read(); if (done) break;
      size += value.byteLength;
      if (size > MAX_BODY) { await reader.cancel(); fail(413, 'body_too_large', 'Request exceeds 64 KiB.'); }
      chunks.push(value);
    }
  }
  const bytes = new Uint8Array(size); let offset = 0;
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
  let body;
  try { body = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes)); } catch { fail(400, 'invalid_json', 'Body must be valid JSON.'); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) fail(422, 'invalid_request', 'Body must be a JSON object.');
  return body;
}
function address(value, name) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) fail(422, 'invalid_address', `${name} is required.`);
  only(value, ['name', 'company', 'street1', 'street2', 'city', 'state', 'postal_code', 'country', 'email', 'phone']);
  const result = {};
  for (const key of ['name', 'street1', 'city', 'postal_code', 'country']) result[key] = required(value[key], `${name}.${key}`);
  result.country = result.country.toUpperCase();
  if (!/^[A-Z]{2}$/.test(result.country)) fail(422, 'invalid_address', `${name}.country must be an ISO two-letter country code.`);
  for (const key of ['company', 'street2', 'state', 'email', 'phone']) if (value[key] !== undefined) result[key] = required(value[key], `${name}.${key}`);
  return result;
}
export function shipmentInput(body) {
  only(body, ['reference', 'from', 'to', 'parcel', 'order_id', 'metadata']);
  const result = { reference: required(body.reference, 'reference'), from: address(body.from, 'from'), to: address(body.to, 'to') };
  const parcel = body.parcel;
  if (!parcel || typeof parcel !== 'object' || Array.isArray(parcel)) fail(422, 'invalid_parcel', 'parcel is required.');
  only(parcel, ['weight_kg', 'length_cm', 'width_cm', 'height_cm']);
  result.parcel = {};
  for (const key of ['weight_kg', 'length_cm', 'width_cm', 'height_cm']) {
    if (typeof parcel[key] !== 'number' || !Number.isFinite(parcel[key]) || parcel[key] <= 0 || parcel[key] > (key === 'weight_kg' ? 1000 : 1000)) fail(422, 'invalid_parcel', `parcel.${key} must be a number greater than 0 and at most 1000.`);
    result.parcel[key] = parcel[key];
  }
  if (body.order_id !== undefined) result.order_id = required(body.order_id, 'order_id', 64);
  if (body.metadata !== undefined) {
    if (!body.metadata || typeof body.metadata !== 'object' || Array.isArray(body.metadata) || JSON.stringify(body.metadata).length > 4096) fail(422, 'invalid_request', 'metadata must be an object of at most 4096 characters.');
    result.metadata = body.metadata;
  }
  return result;
}
function authorize(auth, scope) { if (!auth.scopes.includes(scope)) fail(403, 'insufficient_scope', `This key requires ${scope}.`); }
async function authenticate(request, db) {
  const token = (request.headers.get('authorization') || '').match(/^Bearer (shipide_(?:test|live)_[a-f0-9]{64})$/)?.[1];
  if (!token) fail(401, 'invalid_api_key', 'Provide a Shipide API key using Authorization: Bearer.');
  const key = await stmt(db, 'SELECT * FROM api_keys WHERE key_hash = ?', await hash(token)).first();
  if (!key || key.revoked_at || (key.expires_at && key.expires_at <= now())) fail(401, 'invalid_api_key', 'The API key is invalid, expired, or revoked.');
  const window = Math.floor(Date.now() / 60000);
  const rate = await stmt(db, 'INSERT INTO api_rate_limits (key_id,window,hits) VALUES (?,?,1) ON CONFLICT(key_id,window) DO UPDATE SET hits=hits+1 RETURNING hits', key.id, window).first();
  if (rate.hits > 120) fail(429, 'rate_limit_exceeded', 'Limit of 120 requests per minute per key exceeded.');
  await stmt(db, 'UPDATE api_keys SET last_used_at=? WHERE id=?', now(), key.id).run();
  return { ...key, scopes: JSON.parse(key.scopes) };
}
async function resource(db, auth, kind, resourceId) {
  const row = await stmt(db, 'SELECT data FROM api_resources WHERE id=? AND user_id=? AND mode=? AND kind=?', resourceId, auth.user_id, auth.mode, kind).first();
  if (!row) fail(404, 'not_found', `${kind} not found.`);
  return JSON.parse(row.data);
}
async function listResources(db, auth, kind, url) {
  const limit = Number(url.searchParams.get('limit') || 25);
  if (!Number.isInteger(limit) || limit < 1 || limit > 100) fail(422, 'invalid_request', 'limit must be between 1 and 100.');
  const after = url.searchParams.get('after') || '';
  const { results } = await stmt(db, 'SELECT id,data FROM api_resources WHERE user_id=? AND mode=? AND kind=? AND id>? ORDER BY id LIMIT ?', auth.user_id, auth.mode, kind, after, limit + 1).all();
  const page = results.slice(0, limit);
  return { data: page.map(row => JSON.parse(row.data)), next_cursor: results.length > limit ? page.at(-1).id : null };
}
async function replay(db, auth, key, fingerprint) {
  const record = await stmt(db, 'SELECT * FROM api_idempotency WHERE user_id=? AND mode=? AND key=?', auth.user_id, auth.mode, key).first();
  if (!record) return null;
  if (record.request_hash !== fingerprint) fail(409, 'idempotency_conflict', 'This Idempotency-Key was already used with a different request.');
  return json(JSON.parse(record.response), record.status, { 'Idempotency-Replayed': 'true' });
}
function insertResource(db, auth, kind, data, parent = null) {
  return stmt(db, 'INSERT INTO api_resources(id,user_id,mode,kind,parent_id,data,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)', data.id, auth.user_id, auth.mode, kind, parent, JSON.stringify(data), data.created_at, data.created_at);
}
function eventStatements(db, auth, type, data) {
  const event = { id: id('evt'), type, livemode: auth.mode === 'live', created_at: now(), data };
  return [stmt(db, 'INSERT INTO api_events(id,user_id,mode,type,data,created_at) VALUES(?,?,?,?,?,?)', event.id, auth.user_id, auth.mode, type, JSON.stringify(event), event.created_at),
    stmt(db, `INSERT INTO api_deliveries(id,event_id,webhook_id,next_attempt_at) SELECT ? || id,?,id,? FROM api_webhooks WHERE user_id=? AND mode=? AND disabled_at IS NULL AND EXISTS (SELECT 1 FROM json_each(api_webhooks.events) WHERE value=?)`, `${event.id}_`, event.id, Date.now(), auth.user_id, auth.mode, type)];
}
async function commit(db, auth, operation, statements, data, type, status = 201) {
  const response = { data };
  try {
    // The response reservation and all side effects commit together, including the webhook outbox.
    await db.batch([stmt(db, 'INSERT INTO api_idempotency(user_id,mode,key,request_hash,status,response,created_at) VALUES(?,?,?,?,?,?,?)', auth.user_id, auth.mode, operation.key, operation.fingerprint, status, JSON.stringify(response), now()), ...statements, ...(type ? eventStatements(db, auth, type, data) : [])]);
  } catch (error) {
    const previous = await replay(db, auth, operation.key, operation.fingerprint);
    if (previous) return previous;
    if (/UNIQUE constraint failed/.test(error.message)) fail(409, 'resource_conflict', 'This operation has already been completed for this resource.');
    throw error;
  }
  return json(response, status);
}
function sandboxOnly(auth) { if (auth.mode !== 'test') fail(503, 'carrier_not_configured', 'Live carrier booking is not configured. Use a test key for sandbox labels. No charge was made.'); }
function rate(shipment) {
  return { id: `rate_test_${shipment.id}`, shipment_id: shipment.id, carrier: 'shipide_sandbox', service: 'sandbox_standard', currency: 'EUR', amount: 5.00, test: true };
}
export function webhookUrl(raw) {
  let url; try { url = new URL(raw); } catch { fail(422, 'invalid_webhook_url', 'A public HTTPS webhook URL is required.'); }
  const host = url.hostname.toLowerCase();
  if (url.protocol !== 'https:' || url.username || url.password || url.hash || (url.port && url.port !== '443') || !host.includes('.') || /(^|\.)(localhost|local|internal|test|invalid|example)$/.test(host) || /^[\d.]+$/.test(host) || host.includes(':') || host.startsWith('[')) fail(422, 'invalid_webhook_url', 'Use a public HTTPS hostname on port 443 without credentials or a fragment.');
  return url.href;
}
async function routes(request, env, auth, url) {
  const db = env.CUSTOMER_API_DB, method = request.method;
  const parts = url.pathname.replace(/\/+$/, '').split('/').slice(3);
  const [collection, resourceId, action] = parts;
  if (parts.length > 3) fail(404, 'not_found', 'API route not found.');
  if (method === 'GET' && collection === 'account' && !resourceId) {
    authorize(auth, 'account:read');
    return json({ data: { id: auth.user_id, mode: auth.mode, scopes: auth.scopes, capabilities: { orders: true, shipments: true, labels: auth.mode === 'test', rates: auth.mode === 'test', tracking: 'sandbox_only', live_carrier_booking: false }, docs_url: 'https://docs.shipide.com' } });
  }
  if (['orders', 'shipments', 'labels'].includes(collection) && method === 'GET') {
    authorize(auth, `${collection}:read`);
    const kind = { orders: 'order', shipments: 'shipment', labels: 'label' }[collection];
    if (!resourceId && !action) return json(await listResources(db, auth, kind, url));
    const item = await resource(db, auth, kind, resourceId);
    if (!action) return json({ data: item });
    if (collection === 'labels' && action === 'file') {
      if (item.status === 'voided') fail(409, 'label_voided', 'This label has been voided.');
      const shipment = await resource(db, auth, 'shipment', item.shipment_id);
      return new Response(await sandboxPdf(item, shipment), { headers: { 'Content-Type': 'application/pdf', 'Content-Disposition': `attachment; filename="${item.id}.pdf"`, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' } });
    }
  }
  if (collection === 'tracking' && resourceId && !action && method === 'GET') {
    authorize(auth, 'shipments:read');
    const row = await stmt(db, "SELECT data FROM api_resources WHERE user_id=? AND mode=? AND kind='label' AND json_extract(data,'$.tracking_number')=?", auth.user_id, auth.mode, resourceId).first();
    if (!row) fail(404, 'not_found', 'Tracking number not found.');
    const label = JSON.parse(row.data);
    return json({ data: { tracking_number: label.tracking_number, shipment_id: label.shipment_id, status: label.status === 'voided' ? 'voided' : 'pre_transit', test: true, events: [] } });
  }
  if (collection === 'webhooks' && method === 'GET' && !action) {
    authorize(auth, 'webhooks:read');
    if (resourceId) {
      const row = await stmt(db, 'SELECT id,url,events,created_at,disabled_at FROM api_webhooks WHERE id=? AND user_id=? AND mode=?', resourceId, auth.user_id, auth.mode).first();
      if (!row) fail(404, 'not_found', 'Webhook not found.');
      return json({ data: { ...row, events: JSON.parse(row.events) } });
    }
    const { results } = await stmt(db, 'SELECT id,url,events,created_at,disabled_at FROM api_webhooks WHERE user_id=? AND mode=? ORDER BY created_at DESC LIMIT 100', auth.user_id, auth.mode).all();
    return json({ data: results.map(row => ({ ...row, events: JSON.parse(row.events) })) });
  }
  const supported = (method === 'POST' && ((['orders', 'shipments', 'rates', 'labels', 'webhooks'].includes(collection) && !resourceId) || (collection === 'labels' && resourceId && action === 'void'))) || (method === 'DELETE' && collection === 'webhooks' && resourceId && !action);
  if (!supported) fail(404, 'not_found', 'API route or method not found.');
  authorize(auth, collection === 'rates' ? 'rates:read' : `${collection}:write`);
  const body = method === 'DELETE' ? {} : await readBody(request);
  const key = required(request.headers.get('idempotency-key'), 'Idempotency-Key', 128);
  const operation = { key, fingerprint: await hash(`${method}\n${url.pathname.replace(/\/+$/, '')}\n${canonical(body)}`) };
  const previous = await replay(db, auth, key, operation.fingerprint); if (previous) return previous;
  if (['orders', 'shipments'].includes(collection)) {
    const input = shipmentInput(body);
    if (collection === 'orders' && input.order_id) fail(422, 'invalid_request', 'An order cannot contain order_id.');
    if (input.order_id) await resource(db, auth, 'order', input.order_id);
    const kind = collection === 'orders' ? 'order' : 'shipment';
    const data = { id: id(kind === 'order' ? 'ord' : 'shp'), object: kind, ...input, status: kind === 'order' ? 'received' : 'draft', livemode: auth.mode === 'live', created_at: now() };
    return commit(db, auth, operation, [insertResource(db, auth, kind, data)], data, `${kind}.created`);
  }
  if (collection === 'rates') {
    only(body, ['shipment_id']); sandboxOnly(auth);
    const shipment = await resource(db, auth, 'shipment', required(body.shipment_id, 'shipment_id', 64));
    return commit(db, auth, operation, [], [rate(shipment)], null, 200);
  }
  if (collection === 'labels' && !resourceId) {
    only(body, ['shipment_id', 'rate_id', 'format']); sandboxOnly(auth);
    if (body.format !== undefined && body.format !== 'pdf') fail(422, 'unsupported_format', 'Only PDF labels are currently supported.');
    const shipment = await resource(db, auth, 'shipment', required(body.shipment_id, 'shipment_id', 64));
    if (shipment.status !== 'draft') fail(409, 'shipment_not_draft', 'Create a new shipment to purchase another label.');
    if (body.rate_id !== rate(shipment).id) fail(422, 'invalid_rate', 'Use a rate_id returned for this shipment.');
    const labelId = id('lbl');
    const data = { id: labelId, object: 'label', shipment_id: shipment.id, status: 'created', carrier: 'shipide_sandbox', tracking_number: `TEST${crypto.randomUUID().replaceAll('-', '').slice(0, 20).toUpperCase()}`, format: 'pdf', download_path: `/api/v1/labels/${labelId}/file`, charge: { currency: 'EUR', amount: 0 }, test: true, created_at: now() };
    const updated = { ...shipment, status: 'label_created', label_id: labelId, tracking_number: data.tracking_number };
    return commit(db, auth, operation, [insertResource(db, auth, 'label', data, shipment.id), stmt(db, 'UPDATE api_resources SET data=?,updated_at=? WHERE id=? AND user_id=? AND mode=?', JSON.stringify(updated), now(), shipment.id, auth.user_id, auth.mode)], data, 'label.created');
  }
  if (collection === 'labels' && action === 'void') {
    only(body, []); sandboxOnly(auth);
    const label = await resource(db, auth, 'label', resourceId);
    if (label.status === 'voided') return commit(db, auth, operation, [], label, null, 200);
    const data = { ...label, status: 'voided', voided_at: now() };
    const shipment = await resource(db, auth, 'shipment', label.shipment_id);
    return commit(db, auth, operation, [stmt(db, 'UPDATE api_resources SET data=?,updated_at=? WHERE id=?', JSON.stringify(data), now(), label.id), stmt(db, 'UPDATE api_resources SET data=?,updated_at=? WHERE id=?', JSON.stringify({ ...shipment, status: 'voided' }), now(), shipment.id)], data, 'label.voided', 200);
  }
  if (collection === 'webhooks' && method === 'POST') {
    only(body, ['url', 'events']);
    const webhook = { id: id('wh'), url: webhookUrl(required(body.url, 'url', 2000)), events: body.events, created_at: now(), secret: randomSecret('whsec_') };
    if (!Array.isArray(body.events) || !body.events.length || body.events.length > EVENTS.length || body.events.some(event => !EVENTS.includes(event))) fail(422, 'invalid_request', `events must contain one or more of: ${EVENTS.join(', ')}.`);
    const count = await stmt(db, 'SELECT count(*) AS n FROM api_webhooks WHERE user_id=? AND mode=? AND disabled_at IS NULL', auth.user_id, auth.mode).first();
    if (count.n >= 10) fail(422, 'webhook_limit', 'At most 10 active webhooks are allowed per environment.');
    return commit(db, auth, operation, [stmt(db, 'INSERT INTO api_webhooks(id,user_id,mode,url,secret,events,created_at) VALUES(?,?,?,?,?,?,?)', webhook.id, auth.user_id, auth.mode, webhook.url, webhook.secret, JSON.stringify(webhook.events), webhook.created_at)], webhook, null);
  }
  if (collection === 'webhooks' && method === 'DELETE') {
    const row = await stmt(db, 'SELECT id FROM api_webhooks WHERE id=? AND user_id=? AND mode=?', resourceId, auth.user_id, auth.mode).first();
    if (!row) fail(404, 'not_found', 'Webhook not found.');
    return commit(db, auth, operation, [stmt(db, 'UPDATE api_webhooks SET disabled_at=? WHERE id=?', now(), resourceId)], { id: resourceId, disabled: true }, null, 200);
  }
  fail(404, 'not_found', 'API route not found.');
}
export async function handleCustomerApi(request, env) {
  const requestId = id('req');
  try {
    if (!env.CUSTOMER_API_DB) fail(503, 'api_unavailable', 'Customer API is not configured.');
    const auth = await authenticate(request, env.CUSTOMER_API_DB);
    const response = await routes(request, env, auth, new URL(request.url));
    response.headers.set('X-Request-Id', requestId); return response;
  } catch (error) {
    if (!(error instanceof ApiError)) console.error(JSON.stringify({ event: 'customer_api_error', request_id: requestId, message: error.message }));
    return json({ error: { code: error.code || 'internal_error', message: error instanceof ApiError ? error.message : 'An unexpected error occurred. Contact support with the request_id.', request_id: requestId } }, error.status || 500, { 'X-Request-Id': requestId, ...(error.status === 429 ? { 'Retry-After': '60' } : {}) });
  }
}
export async function handleDeveloperApi(request, env, authenticateUser) {
  try {
    const user = await authenticateUser(request, env);
    if (!user?.id) fail(401, 'authentication_required', 'Sign in to your Shipide account.');
    if (!env.CUSTOMER_API_DB) fail(503, 'api_unavailable', 'Customer API is not configured.');
    const db = env.CUSTOMER_API_DB;
    const path = new URL(request.url).pathname.replace(/\/+$/, '');
    const fields = 'id,name,prefix,mode,scopes,created_at,expires_at,revoked_at,last_used_at';
    if (path === '/api/developer/keys' && request.method === 'GET') {
      const { results } = await stmt(db, `SELECT ${fields} FROM api_keys WHERE user_id=? ORDER BY created_at DESC LIMIT 100`, user.id).all();
      return json({ data: results.map(row => ({ ...row, scopes: JSON.parse(row.scopes) })), scopes: SCOPES });
    }
    if (path === '/api/developer/keys' && request.method === 'POST') {
      const body = await readBody(request); only(body, ['name', 'mode', 'scopes', 'expires_at']);
      const name = required(body.name, 'name', 80), mode = body.mode || 'test', scopes = body.scopes || SCOPES;
      if (!['test', 'live'].includes(mode)) fail(422, 'invalid_request', 'mode must be test or live.');
      if (!Array.isArray(scopes) || !scopes.length || scopes.some(scope => !SCOPES.includes(scope))) fail(422, 'invalid_request', 'Select valid scopes.');
      let expires = null;
      if (body.expires_at) { const date = new Date(body.expires_at); if (!Number.isFinite(date.getTime()) || date <= new Date()) fail(422, 'invalid_request', 'expires_at must be a future timestamp.'); expires = date.toISOString(); }
      const secret = randomSecret(`shipide_${mode}_`);
      const data = { id: id('key'), name, mode, scopes: [...new Set(scopes)], prefix: secret.slice(0, 21), created_at: now(), expires_at: expires };
      const inserted = await stmt(db, `INSERT INTO api_keys(id,user_id,name,key_hash,prefix,mode,scopes,created_at,expires_at) SELECT ?,?,?,?,?,?,?,?,? WHERE (SELECT count(*) FROM api_keys WHERE user_id=? AND revoked_at IS NULL) < 20 RETURNING id`, data.id, user.id, name, await hash(secret), data.prefix, mode, JSON.stringify(data.scopes), data.created_at, expires, user.id).first();
      if (!inserted) fail(422, 'key_limit', 'Revoke an existing key before creating more than 20 active keys.');
      return json({ data: { ...data, key: secret } }, 201);
    }
    if (/^\/api\/developer\/keys\/key_[a-f0-9]{32}$/.test(path) && request.method === 'DELETE') {
      const keyId = path.split('/').at(-1);
      const updated = await stmt(db, 'UPDATE api_keys SET revoked_at=COALESCE(revoked_at,?) WHERE id=? AND user_id=? RETURNING id', now(), keyId, user.id).first();
      if (!updated) fail(404, 'not_found', 'API key not found.');
      return json({ data: { id: keyId, revoked: true } });
    }
    fail(404, 'not_found', 'Developer route not found.');
  } catch (error) {
    if (!(error instanceof ApiError)) console.error('developer_api_error', error.message);
    return json({ error: { code: error.code || 'internal_error', message: error instanceof ApiError ? error.message : 'Could not manage API keys.' } }, error.status || 500);
  }
}
export async function sandboxPdf(label, shipment) {
  const pdf = await PDFDocument.create(), font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage([288, 432]);
  const lines = ['SHIPIDE SANDBOX', 'TEST LABEL - NOT VALID FOR SHIPPING', '', `Shipment: ${shipment.id}`, `Reference: ${shipment.reference}`, '', 'FROM', shipment.from.name, shipment.from.street1, `${shipment.from.postal_code} ${shipment.from.city}`, shipment.from.country, '', 'TO', shipment.to.name, shipment.to.street1, `${shipment.to.postal_code} ${shipment.to.city}`, shipment.to.country, '', `Weight: ${shipment.parcel.weight_kg} kg`, label.tracking_number, '', 'No carrier booking. No payment charged.'];
  let y = 409;
  for (const line of lines) {
    const safe = line.replace(/[^\x20-\x7e]/g, '?');
    const size = Math.min(10, 254 / Math.max(font.widthOfTextAtSize(safe, 1), 1));
    page.drawText(safe, { x: 17, y, size, font, color: rgb(0.12, 0.14, 0.16) }); y -= 17;
  }
  return pdf.save();
}
export async function webhookSignature(secret, timestamp, body) {
  const key = await crypto.subtle.importKey('raw', encoder.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return hex(await crypto.subtle.sign('HMAC', key, encoder.encode(`${timestamp}.${body}`)));
}
export async function dispatchWebhooks(env) {
  if (!env.CUSTOMER_API_DB) return;
  const db = env.CUSTOMER_API_DB, timestamp = Date.now();
  const { results } = await stmt(db, `UPDATE api_deliveries SET lease_until=? WHERE id IN (SELECT d.id FROM api_deliveries d JOIN api_webhooks w ON w.id=d.webhook_id WHERE d.delivered_at IS NULL AND d.failed_at IS NULL AND d.next_attempt_at<=? AND d.lease_until<? AND w.disabled_at IS NULL ORDER BY d.next_attempt_at LIMIT 20) RETURNING *`, timestamp + 60000, timestamp, timestamp).all();
  await Promise.allSettled(results.map(async delivery => {
    const record = await stmt(db, 'SELECT w.url,w.secret,w.disabled_at,e.data FROM api_webhooks w JOIN api_events e ON e.id=? WHERE w.id=?', delivery.event_id, delivery.webhook_id).first();
    if (!record || record.disabled_at) return;
    let status = 0;
    try {
      const url = webhookUrl(record.url);
      const time = String(Math.floor(Date.now() / 1000));
      // Workers fetch uses public routing; redirects are never followed to another destination.
      const response = await fetch(url, { method: 'POST', redirect: 'manual', signal: AbortSignal.timeout(10000), headers: { 'Content-Type': 'application/json', 'User-Agent': 'Shipide-Webhooks/1.0', 'Shipide-Event-Id': delivery.event_id, 'Shipide-Signature': `t=${time},v1=${await webhookSignature(record.secret, time, record.data)}` }, body: record.data });
      status = response.status; await response.body?.cancel();
    } catch { status = 0; }
    const attempts = delivery.attempts + 1, success = status >= 200 && status < 300;
    await stmt(db, 'UPDATE api_deliveries SET attempts=?,last_status=?,delivered_at=?,failed_at=?,next_attempt_at=?,lease_until=0 WHERE id=? AND lease_until=?', attempts, status, success ? now() : null, !success && attempts >= 8 ? now() : null, Date.now() + Math.min(3600000, 60000 * 2 ** (attempts - 1)), delivery.id, delivery.lease_until).run();
  }));
  await stmt(db, 'DELETE FROM api_rate_limits WHERE window < ?', Math.floor(Date.now() / 60000) - 5).run();
}
