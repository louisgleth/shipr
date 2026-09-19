import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readFileSync } from 'node:fs';
import { PDFDocument } from 'pdf-lib';
import { handleCustomerApi, handleDeveloperApi, dispatchWebhooks, webhookSignature, webhookUrl } from '../customer-api.mjs';

// Exercise the actual SQL and transaction rollback semantics, not a query-string mock.
function database() {
  const sqlite = new DatabaseSync(':memory:');
  sqlite.exec(readFileSync(new URL('../migrations/customer-api/0001_customer_api.sql', import.meta.url), 'utf8'));
  class Statement {
    constructor(sql, args = []) { this.sql = sql; this.args = args; }
    bind(...args) { return new Statement(this.sql, args); }
    async first() { return sqlite.prepare(this.sql).get(...this.args) || null; }
    async all() { return { results: sqlite.prepare(this.sql).all(...this.args) }; }
    async run() { return sqlite.prepare(this.sql).run(...this.args); }
  }
  return { prepare: sql => new Statement(sql), async batch(statements) { sqlite.exec('BEGIN'); try { const result = []; for (const statement of statements) result.push(await statement.run()); sqlite.exec('COMMIT'); return result; } catch (error) { sqlite.exec('ROLLBACK'); throw error; } }, sqlite };
}
function harness() {
  const db = database(), env = { CUSTOMER_API_DB: db };
  const request = (path, method = 'GET', body, token, idem) => new Request(`https://portal.shipide.com${path}`, { method, headers: { ...(body !== undefined ? { 'Content-Type': 'application/json' } : {}), ...(token ? { Authorization: `Bearer ${token}` } : {}), ...(idem ? { 'Idempotency-Key': idem } : {}) }, ...(body !== undefined ? { body: JSON.stringify(body) } : {}) });
  const developer = (user, path = '/api/developer/keys', method = 'GET', body) => handleDeveloperApi(request(path, method, body), env, async () => user ? { id: user } : null);
  const api = (key, path, method, body, idem) => handleCustomerApi(request(`/api/v1${path}`, method, body, key, idem), env);
  const key = async (user = 'alice', mode = 'test', scopes) => (await (await developer(user, '/api/developer/keys', 'POST', { name: 'Warehouse', mode, ...(scopes ? { scopes } : {}) })).json()).data;
  return { db, env, developer, api, key };
}
const input = { reference: 'order-1001', from: { name: 'Test warehouse', street1: 'Test street 1', city: 'Brussels', postal_code: '1000', country: 'BE' }, to: { name: 'Test recipient', street1: 'Example street 2', city: 'Antwerp', postal_code: '2000', country: 'BE' }, parcel: { weight_kg: 1, length_cm: 20, width_cm: 15, height_cm: 10 } };
async function shipment(h, key, idem = 'shipment-1') { const response = await h.api(key, '/shipments', 'POST', input, idem); assert.equal(response.status, 201); return (await response.json()).data; }
async function label(h, key, s) { const response = await h.api(key, '/rates', 'POST', { shipment_id: s.id }, 'rates-1'); assert.equal(response.status, 200); const rates = (await response.json()).data; return h.api(key, '/labels', 'POST', { shipment_id: s.id, rate_id: rates[0].id, format: 'pdf' }, 'label-1'); }

test('session authentication and one-time hashed scoped API keys', async () => {
  const h = harness(); assert.equal((await h.developer(null)).status, 401);
  const k = await h.key(); assert.match(k.key, /^shipide_test_[a-f0-9]{64}$/);
  const stored = h.db.sqlite.prepare('SELECT * FROM api_keys').get(); assert.notEqual(stored.key_hash, k.key); assert.equal(stored.key_hash.length, 64);
  const listing = await (await h.developer('alice')).json(); assert.equal(listing.data.length, 1); assert.equal(JSON.stringify(listing).includes(k.key), false); assert.equal(listing.data[0].key_hash, undefined);
  assert.equal((await h.api(null, '/account')).status, 401);
  assert.equal((await h.api(k.key, '/account')).status, 200);
  const readonly = await h.key('alice', 'test', ['shipments:read']); assert.equal((await h.api(readonly.key, '/shipments', 'POST', input, 'write')).status, 403);
  assert.equal((await h.developer('bob', `/api/developer/keys/${k.id}`, 'DELETE')).status, 404);
  assert.equal((await h.developer('alice', `/api/developer/keys/${k.id}`, 'DELETE')).status, 200);
  assert.equal((await h.api(k.key, '/account')).status, 401);
});
test('owner and environment isolation for every resource read', async () => {
  const h = harness(), a = await h.key(), b = await h.key('bob'), live = await h.key('alice', 'live');
  const s = await shipment(h, a.key);
  for (const key of [b.key, live.key]) { assert.equal((await h.api(key, `/shipments/${s.id}`)).status, 404); assert.deepEqual((await (await h.api(key, '/shipments')).json()).data, []); }
  const l = (await (await label(h, a.key, s)).json()).data;
  for (const key of [b.key, live.key]) for (const path of [`/labels/${l.id}`, `/labels/${l.id}/file`, `/tracking/${l.tracking_number}`]) assert.equal((await h.api(key, path)).status, 404);
});
test('idempotency preserves the response and rejects conflicting requests', async () => {
  const h = harness(), k = await h.key(), s = await shipment(h, k.key);
  const retry = await h.api(k.key, '/shipments', 'POST', { ...input }, 'shipment-1'); assert.equal(retry.status, 201); assert.equal(retry.headers.get('Idempotency-Replayed'), 'true'); assert.deepEqual((await retry.json()).data, s);
  assert.equal((await h.api(k.key, '/shipments', 'POST', { ...input, reference: 'other' }, 'shipment-1')).status, 409);
  assert.equal((await h.api(k.key, '/orders', 'POST', input, 'shipment-1')).status, 409);
  assert.equal(h.db.sqlite.prepare('SELECT count(*) AS n FROM api_resources').get().n, 1);
  assert.equal((await h.api(k.key, '/shipments', 'POST', input)).status, 422);
});
test('sandbox lifecycle creates real PDF bytes, tracks and voids without charges', async () => {
  const h = harness(), k = await h.key(), s = await shipment(h, k.key);
  const response = await label(h, k.key, s); assert.equal(response.status, 201); const l = (await response.json()).data;
  assert.equal(l.test, true); assert.equal(l.charge.amount, 0); assert.match(l.tracking_number, /^TEST/);
  const file = await h.api(k.key, `/labels/${l.id}/file`); assert.equal(file.headers.get('content-type'), 'application/pdf'); assert.equal((await PDFDocument.load(await file.arrayBuffer())).getPageCount(), 1);
  assert.equal((await (await h.api(k.key, `/tracking/${l.tracking_number}`)).json()).data.status, 'pre_transit');
  const duplicate = await h.api(k.key, '/labels', 'POST', { shipment_id: s.id, rate_id: `rate_test_${s.id}` }, 'different-label-key'); assert.equal(duplicate.status, 409);
  const voided = await h.api(k.key, `/labels/${l.id}/void`, 'POST', {}, 'void-1'); assert.equal(voided.status, 200);
  assert.equal((await h.api(k.key, `/labels/${l.id}/file`)).status, 409);
  assert.equal((await (await h.api(k.key, `/tracking/${l.tracking_number}`)).json()).data.status, 'voided');
  const replay = await label(h, k.key, s); assert.equal((await replay.json()).data.id, l.id);
});
test('live intake works but cannot fabricate a carrier booking or charge', async () => {
  const h = harness(), k = await h.key('alice', 'live'), s = await shipment(h, k.key);
  assert.equal(s.livemode, true);
  for (const path of ['/rates', '/labels']) { const response = await h.api(k.key, path, 'POST', { shipment_id: s.id }, path); assert.equal(response.status, 503); assert.equal((await response.json()).error.code, 'carrier_not_configured'); }
  assert.equal(h.db.sqlite.prepare("SELECT count(*) AS n FROM api_resources WHERE kind='label'").get().n, 0);
});
test('invalid input, unknown fields, unsupported formats, and body limits', async () => {
  const h = harness(), k = await h.key();
  for (const body of [{ ...input, parcel: { ...input.parcel, weight_kg: -1 } }, { ...input, total_price: 0 }, { ...input, to: { ...input.to, country: 'Belgium' } }, { ...input, reference: '' }, { ...input, parcel: null }]) assert.equal((await h.api(k.key, '/shipments', 'POST', body, crypto.randomUUID())).status, 422);
  assert.equal((await h.api(k.key, '/shipments', 'POST', { ...input, metadata: { large: 'a'.repeat(70000) } }, 'large')).status, 413);
  const s = await shipment(h, k.key); assert.equal((await h.api(k.key, '/labels', 'POST', { shipment_id: s.id, rate_id: `rate_test_${s.id}`, format: 'zpl' }, 'zpl')).status, 422);
  assert.equal((await h.api(k.key, '/shipments?limit=101')).status, 422);
  assert.equal((await h.api(k.key, '/does-not-exist')).status, 404);
});
test('orders can be linked only within their account and environment', async () => {
  const h = harness(), k = await h.key(), b = await h.key('bob');
  const order = (await (await h.api(k.key, '/orders', 'POST', input, 'order')).json()).data;
  assert.equal((await h.api(k.key, '/shipments', 'POST', { ...input, order_id: order.id }, 'linked')).status, 201);
  assert.equal((await h.api(b.key, '/shipments', 'POST', { ...input, order_id: order.id }, 'linked')).status, 404);
});
test('rate limits and expiry are enforced', async () => {
  const h = harness(), k = await h.key();
  h.db.sqlite.prepare('INSERT INTO api_rate_limits VALUES(?,?,?)').run(k.id, Math.floor(Date.now()/60000), 120);
  const response = await h.api(k.key, '/account'); assert.equal(response.status, 429); assert.equal(response.headers.get('Retry-After'), '60');
  h.db.sqlite.prepare('UPDATE api_keys SET expires_at=? WHERE id=?').run('2020-01-01T00:00:00.000Z', k.id);
  assert.equal((await h.api(k.key, '/account')).status, 401);
});
test('atomic outbox emits one event for a retried shipment and signs deliveries', async () => {
  const h = harness(), k = await h.key();
  const hookResponse = await h.api(k.key, '/webhooks', 'POST', { url: 'https://warehouse.customer.com/events', events: ['shipment.created'] }, 'hook'); assert.equal(hookResponse.status, 201); const hook = (await hookResponse.json()).data;
  await shipment(h, k.key); await shipment(h, k.key);
  assert.equal(h.db.sqlite.prepare('SELECT count(*) AS n FROM api_deliveries').get().n, 1);
  const original = globalThis.fetch;
  let calls = 0;
  globalThis.fetch = async (url, options) => { calls++; assert.equal(url, hook.url); assert.equal(options.redirect, 'manual'); const match = options.headers['Shipide-Signature'].match(/^t=(\d+),v1=([a-f0-9]+)$/); assert.equal(match[2], await webhookSignature(hook.secret, match[1], options.body)); return new Response('', { status: 200 }); };
  try { await dispatchWebhooks(h.env); await dispatchWebhooks(h.env); } finally { globalThis.fetch = original; }
  assert.equal(calls, 1); assert.ok(h.db.sqlite.prepare('SELECT delivered_at FROM api_deliveries').get().delivered_at);
  const listed = await (await h.api(k.key, '/webhooks')).json(); assert.equal(listed.data[0].secret, undefined);
});
test('webhook failures are retried and disable stops delivery', async () => {
  const h = harness(), k = await h.key();
  const hook = (await (await h.api(k.key, '/webhooks', 'POST', { url: 'https://warehouse.customer.com/events', events: ['shipment.created'] }, 'hook')).json()).data;
  await shipment(h, k.key);
  const original = globalThis.fetch; globalThis.fetch = async () => new Response(null, { status: 500 });
  try { await dispatchWebhooks(h.env); } finally { globalThis.fetch = original; }
  const delivery = h.db.sqlite.prepare('SELECT * FROM api_deliveries').get(); assert.equal(delivery.attempts, 1); assert.equal(delivery.delivered_at, null); assert.ok(delivery.next_attempt_at > Date.now());
  assert.equal((await h.api(k.key, `/webhooks/${hook.id}`, 'DELETE', undefined, 'delete')).status, 200);
  h.db.sqlite.prepare('UPDATE api_deliveries SET next_attempt_at=0').run();
  globalThis.fetch = async () => { throw new Error('Disabled webhook should never be called'); };
  try { await dispatchWebhooks(h.env); } finally { globalThis.fetch = original; }
  assert.equal(h.db.sqlite.prepare('SELECT attempts FROM api_deliveries').get().attempts, 1);
});
test('webhook URL validation rejects local addresses, credentials and redirects as destinations', () => {
  for (const url of ['http://example.com', 'https://127.0.0.1/x', 'https://[::1]/', 'https://localhost/', 'https://service.internal/', 'https://user:pass@customer.com/', 'https://customer.com:8080/', 'https://customer.com/#fragment']) assert.throws(() => webhookUrl(url));
  assert.equal(webhookUrl('https://warehouse.customer.com/hooks'), 'https://warehouse.customer.com/hooks');
});
test('failed atomic mutation leaves no idempotency response or event', async () => {
  const h = harness(), k = await h.key();
  h.db.sqlite.exec("CREATE TRIGGER simulate_failure BEFORE INSERT ON api_resources BEGIN SELECT RAISE(ABORT, 'storage failure'); END");
  assert.equal((await h.api(k.key, '/shipments', 'POST', input, 'failure')).status, 500);
  assert.equal(h.db.sqlite.prepare('SELECT count(*) AS n FROM api_idempotency').get().n, 0);
  assert.equal(h.db.sqlite.prepare('SELECT count(*) AS n FROM api_events').get().n, 0);
});
