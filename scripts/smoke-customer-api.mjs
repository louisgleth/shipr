import assert from 'node:assert/strict';

const base = process.env.SHIPIDE_API_ORIGIN || 'https://portal.shipide.com';
async function password() {
  if (process.env.SHIPIDE_TEST_PASSWORD) return process.env.SHIPIDE_TEST_PASSWORD;
  if (!process.stdin.isTTY) throw new Error('Set SHIPIDE_TEST_PASSWORD or run in an interactive terminal.');
  process.stdout.write('Demo account password (hidden): ');
  process.stdin.setRawMode(true); process.stdin.resume();
  return new Promise((resolve, reject) => {
    let value = '';
    const done = () => { process.stdin.setRawMode(false); process.stdin.pause(); process.stdin.off('data', onData); process.stdout.write('\n'); };
    function onData(bytes) {
      for (const char of bytes.toString()) {
        if (char === '\u0003') { done(); reject(new Error('Cancelled')); return; }
        if (char === '\r' || char === '\n') { done(); resolve(value); return; }
        if (char === '\u007f') value = value.slice(0, -1); else value += char;
      }
    }
    process.stdin.on('data', onData);
  });
}
const auth = await fetch('https://pxcqxubehvnyaubqjcrf.supabase.co/auth/v1/token?grant_type=password', { method:'POST', headers:{ 'Content-Type':'application/json', apikey:'sb_publishable_MfL9s44GmmR9peD1jouetw_aZOa8xVa' }, body:JSON.stringify({ email:process.env.SHIPIDE_TEST_EMAIL || 'demo@shipide.com', password:await password() }) });
assert.equal(auth.status, 200, 'Demo sign-in failed. No API test keys were created.');
const { access_token:session } = await auth.json();
const sessionHeaders = { Authorization:`Bearer ${session}`, 'Content-Type':'application/json' };
let keyId;
try {
  const created = await fetch(`${base}/api/developer/keys`, { method:'POST', headers:sessionHeaders, body:JSON.stringify({ name:'API deployment smoke test', mode:'test', expires_at:new Date(Date.now()+3600000).toISOString() }) });
  assert.equal(created.status, 201, 'Create test key');
  const { data:key } = await created.json(); keyId = key.id;
  const run = crypto.randomUUID();
  async function call(path, body, operation, expected = 200) {
    const response = await fetch(`${base}/api/v1${path}`, { method:body ? 'POST' : 'GET', headers:{ Authorization:`Bearer ${key.key}`, ...(body ? { 'Content-Type':'application/json', 'Idempotency-Key':`${run}-${operation}` } : {}) }, ...(body ? { body:JSON.stringify(body) } : {}) });
    assert.equal(response.status, expected, `${path}: unexpected HTTP status`);
    return response;
  }
  assert.equal((await (await call('/account')).json()).data.mode, 'test');
  const address = { name:'API test', street1:'Example street 1', city:'Brussels', postal_code:'1000', country:'BE' };
  const input = { reference:`deployment-test-${run}`, from:address, to:address, parcel:{ weight_kg:1, length_cm:20, width_cm:15, height_cm:10 } };
  const [first, duplicate] = await Promise.all([call('/shipments', input, 'shipment', 201), call('/shipments', input, 'shipment', 201)]);
  const s = (await first.json()).data; assert.equal((await duplicate.json()).data.id, s.id);
  const rate = (await (await call('/rates', { shipment_id:s.id }, 'rates')).json()).data[0];
  const labelBody = { shipment_id:s.id, rate_id:rate.id, format:'pdf' };
  const label = (await (await call('/labels', labelBody, 'label', 201)).json()).data;
  const replay = await call('/labels', labelBody, 'label', 201); assert.equal(replay.headers.get('Idempotency-Replayed'), 'true');
  const pdf = await call(`/labels/${label.id}/file`); assert.equal(pdf.headers.get('content-type'), 'application/pdf'); assert.ok((await pdf.arrayBuffer()).byteLength > 500);
  const tracking = (await (await call(`/tracking/${label.tracking_number}`)).json()).data; assert.equal(tracking.status, 'pre_transit');
  assert.equal((await (await call(`/labels/${label.id}/void`, {}, 'void')).json()).data.status, 'voided');
  console.log('Production sandbox smoke test passed: key, account, concurrent idempotency, rates, label, PDF, tracking, void.');
} finally {
  if (keyId) { const result = await fetch(`${base}/api/developer/keys/${keyId}`, { method:'DELETE', headers:sessionHeaders }); assert.equal(result.status, 200, 'Revoke smoke-test key'); console.log('Temporary API key revoked.'); }
}
