(() => {
  'use strict';
  const el = id => document.getElementById(id);
  const client = window.supabase?.createClient('https://pxcqxubehvnyaubqjcrf.supabase.co', 'sb_publishable_MfL9s44GmmR9peD1jouetw_aZOa8xVa', { auth: { storageKey:'sb-pxcqxubehvnyaubqjcrf-auth-token-shipr', persistSession:true, autoRefreshToken:true } });
  let selectedKey = null;
  async function api(path, options = {}) {
    const { data: { session } } = await client.auth.getSession();
    if (!session) throw new Error('Sign in to your Shipide account to manage API keys.');
    const response = await fetch(path, { ...options, headers: { 'Content-Type':'application/json', Authorization:`Bearer ${session.access_token}` } });
    const body = await response.json();
    if (!response.ok) throw new Error(body.error?.message || 'Could not complete the request.');
    return body;
  }
  function text(tag, value, className) { const node = document.createElement(tag); node.textContent = value; if (className) node.className = className; return node; }
  function date(value) { return value ? new Date(value).toLocaleDateString() : 'Never'; }
  async function refresh() {
    try {
      if (!client) throw new Error('Could not load sign-in. Reload the page.');
      const { data: keys, scopes } = await api('/api/developer/keys');
      el('workspace').hidden = false; el('login').hidden = true; el('status').textContent = '';
      el('keyList').replaceChildren();
      if (!keys.length) el('keyList').append(text('p', 'No API keys yet.'));
      for (const key of keys) {
        const row = document.createElement('div'); row.className = 'key-row'; row.dataset.revoked = String(Boolean(key.revoked_at));
        const name = text('div', key.name); name.append(text('small', key.revoked_at ? 'Revoked' : key.expires_at && Date.parse(key.expires_at) <= Date.now() ? 'Expired' : `Last used: ${date(key.last_used_at)}`));
        const revoke = text('button', 'Revoke', 'btn btn-secondary btn-sm'); revoke.type = 'button'; revoke.disabled = Boolean(key.revoked_at); revoke.setAttribute('aria-label', `Revoke ${key.name}`);
        revoke.addEventListener('click', () => { selectedKey = key; el('revokeName').textContent = key.name; el('revokeDialog').returnValue = ''; el('revokeDialog').showModal(); });
        row.append(name, text('code', `${key.prefix}...`), text('span', key.mode === 'test' ? 'Sandbox' : 'Live'), text('small', `Expires: ${date(key.expires_at)}`, 'key-date'), revoke); el('keyList').append(row);
      }
      if (!el('scopes').children.length) for (const scope of scopes) {
        const label = document.createElement('label'), input = document.createElement('input'); input.type = 'checkbox'; input.value = scope; input.name = 'scope'; input.defaultChecked = true; label.append(input, document.createTextNode(scope)); el('scopes').append(label);
      }
    } catch (error) { el('status').textContent = error.message; el('workspace').hidden = true; el('login').hidden = false; }
  }
  el('newKey').addEventListener('click', () => { el('keyForm').reset(); el('formError').textContent = ''; el('keyForm').hidden = false; el('keyResult').hidden = true; el('keyDialog').showModal(); });
  const close = () => el('keyDialog').close();
  el('closeDialog').addEventListener('click', close); el('doneKey').addEventListener('click', close);
  el('keyDialog').addEventListener('close', () => { el('secret').textContent = ''; el('copyKey').textContent = 'Copy key'; });
  el('copyKey').addEventListener('click', async () => { try { await navigator.clipboard.writeText(el('secret').textContent); el('copyKey').textContent = 'Copied'; } catch { el('copyKey').textContent = 'Copy unavailable'; } });
  el('keyForm').addEventListener('submit', async event => {
    event.preventDefault(); el('createKey').disabled = true; el('formError').textContent = '';
    try {
      const days = Number(el('keyExpiry').value);
      const { data } = await api('/api/developer/keys', { method:'POST', body:JSON.stringify({ name:el('keyName').value, mode:el('keyMode').value, scopes:Array.from(document.querySelectorAll('[name="scope"]:checked'), input => input.value), expires_at:days ? new Date(Date.now() + days * 86400000).toISOString() : null }) });
      el('secret').textContent = data.key; el('keyForm').hidden = true; el('keyResult').hidden = false; await refresh();
    } catch (error) { el('formError').textContent = error.message; } finally { el('createKey').disabled = false; }
  });
  el('revokeDialog').addEventListener('close', async () => {
    if (el('revokeDialog').returnValue !== 'revoke' || !selectedKey) return;
    try { await api(`/api/developer/keys/${selectedKey.id}`, { method:'DELETE' }); await refresh(); } catch (error) { el('status').textContent = error.message; }
    selectedKey = null;
  });
  client?.auth.onAuthStateChange(event => { if (event === 'SIGNED_OUT') { el('workspace').hidden = true; el('keyList').replaceChildren(); close(); el('status').textContent = 'You have signed out.'; el('login').hidden = false; } });
  void refresh();
})();
