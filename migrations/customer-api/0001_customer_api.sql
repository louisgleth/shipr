CREATE TABLE api_keys (
  id TEXT PRIMARY KEY, user_id TEXT NOT NULL, name TEXT NOT NULL,
  key_hash TEXT NOT NULL UNIQUE, prefix TEXT NOT NULL,
  mode TEXT NOT NULL CHECK(mode IN ('test','live')), scopes TEXT NOT NULL,
  created_at TEXT NOT NULL, expires_at TEXT, revoked_at TEXT, last_used_at TEXT
);
CREATE INDEX api_keys_owner ON api_keys(user_id, created_at);
CREATE TABLE api_resources (
  id TEXT PRIMARY KEY, user_id TEXT NOT NULL, mode TEXT NOT NULL,
  kind TEXT NOT NULL, parent_id TEXT, data TEXT NOT NULL,
  created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE INDEX api_resources_owner ON api_resources(user_id, mode, kind, id);
CREATE UNIQUE INDEX api_one_label_per_shipment ON api_resources(parent_id) WHERE kind = 'label';
CREATE TABLE api_idempotency (
  user_id TEXT NOT NULL, mode TEXT NOT NULL, key TEXT NOT NULL,
  request_hash TEXT NOT NULL, status INTEGER NOT NULL, response TEXT NOT NULL,
  created_at TEXT NOT NULL, PRIMARY KEY(user_id, mode, key)
);
CREATE TABLE api_rate_limits (key_id TEXT NOT NULL, window INTEGER NOT NULL, hits INTEGER NOT NULL, PRIMARY KEY(key_id,window));
CREATE TABLE api_webhooks (
  id TEXT PRIMARY KEY, user_id TEXT NOT NULL, mode TEXT NOT NULL,
  url TEXT NOT NULL, secret TEXT NOT NULL, events TEXT NOT NULL,
  created_at TEXT NOT NULL, disabled_at TEXT
);
CREATE INDEX api_webhooks_owner ON api_webhooks(user_id, mode);
CREATE TABLE api_events (
  id TEXT PRIMARY KEY, user_id TEXT NOT NULL, mode TEXT NOT NULL,
  type TEXT NOT NULL, data TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE INDEX api_events_owner ON api_events(user_id, mode, id);
CREATE TABLE api_deliveries (
  id TEXT PRIMARY KEY, event_id TEXT NOT NULL REFERENCES api_events(id),
  webhook_id TEXT NOT NULL REFERENCES api_webhooks(id), attempts INTEGER NOT NULL DEFAULT 0,
  next_attempt_at INTEGER NOT NULL, lease_until INTEGER NOT NULL DEFAULT 0,
  delivered_at TEXT, last_status INTEGER, failed_at TEXT,
  UNIQUE(event_id,webhook_id)
);
CREATE INDEX api_deliveries_pending ON api_deliveries(next_attempt_at, lease_until) WHERE delivered_at IS NULL AND failed_at IS NULL;
