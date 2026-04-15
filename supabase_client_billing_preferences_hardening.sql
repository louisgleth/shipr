alter table public.client_billing_preferences enable row level security;

drop policy if exists "client_billing_insert_own" on public.client_billing_preferences;
drop policy if exists "client_billing_update_own" on public.client_billing_preferences;

-- Billing eligibility is managed by trusted backend/admin flows only.
-- Authenticated clients may keep read access to their own settings, but all
-- writes must go through server-side admin paths.
