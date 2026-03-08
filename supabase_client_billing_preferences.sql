create table if not exists public.client_billing_preferences (
  user_id uuid primary key references auth.users(id) on delete cascade,
  invoice_enabled boolean not null default false,
  card_enabled boolean not null default true,
  updated_at timestamptz not null default now(),
  updated_by uuid references auth.users(id),
  constraint client_billing_preferences_at_least_one_method
    check (invoice_enabled or card_enabled)
);

alter table public.client_billing_preferences
  alter column invoice_enabled set default false;

alter table public.client_billing_preferences
  alter column card_enabled set default true;

create index if not exists client_billing_preferences_updated_at_idx
  on public.client_billing_preferences (updated_at desc);

alter table public.client_billing_preferences enable row level security;

drop policy if exists "client_billing_select_own" on public.client_billing_preferences;
create policy "client_billing_select_own"
  on public.client_billing_preferences
  for select
  to authenticated
  using (auth.uid() = user_id);

drop policy if exists "client_billing_insert_own" on public.client_billing_preferences;
create policy "client_billing_insert_own"
  on public.client_billing_preferences
  for insert
  to authenticated
  with check (auth.uid() = user_id);

drop policy if exists "client_billing_update_own" on public.client_billing_preferences;
create policy "client_billing_update_own"
  on public.client_billing_preferences
  for update
  to authenticated
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);
