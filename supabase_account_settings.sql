create table if not exists public.account_settings (
  user_id uuid primary key references auth.users(id) on delete cascade,
  warehouse_origins jsonb not null default '[]'::jsonb,
  updated_at timestamptz not null default now()
);

alter table public.account_settings enable row level security;

drop policy if exists "account_settings_select_own" on public.account_settings;
create policy "account_settings_select_own"
  on public.account_settings
  for select
  using (auth.uid() = user_id);

drop policy if exists "account_settings_insert_own" on public.account_settings;
create policy "account_settings_insert_own"
  on public.account_settings
  for insert
  with check (auth.uid() = user_id);

drop policy if exists "account_settings_update_own" on public.account_settings;
create policy "account_settings_update_own"
  on public.account_settings
  for update
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);
