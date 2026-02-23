create table if not exists public.provider_connections (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  provider text not null,
  shop_domain text not null,
  status text not null default 'connected',
  scopes text not null default '',
  access_token text not null,
  connected_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (user_id, provider, shop_domain)
);

create index if not exists provider_connections_user_provider_idx
  on public.provider_connections (user_id, provider, updated_at desc);

alter table public.provider_connections enable row level security;

drop policy if exists "provider_connections_insert_own" on public.provider_connections;
create policy "provider_connections_insert_own"
  on public.provider_connections
  for insert
  with check (auth.uid() = user_id);

drop policy if exists "provider_connections_update_own" on public.provider_connections;
create policy "provider_connections_update_own"
  on public.provider_connections
  for update
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

