create table if not exists public.provider_connections (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  provider text not null,
  shop_domain text not null,
  status text not null default 'connected',
  scopes text not null default '',
  access_token text not null,
  import_settings jsonb not null default '{}'::jsonb,
  connected_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (user_id, provider, shop_domain)
);

alter table public.provider_connections
  add column if not exists import_settings jsonb not null default '{}'::jsonb;

create index if not exists provider_connections_user_provider_idx
  on public.provider_connections (user_id, provider, updated_at desc);

alter table public.provider_connections enable row level security;

drop policy if exists "provider_connections_insert_own" on public.provider_connections;

drop policy if exists "provider_connections_update_own" on public.provider_connections;

alter table public.provider_connections
  drop constraint if exists provider_connections_access_token_encrypted_check;

alter table public.provider_connections
  add constraint provider_connections_access_token_encrypted_check
  check (access_token like 'v1:%') not valid;

revoke insert, update, delete on public.provider_connections from anon, authenticated;
grant select, insert, update, delete on public.provider_connections to service_role;
