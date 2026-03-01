create table if not exists public.admin_settings (
  scope text primary key,
  carrier_discount_pct numeric(5,2) not null default 25,
  client_discount_pct numeric(5,2) not null default 20,
  updated_at timestamptz not null default now(),
  updated_by uuid references auth.users(id) on delete set null
);

insert into public.admin_settings (scope, carrier_discount_pct, client_discount_pct)
values ('global', 25, 20)
on conflict (scope) do nothing;

alter table public.admin_settings enable row level security;

alter table public.registration_invites
  add column if not exists token_encrypted text,
  add column if not exists revoked_at timestamptz,
  add column if not exists revoked_by uuid references auth.users(id) on delete set null;

create index if not exists registration_invites_revoked_at_idx
  on public.registration_invites (revoked_at desc);
