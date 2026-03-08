create extension if not exists pgcrypto;

create table if not exists public.billing_wallets (
  user_id uuid primary key references auth.users(id) on delete cascade,
  balance_cents bigint not null default 0 check (balance_cents >= 0),
  currency text not null default 'EUR',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.billing_wallet_topups (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  amount_cents bigint not null check (amount_cents > 0),
  currency text not null default 'EUR',
  status text not null default 'pending' check (status in ('pending', 'received', 'credited', 'cancelled', 'failed')),
  reference text not null unique,
  requested_at timestamptz not null default now(),
  received_at timestamptz,
  credited_at timestamptz,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.billing_wallet_transactions (
  id bigint generated always as identity primary key,
  user_id uuid not null references auth.users(id) on delete cascade,
  source text not null,
  amount_cents bigint not null,
  balance_after_cents bigint not null check (balance_after_cents >= 0),
  reference text,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists billing_wallet_topups_user_requested_idx
  on public.billing_wallet_topups (user_id, requested_at desc);

create index if not exists billing_wallet_topups_status_idx
  on public.billing_wallet_topups (status, requested_at desc);

create index if not exists billing_wallet_transactions_user_created_idx
  on public.billing_wallet_transactions (user_id, created_at desc, id desc);

alter table public.billing_wallets enable row level security;
alter table public.billing_wallet_topups enable row level security;
alter table public.billing_wallet_transactions enable row level security;

drop policy if exists billing_wallets_select_own on public.billing_wallets;
create policy billing_wallets_select_own
  on public.billing_wallets
  for select
  to authenticated
  using (auth.uid() = user_id);

drop policy if exists billing_wallet_topups_select_own on public.billing_wallet_topups;
create policy billing_wallet_topups_select_own
  on public.billing_wallet_topups
  for select
  to authenticated
  using (auth.uid() = user_id);

drop policy if exists billing_wallet_topups_insert_own on public.billing_wallet_topups;
create policy billing_wallet_topups_insert_own
  on public.billing_wallet_topups
  for insert
  to authenticated
  with check (auth.uid() = user_id);

drop policy if exists billing_wallet_transactions_select_own on public.billing_wallet_transactions;
create policy billing_wallet_transactions_select_own
  on public.billing_wallet_transactions
  for select
  to authenticated
  using (auth.uid() = user_id);
