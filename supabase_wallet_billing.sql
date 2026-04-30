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
  user_id uuid references auth.users(id) on delete set null,
  amount_cents bigint not null default 0 check (amount_cents >= 0),
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

alter table public.billing_wallet_topups
  alter column amount_cents set default 0;

alter table public.billing_wallet_topups
  drop constraint if exists billing_wallet_topups_amount_cents_check;

alter table public.billing_wallet_topups
  add constraint billing_wallet_topups_amount_cents_check
  check (amount_cents >= 0);

create table if not exists public.billing_wallet_transactions (
  id bigint generated always as identity primary key,
  user_id uuid references auth.users(id) on delete set null,
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

drop policy if exists billing_wallet_transactions_select_own on public.billing_wallet_transactions;
create policy billing_wallet_transactions_select_own
  on public.billing_wallet_transactions
  for select
  to authenticated
  using (auth.uid() = user_id);

revoke insert, update, delete on public.billing_wallets from anon, authenticated;
revoke insert, update, delete on public.billing_wallet_topups from anon, authenticated;
revoke insert, update, delete on public.billing_wallet_transactions from anon, authenticated;

grant select on public.billing_wallets to authenticated;
grant select on public.billing_wallet_topups to authenticated;
grant select on public.billing_wallet_transactions to authenticated;

grant select, insert, update, delete on public.billing_wallets to service_role;
grant select, insert, update, delete on public.billing_wallet_topups to service_role;
grant select, insert, update, delete on public.billing_wallet_transactions to service_role;

drop function if exists public.apply_billing_wallet_transaction(uuid, bigint, text, text, jsonb, text);

create function public.apply_billing_wallet_transaction(
  p_user_id uuid,
  p_amount_cents bigint,
  p_source text,
  p_reference text default null,
  p_metadata jsonb default '{}'::jsonb,
  p_currency text default 'EUR'
)
returns table (
  wallet_user_id uuid,
  wallet_balance_cents bigint,
  wallet_currency text,
  wallet_updated_at timestamptz,
  transaction_id bigint,
  transaction_reference text,
  balance_after_cents bigint
)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_wallet public.billing_wallets%rowtype;
  v_next_balance bigint;
  v_now timestamptz := now();
  v_source text := nullif(trim(coalesce(p_source, '')), '');
  v_reference text := nullif(trim(coalesce(p_reference, '')), '');
  v_currency text := coalesce(nullif(trim(coalesce(p_currency, '')), ''), 'EUR');
  v_metadata jsonb := coalesce(p_metadata, '{}'::jsonb);
begin
  if p_user_id is null then
    raise exception 'Wallet user id is required.' using errcode = '22023';
  end if;

  if p_amount_cents is null or p_amount_cents = 0 then
    raise exception 'Wallet transaction amount must be non-zero.' using errcode = '22023';
  end if;

  if v_source is null then
    raise exception 'Wallet transaction source is required.' using errcode = '22023';
  end if;

  if jsonb_typeof(v_metadata) <> 'object' then
    raise exception 'Wallet transaction metadata must be a JSON object.' using errcode = '22023';
  end if;

  insert into public.billing_wallets (user_id, balance_cents, currency, created_at, updated_at)
  values (p_user_id, 0, v_currency, v_now, v_now)
  on conflict (user_id) do nothing;

  select *
    into v_wallet
    from public.billing_wallets
   where user_id = p_user_id
   for update;

  if not found then
    raise exception 'Billing wallet could not be created.' using errcode = 'P0001';
  end if;

  v_next_balance := coalesce(v_wallet.balance_cents, 0) + p_amount_cents;

  if v_next_balance < 0 then
    raise exception 'Insufficient wallet balance.' using errcode = 'P0001';
  end if;

  update public.billing_wallets
     set balance_cents = v_next_balance,
         currency = coalesce(nullif(v_wallet.currency, ''), v_currency),
         updated_at = v_now
   where user_id = p_user_id
   returning *
    into v_wallet;

  insert into public.billing_wallet_transactions (
    user_id,
    source,
    amount_cents,
    balance_after_cents,
    reference,
    metadata,
    created_at
  )
  values (
    p_user_id,
    v_source,
    p_amount_cents,
    v_next_balance,
    v_reference,
    v_metadata,
    v_now
  )
  returning id, reference
    into transaction_id, transaction_reference;

  wallet_user_id := v_wallet.user_id;
  wallet_balance_cents := v_wallet.balance_cents;
  wallet_currency := v_wallet.currency;
  wallet_updated_at := v_wallet.updated_at;
  balance_after_cents := v_next_balance;

  return next;
end;
$$;

revoke all on function public.apply_billing_wallet_transaction(uuid, bigint, text, text, jsonb, text) from public;
revoke all on function public.apply_billing_wallet_transaction(uuid, bigint, text, text, jsonb, text) from anon;
revoke all on function public.apply_billing_wallet_transaction(uuid, bigint, text, text, jsonb, text) from authenticated;
grant execute on function public.apply_billing_wallet_transaction(uuid, bigint, text, text, jsonb, text) to service_role;
