create extension if not exists pgcrypto;

create table if not exists public.wise_webhook_events (
  id uuid primary key default gen_random_uuid(),
  delivery_id text not null unique,
  event_type text not null,
  subscription_id text,
  schema_version text,
  signature_valid boolean not null default false,
  is_test boolean not null default false,
  processing_status text not null default 'received' check (
    processing_status in ('received', 'processed', 'ignored', 'failed')
  ),
  processing_error text,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  processed_at timestamptz
);

create index if not exists wise_webhook_events_created_idx
  on public.wise_webhook_events (created_at desc);

create index if not exists wise_webhook_events_status_idx
  on public.wise_webhook_events (processing_status, created_at desc);

create table if not exists public.billing_bank_receipts (
  id uuid primary key default gen_random_uuid(),
  provider text not null default 'wise',
  source text not null default 'statement',
  external_key text not null unique,
  delivery_id text,
  event_type text,
  status text not null default 'pending' check (
    status in ('pending', 'manual_review', 'matched', 'credited', 'ignored', 'failed')
  ),
  amount_cents bigint not null default 0 check (amount_cents >= 0),
  currency text not null default 'EUR',
  payment_reference text,
  reference_number text,
  sender_name text,
  sender_account text,
  sender_bank text,
  occurred_at timestamptz,
  matched_user_id uuid references auth.users(id) on delete set null,
  matched_topup_id uuid references public.billing_wallet_topups(id) on delete set null,
  matched_invoice_id uuid references public.billing_invoices(id) on delete set null,
  match_reason text,
  review_note text,
  credited_at timestamptz,
  metadata jsonb not null default '{}'::jsonb,
  raw_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists billing_bank_receipts_status_idx
  on public.billing_bank_receipts (status, occurred_at desc, created_at desc);

create index if not exists billing_bank_receipts_reference_idx
  on public.billing_bank_receipts (payment_reference, occurred_at desc);

create index if not exists billing_bank_receipts_user_idx
  on public.billing_bank_receipts (matched_user_id, occurred_at desc);

create table if not exists public.billing_invoice_payments (
  id bigint generated always as identity primary key,
  invoice_id uuid not null references public.billing_invoices(id) on delete cascade,
  bank_receipt_id uuid not null references public.billing_bank_receipts(id) on delete cascade,
  amount_cents bigint not null check (amount_cents >= 0),
  currency text not null default 'EUR',
  payment_reference text,
  source text not null default 'wise_bank_receipt',
  received_at timestamptz,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  unique (bank_receipt_id)
);

create index if not exists billing_invoice_payments_invoice_idx
  on public.billing_invoice_payments (invoice_id, created_at desc);

alter table public.wise_webhook_events enable row level security;
alter table public.billing_bank_receipts enable row level security;
alter table public.billing_invoice_payments enable row level security;

create or replace function public.apply_billing_bank_receipt_resolution(
  p_receipt_id uuid,
  p_resolution text,
  p_actor text default 'system',
  p_topup_id uuid default null,
  p_invoice_id uuid default null,
  p_note text default null
)
returns public.billing_bank_receipts
language plpgsql
security definer
set search_path = public
as $$
declare
  v_resolution text := lower(trim(coalesce(p_resolution, '')));
  v_actor text := nullif(trim(coalesce(p_actor, '')), '');
  v_note text := nullif(trim(coalesce(p_note, '')), '');
  v_now timestamptz := now();
  v_receipt public.billing_bank_receipts%rowtype;
  v_topup public.billing_wallet_topups%rowtype;
  v_wallet public.billing_wallets%rowtype;
  v_invoice public.billing_invoices%rowtype;
  v_invoice_total_cents bigint := 0;
  v_paid_cents bigint := 0;
  v_balance_after_cents bigint := 0;
begin
  if p_receipt_id is null then
    raise exception 'Bank receipt id is required.';
  end if;
  if v_resolution not in ('topup', 'invoice', 'ignore') then
    raise exception 'Resolution must be topup, invoice, or ignore.';
  end if;

  select *
  into v_receipt
  from public.billing_bank_receipts
  where id = p_receipt_id
  for update;

  if not found then
    raise exception 'Bank receipt not found.';
  end if;

  if v_resolution = 'ignore' then
    update public.billing_bank_receipts
    set
      status = 'ignored',
      review_note = coalesce(v_note, review_note),
      match_reason = 'ignored_by_admin',
      updated_at = v_now,
      metadata = coalesce(metadata, '{}'::jsonb) || jsonb_strip_nulls(
        jsonb_build_object(
          'last_actor', coalesce(v_actor, 'system'),
          'resolution_note', v_note
        )
      )
    where id = p_receipt_id
    returning * into v_receipt;
    return v_receipt;
  end if;

  if v_receipt.matched_topup_id is not null and v_resolution <> 'topup' then
    raise exception 'Receipt is already linked to a wallet top-up.';
  end if;
  if v_receipt.matched_invoice_id is not null and v_resolution <> 'invoice' then
    raise exception 'Receipt is already linked to an invoice payment.';
  end if;

  if v_resolution = 'topup' then
    if p_topup_id is null then
      raise exception 'Top-up id is required.';
    end if;
    if v_receipt.status = 'credited' and v_receipt.matched_topup_id = p_topup_id then
      return v_receipt;
    end if;

    select *
    into v_topup
    from public.billing_wallet_topups
    where id = p_topup_id
    for update;

    if not found then
      raise exception 'Wallet top-up not found.';
    end if;
    if v_topup.user_id is null then
      raise exception 'Wallet top-up has no account owner.';
    end if;
    if nullif(coalesce(v_topup.metadata ->> 'wise_receipt_id', ''), '') is not null
      and v_topup.metadata ->> 'wise_receipt_id' <> v_receipt.id::text then
      raise exception 'Wallet top-up was already credited by another bank receipt.';
    end if;

    insert into public.billing_wallets (user_id, balance_cents, currency, created_at, updated_at)
    values (v_topup.user_id, 0, coalesce(nullif(v_receipt.currency, ''), 'EUR'), v_now, v_now)
    on conflict (user_id) do nothing;

    select *
    into v_wallet
    from public.billing_wallets
    where user_id = v_topup.user_id
    for update;

    v_balance_after_cents := coalesce(v_wallet.balance_cents, 0) + coalesce(v_receipt.amount_cents, 0);

    update public.billing_wallets
    set
      balance_cents = v_balance_after_cents,
      currency = coalesce(nullif(v_wallet.currency, ''), nullif(v_receipt.currency, ''), 'EUR'),
      updated_at = v_now
    where user_id = v_topup.user_id;

    insert into public.billing_wallet_transactions (
      user_id,
      source,
      amount_cents,
      balance_after_cents,
      reference,
      metadata
    )
    values (
      v_topup.user_id,
      'wise_bank_receipt',
      coalesce(v_receipt.amount_cents, 0),
      v_balance_after_cents,
      coalesce(
        nullif(v_receipt.payment_reference, ''),
        nullif(v_topup.reference, ''),
        concat('WISE-', upper(left(replace(v_receipt.id::text, '-', ''), 12)))
      ),
      jsonb_strip_nulls(
        jsonb_build_object(
          'receipt_id', v_receipt.id,
          'external_key', v_receipt.external_key,
          'delivery_id', v_receipt.delivery_id,
          'actor', coalesce(v_actor, 'system'),
          'note', v_note
        )
      )
    );

    update public.billing_wallet_topups
    set
      status = 'credited',
      amount_cents = coalesce(v_receipt.amount_cents, amount_cents, 0),
      received_at = coalesce(received_at, v_receipt.occurred_at, v_now),
      credited_at = coalesce(credited_at, v_receipt.occurred_at, v_now),
      updated_at = v_now,
      metadata = coalesce(metadata, '{}'::jsonb) || jsonb_strip_nulls(
        jsonb_build_object(
          'wise_receipt_id', v_receipt.id,
          'wise_external_key', v_receipt.external_key,
          'requested_amount_cents', v_topup.amount_cents,
          'credited_amount_cents', v_receipt.amount_cents,
          'last_actor', coalesce(v_actor, 'system'),
          'resolution_note', v_note
        )
      )
    where id = v_topup.id
    returning * into v_topup;

    update public.billing_bank_receipts
    set
      status = 'credited',
      matched_user_id = v_topup.user_id,
      matched_topup_id = v_topup.id,
      matched_invoice_id = null,
      credited_at = coalesce(credited_at, v_receipt.occurred_at, v_now),
      review_note = coalesce(v_note, review_note),
      match_reason = coalesce(match_reason, 'topup_reference_exact'),
      updated_at = v_now,
      metadata = coalesce(metadata, '{}'::jsonb) || jsonb_strip_nulls(
        jsonb_build_object(
          'last_actor', coalesce(v_actor, 'system'),
          'resolution_note', v_note,
          'wallet_balance_after_cents', v_balance_after_cents
        )
      )
    where id = p_receipt_id
    returning * into v_receipt;

    return v_receipt;
  end if;

  if p_invoice_id is null then
    raise exception 'Invoice id is required.';
  end if;
  if v_receipt.status = 'matched' and v_receipt.matched_invoice_id = p_invoice_id then
    return v_receipt;
  end if;

  select *
  into v_invoice
  from public.billing_invoices
  where id = p_invoice_id
  for update;

  if not found then
    raise exception 'Invoice not found.';
  end if;
  if nullif(coalesce(v_receipt.matched_topup_id::text, ''), '') is not null then
    raise exception 'Receipt is already linked to a wallet top-up.';
  end if;

  insert into public.billing_invoice_payments (
    invoice_id,
    bank_receipt_id,
    amount_cents,
    currency,
    payment_reference,
    source,
    received_at,
    metadata
  )
  values (
    v_invoice.id,
    v_receipt.id,
    coalesce(v_receipt.amount_cents, 0),
    coalesce(nullif(v_receipt.currency, ''), 'EUR'),
    nullif(v_receipt.payment_reference, ''),
    'wise_bank_receipt',
    coalesce(v_receipt.occurred_at, v_now),
    jsonb_strip_nulls(
      jsonb_build_object(
        'external_key', v_receipt.external_key,
        'delivery_id', v_receipt.delivery_id,
        'actor', coalesce(v_actor, 'system'),
        'note', v_note
      )
    )
  )
  on conflict (bank_receipt_id) do update
  set
    invoice_id = excluded.invoice_id,
    amount_cents = excluded.amount_cents,
    currency = excluded.currency,
    payment_reference = excluded.payment_reference,
    received_at = excluded.received_at,
    metadata = public.billing_invoice_payments.metadata || excluded.metadata;

  select coalesce(sum(amount_cents), 0)
  into v_paid_cents
  from public.billing_invoice_payments
  where invoice_id = v_invoice.id;

  v_invoice_total_cents := round(coalesce(v_invoice.total_inc_vat, 0)::numeric * 100);

  update public.billing_invoices
  set
    payment_reference = coalesce(nullif(payment_reference, ''), nullif(v_receipt.payment_reference, '')),
    payment_received_amount = round(v_paid_cents::numeric / 100, 2),
    status = case
      when v_paid_cents >= v_invoice_total_cents then 'paid'
      else status
    end,
    paid_at = case
      when v_paid_cents >= v_invoice_total_cents then coalesce(paid_at, v_receipt.occurred_at, v_now)
      else paid_at
    end,
    updated_at = v_now
  where id = v_invoice.id
  returning * into v_invoice;

  insert into public.billing_invoice_events (
    invoice_id,
    event_type,
    event_at,
    actor,
    channel,
    message,
    metadata
  )
  values (
    v_invoice.id,
    case when v_paid_cents >= v_invoice_total_cents then 'paid' else 'updated' end,
    v_now,
    coalesce(v_actor, 'system'),
    'wise_bank_receipt',
    case
      when v_paid_cents >= v_invoice_total_cents then 'Invoice settled from Wise receipt.'
      else 'Partial invoice payment recorded from Wise receipt.'
    end,
    jsonb_strip_nulls(
      jsonb_build_object(
        'receipt_id', v_receipt.id,
        'external_key', v_receipt.external_key,
        'payment_reference', v_receipt.payment_reference,
        'amount_cents', v_receipt.amount_cents,
        'total_paid_cents', v_paid_cents
      )
    )
  );

  update public.billing_bank_receipts
  set
    status = 'matched',
    matched_user_id = v_invoice.user_id,
    matched_topup_id = null,
    matched_invoice_id = v_invoice.id,
    review_note = coalesce(v_note, review_note),
    match_reason = coalesce(
      match_reason,
      case
        when v_paid_cents >= v_invoice_total_cents then 'invoice_reference_exact'
        else 'invoice_reference_partial'
      end
    ),
    updated_at = v_now,
    metadata = coalesce(metadata, '{}'::jsonb) || jsonb_strip_nulls(
      jsonb_build_object(
        'last_actor', coalesce(v_actor, 'system'),
        'resolution_note', v_note,
        'invoice_total_paid_cents', v_paid_cents
      )
    )
  where id = p_receipt_id
  returning * into v_receipt;

  return v_receipt;
end;
$$;
