create extension if not exists pgcrypto;

create table if not exists public.billing_invoices (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  period_start date not null,
  period_end date not null,
  due_at timestamptz not null,
  issued_at timestamptz,
  sent_at timestamptz,
  paid_at timestamptz,
  status text not null default 'draft' check (status in ('draft', 'sent', 'overdue', 'paid', 'cancelled')),
  payment_mode text not null default 'invoice' check (payment_mode in ('invoice', 'card', 'hybrid')),
  currency text not null default 'EUR',
  vat_rate numeric(5,4) not null default 0.21,
  subtotal_ex_vat numeric(12,2) not null default 0,
  vat_amount numeric(12,2) not null default 0,
  total_inc_vat numeric(12,2) not null default 0,
  labels_count integer not null default 0,
  line_count integer not null default 0,
  reminder_stage integer not null default 0,
  last_reminder_sent_at timestamptz,
  email_message_id text,
  company_name text,
  contact_name text,
  contact_email text,
  customer_id text,
  account_manager text,
  payment_reference text,
  payment_received_amount numeric(12,2),
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint billing_invoices_unique_period unique (user_id, period_start, period_end)
);

create index if not exists billing_invoices_user_idx
  on public.billing_invoices (user_id, period_start desc);

create index if not exists billing_invoices_status_due_idx
  on public.billing_invoices (status, due_at asc);

create index if not exists billing_invoices_updated_idx
  on public.billing_invoices (updated_at desc);

create table if not exists public.billing_invoice_items (
  id bigint generated always as identity primary key,
  invoice_id uuid not null references public.billing_invoices(id) on delete cascade,
  generation_id uuid,
  generated_at timestamptz,
  service_type text,
  quantity integer not null default 1,
  recipient_name text,
  recipient_city text,
  recipient_country text,
  tracking_id text,
  label_id text,
  unit_price_ex_vat numeric(12,2) not null default 0,
  amount_ex_vat numeric(12,2) not null default 0,
  vat_amount numeric(12,2) not null default 0,
  amount_inc_vat numeric(12,2) not null default 0,
  sort_index integer not null default 0,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists billing_invoice_items_invoice_idx
  on public.billing_invoice_items (invoice_id, sort_index asc, id asc);

create table if not exists public.billing_invoice_events (
  id bigint generated always as identity primary key,
  invoice_id uuid not null references public.billing_invoices(id) on delete cascade,
  event_type text not null check (
    event_type in (
      'created',
      'updated',
      'sent',
      'reminder',
      'paid',
      'cancelled',
      'failed'
    )
  ),
  event_at timestamptz not null default now(),
  actor text not null default 'system',
  channel text,
  message text,
  metadata jsonb not null default '{}'::jsonb
);

create index if not exists billing_invoice_events_invoice_idx
  on public.billing_invoice_events (invoice_id, event_at desc);

alter table public.label_generations
  add column if not exists billed_invoice_id uuid references public.billing_invoices(id) on delete set null;

alter table public.label_generations
  add column if not exists billed_at timestamptz;

create index if not exists label_generations_unbilled_idx
  on public.label_generations (user_id, created_at desc)
  where billed_invoice_id is null;

alter table public.billing_invoices enable row level security;
alter table public.billing_invoice_items enable row level security;
alter table public.billing_invoice_events enable row level security;

drop policy if exists billing_invoices_select_own on public.billing_invoices;
create policy billing_invoices_select_own
  on public.billing_invoices
  for select
  to authenticated
  using (auth.uid() = user_id);

drop policy if exists billing_invoice_items_select_own on public.billing_invoice_items;
create policy billing_invoice_items_select_own
  on public.billing_invoice_items
  for select
  to authenticated
  using (
    exists (
      select 1
      from public.billing_invoices inv
      where inv.id = billing_invoice_items.invoice_id
        and inv.user_id = auth.uid()
    )
  );

drop policy if exists billing_invoice_events_select_own on public.billing_invoice_events;
create policy billing_invoice_events_select_own
  on public.billing_invoice_events
  for select
  to authenticated
  using (
    exists (
      select 1
      from public.billing_invoices inv
      where inv.id = billing_invoice_events.invoice_id
        and inv.user_id = auth.uid()
    )
  );
