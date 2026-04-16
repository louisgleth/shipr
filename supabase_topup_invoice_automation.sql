alter table public.billing_invoices
  add column if not exists invoice_kind text;

update public.billing_invoices
set invoice_kind = 'monthly'
where coalesce(trim(invoice_kind), '') = '';

alter table public.billing_invoices
  alter column invoice_kind set default 'monthly';

alter table public.billing_invoices
  alter column invoice_kind set not null;

alter table public.billing_invoices
  drop constraint if exists billing_invoices_invoice_kind_check;

alter table public.billing_invoices
  add constraint billing_invoices_invoice_kind_check
  check (invoice_kind in ('monthly', 'topup'));

alter table public.billing_invoices
  add column if not exists source_topup_id uuid;

alter table public.billing_invoices
  add column if not exists invoice_number text;

alter table public.billing_invoices
  drop constraint if exists billing_invoices_payment_mode_check;

alter table public.billing_invoices
  add constraint billing_invoices_payment_mode_check
  check (payment_mode in ('invoice', 'card', 'hybrid', 'wallet'));

alter table public.billing_invoices
  drop constraint if exists billing_invoices_unique_period;

create unique index if not exists billing_invoices_unique_monthly_period_idx
  on public.billing_invoices (user_id, period_start, period_end)
  where invoice_kind = 'monthly';

create unique index if not exists billing_invoices_unique_topup_idx
  on public.billing_invoices (source_topup_id)
  where invoice_kind = 'topup' and source_topup_id is not null;

create index if not exists billing_invoices_kind_user_idx
  on public.billing_invoices (invoice_kind, user_id, updated_at desc);

notify pgrst, 'reload schema';
