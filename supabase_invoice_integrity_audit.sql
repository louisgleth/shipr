-- Invoice integrity audit queries for billing_invoices / billing_wallet_topups.
-- Run these in the Supabase SQL editor to verify canonical invoice mapping.

-- 1. Duplicate top-up invoices for the same credited event.
select
  source_topup_id,
  count(*) as invoice_count,
  array_agg(id order by created_at asc, id asc) as invoice_ids,
  array_agg(invoice_number order by created_at asc, id asc) as invoice_numbers
from public.billing_invoices
where invoice_kind = 'topup'
  and source_topup_id is not null
group by source_topup_id
having count(*) > 1
order by invoice_count desc, source_topup_id;

-- 2. Duplicate monthly invoices for the same client + period.
select
  user_id,
  period_start,
  period_end,
  count(*) as invoice_count,
  array_agg(id order by created_at asc, id asc) as invoice_ids,
  array_agg(coalesce(invoice_number, '')) as invoice_numbers
from public.billing_invoices
where invoice_kind = 'monthly'
group by user_id, period_start, period_end
having count(*) > 1
order by invoice_count desc, period_start desc, user_id;

-- 3. Duplicate invoice numbers across the whole ledger.
select
  invoice_number,
  count(*) as invoice_count,
  array_agg(id order by created_at asc, id asc) as invoice_ids,
  array_agg(invoice_kind order by created_at asc, id asc) as invoice_kinds
from public.billing_invoices
where coalesce(trim(invoice_number), '') <> ''
group by invoice_number
having count(*) > 1
order by invoice_count desc, invoice_number;

-- 4. Credited top-ups missing a canonical invoice row.
select
  t.id as topup_id,
  t.user_id,
  t.reference as topup_reference,
  t.credited_at,
  t.metadata ->> 'invoice_id' as metadata_invoice_id,
  t.metadata ->> 'invoice_number' as metadata_invoice_number
from public.billing_wallet_topups t
left join public.billing_invoices i
  on i.source_topup_id = t.id
 and i.invoice_kind = 'topup'
where coalesce(t.status, '') = 'credited'
  and i.id is null
order by t.credited_at desc nulls last, t.created_at desc nulls last;

-- 5. Credited top-ups whose metadata points to a different invoice than the canonical linked invoice.
select
  t.id as topup_id,
  t.reference as topup_reference,
  t.metadata ->> 'invoice_id' as metadata_invoice_id,
  i.id as canonical_invoice_id,
  t.metadata ->> 'invoice_number' as metadata_invoice_number,
  i.invoice_number as canonical_invoice_number
from public.billing_wallet_topups t
join public.billing_invoices i
  on i.source_topup_id = t.id
 and i.invoice_kind = 'topup'
where coalesce(t.status, '') = 'credited'
  and (
    coalesce(trim(t.metadata ->> 'invoice_id'), '') <> coalesce(trim(i.id::text), '')
    or coalesce(trim(t.metadata ->> 'invoice_number'), '') <> coalesce(trim(i.invoice_number), '')
  )
order by t.credited_at desc nulls last, t.created_at desc nulls last;
