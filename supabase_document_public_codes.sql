create unique index if not exists billing_invoices_unique_invoice_number_idx
  on public.billing_invoices (invoice_number)
  where coalesce(trim(invoice_number), '') <> '';

create unique index if not exists label_generations_unique_receipt_public_code_idx
  on public.label_generations (((payload -> 'document' ->> 'public_code')))
  where coalesce(trim(payload -> 'document' ->> 'public_code'), '') <> '';
