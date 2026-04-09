-- Preserve legally required financial records when a portal account is deleted.
-- Run this in the Supabase SQL editor for existing environments.

alter table public.billing_invoices
  alter column user_id drop not null;

alter table public.billing_invoices
  drop constraint if exists billing_invoices_user_id_fkey;

alter table public.billing_invoices
  add constraint billing_invoices_user_id_fkey
  foreign key (user_id) references auth.users(id) on delete set null;

alter table public.billing_wallet_topups
  alter column user_id drop not null;

alter table public.billing_wallet_topups
  drop constraint if exists billing_wallet_topups_user_id_fkey;

alter table public.billing_wallet_topups
  add constraint billing_wallet_topups_user_id_fkey
  foreign key (user_id) references auth.users(id) on delete set null;

alter table public.billing_wallet_transactions
  alter column user_id drop not null;

alter table public.billing_wallet_transactions
  drop constraint if exists billing_wallet_transactions_user_id_fkey;

alter table public.billing_wallet_transactions
  add constraint billing_wallet_transactions_user_id_fkey
  foreign key (user_id) references auth.users(id) on delete set null;
