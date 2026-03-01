create extension if not exists pgcrypto;

create table if not exists public.registration_invites (
  id uuid primary key default gen_random_uuid(),
  token_hash text not null unique,
  invited_email text,
  company_name text,
  contact_name text,
  contact_email text,
  contact_phone text,
  billing_address text,
  tax_id text,
  customer_id text,
  plan_name text not null default 'Growth',
  expires_at timestamptz not null,
  created_at timestamptz not null default now(),
  claimed_at timestamptz,
  claimed_user_id uuid references auth.users(id) on delete set null,
  claimed_email text,
  created_by uuid references auth.users(id) on delete set null,
  notes text
);

create index if not exists registration_invites_expires_at_idx
  on public.registration_invites (expires_at desc);

create index if not exists registration_invites_invited_email_idx
  on public.registration_invites (lower(invited_email));

alter table public.registration_invites enable row level security;
