create extension if not exists pgcrypto;

create table if not exists public.shipment_extract_requests (
  id uuid primary key default gen_random_uuid(),
  token_hash text not null unique,
  token_encrypted text,
  client_email text not null,
  expires_at timestamptz not null,
  created_at timestamptz not null default now(),
  created_by uuid references auth.users(id) on delete set null,
  submitted_at timestamptz,
  submitted_filename text,
  submitted_rows integer,
  submitted_columns integer,
  kept_columns text[],
  removed_columns text[],
  registration_invite_id uuid references public.registration_invites(id) on delete set null,
  metadata jsonb not null default '{}'::jsonb
);

create index if not exists shipment_extract_requests_client_email_idx
  on public.shipment_extract_requests (lower(client_email));

create index if not exists shipment_extract_requests_created_at_idx
  on public.shipment_extract_requests (created_at desc);

create index if not exists shipment_extract_requests_submitted_at_idx
  on public.shipment_extract_requests (submitted_at desc);

alter table public.shipment_extract_requests enable row level security;
