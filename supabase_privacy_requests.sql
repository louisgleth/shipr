
create extension if not exists pgcrypto;

create table if not exists public.privacy_requests (
  id uuid primary key default gen_random_uuid(),
  source text not null default 'portal',
  request_type text not null,
  status text not null default 'received',
  subject_user_id uuid,
  requested_by_user_id uuid references auth.users(id) on delete set null,
  provider text,
  shop_domain text,
  external_customer_id text,
  external_customer_email text,
  request_payload jsonb not null default '{}'::jsonb,
  result_payload jsonb not null default '{}'::jsonb,
  requested_at timestamptz not null default now(),
  processed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint privacy_requests_source_check
    check (source in ('portal', 'shopify_webhook', 'admin', 'system')),
  constraint privacy_requests_type_check
    check (
      request_type in (
        'export',
        'delete',
        'shop_redact',
        'customer_data_request',
        'customer_redact'
      )
    ),
  constraint privacy_requests_status_check
    check (
      status in (
        'received',
        'processing',
        'completed',
        'manual_review_required',
        'failed'
      )
    )
);

create index if not exists privacy_requests_requested_at_idx
  on public.privacy_requests (requested_at desc, created_at desc);

create index if not exists privacy_requests_requested_by_idx
  on public.privacy_requests (requested_by_user_id, requested_at desc);

create index if not exists privacy_requests_subject_user_idx
  on public.privacy_requests (subject_user_id, requested_at desc);

create index if not exists privacy_requests_provider_shop_idx
  on public.privacy_requests (provider, shop_domain, requested_at desc);

alter table public.privacy_requests enable row level security;

drop policy if exists privacy_requests_select_own on public.privacy_requests;
create policy privacy_requests_select_own
  on public.privacy_requests
  for select
  to authenticated
  using (
    auth.uid() = requested_by_user_id
    or auth.uid()::text = coalesce(subject_user_id::text, '')
  );

drop policy if exists privacy_requests_insert_own on public.privacy_requests;
create policy privacy_requests_insert_own
  on public.privacy_requests
  for insert
  to authenticated
  with check (
    requested_by_user_id is null
    or auth.uid() = requested_by_user_id
  );
