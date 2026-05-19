create table if not exists public.lead_prospects (
  id uuid primary key default gen_random_uuid(),
  dedupe_key text not null,
  source text not null default 'csv',
  order_index integer not null default 0,
  payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists lead_prospects_dedupe_key_idx
  on public.lead_prospects (dedupe_key);

create index if not exists lead_prospects_order_index_idx
  on public.lead_prospects (order_index asc, created_at asc);

create index if not exists lead_prospects_updated_at_idx
  on public.lead_prospects (updated_at desc);

create index if not exists lead_prospects_payload_disposition_idx
  on public.lead_prospects ((payload ->> 'disposition'));

alter table public.lead_prospects enable row level security;

revoke all on public.lead_prospects from public;
revoke all on public.lead_prospects from anon;
revoke all on public.lead_prospects from authenticated;

grant select, insert, update, delete on public.lead_prospects to service_role;

drop policy if exists "lead_prospects_service_role_all" on public.lead_prospects;

create policy "lead_prospects_service_role_all"
  on public.lead_prospects
  for all
  to service_role
  using (true)
  with check (true);
