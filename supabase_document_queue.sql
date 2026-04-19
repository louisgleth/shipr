create table if not exists public.billing_document_jobs (
  id uuid primary key default gen_random_uuid(),
  job_type text not null,
  job_key text not null unique,
  invoice_id uuid references public.billing_invoices(id) on delete cascade,
  user_id uuid references auth.users(id) on delete set null,
  status text not null default 'queued',
  priority integer not null default 100,
  attempts integer not null default 0,
  max_attempts integer not null default 4,
  payload jsonb not null default '{}'::jsonb,
  result jsonb not null default '{}'::jsonb,
  error_message text,
  available_at timestamptz not null default timezone('utc', now()),
  claimed_at timestamptz,
  claimed_by text,
  lease_expires_at timestamptz,
  completed_at timestamptz,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  constraint billing_document_jobs_job_type_check
    check (job_type in ('invoice_pdf', 'invoice_email', 'invoice_render', 'receipt_render', 'preview_email')),
  constraint billing_document_jobs_status_check
    check (status in ('queued', 'processing', 'completed', 'failed'))
);

create index if not exists billing_document_jobs_status_priority_idx
  on public.billing_document_jobs (status, priority asc, available_at asc, created_at asc);

create index if not exists billing_document_jobs_invoice_idx
  on public.billing_document_jobs (invoice_id, created_at desc);

create index if not exists billing_document_jobs_user_idx
  on public.billing_document_jobs (user_id, created_at desc);

alter table public.billing_document_jobs enable row level security;

revoke all on public.billing_document_jobs from public;
revoke all on public.billing_document_jobs from anon;
revoke all on public.billing_document_jobs from authenticated;

grant select, insert, update, delete on public.billing_document_jobs to service_role;

create or replace function public.enqueue_billing_document_job(
  p_job_type text,
  p_job_key text,
  p_payload jsonb default '{}'::jsonb,
  p_invoice_id uuid default null,
  p_user_id uuid default null,
  p_priority integer default 100,
  p_max_attempts integer default 4
)
returns jsonb
language sql
security definer
set search_path = public
as $$
  with
  _lock as (
    select pg_advisory_xact_lock(hashtext('billing_document_queue')::bigint)
  ),
  existing as (
    select j.id, j.status, j.lease_expires_at
    from public.billing_document_jobs j
    where j.job_key = nullif(trim(coalesce(p_job_key, '')), '')
    limit 1
    for update
  ),
  completed_existing as (
    select to_jsonb(j) as payload
    from public.billing_document_jobs j
    where j.id = (
      select e.id
      from existing e
      where e.status = 'completed'
      limit 1
    )
  ),
  updated as (
    update public.billing_document_jobs j
    set
      job_type = lower(trim(coalesce(p_job_type, j.job_type))),
      invoice_id = coalesce(p_invoice_id, j.invoice_id),
      user_id = coalesce(p_user_id, j.user_id),
      priority = coalesce(p_priority, j.priority),
      max_attempts = greatest(1, coalesce(p_max_attempts, j.max_attempts)),
      payload = coalesce(p_payload, j.payload, '{}'::jsonb),
      result = '{}'::jsonb,
      error_message = null,
      status = case
        when e.status = 'processing'
          and e.lease_expires_at is not null
          and e.lease_expires_at > timezone('utc', now()) then 'processing'
        else 'queued'
      end,
      available_at = case
        when e.status = 'processing'
          and e.lease_expires_at is not null
          and e.lease_expires_at > timezone('utc', now()) then j.available_at
        else timezone('utc', now())
      end,
      claimed_at = case
        when e.status = 'processing'
          and e.lease_expires_at is not null
          and e.lease_expires_at > timezone('utc', now()) then j.claimed_at
        else null
      end,
      claimed_by = case
        when e.status = 'processing'
          and e.lease_expires_at is not null
          and e.lease_expires_at > timezone('utc', now()) then j.claimed_by
        else null
      end,
      lease_expires_at = case
        when e.status = 'processing'
          and e.lease_expires_at is not null
          and e.lease_expires_at > timezone('utc', now()) then j.lease_expires_at
        else null
      end,
      completed_at = null,
      updated_at = timezone('utc', now())
    from existing e
    where j.id = e.id
      and e.status <> 'completed'
    returning to_jsonb(j) as payload
  ),
  inserted as (
    insert into public.billing_document_jobs as j (
      job_type,
      job_key,
      invoice_id,
      user_id,
      status,
      priority,
      attempts,
      max_attempts,
      payload,
      result,
      available_at,
      created_at,
      updated_at
    )
    select
      lower(trim(coalesce(p_job_type, ''))),
      trim(coalesce(p_job_key, '')),
      p_invoice_id,
      p_user_id,
      'queued',
      coalesce(p_priority, 100),
      0,
      greatest(1, coalesce(p_max_attempts, 4)),
      coalesce(p_payload, '{}'::jsonb),
      '{}'::jsonb,
      timezone('utc', now()),
      timezone('utc', now()),
      timezone('utc', now())
    where not exists (select 1 from existing)
      and nullif(trim(coalesce(p_job_type, '')), '') is not null
      and nullif(trim(coalesce(p_job_key, '')), '') is not null
    returning to_jsonb(j) as payload
  )
  select payload
  from (
    select payload from completed_existing
    union all
    select payload from updated
    union all
    select payload from inserted
  ) queue_result
  limit 1;
$$;

revoke all on function public.enqueue_billing_document_job(
  text,
  text,
  jsonb,
  uuid,
  uuid,
  integer,
  integer
) from public;
revoke all on function public.enqueue_billing_document_job(
  text,
  text,
  jsonb,
  uuid,
  uuid,
  integer,
  integer
) from anon;
revoke all on function public.enqueue_billing_document_job(
  text,
  text,
  jsonb,
  uuid,
  uuid,
  integer,
  integer
) from authenticated;
grant execute on function public.enqueue_billing_document_job(
  text,
  text,
  jsonb,
  uuid,
  uuid,
  integer,
  integer
) to service_role;

create or replace function public.claim_billing_document_job(
  p_worker_id text,
  p_lease_seconds integer default 180
)
returns jsonb
language sql
security definer
set search_path = public
as $$
  with
  _lock as (
    select pg_advisory_xact_lock(hashtext('billing_document_queue')::bigint)
  ),
  expired as (
    update public.billing_document_jobs
    set
      status = 'queued',
      available_at = timezone('utc', now()),
      claimed_at = null,
      claimed_by = null,
      lease_expires_at = null,
      updated_at = timezone('utc', now())
    where status = 'processing'
      and lease_expires_at is not null
      and lease_expires_at < timezone('utc', now())
    returning id
  ),
  active_processing as (
    select 1
    from public.billing_document_jobs
    where status = 'processing'
      and lease_expires_at is not null
      and lease_expires_at >= timezone('utc', now())
    limit 1
  ),
  next_job as (
    select j.id
    from public.billing_document_jobs j
    where not exists (select 1 from active_processing)
      and j.status in ('queued', 'failed')
      and coalesce(j.available_at, timezone('utc', now())) <= timezone('utc', now())
      and j.attempts < greatest(1, j.max_attempts)
    order by j.priority asc, j.created_at asc
    limit 1
    for update skip locked
  ),
  claimed as (
    update public.billing_document_jobs j
    set
      status = 'processing',
      attempts = j.attempts + 1,
      claimed_at = timezone('utc', now()),
      claimed_by = trim(coalesce(p_worker_id, '')),
      lease_expires_at = timezone('utc', now()) + make_interval(secs => greatest(30, coalesce(p_lease_seconds, 180))),
      error_message = null,
      updated_at = timezone('utc', now())
    where j.id in (select id from next_job)
    returning to_jsonb(j) as payload
  )
  select payload
  from claimed
  limit 1;
$$;

revoke all on function public.claim_billing_document_job(text, integer) from public;
revoke all on function public.claim_billing_document_job(text, integer) from anon;
revoke all on function public.claim_billing_document_job(text, integer) from authenticated;
grant execute on function public.claim_billing_document_job(text, integer) to service_role;
