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

create or replace function public.enqueue_billing_document_job(
  p_job_type text,
  p_job_key text,
  p_payload jsonb default '{}'::jsonb,
  p_invoice_id uuid default null,
  p_user_id uuid default null,
  p_priority integer default 100,
  p_max_attempts integer default 4
)
returns public.billing_document_jobs
language plpgsql
security definer
set search_path = public
as $$
declare
  v_existing public.billing_document_jobs%rowtype;
  v_now timestamptz := timezone('utc', now());
begin
  if nullif(trim(coalesce(p_job_type, '')), '') is null then
    raise exception 'Job type is required.';
  end if;
  if nullif(trim(coalesce(p_job_key, '')), '') is null then
    raise exception 'Job key is required.';
  end if;

  perform pg_advisory_xact_lock(hashtext('billing_document_queue')::bigint);

  select *
  into v_existing
  from public.billing_document_jobs
  where job_key = p_job_key
  limit 1
  for update;

  if found then
    if v_existing.status = 'completed' then
      return v_existing;
    end if;

    update public.billing_document_jobs
    set
      job_type = p_job_type,
      invoice_id = coalesce(p_invoice_id, invoice_id),
      user_id = coalesce(p_user_id, user_id),
      priority = coalesce(p_priority, priority),
      max_attempts = greatest(1, coalesce(p_max_attempts, max_attempts)),
      payload = coalesce(p_payload, payload, '{}'::jsonb),
      result = case
        when status = 'completed' then result
        else '{}'::jsonb
      end,
      error_message = null,
      status = case
        when status = 'processing' and lease_expires_at is not null and lease_expires_at > v_now then status
        when status = 'completed' then status
        else 'queued'
      end,
      available_at = case
        when status = 'processing' and lease_expires_at is not null and lease_expires_at > v_now
          then available_at
        else v_now
      end,
      claimed_at = case
        when status = 'processing' and lease_expires_at is not null and lease_expires_at > v_now
          then claimed_at
        else null
      end,
      claimed_by = case
        when status = 'processing' and lease_expires_at is not null and lease_expires_at > v_now
          then claimed_by
        else null
      end,
      lease_expires_at = case
        when status = 'processing' and lease_expires_at is not null and lease_expires_at > v_now
          then lease_expires_at
        else null
      end,
      completed_at = case
        when status = 'completed' then completed_at
        else null
      end,
      updated_at = v_now
    where id = v_existing.id
    returning * into v_existing;

    return v_existing;
  end if;

  insert into public.billing_document_jobs (
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
  values (
    p_job_type,
    p_job_key,
    p_invoice_id,
    p_user_id,
    'queued',
    coalesce(p_priority, 100),
    0,
    greatest(1, coalesce(p_max_attempts, 4)),
    coalesce(p_payload, '{}'::jsonb),
    '{}'::jsonb,
    v_now,
    v_now,
    v_now
  )
  returning * into v_existing;

  return v_existing;
end;
$$;

create or replace function public.claim_billing_document_job(
  p_worker_id text,
  p_lease_seconds integer default 180
)
returns public.billing_document_jobs
language plpgsql
security definer
set search_path = public
as $$
declare
  v_claimed public.billing_document_jobs%rowtype;
  v_now timestamptz := timezone('utc', now());
  v_lease interval := make_interval(secs => greatest(30, coalesce(p_lease_seconds, 180)));
begin
  if nullif(trim(coalesce(p_worker_id, '')), '') is null then
    raise exception 'Worker id is required.';
  end if;

  perform pg_advisory_xact_lock(hashtext('billing_document_queue')::bigint);

  update public.billing_document_jobs
  set
    status = 'queued',
    available_at = v_now,
    claimed_at = null,
    claimed_by = null,
    lease_expires_at = null,
    updated_at = v_now
  where status = 'processing'
    and lease_expires_at is not null
    and lease_expires_at < v_now;

  if exists (
    select 1
    from public.billing_document_jobs
    where status = 'processing'
      and lease_expires_at is not null
      and lease_expires_at >= v_now
  ) then
    return null;
  end if;

  with next_job as (
    select id
    from public.billing_document_jobs
    where status in ('queued', 'failed')
      and coalesce(available_at, v_now) <= v_now
      and attempts < greatest(1, max_attempts)
    order by priority asc, created_at asc
    limit 1
    for update skip locked
  )
  update public.billing_document_jobs
  set
    status = 'processing',
    attempts = attempts + 1,
    claimed_at = v_now,
    claimed_by = p_worker_id,
    lease_expires_at = v_now + v_lease,
    error_message = null,
    updated_at = v_now
  where id in (select id from next_job)
  returning * into v_claimed;

  return v_claimed;
end;
$$;
