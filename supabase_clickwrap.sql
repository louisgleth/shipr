create extension if not exists pgcrypto;

create table if not exists public.clickwrap_contract_versions (
  id uuid primary key default gen_random_uuid(),
  version text not null unique,
  title text not null,
  body_text text not null,
  hash_sha256 text not null check (char_length(hash_sha256) = 64),
  effective_at timestamptz not null default now(),
  is_active boolean not null default false,
  created_at timestamptz not null default now(),
  created_by uuid references auth.users(id) on delete set null,
  notes text
);

create unique index if not exists clickwrap_contract_single_active_idx
  on public.clickwrap_contract_versions ((is_active))
  where is_active;

create index if not exists clickwrap_contract_effective_idx
  on public.clickwrap_contract_versions (effective_at desc, created_at desc);

create table if not exists public.clickwrap_acceptances (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references auth.users(id) on delete set null,
  invite_id uuid references public.registration_invites(id) on delete set null,
  accepted_email text not null,
  invited_email text,
  contract_id uuid references public.clickwrap_contract_versions(id) on delete set null,
  contract_version text not null,
  contract_hash_sha256 text not null check (char_length(contract_hash_sha256) = 64),
  agreement_method text not null default 'scroll_clickwrap',
  scrolled_to_end boolean not null default false,
  scrolled_to_end_at timestamptz,
  agreed_at timestamptz not null,
  ip_address text,
  user_agent text,
  client_timezone text,
  client_locale text,
  proof_digest_sha256 text not null check (char_length(proof_digest_sha256) = 64),
  proof_signature_hmac text not null,
  proof_payload jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists clickwrap_acceptances_user_agreed_idx
  on public.clickwrap_acceptances (user_id, agreed_at desc, created_at desc);

create index if not exists clickwrap_acceptances_invite_idx
  on public.clickwrap_acceptances (invite_id, agreed_at desc);

create index if not exists clickwrap_acceptances_contract_idx
  on public.clickwrap_acceptances (contract_version, agreed_at desc);

alter table public.clickwrap_contract_versions enable row level security;
alter table public.clickwrap_acceptances enable row level security;

drop policy if exists clickwrap_contract_versions_select_authenticated on public.clickwrap_contract_versions;
create policy clickwrap_contract_versions_select_authenticated
  on public.clickwrap_contract_versions
  for select
  to authenticated
  using (true);

drop policy if exists clickwrap_acceptances_select_own on public.clickwrap_acceptances;
create policy clickwrap_acceptances_select_own
  on public.clickwrap_acceptances
  for select
  to authenticated
  using (auth.uid() = user_id);

insert into public.clickwrap_contract_versions (
  version,
  title,
  body_text,
  hash_sha256,
  effective_at,
  is_active,
  notes
) values (
  '2026-03-13.1',
  'Shipide Portal Service Agreement',
  $contract$
Electronic Acceptance Notice

By clicking "I agree", you consent to use an electronic signature and electronic records for this agreement.

1. Scope of Services
Shipide provides software to prepare shipment data, request transport labels, and manage related billing and reporting operations.

2. Account Security
You must keep credentials confidential and promptly report unauthorized access. Activity under your account is treated as authorized unless proven otherwise.

3. Billing and Payment
Charges may apply per label, shipment, subscription, or wallet usage. Applicable taxes, including VAT, are charged where required by law.

4. Wallet and Top Ups
Wallet balances are pre-funded. Processing time may vary by payment rail. You are responsible for using the exact transfer reference provided.

5. Acceptable Use
You will not use the platform for illegal goods, sanctioned parties, fraudulent shipments, or any prohibited postal content.

6. Data and Retention
Operational shipment and billing records are retained for compliance, service delivery, and dispute resolution.

7. Third-Party Integrations
Provider integrations such as ecommerce platforms and carriers are subject to each provider terms and service availability.

8. Service Availability
Shipide targets high availability but does not guarantee uninterrupted service. Planned maintenance and provider downtime may occur.

9. Limitation of Liability
To the maximum extent permitted by law, liability is limited to direct damages and capped to the fees paid for the affected billing period.

10. Governing Law and Forum
Unless otherwise agreed in writing, this agreement is governed by Belgian law. Disputes may be brought before competent courts in Belgium.

11. Updates to Terms
Updated terms may be published with a new version number and effective date. Continued use after notice constitutes acceptance.

12. Contact
Questions related to this agreement can be sent to legal@shipide.com.

By selecting "I agree", you confirm that you had the opportunity to review the full text and agree to be legally bound.
$contract$,
  encode(digest($contract$
Electronic Acceptance Notice

By clicking "I agree", you consent to use an electronic signature and electronic records for this agreement.

1. Scope of Services
Shipide provides software to prepare shipment data, request transport labels, and manage related billing and reporting operations.

2. Account Security
You must keep credentials confidential and promptly report unauthorized access. Activity under your account is treated as authorized unless proven otherwise.

3. Billing and Payment
Charges may apply per label, shipment, subscription, or wallet usage. Applicable taxes, including VAT, are charged where required by law.

4. Wallet and Top Ups
Wallet balances are pre-funded. Processing time may vary by payment rail. You are responsible for using the exact transfer reference provided.

5. Acceptable Use
You will not use the platform for illegal goods, sanctioned parties, fraudulent shipments, or any prohibited postal content.

6. Data and Retention
Operational shipment and billing records are retained for compliance, service delivery, and dispute resolution.

7. Third-Party Integrations
Provider integrations such as ecommerce platforms and carriers are subject to each provider terms and service availability.

8. Service Availability
Shipide targets high availability but does not guarantee uninterrupted service. Planned maintenance and provider downtime may occur.

9. Limitation of Liability
To the maximum extent permitted by law, liability is limited to direct damages and capped to the fees paid for the affected billing period.

10. Governing Law and Forum
Unless otherwise agreed in writing, this agreement is governed by Belgian law. Disputes may be brought before competent courts in Belgium.

11. Updates to Terms
Updated terms may be published with a new version number and effective date. Continued use after notice constitutes acceptance.

12. Contact
Questions related to this agreement can be sent to legal@shipide.com.

By selecting "I agree", you confirm that you had the opportunity to review the full text and agree to be legally bound.
$contract$, 'sha256'), 'hex'),
  now(),
  true,
  'Seeded by supabase_clickwrap.sql'
)
on conflict (version) do update
set
  title = excluded.title,
  body_text = excluded.body_text,
  hash_sha256 = excluded.hash_sha256,
  effective_at = excluded.effective_at,
  notes = excluded.notes;

update public.clickwrap_contract_versions
set is_active = (version = '2026-03-13.1');
