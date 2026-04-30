create table if not exists public.label_generations (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  service_type text not null,
  quantity integer not null check (quantity > 0),
  total_price numeric(10,2) not null check (total_price >= 0),
  payload jsonb not null,
  created_at timestamptz not null default now()
);

create index if not exists label_generations_user_created_idx
  on public.label_generations (user_id, created_at desc);

alter table public.label_generations enable row level security;

drop policy if exists "label_generations_select_own" on public.label_generations;
create policy "label_generations_select_own"
  on public.label_generations
  for select
  using (auth.uid() = user_id);

drop policy if exists "label_generations_insert_own" on public.label_generations;

revoke insert, update, delete on public.label_generations from anon, authenticated;
grant select on public.label_generations to authenticated;
grant select, insert, update, delete on public.label_generations to service_role;
