do $$
begin
  if to_regclass('public."Session"') is not null then
    execute 'alter table public."Session" enable row level security';
    execute 'revoke all on table public."Session" from public';
    execute 'revoke all on table public."Session" from anon';
    execute 'revoke all on table public."Session" from authenticated';
    execute 'grant all on table public."Session" to service_role';
    execute 'drop policy if exists "session_service_role_all" on public."Session"';
    execute 'create policy "session_service_role_all" on public."Session" for all to service_role using (true) with check (true)';
  end if;

  if to_regclass('public._prisma_migrations') is not null then
    execute 'alter table public._prisma_migrations enable row level security';
    execute 'revoke all on table public._prisma_migrations from public';
    execute 'revoke all on table public._prisma_migrations from anon';
    execute 'revoke all on table public._prisma_migrations from authenticated';
    execute 'grant all on table public._prisma_migrations to service_role';
    execute 'drop policy if exists "prisma_migrations_service_role_all" on public._prisma_migrations';
    execute 'create policy "prisma_migrations_service_role_all" on public._prisma_migrations for all to service_role using (true) with check (true)';
  end if;
end
$$;
