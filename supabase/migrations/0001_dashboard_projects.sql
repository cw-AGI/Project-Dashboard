-- Project dashboard: one row per project (config + data JSON).
-- Run in Supabase SQL Editor, or: supabase db push (if using Supabase CLI).

create table if not exists public.dashboard_projects (
  project_id text primary key,
  display_name text not null,
  color text not null default '#1d4ed8',
  active boolean not null default true,
  path text,
  config jsonb not null default '{}'::jsonb,
  data jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

create index if not exists dashboard_projects_active_idx
  on public.dashboard_projects (active);

alter table public.dashboard_projects enable row level security;

-- Public read (dashboard visitors load data without Supabase login).
create policy "dashboard_projects_select_anon"
  on public.dashboard_projects for select
  using (true);

-- Writes require a logged-in Supabase Auth user (create Editor users in Authentication).
create policy "dashboard_projects_insert_auth"
  on public.dashboard_projects for insert
  to authenticated
  with check (true);

create policy "dashboard_projects_update_auth"
  on public.dashboard_projects for update
  to authenticated
  using (true)
  with check (true);

create policy "dashboard_projects_delete_auth"
  on public.dashboard_projects for delete
  to authenticated
  using (true);
