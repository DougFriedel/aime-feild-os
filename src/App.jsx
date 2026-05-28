-- ============================================================
-- AIME FIELD OS  —  Database Update (Run AFTER initial setup)
-- Adds: Time Cards, Weather, Subcontractors, Equipment on Site
-- ============================================================

-- Time Cards
create table if not exists public.time_cards (
  id           uuid primary key default gen_random_uuid(),
  project_id   uuid references public.projects(id) on delete cascade,
  worker_name  text not null,
  date         date not null,
  clock_in     text,
  clock_out    text,
  total_hours  numeric,
  ot_hours     numeric default 0,
  notes        text,
  created_at   timestamptz default now()
);

-- Weather Logs
create table if not exists public.weather_logs (
  id            uuid primary key default gen_random_uuid(),
  project_id    uuid references public.projects(id) on delete cascade,
  date          date not null,
  temp_high     numeric,
  temp_low      numeric,
  conditions    text,
  wind_speed    numeric,
  precipitation numeric,
  notes         text,
  created_at    timestamptz default now(),
  unique(project_id, date)
);

-- Subcontractors
create table if not exists public.subcontractors (
  id               uuid primary key default gen_random_uuid(),
  project_id       uuid references public.projects(id) on delete cascade,
  date             date not null,
  company_name     text not null,
  trade            text,
  contact_name     text,
  contact_phone    text,
  workers_count    int default 0,
  hours_worked     numeric,
  work_description text,
  created_at       timestamptz default now()
);

-- Equipment on Site (daily log)
create table if not exists public.equipment_on_site (
  id             uuid primary key default gen_random_uuid(),
  project_id     uuid references public.projects(id) on delete cascade,
  date           date not null,
  equipment_name text not null,
  quantity       int default 1,
  operator_name  text,
  hours_used     numeric,
  notes          text,
  created_at     timestamptz default now()
);

-- Add approval columns to daily_reports
alter table public.daily_reports
  add column if not exists status       text default 'submitted',
  add column if not exists pm_notes     text,
  add column if not exists approved_by  text,
  add column if not exists approved_at  timestamptz;

-- RLS policies for new tables
alter table public.time_cards         enable row level security;
alter table public.weather_logs       enable row level security;
alter table public.subcontractors     enable row level security;
alter table public.equipment_on_site  enable row level security;

create policy "open" on public.time_cards         for all using (true) with check (true);
create policy "open" on public.weather_logs       for all using (true) with check (true);
create policy "open" on public.subcontractors     for all using (true) with check (true);
create policy "open" on public.equipment_on_site  for all using (true) with check (true);
