-- AIIMS Bibinagar Examiner Portal — Supabase schema
-- Run this once in your Supabase project's SQL editor (Database > SQL Editor > New query).
--
-- Security model note (matches the existing Hall Ticket app's approach): this is a
-- client-only React app using the public "anon" key, with a single shared admin
-- password checked in the browser rather than real Supabase Auth. That means anyone
-- who can reach your Supabase URL + anon key can read/write these tables directly
-- (e.g. via browser devtools), not just through the app's UI — the same trade-off
-- the Hall Ticket app already makes with its own admin password. Keep this suitable
-- for a trusted internal deployment; ask if you'd like real Supabase Auth + properly
-- scoped Row Level Security added later for a hardened version.
--
-- RLS is turned ON below (Supabase flags new tables with it off), but each
-- table gets a single permissive "allow all" policy — so behavior today is
-- identical to RLS being off, it just gives you a named policy per table to
-- edit/replace later instead of having to remember to enable RLS from scratch.

create extension if not exists "pgcrypto";

-- Singleton row (id = 1) holding portal-wide settings.
create table if not exists examiner_settings (
  id int primary key default 1,
  institute_name text default 'All India Institute of Medical Sciences, Bibinagar',
  portal_subtitle text default 'Examination Cell — Examiner Bill / Requisition Portal',
  public_base_url text default '',
  admin_password text default 'admin123',
  logo_url text default '',
  min_remuneration numeric default 3000,
  updated_at timestamptz default now()
);
insert into examiner_settings (id) values (1) on conflict (id) do nothing;
-- Safety net for a project whose examiner_settings table was already created before
-- these two columns existed in this file (they were added to the app after the table
-- itself) — re-running this whole script is otherwise idempotent, but a plain CREATE
-- TABLE IF NOT EXISTS above silently skips column changes on an existing table, so
-- these explicit ALTERs make sure both columns are actually present either way. If
-- Admin > Settings was throwing an error on every save (not just when touching the
-- logo), a missing logo_url column here was almost certainly why.
alter table examiner_settings add column if not exists logo_url text default '';
alter table examiner_settings add column if not exists min_remuneration numeric default 3000;

-- One row per examiner account. Admin creates these (singly or via CSV bulk
-- import) with an auto-generated password, then shares the email + password
-- with that examiner out of band (email/print) — the examiner never self-registers.
create table if not exists examiners (
  id text primary key,
  email text unique not null,
  password text not null,
  name text default '',
  designation text default '',
  mobile text default '',
  category_hint text default '', -- MBBS | MSC | PG, just pre-selects the category on login
  status text default 'invited', -- invited | submitted
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

-- Admin-editable field & rate templates for the three documents. One row per
-- form ("remuneration" | "vehicle" | "flight"). `schema` holds the full field
-- definition (sections, fields, table columns, computed formulas/rates/caps,
-- conditional visibility) as JSON — see src/App.jsx defaultFormConfigs() for
-- the shape and the seeded defaults matching AIIMS Bibinagar's current forms.
create table if not exists examiner_form_configs (
  form_id text primary key,
  schema jsonb not null,
  updated_at timestamptz default now()
);

-- One row per examiner submission. `data` holds everything the examiner (and
-- later the admin) filled in: profile, remuneration workload + travel/
-- accommodation/DA claims, vehicle requisition, flight booking — keyed the
-- same way as examiner_form_configs.schema so the same renderer can display
-- either. Examiners can only create/update their own row while status is
-- 'submitted'; only the admin marks a row 'generated' once the three
-- documents have been produced.
create table if not exists examiner_submissions (
  id text primary key,
  examiner_id text references examiners(id) on delete cascade,
  ref_no text unique,
  status text default 'submitted', -- submitted | generated
  data jsonb not null default '{}'::jsonb,
  generated boolean default false,
  generated_at timestamptz,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create index if not exists idx_submissions_examiner on examiner_submissions(examiner_id);

-- Enable RLS + one permissive policy per table. This keeps the anon key able
-- to read/write everything (required — the app has no other credential), but
-- gives each table an explicit, named policy rather than relying on RLS
-- simply being off. To harden later: replace these four policies with ones
-- scoped to an authenticated role instead of "true" for every operation.
-- drop-then-create makes this whole script safe to re-run in full (e.g. after a
-- schema.sql update) without the "policy already exists" error stopping the batch
-- partway through and rolling back everything before it, including the ALTER TABLE
-- column additions above.
alter table examiner_settings enable row level security;
drop policy if exists "allow all (anon key) - examiner_settings" on examiner_settings;
create policy "allow all (anon key) - examiner_settings" on examiner_settings
  for all using (true) with check (true);

alter table examiners enable row level security;
drop policy if exists "allow all (anon key) - examiners" on examiners;
create policy "allow all (anon key) - examiners" on examiners
  for all using (true) with check (true);

alter table examiner_form_configs enable row level security;
drop policy if exists "allow all (anon key) - examiner_form_configs" on examiner_form_configs;
create policy "allow all (anon key) - examiner_form_configs" on examiner_form_configs
  for all using (true) with check (true);

alter table examiner_submissions enable row level security;
drop policy if exists "allow all (anon key) - examiner_submissions" on examiner_submissions;
create policy "allow all (anon key) - examiner_submissions" on examiner_submissions
  for all using (true) with check (true);
