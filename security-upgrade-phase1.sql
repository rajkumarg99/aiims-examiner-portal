--
-- =====================================================================
-- SECURITY UPGRADE — PHASE 1 — run this on your EXISTING Supabase project
-- =====================================================================
--
-- Replaces the "allow all (anon key)" policies from schema.sql (anyone with
-- the anon key can do anything) with real role-based access:
--   - Reading data the app genuinely needs to show without anyone being
--     signed in (institute name/logo/settings, the form field/rate
--     configuration, the examiner list for login, submissions so an
--     examiner can reload and continue their own entry) stays open — there
--     is still no separate "examiner login via Supabase Auth" system, so
--     this can't require authentication without breaking the examiner side.
--   - Every CHANGE an admin makes (Settings, Form Builder / Exam Rates,
--     adding/removing/editing examiners, deleting a submission) now
--     requires being signed in as a real Supabase user — no longer just
--     "knew a password stored in a public table."
--   - The one exception: an examiner submitting their form still needs to
--     create/update their own submission row and flip their own examiner
--     record to "submitted" without being signed in (there's no examiner
--     login system, by design). The submissions table stays open to
--     insert/update for that reason (same trade-off the reference PDF
--     makes for its own "applications" table). The examiner-status flip
--     goes through a narrow function instead of a raw table update, so an
--     anonymous caller can only ever set that one flag on their own row —
--     nothing else on the examiners table is writable without signing in.
--
-- BEFORE RUNNING THIS: create at least one admin account, or you'll lock
-- yourself out of the admin console entirely.
--   1. Supabase dashboard -> Authentication -> Users -> Add user
--   2. Enter an email and password for yourself (tick "Auto Confirm User")
--   3. That email/password is what you'll use to sign in to Admin from now on
--
-- The app no longer reads or writes the old admin_password column at all —
-- it's fine to leave the column in place (this script doesn't drop it), or
-- drop it yourself later once you've confirmed real sign-in is working:
--     alter table examiner_settings drop column admin_password;
--
-- Still open for later (not covered by this script): reading the examiners
-- table (names/emails/mobiles/their issued passwords) is still public, same
-- as submissions data. Locking that down too needs a bigger change (secure
-- lookup functions instead of raw table reads, matching the Phase 2 pattern
-- from the reference PDFs) — ask if you want to go ahead with that next.
-- =====================================================================

-- ---------------------------------------------------------------------
-- examiner_settings — public read (institute name, logo, min remuneration
-- all need to show before anyone signs in), admin-only writes
-- ---------------------------------------------------------------------
drop policy if exists "allow all (anon key) - examiner_settings" on examiner_settings;
drop policy if exists "public_read_examiner_settings" on examiner_settings;
drop policy if exists "admin_write_examiner_settings" on examiner_settings;
drop policy if exists "admin_update_examiner_settings" on examiner_settings;
create policy "public_read_examiner_settings" on examiner_settings for select using (true);
create policy "admin_write_examiner_settings" on examiner_settings for insert to authenticated with check (true);
create policy "admin_update_examiner_settings" on examiner_settings for update to authenticated using (true) with check (true);

-- ---------------------------------------------------------------------
-- examiner_form_configs — public read (the examiner portal needs the field/
-- rate schema to render forms), admin-only writes (Form Builder, Exam Rates)
--
-- Note: on a brand-new project with zero rows, the app's very first ever
-- page load seeds the four default form configs automatically. If that
-- first-ever visitor is an examiner (not signed in), the seed insert will
-- be rejected and the examiner still sees a working form locally that
-- visit, just not saved to the database yet — the first time an admin
-- actually signs in, the seed succeeds and everything is fine from then
-- on. Signing in as admin once right after deploying avoids this entirely.
-- ---------------------------------------------------------------------
drop policy if exists "allow all (anon key) - examiner_form_configs" on examiner_form_configs;
drop policy if exists "public_read_examiner_form_configs" on examiner_form_configs;
drop policy if exists "admin_write_examiner_form_configs" on examiner_form_configs;
drop policy if exists "admin_update_examiner_form_configs" on examiner_form_configs;
create policy "public_read_examiner_form_configs" on examiner_form_configs for select using (true);
create policy "admin_write_examiner_form_configs" on examiner_form_configs for insert to authenticated with check (true);
create policy "admin_update_examiner_form_configs" on examiner_form_configs for update to authenticated using (true) with check (true);

-- ---------------------------------------------------------------------
-- examiners — public read stays for now (needed for examiner email/
-- password login and for the admin's own list; locking reads down is
-- Phase 2 territory). Adding, editing or removing an examiner record now
-- requires a real admin sign-in.
-- ---------------------------------------------------------------------
drop policy if exists "allow all (anon key) - examiners" on examiners;
drop policy if exists "public_read_examiners" on examiners;
drop policy if exists "admin_insert_examiners" on examiners;
drop policy if exists "admin_update_examiners" on examiners;
drop policy if exists "admin_delete_examiners" on examiners;
create policy "public_read_examiners" on examiners for select using (true);
create policy "admin_insert_examiners" on examiners for insert to authenticated with check (true);
create policy "admin_update_examiners" on examiners for update to authenticated using (true) with check (true);
create policy "admin_delete_examiners" on examiners for delete to authenticated using (true);

-- ---------------------------------------------------------------------
-- mark_examiner_submitted — the one write an anonymous examiner still
-- needs on the examiners table: flip their own record's status to
-- "submitted" once they submit their form. This function can ONLY do that
-- one thing to that one row (nothing else is exposed), so it's safe to
-- grant to anon even though the raw table's UPDATE above is admin-only.
-- ---------------------------------------------------------------------
create or replace function mark_examiner_submitted(p_examiner_id text)
returns void
language sql
security definer
set search_path = public
as $$
  update examiners set status = 'submitted', updated_at = now() where id = p_examiner_id;
$$;
revoke all on function mark_examiner_submitted(text) from public;
grant execute on function mark_examiner_submitted(text) to anon, authenticated;

-- ---------------------------------------------------------------------
-- examiner_submissions — public read/insert/update stays (an examiner
-- creates their own submission and can re-edit it before it's generated,
-- all without being signed in — there's no examiner login system). Only
-- DELETING a submission (an admin-only action in the Submissions tab) now
-- requires a real admin sign-in.
-- ---------------------------------------------------------------------
drop policy if exists "allow all (anon key) - examiner_submissions" on examiner_submissions;
drop policy if exists "public_read_examiner_submissions" on examiner_submissions;
drop policy if exists "public_insert_examiner_submissions" on examiner_submissions;
drop policy if exists "public_update_examiner_submissions" on examiner_submissions;
drop policy if exists "admin_delete_examiner_submissions" on examiner_submissions;
create policy "public_read_examiner_submissions" on examiner_submissions for select using (true);
create policy "public_insert_examiner_submissions" on examiner_submissions for insert with check (true);
create policy "public_update_examiner_submissions" on examiner_submissions for update using (true) with check (true);
create policy "admin_delete_examiner_submissions" on examiner_submissions for delete to authenticated using (true);
