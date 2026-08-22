--
-- =====================================================================
-- EGRESS REDUCTION — run this on your EXISTING Supabase project
-- =====================================================================
--
-- Creates a public Storage bucket for the institute logo. Previously, an
-- uploaded logo was stored as a base64 image directly inside the
-- examiner_settings row — and that row gets re-downloaded by every open
-- tab (examiner or admin) every time the app refreshes, so a large embedded
-- image was being re-sent over and over for no reason. Moving it to
-- Storage means the app now just stores a short link in that row, and the
-- actual image is served with normal HTTP caching — each browser downloads
-- it once, not on every refresh.
--
-- This must be run BEFORE anyone uploads a new logo through Admin >
-- Settings — the app's upload button (src/App.jsx, AdminSettings) uploads
-- to this exact bucket name ("branding") and will fail with a clear error
-- if the bucket doesn't exist yet.
--
-- Run this AFTER security-upgrade-phase1.sql (it assumes admin sign-in is
-- already real Supabase Auth) — the bucket's write policies below require
-- being signed in as admin, same as every other admin-only action.
--
-- If you already have a logo saved as an uploaded (base64) image, it will
-- keep displaying exactly as before — nothing breaks — but it stays a
-- base64 blob, and the egress saving above won't apply to it, until you
-- re-upload it once through Admin > Settings after running this script.
-- =====================================================================

insert into storage.buckets (id, name, public)
values ('branding', 'branding', true)
on conflict (id) do nothing;

drop policy if exists "public_read_branding" on storage.objects;
create policy "public_read_branding" on storage.objects
  for select using (bucket_id = 'branding');

drop policy if exists "admin_insert_branding" on storage.objects;
create policy "admin_insert_branding" on storage.objects
  for insert to authenticated with check (bucket_id = 'branding');

drop policy if exists "admin_update_branding" on storage.objects;
create policy "admin_update_branding" on storage.objects
  for update to authenticated using (bucket_id = 'branding') with check (bucket_id = 'branding');

drop policy if exists "admin_delete_branding" on storage.objects;
create policy "admin_delete_branding" on storage.objects
  for delete to authenticated using (bucket_id = 'branding');
