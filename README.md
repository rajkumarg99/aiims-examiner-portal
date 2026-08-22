# AIIMS Bibinagar — Examiner Portal (React + Supabase)

Built the same way as the Hall Ticket app: a single `src/App.jsx`, plain inline
styles (no CSS framework), Supabase for the database, and the same
Landing → role portal → tabbed admin console pattern.

## What it does

- **Examiners** sign in with an email + password issued by the admin (no
  self-registration) and fill in one combined form covering their
  Remuneration Bill, Vehicle Requisition, and Flight Booking details. They
  submit once and never see a "generate" or print button.
- **Admins** create examiner accounts (one at a time or bulk via CSV, each
  gets a random password to hand out), configure every field and fee
  rate/cap on all three documents from a Form Builder (no code), review and
  adjust each submission, then **Generate** the three documents — singly or
  in bulk by category — and print/save them as PDF. A CSV register of all
  submissions is available from Reports.

## 1. Create a Supabase project

1. Go to [supabase.com](https://supabase.com), create a free project.
2. Open **SQL Editor → New query**, paste the contents of `schema.sql` from
   this folder, and run it. This creates the four tables the app needs.
3. Open **Project Settings → API** and copy the **Project URL** and the
   **anon public** key.
4. **Lock down admin access** — open **Authentication → Users → Add user**
   and create a real account for yourself (email + password, tick "Auto
   Confirm User"). Then open **SQL Editor → New query** again, paste the
   contents of `security-upgrade-phase1.sql`, and run it. This switches
   Admin sign-in from a shared password to that real Supabase account, and
   requires it for every admin write (Settings, Form Builder, Exam Rates,
   examiners, deleting submissions). Do this before sharing the portal's
   link with anyone — see the comments at the top of that file for exactly
   what it changes.
5. **Cut down data usage** — open **SQL Editor → New query** once more,
   paste the contents of `egress-reduction.sql`, and run it. This creates a
   public "branding" Storage bucket so an uploaded institute logo is served
   as a real cached file instead of being re-downloaded in full every time
   the app refreshes. Run this after step 4 (it reuses the same admin
   sign-in for uploads). If you already have a logo saved, re-upload it
   once from Admin → Settings afterward to move it into Storage.

> **Security note:** like the Hall Ticket app, this is a client-only React
> app. Once `security-upgrade-phase1.sql` has been run, admin sign-in is a
> real Supabase account and every admin write requires it — but examiner
> passwords are still checked in the browser (there's no examiner-side
> Supabase Auth), and the `examiners`/`examiner_submissions` tables are
> still publicly *readable* via the anon key (locking reads down too is a
> bigger change — secure lookup functions instead of raw table reads — ask
> if you want that added next). Keep this off the public internet regardless.

## 2. Configure and run

```
cd aiims-examiner-react
cp .env.example .env        # then paste your Supabase URL + anon key into .env
npm install
npm run dev
```

Open the URL Vite prints (usually `http://localhost:5173`).

- Examiner portal: click **Examiner login** on the landing page.
- Admin console: click **Administrator login** and sign in with the Supabase
  account you created in step 1.4 above (email + password). Add or remove
  admin accounts, or reset a password, from Supabase → Authentication → Users
  — there's no admin password field in the app's Settings tab anymore.

To build a static production bundle for real hosting: `npm run build`, then
deploy the `dist/` folder anywhere that serves static files (Netlify,
Vercel, your own web server, etc.) — no server-side code is needed since all
data access goes straight to Supabase from the browser.

## 3. Day-to-day use

1. **Examiners tab** — add examiners (single or CSV bulk import; template
   download provided). Each gets an auto-generated password shown once and
   downloadable as a credentials CSV — email/print that to the examiner
   yourself.
2. **Form Builder tab** — edit labels, field types, required/optional,
   dropdown options, section order, conditional visibility (e.g. which
   remuneration fields show for MBBS vs M.Sc. vs PG), and fee
   formulas/rates/min-max caps for all three documents. Changes are live on
   the examiner form immediately.
3. Examiners log in and submit.
4. **Submissions tab** — expand any submission to review or correct what was
   entered (including the workload counts that drive the fee formulas),
   then **Generate documents**. Generated submissions can be printed
   individually, or generate + print in bulk by category.
5. **Reports tab** — export a CSV register of all submissions.
6. **Backup tab** — export/import the Form Builder configuration + settings
   as one JSON file (examiners and submissions aren't included; export those
   from their own tabs).

## Files

- `src/App.jsx` — the whole application.
- `src/supabaseClient.js` — Supabase client setup (reads `.env`).
- `schema.sql` — run once in Supabase's SQL editor.
- `security-upgrade-phase1.sql` — run once, after `schema.sql` and after
  creating your admin account, to require real sign-in for admin writes.
- `egress-reduction.sql` — run once, after `security-upgrade-phase1.sql`, to
  set up Storage for the logo instead of storing it in the database.
- `.env.example` — copy to `.env` and fill in your project's credentials.
