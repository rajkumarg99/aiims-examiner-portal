# Deploying the Examiner Portal via GitHub

This app is 100% static once built (`npm run build` → a `dist/` folder of
plain HTML/JS/CSS) — it talks directly to Supabase from the browser, so
there's no server process to run. That means once the code is on GitHub, any
static host can serve it. Two options below; **Vercel is the least fiddly**,
GitHub Pages keeps everything inside GitHub if you'd rather not add a third
service.

I can't push to GitHub on your behalf (no credentials), but everything here
is ready to go — the steps below are copy/paste.

## 1. Push the code to GitHub

1. On [github.com](https://github.com), click **New repository** (e.g. name
   it `aiims-examiner-portal`). Leave it empty — don't add a README/.gitignore
   there, this project already has them.
2. In a terminal, from inside this `aiims-examiner-react` folder:
   ```
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/aiims-examiner-portal.git
   git push -u origin main
   ```
   (If you don't have `git` installed or aren't sure how to authenticate,
   GitHub Desktop — desktop.github.com — does steps 2 onward with buttons
   instead of commands.)

Your `.env` file (with the real Supabase keys) is already excluded via
`.gitignore` and won't be pushed — that's expected, see step 2B below for
where those values go instead.

## 2A. Option A — Deploy with Vercel (recommended)

1. Go to [vercel.com](https://vercel.com) → sign in with GitHub → **Add New
   → Project** → pick the repo you just pushed. Vercel auto-detects Vite;
   leave the build settings as-is.
2. Before clicking Deploy, expand **Environment Variables** and add:
   - `VITE_SUPABASE_URL` = your Supabase project URL
   - `VITE_SUPABASE_ANON_KEY` = your Supabase anon/public key
3. Click **Deploy**. A couple of minutes later you'll have a live URL like
   `https://aiims-examiner-portal.vercel.app`. Every future push to `main`
   redeploys automatically.
4. Optional: **Project Settings → Domains** to attach a custom domain
   (e.g. `examiners.aiimsbibinagar.edu.in`) if you have one.

Netlify works the same way (netlify.com → Add new site → Import from
GitHub → same two env vars) if you'd prefer that instead.

## 2B. Option B — Deploy with GitHub Pages (stays fully inside GitHub)

This repo already includes `.github/workflows/deploy-pages.yml`, which
builds and publishes the app automatically on every push to `main`.

1. In the GitHub repo, go to **Settings → Secrets and variables → Actions →
   New repository secret** and add two secrets:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
2. Go to **Settings → Pages** and set **Source** to **GitHub Actions**.
3. Push to `main` (or re-run the workflow from the **Actions** tab). After
   it finishes, your site is live at `https://YOUR-USERNAME.github.io/aiims-examiner-portal/`.

Note the trailing subpath (`/repo-name/`) — the workflow sets that
automatically via `vite.config.js`'s `base` option, so links and assets
resolve correctly. If you ever move to a custom domain on GitHub Pages
instead, remove the `GH_PAGES_BASE` line from the workflow (custom domains
serve from `/`, not a subpath).

## Reminder: this is a client-only app

Both options above just serve static files — your Supabase anon key ends up
readable in the deployed JavaScript, which is expected and fine (that's what
the anon key is *for*). Run `security-upgrade-phase1.sql` (see README.md)
before your first real deploy if you haven't already — it switches admin
sign-in to a real Supabase account and locks every admin write behind it.
Examiner passwords are still checked in the browser (no examiner-side
Supabase Auth), so treat the deployed URL the same way you'd treat an
internal tool: don't publish the link beyond people who should have it.
