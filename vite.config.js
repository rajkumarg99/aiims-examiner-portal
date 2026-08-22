import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// GH_PAGES_BASE is only needed if you deploy to GitHub Pages under a repo
// subpath (https://youruser.github.io/repo-name/) — the included
// .github/workflows/deploy-pages.yml sets it automatically from the repo
// name. Deploying to Vercel/Netlify or your own domain doesn't need this;
// it defaults to "/" (root) so `npm run dev` / `npm run build` work as-is.
export default defineConfig({
  plugins: [react()],
  base: process.env.GH_PAGES_BASE || "/",
  server: { port: 5173 },
});
