import { createClient } from "@supabase/supabase-js";

// Vite exposes only VITE_-prefixed env vars to the browser. Create a
// .env file (copy .env.example) with your Supabase project's URL and
// anon/public key — Project Settings > API in the Supabase dashboard.
const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  // Fails loudly at startup rather than silently returning empty data,
  // matching how the rest of the app assumes a working connection.
  console.error(
    "Missing Supabase configuration. Copy .env.example to .env and fill in " +
    "VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY from your Supabase project's Settings > API page."
  );
}

export const supabase = createClient(url || "", anonKey || "");
