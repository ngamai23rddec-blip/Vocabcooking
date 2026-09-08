import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

// Login is completely optional. If these two environment variables aren't
// set (in Vercel: Project Settings -> Environment Variables), the app runs
// exactly as before — no login screen, data saved in the browser's
// localStorage. Once both variables are set and redeployed, a login/sign-up
// screen appears automatically and data syncs to Supabase per account.
export const isSupabaseConfigured = Boolean(url && anonKey);

export const supabase = isSupabaseConfigured ? createClient(url, anonKey) : null;
