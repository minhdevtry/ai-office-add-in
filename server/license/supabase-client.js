/**
 * Supabase Client Singleton
 * Sử dụng service_role key (full quyền) — KHÔNG BAO GIỜ expose ra client.
 * Server-side only.
 */

import { createClient } from "@supabase/supabase-js";
import { config } from "../config.js";

let _supabase = null;

export function getSupabase() {
  if (_supabase) return _supabase;

  if (!config.supabase.url || !config.supabase.serviceKey) {
    throw new Error(
      "Supabase chưa được cấu hình. Set SUPABASE_URL và SUPABASE_SERVICE_KEY trong .env"
    );
  }

  _supabase = createClient(config.supabase.url, config.supabase.serviceKey, {
    auth: {
      autoRefreshToken: false,
      persistSession: false,
    },
  });

  return _supabase;
}

export function isSupabaseConfigured() {
  return Boolean(config.supabase.url && config.supabase.serviceKey);
}
