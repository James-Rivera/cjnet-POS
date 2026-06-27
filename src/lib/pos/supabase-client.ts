import { createClient as createBrowserSupabaseClient } from "@/utils/supabase/client";
import { getSupabaseConfig } from "@/utils/supabase/env";

export function hasSupabaseConfig() {
  return getSupabaseConfig().hasConfig;
}

export function createSupabaseBrowserClient() {
  return createBrowserSupabaseClient();
}
