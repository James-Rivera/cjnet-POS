import { createBrowserClient } from "@supabase/ssr";
import { getSupabaseConfig } from "./env";

export function createClient() {
  const { url, key } = getSupabaseConfig();

  if (!url || !key) {
    return null;
  }

  return createBrowserClient(url, key);
}