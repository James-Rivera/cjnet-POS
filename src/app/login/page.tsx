"use client";

import { FormEvent, Suspense, useMemo, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { createSupabaseBrowserClient, hasSupabaseConfig } from "@/lib/pos/supabase-client";

function LoginPageContent() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const supabase = useMemo(() => createSupabaseBrowserClient(), []);
  const redirectTo = searchParams.get("redirectTo") || "/";
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!supabase) {
      setMessage("Supabase is not configured for this workspace.");
      return;
    }

    try {
      setLoading(true);
      setMessage("");
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) throw error;
      router.replace(redirectTo);
      router.refresh();
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Could not sign in.");
    } finally {
      setLoading(false);
    }
  }

  if (!hasSupabaseConfig()) {
    return (
      <main className="min-h-screen bg-[radial-gradient(circle_at_top,_rgba(255,212,0,0.16),_transparent_42%),linear-gradient(180deg,#fff9e6_0%,#f5efe1_100%)] px-4 py-10 text-foreground">
        <section className="mx-auto max-w-xl rounded-[1.5rem] border border-surface-border bg-white/85 p-6 shadow-[var(--shadow-soft)] backdrop-blur">
          <h1 className="text-2xl font-bold">CJNET POS local mode</h1>
          <p className="mt-2 text-sm leading-6 text-text-secondary">
            Supabase auth is not configured yet, so the app is running in local development mode.
          </p>
          <button type="button" className="primary-btn mt-6" onClick={() => router.replace("/") }>
            Open POS
          </button>
        </section>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-[radial-gradient(circle_at_top,_rgba(255,212,0,0.16),_transparent_42%),linear-gradient(180deg,#fff9e6_0%,#f5efe1_100%)] px-4 py-10 text-foreground">
      <section className="mx-auto grid w-full max-w-5xl gap-6 rounded-[2rem] border border-surface-border bg-white/85 p-4 shadow-[var(--shadow-soft)] backdrop-blur lg:grid-cols-[1.1fr_0.9fr] lg:p-6">
        <div className="rounded-[1.5rem] bg-[linear-gradient(145deg,#151515_0%,#2a240f_100%)] p-6 text-white">
          <p className="text-xs font-semibold uppercase tracking-[0.22em] text-[rgba(255,212,0,0.82)]">CJNET POS</p>
          <h1 className="mt-4 max-w-md text-4xl font-black leading-tight">Owner and cashier access, enforced at the database layer.</h1>
          <p className="mt-4 max-w-lg text-sm leading-6 text-[rgba(255,255,255,0.78)]">
            Sign in with your own Supabase account. The app uses protected routes, RLS, and profile roles instead of shared staff passwords.
          </p>
          <div className="mt-8 grid gap-3 text-sm text-[rgba(255,255,255,0.82)]">
            <div className="rounded-2xl border border-[rgba(255,255,255,0.12)] bg-[rgba(255,255,255,0.06)] px-4 py-3">Owner: dashboard, reports, prices, staff, audit logs.</div>
            <div className="rounded-2xl border border-[rgba(255,255,255,0.12)] bg-[rgba(255,255,255,0.06)] px-4 py-3">Staff: register, limited sales, and approved expenses only.</div>
          </div>
        </div>

        <form onSubmit={handleSubmit} className="rounded-[1.5rem] border border-surface-border bg-surface-card p-6 shadow-[var(--shadow-card)]">
          <h2 className="text-2xl font-bold">Sign in</h2>
          <p className="mt-2 text-sm leading-6 text-text-secondary">Use your individual account. Disabled users are blocked automatically.</p>

          <label className="mt-6 block">
            <span className="mb-2 block text-xs font-semibold uppercase tracking-[0.12em] text-text-secondary">Email</span>
            <input className="input-field" type="email" value={email} onChange={(event) => setEmail(event.target.value)} autoComplete="email" required />
          </label>

          <label className="mt-4 block">
            <span className="mb-2 block text-xs font-semibold uppercase tracking-[0.12em] text-text-secondary">Password</span>
            <input className="input-field" type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete="current-password" required />
          </label>

          {message ? <p className="mt-4 rounded-2xl border border-[rgba(176,54,54,0.2)] bg-[rgba(176,54,54,0.08)] px-4 py-3 text-sm text-[#8f2d2d]">{message}</p> : null}

          <button type="submit" className="primary-btn mt-6 w-full" disabled={loading}>
            {loading ? "Signing in..." : "Sign in to CJNET POS"}
          </button>
        </form>
      </section>
    </main>
  );
}

export default function LoginPage() {
  return (
    <Suspense fallback={null}>
      <LoginPageContent />
    </Suspense>
  );
}