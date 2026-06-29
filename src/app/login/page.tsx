"use client";

import Image from "next/image";
import dynamic from "next/dynamic";
import { FormEvent, Suspense, useMemo, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { createSupabaseBrowserClient, hasSupabaseConfig } from "@/lib/pos/supabase-client";

const LoginShaderGradient = dynamic(
  () => import("@/app/_components/login-shader-gradient").then((module) => module.LoginShaderGradient),
  { ssr: false },
);

function LoginPageContent() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const supabase = useMemo(() => createSupabaseBrowserClient(), []);
  const redirectTo = searchParams.get("redirectTo") || "/";
  const accessDenied = searchParams.get("error") === "access_denied";
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState(() => (accessDenied ? "Your account cannot access this register." : ""));

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
      <main className="login-screen">
        <BrandPanel />
        <section className="login-form-panel" aria-labelledby="local-login-title">
          <div className="login-form-shell">
            <Image src="/logo.png" alt="CJNET Internet Cafe and Xerox Copier" width={920} height={311} className="login-logo" priority />

            <div className="login-title-block">
              <h1 id="local-login-title">Local setup</h1>
            </div>

            <p className="login-helper">Connect Supabase when you are ready to use staff accounts.</p>

            <button type="button" className="login-submit" onClick={() => router.replace("/")}>
              Open Register
            </button>
          </div>
        </section>
      </main>
    );
  }

  return (
    <main className="login-screen">
      <BrandPanel />
      <section className="login-form-panel" aria-labelledby="login-title">
        <form onSubmit={handleSubmit} className="login-form-shell">
          <Image src="/logo.png" alt="CJNET Internet Cafe and Xerox Copier" width={920} height={311} className="login-logo" priority />

          <div className="login-title-block">
            <h1 id="login-title">Login</h1>
          </div>

          <label className="login-field">
            <span>Email</span>
            <input
              className="login-input"
              type="email"
              value={email}
              onChange={(event) => setEmail(event.target.value)}
              autoComplete="email"
              required
            />
          </label>

          <label className="login-field">
            <span>Password</span>
            <div className="login-password-wrap">
              <input
                className="login-input"
                type={showPassword ? "text" : "password"}
                value={password}
                onChange={(event) => setPassword(event.target.value)}
                autoComplete="current-password"
                required
              />
              <button
                type="button"
                className="login-password-toggle"
                onClick={() => setShowPassword((current) => !current)}
                aria-label={showPassword ? "Hide password" : "Show password"}
              >
                <Image src="/icons/hide-outline.svg" alt="" width={22} height={22} aria-hidden="true" />
              </button>
            </div>
          </label>

          {message ? <p className="login-message">{message}</p> : null}

          <button type="submit" className="login-submit" disabled={loading}>
            {loading ? "Logging in..." : "Login"}
          </button>
        </form>
      </section>
    </main>
  );
}

function BrandPanel() {
  return (
    <section className="login-brand-panel" aria-hidden="true">
      <LoginShaderGradient />
      <Image src="/cjnet-mark.png" alt="" width={256} height={256} className="login-mark" priority />
    </section>
  );
}

export default function LoginPage() {
  return (
    <Suspense fallback={null}>
      <LoginPageContent />
    </Suspense>
  );
}
