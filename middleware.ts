import { NextRequest, NextResponse } from "next/server";
import { createClient } from "@/utils/supabase/middleware";

export async function middleware(request: NextRequest) {
  const { pathname } = request.nextUrl;
  const isStatic =
    pathname.startsWith("/_next") ||
    pathname.startsWith("/icons") ||
    pathname === "/favicon.ico" ||
    pathname === "/logo.png" ||
    pathname === "/cjnet-mark.png";
  if (isStatic) {
    return NextResponse.next();
  }

  const { supabase, response } = createClient(request);

  if (!supabase) {
    return NextResponse.next();
  }

  const { data: authData } = await supabase.auth.getUser();
  const user = authData.user;

  if (pathname === "/login") {
    if (user) {
      return NextResponse.redirect(new URL("/", request.url));
    }
    return response;
  }

  if (!user) {
    const loginUrl = new URL("/login", request.url);
    loginUrl.searchParams.set("redirectTo", pathname);
    return NextResponse.redirect(loginUrl);
  }

  const { data: profile } = await supabase.from("profiles").select("status").eq("id", user.id).maybeSingle();
  if (!profile || profile.status === "disabled") {
    await supabase.auth.signOut();
    return NextResponse.redirect(new URL("/login?error=access_denied", request.url));
  }

  return response;
}

export const config = {
  matcher: ["/((?!_next/static|_next/image|favicon.ico|logo.png|cjnet-mark.png|icons/).*)"],
};