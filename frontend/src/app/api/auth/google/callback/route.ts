import { NextRequest, NextResponse } from "next/server";
import { clerkClient } from "@clerk/nextjs/server";
import crypto from "crypto";

const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || "";
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || "";
const GOOGLE_REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI || (process.env.NEXT_PUBLIC_APP_URL ? `${process.env.NEXT_PUBLIC_APP_URL}/api/auth/google/callback` : "http://localhost:3000/api/auth/google/callback");

function verifyStateHmac(nonce: string, userId: string, hmac: string): boolean {
  const secret = process.env.CLERK_SECRET_KEY;
  if (!secret) throw new Error("CLERK_SECRET_KEY is required for HMAC verification");
  const expected = crypto.createHmac("sha256", secret).update(`${nonce}:${userId}`).digest("hex");
  return crypto.timingSafeEqual(Buffer.from(expected), Buffer.from(hmac));
}

export async function GET(req: NextRequest) {
  try {
    const code = req.nextUrl.searchParams.get("code");
    const state = req.nextUrl.searchParams.get("state");

    if (!code || !state) {
      console.error("[Google OAuth] Missing params - code:", !!code, "state:", !!state);
      return NextResponse.redirect(new URL("/settings?error=missing_params", req.url));
    }

    // Parse state: "nonce:userId:hmac"
    const parts = state.split(":");
    if (parts.length < 3) {
      console.error("[Google OAuth] Invalid state format");
      return NextResponse.redirect(new URL("/settings?error=invalid_state", req.url));
    }

    const nonce = parts[0];
    // userId may contain colons (unlikely but safe), hmac is always last 64 chars
    const hmac = parts[parts.length - 1];
    const userId = parts.slice(1, parts.length - 1).join(":");

    if (!nonce || !userId || !hmac) {
      console.error("[Google OAuth] Empty nonce, userId, or hmac in state");
      return NextResponse.redirect(new URL("/settings?error=invalid_state", req.url));
    }

    // Verify the nonce matches the cookie set during authorization
    const storedNonce = req.cookies.get("google_oauth_state")?.value;
    if (!storedNonce || storedNonce !== nonce) {
      console.error("[Google OAuth] State mismatch - possible CSRF attack");
      return NextResponse.redirect(new URL("/settings?error=state_mismatch", req.url));
    }

    // Verify HMAC to ensure userId hasn't been tampered with
    try {
      if (!verifyStateHmac(nonce, userId, hmac)) {
        console.error("[Google OAuth] HMAC verification failed - possible userId tampering");
        return NextResponse.redirect(new URL("/settings?error=state_mismatch", req.url));
      }
    } catch {
      console.error("[Google OAuth] HMAC verification error");
      return NextResponse.redirect(new URL("/settings?error=state_mismatch", req.url));
    }

    console.log("[Google OAuth] State verified, exchanging code for tokens, userId:", userId);

    // Exchange code for tokens
    const tokenRes = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        code,
        client_id: GOOGLE_CLIENT_ID,
        client_secret: GOOGLE_CLIENT_SECRET,
        redirect_uri: GOOGLE_REDIRECT_URI,
        grant_type: "authorization_code",
      }),
    });

    if (!tokenRes.ok) {
      console.error("[Google OAuth] Token exchange failed with status:", tokenRes.status);
      return NextResponse.redirect(new URL("/settings?error=token_exchange", req.url));
    }

    const tokens = await tokenRes.json();
    console.log("[Google OAuth] Token exchange success, has refresh_token:", !!tokens.refresh_token);

    // Store tokens in Clerk private metadata (merge with existing to preserve other keys)
    const clerk = await clerkClient();
    const user = await clerk.users.getUser(userId);
    const currentPublic = (user.publicMetadata || {}) as Record<string, unknown>;
    const currentPrivate = (user.privateMetadata || {}) as Record<string, unknown>;

    console.log("[Google OAuth] Updating Clerk metadata for user:", userId);
    await clerk.users.updateUserMetadata(userId, {
      publicMetadata: {
        ...currentPublic,
        googleConnected: true,
      },
      privateMetadata: {
        ...currentPrivate,
        googleAccessToken: tokens.access_token,
        googleRefreshToken: tokens.refresh_token,
        googleTokenExpiry: Date.now() + tokens.expires_in * 1000,
      },
    });

    console.log("[Google OAuth] Metadata updated, redirecting to /settings?google=connected");

    // Clear the state cookie
    const response = NextResponse.redirect(new URL("/settings?google=connected", req.url));
    response.cookies.delete("google_oauth_state");
    return response;
  } catch (err) {
    console.error("[Google OAuth] Unhandled error:", err);
    return NextResponse.redirect(new URL("/settings?error=server_error", req.url));
  }
}
