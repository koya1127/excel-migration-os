import { NextRequest, NextResponse } from "next/server";
import { clerkClient } from "@clerk/nextjs/server";

const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || "";
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || "";
const GOOGLE_REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI || "http://localhost:3000/api/auth/google/callback";

export async function GET(req: NextRequest) {
  try {
    const code = req.nextUrl.searchParams.get("code");
    const userId = req.nextUrl.searchParams.get("state");

    if (!code || !userId) {
      console.error("[Google OAuth] Missing params - code:", !!code, "userId:", !!userId);
      return NextResponse.redirect(new URL("/settings?error=missing_params", req.url));
    }

    console.log("[Google OAuth] Exchanging code for tokens, userId:", userId);

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
      const errorBody = await tokenRes.text();
      console.error("[Google OAuth] Token exchange failed:", tokenRes.status, errorBody);
      return NextResponse.redirect(new URL("/settings?error=token_exchange", req.url));
    }

    const tokens = await tokenRes.json();
    console.log("[Google OAuth] Token exchange success, has refresh_token:", !!tokens.refresh_token);

    // Store tokens in Clerk private metadata
    const clerk = await clerkClient();
    const user = await clerk.users.getUser(userId);
    const currentPublic = (user.publicMetadata || {}) as Record<string, unknown>;

    console.log("[Google OAuth] Updating Clerk metadata for user:", userId);
    await clerk.users.updateUserMetadata(userId, {
      publicMetadata: {
        ...currentPublic,
        googleConnected: true,
      },
      privateMetadata: {
        googleAccessToken: tokens.access_token,
        googleRefreshToken: tokens.refresh_token,
        googleTokenExpiry: Date.now() + tokens.expires_in * 1000,
      },
    });

    console.log("[Google OAuth] Metadata updated, redirecting to /settings?google=connected");
    return NextResponse.redirect(new URL("/settings?google=connected", req.url));
  } catch (err) {
    console.error("[Google OAuth] Unhandled error:", err);
    return NextResponse.redirect(new URL("/settings?error=server_error", req.url));
  }
}
