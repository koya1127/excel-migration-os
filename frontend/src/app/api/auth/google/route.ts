import { auth } from "@clerk/nextjs/server";
import { NextResponse } from "next/server";
import crypto from "crypto";

const GOOGLE_CLIENT_ID = (process.env.GOOGLE_CLIENT_ID || "").trim();
const GOOGLE_REDIRECT_URI = (process.env.GOOGLE_REDIRECT_URI || (process.env.NEXT_PUBLIC_APP_URL ? `${process.env.NEXT_PUBLIC_APP_URL}/api/auth/google/callback` : "")).trim();

const SCOPES = [
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/script.projects",
].join(" ");

/**
 * Create HMAC of nonce:userId to cryptographically bind them together.
 * This prevents an attacker from swapping the userId in the state parameter.
 */
function createStateHmac(nonce: string, userId: string): string {
  const secret = process.env.CLERK_SECRET_KEY;
  if (!secret) throw new Error("CLERK_SECRET_KEY is required for HMAC signing");
  return crypto.createHmac("sha256", secret).update(`${nonce}:${userId}`).digest("hex");
}

export { createStateHmac };

export async function GET() {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  if (!GOOGLE_REDIRECT_URI) {
    return NextResponse.json({ error: "GOOGLE_REDIRECT_URI or NEXT_PUBLIC_APP_URL is not configured" }, { status: 500 });
  }

  // Generate cryptographic random state to prevent CSRF
  const nonce = crypto.randomBytes(32).toString("hex");
  const hmac = createStateHmac(nonce, userId);
  // state = nonce:userId:hmac — hmac binds nonce to userId
  const state = `${nonce}:${userId}:${hmac}`;

  const params = new URLSearchParams({
    client_id: GOOGLE_CLIENT_ID,
    redirect_uri: GOOGLE_REDIRECT_URI,
    response_type: "code",
    scope: SCOPES,
    access_type: "offline",
    prompt: "consent",
    state,
  });

  const response = NextResponse.json({
    url: `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`,
  });

  // Store the nonce in a secure, httpOnly cookie for callback verification
  response.cookies.set("google_oauth_state", nonce, {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    maxAge: 600, // 10 minutes
    path: "/api/auth/google",
  });

  return response;
}
