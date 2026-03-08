import { auth } from "@clerk/nextjs/server";
import { clerkClient } from "@clerk/nextjs/server";
import { NextResponse } from "next/server";

const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || "";
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || "";

export async function GET() {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const clerk = await clerkClient();
  const user = await clerk.users.getUser(userId);
  const meta = (user.privateMetadata || {}) as Record<string, unknown>;

  const accessToken = meta.googleAccessToken as string | undefined;
  const refreshToken = meta.googleRefreshToken as string | undefined;
  const expiry = meta.googleTokenExpiry as number | undefined;

  if (!accessToken || !refreshToken) {
    return NextResponse.json({ error: "Google not connected" }, { status: 400 });
  }

  // If token is still valid (with 5 min buffer), return it
  if (expiry && Date.now() < expiry - 5 * 60 * 1000) {
    return NextResponse.json({ accessToken });
  }

  // Refresh the token
  const tokenRes = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: GOOGLE_CLIENT_ID,
      client_secret: GOOGLE_CLIENT_SECRET,
      refresh_token: refreshToken,
      grant_type: "refresh_token",
    }),
  });

  if (!tokenRes.ok) {
    return NextResponse.json({ error: "Token refresh failed" }, { status: 500 });
  }

  const tokens = await tokenRes.json();

  // Update stored token
  await clerk.users.updateUserMetadata(userId, {
    privateMetadata: {
      ...meta,
      googleAccessToken: tokens.access_token,
      googleTokenExpiry: Date.now() + tokens.expires_in * 1000,
    },
  });

  return NextResponse.json({ accessToken: tokens.access_token });
}
