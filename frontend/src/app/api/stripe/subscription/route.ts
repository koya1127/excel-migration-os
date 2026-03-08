import { auth } from "@clerk/nextjs/server";
import { clerkClient } from "@clerk/nextjs/server";
import { NextResponse } from "next/server";

export async function GET() {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const clerk = await clerkClient();
  const user = await clerk.users.getUser(userId);
  const meta = (user.publicMetadata || {}) as Record<string, unknown>;

  const subscriptionStatus = meta.subscriptionStatus as string | undefined;
  const hasSubscription = subscriptionStatus === "active";

  return NextResponse.json({ hasSubscription, subscriptionStatus });
}
