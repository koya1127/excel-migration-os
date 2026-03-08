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

  // Usage tracked in Clerk metadata (updated by backend on each conversion)
  const filesConverted = (meta.filesConvertedThisMonth as number) || 0;
  const tokensUsed = (meta.tokensUsedThisMonth as number) || 0;

  // Claude API pricing: ~¥0.5 per 1K tokens (approximate)
  const estimatedCost = Math.ceil(tokensUsed / 1000 * 0.5);

  return NextResponse.json({
    filesConverted,
    tokensUsed,
    estimatedCost,
  });
}
