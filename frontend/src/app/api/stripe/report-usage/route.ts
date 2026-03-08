import { auth } from "@clerk/nextjs/server";
import { clerkClient } from "@clerk/nextjs/server";
import { NextRequest, NextResponse } from "next/server";
import { getStripe } from "@/lib/stripe";

export async function POST(req: NextRequest) {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { inputTokens, outputTokens } = await req.json();
  const totalTokens = (inputTokens || 0) + (outputTokens || 0);

  if (totalTokens <= 0) {
    return NextResponse.json({ error: "No tokens to report" }, { status: 400 });
  }

  // Get Stripe customer ID from Clerk metadata
  const clerk = await clerkClient();
  const user = await clerk.users.getUser(userId);
  const meta = (user.publicMetadata || {}) as Record<string, unknown>;
  const stripeCustomerId = meta.stripeCustomerId as string | undefined;

  if (!stripeCustomerId) {
    return NextResponse.json({ error: "No active subscription" }, { status: 400 });
  }

  const meterEventName = process.env.STRIPE_METER_EVENT_NAME || "ai_tokens";

  try {
    const stripe = getStripe();

    // Report usage via Stripe Billing Meter Events API
    await stripe.billing.meterEvents.create({
      event_name: meterEventName,
      payload: {
        value: String(totalTokens),
        stripe_customer_id: stripeCustomerId,
      },
      timestamp: Math.floor(Date.now() / 1000),
    });

    // Also update Clerk metadata for display in settings
    const currentTokens = (meta.tokensUsedThisMonth as number) || 0;
    const currentFiles = (meta.filesConvertedThisMonth as number) || 0;
    await clerk.users.updateUserMetadata(userId, {
      publicMetadata: {
        ...meta,
        tokensUsedThisMonth: currentTokens + totalTokens,
        filesConvertedThisMonth: currentFiles + 1,
      },
    });

    return NextResponse.json({ reported: totalTokens });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return NextResponse.json({ error: `Failed to report usage: ${message}` }, { status: 500 });
  }
}
