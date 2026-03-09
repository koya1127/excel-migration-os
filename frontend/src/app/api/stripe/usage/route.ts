import { auth } from "@clerk/nextjs/server";
import { clerkClient } from "@clerk/nextjs/server";
import { NextResponse } from "next/server";
import { getStripe } from "@/lib/stripe";

const STRIPE_METER_ID = process.env.STRIPE_METER_ID || "";

export async function GET() {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const clerk = await clerkClient();
  const user = await clerk.users.getUser(userId);
  const meta = (user.publicMetadata || {}) as Record<string, unknown>;
  const customerId = meta.stripeCustomerId as string | undefined;

  if (!customerId || !STRIPE_METER_ID) {
    return NextResponse.json({ filesConverted: 0, tokensUsed: 0, estimatedCost: 0 });
  }

  try {
    const stripe = getStripe();
    const now = new Date();
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

    const summaries = await stripe.billing.meters.listEventSummaries(
      STRIPE_METER_ID,
      {
        customer: customerId,
        start_time: Math.floor(startOfMonth.getTime() / 1000),
        end_time: Math.floor(now.getTime() / 1000),
      }
    );

    let tokensUsed = 0;
    for (const summary of summaries.data) {
      tokensUsed += summary.aggregated_value;
    }

    // ¥3 per 1K tokens (integer math to avoid floating point precision issues)
    const estimatedCost = Math.ceil((tokensUsed * 3) / 1000);

    return NextResponse.json({
      filesConverted: 0, // Not tracked per-file anymore
      tokensUsed,
      estimatedCost,
    });
  } catch (e) {
    console.error("Usage fetch error:", e);
    return NextResponse.json({ filesConverted: 0, tokensUsed: 0, estimatedCost: 0, error: "使用量データを取得できませんでした" });
  }
}
