import { auth } from "@clerk/nextjs/server";
import { NextRequest, NextResponse } from "next/server";
import { getStripe } from "@/lib/stripe";
import { getPlanById, METERED_PRICE_ID } from "@/config/plans";

export async function POST(req: NextRequest) {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { planId } = await req.json();
  const plan = getPlanById(planId);
  if (!plan || !plan.stripePriceId) {
    return NextResponse.json({ error: "Invalid plan" }, { status: 400 });
  }

  const stripe = getStripe();
  const origin = req.headers.get("origin") || "http://localhost:3000";

  // Build line items: base plan + metered AI usage
  const lineItems: { price: string; quantity?: number }[] = [
    { price: plan.stripePriceId, quantity: 1 },
  ];

  // Add metered price for AI token usage (if configured)
  if (METERED_PRICE_ID) {
    lineItems.push({ price: METERED_PRICE_ID });
  }

  const session = await stripe.checkout.sessions.create({
    mode: "subscription",
    line_items: lineItems,
    success_url: `${origin}/checkout/complete?session_id={CHECKOUT_SESSION_ID}`,
    cancel_url: `${origin}/pricing?canceled=true`,
    metadata: { userId, planId: plan.id },
  });

  return NextResponse.json({ url: session.url });
}
