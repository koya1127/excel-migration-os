import { auth } from "@clerk/nextjs/server";
import { NextRequest, NextResponse } from "next/server";
import { getStripe } from "@/lib/stripe";
import { METERED_PRICE_ID } from "@/config/plans";

export async function POST(req: NextRequest) {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  if (!METERED_PRICE_ID) {
    return NextResponse.json({ error: "Metered price not configured" }, { status: 500 });
  }

  const stripe = getStripe();
  const origin = req.headers.get("origin") || "http://localhost:3000";

  const session = await stripe.checkout.sessions.create({
    mode: "subscription",
    line_items: [{ price: METERED_PRICE_ID }],
    success_url: `${origin}/checkout/complete?session_id={CHECKOUT_SESSION_ID}`,
    cancel_url: `${origin}/pricing?canceled=true`,
    metadata: { userId },
  });

  return NextResponse.json({ url: session.url });
}
