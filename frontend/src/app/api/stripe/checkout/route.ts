import { auth } from "@clerk/nextjs/server";
import { clerkClient } from "@clerk/nextjs/server";
import { NextResponse } from "next/server";
import { getStripe } from "@/lib/stripe";
import { METERED_PRICE_ID } from "@/config/plans";

export async function POST() {
  const { userId } = await auth();
  if (!userId) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  if (!METERED_PRICE_ID) {
    return NextResponse.json({ error: "Metered price not configured" }, { status: 500 });
  }

  const stripe = getStripe();
  const appUrl = process.env.NEXT_PUBLIC_APP_URL;
  if (!appUrl) {
    return NextResponse.json({ error: "NEXT_PUBLIC_APP_URL is not configured" }, { status: 500 });
  }

  // Reuse existing Stripe customer to prevent duplicate customers
  const clerk = await clerkClient();
  const user = await clerk.users.getUser(userId);
  const meta = (user.publicMetadata || {}) as Record<string, unknown>;
  const existingCustomerId = meta.stripeCustomerId as string | undefined;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sessionParams: Record<string, any> = {
    mode: "subscription",
    line_items: [{ price: METERED_PRICE_ID }],
    success_url: `${appUrl}/checkout/complete?session_id={CHECKOUT_SESSION_ID}`,
    cancel_url: `${appUrl}/pricing?canceled=true`,
    metadata: { userId },
  };

  if (existingCustomerId) {
    sessionParams.customer = existingCustomerId;
  }

  const session = await stripe.checkout.sessions.create(sessionParams);

  return NextResponse.json({ url: session.url });
}
