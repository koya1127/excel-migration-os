import { NextRequest, NextResponse } from "next/server";
import { getStripe } from "@/lib/stripe";
import { clerkClient } from "@clerk/nextjs/server";

export async function POST(req: NextRequest) {
  const stripe = getStripe();
  const body = await req.text();
  const sig = req.headers.get("stripe-signature");
  const webhookSecret = process.env.STRIPE_WEBHOOK_SECRET;

  if (!sig || !webhookSecret) {
    return NextResponse.json({ error: "Missing signature or secret" }, { status: 400 });
  }

  let event;
  try {
    event = stripe.webhooks.constructEvent(body, sig, webhookSecret);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return NextResponse.json({ error: `Invalid signature: ${message}` }, { status: 400 });
  }

  if (event.type === "checkout.session.completed") {
    const session = event.data.object;
    const userId = session.metadata?.userId;
    const planId = session.metadata?.planId;

    if (userId && planId) {
      const clerk = await clerkClient();
      const user = await clerk.users.getUser(userId);
      const currentMeta = (user.publicMetadata || {}) as Record<string, unknown>;

      await clerk.users.updateUserMetadata(userId, {
        publicMetadata: {
          ...currentMeta,
          planId,
          stripeCustomerId: session.customer as string,
          stripeSubscriptionId: session.subscription as string,
          subscriptionStatus: "active",
        },
      });
    }
  }

  if (event.type === "customer.subscription.deleted") {
    const subscription = event.data.object;
    const customerId = subscription.customer as string;

    // Find user by stripeCustomerId in Clerk
    const clerk = await clerkClient();
    const users = await clerk.users.getUserList({ limit: 100 });
    const user = users.data.find(
      (u) => (u.publicMetadata as Record<string, unknown>)?.stripeCustomerId === customerId
    );

    if (user) {
      const currentMeta = (user.publicMetadata || {}) as Record<string, unknown>;
      await clerk.users.updateUserMetadata(user.id, {
        publicMetadata: {
          ...currentMeta,
          planId: "free",
          subscriptionStatus: "canceled",
          stripeSubscriptionId: null,
        },
      });
    }
  }

  return NextResponse.json({ received: true });
}
