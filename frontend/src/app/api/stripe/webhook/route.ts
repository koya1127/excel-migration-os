import { NextRequest, NextResponse } from "next/server";
import { getStripe } from "@/lib/stripe";
import { clerkClient } from "@clerk/nextjs/server";

export async function POST(req: NextRequest) {
  const stripe = getStripe();
  const body = await req.text();
  const sig = req.headers.get("stripe-signature");
  const webhookSecret = process.env.STRIPE_WEBHOOK_SECRET;

  if (!sig || !webhookSecret) {
    console.error("[Stripe Webhook] Missing signature or webhook secret");
    return NextResponse.json({ received: true }, { status: 200 });
  }

  let event;
  try {
    event = stripe.webhooks.constructEvent(body, sig, webhookSecret);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    console.error("[Stripe Webhook] Signature verification failed:", message);
    return NextResponse.json({ error: "Invalid signature" }, { status: 400 });
  }

  const clerk = await clerkClient();

  if (event.type === "checkout.session.completed") {
    const session = event.data.object;
    const userId = session.metadata?.userId;
    const customerId = session.customer as string;

    if (userId) {
      const user = await clerk.users.getUser(userId);
      const currentMeta = (user.publicMetadata || {}) as Record<string, unknown>;

      await clerk.users.updateUserMetadata(userId, {
        publicMetadata: {
          ...currentMeta,
          stripeCustomerId: customerId,
          stripeSubscriptionId: session.subscription as string,
          subscriptionStatus: "active",
        },
      });

      // Store userId in Stripe Customer metadata for reverse lookup in future webhooks
      await stripe.customers.update(customerId, {
        metadata: { clerkUserId: userId },
      });
    }
  }

  // Handle subscription lifecycle changes
  if (event.type === "customer.subscription.updated" || event.type === "customer.subscription.deleted") {
    const subscription = event.data.object;
    const customerId = subscription.customer as string;

    let newStatus: string;
    if (event.type === "customer.subscription.deleted") {
      newStatus = "canceled";
    } else {
      const stripeStatus = subscription.status;
      if (stripeStatus === "active" || stripeStatus === "trialing") {
        newStatus = "active";
      } else if (stripeStatus === "past_due") {
        newStatus = "past_due";
      } else if (stripeStatus === "unpaid") {
        newStatus = "unpaid";
      } else if (stripeStatus === "paused") {
        newStatus = "paused";
      } else {
        newStatus = "canceled";
      }
    }

    const userId = await findClerkUserByCustomerId(stripe, clerk, customerId);
    if (userId) {
      const user = await clerk.users.getUser(userId);
      const currentMeta = (user.publicMetadata || {}) as Record<string, unknown>;
      await clerk.users.updateUserMetadata(userId, {
        publicMetadata: {
          ...currentMeta,
          subscriptionStatus: newStatus,
          ...(newStatus === "canceled" ? { stripeSubscriptionId: null } : {}),
        },
      });
    }
  }

  // Handle payment failures
  if (event.type === "invoice.payment_failed") {
    const invoice = event.data.object;
    const customerId = invoice.customer as string;

    const userId = await findClerkUserByCustomerId(stripe, clerk, customerId);
    if (userId) {
      const user = await clerk.users.getUser(userId);
      const currentMeta = (user.publicMetadata || {}) as Record<string, unknown>;
      if (currentMeta.subscriptionStatus === "active") {
        await clerk.users.updateUserMetadata(userId, {
          publicMetadata: {
            ...currentMeta,
            subscriptionStatus: "past_due",
          },
        });
      }
    }
  }

  return NextResponse.json({ received: true });
}

/**
 * Find Clerk userId from Stripe Customer metadata (O(1) lookup).
 * Falls back to Clerk search if metadata not set (for legacy customers).
 */
async function findClerkUserByCustomerId(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  stripe: any,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  clerk: any,
  customerId: string
): Promise<string | null> {
  try {
    // Fast path: read userId from Stripe Customer metadata
    const customer = await stripe.customers.retrieve(customerId);
    if (customer.metadata?.clerkUserId) {
      return customer.metadata.clerkUserId;
    }

    // Slow fallback for legacy customers without metadata
    console.warn(`[Stripe Webhook] Customer ${customerId} missing clerkUserId metadata, falling back to user scan`);
    let offset = 0;
    const limit = 100;
    const maxPages = 5; // Cap at 500 users to prevent timeout

    for (let page = 0; page < maxPages; page++) {
      const users = await clerk.users.getUserList({ limit, offset });
      if (users.data.length === 0) break;

      const user = users.data.find(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (u: any) => (u.publicMetadata as Record<string, unknown>)?.stripeCustomerId === customerId
      );

      if (user) {
        // Backfill Stripe metadata for future lookups
        await stripe.customers.update(customerId, {
          metadata: { clerkUserId: user.id },
        });
        return user.id;
      }

      if (users.data.length < limit) break;
      offset += limit;
    }
  } catch (e) {
    console.error("findClerkUserByCustomerId error:", e);
  }

  return null;
}
