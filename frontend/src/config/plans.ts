// Metered pricing for AI conversion usage (per 1K tokens)
export const METERED_PRICE_ID = process.env.NEXT_PUBLIC_STRIPE_PRICE_METERED || "";
export const TOKEN_UNIT_PRICE_YEN = 3; // ¥3 per 1K tokens
