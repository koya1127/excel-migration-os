export interface Plan {
  id: string;
  name: string;
  price: number;
  fileLimit: number;
  features: string[];
  stripePriceId: string;
  recommended?: boolean;
}

export const PLANS: Plan[] = [
  {
    id: "free",
    name: "Free",
    price: 0,
    fileLimit: 0,
    features: [
      "Scan（ファイル分析）無制限",
      "リスクスコア・互換性レポート",
    ],
    stripePriceId: "",
  },
  {
    id: "starter",
    name: "Starter",
    price: 2980,
    fileLimit: 10,
    features: [
      "Freeの全機能",
      "月10ファイルまで変換",
      "VBA → GAS 自動変換",
      "Google Drive アップロード",
      "AI変換: ¥1/1Kトークン従量課金",
    ],
    stripePriceId: process.env.NEXT_PUBLIC_STRIPE_PRICE_STARTER || "",
  },
  {
    id: "pro",
    name: "Pro",
    price: 9800,
    fileLimit: 50,
    features: [
      "Starterの全機能",
      "月50ファイルまで変換",
      "GAS自動デプロイ",
      "一括マイグレーション",
      "AI変換: ¥1/1Kトークン従量課金",
    ],
    stripePriceId: process.env.NEXT_PUBLIC_STRIPE_PRICE_PRO || "",
    recommended: true,
  },
  {
    id: "business",
    name: "Business",
    price: 29800,
    fileLimit: -1,
    features: [
      "Proの全機能",
      "ファイル数無制限",
      "優先サポート",
      "チーム利用対応",
      "AI変換: ¥1/1Kトークン従量課金",
    ],
    stripePriceId: process.env.NEXT_PUBLIC_STRIPE_PRICE_BUSINESS || "",
  },
];

// Metered pricing for AI conversion usage (per 1K tokens)
export const METERED_PRICE_ID = process.env.NEXT_PUBLIC_STRIPE_PRICE_METERED || "";
export const TOKEN_UNIT_PRICE_YEN = 1; // ¥1 per 1K tokens

export function getPlanById(id: string): Plan | undefined {
  return PLANS.find((p) => p.id === id);
}

export function getPaidPlans(): Plan[] {
  return PLANS.filter((p) => p.price > 0);
}
