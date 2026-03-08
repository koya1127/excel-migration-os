"use client";

import { useState } from "react";
import { useUser } from "@clerk/nextjs";
import { SignInButton } from "@clerk/nextjs";
import { PLANS } from "@/config/plans";

export default function PricingPage() {
  const { user, isSignedIn, isLoaded } = useUser();
  const [loading, setLoading] = useState<string | null>(null);

  const currentPlan = (user?.publicMetadata as Record<string, unknown>)?.planId as string | undefined;

  async function handleSubscribe(planId: string) {
    setLoading(planId);
    try {
      const res = await fetch("/api/stripe/checkout", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ planId }),
      });
      const data = await res.json();
      if (data.url) {
        window.location.href = data.url;
      }
    } catch {
      alert("エラーが発生しました。もう一度お試しください。");
    } finally {
      setLoading(null);
    }
  }

  return (
    <div className="mx-auto max-w-6xl px-6 py-16">
      <div className="text-center mb-12">
        <h1 className="text-3xl font-bold text-gray-900">料金プラン</h1>
        <p className="mt-3 text-gray-600">
          Excel から Google Sheets への移行を、あなたのペースで。
        </p>
      </div>

      <div className="grid gap-6 md:grid-cols-2 lg:grid-cols-4">
        {PLANS.map((plan) => {
          const isCurrent = currentPlan === plan.id;
          const isFree = plan.price === 0;

          return (
            <div
              key={plan.id}
              className={`relative rounded-2xl border p-6 flex flex-col ${
                plan.recommended
                  ? "border-blue-500 ring-2 ring-blue-500"
                  : "border-gray-200"
              }`}
            >
              {plan.recommended && (
                <span className="absolute -top-3 left-1/2 -translate-x-1/2 rounded-full bg-blue-600 px-3 py-0.5 text-xs font-medium text-white">
                  おすすめ
                </span>
              )}

              <h2 className="text-lg font-semibold text-gray-900">{plan.name}</h2>
              <div className="mt-4">
                <span className="text-3xl font-bold text-gray-900">
                  {isFree ? "¥0" : `¥${plan.price.toLocaleString()}`}
                </span>
                {!isFree && (
                  <span className="text-sm text-gray-500">/月</span>
                )}
              </div>

              <ul className="mt-6 flex-1 space-y-3">
                {plan.features.map((feature) => (
                  <li key={feature} className="flex items-start gap-2 text-sm text-gray-600">
                    <span className="mt-0.5 text-green-500 font-bold">✓</span>
                    {feature}
                  </li>
                ))}
              </ul>

              <div className="mt-6">
                {!isLoaded ? (
                  <div className="h-10" />
                ) : isCurrent ? (
                  <button
                    disabled
                    className="w-full rounded-lg bg-gray-100 px-4 py-2.5 text-sm font-medium text-gray-400"
                  >
                    現在のプラン
                  </button>
                ) : isFree ? (
                  !isSignedIn ? (
                    <SignInButton mode="modal">
                      <button className="w-full rounded-lg bg-gray-900 px-4 py-2.5 text-sm font-medium text-white hover:bg-gray-800 transition-colors">
                        無料で始める
                      </button>
                    </SignInButton>
                  ) : (
                    <button
                      disabled
                      className="w-full rounded-lg bg-gray-100 px-4 py-2.5 text-sm font-medium text-gray-400"
                    >
                      利用中
                    </button>
                  )
                ) : !isSignedIn ? (
                  <SignInButton mode="modal">
                    <button
                      className={`w-full rounded-lg px-4 py-2.5 text-sm font-medium text-white transition-colors ${
                        plan.recommended
                          ? "bg-blue-600 hover:bg-blue-500"
                          : "bg-gray-900 hover:bg-gray-800"
                      }`}
                    >
                      ログインして申し込み
                    </button>
                  </SignInButton>
                ) : (
                  <button
                    onClick={() => handleSubscribe(plan.id)}
                    disabled={loading === plan.id}
                    className={`w-full rounded-lg px-4 py-2.5 text-sm font-medium text-white transition-colors ${
                      plan.recommended
                        ? "bg-blue-600 hover:bg-blue-500"
                        : "bg-gray-900 hover:bg-gray-800"
                    } disabled:opacity-50`}
                  >
                    {loading === plan.id ? "処理中..." : "申し込む"}
                  </button>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
