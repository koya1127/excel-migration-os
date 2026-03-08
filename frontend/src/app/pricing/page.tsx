"use client";

import { useState } from "react";
import { useUser } from "@clerk/nextjs";
import { SignInButton } from "@clerk/nextjs";
import { TOKEN_UNIT_PRICE_YEN } from "@/config/plans";

export default function PricingPage() {
  const { user, isSignedIn, isLoaded } = useUser();
  const [loading, setLoading] = useState(false);

  const meta = (user?.publicMetadata || {}) as Record<string, unknown>;
  const hasSubscription = meta.subscriptionStatus === "active";

  async function handleSubscribe() {
    setLoading(true);
    try {
      const res = await fetch("/api/stripe/checkout", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({}),
      });
      const data = await res.json();
      if (data.url) {
        window.location.href = data.url;
      }
    } catch {
      alert("エラーが発生しました。もう一度お試しください。");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="mx-auto max-w-4xl px-6 py-16">
      <div className="text-center mb-12">
        <h1 className="text-3xl font-bold text-gray-900">料金</h1>
        <p className="mt-3 text-gray-600">
          使った分だけ。シンプルな従量課金。
        </p>
      </div>

      <div className="mx-auto max-w-lg">
        {/* Free tier */}
        <div className="rounded-2xl border border-gray-200 p-8 mb-6">
          <h2 className="text-lg font-semibold text-gray-900">無料機能</h2>
          <div className="mt-4">
            <span className="text-3xl font-bold text-gray-900">¥0</span>
          </div>
          <ul className="mt-6 space-y-3">
            <li className="flex items-start gap-2 text-sm text-gray-600">
              <span className="mt-0.5 text-green-500 font-bold">✓</span>
              Scan（ファイル分析）無制限
            </li>
            <li className="flex items-start gap-2 text-sm text-gray-600">
              <span className="mt-0.5 text-green-500 font-bold">✓</span>
              リスクスコア・互換性レポート
            </li>
            <li className="flex items-start gap-2 text-sm text-gray-600">
              <span className="mt-0.5 text-green-500 font-bold">✓</span>
              VBA抽出・コード閲覧
            </li>
          </ul>
        </div>

        {/* Metered tier */}
        <div className="rounded-2xl border-2 border-blue-500 ring-2 ring-blue-500 p-8 relative">
          <span className="absolute -top-3 left-1/2 -translate-x-1/2 rounded-full bg-blue-600 px-3 py-0.5 text-xs font-medium text-white">
            従量課金
          </span>
          <h2 className="text-lg font-semibold text-gray-900">AI変換プラン</h2>
          <div className="mt-4">
            <span className="text-3xl font-bold text-gray-900">¥{TOKEN_UNIT_PRICE_YEN}</span>
            <span className="text-sm text-gray-500"> / 1Kトークン</span>
          </div>
          <p className="mt-2 text-sm text-gray-500">
            使った分だけ月末にお支払い。最低料金なし。
          </p>
          <ul className="mt-6 space-y-3">
            <li className="flex items-start gap-2 text-sm text-gray-600">
              <span className="mt-0.5 text-green-500 font-bold">✓</span>
              無料機能すべて
            </li>
            <li className="flex items-start gap-2 text-sm text-gray-600">
              <span className="mt-0.5 text-green-500 font-bold">✓</span>
              VBA → GAS AI自動変換
            </li>
            <li className="flex items-start gap-2 text-sm text-gray-600">
              <span className="mt-0.5 text-green-500 font-bold">✓</span>
              Google Drive アップロード
            </li>
            <li className="flex items-start gap-2 text-sm text-gray-600">
              <span className="mt-0.5 text-green-500 font-bold">✓</span>
              GAS自動デプロイ
            </li>
            <li className="flex items-start gap-2 text-sm text-gray-600">
              <span className="mt-0.5 text-green-500 font-bold">✓</span>
              一括マイグレーション
            </li>
          </ul>

          <div className="mt-8">
            {!isLoaded ? (
              <div className="h-10" />
            ) : hasSubscription ? (
              <button
                disabled
                className="w-full rounded-lg bg-gray-100 px-4 py-2.5 text-sm font-medium text-gray-400"
              >
                利用中
              </button>
            ) : !isSignedIn ? (
              <SignInButton mode="modal">
                <button className="w-full rounded-lg bg-blue-600 px-4 py-2.5 text-sm font-medium text-white hover:bg-blue-500 transition-colors">
                  ログインして申し込み
                </button>
              </SignInButton>
            ) : (
              <button
                onClick={handleSubscribe}
                disabled={loading}
                className="w-full rounded-lg bg-blue-600 px-4 py-2.5 text-sm font-medium text-white hover:bg-blue-500 transition-colors disabled:opacity-50"
              >
                {loading ? "処理中..." : "従量課金を開始する"}
              </button>
            )}
          </div>
        </div>

        {/* Cost examples */}
        <div className="mt-8 rounded-xl bg-gray-50 p-6">
          <h3 className="text-sm font-semibold text-gray-900 mb-3">料金の目安</h3>
          <div className="space-y-2 text-sm text-gray-600">
            <div className="flex justify-between">
              <span>小規模VBAモジュール（〜500行）</span>
              <span className="font-medium text-gray-900">約¥30〜60</span>
            </div>
            <div className="flex justify-between">
              <span>中規模VBAモジュール（〜2000行）</span>
              <span className="font-medium text-gray-900">約¥60〜120</span>
            </div>
            <div className="flex justify-between">
              <span>大規模マクロブック一括変換</span>
              <span className="font-medium text-gray-900">約¥200〜500</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
