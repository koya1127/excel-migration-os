"use client";

import { useUser } from "@clerk/nextjs";
import { Suspense, useState, useEffect } from "react";
import { useSearchParams } from "next/navigation";

interface UsageData {
  filesConverted: number;
  tokensUsed: number;
  estimatedCost: number;
}

function SettingsContent() {
  const { user, isLoaded } = useUser();
  const [usage, setUsage] = useState<UsageData | null>(null);
  const [portalLoading, setPortalLoading] = useState(false);
  const [googleStatus, setGoogleStatus] = useState<string | null>(null);
  const searchParams = useSearchParams();

  const meta = (user?.publicMetadata || {}) as Record<string, unknown>;
  const subscriptionStatus = meta.subscriptionStatus as string | undefined;
  const hasSubscription = subscriptionStatus === "active";
  const googleConnected = !!meta.googleConnected;

  // Reload user metadata after Google OAuth callback redirect
  useEffect(() => {
    const google = searchParams.get("google");
    const error = searchParams.get("error");

    if (error) {
      setGoogleStatus(
        error === "token_exchange"
          ? "Google認証に失敗しました。もう一度お試しください。"
          : error === "missing_params"
            ? "認証パラメータが不足しています。"
            : `エラー: ${error}`
      );
      return;
    }

    if (google === "connected" && user) {
      user.reload().then(() => {
        setGoogleStatus("Google アカウントを連携しました！");
      });
    }
  }, [searchParams, user]);

  useEffect(() => {
    if (!isLoaded || !user) return;
    fetch("/api/stripe/usage")
      .then((r) => r.json())
      .then(setUsage)
      .catch(() => {});
  }, [isLoaded, user]);

  async function handleGoogleConnect() {
    const res = await fetch("/api/auth/google");
    const data = await res.json();
    if (data.url) {
      try {
        const url = new URL(data.url);
        if (url.hostname.endsWith("google.com")) {
          window.location.href = data.url;
        }
      } catch { /* invalid URL */ }
    }
  }

  async function handleManageBilling() {
    setPortalLoading(true);
    try {
      const res = await fetch("/api/stripe/portal", { method: "POST" });
      const data = await res.json();
      if (data.url) {
        try {
          const url = new URL(data.url);
          if (url.hostname.endsWith("stripe.com")) {
            window.location.href = data.url;
          }
        } catch { /* invalid URL */ }
      }
    } catch {
      alert("エラーが発生しました");
    } finally {
      setPortalLoading(false);
    }
  }

  if (!isLoaded) {
    return (
      <div className="flex min-h-[60vh] items-center justify-center">
        <p className="text-gray-500">読み込み中...</p>
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-3xl px-6 py-12">
      <h1 className="text-2xl font-bold text-gray-900 mb-8">設定</h1>

      {/* 課金ステータス */}
      <section className="mb-8 rounded-xl border border-gray-200 p-6">
        <h2 className="text-lg font-semibold text-gray-900 mb-4">課金</h2>
        <div className="flex items-center justify-between">
          <div>
            <p className="text-sm text-gray-500">ステータス</p>
            <p className="text-xl font-bold text-gray-900">
              {hasSubscription ? "従量課金（有効）" : "未登録"}
            </p>
            {hasSubscription && (
              <span className="mt-1 inline-block rounded-full px-2.5 py-0.5 text-xs font-medium bg-green-100 text-green-700">
                ¥3 / 1Kトークン
              </span>
            )}
          </div>
          <div className="flex gap-3">
            {!hasSubscription && (
              <a
                href="/pricing"
                className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-500 transition-colors"
              >
                従量課金を開始
              </a>
            )}
            {hasSubscription && (
              <button
                onClick={handleManageBilling}
                disabled={portalLoading}
                className="rounded-lg bg-gray-900 px-4 py-2 text-sm font-medium text-white hover:bg-gray-800 transition-colors disabled:opacity-50"
              >
                {portalLoading ? "読込中..." : "請求管理"}
              </button>
            )}
          </div>
        </div>
      </section>

      {/* 今月の利用状況 */}
      <section className="mb-8 rounded-xl border border-gray-200 p-6">
        <h2 className="text-lg font-semibold text-gray-900 mb-4">今月の利用状況</h2>
        {usage ? (
          <div className="grid grid-cols-3 gap-4">
            <div className="rounded-lg bg-gray-50 p-4">
              <p className="text-sm text-gray-500">変換ファイル数</p>
              <p className="text-2xl font-bold text-gray-900">{usage.filesConverted}</p>
            </div>
            <div className="rounded-lg bg-gray-50 p-4">
              <p className="text-sm text-gray-500">API トークン使用量</p>
              <p className="text-2xl font-bold text-gray-900">
                {usage.tokensUsed.toLocaleString()}
              </p>
            </div>
            <div className="rounded-lg bg-gray-50 p-4">
              <p className="text-sm text-gray-500">API従量課金（税抜）</p>
              <p className="text-2xl font-bold text-gray-900">
                ¥{usage.estimatedCost.toLocaleString()}
              </p>
            </div>
          </div>
        ) : (
          <p className="text-sm text-gray-500">データを取得中...</p>
        )}
      </section>

      {/* Google連携 */}
      <section className="mb-8 rounded-xl border border-gray-200 p-6">
        <h2 className="text-lg font-semibold text-gray-900 mb-4">Google アカウント連携</h2>
        <p className="text-sm text-gray-600 mb-4">
          Google Drive へのアップロードや Apps Script のデプロイに必要です。
          ログイン認証とは別に、ファイル操作の権限を許可する必要があります。
          変換後のファイルはあなた自身の Google アカウントに保存されます。
        </p>
        {googleStatus && (
          <div className={`mb-4 rounded-lg px-4 py-2 text-sm ${
            googleStatus.includes("失敗") || googleStatus.includes("エラー")
              ? "bg-red-50 text-red-700"
              : "bg-green-50 text-green-700"
          }`}>
            {googleStatus}
          </div>
        )}
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div
              className={`h-3 w-3 rounded-full ${
                googleConnected ? "bg-green-500" : "bg-gray-300"
              }`}
            />
            <span className="text-sm text-gray-700">
              {googleConnected ? "接続済み" : "未接続"}
            </span>
          </div>
          <button
            onClick={handleGoogleConnect}
            className={`rounded-lg px-4 py-2 text-sm font-medium transition-colors ${
              googleConnected
                ? "border border-gray-300 text-gray-700 hover:bg-gray-50"
                : "bg-blue-600 text-white hover:bg-blue-500"
            }`}
          >
            {googleConnected ? "再連携" : "Google アカウントを連携"}
          </button>
        </div>
      </section>

      {/* アカウント情報 */}
      <section className="rounded-xl border border-gray-200 p-6">
        <h2 className="text-lg font-semibold text-gray-900 mb-4">アカウント</h2>
        <div className="space-y-2 text-sm">
          <div className="flex justify-between">
            <span className="text-gray-500">メールアドレス</span>
            <span className="text-gray-900">
              {user?.emailAddresses?.[0]?.emailAddress || "-"}
            </span>
          </div>
          <div className="flex justify-between">
            <span className="text-gray-500">ユーザーID</span>
            <span className="font-mono text-gray-500 text-xs">{user?.id || "-"}</span>
          </div>
        </div>
      </section>
    </div>
  );
}

export default function SettingsPage() {
  return (
    <Suspense fallback={
      <div className="flex min-h-[60vh] items-center justify-center">
        <p className="text-gray-500">読み込み中...</p>
      </div>
    }>
      <SettingsContent />
    </Suspense>
  );
}
