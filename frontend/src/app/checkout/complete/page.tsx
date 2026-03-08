"use client";

import Link from "next/link";
import { useUser } from "@clerk/nextjs";

export default function CheckoutCompletePage() {
  const { user, isLoaded } = useUser();
  const planId = (user?.publicMetadata as Record<string, unknown>)?.planId as string | undefined;

  if (!isLoaded) {
    return (
      <div className="flex min-h-[60vh] items-center justify-center">
        <p className="text-gray-500">読み込み中...</p>
      </div>
    );
  }

  return (
    <div className="flex min-h-[60vh] items-center justify-center">
      <div className="text-center max-w-md">
        <div className="mx-auto mb-6 flex h-16 w-16 items-center justify-center rounded-full bg-green-100">
          <span className="text-3xl">✓</span>
        </div>
        <h1 className="text-2xl font-bold text-gray-900">お申し込み完了</h1>
        <p className="mt-3 text-gray-600">
          {planId ? `${planId.charAt(0).toUpperCase() + planId.slice(1)} プラン` : "プラン"}
          のお申し込みが完了しました。すべての機能をご利用いただけます。
        </p>
        <div className="mt-8 flex gap-3 justify-center">
          <Link
            href="/migrate"
            className="rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-medium text-white hover:bg-blue-500 transition-colors"
          >
            移行を始める
          </Link>
          <Link
            href="/"
            className="rounded-lg border border-gray-300 px-6 py-2.5 text-sm font-medium text-gray-700 hover:bg-gray-50 transition-colors"
          >
            ホームに戻る
          </Link>
        </div>
      </div>
    </div>
  );
}
