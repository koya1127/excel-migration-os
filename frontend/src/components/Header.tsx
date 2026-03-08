"use client";

import Link from "next/link";
import {
  SignInButton,
  UserButton,
  useUser,
} from "@clerk/nextjs";

export function Header() {
  const { isSignedIn, isLoaded } = useUser();

  return (
    <header className="border-b border-gray-200 bg-white">
      <div className="mx-auto flex max-w-7xl items-center justify-between px-6 py-4">
        <Link href="/" className="text-xl font-bold text-gray-900">
          Excel Migration OS
        </Link>
        <nav className="flex items-center gap-6">
          <Link
            href="/"
            className="text-sm font-medium text-gray-600 hover:text-gray-900 transition-colors"
          >
            Home
          </Link>
          <Link
            href="/pricing"
            className="text-sm font-medium text-gray-600 hover:text-gray-900 transition-colors"
          >
            Pricing
          </Link>
          {isLoaded && isSignedIn ? (
            <>
              <Link
                href="/scan"
                className="text-sm font-medium text-gray-600 hover:text-gray-900 transition-colors"
              >
                Scan
              </Link>
              <Link
                href="/extract"
                className="text-sm font-medium text-gray-600 hover:text-gray-900 transition-colors"
              >
                Extract
              </Link>
              <Link
                href="/convert"
                className="text-sm font-medium text-gray-600 hover:text-gray-900 transition-colors"
              >
                Convert
              </Link>
              <Link
                href="/upload"
                className="text-sm font-medium text-gray-600 hover:text-gray-900 transition-colors"
              >
                Upload
              </Link>
              <Link
                href="/migrate"
                className="rounded-lg bg-blue-600 px-3 py-1.5 text-sm font-medium text-white hover:bg-blue-500 transition-colors"
              >
                Migrate
              </Link>
              <Link
                href="/settings"
                className="text-sm font-medium text-gray-600 hover:text-gray-900 transition-colors"
              >
                Settings
              </Link>
              <UserButton />
            </>
          ) : isLoaded ? (
            <SignInButton mode="modal">
              <button className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-500 transition-colors">
                ログイン
              </button>
            </SignInButton>
          ) : null}
        </nav>
      </div>
    </header>
  );
}
