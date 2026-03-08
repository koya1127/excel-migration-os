import Link from "next/link";

export default function Home() {
  return (
    <div className="flex min-h-[calc(100vh-65px)] items-center justify-center">
      <div className="mx-auto max-w-3xl text-center px-6">
        <h1 className="text-5xl font-bold tracking-tight text-gray-900">
          Excel Migration OS
        </h1>
        <p className="mt-4 text-lg text-gray-600">
          Excel/VBA から Google Sheets/Apps Script への移行を自動化
        </p>
        <p className="mt-2 text-sm text-gray-500">
          ファイルをスキャンして移行リスクを分析し、VBA を Apps Script に変換します
        </p>
        <div className="mt-10 flex justify-center gap-4">
          <Link
            href="/migrate"
            className="rounded-lg bg-gradient-to-r from-blue-600 to-indigo-600 px-8 py-3 text-sm font-semibold text-white shadow-md hover:from-blue-500 hover:to-indigo-500 transition-all"
          >
            一括移行を開始する
          </Link>
          <Link
            href="/scan"
            className="rounded-lg bg-white px-8 py-3 text-sm font-semibold text-gray-700 border border-gray-300 shadow-sm hover:bg-gray-50 transition-colors"
          >
            スキャンから始める
          </Link>
        </div>

        <div className="mt-16 grid grid-cols-1 gap-6 sm:grid-cols-5">
          <Link href="/scan" className="rounded-xl border border-gray-200 bg-white p-6 text-left shadow-sm hover:shadow-md hover:border-gray-300 transition-all">
            <div className="text-2xl mb-3 text-blue-600 font-bold">1</div>
            <h3 className="font-semibold text-gray-900">スキャン</h3>
            <p className="mt-1 text-sm text-gray-500">
              Excel ファイルを分析しリスクスコアを算出
            </p>
          </Link>
          <Link href="/extract" className="rounded-xl border border-gray-200 bg-white p-6 text-left shadow-sm hover:shadow-md hover:border-gray-300 transition-all">
            <div className="text-2xl mb-3 text-purple-600 font-bold">2</div>
            <h3 className="font-semibold text-gray-900">抽出</h3>
            <p className="mt-1 text-sm text-gray-500">
              VBA モジュールとコントロールを抽出
            </p>
          </Link>
          <Link href="/convert" className="rounded-xl border border-gray-200 bg-white p-6 text-left shadow-sm hover:shadow-md hover:border-gray-300 transition-all">
            <div className="text-2xl mb-3 text-amber-600 font-bold">3</div>
            <h3 className="font-semibold text-gray-900">変換</h3>
            <p className="mt-1 text-sm text-gray-500">
              VBA を Google Apps Script に自動変換
            </p>
          </Link>
          <Link href="/upload" className="rounded-xl border border-gray-200 bg-white p-6 text-left shadow-sm hover:shadow-md hover:border-gray-300 transition-all">
            <div className="text-2xl mb-3 text-green-600 font-bold">4</div>
            <h3 className="font-semibold text-gray-900">アップロード</h3>
            <p className="mt-1 text-sm text-gray-500">
              Google Drive にファイルをアップロード
            </p>
          </Link>
          <Link href="/migrate" className="rounded-xl border-2 border-blue-200 bg-blue-50 p-6 text-left shadow-sm hover:shadow-md hover:border-blue-300 transition-all">
            <div className="text-2xl mb-3 text-blue-600 font-bold">ALL</div>
            <h3 className="font-semibold text-blue-900">一括移行</h3>
            <p className="mt-1 text-sm text-blue-600">
              全ステップを一括で自動実行
            </p>
          </Link>
        </div>
      </div>
    </div>
  );
}
