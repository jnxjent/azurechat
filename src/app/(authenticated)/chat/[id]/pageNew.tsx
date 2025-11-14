// 最上部に追記
export const dynamic = 'force-dynamic';
export const fetchCache = 'force-no-store';
export const revalidate = 0;

import { unstable_noStore as noStore } from 'next/cache';

export default async function Home(props: HomeParams) {
  noStore(); // ← 関数冒頭に追加（実行時もキャッシュ断つ）
  // 既存コード…
}
