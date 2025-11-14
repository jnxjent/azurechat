// server-only: このモジュールを読み込んだツリー内では fetch 既定を no-store にする
import "server-only";

// すでにパッチ済みでも二重適用しないようにガード
if (!(global as any).__NO_STORE_FETCH_PATCHED__) {
  const origFetch = global.fetch;
  global.fetch = (input: RequestInfo | URL, init: RequestInit = {}) => {
    // 既に no-store 指定がある場合は尊重し、無いときだけ付与
    const patchedInit: RequestInit = {
      cache: "no-store",
      next: { revalidate: 0, ...(init as any)?.next },
      ...init,
    };
    return origFetch(input, patchedInit);
  };
  (global as any).__NO_STORE_FETCH_PATCHED__ = true;
}
