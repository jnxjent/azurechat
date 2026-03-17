// src/features/common/http/post-json.ts（新規）
export async function postJson<T>(
  url: string,
  body: unknown,
  signal?: AbortSignal
): Promise<T> {
  // 安全ブレーキ：FormData/URLSearchParams は許可しない
  if (typeof FormData !== "undefined" && body instanceof FormData) {
    throw new Error("postJson() に FormData は渡せません（ファイル送信は既存ルートを使用）");
  }
  if (typeof URLSearchParams !== "undefined" && body instanceof URLSearchParams) {
    throw new Error("postJson() に URLSearchParams は渡せません");
  }

  const json = JSON.stringify(body); // ← UTF-8 で表現される JSON 文字列
  const res = await fetch(url, {
    method: "POST",
    headers: {
      "content-type": "application/json; charset=utf-8",
      accept: "application/json",
    },
    body: json, // JSON文字列をそのまま
    cache: "no-store",
    signal,
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`HTTP ${res.status}: ${text.slice(0, 400)}`);
  }
  const ct = res.headers.get("content-type") || "";
  if (ct.includes("application/json")) return (await res.json()) as T;
  return ({ raw: await res.text() } as unknown as T);
}
