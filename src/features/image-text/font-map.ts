// File: src/features/image-text/font-map.ts
// フォントファミリ種別 → 実際の TTF ファイル名を解決する小さなユーティリティ
// 事前に /public/fonts に以下が置いてある想定：
// - NotoSansJP-Regular.ttf
// - NotoSansJP-Bold.ttf
// - NotoSerifJP-Regular.ttf
// - NotoSerifJP-Bold.ttf

export type FontFamilyKey = "gothic" | "mincho" | "meiryo";

/**
 * family + bold から /public/fonts 配下の TTF ファイル名を返す。
 * 実際のファイルパスは呼び出し側で `path.join(process.cwd(), "public", "fonts", fileName)` などで組み立てる。
 */
export function resolveFontFileName(
  family: FontFamilyKey | undefined,
  bold: boolean | undefined
): string {
  const fam: FontFamilyKey = family ?? "gothic";

  if (fam === "mincho") {
    return bold ? "NotoSerifJP-Bold.ttf" : "NotoSerifJP-Regular.ttf";
  }

  // "gothic" / "meiryo" は NotoSansJP で代用
  return bold ? "NotoSansJP-Bold.ttf" : "NotoSansJP-Regular.ttf";
}
