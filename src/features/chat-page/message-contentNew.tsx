/**
 * SF連携の「レコードURL: https://...」をクリック可能なMarkdownに変換する。
 * - citation 行（{% citation ... %}）は一切触らない
 * - ★FIX: GFMテーブル行は触らない（黒い箱の主因）
 * - ★FIX: URL末尾に混入しがちな記号（| ) ] . , 等）を除去
 */
const normalizeContent = (src: string): string => {
  if (!src) return "";

  const lines = src.split(/\r?\n/).map((line) => {
    // 引用タグは絶対に触らない
    if (line.includes("{% citation")) return line;

    // ★FIX(1): GFMテーブルっぽい行は触らない（最重要）
    // 例: "| 商談名 | ... | [開く](https://...) |"
    const trimmed = line.trim();
    const looksLikeTableRow =
      trimmed.startsWith("|") && trimmed.includes("|") && !trimmed.startsWith("|-");
    if (looksLikeTableRow) return line;

    // 「レコードURL: https://...」/「URL: https://...」などを検出
    // ★FIX(2): URL は「空白/|/)/]」などで止める（\S+ をやめる）
    const m = line.match(
      /^(.*?(?:レコードURL|URL|画像URL))\s*[:：]\s*(https?:\/\/[^\s|)\]]+)\s*$/i
    );
    if (m) {
      const labelPart = m[1].trim(); // "レコードURL" など
      let url = (m[2] || "").trim();

      // ★FIX(3): 念のため末尾の句読点・区切りを落とす
      while (/[|)\].,}。、】【]$/.test(url)) {
        url = url.slice(0, -1);
      }

      // 空になったら元に戻す
      if (!url) return line;

      // ラベル行 + URLを別行でMarkdownリンク化
      return `${labelPart}:\n[${url}](${url})`;
    }

    return line;
  });

  return lines.join("\n");
};
