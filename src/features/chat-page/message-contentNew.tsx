// 2) 「名前：URL」→ Markdownリンク化、残りの生URLは <URL> で保険
const normalizeToMarkdown = (src: string): string => {
  if (!src) return "";

  const mdLinkPattern = /\[[^\]]+\]\([^)]+\)/; // 既に markdown リンクなら触らない

  const lines = src.split(/\r?\n/).map((line) => {
    // 既に Markdown リンクが含まれる行はそのまま返す（副作用防止）
    if (mdLinkPattern.test(line)) return line;

    // 「- ラベル：URL」/「ラベル：URL」/ 全角・半角コロン両対応 → 正しく Markdown 化
    const m = line.match(/^\s*(-\s*)?([^\n:：]+?)\s*[：:]\s*(https?:\/\/\S+)\s*$/);
    if (m) {
      const dash = m[1] ? "- " : "";
      const label = m[2].trim();
      const url = m[3].trim(); // ★ 末尾に / を付けない
      return `${dash}[${label}](${url})`;
    }

    // 残った生URLだけ <URL> 化。ただし markdown のリンク先URLは除外したいので
    // 直前文字が '(' または '<' の場合はスキップ
    return line.replace(/https?:\/\/\S+/g, (u, offset) => {
      const prev = line[offset - 1] || "";
      if (prev === "(" || prev === "<") return u; // リンク先 or 既に <URL> の一部
      return `<${u}>`;
    });
  });

  return lines.join("\n");
};
