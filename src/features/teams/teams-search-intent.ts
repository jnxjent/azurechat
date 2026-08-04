/**
 * Teams chat uses internal Azure AI Search only when the user explicitly asks
 * for company knowledge. General questions must not be grounded accidentally
 * by unrelated internal documents.
 */
export function requiresTeamsInternalSearch(message: string): boolean {
  const normalized = message.normalize("NFKC").trim();
  if (!normalized) return false;

  return /(?:社内(?:の)?(?:資料|文書|書類|規程|規定|ルール|マニュアル|情報|データ|ナレッジ|検索)|(?:会社|当社)(?:の)?(?:資料|文書|書類|規程|規定|ルール|マニュアル)|社内検索|share\s*point|ai\s*search|(?:SP|SL)(?:上|内|にある|の)(?:資料|文書|書類|ファイル|規程|規定)|(?:全社共通|全社共有|部署共通|部門共通|部署共有|部門共有)(?:資料|ファイル|フォルダ(?:ー)?)?|(?:個人|自分の|私の)(?:資料|ファイル|フォルダ(?:ー)?)|(?:という|と呼ばれる)\s*フォルダ(?:ー)?)/i.test(
    normalized
  );
}
