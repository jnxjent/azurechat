export function normalizeCompanyName(value: string): string {
  return value.normalize("NFKC").replace(/株式会社|有限会社|合同会社|\(株\)|\(有\)/g, "").replace(/[\s・]/g, "").toLowerCase();
}

export function extractPresentationCompanyName(title: string, prompt: string): string {
  const text = `${prompt} ${title}`.normalize("NFKC");
  const prefixed = text.match(/(?:株式会社|有限会社|合同会社|\(株\)|\(有\))\s*([ァ-ヶー一-龠A-Za-z0-9・＆&]{2,60}?)(?=という|[はが](?:産業|企業|会社|事業|提供|展開|行|営|扱|製造|販売|サービス)|[のをで](?:公式|HP|ホームページ|Web|ウェブ|会社|営業|初回|紹介|資料)|[、。\s]|$)/i)?.[1];
  if (prefixed) return prefixed.trim();
  const named = text.match(/(?:^|[。\s])([ァ-ヶー一-龠A-Za-z0-9・＆&]{2,60}?)の(?:公式サイト|HP|ホームページ|会社紹介|会社案内|営業資料|初回(?:客先)?訪問)/i)?.[1];
  return named?.trim() ?? "";
}

export function assertCompanyIdentity(expected: string, actual: string): void {
  if (!expected || !actual || normalizeCompanyName(expected) !== normalizeCompanyName(actual)) {
    throw new Error("取得した会社情報が対象会社と一致しないため、PowerPoint作成を停止しました。対象会社の公式URLを教えてください。");
  }
}

export function isCompanySearchResult(name: string, title: string, description: string): boolean {
  const key = normalizeCompanyName(name);
  return key.length >= 3 && normalizeCompanyName(`${title} ${description}`).includes(key);
}
