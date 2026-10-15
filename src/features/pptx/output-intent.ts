/** Explicit Office output takes precedence over raster image composition. */
export const PPT_PRODUCTION_INSTRUCTION = "ロゴ配置・配色・枚数・ファイル形式・添付ファイル名などの制作指示はスライド本文・見出しに書かない。これらは描画設定として扱い、本文は読者向けの内容だけにする。";
export function removePptProductionNotes<T>(value: T): T {
  if (typeof value === "string") {
    return value.split(/\n/).filter(line => !/(?:添付|会社)?(?:ロゴ|logo).{0,60}(?:配置|挿入|貼り|右上)|右上.{0,30}(?:ロゴ|logo)|(?:添付|ロゴ|logo).{0,60}\.(?:png|jpe?g|webp)/i.test(line)).join("\n") as T;
  }
  if (Array.isArray(value)) return value.map(removePptProductionNotes).filter(v => v !== "") as T;
  if (value && typeof value === "object") return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, removePptProductionNotes(v)])) as T;
  return value;
}

export function requestsOfficialWebsiteContent(text: string): boolean {
  return /(?:hp|ホームページ|web|ウェブ|公式サイト).{0,40}(?:入手|取得|反映|引用|参考|参照|調べ|検索|情報|会社案内|会社概要|企業情報|会社情報|内容を埋め)/i.test(text.normalize("NFKC"));
}

export function isCompanyProfileContent(text: string, slideCount: number): boolean {
  return requestsOfficialWebsiteContent(text) || (/会社紹介|(?<!機能)紹介資料|company profile|初回(?:客先|顧客|お客様)?訪問|初回営業/i.test(text) && slideCount <= 16);
}

export function resolveOfficeChatRoute(input: { explicitPptRequest: boolean; requiredTools: boolean; imageCount: number; documentCount: number }): "extensions" | "multimodal" | "chat-with-file" {
  if (input.explicitPptRequest || input.requiredTools) return "extensions";
  if (input.imageCount > 0) return "multimodal";
  if (input.documentCount > 0) return "chat-with-file";
  return "extensions";
}

export function isExplicitPptOutputRequest(message: string): boolean {
  const value = message.normalize("NFKC");
  return /(?:\bpptx?\b|powerpoint|パワーポイント)/i.test(value) &&
    /作成|生成|作って|出力|変換|にして|編集|修正|変更|create|export|convert/i.test(value) &&
    !/(?:PPTX?|PowerPoint|パワーポイント).{0,8}(?:不要|いらない|ではなく)/i.test(value);
}

export function isNewPptCreationRequest(message: string): boolean {
  if (isAttachedPptConversionRequest(message)) return false;
  return isExplicitPptOutputRequest(message) &&
    !/(?:PDF|添付ファイル).{0,20}(?:PPT|PowerPoint|変換|にして)/i.test(message) &&
    /作成|生成|作って|create/i.test(message) &&
    !/(?:既存|先ほど|直前|出力した|作成した|今の|このPPT).{0,20}(?:編集|変更|修正)/i.test(message);
}

/** Conversion of an uploaded source, including requests that omit its extension. */
export function isAttachedPptConversionRequest(message: string): boolean {
  return isExplicitPptOutputRequest(message) && /添付|アップロード|このファイル|この画像|png|jpe?g|webp|PDF|画像|ポスター|スクリーンショット/i.test(message) &&
    /変換|にして|スライド化|復元|編集可能/i.test(message) &&
    !/参考|(?:ロゴ|logo).{0,30}(?:右上|挿入|追加|配置|差し替)/i.test(message);
}

export function isImagePptConversionRequest(message: string): boolean {
  return isExplicitPptOutputRequest(message) && /png|jpe?g|webp|画像|ポスター|スクリーンショット/i.test(message) &&
    /変換|にして|スライド化|復元|編集可能/i.test(message) && !/参考|(?:ロゴ|logo).{0,30}(?:右上|挿入|追加|配置|差し替)/i.test(message);
}

export function isExcelTableToSlideRequest(message: string): boolean {
  const value = message.normalize("NFKC");
  return /(?:\bpptx?\b|\bpptx?(?=\d+\s*枚)|powerpoint|パワーポイント|パワポ|スライド)/i.test(value) &&
    /にして|変換|作成|作って|生成|化して|まとめて/.test(value) &&
    /1\s*枚|一枚|見栄え|みばえ|綺麗|きれい|キレイ|見やす|体裁|清書|整えて|おしゃれ|かっこ(?:よ|い)|格好(?:よ|い)/.test(value) &&
    !/(?:[2-9]|[1-9]\d|二|三|四|五|六|七|八|九|十)\s*枚/.test(value) &&
    !/編集|修正|変更|フォント|ロゴ|カード|翻訳|追加|削除|グラフ|チャート|参考|拡充|追記|肉付け|PDF|画像|Word|ワード/i.test(value) &&
    !/(?:この|その|先ほど|さっき|前の|今の|作成した|出力した)\s*(?:PPT|パワポ|パワーポイント|スライド)/i.test(value);
}

export function pdfPptConversionMode(message: string): "faithful" | "redesign" {
  // "編集可能" alone is served by faithful mode (source images kept, text editable).
  return /再構成|リデザイン|デザイン.{0,8}(?:改善|変更)|要約|redesign/i.test(message)
    ? "redesign" : "faithful";
}
