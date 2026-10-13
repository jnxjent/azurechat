/** Explicit Office output takes precedence over raster image composition. */
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
  return isExplicitPptOutputRequest(message) &&
    !/(?:PDF|添付ファイル).{0,20}(?:PPT|PowerPoint|変換|にして)/i.test(message) &&
    /作成|生成|作って|create/i.test(message) &&
    !/(?:既存|先ほど|直前|出力した|作成した|今の|このPPT).{0,20}(?:編集|変更|修正)/i.test(message);
}

export function pdfPptConversionMode(message: string): "faithful" | "redesign" {
  // "編集可能" alone is served by faithful mode (source images kept, text editable).
  return /再構成|リデザイン|デザイン.{0,8}(?:改善|変更)|要約|redesign/i.test(message)
    ? "redesign" : "faithful";
}
