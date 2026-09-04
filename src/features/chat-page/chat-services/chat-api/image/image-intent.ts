const IMAGE_REFERENCE_PATTERN =
  /(?:この|今の|現在の|さっきの|先ほどの|生成した|作った)?\s*(?:画像|絵|イラスト|写真)|(?:この|今の|現在の|さっきの|先ほどの)\s*(?:もの|やつ)/i;

const TEXT_NOUN_PATTERN =
  /(?:文字|テキスト|文言|コピー|キャプション|タイトル|見出し|メッセージ|セリフ)/i;

const TEXT_ADD_ACTION_PATTERN =
  /(?:加え|追加|入れ|載せ|のせ|表示|記載|書き込|書き足|文字入れ)/i;

const QUOTED_TEXT_PATTERN = /[「『“"]\s*[^」』”"]+\s*[」』”"]/;

/** チャット添付のdata URLを画像編集APIへ渡せるバイト列に戻す。 */
export function decodeChatImageDataUrl(value: string): Buffer | null {
  const match = String(value ?? "").match(
    /^data:image\/(?:png|jpeg|webp);base64,([a-z0-9+/=\s]+)$/i
  );
  if (!match) return null;

  try {
    const buffer = Buffer.from(match[1].replace(/\s/g, ""), "base64");
    const isPng =
      buffer.length >= 8 &&
      buffer.subarray(0, 8).equals(
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
      );
    const isJpeg =
      buffer.length >= 3 &&
      buffer[0] === 0xff &&
      buffer[1] === 0xd8 &&
      buffer[2] === 0xff;
    const isWebp =
      buffer.length >= 12 &&
      buffer.toString("ascii", 0, 4) === "RIFF" &&
      buffer.toString("ascii", 8, 12) === "WEBP";

    return buffer.length <= 50 * 1024 * 1024 && (isPng || isJpeg || isWebp)
      ? buffer
      : null;
  } catch {
    return null;
  }
}

/**
 * 旧Sharp文字合成を許可するのは、既存画像への文字追加が明示された場合だけ。
 * 「明るくして」「人物を追加して」等の通常の画像編集は意図的に除外する。
 */
export function isExplicitTextOverlayRequest(message: string): boolean {
  const value = String(message ?? "").trim();
  if (!value || !IMAGE_REFERENCE_PATTERN.test(value)) return false;
  if (!TEXT_ADD_ACTION_PATTERN.test(value)) return false;

  return TEXT_NOUN_PATTERN.test(value) || QUOTED_TEXT_PATTERN.test(value);
}

/** 添付画像をVision分析ではなく画像編集ツールへ送るべき依頼かを判定する。 */
export function isImageEditOrCompositionRequest(message: string): boolean {
  const value = String(message ?? "").trim();
  if (!value) return false;

  const hasEditAction =
    /(?:編集|修正|変更|変え|差し替|置き換|置換|交換|反映|適用|使用|使って|にして|貼り|貼って|追加|加え|付け|つけ|組み込|合成|配置|載せ|のせ|入れ|消し|削除|除去|明るく|暗く|大きく|小さく|移動)/i.test(
      value
    );
  if (!hasEditAction) return false;

  return (
    /(?:画像|絵|イラスト|写真|ロゴ|マーク|ラベル|添付|今の|現在の|この)/i.test(
      value
    ) || value.length <= 40
  );
}

export function resolveRequiredImageToolName(
  message: string,
  attachmentCount: number
): "edit_existing_image" | "add_text_to_existing_image" | undefined {
  if (
    !Number.isFinite(attachmentCount) ||
    attachmentCount <= 0 ||
    (!isImageEditOrCompositionRequest(message) &&
      !isNewImageReferenceCompositionRequest(message) &&
      !isSharePointImageUseRequest(message))
  ) {
    return undefined;
  }

  return isExplicitTextOverlayRequest(message)
    ? "add_text_to_existing_image"
    : "edit_existing_image";
}

/** 添付素材を使って、既存画像ではなく新しい画像を構成する依頼かを判定する。 */
export function isNewImageReferenceCompositionRequest(message: string): boolean {
  const value = String(message ?? "").trim();
  if (!value) return false;

  const hasCreateAction =
    /(?:デザイン|新規|新しい|作成|生成|描いて|描画|イメージ化|ビジュアル化)/i.test(
      value
    );
  const hasReferenceAsset =
    /(?:添付|参照|ロゴ|マーク|ラベル|素材|写真|画像)/i.test(value);
  return hasCreateAction && hasReferenceAsset;
}

/**
 * ツール引数へ整形される過程でユーザー原文が失われないよう、原文を最優先で残す。
 * toolPrompt は「それでお願いします」のような省略指示を会話文脈で補う用途に限定する。
 */
export function buildFaithfulImagePrompt(
  userMessage: string,
  toolPrompt: string,
  mode: "generate" | "edit"
): string {
  const original = String(userMessage ?? "").trim();
  const supplement = String(toolPrompt ?? "").trim();

  if (!original) return supplement;
  if (!supplement || supplement === original) return original;

  const modeInstruction =
    mode === "edit"
      ? [
          "入力画像を部分編集してください。全体を新しく描き直してはいけません。",
          "【編集範囲の固定ルール】",
          "- ユーザーが最新メッセージで明示した対象と、その変更に必然的に伴う最小範囲だけを変更してください。",
          "- 明示されていない人物、顔、表情、体格、ポーズ、服装、物体、背景、構図、カメラ位置、画角、遠近感、配置、比率、色、照明、画風、質感、文字、ロゴは元画像のまま維持してください。",
          "- 要素を勝手に追加・削除・複製・移動しないでください。",
          "- 元画像と変更後画像を重ねたとき、指定箇所以外に差分が出ないことを目標にしてください。",
          "- ユーザーが全体変更を明示した場合でも、その指示に関係しない内容・構図・配置は維持してください。",
        ].join("\n")
      : "新しい画像を生成してください。";

  return [
    modeInstruction,
    "以下のユーザー原文を最優先し、省略・一般化・翻訳せず、その言語のまま忠実に反映してください。",
    "",
    "【ユーザー原文】",
    original,
    "",
    "【会話文脈からの補足（原文にない変更対象を増やしてはいけません）】",
    supplement,
  ].join("\n");
}

export function normalizeGptImageSize(
  size?: string
): "1024x1024" | "1024x1536" | "1536x1024" | "auto" {
  if (size === "1024x1792") return "1024x1536";
  if (size === "1792x1024") return "1536x1024";
  if (
    size === "1024x1024" ||
    size === "1024x1536" ||
    size === "1536x1024"
  ) {
    return size;
  }
  return "auto";
}

export function isSharePointImageRequest(message: string): boolean {
  const value = String(message ?? "");
  return (
    /(?:SharePoint|\bSP\b|ＳＰ|\bSL\b|ＳＬ)/i.test(value) &&
    /(?:ロゴ|画像|イメージ|\.(?:png|jpe?g|webp)\b|logo|image)/i.test(value)
  );
}

export function isSharePointImageUseRequest(message: string): boolean {
  const value = String(message ?? "");
  return (
    isSharePointImageRequest(value) &&
    /(?:使って|使用|利用|参照|差し替|置き換|入れて|追加|組み込|デザイン|生成|作って|create|use|replace|insert|add)/i.test(
      value
    )
  );
}

export function extractSharePointImageQuery(message: string): string | null {
  const value = String(message ?? "").trim();
  if (!isSharePointImageRequest(value)) return null;

  const afterSharePoint = value.split(
    /(?:SharePoint|\bSP\b|ＳＰ|\bSL\b|ＳＬ)/i
  )[1];
  const sharePointTail = String(afterSharePoint ?? "").replace(
    /^(?:上|内|側)?(?:にある|に保存された|の中の|の)?\s*/,
    ""
  );
  const explicitFileName = sharePointTail.match(
    /^[「『'"\s]*([^「」『』'"/\\]+?\.(?:png|jpe?g|webp))\b/i
  );
  if (explicitFileName?.[1]) return explicitFileName[1].trim();

  const query = sharePointTail
    .replace(/^[「『'"\s]+|[」』'"\s]+$/g, "")
    .replace(
      /(?:を|に)?(?:使って|使用して|利用して|参照して|差し替えて|置き換えて|入れて|追加して).*$/,
      ""
    )
    .replace(/[「」『』'"\s]+$/g, "")
    .trim();

  if (query) return query;
  return /(?:ロゴ|logo)/i.test(value) ? "logo" : "image";
}

/** gpt-image-2へ渡す複数画像の役割を明示し、参照素材の勝手な改変を抑える。 */
/** Model-provided reference assets must be fetchable URLs or supported image data URLs. */
export function isSupportedImageReferenceUrl(value: string): boolean {
  const normalized = String(value ?? "").trim();
  if (/^data:image\/(?:png|jpe?g|webp);base64,/i.test(normalized)) {
    return true;
  }
  try {
    const parsed = new URL(normalized);
    return parsed.protocol === "https:" || parsed.protocol === "http:";
  } catch {
    return false;
  }
}

/**
 * Resolve the image source for a PowerPoint edit. An image attached in the
 * current turn is authoritative; model-authored placeholders such as
 * "添付画像" or "attachment://logo" must never shadow it.
 */
export function resolvePptxEditImageSource(
  modelImageUrl: unknown,
  currentAttachmentUrls: unknown
): string {
  const attachments = Array.isArray(currentAttachmentUrls)
    ? currentAttachmentUrls
    : [];
  for (let index = attachments.length - 1; index >= 0; index--) {
    const candidate = String(attachments[index] ?? "").trim();
    if (isSupportedImageReferenceUrl(candidate)) return candidate;
  }

  const modelCandidate = String(modelImageUrl ?? "").trim();
  return isSupportedImageReferenceUrl(modelCandidate) ? modelCandidate : "";
}

/** Removes query strings such as SAS signatures before an image location is logged. */
export function sanitizeImageLocationForLog(value: string): string {
  const normalized = String(value ?? "").trim();
  if (!normalized) return "unknown";
  if (/^data:image\//i.test(normalized)) return "chat-attachment:data-url";
  if (/^thread:/i.test(normalized)) return normalized;

  try {
    const parsed = new URL(normalized);
    if (parsed.protocol !== "https:" && parsed.protocol !== "http:") {
      return "non-http-reference";
    }
    return `${parsed.protocol}//${parsed.host}${parsed.pathname}`;
  } catch {
    return "non-url-reference";
  }
}

export function buildMultiImageReferenceInstruction(
  referenceCount: number
): string {
  if (!Number.isFinite(referenceCount) || referenceCount <= 0) return "";

  const references = Array.from(
    { length: Math.min(Math.floor(referenceCount), 15) },
    (_, index) =>
      `- 画像${index + 2}: ユーザーが組み込みを求めた参照素材${
        index === 0 ? "（例: 添付ロゴ）" : ""
      }。`
  );

  return [
    "【複数画像の役割】",
    "- 画像1: 完成画像の土台となる編集元画像。指定箇所以外は変更しないでください。",
    ...references,
    "- 参照素材は、ユーザーが指定した対象・位置にだけ組み込んでください。参照画像の背景や余白は、ユーザーが求めない限り持ち込まないでください。",
    "- ロゴ、ラベル、マークを組み込む場合は、綴り、文字、図形、色、縦横比、要素間隔を参照画像どおりに維持し、勝手に描き換えないでください。",
    "- 配置先の遠近感、面の角度、縮尺、照明、影、反射だけを自然に合わせてください。",
  ].join("\n");
}

/** 新規生成時は、入力画像をキャンバスではなく合成用の参照素材として扱わせる。 */
export function buildNewImageReferenceInstruction(imageCount: number): string {
  if (!Number.isFinite(imageCount) || imageCount <= 0) return "";

  return [
    "【添付画像の扱い】",
    `- 入力画像1〜${Math.min(Math.floor(imageCount), 16)}は編集元キャンバスではなく、新しい画像へ組み込む参照素材です。`,
    "- ユーザーの指示どおり新しい場面・構図を生成し、添付素材は指定された対象と位置に自然に組み込んでください。",
    "- ロゴ、マーク、ラベルは綴り、文字、形状、色、縦横比を参照画像どおりに保ち、勝手に描き替えないでください。",
    "- 参照素材の背景や余白は、ユーザーが求めない限り新しい画像へ持ち込まないでください。",
  ].join("\n");
}
