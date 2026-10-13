/** Route image output requests, while leaving advice and Office output alone. */
export function isTeamsImageEditRequest(message: string): boolean {
  return /(?:差し替|差替|置き換|置換|変更|編集|修正|消して|削除|replace|swap)/i.test(message)
    || /(?:今出した|先ほどの|さっきの|今の|この)(?:画像|絵|写真|車|トラック|廃棄物処理車)/i.test(message);
}

export function isTeamsImageGenerationRequest(message: string): boolean {
  const value = String(message ?? "").normalize("NFKC").trim();
  if (/(?:PPTX?|PowerPoint|パワーポイント|スライド|プレゼン|Excel|Word|PDF)/i.test(value)) return false;
  if (/(?:プロンプト|作り方|生成方法|操作方法|デザイン案|文章だけ|テキストだけ|画像は不要)/i.test(value)) return false;
  return /(?:画像|絵|イラスト|写真|ロゴ|logo|文字).{0,60}(?:差し替|差替|置き換|置換|変更|replace|swap)/i.test(value)
    || /(?:画像|絵|イラスト|写真|ロゴ).{0,30}(?:生成|作成|作って|描|出して)/i.test(value)
    || /(?:描いて|描画して|デザインして|デザインしてください|画像化して|イメージ化して)/i.test(value);
}
