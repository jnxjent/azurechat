/**
 * DeckSpec — PPTX生成時の意味構造をサイドカーJSONとして保存するモデル。
 * ・自システム生成PPTXにのみ付随する（外部PPTXにはない）
 * ・編集時にlayoutType/itemsを参照してレイアウト別再描画を行う
 * ・Blob: pptx-specs コンテナ内 {blobKey}.spec.json に保存（非公開）
 */

export type DeckSpecItem = {
  id: string;
  heading?: string;
  body: string;
  iconKey?: string;
};

export type DeckSpecSlide = {
  id: string;
  title: string;
  /** 元スライドの visualType（再描画時に spotlight 固定を防ぐ） */
  visualHint?: { visualType: string };
  /** 生成時のlayoutType（gen-pptxのPptxSlide.layoutTypeと同値） */
  layoutType:
    | "title"
    | "bullets"
    | "card_grid"
    | "icon_rows"
    | "table"
    | "multi-column"
    | "diagram"
    | "stat_callouts"
    | "metric-cards"
    | "timeline"
    | "roadmap"
    | "conversation"
    | "company-overview"
    | "process-cards"
    | "split_visual"
    | "comparison_matrix"
    | "decision_summary"
    | "editorial_statement"
    | "asymmetric_list"
    | "closing";
  /** 正規化済み項目リスト（renderer と同じ優先順で抽出） */
  items: DeckSpecItem[];
  /**
   * 実際のPPTXにおける0-basedスライド番号。
   * 非faithfulモードでは表紙スライドが0番に挿入されるため、
   * sanitizedSlides[i] → pptxSlideIndex = i + 1（faithfulモードは i）
   */
  pptxSlideIndex: number;
  /**
   * 元のPptxSlideオブジェクト（faithful再生成用）。
   * 内部用途のみ——公開URLには出さない（pptx-specsコンテナは非公開）
   */
  rawSlide: Record<string, unknown>;
};

/** gen-pptxリクエストのうち再生成に必要なメタ情報 */
export type DeckSpecGenMeta = {
  title: string;
  palette?: string;
  /** designInstructionは内部用のみ（サイドカーは非公開コンテナ保存） */
  designInstruction?: string;
  mode?: string;
  /** 表紙スライド用フィールド（TypeScript再描画で buildTitleSlide に渡す） */
  coverSubtitle?: string;
  coverKicker?: string;
  footerNote?: string;
  mood?: string;
  /**
   * 最終的な14色パレット全スナップショット（LLM調整・accentOverride・normalizePaletteDiversity反映済み）。
   * paletteKey から再構築するよりも正確。旧DeckSpec向けフォールバックとして paletteKey も残す。
   * キーは Palette の全フィールド名、値は6桁hex（# なし）。
   */
  paletteSnapshot?: Record<string, string>;
  /**
   * DALL-E 生成イラストの Blob キー（pptx-specs コンテナ内 "{key}.illustration"）。
   * イラスト本体をDeckSpec JSONに入れると巨大化するため別Blobに保存。
   */
  illustrationBlobKey?: string;
  illustrationPrompt?: string;
  /** DeckStyleSpec 由来フィールド（TypeScript再描画でThemeを再構築するために保存） */
  cardStyle?: string;
  headerStyle?: string;
  visualStyle?: string;
  deckPurpose?: string;
  execMode?: boolean;
  playfulMode?: boolean;
  minimalMode?: boolean;
  /** 解決済みフォント名（fontFaceリクエストパラメータではなく resolveTheme 後の値） */
  fontFace?: string;
  language?: string;
  useJapaneseLabels?: boolean;
  /** 元生成時のフォントサイズ（再描画でスライド枚数依存の計算を省くため保存） */
  titleFontSize?: number;
  bodyFontSize?: number;
  smallFontSize?: number;
  /** B4 Art Director メタ情報（再描画では Art Director を再実行しない） */
  artDirectorMode?: "off" | "audit" | "apply";
  artDirectionVersion?: 1;
  artDirectionMotif?: string;
  artDirectionApplied?: boolean;
};

export type DeckSpec = {
  version: 1;
  /** 生成時に付与した一意ID（PPTX blobKeyに対応） */
  deckId: string;
  /** 編集ごとに+1するリビジョン番号 */
  revision: number;
  /** パレットキー（PPTX_NAMED_PALETTESのキー） */
  paletteKey?: string;
  slides: DeckSpecSlide[];
  /** 再生成用メタ情報 */
  genMeta: DeckSpecGenMeta;
  /** 保存日時（ISO 8601） */
  savedAt: string;
};
