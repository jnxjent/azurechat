export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import PptxGenJS from "pptxgenjs";
import JSZip from "jszip";
import {
  BlobSASPermissions,
  BlobServiceClient,
} from "@azure/storage-blob";
import { uniqueId } from "@/features/common/util";
import { OpenAIDALLEInstance, OpenAIInstance } from "@/features/common/services/openai";

// ── SVGアイコン（process-cards用） ────────────────────────────────────────────
function _svgUri(svg: string): string {
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`;
}

const STEP_ICON_URIS: Record<string, string> = {
  // 収集・運搬 → トラック
  truck: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M20 8h-3V4H3c-1.1 0-2 .9-2 2v11h2c0 1.66 1.34 3 3 3s3-1.34 3-3h6c0 1.66 1.34 3 3 3s3-1.34 3-3h2v-5l-3-4zM6 18.5c-.83 0-1.5-.67-1.5-1.5s.67-1.5 1.5-1.5 1.5.67 1.5 1.5-.67 1.5-1.5 1.5zm13.5-9l1.96 2.5H17V9.5h2.5zm-1.5 9c-.83 0-1.5-.67-1.5-1.5s.67-1.5 1.5-1.5 1.5.67 1.5 1.5-.67 1.5-1.5 1.5z"/></svg>`),
  // 処理・加工 → ギア
  gear: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M19.14 12.94c.04-.3.06-.61.06-.94 0-.32-.02-.64-.07-.94l2.03-1.58c.18-.14.23-.41.12-.61l-1.92-3.32c-.12-.22-.37-.29-.59-.22l-2.39.96c-.5-.38-1.03-.7-1.62-.94l-.36-2.54c-.04-.24-.24-.41-.48-.41h-3.84c-.24 0-.43.17-.47.41l-.36 2.54c-.59.24-1.13.57-1.62.94l-2.39-.96c-.22-.08-.47 0-.59.22L2.74 8.87c-.12.21-.08.47.12.61l2.03 1.58c-.05.3-.09.63-.09.94s.02.64.07.94l-2.03 1.58c-.18.14-.23.41-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.38 1.03.7 1.62.94l.36 2.54c.05.24.24.41.48.41h3.84c.24 0 .44-.17.47-.41l.36-2.54c.59-.24 1.13-.56 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32c.12-.22.07-.47-.12-.61l-2.01-1.58zM12 15.6c-1.98 0-3.6-1.62-3.6-3.6s1.62-3.6 3.6-3.6 3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z"/></svg>`),
  // 処分・最終 → 倉庫
  archive: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M20.54 5.23l-1.39-1.68C18.88 3.21 18.47 3 18 3H6c-.47 0-.88.21-1.16.55L3.46 5.23C3.17 5.57 3 6.02 3 6.5V19c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V6.5c0-.48-.17-.93-.46-1.27zM12 17.5L6.5 12H10v-2h4v2h3.5L12 17.5zM5.12 5l.81-1h12l.94 1H5.12z"/></svg>`),
  // 法令・遵守 → シールド
  shield: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm-2 16l-4-4 1.41-1.41L10 14.17l6.59-6.59L18 9l-8 8z"/></svg>`),
  // コスト・費用 → 円マーク
  coins: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M11.8 10.9c-2.27-.59-3-1.2-3-2.15 0-1.09 1.01-1.85 2.7-1.85 1.78 0 2.44.85 2.5 2.1h2.21c-.07-1.72-1.12-3.3-3.21-3.81V3h-3v2.16c-1.94.42-3.5 1.68-3.5 3.61 0 2.31 1.91 3.46 4.7 4.13 2.5.6 3 1.48 3 2.41 0 .69-.49 1.79-2.7 1.79-2.06 0-2.87-.92-2.98-2.1h-2.2c.12 2.19 1.76 3.42 3.68 3.83V21h3v-2.15c1.95-.37 3.5-1.5 3.5-3.55 0-2.84-2.43-3.81-4.7-4.4z"/></svg>`),
  // 環境・ESG → エコ
  leaf: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M6.05 8.5c.44-4.06 4-7.5 8.95-7.5C16 1 22 1 22 1s0 6-1 10c-1.25 5.19-6.5 8-11 7-.49-1.27-.76-2.61-.76-4C9.24 11.13 7.36 9.45 6.05 8.5zM2 21c0-2.76 2.24-5 5-5l-1 3 4-4c-3.15-.54-5 2.62-5 6H2z"/></svg>`),
  // 安全・管理 → 目
  eye: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>`),
  // 創業・日付 → カレンダー
  calendar: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M20 3h-1V1h-2v2H7V1H5v2H4c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 18H4V8h16v13zm-2-7H6v-2h12v2zm-4 4H6v-2h8v2z"/></svg>`),
  // 住所・地図 → ピン
  location: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z"/></svg>`),
  // 上場・株式 → トレンドアップ（既存chartと区別）
  stock: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg>`),
  // 会社・拠点 → ビル
  building: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M12 7V3H2v18h20V7H12zM6 19H4v-2h2v2zm0-4H4v-2h2v2zm0-4H4V9h2v2zm0-4H4V5h2v2zm4 12H8v-2h2v2zm0-4H8v-2h2v2zm0-4H8V9h2v2zm0-4H8V5h2v2zm10 12h-8v-2h2v-2h-2v-2h2v-2h-2V9h8v10zm-2-8h-2v2h2v-2zm0 4h-2v2h2v-2z"/></svg>`),
  // 人材・チーム → グループ
  people: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>`),
  // 実績・成長 → 右上トレンド
  chart: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M16 6l2.29 2.29-4.88 4.88-4-4L2 16.59 3.41 18l6-6 4 4 6.3-6.29L22 12V6z"/></svg>`),
  // 強み・品質 → スター
  star: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M12 17.27L18.18 21l-1.64-7.03L22 9.24l-7.19-.61L12 2 9.19 8.63 2 9.24l5.46 4.73L5.82 21z"/></svg>`),
  // 認定・信頼 → バッジ
  verified: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M23 12l-2.44-2.78.34-3.68-3.61-.82-1.89-3.18L12 3 8.6 1.54 6.71 4.72l-3.61.81.34 3.68L1 12l2.44 2.78-.34 3.69 3.61.82 1.89 3.18L12 21l3.4 1.46 1.89-3.18 3.61-.82-.34-3.68L23 12zm-12 2.5l-3.5-3.5 1.41-1.41L11 12.67l5.59-5.59L18 8.5l-7 7z"/></svg>`),
  // 提案・アイデア → 電球
  lightbulb: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M9 21c0 .55.45 1 1 1h4c.55 0 1-.45 1-1v-1H9v1zm3-19C8.14 2 5 5.14 5 9c0 2.38 1.19 4.47 3 5.74V17c0 .55.45 1 1 1h6c.55 0 1-.45 1-1v-2.26c1.81-1.27 3-3.36 3-5.74 0-3.86-3.14-7-7-7z"/></svg>`),
  // 革新・DX → ロケット
  rocket: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M9.19 6.35c-2.04 2.29-3.44 5.58-3.57 5.89L2 10l4.06-4.06c.61-.61 1.46-.95 2.33-.95.87 0 1.73.34 2.33.95l.21.21c-.29.06-.59.12-.74.2zm9.46 4.1c-.31.13-3.58 1.54-5.87 3.57l-.2.74-.21-.21c-.61-.61-.95-1.46-.95-2.33 0-.87.34-1.73.95-2.33L16.44 5.82l2.21 4.63zm-8.84 7.37l-3.81-3.81c-1.57 1.58-3 3.84-3 6l3-1 1 3 2.81-4.19zM15.5 7c-.55 0-1 .45-1 1s.45 1 1 1 1-.45 1-1-.45-1-1-1z"/></svg>`),
  // 取引先・ネットワーク → ハブ
  network: _svgUri(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z"/></svg>`),
};

function resolveStepIconKey(hintOrTitle: string): keyof typeof STEP_ICON_URIS | null {
  const t = hintOrTitle;
  // explicit key pass-through
  if (t in STEP_ICON_URIS) return t as keyof typeof STEP_ICON_URIS;
  // waste-industry specific
  if (/収集|運搬|輸送|トラック/.test(t)) return "truck";
  if (/処理|焼却|破砕|脱水|中間/.test(t)) return "gear";
  if (/処分|埋立|最終|廃棄/.test(t)) return "archive";
  if (/法令|遵守|法規|コンプラ|許認可|認定/.test(t)) return "shield";
  if (/コスト|費用|節約|削減|最適化|経費/.test(t)) return "coins";
  if (/環境|ESG|エコ|リサイクル|循環|持続|脱炭素/.test(t)) return "leaf";
  if (/管理|追跡|モニタ|GPS|電子/.test(t)) return "eye";
  // general business
  if (/会社|本社|拠点|施設|オフィス|ビル|企業/.test(t)) return "building";
  if (/人材|チーム|組織|従業員|スタッフ|社員|採用|人/.test(t)) return "people";
  if (/実績|業績|成長|売上|収益|拡大|展開|トレンド|グラフ/.test(t)) return "chart";
  if (/強み|品質|特徴|差別化|優位|クオリティ|ハイ|ベスト/.test(t)) return "star";
  if (/認定|信頼|安心|保証|資格|ライセンス|基準|品質保証/.test(t)) return "verified";
  if (/提案|ソリューション|アイデア|改善|革新|イノベーション|企画/.test(t)) return "lightbulb";
  if (/DX|IT|デジタル|テクノロジー|システム|AI|自動化|スマート/.test(t)) return "rocket";
  if (/取引|顧客|パートナー|ネットワーク|エリア|地域|広域/.test(t)) return "network";
  return null;
}

export type PptxColumn = { header: string; bullets: string[] };

function resolveMetricIconKey(label: string, hint?: string): keyof typeof STEP_ICON_URIS | null {
  const src = hint ?? label;
  if (src in STEP_ICON_URIS) return src as keyof typeof STEP_ICON_URIS;
  if (/創業|設立|創設|歴史|年|創/.test(src)) return "calendar";
  if (/本社|所在地|住所|拠点|オフィス|所在|地/.test(src)) return "location";
  if (/上場|株式|証券|市場|東証|マザーズ|グロース|プライム|スタンダード/.test(src)) return "stock";
  if (/取引先|顧客|顧客数|クライアント|契約/.test(src)) return "network";
  if (/従業員|社員|スタッフ|人材|社数|人/.test(src)) return "people";
  if (/売上|収益|営業|業績|利益|収入/.test(src)) return "chart";
  if (/施設|工場|設備|処理施設|プラント/.test(src)) return "building";
  if (/処理能力|処理量|年間|トン|容量/.test(src)) return "gear";
  if (/認定|認証|資格|ライセンス|許可/.test(src)) return "verified";
  if (/強み|特徴|優位|品質|クオリティ/.test(src)) return "star";
  return "star";
}

export type PptxMetric = {
  label: string;
  value: string;           // 表示用短縮値（LLMが設定）
  note?: string;           // 補足詳細（LLMが設定）
  iconKey?: string;
  displayValue?: string;   // value の代替（LLMが明示的に短縮した場合）
  colorRole?: "primary" | "accent" | "neutral";  // カードの色役割
};

export type PptxStep = {
  title: string;
  body: string;
  iconKey?: string;
};

export type PptxVisualBlock = {
  kind: "callout" | "node" | "badge" | "figure";
  role?: "primary" | "supporting" | "annotation";
  groupId?: string;
  text: string;
  x: number;
  y: number;
  w: number;
  h: number;
  emphasis?: boolean;
};

export type PptxConnector = {
  from: number;
  to: number;
  label?: string;
  style?: "arrow" | "line";
  relationshipType?: "flow" | "compare" | "annotation" | "support";
};

export type PptxConversationTurn = {
  speakerRole: string;
  speakerType?: "agent" | "customer" | "staff" | "other";
  text: string;
  turnIndex: number;
};

export type PptxStatCallout = {
  value: string;   // 表示数値 (例: "94")
  unit: string;    // 単位   (例: "%", "名", "億円")
  label: string;   // 説明ラベル (例: "顧客満足度")
};

export type PptxCard = {
  iconKey?: string;
  heading: string;
  body: string;
  statusLabel?: string;  // ステータスピル用 (例: "導入済", "計画中")
};

export type PptxSlide = {
  title: string;
  bullets: string[];
  layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram" | "conversation"
             | "company-overview" | "process-cards" | "closing"
             | "metric-cards" | "timeline"
             | "stat_callouts" | "card_grid" | "icon_rows" | "roadmap";
  tableRows?: string[][];
  columns?: PptxColumn[];
  visualBlocks?: PptxVisualBlock[];
  connectors?: PptxConnector[];
  conversationStyle?: "chat-ui" | "interview" | "dialog-list";
  conversationTurns?: PptxConversationTurn[];
  // company-overview
  leadText?: string;
  metrics?: PptxMetric[];
  callout?: { title: string; body: string };
  // process-cards / timeline
  subtitle?: string;
  steps?: PptxStep[];
  benefits?: string[];
  // stat_callouts — 数字3つ大きく表示
  statCallouts?: PptxStatCallout[];
  // card_grid / icon_rows — アイコン+見出し+本文カード
  cards?: PptxCard[];
  // LLMデザイン判断フィールド
  visualIntent?: string;
  density?: "low" | "medium" | "high";
  textTreatment?: "short" | "normal" | "explanatory";
  // Vision fix で設定されるフォント・テキスト調整フィールド
  fontScale?: number;       // 0.70–1.0 実フォント縮小率（build 関数で反映、Vision fontScaleDown で設定）
  fitTextToShape?: boolean; // true: テキストボックス内収縮を build 関数で強制
};

export type DeckPreferencesInput = {
  designInstruction?: string;
  accentColor?: string;
  fontScale?: "small" | "medium" | "large" | "xlarge";
  avoidEnglishLabels?: boolean;
  language?: "ja" | "en";
  recentDesignNotes?: string[];
};

// ユーザープロンプトの意図を構造化して描画まで伝搬する
export type PromptIntent = {
  documentPurpose: "proposal" | "company-intro" | "recruitment" | "training" | "analysis" | "internal" | "ir" | "campaign" | "other";
  audience: "executive" | "customer" | "employee" | "candidate" | "general";
  designFreedom: "conservative" | "balanced" | "expressive";
  toneKeywords: string[];
  colorDirectives?: {
    primary?: string;   // 6桁 HEX (# なし)
    accent?: string;
    background?: string;
  };
  layoutDirectives: {
    preferTwoColumn?: boolean;
    includeTables?: boolean;
    avoidBulletOnly?: boolean;
    preferMetrics?: boolean;
    preferProcess?: boolean;
  };
  styleGuardrails: {
    allowModernDark?: boolean;
    allowPlayful?: boolean;
    allowGlass?: boolean;
    maxAccentIntensity?: "low" | "medium" | "high";
  };
};

export type GenPptxRequest = {
  title: string;
  slides: PptxSlide[];
  threadId: string;
  fontFace?: string;
  designInstruction?: string;
  deckPreferences?: DeckPreferencesInput;
  /** "faithful": 元ページ数維持・タイトルスライド追加なし・デザインAI最小化 */
  mode?: "faithful" | "redesign";
  /** ダウンロード時のファイル名ベース（拡張子なし）。例: "ミダック会社紹介" */
  fileBaseName?: string;
  /** ユーザープロンプトの意図を構造化して渡す */
  promptIntent?: PromptIntent;
  /**
   * 5パレット名のいずれか。LLM または create_pptx ツールが業種・用途から選択して渡す。
   * 未指定時は既存の selectStrictPaletteKey() でキーワード選択。
   */
  palette?: string;
};

type Palette = {
  canvas: string;
  surface: string;
  titleBg: string;
  headerBg: string;
  accentA: string;
  accentB: string;
  headerText: string;
  bodyText: string;
  mutedText: string;
  sectionBg: string;
  tableHeaderBg: string;
  tableHeaderText: string;
  tableAltBg: string;
  border: string;
};

type SlideVisualType =
  | "editorial"
  | "process"
  | "comparison"
  | "spotlight"
  | "cards"
  | "timeline"
  | "table";

type SlideVisualHint = {
  title: string;
  visualType: SlideVisualType;
  emphasis?: string;
};

// LLM が毎回自由生成するデザイン仕様
type DeckStyleSpec = {
  deckPurpose: "recruitment" | "proposal" | "company-intro" | "training" | "analysis" | "ir" | "internal" | "campaign" | "other";
  visualStyle: "corporate-light" | "modern-dark" | "editorial" | "playful" | "minimal" | "bold";
  cardStyle: "default" | "filled" | "glass" | "flat";
  headerStyle: "band" | "minimal" | "accent-line";
};

type DeckDesignBrief = {
  palette: Palette;
  coverKicker: string;
  coverSubtitle: string;
  footerNote: string;
  mood: string;
  visualHints: SlideVisualHint[];
  styleSpec: DeckStyleSpec; // LLM が毎回生成する汎用デザイン仕様
};

type GeneratedIllustration = {
  dataUri: string;
  prompt: string;
};

type Theme = {
  palette: Palette;
  fontFace: string;
  titleFontSize: number;
  bodyFontSize: number;
  smallFontSize: number;
  execMode: boolean;
  playfulMode: boolean;
  minimalMode: boolean;
  language: "ja" | "en";
  useJapaneseLabels: boolean;
  // StyleSpec 由来フィールド（LLM 生成）
  cardStyle: DeckStyleSpec["cardStyle"];
  headerStyle: DeckStyleSpec["headerStyle"];
  visualStyle: DeckStyleSpec["visualStyle"];
  deckPurpose: DeckStyleSpec["deckPurpose"];
};

const DEFAULT_PALETTES: Record<string, Palette> = {
  blue: { canvas: "F4F8FF", surface: "FFFFFF", titleBg: "163D77", headerBg: "214C8F", accentA: "3B82F6", accentB: "93C5FD", headerText: "FFFFFF", bodyText: "122033", mutedText: "58667A", sectionBg: "EAF2FF", tableHeaderBg: "3B82F6", tableHeaderText: "FFFFFF", tableAltBg: "EDF5FF", border: "D4E4FF" },
  red: { canvas: "FFF6F6", surface: "FFFFFF", titleBg: "7B1E2B", headerBg: "A3293B", accentA: "E24A5A", accentB: "F7B0B8", headerText: "FFFFFF", bodyText: "35141B", mutedText: "7B5960", sectionBg: "FFF0F1", tableHeaderBg: "E24A5A", tableHeaderText: "FFFFFF", tableAltBg: "FFF4F5", border: "F2D2D6" },
  green: { canvas: "F3FBF7", surface: "FFFFFF", titleBg: "184B3A", headerBg: "21644D", accentA: "35A073", accentB: "A7E1C7", headerText: "FFFFFF", bodyText: "163128", mutedText: "5B7268", sectionBg: "EAF8F0", tableHeaderBg: "35A073", tableHeaderText: "FFFFFF", tableAltBg: "F2FBF6", border: "D2ECDD" },
  gold: { canvas: "FFFBEF", surface: "FFFFFF", titleBg: "5B4212", headerBg: "7B5A18", accentA: "C6922D", accentB: "F0D99A", headerText: "FFF8E5", bodyText: "36280E", mutedText: "746344", sectionBg: "FFF6DD", tableHeaderBg: "C6922D", tableHeaderText: "FFFFFF", tableAltBg: "FFF9EA", border: "EADDAE" },
  pastel: { canvas: "FFF9FD", surface: "FFFFFF", titleBg: "8A6BBE", headerBg: "A085D6", accentA: "F39BCB", accentB: "CBB7F7", headerText: "FFFFFF", bodyText: "3B3150", mutedText: "7B7091", sectionBg: "F7F1FF", tableHeaderBg: "A085D6", tableHeaderText: "FFFFFF", tableAltBg: "FBF7FF", border: "E6DBFB" },
  pop: { canvas: "FFFDF1", surface: "FFFFFF", titleBg: "8B0D57", headerBg: "D61F69", accentA: "FF9F1C", accentB: "FFE066", headerText: "FFFFFF", bodyText: "2D2230", mutedText: "7F6570", sectionBg: "FFF4DB", tableHeaderBg: "FF9F1C", tableHeaderText: "FFFFFF", tableAltBg: "FFF9E6", border: "FFE2A8" },
  dark: { canvas: "0A0F1A", surface: "111827", titleBg: "000D1A", headerBg: "0D1B2A", accentA: "00FF88", accentB: "00C87A", headerText: "00FF88", bodyText: "D0F0E8", mutedText: "6DBFA0", sectionBg: "0D1F2D", tableHeaderBg: "003D26", tableHeaderText: "00FF88", tableAltBg: "0A1A14", border: "1A4D3A" },
  forest: { canvas: "F0F7F4", surface: "FFFFFF", titleBg: "1B4D3E", headerBg: "256B53", accentA: "2E8B68", accentB: "5DB89A", headerText: "FFFFFF", bodyText: "162E25", mutedText: "4A7265", sectionBg: "E8F5F0", tableHeaderBg: "2E8B68", tableHeaderText: "FFFFFF", tableAltBg: "F0FAF5", border: "C8E8DC" },
  forestBalanced: { canvas: "F7F8F6", surface: "FFFFFF", titleBg: "183F34", headerBg: "245447", accentA: "1F8A70", accentB: "2E86AB", headerText: "FFFFFF", bodyText: "172B25", mutedText: "61736D", sectionBg: "EEF3F0", tableHeaderBg: "1F8A70", tableHeaderText: "FFFFFF", tableAltBg: "F4F7F5", border: "D7E0DC" },
  // 役員向け提案書: navy × orange × gray（お手本ベース）
  executiveProposal: { canvas: "FFFFFF", surface: "F4F5F8", titleBg: "13294B", headerBg: "13294B", accentA: "13294B", accentB: "F5821F", headerText: "FFFFFF", bodyText: "1D2435", mutedText: "6B7488", sectionBg: "EEF2F9", tableHeaderBg: "13294B", tableHeaderText: "FFFFFF", tableAltBg: "F7F8FB", border: "D9E0EC" },
};

// ─── 固定パレット（決め打ち1パターン）────────────────────────────────────────
// コンテキストによって variant キーで切り替えられる構造を維持しつつ、
// 現在は全 variant で同一の統一パレットを使用する。
type StrictPaletteKey = "calm" | "friendly" | "passionate" | "eco";

// 決め打ち7色（ここを変えるだけで全スライドに反映）
const PALETTE = {
  main:       "13294B",  // 濃紺の帯・章扉・濃色背景
  bg:         "FFFFFF",  // 本文スライドのベース背景（白）
  card:       "FFFFFF",  // カード/ボックスの塗り
  accent:     "F5821F",  // 重要・強調・CTA（明るいオレンジ）← 1か所で調整
  accent_bg:  "EEF2F9",  // 薄い強調ボックス（クリーム禁止）
  sub:        "6B7488",  // 補助・補足テキスト・枠線
  border:     "E4E8F0",  // 区切り線・細枠
} as const;

// ヘッダースタイル: 全幅バンド禁止 — 常に左アクセントバー+テキストのみ
const STRICT_HEADER_STYLE = "minimal" as const;

// ── 5パレット定義（gen_pptx_profile.py の PALETTES と役割キーを同期） ───────
// 役割キー: main / accent / accent_light / main_light / text_muted
// 背景・帯上文字は全パレット共通で FFFFFF 固定。
const PPTX_PALETTES: Record<string, {
  main: string; accent: string; accent_light: string; main_light: string; text_muted: string;
}> = {
  navy_orange:    { main: "13294B", accent: "F5821F", accent_light: "EEF2F9", main_light: "E4E8F0", text_muted: "6B7488" },
  forest_amber:   { main: "1B4D3E", accent: "F4A300", accent_light: "FBEFD5", main_light: "E3EDE8", text_muted: "5E6E66" },
  burgundy_gold:  { main: "8C1D18", accent: "E0A33B", accent_light: "F7ECD6", main_light: "F3E5E4", text_muted: "6E5A58" },
  teal_coral:     { main: "0E4D5C", accent: "EE6C4D", accent_light: "FBE6DE", main_light: "DCE9EC", text_muted: "5A6B70" },
  charcoal_terra: { main: "333333", accent: "C15F3C", accent_light: "F3E3DA", main_light: "ECECEA", text_muted: "6E6E6E" },
};

/**
 * パレット名から Palette 型を構築する。
 * 既存の buildStrictPalette / PALETTE と同じ役割マッピングを使用。
 */
function buildPaletteFromName(name: string): Palette {
  const src = PPTX_PALETTES[name] ?? PPTX_PALETTES["navy_orange"];
  return {
    canvas:          "FFFFFF",           // 背景は常に白
    surface:         "FFFFFF",           // カード白
    titleBg:         src.main,
    headerBg:        src.main,
    accentA:         src.main,           // 帯・アイコン円・構造主色
    accentB:         src.accent,         // 強調差し色（5-10%）
    headerText:      "FFFFFF",
    bodyText:        src.main,
    mutedText:       src.text_muted,
    sectionBg:       src.accent_light,   // コールアウト・強調ボックス背景
    tableHeaderBg:   src.main,
    tableHeaderText: "FFFFFF",
    tableAltBg:      src.main_light,     // テーブル交互行
    border:          src.main_light,     // 区切り線・カード枠
  };
}

// 7色 → Palette 14色マッピング（全 variant 共通）
function buildStrictPalette(_key: StrictPaletteKey): Palette {
  return {
    canvas:          PALETTE.bg,         // スライド背景
    surface:         PALETTE.card,       // カード/ボックス塗り
    titleBg:         PALETTE.main,
    headerBg:        PALETTE.main,
    accentA:         PALETTE.main,       // 構造的な主色（帯・アイコン円・番号）
    accentB:         PALETTE.accent,     // 差し色（重要数値・CTA・ライン） 5-10%
    headerText:      "FFFFFF",
    bodyText:        PALETTE.main,       // 本文・見出し → main #13294B
    mutedText:       PALETTE.sub,
    sectionBg:       PALETTE.accent_bg,  // 強調ボックスの薄オレンジ背景
    tableHeaderBg:   PALETTE.main,
    tableHeaderText: "FFFFFF",
    tableAltBg:      "DCE3EF",           // bg をわずかに暗くした交互行
    border:          PALETTE.border,
  };
}

// calm が優先されるキーワード（IT/AI/導入系は社内向けでも calm）
const CALM_KEYWORDS = [
  "役員","経営層","稟議","承認","提案","proposal","ir","board","executive",
  "ネイビー","濃紺","calm",
  "azurechat","azure chat","ai","生成ai","生成 ai","llm","gpt","chatgpt",
  "it","dx","デジタル","プラットフォーム","システム","導入","説明","紹介",
  "セキュリティ","ガバナンス","コンプライアンス","情報管理","クラウド",
];

// friendly が適用されるキーワード（柔らかいトーンが明確な場合のみ）
const FRIENDLY_KEYWORDS = [
  "研修","教育","オンボーディング","onboarding","training",
  "親しみ","やわらか","初学者","初心者","入門","フレンドリー","friendly",
  "採用","リクルート","recruitment",
];

/**
 * PromptIntent と instruction から最適な固定パターンを選択する。
 *
 * 優先順位:
 * 1. 明示指定 (calm/friendly/passionate/eco/①〜④)
 * 2. eco: 環境/廃棄物/サステナ
 * 3. passionate: 予算確保/決断/キャンペーン
 * 4. calm: 役員・IT・AI・導入・説明系 (広め)
 * 5. friendly: 研修・教育・親しみ明示 (calm キーワードが優先)
 * 6. default: calm
 */
function selectStrictPaletteKey(
  instructionText: string,
  intent?: PromptIntent
): StrictPaletteKey {
  const h = instructionText.toLowerCase();
  const purpose  = intent?.documentPurpose ?? "";
  const audience = intent?.audience ?? "";

  // 1. 明示指定
  if (containsAny(h, ["①","calm","冷静パターン","ネイビーオレンジ"])) return "calm";
  if (containsAny(h, ["②","friendly","親しみパターン"])) return "friendly";
  if (containsAny(h, ["③","passionate","熱心パターン"])) return "passionate";
  if (containsAny(h, ["④","eco","エコパターン"])) return "eco";

  // 2. eco: 環境・廃棄物処理
  if (containsAny(h, ["廃棄物","産廃","環境","エコ","recycl","sustainable","サステナ","リサイクル","インフラ"])) return "eco";

  // 3. passionate: 予算確保・決断・キャンペーン
  if (purpose === "campaign" ||
      containsAny(h, ["キャンペーン","予算確保","投資判断","決断","campaign","緊急","今すぐ","即決","強い訴求"])) return "passionate";

  // 4. calm: 役員 / IT・AI・導入系（「社内向け」でも calm 優先）
  if (audience === "executive" ||
      purpose === "proposal" || purpose === "ir" ||
      CALM_KEYWORDS.some((kw) => h.includes(kw))) return "calm";

  // 5. friendly: 研修・教育・親しみ明示（calm キーワードがない場合のみ）
  if (FRIENDLY_KEYWORDS.some((kw) => h.includes(kw))) return "friendly";

  // 6. default: calm
  return "calm";
}

function normalizeHex(input: string, fallback: string): string {
  const value = String(input ?? "").replace("#", "").trim();
  return /^[0-9A-Fa-f]{6}$/.test(value) ? value.toUpperCase() : fallback;
}

function hexLuminance(hex: string): number {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

// 0-360 の色相を返す（HEXから）
function hexHue(hex: string): number {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  if (d === 0) return 0;
  let h = 0;
  if (max === r) h = ((g - b) / d + 6) % 6;
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return h * 60;
}

// 2色の色相距離 (0-180)
function hueDist(hexA: string, hexB: string): number {
  const d = Math.abs(hexHue(hexA) - hexHue(hexB));
  return Math.min(d, 360 - d);
}

// 色が青系 (200-260°) か判定
function isBluishHex(hex: string): boolean {
  const h = hexHue(hex);
  return h >= 195 && h <= 265;
}

// 色が暖色 (orange/amber: 15-55°) か判定
function isWarmHex(hex: string): boolean {
  const h = hexHue(hex);
  return h >= 15 && h <= 55;
}

// Azure Blue 系の明るいブルーを判定（#3B82F6 / #214C8F / #93C5FD 相当）
function isAzureBluish(hex: string): boolean {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const h = hexHue(hex);
  // 青系色相 (210-250°) かつ青成分が赤より明らかに強い
  return h >= 210 && h <= 252 && b > r * 1.2 && b > g * 0.9;
}

// 鮮やかすぎるオレンジ判定（#F97316 / #F59E0B など）
function isVividOrange(hex: string): boolean {
  const h = hexHue(hex);
  const lum = hexLuminance(hex);
  return h >= 20 && h <= 45 && lum > 0.28; // 明るすぎるオレンジ
}

// ─── セマンティックカラートークン ─────────────────────────────────────────────
// tint = 白と混合（ratio 0=原色 〜 1=白）/ shade = 黒と混合（ratio 0=原色 〜 1=黒）

function tintHex(hex: string, ratio: number): string {
  const ch = (off: number) => Math.round(parseInt(hex.slice(off, off + 2), 16) + (255 - parseInt(hex.slice(off, off + 2), 16)) * ratio);
  return [ch(0), ch(2), ch(4)].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}

function shadeHex(hex: string, ratio: number): string {
  const ch = (off: number) => Math.round(parseInt(hex.slice(off, off + 2), 16) * (1 - ratio));
  return [ch(0), ch(2), ch(4)].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}

type SemanticTokens = {
  primaryDark: string;    // primary shade25%
  primary: string;        // 構造主色（= palette.accentA）
  primaryMid: string;     // primary tint30%
  primaryLight: string;   // primary tint70%
  blueGray: string;       // primary tint55%（低彩度灰青）
  surfaceBlue: string;    // primary tint92%（淡面）
  accentDark: string;     // accent shade18%
  accent: string;         // 差し色（= palette.accentB）
  accentLight: string;    // accent tint45%
  surfaceOrange: string;  // accent tint90%（淡面）
  textPrimary: string;    // 本文主色
  textMuted: string;      // 補助テキスト
  surfaceGray: string;    // 汎用淡灰面
  borderGray: string;     // 枠線
  white: string;
  info: string;           // シアン系補助色（#2A7DA3）
  surfaceInfo: string;    // info tint90%
  success: string;        // グリーン系補助色（#2E7D5B）
  surfaceSuccess: string; // success tint90%
};

/**
 * Palette の主色・差し色から tint/shade でセマンティックトークンを導出する。
 * navy_orange 専用値を持つが、他パレットでも同比率で自動導出するため汎用。
 */
function resolveSemanticTokens(palette: Palette): SemanticTokens {
  const primary = palette.accentA;  // navy_orange → "13294B"
  const accent  = palette.accentB;  // navy_orange → "F5821F"
  const INFO    = "2A7DA3";         // 落ち着いたシアン（全パレット共通補助色）
  const SUCCESS = "2E7D5B";         // 落ち着いたグリーン（全パレット共通補助色）
  return {
    primaryDark:    shadeHex(primary, 0.25),  // navy_orange → "0E1F38"
    primary,
    primaryMid:     tintHex(primary, 0.30),   // navy_orange → "5A6981"
    primaryLight:   tintHex(primary, 0.70),   // navy_orange → "B8BFC9"
    blueGray:       tintHex(primary, 0.55),   // navy_orange → "959FAE"
    surfaceBlue:    tintHex(primary, 0.92),   // navy_orange → "E8EAF0"
    accentDark:     shadeHex(accent, 0.18),   // navy_orange → "C96B19"
    accent,
    accentLight:    tintHex(accent, 0.45),    // navy_orange → "FABA84"
    surfaceOrange:  tintHex(accent, 0.90),    // navy_orange → "FEF3E9"
    textPrimary:    palette.bodyText,
    textMuted:      palette.mutedText,
    surfaceGray:    "F4F5F7",
    borderGray:     palette.border,
    white:          "FFFFFF",
    info:           INFO,
    surfaceInfo:    tintHex(INFO, 0.90),      // "EAF2F6"
    success:        SUCCESS,
    surfaceSuccess: tintHex(SUCCESS, 0.90),   // "EAF3EE"
  };
}

/**
 * パレット多様性を検証・補正する。
 * - 青単色パレット (blue monochrome) を検出して accentB を暖色に修正
 * - executive/proposal では Azure ブルー系を深い navy (#13294B) に強制置換
 * - executive/proposal で accentB が暖色でない場合に補正
 * - accentA と accentB の色相距離が近すぎる場合に補正
 * generateDesignBrief / createFallbackBrief / regenerateStyle 後に必ず通す。
 */
function normalizePaletteDiversity(
  palette: Palette,
  deckPurpose?: string,
  audience?: string
): Palette {
  const isExecProposal =
    (deckPurpose === "proposal" || deckPurpose === "ir") &&
    (audience === "executive");
  const isConservative = deckPurpose === "proposal" || deckPurpose === "ir";

  let p = { ...palette };
  const NAVY      = "13294B";   // Claude版のDeep Navy
  const ORANGE    = "F5821F";   // Accent Orange（明るいオレンジ）
  const BODY_TEXT = "1D2435";
  const MUTED     = "6B7488";
  const BORDER    = "D9E0EC";
  const SURFACE   = "F4F5F8";
  const WARM_AMBER = "D97706";

  // ── executive/proposal: Claude版パレットへ完全固定 ─────────────────────
  if (isExecProposal) {
    // titleBg / headerBg: Azure ブルー系 → deep navy
    if (isAzureBluish(p.titleBg) || hexLuminance(p.titleBg) > 0.10) {
      console.log(`[paletteDiversity] executive proposal: titleBg #${p.titleBg} → navy #${NAVY}`);
      p = { ...p, titleBg: NAVY };
    }
    if (isAzureBluish(p.headerBg) || hexLuminance(p.headerBg) > 0.10) {
      console.log(`[paletteDiversity] replaced azure blue #${p.headerBg} -> #${NAVY}`);
      p = { ...p, headerBg: NAVY, tableHeaderBg: NAVY };
    }
    // accentA: Azure ブルー系 → deep navy
    if (isAzureBluish(p.accentA) || (isBluishHex(p.accentA) && hexLuminance(p.accentA) > 0.12)) {
      console.log(`[paletteDiversity] replaced azure blue #${p.accentA} -> #${NAVY}`);
      p = { ...p, accentA: NAVY };
    }
    // accentB: 非暖色 / 鮮やかすぎるオレンジ → #DD6B17
    if (!isWarmHex(p.accentB) || isVividOrange(p.accentB)) {
      console.log(`[paletteDiversity] executive proposal: accentB #${p.accentB} → #${ORANGE}`);
      p = { ...p, accentB: ORANGE };
    }
    // accentB が水色系 (#93C5FD など) → グレーに
    if (isAzureBluish(p.accentB) && hexLuminance(p.accentB) > 0.3) {
      p = { ...p, accentB: ORANGE };
    }
    // bodyText / mutedText / border / surface を固定
    p = {
      ...p,
      bodyText:  BODY_TEXT,
      mutedText: MUTED,
      border:    BORDER,
      surface:   SURFACE,
    };
    console.log(`[paletteDiversity] executive proposal palette locked to Claude-like navy/orange`);
    return p;
  }

  // ── 非executive: 青単色検出 → accentB を暖色に ─────────────────────────
  const key5 = [p.titleBg, p.headerBg, p.accentA, p.accentB, p.sectionBg];
  const blueCount = key5.filter(isBluishHex).length;
  if (blueCount >= 4) {
    const warm = isConservative ? ORANGE : WARM_AMBER;
    console.log(`[paletteDiversity] blue monochrome detected (${blueCount}/5): accentB #${p.accentB} → #${warm}`);
    p = { ...p, accentB: warm };
  }

  // ── accentA vs accentB 色相距離チェック ─────────────────────────────────
  const dist = hueDist(p.accentA, p.accentB);
  if (dist < 25) {
    const warm = isConservative ? ORANGE : WARM_AMBER;
    console.log(`[paletteDiversity] accentA/B too similar (hue dist ${dist.toFixed(0)}°): accentB → #${warm}`);
    p = { ...p, accentB: warm };
  }

  // ── conservative: accentB が暖色でなければ修正 ──────────────────────────
  if (isConservative && !isWarmHex(p.accentB)) {
    console.log(`[paletteDiversity] proposal: accentB #${p.accentB} not warm → #${ORANGE}`);
    p = { ...p, accentB: ORANGE };
  }

  return p;
}

// ─── visualIntent → renderMode 変換 ─────────────────────────────────────────

type VisualRenderMode = "trust" | "appeal" | "data" | "process" | "default";

function parseVisualIntent(intent?: string): VisualRenderMode {
  if (!intent) return "default";
  if (/信頼|安心|実績|証明|認定|保証|stable|trust/.test(intent)) return "trust";
  if (/訴求|強み|差別化|インパクト|アピール|大きく|appeal/.test(intent)) return "appeal";
  if (/数値|データ|KPI|指標|数字|定量|metric|data/.test(intent)) return "data";
  if (/プロセス|フロー|手順|工程|ステップ|流れ|process/.test(intent)) return "process";
  return "default";
}

// renderMode → value フォントサイズ調整量
function intentValueFontAdj(mode: VisualRenderMode): number {
  if (mode === "data"   || mode === "appeal") return  3;
  if (mode === "trust")                       return  1;
  if (mode === "process")                     return -1;
  return 0;
}

// renderMode → note の表示重み（1.0=通常, >1=広め）
function intentNoteExpand(mode: VisualRenderMode): number {
  if (mode === "trust")   return 1.3;
  if (mode === "process") return 1.2;
  return 1.0;
}

// ─── density / textTreatment ヘルパー ────────────────────────────────────────

function densityFontAdj(density?: string): number {
  if (density === "low")  return  2;
  if (density === "high") return -2;
  return 0;
}

function densityMaxItems(density?: string, def = 4): number {
  if (density === "low")  return Math.min(def, 3);
  if (density === "high") return def;
  return Math.min(def, 4);
}

function densitySpacingMult(density?: string): number {
  if (density === "low")  return 1.25;
  if (density === "high") return 0.82;
  return 1.0;
}

// textTreatment が "explanatory" のときリードテキスト領域を拡張
function leadTextHeightBoost(textTreatment?: string): number {
  if (textTreatment === "explanatory") return 0.5;
  if (textTreatment === "short")       return -0.3;
  return 0;
}

/**
 * PromptIntent の guardrails に基づいて styleSpec を正規化する。
 * LLM が intent に反するスタイルを選んだ場合に上書きする。
 */
function normalizeStyleSpecFromIntent(
  styleSpec: DeckStyleSpec,
  intent: PromptIntent
): DeckStyleSpec {
  let { visualStyle, cardStyle } = styleSpec;
  const { documentPurpose, audience, designFreedom, styleGuardrails } = intent;

  const isConservativePurpose =
    documentPurpose === "proposal" ||
    documentPurpose === "company-intro" ||
    documentPurpose === "ir" ||
    audience === "executive";

  const isExpressivePurpose =
    documentPurpose === "recruitment" ||
    documentPurpose === "campaign";

  // conservative コンテキスト: modern-dark / playful / glass を禁止
  if (isConservativePurpose || designFreedom === "conservative") {
    if (!styleGuardrails.allowModernDark && visualStyle === "modern-dark") visualStyle = "corporate-light";
    if (!styleGuardrails.allowPlayful   && visualStyle === "playful")      visualStyle = "editorial";
    if (!styleGuardrails.allowGlass     && cardStyle  === "glass")         cardStyle  = "default";
  }

  // expressive コンテキスト: corporate-light が選ばれたら editorial か bold に引き上げ
  if (isExpressivePurpose && designFreedom === "expressive") {
    if (visualStyle === "corporate-light") visualStyle = "editorial";
  }

  return { ...styleSpec, visualStyle, cardStyle };
}

// データ欠如時にレイアウトタイプを安全にフォールバックする
function hasUsableColumns(columns?: PptxColumn[]): boolean {
  const usable = (columns ?? []).filter((col) =>
    Boolean(col.header?.trim()) || (col.bullets ?? []).some((b) => Boolean(b?.trim()))
  );
  return usable.length >= 2;
}

function hasUsableTableRows(rows?: string[][]): boolean {
  if (!rows || rows.length < 2) return false;
  const nonEmptyRows = rows.filter((row) => row.some((cell) => Boolean(cell?.trim())));
  return nonEmptyRows.length >= 2;
}

function resolveLayoutType(slide: PptxSlide): NonNullable<PptxSlide["layoutType"]> {
  const lt = slide.layoutType ?? "bullets";
  if (lt === "multi-column" && !hasUsableColumns(slide.columns)) {
    console.log(`[gen-pptx] multi-column→bullets fallback (no columns): "${slide.title}"`);
    return "bullets";
  }
  if (lt === "table" && !hasUsableTableRows(slide.tableRows)) {
    console.log(`[gen-pptx] table→bullets fallback (no tableRows): "${slide.title}"`);
    return "bullets";
  }
  if (lt === "process-cards") {
    const validSteps = (slide.steps ?? []).filter((s) => s.title?.trim() && s.body?.trim());
    if (validSteps.length < 2) {
      console.log(`[gen-pptx] process-cards→bullets fallback (insufficient steps): "${slide.title}"`);
      return "bullets";
    }
  }
  if (lt === "timeline") {
    const validSteps = (slide.steps ?? []).filter((s) => s.title?.trim() && s.body?.trim());
    if (validSteps.length < 3) {
      console.log(`[gen-pptx] timeline→bullets fallback (insufficient steps): "${slide.title}"`);
      return "bullets";
    }
  }
  if (lt === "metric-cards" && (!slide.metrics || slide.metrics.length === 0)) {
    console.log(`[gen-pptx] metric-cards→bullets fallback (no metrics): "${slide.title}"`);
    return "bullets";
  }
  if (lt === "diagram" && (!slide.visualBlocks || slide.visualBlocks.length === 0)) {
    console.log(`[gen-pptx] diagram→bullets fallback (no visualBlocks): "${slide.title}"`);
    return "bullets";
  }
  if (lt === "stat_callouts") {
    const hasCallouts = (slide.statCallouts ?? []).some((c) => c.value?.trim());
    const hasMetrics  = (slide.metrics ?? []).some((m) => m.value?.trim());
    if (!hasCallouts && !hasMetrics) {
      console.log(`[gen-pptx] stat_callouts→bullets fallback (no data): "${slide.title}"`);
      return "bullets";
    }
  }
  if (lt === "card_grid") {
    const hasCards = (slide.cards ?? []).some((c) => c.heading?.trim());
    const hasSteps = (slide.steps ?? []).some((st) => st.title?.trim() && st.body?.trim());
    if (!hasCards && !hasSteps) {
      console.log(`[gen-pptx] card_grid→bullets fallback (no cards/steps): "${slide.title}"`);
      return "bullets";
    }
  }
  if (lt === "icon_rows") {
    const hasCards = (slide.cards ?? []).some((c) => c.heading?.trim());
    const hasSteps = (slide.steps ?? []).some((st) => st.title?.trim());
    if (!hasCards && !hasSteps) {
      console.log(`[gen-pptx] icon_rows→bullets fallback (no rows): "${slide.title}"`);
      return "bullets";
    }
  }
  return lt;
}

// レイアウト変換のキーワードパターン
const TWO_COLUMN_KEYWORDS = ["比較","特徴","強み","違い","メリット","ポイント","vs","デメリット","対比","Before","After"];
const TABLE_KEYWORDS      = ["費用","料金","仕様","スペック","一覧","まとめ","価格","プラン","コスト","スケジュール","工程"];
const METRIC_KEYWORDS     = ["実績","KPI","数値","指標","業績","売上","成果","効果","達成"];
const PROCESS_KEYWORDS    = ["手順","ステップ","流れ","フロー","プロセス","導入","工程","手続き","方法"];

/**
 * PromptIntent の layoutDirectives に基づいてスライドのレイアウトを補正する。
 * validateAndRepairSlides の前に呼ぶことで、columns/tableRows 欠如による
 * bullets フォールバックを防ぐ。
 */
function applyPromptIntentToSlides(slides: PptxSlide[], intent: PromptIntent): PptxSlide[] {
  const ld = intent.layoutDirectives;
  if (!ld.preferTwoColumn && !ld.includeTables && !ld.preferMetrics && !ld.preferProcess) return slides;

  const alreadyHas = (lt: string) => slides.some((s) => s.layoutType === lt);

  let result = [...slides];

  // ── preferTwoColumn: multi-column がなければ適したスライドを変換 ──
  if (ld.preferTwoColumn && !alreadyHas("multi-column")) {
    const idx = result.findIndex((s, i) => {
      if (i === 0) return false; // 表紙は変換しない
      if (s.layoutType && s.layoutType !== "bullets") return false;
      const bullets = (s.bullets ?? []).filter((b) => b?.trim());
      if (bullets.length < 4) return false;
      return TWO_COLUMN_KEYWORDS.some((kw) => s.title.includes(kw));
    }) ?? result.findIndex((s, i) => {
      // キーワードなしのフォールバック: 中間の bullets スライド
      if (i === 0 || i === result.length - 1) return false;
      if (s.layoutType && s.layoutType !== "bullets") return false;
      return (s.bullets ?? []).filter((b) => b?.trim()).length >= 4;
    });

    if (idx >= 0) {
      const slide = result[idx];
      const bullets = (slide.bullets ?? []).filter((b) => b?.trim());
      const half = Math.ceil(bullets.length / 2);
      const col1 = bullets.slice(0, half);
      const col2 = bullets.slice(half);
      result[idx] = {
        ...slide,
        layoutType: "multi-column",
        columns: [
          { header: slide.title.length > 12 ? slide.title.slice(0, 10) + "①" : `${slide.title} (1)`, bullets: col1 },
          { header: slide.title.length > 12 ? slide.title.slice(0, 10) + "②" : `${slide.title} (2)`, bullets: col2 },
        ],
      };
      console.log(`[applyIntent] twoColumn: "${slide.title}" → multi-column (${col1.length}+${col2.length} bullets)`);
    }
  }

  // ── includeTables: table がなければ適したスライドを変換 ──
  if (ld.includeTables && !alreadyHas("table")) {
    const idx = result.findIndex((s, i) => {
      if (i === 0) return false;
      if (s.layoutType && s.layoutType !== "bullets") return false;
      const bullets = (s.bullets ?? []).filter((b) => b?.trim());
      if (bullets.length < 3) return false;
      return TABLE_KEYWORDS.some((kw) => s.title.includes(kw));
    });

    if (idx >= 0) {
      const slide = result[idx];
      const bullets = (slide.bullets ?? []).filter((b) => b?.trim());
      // bullets に "：" or ":" があれば key:value テーブルに、なければ番号付き一覧に
      const hasColon = bullets.some((b) => b.includes("：") || b.includes(":"));
      const tableRows: string[][] = hasColon
        ? [["項目", "内容"], ...bullets.map((b) => {
            const ci = b.indexOf("：") >= 0 ? b.indexOf("：") : b.indexOf(":");
            return ci > 0 ? [b.slice(0, ci).trim(), b.slice(ci + 1).trim()] : [b, ""];
          })]
        : [["No.", slide.title], ...bullets.map((b, i) => [String(i + 1), b])];

      result[idx] = { ...slide, layoutType: "table", tableRows };
      console.log(`[applyIntent] includeTables: "${slide.title}" → table (${tableRows.length} rows)`);
    }
  }

  return result;
}

/**
 * LLM由来データの一括正規化。validateAndRepairSlides の前に通す。
 * 数値・null・undefined が文字列・配列フィールドに混入しても
 * .trim() / .length 等で描画クラッシュしないよう強制変換する。
 */
function normalizeSlidesForPptx(slides: PptxSlide[]): PptxSlide[] {
  const s  = (v: unknown): string   => (v == null ? "" : String(v));
  const sa = (v: unknown): string[] =>
    Array.isArray(v) ? v.map(s) : typeof v === "string" ? [v] : [];

  return slides.map((slide) => ({
    ...slide,
    title:       s(slide.title),
    bullets:     sa(slide.bullets),
    leadText:    slide.leadText    != null ? s(slide.leadText)    : undefined,
    subtitle:    slide.subtitle    != null ? s(slide.subtitle)    : undefined,
    visualIntent:slide.visualIntent!= null ? s(slide.visualIntent): undefined,
    callout: slide.callout
      ? { title: s(slide.callout.title), body: s(slide.callout.body) }
      : undefined,
    metrics: slide.metrics?.map((m) => ({
      ...m,
      label:        s(m.label),
      value:        s(m.value),
      displayValue: m.displayValue != null ? s(m.displayValue) : undefined,
      note:         m.note         != null ? s(m.note)         : undefined,
      iconKey:      m.iconKey      != null ? s(m.iconKey)      : undefined,
    })),
    steps: slide.steps?.map((st) => ({
      ...st,
      title:  s(st.title),
      body:   s(st.body),
      iconKey: st.iconKey != null ? s(st.iconKey) : undefined,
    })),
    benefits: slide.benefits ? sa(slide.benefits) : undefined,
    columns:  slide.columns?.map((col) => ({
      ...col,
      header:  s(col.header),
      bullets: sa(col.bullets),
    })),
    tableRows: slide.tableRows?.map((row) =>
      Array.isArray(row) ? row.map(s) : []
    ),
  }));
}

/**
 * 描画前品質ゲート。
 * 空スライド・データ欠如スライドを bullets 化または削除し、
 * 安全に描画できる状態に修復して返す。
 */
function validateAndRepairSlides(slides: PptxSlide[]): PptxSlide[] {
  const result: PptxSlide[] = [];

  for (const slide of slides) {
    if (!slide.title?.trim()) {
      console.log(`[gen-pptx] repair: removed (empty title)`);
      continue;
    }

    const lt = slide.layoutType ?? "bullets";
    let repairTo: "bullets" | "delete" | null = null;
    let reason = "";

    switch (lt) {
      case "multi-column":
        if (!hasUsableColumns(slide.columns)) {
          repairTo = "bullets";
          reason = "columns が空または不足";
        }
        break;

      case "table":
        if (!hasUsableTableRows(slide.tableRows)) {
          repairTo = "bullets";
          reason = "tableRows が空または不足";
        }
        break;

      case "process-cards": {
        const validSteps = (slide.steps ?? []).filter((s) => s.title?.trim() && s.body?.trim());
        if (validSteps.length < 2) {
          repairTo = "bullets";
          reason = "有効な steps が 2 件未満";
        }
        break;
      }

      case "timeline": {
        const validSteps = (slide.steps ?? []).filter((s) => s.title?.trim() && s.body?.trim());
        if (validSteps.length < 3) {
          repairTo = "bullets";
          reason = "有効な steps が 3 件未満";
        }
        break;
      }

      case "metric-cards":
        if (!slide.metrics || slide.metrics.length === 0) {
          repairTo = "bullets";
          reason = "metrics が空";
        }
        break;

      case "diagram":
        if (!slide.visualBlocks || slide.visualBlocks.length === 0) {
          repairTo = "bullets";
          reason = "visualBlocks が空";
        }
        break;

      case "company-overview": {
        const hasLead = Boolean(slide.leadText?.trim());
        const hasMetrics = slide.metrics && slide.metrics.length > 0;
        if (!hasLead && !hasMetrics) {
          repairTo = "bullets";
          reason = "leadText も metrics も空";
        }
        break;
      }
    }

    // bullets 化後または元から bullets で、bullet 本文もすべて空 → 削除
    if (repairTo === "bullets" || lt === "bullets") {
      const validBullets = (slide.bullets ?? []).filter((b) => b?.trim());
      if (validBullets.length === 0) {
        repairTo = "delete";
        reason = reason ? `${reason} + bullets も空` : "bullets が空";
      }
    }

    if (repairTo === "delete") {
      console.log(`[gen-pptx] repair: deleted "${slide.title}" (${reason})`);
      continue;
    }

    if (repairTo === "bullets") {
      console.log(`[gen-pptx] repair: "${slide.title}" → bullets (${reason})`);
      result.push({ ...slide, layoutType: "bullets" });
    } else {
      result.push(slide);
    }
  }

  if (result.length !== slides.length) {
    console.log(`[gen-pptx] validateAndRepairSlides: ${slides.length} → ${result.length} slides`);
  }
  return result;
}

// cardStyle に応じたカード描画パラメータを返す汎用ヘルパー
type CardStyleProps = {
  fill: { color: string; transparency?: number };
  line: { color: string; width: number; transparency?: number };
  shadow?: { type: "outer"; color: string; blur: number; angle: number; opacity: number };
  rectRadius: number;
};

function getCardStyleProps(theme: Theme, _accentOverride?: string): CardStyleProps {
  const { palette } = theme;
  // デザイン原則: 影・グラデーション・装飾なし。フラットで端正に。
  // cardStyle バリアントは統一パレット下では全て同一の白カード+細枠に統一
  return {
    fill: { color: palette.surface },         // card = #FFFFFF
    line: { color: palette.border, width: 0.7 }, // border = #E4E8F0
    rectRadius: 0.06,                         // 控えめな角丸
  };
}

// visualStyle に応じた余白乗数（editorial/minimal は広め）
function getSpacingMult(theme: Theme): number {
  if (theme.visualStyle === "editorial" || theme.visualStyle === "minimal") return 1.18;
  if (theme.visualStyle === "bold") return 0.92;
  return 1.0;
}

// cardStyle が filled の時は背景が濃色なのでテキストを明色にする
function getCardBodyColor(theme: Theme): string {
  return theme.cardStyle === "filled" ? theme.palette.headerText : theme.palette.bodyText;
}
function getCardMutedColor(theme: Theme): string {
  return theme.cardStyle === "filled" ? theme.palette.accentB : theme.palette.mutedText;
}

function ensureDarkIconBg(candidate: string, fallback: string): string {
  return hexLuminance(candidate) > 0.35 ? fallback : candidate;
}

function isGreenish(hex: string): boolean {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return g > r * 1.12 && g > b * 1.08;
}

function correctOvergreenPalette(palette: Palette, deckPurpose?: string): Palette {
  // 採用・IR・研修・internal など非環境用途はデザインエージェントの選択を尊重して修正しない
  const envPurposes = new Set(["company-intro", "other", undefined]);
  if (!envPurposes.has(deckPurpose)) {
    console.log(`[palette] skip correction for deckPurpose=${deckPurpose}`);
    return palette;
  }
  const key5 = [palette.titleBg, palette.headerBg, palette.accentA, palette.accentB, palette.sectionBg];
  const greenCount = key5.filter(isGreenish).length;
  console.log(`[palette] accentA=#${palette.accentA} accentB=#${palette.accentB} titleBg=#${palette.titleBg} greenCount=${greenCount}`);
  if (greenCount >= 4) {
    console.log("[palette] correction: too many greens → forestBalanced");
    return DEFAULT_PALETTES.forestBalanced;
  }
  if (isGreenish(palette.accentB) && isGreenish(palette.accentA)) {
    console.log(`[palette] correction: accentB=#${palette.accentB} is green → forcing professional blue 2E86AB`);
    return { ...palette, accentB: "2E86AB" };
  }
  console.log("[palette] no correction needed");
  return palette;
}

function truncateText(value: string, max: number): string {
  const text = String(value ?? "").trim();
  if (text.length <= max) return text;
  // 文末として扱える境界のみ（読点「、」とカンマ「,」は節の途中のため除外）
  const boundaries = "。．.！!？?）」』）";
  for (let i = max; i >= Math.floor(max * 0.5); i--) {
    if (boundaries.includes(text[i])) return text.slice(0, i + 1).trim();
  }
  return `${text.slice(0, max).trim()}…`;
}

function containsAny(value: string, tokens: string[]): boolean {
  return tokens.some((token) => value.includes(token));
}

function resolvePaletteKeyFromPrompt(input: string): keyof typeof DEFAULT_PALETTES {
  const hint = input.toLowerCase();
  // 役員向け提案 → executiveProposal を優先
  if (containsAny(hint, ["役員", "経営層", "executive", "board", "社長", "ceo"]) &&
      containsAny(hint, ["提案", "proposal", "社内提案", "承認"])) return "executiveProposal";
  if (containsAny(hint, ["dark", "matrix", "マトリックス", "cyber", "サイバー", "future", "未来", "クール", "cool", "近未来", "sf", "noir", "ノワール", "black", "ブラック", "futuristic", "techno", "テクノ", "night", "夜"])) return "dark";
  if (containsAny(hint, ["pastel", "soft", "gentle", "sweet", "やわらか", "パステル"])) return "pastel";
  if (containsAny(hint, ["pop", "playful", "vivid", "colorful", "ポップ", "元気"])) return "pop";
  if (containsAny(hint, ["red", "crimson", "scarlet", "赤"])) return "red";
  if (containsAny(hint, ["廃棄物", "産廃", "環境", "サステナ", "sustainability", "エコシステム", "インフラ", "リサイクル", "recycl"])) return "forestBalanced";
  if (containsAny(hint, ["forest"])) return "forest";
  if (containsAny(hint, ["green", "eco", "natural", "organic", "緑", "エコ"])) return "green";
  if (containsAny(hint, ["gold", "yellow", "黄色", "黄", "luxury", "premium", "golden", "金", "高級", "オレンジ", "orange", "橙"])) return "gold";
  if (containsAny(hint, ["blue", "青", "navy", "ネイビー", "水色", "indigo"])) return "blue";
  return "blue";
}

function getFontScaleMultiplier(fontScale?: DeckPreferencesInput["fontScale"]): number {
  switch (fontScale) {
    case "small":
      return 0.92;
    case "large":
      return 1.12;
    case "xlarge":
      return 1.22;
    default:
      return 1;
  }
}

function applyAccentOverride(palette: Palette, accentColor?: string): Palette {
  const key = String(accentColor ?? "").toLowerCase();
  if (containsAny(key, ["dark", "matrix", "cyber", "black", "night"])) return DEFAULT_PALETTES.dark;
  if (containsAny(key, ["blue", "青"])) return DEFAULT_PALETTES.blue;
  if (containsAny(key, ["red", "赤"])) return DEFAULT_PALETTES.red;
  if (containsAny(key, ["green", "eco", "緑", "エコ"])) return DEFAULT_PALETTES.green;
  if (containsAny(key, ["gold", "yellow", "金"])) return DEFAULT_PALETTES.gold;
  return palette;
}

function getFontFace(input: string, explicitFontFace?: string): string {
  if (explicitFontFace?.trim()) return explicitFontFace.trim();
  const hint = input.toLowerCase();
  if (containsAny(hint, ["yu gothic", "游ゴシック"])) return "Yu Gothic";
  if (containsAny(hint, ["yu mincho", "游明朝", "明朝"])) return "Yu Mincho";
  if (containsAny(hint, ["arial"])) return "Arial";
  return "Meiryo";
}

function localizeLabel(
  key: "coverKicker" | "summary" | "points" | "details" | "pages",
  theme: Theme
): string {
  if (!theme.useJapaneseLabels) {
    if (key === "coverKicker") return "AI-ENHANCED PRESENTATION";
    if (key === "summary") return "SUMMARY";
    if (key === "points") return "POINTS";
    if (key === "details") return "DETAILS";
    return "pages";
  }
  if (key === "coverKicker") return "AI生成プレゼンテーション";
  if (key === "summary") return "要約";
  if (key === "points") return "要点";
  if (key === "details") return "詳細";
  return "ページ";
}

function inferFallbackStyleSpec(instructionText: string): DeckStyleSpec {
  const h = instructionText.toLowerCase();
  const deckPurpose: DeckStyleSpec["deckPurpose"] =
    containsAny(h, ["採用", "recruit", "人材", "求人", "hiring"]) ? "recruitment" :
    containsAny(h, ["キャンペーン", "イベント", "告知", "campaign"]) ? "campaign" :
    containsAny(h, ["提案", "proposal", "企画"]) ? "proposal" :
    containsAny(h, ["会社紹介", "company profile", "紹介資料", "会社概要"]) ? "company-intro" :
    containsAny(h, ["研修", "training", "教育", "onboard"]) ? "training" :
    containsAny(h, ["分析", "調査", "市場", "analysis", "リサーチ"]) ? "analysis" :
    containsAny(h, ["ir", "投資家", "株主", "決算"]) ? "ir" :
    containsAny(h, ["社内", "internal", "報告", "レポート"]) ? "internal" : "other";

  const visualStyle: DeckStyleSpec["visualStyle"] =
    containsAny(h, ["dark", "matrix", "cyber", "night", "modern", "tech", "dx", "デジタル"]) ? "modern-dark" :
    containsAny(h, ["pop", "playful", "cute", "ポップ", "親しみ"]) ? "playful" :
    containsAny(h, ["minimal", "simple", "clean", "ミニマル", "シンプル"]) ? "minimal" :
    containsAny(h, ["bold", "impact", "強調", "インパクト"]) ? "bold" :
    deckPurpose === "recruitment" ? "editorial" : "corporate-light";

  return {
    deckPurpose,
    visualStyle,
    cardStyle: visualStyle === "modern-dark" ? "glass" : visualStyle === "minimal" ? "flat" : "default",
    headerStyle: "minimal",  // 全幅バンド禁止 — 常にフラットヘッダー
  };
}

function createFallbackBrief(
  _title: string,
  slides: PptxSlide[],
  instructionText: string,
  prefs?: DeckPreferencesInput
): DeckDesignBrief {
  const paletteKey = resolvePaletteKeyFromPrompt(instructionText);
  let palette = applyAccentOverride(DEFAULT_PALETTES[paletteKey], prefs?.accentColor);
  const styleSpec = inferFallbackStyleSpec(instructionText);

  // fallback 時も normalizePaletteDiversity を通す
  palette = normalizePaletteDiversity(palette, styleSpec.deckPurpose, undefined);
  if (paletteKey === "blue") {
    console.log(`[paletteDiversity] fallback palette normalized for ${styleSpec.deckPurpose}`);
  }
  const visualCycle: SlideVisualType[] = ["spotlight", "cards", "process", "comparison", "editorial", "timeline"];
  return {
    palette,
    coverKicker: styleSpec.deckPurpose === "proposal" ? "PROPOSAL DECK"
      : styleSpec.deckPurpose === "company-intro" ? "COMPANY PROFILE"
      : styleSpec.deckPurpose === "recruitment" ? "RECRUITING"
      : "",
    coverSubtitle: "",
    footerNote: "",
    mood: instructionText || "editorial",
    styleSpec,
    visualHints: slides.map((slide, index) => {
      // resolveLayoutType でフォールバック後の実効レイアウトを使う
      const lt = resolveLayoutType(slide);
      let visualType: SlideVisualType;
      if (lt === "table") visualType = "table";
      else if (lt === "multi-column") visualType = "comparison";
      else if (lt === "diagram" || lt === "process-cards") visualType = "process";
      else if (lt === "timeline") visualType = "timeline";
      else if (lt === "company-overview" || lt === "metric-cards" || lt === "stat_callouts") visualType = "cards";
      else if (lt === "closing") visualType = "spotlight";
      else if (lt === "conversation") visualType = "editorial";
      else if (lt === "card_grid" || lt === "icon_rows") visualType = "process";
      else visualType = visualCycle[index % visualCycle.length];
      return {
        title: slide.title,
        visualType,
        emphasis: (Array.isArray(slide.bullets) && slide.bullets[0]) || slide.title,
      };
    }),
  };
}

async function generateDesignBrief(
  title: string,
  slides: PptxSlide[],
  instructionText: string,
  prefs?: DeckPreferencesInput,
  styleHint?: string,          // Vision LLM からのスタイル再生成ヒント
  intent?: PromptIntent        // ユーザー意図の構造化データ
): Promise<DeckDesignBrief> {
  const fallback = createFallbackBrief(title, slides, instructionText, prefs);

  try {
    const openai = OpenAIInstance();

    // ── PromptIntent に基づくカラー・スタイル制約を構築 ──
    const intentConstraints: string[] = [];
    if (intent) {
      const { documentPurpose, audience, designFreedom, colorDirectives, styleGuardrails, toneKeywords } = intent;
      intentConstraints.push(`\n=== USER INTENT (MUST follow these constraints) ===`);
      intentConstraints.push(`documentPurpose: ${documentPurpose}, audience: ${audience}, designFreedom: ${designFreedom}`);
      if (toneKeywords.length > 0) intentConstraints.push(`toneKeywords: ${toneKeywords.join(", ")}`);
      if (colorDirectives?.primary) intentConstraints.push(`PRIMARY COLOR: #${colorDirectives.primary} — use this as titleBg or accentA`);
      if (colorDirectives?.accent)  intentConstraints.push(`ACCENT COLOR: #${colorDirectives.accent} — use as accentB or accentA`);
      if (colorDirectives?.background) intentConstraints.push(`BACKGROUND: #${colorDirectives.background} — use as canvas`);
      if (!styleGuardrails.allowModernDark) intentConstraints.push(`FORBIDDEN: visualStyle='modern-dark' (not appropriate for this purpose/audience)`);
      if (!styleGuardrails.allowPlayful)   intentConstraints.push(`FORBIDDEN: visualStyle='playful'`);
      if (!styleGuardrails.allowGlass)     intentConstraints.push(`FORBIDDEN: cardStyle='glass'`);
      if (styleGuardrails.maxAccentIntensity === "low") intentConstraints.push(`Keep colors muted/professional — no vivid neon or heavy saturation`);
      if (designFreedom === "expressive") intentConstraints.push(`ALLOWED: bold/modern-dark/playful/glass — user wants expressive, impactful design`);
    }

    const systemPrompt = [
      "You are a presentation Design Agent. Given a title, instruction, and optional constraints, you decide:",
      "1. A color palette that fits the purpose",
      "2. A visual style matching audience and goal",
      "",
      "Return compact JSON with EXACTLY these fields:",
      "palette: { canvas, surface, titleBg, headerBg, accentA, accentB, headerText, bodyText, mutedText, sectionBg, tableHeaderBg, tableHeaderText, tableAltBg, border } — all 6-char hex WITHOUT #",
      "coverKicker: 3-5 words UPPERCASE label",
      "coverSubtitle: under 60 chars — viewer-facing benefit statement, NOT a title echo",
      "mood: 2-4 descriptive words",
      "styleSpec: { deckPurpose, visualStyle, cardStyle, headerStyle }",
      "",
      "styleSpec values:",
      "  deckPurpose: 'recruitment'|'proposal'|'company-intro'|'training'|'ir'|'internal'|'other'",
      "  visualStyle: 'corporate-light'|'modern-dark'|'editorial'|'playful'|'minimal'|'bold'",
      "  cardStyle: 'default'|'filled'|'glass'|'flat'",
      "  headerStyle: 'minimal'|'accent-line'  (NEVER 'band')",
      "",
      "DESIGN RULES:",
      "- deckPurpose='recruitment' → vibrant, energetic colors. Consider indigo/violet, warm orange, or bold teal.",
      "- deckPurpose='proposal' → authoritative but modern. Dark navy or deep teal with bright accent.",
      "- deckPurpose='company-intro' → trustworthy, balanced. Industry-appropriate hues.",
      "- deckPurpose='ir' → conservative, professional. Blue/grey tones.",
      "- deckPurpose='training' → clear, friendly. Warm blues or greens.",
      "- CRITICAL: accentA and accentB MUST be clearly different hues (hue distance > 30°).",
      "- STRICTLY FORBIDDEN: accentA and accentB both in blue/indigo range (#3B82F6/#93C5FD type blue monochrome).",
      "- STRICTLY FORBIDDEN: Using only blue shades for the entire palette. Every palette needs contrast.",
      "- For deckPurpose='proposal' or 'ir': accentB MUST be warm — orange (#F5821F), amber (#D97706), or muted gold (#C98A2E). NOT another shade of blue.",
      "- Executive proposal baseline: titleBg~13294B (dark navy), accentA~13294B, accentB~F5821F (orange), bodyText~1D2435, sectionBg~EEF2F9, canvas~FFFFFF.",
      "- Orange is used at ~5% — only on: metric values, step numbers, accent lines, CTAs. NOT full card fills.",
      "- Do NOT default to green unless the content is specifically environmental/ecological.",
      "- STRICTLY FORBIDDEN: headerStyle='band'. Full-width header color fills make slides look like AI-generated templates.",
      "- headerStyle='minimal' means left accent bar + title text only — no background fill.",
      "- Keep total JSON under 500 tokens.",
      ...intentConstraints,
      styleHint ? `\nStyle regeneration hint from reviewer: "${styleHint}" — honor this direction.` : "",
    ].filter(Boolean).join("\n");

    const userContent = [
      `title: ${title}`,
      `instruction: ${instructionText.slice(0, 400)}`,
      intent ? `promptIntent: ${JSON.stringify({ documentPurpose: intent.documentPurpose, audience: intent.audience, designFreedom: intent.designFreedom, colorDirectives: intent.colorDirectives })}` : "",
    ].filter(Boolean).join("\n");

    let res = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_PPT_DEPLOYMENT_NAME ?? process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userContent },
      ],
      response_format: { type: "json_object" },
      max_completion_tokens: 2000,
    });

    if (res.choices[0]?.finish_reason === "length") {
      console.warn("[designBrief] truncated → retrying with higher token limit");
      res = await openai.chat.completions.create({
        model: process.env.AZURE_OPENAI_PPT_DEPLOYMENT_NAME ?? process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userContent },
        ],
        response_format: { type: "json_object" },
        max_completion_tokens: 8000,
      });
      if (res.choices[0]?.finish_reason === "length") {
        console.warn("[designBrief] retry also truncated → using fallback");
        return fallback;
      }
    }

    const parsed = JSON.parse(res.choices[0]?.message?.content ?? "{}");
    const p = parsed?.palette ?? {};
    const palette: Palette = {
      canvas:          normalizeHex(p.canvas,          fallback.palette.canvas),
      surface:         normalizeHex(p.surface,         fallback.palette.surface),
      titleBg:         normalizeHex(p.titleBg,         fallback.palette.titleBg),
      headerBg:        normalizeHex(p.headerBg,        fallback.palette.headerBg),
      accentA:         normalizeHex(p.accentA,         fallback.palette.accentA),
      accentB:         normalizeHex(p.accentB,         fallback.palette.accentB),
      headerText:      normalizeHex(p.headerText,      fallback.palette.headerText),
      bodyText:        normalizeHex(p.bodyText,        fallback.palette.bodyText),
      mutedText:       normalizeHex(p.mutedText,       fallback.palette.mutedText),
      sectionBg:       normalizeHex(p.sectionBg,       fallback.palette.sectionBg),
      tableHeaderBg:   normalizeHex(p.tableHeaderBg,   fallback.palette.tableHeaderBg),
      tableHeaderText: normalizeHex(p.tableHeaderText, fallback.palette.tableHeaderText),
      tableAltBg:      normalizeHex(p.tableAltBg,      fallback.palette.tableAltBg),
      border:          normalizeHex(p.border,          fallback.palette.border),
    };

    // styleSpec を LLM 出力から取得、不正値はフォールバック
    const rawSpec = parsed?.styleSpec ?? {};
    const VALID_PURPOSES = new Set(["recruitment","proposal","company-intro","training","analysis","ir","internal","campaign","other"]);
    const VALID_VSTYLES  = new Set(["corporate-light","modern-dark","editorial","playful","minimal","bold"]);
    const VALID_CARDS    = new Set(["default","filled","glass","flat"]);
    const VALID_HEADERS  = new Set(["band","minimal","accent-line"]);
    const styleSpec: DeckStyleSpec = {
      deckPurpose:  VALID_PURPOSES.has(rawSpec.deckPurpose) ? rawSpec.deckPurpose : fallback.styleSpec.deckPurpose,
      visualStyle:  VALID_VSTYLES.has(rawSpec.visualStyle)  ? rawSpec.visualStyle  : fallback.styleSpec.visualStyle,
      cardStyle:    VALID_CARDS.has(rawSpec.cardStyle)      ? rawSpec.cardStyle    : fallback.styleSpec.cardStyle,
      headerStyle:  VALID_HEADERS.has(rawSpec.headerStyle)  ? rawSpec.headerStyle  : fallback.styleSpec.headerStyle,
    };

    // PromptIntent の guardrails で styleSpec を正規化（colorDirectives 適用より先に行う）
    const normalizedStyleSpec = intent ? normalizeStyleSpecFromIntent(styleSpec, intent) : styleSpec;

    // PromptIntent の colorDirectives を palette に反映
    // ただし executive/proposal では Azure ブルー系を navy に丸める
    let finalPalette = { ...palette };
    const isExecCtx = normalizedStyleSpec.deckPurpose === "proposal" || normalizedStyleSpec.deckPurpose === "ir";
    if (intent?.colorDirectives) {
      const cd = intent.colorDirectives;
      if (cd.primary) {
        // executive context で明るいブルー系を指定された場合は deep navy に丸める
        const pHex = normalizeHex(cd.primary, finalPalette.accentA);
        const resolvedPrimary = isExecCtx && isAzureBluish(pHex) ? "13294B" : pHex;
        finalPalette.accentA = resolvedPrimary;
        finalPalette.titleBg = resolvedPrimary;
      }
      if (cd.accent) {
        // executive context で鮮やかすぎるオレンジは PALETTE.accent に丸める
        const aHex = normalizeHex(cd.accent, finalPalette.accentB);
        const resolvedAccent = isExecCtx && isVividOrange(aHex) ? PALETTE.accent : aHex;
        finalPalette.accentB = resolvedAccent;
      }
      if (cd.background) { finalPalette.canvas = normalizeHex(cd.background, finalPalette.canvas); }
    }

    console.log(`[designBrief] purpose=${normalizedStyleSpec.deckPurpose} style=${normalizedStyleSpec.visualStyle} card=${normalizedStyleSpec.cardStyle} header=${normalizedStyleSpec.headerStyle} accentA=#${finalPalette.accentA} accentB=#${finalPalette.accentB}`);

    const correctedPalette = normalizePaletteDiversity(
      correctOvergreenPalette(applyAccentOverride(finalPalette, prefs?.accentColor), normalizedStyleSpec.deckPurpose),
      normalizedStyleSpec.deckPurpose,
      intent?.audience
    );
    return {
      palette: correctedPalette,
      coverKicker:   String(parsed?.coverKicker   ?? "").trim() || fallback.coverKicker,
      coverSubtitle: String(parsed?.coverSubtitle ?? "").trim() || fallback.coverSubtitle,
      footerNote:    fallback.footerNote,
      mood:          String(parsed?.mood ?? "").trim() || fallback.mood,
      visualHints:   fallback.visualHints,
      styleSpec:     normalizedStyleSpec,
    };
  } catch (error) {
    console.warn("[gen-pptx] design brief fallback:", error);
    return fallback;
  }
}

function resolveTheme(
  designBrief: DeckDesignBrief,
  instructionText: string,
  prefs?: DeckPreferencesInput,
  explicitFontFace?: string
): Theme {
  const palette = applyAccentOverride(designBrief.palette, prefs?.accentColor);
  const styleSpec = designBrief.styleSpec;

  // styleSpec.visualStyle を優先。キーワードは補助フォールバックのみ
  const lowered = instructionText.toLowerCase();
  const execMode  = styleSpec.deckPurpose === "ir" || styleSpec.deckPurpose === "proposal" ||
                    containsAny(lowered, ["executive", "board", "役員", "経営"]);
  const playfulMode = styleSpec.visualStyle === "playful" ||
                      containsAny(lowered, ["pop", "ポップ", "親しみ", "やわらか"]);
  const minimalMode = styleSpec.visualStyle === "minimal" ||
                      containsAny(lowered, ["minimal", "ミニマル", "シンプル"]);

  const scale = getFontScaleMultiplier(prefs?.fontScale ?? (execMode ? "large" : "medium"));
  const useJapaneseLabels = prefs?.language === "ja" || prefs?.avoidEnglishLabels === true;
  return {
    palette,
    fontFace: getFontFace(instructionText, explicitFontFace),
    titleFontSize: Math.round(24 * scale),
    bodyFontSize: Math.round(17 * scale),
    smallFontSize: Math.round(11 * scale),
    execMode,
    playfulMode,
    minimalMode,
    language: prefs?.language === "en" ? "en" : "ja",
    useJapaneseLabels,
    cardStyle:    styleSpec.cardStyle,
    headerStyle:  styleSpec.headerStyle,
    visualStyle:  styleSpec.visualStyle,
    deckPurpose:  styleSpec.deckPurpose,
  };
}

/**
 * pptxgenjs は a:latin しか設定しないため日本語が Yu Gothic 等にフォールバックする。
 * PPTX ZIP を書き換えて a:ea typeface を a:latin と同じフォントに揃える。
 */
async function patchEastAsianFont(buffer: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter(
    (name) => name.startsWith("ppt/") && name.endsWith(".xml") && !zip.files[name].dir
  );
  for (const name of targets) {
    const xml = await zip.files[name].async("string");
    if (!xml.includes("<a:latin")) continue;
    // 既存の a:ea 宣言を一旦除去してから a:latin の直後に再挿入
    let patched = xml.replace(/<a:ea typeface="[^"]*"\/>/g, "");
    patched = patched.replace(
      /<a:latin typeface="([^"]+)"\/>/g,
      (_, f) => `<a:latin typeface="${f}"/><a:ea typeface="${f}"/>`
    );
    zip.file(name, patched);
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE",
    compressionOptions: { level: 6 } }) as Promise<Buffer>;
}

async function uploadPptxToBlob(buffer: Buffer, blobKey: string, displayFileName?: string): Promise<string> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const containerName = "pptx";
  const blobServiceClient = BlobServiceClient.fromConnectionString(
    `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
  const containerClient = blobServiceClient.getContainerClient(containerName);
  await containerClient.createIfNotExists({ access: "blob" });
  // Blob key は ASCII のみ（URLを短く保つ）。DL時のファイル名は Content-Disposition で指定
  const blockBlobClient = containerClient.getBlockBlobClient(blobKey);
  const encodedFileName = encodeURIComponent(displayFileName ?? blobKey);
  await blockBlobClient.uploadData(buffer, {
    blobHTTPHeaders: {
      blobContentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      blobContentDisposition: `attachment; filename*=UTF-8''${encodedFileName}`,
    },
  });
  // generateSasUrl は BlockBlobClient が StorageSharedKeyCredential を持つ場合のみ使用可能
  // fromConnectionString（アカウントキー含む）で作成した場合は使用可能
  const sasUrl = await blockBlobClient.generateSasUrl({
    expiresOn: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
    permissions: BlobSASPermissions.parse("r"),
  });
  return sasUrl;
}

const W = 13.33;
const H = 7.5;
const HEADER_H = 1.05;

function addHeaderBand(s: PptxGenJS.Slide, title: string, theme: Theme) {
  // 装飾バー完全撤廃 — テキストのみ
  s.addText(title, {
    x: 0.40, y: 0.08, w: W - 0.80, h: HEADER_H - 0.14,
    fontSize: theme.titleFontSize,
    fontFace: theme.fontFace,
    bold: true,
    color: theme.palette.bodyText,
    valign: "middle",
  });
}

function addChrome(_s: PptxGenJS.Slide, _theme: Theme) {
  // フッターバンド撤廃 — クリーンなフラットデザイン
}

function clampPercent(value: number): number {
  return Math.max(0, Math.min(100, value));
}

function toDiagramX(value: number): number {
  return 0.55 + ((W - 1.2) * clampPercent(value)) / 100;
}

function toDiagramY(value: number): number {
  return HEADER_H + 0.25 + ((H - HEADER_H - 0.6) * clampPercent(value)) / 100;
}

function toDiagramW(value: number): number {
  return ((W - 1.2) * Math.max(6, Math.min(100, value))) / 100;
}

function toDiagramH(value: number): number {
  return ((H - HEADER_H - 0.6) * Math.max(6, Math.min(100, value))) / 100;
}

function buildDiagramRowBands(s: PptxGenJS.Slide, blocks: PptxVisualBlock[], theme: Theme) {
  if (blocks.length < 3) return;
  const sorted = [...blocks].sort((a, b) => a.y - b.y);
  const groups: PptxVisualBlock[][] = [];
  for (const block of sorted) {
    const last = groups[groups.length - 1];
    if (!last) {
      groups.push([block]);
      continue;
    }
    const centerY = last.reduce((sum, item) => sum + item.y + item.h / 2, 0) / last.length;
    const blockCenter = block.y + block.h / 2;
    if (Math.abs(blockCenter - centerY) <= 12) {
      last.push(block);
    } else {
      groups.push([block]);
    }
  }

  groups
    .filter((group) => group.length >= 2)
    .slice(0, 4)
    .forEach((group, index) => {
      const minY = Math.min(...group.map((item) => item.y));
      const maxY = Math.max(...group.map((item) => item.y + item.h));
      s.addShape("roundRect", {
        x: 0.62,
        y: Math.max(HEADER_H + 0.26, toDiagramY(minY) - 0.1),
        w: 9.15,
        h: Math.min(H - HEADER_H - 0.7, toDiagramY(maxY) - toDiagramY(minY) + 0.22),
        rectRadius: 0.04,
        fill: { color: index % 2 === 0 ? theme.palette.sectionBg : theme.palette.tableAltBg, transparency: 55 },
        line: { color: theme.palette.border, transparency: 75, width: 0.7 },
      });
    });
}

function buildDiagramGroupCards(s: PptxGenJS.Slide, blocks: PptxVisualBlock[], theme: Theme) {
  const grouped = new Map<string, PptxVisualBlock[]>();
  for (const block of blocks) {
    if (!block.groupId) continue;
    const list = grouped.get(block.groupId) ?? [];
    list.push(block);
    grouped.set(block.groupId, list);
  }

  Array.from(grouped.values())
    .filter((group) => group.length >= 2)
    .slice(0, 5)
    .forEach((group, index) => {
      const minX = Math.min(...group.map((item) => item.x));
      const minY = Math.min(...group.map((item) => item.y));
      const maxX = Math.max(...group.map((item) => item.x + item.w));
      const maxY = Math.max(...group.map((item) => item.y + item.h));
      s.addShape("roundRect", {
        x: Math.max(0.55, toDiagramX(minX) - 0.12),
        y: Math.max(HEADER_H + 0.25, toDiagramY(minY) - 0.1),
        w: Math.min(9.0, toDiagramX(maxX) - toDiagramX(minX) + 0.24),
        h: Math.min(H - HEADER_H - 0.75, toDiagramY(maxY) - toDiagramY(minY) + 0.22),
        rectRadius: 0.04,
        fill: {
          color: index % 2 === 0 ? theme.palette.tableAltBg : theme.palette.sectionBg,
          transparency: 70,
        },
        line: { color: theme.palette.border, transparency: 82, width: 0.8 },
      });
    });
}

function drawDiagramConnector(
  s: PptxGenJS.Slide,
  from: PptxVisualBlock,
  to: PptxVisualBlock,
  connector: PptxConnector,
  theme: Theme
) {
  const fromX = toDiagramX(from.x) + toDiagramW(from.w) / 2;
  const fromY = toDiagramY(from.y) + toDiagramH(from.h) / 2;
  const toX = toDiagramX(to.x) + toDiagramW(to.w) / 2;
  const toY = toDiagramY(to.y) + toDiagramH(to.h) / 2;
  const midX = fromX + (toX - fromX) / 2;
  const midY = fromY + (toY - fromY) / 2;
  const relationship = connector.relationshipType ?? "flow";
  const lineColor =
    relationship === "compare"
      ? theme.palette.accentB
      : relationship === "annotation"
        ? theme.palette.mutedText
        : relationship === "support"
          ? theme.palette.border
          : theme.palette.accentA;
  const lineWidth =
    relationship === "annotation" ? 1 : relationship === "support" ? 0.9 : relationship === "compare" ? 1.6 : 1.4;
  const endArrow =
    relationship === "compare"
      ? "stealth"
      : connector.style === "line" || relationship === "support"
        ? "none"
        : "triangle";

  s.addShape("line", {
    x: fromX,
    y: fromY,
    w: toX - fromX,
    h: toY - fromY,
    line: {
      color: lineColor,
      width: lineWidth,
      transparency: relationship === "annotation" ? 18 : 0,
      beginArrowType: "none",
      endArrowType: endArrow,
    },
  });

  s.addShape("ellipse", {
    x: fromX - 0.04,
    y: fromY - 0.04,
    w: 0.08,
    h: 0.08,
    fill: { color: lineColor },
    line: { color: lineColor, width: 0 },
  });

  if (connector.label) {
    s.addText(connector.label, {
      x: midX - 0.72,
      y: midY - 0.16,
      w: 1.2,
      h: 0.25,
      fontSize: 9,
      fontFace: theme.fontFace,
      align: "center",
      color: relationship === "annotation" ? theme.palette.mutedText : theme.palette.bodyText,
      margin: 0,
      fill: { color: theme.palette.canvas, transparency: 8 },
    });
  }
}

function buildDiagramSlide(
  pptx: PptxGenJS,
  slide: PptxSlide,
  theme: Theme,
  visual: SlideVisualHint,
  faithfulMode?: boolean
) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  addVisualAccent(s, visual, theme, faithfulMode);

  const blocks = slide.visualBlocks ?? [];
  const connectors = slide.connectors ?? [];
  const primaryBlock =
    blocks.find((block) => block.role === "primary") ??
    blocks.find((block) => block.emphasis) ??
    [...blocks].sort((a, b) => b.w * b.h - a.w * a.h)[0] ??
    null;
  const validConnectors = connectors.filter(
    (conn) => blocks[conn.from] && blocks[conn.to] && conn.from !== conn.to
  );

  // faithfulModeではメインエリアを全幅に拡張、右サマリーパネルをスキップ
  const mainContainerW = faithfulMode ? W - 0.84 : 9.55;
  s.addShape("roundRect", {
    x: 0.42,
    y: HEADER_H + 0.18,
    w: mainContainerW,
    h: H - HEADER_H - 0.48,
    rectRadius: 0.05,
    fill: { color: theme.palette.surface, transparency: 2 },
    line: { color: theme.palette.border, width: 0.9 },
  });

  buildDiagramGroupCards(s, blocks, theme);
  buildDiagramRowBands(s, blocks, theme);

  if (!faithfulMode) {
    s.addShape("roundRect", {
      x: 10.12,
      y: HEADER_H + 0.18,
      w: 2.78,
      h: H - HEADER_H - 0.48,
      rectRadius: 0.05,
      fill: { color: theme.palette.surface },
      line: { color: theme.palette.border, width: 0.9 },
    });

    s.addText(localizeLabel("summary", theme), {
      x: 10.34,
      y: HEADER_H + 0.38,
      w: 2.2,
      h: 0.24,
      fontSize: theme.smallFontSize,
      fontFace: theme.fontFace,
      bold: true,
      color: theme.palette.mutedText,
    });

    if (primaryBlock) {
      s.addShape("roundRect", {
        x: 10.34,
        y: HEADER_H + 0.76,
        w: 2.18,
        h: 0.82,
        rectRadius: 0.05,
        fill: { color: theme.palette.accentA, transparency: theme.execMode ? 8 : 0 },
        line: { color: theme.palette.accentA, width: 0 },
      });
      s.addText(primaryBlock.text, {
        x: 10.52,
        y: HEADER_H + 0.93,
        w: 1.82,
        h: 0.48,
        fontSize: Math.max(theme.bodyFontSize - 1, 12),
        fontFace: theme.fontFace,
        bold: true,
        color: theme.palette.headerText,
        align: "center",
        valign: "middle",
        fit: "shrink",
        margin: 0.04,
      });
    }

    const summaryItems = slide.bullets.slice(0, 4);
    summaryItems.forEach((item, index) => {
      const y = HEADER_H + 1.8 + index * 0.88;
      s.addShape("roundRect", {
        x: 10.34,
        y,
        w: 2.18,
        h: 0.68,
        rectRadius: 0.04,
        fill: { color: "FFFFFF", transparency: 6 },
        line: { color: theme.palette.border, width: 0.8 },
      });
      s.addText(item, {
        x: 10.5,
        y: y + 0.1,
        w: 1.86,
        h: 0.46,
        fontSize: Math.max(theme.smallFontSize - 1, 9),
        fontFace: theme.fontFace,
        color: theme.palette.bodyText,
        valign: "middle",
        fit: "shrink",
        margin: 0.02,
      });
    });
  }

  for (const connector of validConnectors) {
    drawDiagramConnector(s, blocks[connector.from], blocks[connector.to], connector, theme);
  }

  blocks.forEach((block) => {
    const x = toDiagramX(block.x);
    const y = toDiagramY(block.y);
    const w = toDiagramW(block.w);
    const h = toDiagramH(block.h);
    const isBadge = block.kind === "badge";
    const isCallout = block.kind === "callout";
    const isFigure = block.kind === "figure";
    const isAnnotation = block.role === "annotation";
    const isSupporting = block.role === "supporting";
    const isPrimary =
      primaryBlock !== null &&
      block.text === primaryBlock.text &&
      block.x === primaryBlock.x &&
      block.y === primaryBlock.y;
    const fillColor = block.emphasis
      ? theme.palette.accentA
      : isBadge
        ? theme.palette.sectionBg
        : isFigure
          ? theme.palette.tableAltBg
          : isAnnotation
            ? theme.palette.sectionBg
          : theme.palette.surface;
    const textColor = block.emphasis ? theme.palette.headerText : theme.palette.bodyText;

    s.addShape(isCallout ? "roundRect" : isBadge ? "hexagon" : "roundRect", {
      x,
      y,
      w,
      h,
      rectRadius: isCallout ? 0.06 : 0.03,
      fill: { color: fillColor, transparency: isAnnotation ? 8 : 0 },
      line: { color: block.emphasis || isPrimary ? theme.palette.accentA : theme.palette.border, width: isPrimary ? 2 : 1.2 },
      shadow: { type: "outer", color: "AAB7C4", blur: isPrimary ? 2 : 1, angle: 45, opacity: isAnnotation ? 0.08 : isPrimary ? 0.18 : 0.12 },
    });

    if (isCallout) {
      s.addShape("chevron", {
        x: x + Math.max(w / 2 - 0.14, 0.06),
        y: y + h - 0.05,
        w: 0.28,
        h: 0.22,
        rotate: 90,
        fill: { color: fillColor },
        line: { color: block.emphasis || isPrimary ? theme.palette.accentA : theme.palette.border, width: 1 },
      });
    }

    if (isPrimary || isFigure) {
      s.addShape("rect", {
        x: x + 0.08,
        y: y + 0.08,
        w: w - 0.16,
        h: 0.1,
        fill: { color: theme.palette.accentB, transparency: isPrimary ? 0 : 15 },
        line: { color: theme.palette.accentB, width: 0 },
      });
    }

    s.addText(block.text, {
      x: x + 0.12,
      y: y + (isPrimary || isFigure ? 0.2 : 0.08),
      w: w - 0.24,
      h: h - (isPrimary || isFigure ? 0.28 : 0.16),
      fontSize: Math.max(theme.bodyFontSize - (isBadge ? 4 : isPrimary ? -1 : isSupporting ? 1 : 2), 10),
      fontFace: theme.fontFace,
      bold: block.emphasis || isBadge || isPrimary,
      color: textColor,
      align: "center",
      valign: "middle",
      margin: 0.05,
      breakLine: false,
      fit: "shrink",
    });
  });

  addChrome(s, theme);
}

function addAmbientShapes(s: PptxGenJS.Slide, theme: Theme) {
  if (theme.minimalMode) return;
  s.addShape("ellipse", {
    x: W - 1.9,
    y: -0.38,
    w: 2.3,
    h: 2.3,
    fill: { color: theme.palette.accentA, transparency: 78 },
    line: { color: theme.palette.accentA, transparency: 35, width: 1.5 },
  });
  s.addShape("ellipse", {
    x: 0.5,
    y: H - 1.55,
    w: 1.05,
    h: 1.05,
    fill: { color: theme.palette.accentB, transparency: 72 },
    line: { color: theme.palette.accentB, transparency: 45, width: 1 },
  });
  if (theme.playfulMode && !theme.execMode) {
    s.addShape("ellipse", {
      x: W - 3.2,
      y: H - 1.4,
      w: 0.72,
      h: 0.72,
      fill: { color: theme.palette.accentA, transparency: 58 },
      line: { color: theme.palette.accentA, transparency: 25, width: 1 },
    });
    s.addShape("ellipse", {
      x: W - 2.55,
      y: H - 1.12,
      w: 0.34,
      h: 0.34,
      fill: { color: theme.palette.accentB, transparency: 42 },
      line: { color: theme.palette.accentB, transparency: 20, width: 0.8 },
    });
  }
}

function addTitleDecorativeCircles(s: PptxGenJS.Slide, theme: Theme) {
  if (theme.minimalMode) return;
  // 背面の大きい円
  s.addShape("ellipse", {
    x: W - 3.5,
    y: -0.6,
    w: 3.8,
    h: 3.8,
    fill: { color: theme.palette.accentA, transparency: 30 },
    line: { color: theme.palette.accentA, transparency: 20, width: 0 },
  });
  // 前面の中くらいの円
  s.addShape("ellipse", {
    x: W - 2.6,
    y: 1.1,
    w: 2.8,
    h: 2.8,
    fill: { color: theme.palette.accentB, transparency: 38 },
    line: { color: theme.palette.accentB, transparency: 25, width: 0 },
  });
  // 小さい補助円
  s.addShape("ellipse", {
    x: W - 3.8,
    y: 1.6,
    w: 1.8,
    h: 1.8,
    fill: { color: theme.palette.headerBg, transparency: 45 },
    line: { color: theme.palette.accentA, transparency: 60, width: 1 },
  });
}

// 右上の装飾バー・ドットは不要なノイズになるため無効化済み
function addVisualAccent(
  _s: PptxGenJS.Slide,
  _visual: SlideVisualHint,
  _theme: Theme,
  _faithfulMode?: boolean,
  _hasComparisonData = true
) { return; }

function shouldGenerateIllustration(instructionText: string, theme: Theme): boolean {
  const lowered = instructionText.toLowerCase();
  if (containsAny(lowered, ["イラスト不要", "画像不要", "no illustration", "no image"])) return false;
  if (theme.execMode && !containsAny(lowered, ["イラスト", "挿絵", "絵", "robot", "ロボット", "image", "画像"])) {
    return false;
  }
  return containsAny(lowered, [
    "illustration",
    "illustrated",
    "image",
    "insert image",
    "robot",
    "eco mascot",
    "イラスト",
    "挿絵",
    "絵",
    "画像",
    "ロボット",
    "マスコット",
  ]);
}

async function generateCoverIllustration(
  title: string,
  instructionText: string,
  theme: Theme
): Promise<GeneratedIllustration | null> {
  if (!shouldGenerateIllustration(instructionText, theme)) return null;
  try {
    const openai = OpenAIDALLEInstance();
    const prompt = [
      "Create a clean presentation illustration for a business proposal cover.",
      `Topic: ${title}.`,
      `Design direction: ${instructionText || "professional and polished"}.`,
      "Style: modern editorial illustration, no text, no watermark, transparent or plain clean background, suitable for a PowerPoint cover.",
      theme.execMode
        ? "Tone: trustworthy, restrained, executive-ready, not cartoonish."
        : "Tone: approachable, polished, optimistic, presentation-friendly.",
    ].join(" ");
    const response = await openai.images.generate({
      model: "gpt-image-1.5",
      prompt,
    });
    const b64 = response.data?.[0]?.b64_json;
    if (!b64) return null;
    return {
      dataUri: `data:image/png;base64,${b64}`,
      prompt,
    };
  } catch (error) {
    console.warn("[gen-pptx] cover illustration skipped:", error);
    return null;
  }
}

function buildTitleSlide(
  pptx: PptxGenJS,
  title: string,
  brief: DeckDesignBrief,
  theme: Theme,
  slideCount: number,
  illustration?: GeneratedIllustration | null
) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.titleBg };
  // 右上装飾サークル（バンドではなく幾何学的モチーフ）
  addTitleDecorativeCircles(s, theme);
  // kicker ラベル（左上・小文字・アクセントカラー）— 区切り線なし
  const kicker = brief.coverKicker?.trim();
  if (kicker) {
    s.addText(kicker, {
      x: 0.72,
      y: 0.85,
      w: 6.0,
      h: 0.28,
      fontSize: theme.smallFontSize,
      fontFace: theme.fontFace,
      bold: true,
      color: theme.palette.accentB,
      charSpacing: 2,
    });
  }
  // メインタイトル（左揃え・大きく）
  const titleW = illustration?.dataUri ? W - 4.2 : W - 1.45;
  s.addText(title, {
    x: 0.72,
    y: kicker ? 1.22 : 1.0,
    w: titleW,
    h: 2.4,
    fontSize: theme.titleFontSize + 10,
    fontFace: theme.fontFace,
    bold: true,
    color: theme.palette.headerText,
    align: "left",
    valign: "middle",
    fit: "shrink",
  });
  // サブタイトル（左揃え・中くらい）— 空なら非表示
  const subtitle = brief.coverSubtitle?.trim();
  if (subtitle) {
    s.addText(subtitle, {
      x: 0.72,
      y: kicker ? 3.82 : 3.60,
      w: titleW,
      h: 0.55,
      fontSize: theme.bodyFontSize,
      fontFace: theme.fontFace,
      color: theme.palette.accentB,
      align: "left",
      fit: "shrink",
    });
  }
  if (illustration?.dataUri) {
    s.addShape("roundRect", {
      x: 9.55,
      y: 1.18,
      w: 2.75,
      h: 2.75,
      rectRadius: 0.08,
      fill: { color: theme.palette.surface, transparency: 4 },
      line: { color: theme.palette.accentB, transparency: 35, width: 1.2 },
      shadow: { type: "outer", color: "1B1B1B", blur: 2, angle: 45, opacity: 0.18 },
    });
    s.addImage({
      data: illustration.dataUri,
      x: 9.72,
      y: 1.33,
      w: 2.42,
      h: 2.42,
      sizing: { type: "contain", x: 9.72, y: 1.33, w: 2.42, h: 2.42 },
    });
  }
  s.addText(`${slideCount} ${localizeLabel("pages", theme)}`, {
    x: W - 2.2,
    y: 6.65,
    w: 1.7,
    h: 0.2,
    fontSize: theme.smallFontSize,
    fontFace: theme.fontFace,
    bold: true,
    align: "right",
    color: theme.palette.headerText,
  });
}

type Section = { header: string | null; items: string[] };

function parseSections(bullets: string[]): { hasSections: boolean; sections: Section[] } {
  const sections: Section[] = [];
  let current: Section = { header: null, items: [] };
  for (const bullet of bullets) {
    if (bullet.startsWith("[H]")) {
      if (current.header !== null || current.items.length > 0) sections.push(current);
      current = { header: bullet.slice(3).trim(), items: [] };
    } else {
      current.items.push(bullet);
    }
  }
  if (current.header !== null || current.items.length > 0) sections.push(current);
  return { hasSections: sections.some((section) => section.header !== null), sections };
}

function buildSectionSlide(pptx: PptxGenJS, title: string, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  s.addShape("rect", {
    x: 0,
    y: 2.82,
    w: W,
    h: 1.45,
    fill: { color: theme.palette.headerBg },
    line: { color: theme.palette.headerBg, width: 0 },
  });
  s.addShape("rect", {
    x: 0,
    y: 4.27,
    w: W,
    h: 0.06,
    fill: { color: theme.palette.accentA },
    line: { color: theme.palette.accentA, width: 0 },
  });
  s.addText(title, {
    x: 0.6,
    y: 2.93,
    w: W - 1.2,
    h: 1.18,
    fontSize: theme.titleFontSize + 4,
    fontFace: theme.fontFace,
    bold: true,
    color: theme.palette.headerText,
    align: "center",
    valign: "middle",
  });
  addChrome(s, theme);
}

// body が助詞で始まる場合は「は完了しており...」のような不自然な文になるためガード
const LEADING_PARTICLE_RE = /^[はがをにでとへからまでよりも]/;

function splitBulletTitle(text: string): { title: string; body: string } {
  const full = text.trim();
  const colonMatch = full.match(/^([^：:]{2,22})[：:]\s*([\s\S]*)$/);
  if (colonMatch) {
    const bd = colonMatch[2];
    return { title: colonMatch[1], body: LEADING_PARTICLE_RE.test(bd) ? full : bd };
  }
  const periodMatch = full.match(/^([^。！!]{5,20}[。！!])\s*([\s\S]*)$/);
  if (periodMatch && periodMatch[1].length <= 22) {
    const bd = periodMatch[2];
    return { title: periodMatch[1], body: LEADING_PARTICLE_RE.test(bd) ? full : bd };
  }
  const spaceIdx = full.search(/[\s　]/);
  if (spaceIdx > 3 && spaceIdx < 16) {
    const bd = full.slice(spaceIdx).trim();
    return { title: full.slice(0, spaceIdx), body: LEADING_PARTICLE_RE.test(bd) ? full : bd };
  }
  // は/が: title = 助詞前, body = 全文（助詞始まりの body を生成しない）
  for (let i = 2; i <= Math.min(22, full.length - 2); i++) {
    if (full[i] === "は" || full[i] === "が") {
      return { title: full.slice(0, i).trim(), body: full };
    }
  }
  return { title: full.slice(0, 14), body: full.slice(14).trim() };
}

function renderHorizontalCards(
  s: PptxGenJS.Slide,
  bullets: string[],
  theme: Theme,
  startY: number,
  totalW: number
) {
  const count = bullets.length;
  const GAP = 0.14;
  const cardW = (totalW - GAP * (count - 1)) / count;
  const cardH = H - startY - 0.26;
  const x0 = 0.42;
  const ICON_D = 0.68;
  const ICON_TOP = 0.22;

  const spaceMult = getSpacingMult(theme);
  const adjustedCardH = cardH * (spaceMult > 1 ? 1 : 1); // 高さは変えず余白内で調整

  bullets.forEach((bullet, idx) => {
    const cx = x0 + idx * (cardW + GAP);
    const numStr = String(idx + 1).padStart(2, "0");
    const { title, body } = splitBulletTitle(bullet);
    const accentFill = idx === 0 ? theme.palette.accentA : theme.palette.headerBg;

    // カード背景（cardStyle 対応）
    const hcp = getCardStyleProps(theme, accentFill);
    s.addShape("roundRect", {
      x: cx, y: startY, w: cardW, h: adjustedCardH,
      rectRadius: hcp.rectRadius,
      fill: hcp.fill as any,
      line: hcp.line as any,
      ...(hcp.shadow ? { shadow: hcp.shadow } : {}),
    });

    // ウォーターマーク数字
    s.addText(numStr, {
      x: cx + cardW - 1.0, y: startY + 0.04,
      w: 0.9, h: 0.82,
      fontSize: theme.bodyFontSize + 6,
      fontFace: theme.fontFace, bold: true,
      color: theme.cardStyle === "filled" ? theme.palette.headerText : theme.palette.accentA,
      align: "right", transparency: 70,
    });

    // アイコン円（上部中央）
    const iconPadTop = ICON_TOP * spaceMult;
    const iconX = cx + cardW / 2 - ICON_D / 2;
    const iconY = startY + iconPadTop;
    s.addShape("ellipse", {
      x: iconX, y: iconY, w: ICON_D, h: ICON_D,
      fill: { color: accentFill },
      line: { color: accentFill, width: 0 },
    });
    s.addText(numStr, {
      x: iconX + 0.04, y: iconY + 0.04,
      w: ICON_D - 0.08, h: ICON_D - 0.08,
      fontSize: theme.bodyFontSize + 1,
      fontFace: theme.fontFace, bold: true,
      color: "FFFFFF",
      align: "center", valign: "middle",
    });

    // タイトル
    const titleY = iconY + ICON_D + 0.14;
    const titleH = 0.44;
    s.addText(title, {
      x: cx + 0.1, y: titleY,
      w: cardW - 0.2, h: titleH,
      fontSize: theme.bodyFontSize + 1,
      fontFace: theme.fontFace, bold: true,
      color: getCardBodyColor(theme),
      align: "center", valign: "middle", fit: "shrink",
    });

    // 本文テキスト
    const bodyY = titleY + titleH + 0.1;
    s.addText(body, {
      x: cx + 0.14, y: bodyY,
      w: cardW - 0.28,
      h: adjustedCardH - (bodyY - startY) - 0.16,
      fontSize: Math.max(theme.bodyFontSize - 2, 10),
      fontFace: theme.fontFace,
      color: getCardMutedColor(theme),
      valign: "top", fit: "shrink",
    });

    // 矢印コネクタ（カード間）
    if (idx < count - 1) {
      const arrowW = GAP * 0.65;
      const arrowH = cardH * 0.13;
      s.addShape("chevron", {
        x: cx + cardW + GAP * 0.12,
        y: startY + cardH / 2 - arrowH / 2,
        w: arrowW, h: arrowH,
        rotate: 0,
        fill: { color: theme.palette.accentA },
        line: { color: theme.palette.accentA, width: 0 },
      });
    }
  });
}

function renderCardBullets(
  s: PptxGenJS.Slide,
  bullets: string[],
  theme: Theme,
  startY: number,
  cardW: number,
  fontScaleMult: number = 1.0
) {
  if (bullets.length === 0) return;

  // 6件以上: カード化せず dense bullets として全件表示（情報欠落防止）
  if (bullets.length > 5) {
    const denseFontSize = Math.max(9, Math.round((theme.bodyFontSize - 1) * fontScaleMult));
    const denseItems = bullets.map((item) => ({
      text: item,
      options: {
        bullet: { indent: 14 },
        breakLine: true,
        fontSize: denseFontSize,
        fontFace: theme.fontFace,
        color: theme.palette.bodyText,
        paraSpaceAfter: 5,
      },
    }));
    s.addText(denseItems, {
      x: 0.42,
      y: startY,
      w: cardW,
      h: H - startY - 0.26,
      margin: 0.1,
      valign: "top",
    });
    return;
  }

  const count = bullets.length;

  // 2〜3アイテム → 横並びカード（プロセスフロー）
  if (count <= 3) {
    renderHorizontalCards(s, bullets.slice(0, count), theme, startY, cardW);
    return;
  }

  // 4〜5アイテム → 縦並びカード
  const STRIP_W = 0.38;
  const GAP = 0.1;
  const totalH = H - startY - 0.26;
  const cardH = Math.max(0.88, (totalH - GAP * (count - 1)) / count);

  const CARD_STRIP_CYCLE = [
    theme.palette.accentA,
    theme.palette.accentB,
    theme.palette.headerBg,
    theme.palette.accentA,
    theme.palette.accentB,
  ];

  bullets.slice(0, count).forEach((bullet, idx) => {
    const y = startY + idx * (cardH + GAP);
    const numStr = String(idx + 1).padStart(2, "0");
    const stripFill = CARD_STRIP_CYCLE[idx % CARD_STRIP_CYCLE.length];

    const bcp = getCardStyleProps(theme, stripFill);
    s.addShape("roundRect", {
      x: 0.42, y, w: cardW, h: cardH,
      rectRadius: bcp.rectRadius,
      fill: bcp.fill as any,
      line: bcp.line as any,
      ...(bcp.shadow ? { shadow: bcp.shadow } : {}),
    });
    // flat/glass ではサイドストリップの代わりに細いアクセントラインのみ
    if (theme.cardStyle === "flat" || theme.cardStyle === "glass") {
      s.addShape("rect", {
        x: 0.42, y, w: 0.06, h: cardH,
        fill: { color: stripFill },
        line: { color: stripFill, width: 0 },
      });
    } else {
      s.addShape("roundRect", {
        x: 0.42, y, w: STRIP_W, h: cardH,
        rectRadius: 0.06,
        fill: { color: stripFill },
        line: { color: stripFill, width: 0 },
      });
    }
    s.addText(numStr, {
      x: 0.44, y: y + cardH / 2 - 0.22,
      w: STRIP_W - 0.06, h: 0.44,
      fontSize: Math.max(8, Math.round((theme.bodyFontSize - 1) * fontScaleMult)),
      fontFace: theme.fontFace, bold: true,
      color: "FFFFFF",
      align: "center", valign: "middle",
    });
    s.addText(bullet, {
      x: 0.42 + STRIP_W + 0.16, y: y + 0.1,
      w: cardW - STRIP_W - 0.24, h: cardH - 0.2,
      fontSize: Math.max(9, Math.round(theme.bodyFontSize * fontScaleMult)),
      fontFace: theme.fontFace,
      color: theme.palette.bodyText,
      valign: "middle", fit: "shrink",
    });
  });
}

function buildBulletsSlide(
  pptx: PptxGenJS,
  slide: PptxSlide,
  theme: Theme,
  visual: SlideVisualHint,
  illustration?: GeneratedIllustration | null,
  faithfulMode?: boolean
) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  const bFontAdj = densityFontAdj(slide.density);
  const fontScaleMult = slide.fontScale ?? 1.0;
  const bMaxItems = densityMaxItems(slide.density, 5);
  const adjustedBullets = slide.bullets.slice(0, bMaxItems);
  const { hasSections, sections } = parseSections(adjustedBullets);
  const useHorizontalCards = !faithfulMode && !theme.execMode && !hasSections &&
    adjustedBullets.length >= 2 && adjustedBullets.length <= 3;
  // 横並びカードのとき右側アクセントパネルは非表示（カードが全幅を使う）
  // comparison アイコンは列データがある場合のみ（bullets フォールバック時は非表示）
  if (!useHorizontalCards) {
    const hasColData = visual.visualType !== "comparison" || Boolean(slide.columns && slide.columns.length >= 2);
    addVisualAccent(s, visual, theme, faithfulMode, hasColData);
  }
  const showIllustration = Boolean(illustration?.dataUri) && !theme.execMode && !faithfulMode && !useHorizontalCards;
  const textWidth = showIllustration ? 6.85 : 8.15;
  // faithfulMode では execMode による内容圧縮（キーメッセージ1件＋最大3件）をスキップ
  if (theme.execMode && !faithfulMode) {
    // 6件以上は圧縮せず dense bullets fallback で全件表示
    if (slide.bullets.length > 5) {
      const denseFontSize = Math.max(9, theme.bodyFontSize - 1);
      const denseItems = slide.bullets.map((item) => ({
        text: item,
        options: {
          bullet: { indent: 14 },
          breakLine: true,
          fontSize: denseFontSize,
          fontFace: theme.fontFace,
          color: theme.palette.bodyText,
          paraSpaceAfter: 5,
        },
      }));
      s.addText(denseItems, {
        x: 0.55,
        y: HEADER_H + 0.24,
        w: 8.15,
        h: H - HEADER_H - 0.55,
        margin: 0.1,
        valign: "top",
      });
      addChrome(s, theme);
      return;
    }
    const allBullets = sections.flatMap((section) => section.items);
    const keyMessage = allBullets[0] ?? (Array.isArray(slide.bullets) && slide.bullets[0]) ?? "";
    const rest = allBullets.slice(1, 4);

    if (keyMessage) {
      s.addShape("roundRect", {
        x: 0.55,
        y: 1.45,
        w: 8.15,
        h: 1.0,
        rectRadius: 0.04,
        fill: { color: theme.palette.surface },
        line: { color: theme.palette.accentA, width: 1.5 },
      });
      s.addText(keyMessage, {
        x: 0.75,
        y: 1.63,
        w: 7.75,
        h: 0.65,
        fontSize: theme.bodyFontSize + 2,
        fontFace: theme.fontFace,
        bold: true,
        color: theme.palette.bodyText,
        valign: "middle",
      });
    }

    if (rest.length > 0) {
      const bulletItems = rest.map((item) => ({
        text: item,
        options: {
          bullet: { indent: 14 },
          breakLine: true,
          fontSize: theme.bodyFontSize,
          fontFace: theme.fontFace,
          color: theme.palette.bodyText,
          paraSpaceAfter: 10,
        },
      }));
      s.addText(bulletItems, {
        x: 0.55,
        y: 2.72,
        w: 8.15,
        h: H - 3.05,
        margin: 0.1,
        valign: "top",
      });
    }
      addChrome(s, theme);
      return;
  }

  if (showIllustration) {
    s.addShape("roundRect", {
      x: 9.05,
      y: 1.72,
      w: 3.25,
      h: 3.25,
      rectRadius: 0.08,
      fill: { color: theme.palette.surface },
      line: { color: theme.palette.border, width: 1 },
    });
    s.addImage({
      data: illustration!.dataUri,
      x: 9.22,
      y: 1.9,
      w: 2.92,
      h: 2.92,
      sizing: { type: "contain", x: 9.22, y: 1.9, w: 2.92, h: 2.92 },
    });
  }

  if (hasSections) {
    let currentY = HEADER_H + 0.24;
    const maxY = H - 0.28;
    for (const section of sections) {
      if (section.header) {
        if (currentY + 0.42 > maxY) break;
        s.addShape("rect", {
          x: 0.42,
          y: currentY,
          w: 8.35,
          h: 0.42,
          fill: { color: theme.palette.surface },
          line: { color: theme.palette.border, width: 0.8 },
        });
        s.addText(section.header, {
          x: 0.56,
          y: currentY + 0.04,
          w: textWidth - 0.15,
          h: 0.28,
          fontSize: theme.bodyFontSize - 1,
          fontFace: theme.fontFace,
          bold: true,
          color: theme.palette.bodyText,
          valign: "middle",
        });
        currentY += 0.5;
      }
      for (const item of section.items) {
        if (currentY + 0.31 > maxY) break;
        s.addText(`• ${item}`, {
          x: 0.72,
          y: currentY,
          w: textWidth - 0.3,
          h: 0.28,
          fontSize: theme.bodyFontSize - 1,
          fontFace: theme.fontFace,
          color: theme.palette.bodyText,
          valign: "middle",
        });
        currentY += 0.33;
      }
      currentY += 0.08;
    }
  } else if (slide.bullets.length > 0) {
    if (faithfulMode) {
      const faithfulFontSize = Math.max(9, Math.round(theme.bodyFontSize * fontScaleMult));
      const bulletItems = slide.bullets.map((item) => ({
        text: item,
        options: {
          bullet: { indent: 14 },
          breakLine: true,
          fontSize: faithfulFontSize,
          fontFace: theme.fontFace,
          color: theme.palette.bodyText,
          paraSpaceAfter: 9,
        },
      }));
      s.addText(bulletItems, {
        x: 0.55,
        y: HEADER_H + 0.24,
        w: textWidth,
        h: H - HEADER_H - 0.55,
        margin: 0.1,
        valign: "top",
        ...(slide.fitTextToShape ? { shrinkText: true } : {}),
      });
    } else {
      const cardW = showIllustration ? 6.55 : 8.88;
      renderCardBullets(s, slide.bullets, theme, HEADER_H + 0.18, cardW, fontScaleMult);
    }
  }

  addChrome(s, theme);
}

function chunkConversationTurns(turns: PptxConversationTurn[], chunkSize: number): PptxConversationTurn[][] {
  const chunks: PptxConversationTurn[][] = [];
  for (let i = 0; i < turns.length; i += chunkSize) {
    chunks.push(turns.slice(i, i + chunkSize));
  }
  return chunks;
}

function buildConversationSlide(
  pptx: PptxGenJS,
  slide: PptxSlide,
  theme: Theme
) {
  const turns = [...(slide.conversationTurns ?? [])].sort((a, b) => a.turnIndex - b.turnIndex);
  const style = slide.conversationStyle ?? "chat-ui";
  const chunks = chunkConversationTurns(turns, style === "chat-ui" ? 8 : 10);
  const pages = chunks.length > 0 ? chunks : [[]];

  pages.forEach((pageTurns, pageIndex) => {
    const s = pptx.addSlide();
    s.background = { color: theme.palette.canvas };
    addHeaderBand(s, pages.length > 1 ? `${slide.title} (${pageIndex + 1}/${pages.length})` : slide.title, theme);

    if (pageTurns.length === 0) {
      addChrome(s, theme);
      return;
    }

    const contentTop = HEADER_H + 0.18;
    const contentBottom = H - 0.24;
    const availH = contentBottom - contentTop;
    const turnH = availH / pageTurns.length;

    pageTurns.forEach((turn, idx) => {
      const y = contentTop + idx * turnH;
      const label = turn.speakerRole.trim() || `Speaker ${idx + 1}`;
      const type = turn.speakerType ?? "other";

      if (style === "chat-ui") {
        const alignLeft = type === "agent" || type === "staff";
        const bubbleX = alignLeft ? 0.52 : 4.18;
        const bubbleW = 8.63;
        const labelX = alignLeft ? bubbleX : bubbleX + bubbleW - 2.0;
        const labelAlign = alignLeft ? "left" : "right";
        const bubbleFill = alignLeft ? theme.palette.accentA : theme.palette.surface;
        const bubbleLine = alignLeft ? theme.palette.accentA : theme.palette.border;
        const textColor = alignLeft ? theme.palette.headerText : theme.palette.bodyText;
        const labelColor = alignLeft ? theme.palette.accentA : theme.palette.mutedText;
        const labelH = Math.min(0.22, turnH * 0.26);
        const bubbleH = Math.max(0.42, turnH - labelH - 0.12);

        s.addText(label, {
          x: labelX,
          y,
          w: 2.0,
          h: labelH,
          fontSize: Math.max(theme.smallFontSize - 1, 9),
          fontFace: theme.fontFace,
          bold: true,
          color: labelColor,
          align: labelAlign,
          valign: "middle",
          fit: "shrink",
        });
        s.addShape("roundRect", {
          x: bubbleX,
          y: y + labelH,
          w: bubbleW,
          h: bubbleH,
          rectRadius: 0.06,
          fill: { color: bubbleFill },
          line: { color: bubbleLine, width: alignLeft ? 0 : 1 },
          shadow: alignLeft ? undefined : { type: "outer", color: "C8D4E0", blur: 1, angle: 45, opacity: 0.15 },
        });
        s.addText(turn.text, {
          x: bubbleX + 0.14,
          y: y + labelH + 0.06,
          w: bubbleW - 0.28,
          h: bubbleH - 0.12,
          fontSize: Math.max(theme.bodyFontSize - 2, 11),
          fontFace: theme.fontFace,
          color: textColor,
          valign: "middle",
          fit: "shrink",
          margin: 0.04,
        });
        return;
      }

      if (style === "interview") {
        const roleW = 2.2;
        const bubbleX = 2.85;
        const bubbleW = W - bubbleX - 0.55;
        const boxH = Math.max(0.5, turnH - 0.08);
        const roleFill =
          type === "agent" || type === "staff" ? theme.palette.sectionBg :
          type === "customer" ? theme.palette.tableAltBg :
          "F7F9FC";

        s.addShape("roundRect", {
          x: 0.48,
          y,
          w: roleW,
          h: boxH,
          rectRadius: 0.05,
          fill: { color: roleFill },
          line: { color: theme.palette.border, width: 0.8 },
        });
        s.addText(label, {
          x: 0.62,
          y: y + 0.08,
          w: roleW - 0.28,
          h: boxH - 0.16,
          fontSize: Math.max(theme.bodyFontSize - 2, 11),
          fontFace: theme.fontFace,
          bold: true,
          color: theme.palette.bodyText,
          align: "center",
          valign: "middle",
          fit: "shrink",
        });
        s.addShape("roundRect", {
          x: bubbleX,
          y,
          w: bubbleW,
          h: boxH,
          rectRadius: 0.04,
          fill: { color: theme.palette.surface },
          line: { color: theme.palette.border, width: 0.8 },
        });
        s.addText(turn.text, {
          x: bubbleX + 0.16,
          y: y + 0.08,
          w: bubbleW - 0.32,
          h: boxH - 0.16,
          fontSize: Math.max(theme.bodyFontSize - 2, 11),
          fontFace: theme.fontFace,
          color: theme.palette.bodyText,
          valign: "middle",
          fit: "shrink",
          margin: 0.03,
        });
        return;
      }

      const cardH = Math.max(0.56, turnH - 0.08);
      s.addShape("roundRect", {
        x: 0.48,
        y,
        w: W - 0.96,
        h: cardH,
        rectRadius: 0.04,
        fill: { color: idx % 2 === 0 ? theme.palette.surface : theme.palette.tableAltBg },
        line: { color: theme.palette.border, width: 0.8 },
      });
      s.addText(label, {
        x: 0.7,
        y: y + 0.08,
        w: 2.0,
        h: 0.2,
        fontSize: Math.max(theme.smallFontSize - 1, 9),
        fontFace: theme.fontFace,
        bold: true,
        color: theme.palette.mutedText,
        fit: "shrink",
      });
      s.addText(turn.text, {
        x: 0.7,
        y: y + 0.28,
        w: W - 1.4,
        h: cardH - 0.34,
        fontSize: Math.max(theme.bodyFontSize - 2, 11),
        fontFace: theme.fontFace,
        color: theme.palette.bodyText,
        valign: "middle",
        fit: "shrink",
        margin: 0.02,
      });
    });

    addChrome(s, theme);
  });
}

function buildTableSlide(
  pptx: PptxGenJS,
  slide: PptxSlide,
  theme: Theme,
  visual: SlideVisualHint,
  faithfulMode?: boolean
) {
  // tableRows が不十分なら bullets にフォールバック
  if (!slide.tableRows || slide.tableRows.length < 2) {
    buildBulletsSlide(pptx, slide, theme, visual, null, faithfulMode);
    return;
  }

  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  addVisualAccent(s, visual, theme, faithfulMode);

  const rows = slide.tableRows;

  const colCount = Math.max(...rows.map((row) => row.length), 1);
  const colW = (W - 1.05) / colCount;
  const tableData = rows.map((row, rowIndex) =>
    Array.from({ length: colCount }, (_, colIndex) => ({
      text: row[colIndex] ?? "",
      options: {
        bold: rowIndex === 0,
        fontSize: rowIndex === 0 ? theme.bodyFontSize - 1 : theme.bodyFontSize - 2,
        fontFace: theme.fontFace,
        color: rowIndex === 0 ? theme.palette.tableHeaderText : theme.palette.bodyText,
        fill: { color: rowIndex === 0 ? theme.palette.tableHeaderBg : rowIndex % 2 === 0 ? theme.palette.tableAltBg : "FFFFFF" },
        align: "center" as const,
        valign: "middle" as const,
        margin: 5,
        border: [
          { type: "solid" as const, pt: 0.5, color: theme.palette.border },
          { type: "solid" as const, pt: 0.5, color: theme.palette.border },
          { type: "solid" as const, pt: 0.5, color: theme.palette.border },
          { type: "solid" as const, pt: 0.5, color: theme.palette.border },
        ] as [any, any, any, any],
      },
    }))
  );

  s.addTable(tableData, {
    x: 0.52,
    y: HEADER_H + 0.24,
    w: 8.25,
    colW: Array(colCount).fill(colW),
  });
  addChrome(s, theme);
}

function buildMultiColumnSlide(
  pptx: PptxGenJS,
  slide: PptxSlide,
  theme: Theme,
  visual: SlideVisualHint,
  faithfulMode?: boolean
) {
  const columns = (slide.columns ?? []).filter((column) =>
    Boolean(column.header?.trim()) || (column.bullets ?? []).some((b) => Boolean(b?.trim()))
  );

  if (columns.length < 2) {
    buildBulletsSlide(pptx, slide, theme, visual, null, faithfulMode);
    return;
  }

  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  // multi-columnはコンテンツが全幅を使うためvisualAccentを省略

  const MARGIN = 0.42;
  const totalW = W - MARGIN * 2;
  const COL_GAP = 0.16;
  const colW = (totalW - COL_GAP * (columns.length - 1)) / columns.length;
  const colX0 = MARGIN;
  const colY = HEADER_H + 0.22;
  const headerH = faithfulMode ? 0.48 : 0.78;
  const contentH = H - colY - headerH - 0.28;

  columns.forEach((column, index) => {
    const x = colX0 + index * colW;
    const colInnerW = colW - 0.1;

    if (faithfulMode) {
      s.addShape("rect", {
        x,
        y: colY,
        w: colInnerW,
        h: headerH,
        fill: { color: theme.palette.surface },
        line: { color: theme.palette.border, width: 0.8 },
      });
      s.addText(column.header, {
        x: x + 0.06,
        y: colY + 0.06,
        w: colInnerW - 0.12,
        h: headerH - 0.12,
        fontSize: theme.bodyFontSize - 1,
        fontFace: theme.fontFace,
        bold: true,
        color: theme.palette.bodyText,
        align: "center",
        valign: "middle",
        fit: "shrink",
      });
    } else {
      // 濃色ヘッダー + 番号バッジ
      s.addShape("roundRect", {
        x,
        y: colY,
        w: colInnerW,
        h: headerH,
        rectRadius: 0.06,
        fill: { color: theme.palette.headerBg },
        line: { color: theme.palette.headerBg, width: 0 },
      });
      s.addText(String(index + 1).padStart(2, "0"), {
        x: x + 0.1,
        y: colY + 0.06,
        w: 0.44,
        h: 0.28,
        fontSize: theme.smallFontSize,
        fontFace: theme.fontFace,
        bold: true,
        color: theme.palette.accentB,
      });
      s.addText(column.header, {
        x: x + 0.06,
        y: colY + 0.36,
        w: colInnerW - 0.12,
        h: 0.36,
        fontSize: theme.bodyFontSize,
        fontFace: theme.fontFace,
        bold: true,
        color: theme.palette.headerText,
        align: "center",
        valign: "middle",
        fit: "shrink",
      });
    }

    // 列コンテンツ（cardStyle 対応）
    const bulletCount = Math.min(column.bullets.length, 5);
    if (bulletCount > 0 && !faithfulMode) {
      const itemGap = 0.07 * getSpacingMult(theme);
      const itemH = Math.max(0.52, (contentH - itemGap * (bulletCount - 1)) / bulletCount);
      column.bullets.slice(0, bulletCount).forEach((item, bulletIdx) => {
        const itemY = colY + headerH + 0.1 + bulletIdx * (itemH + itemGap);
        const mcp = getCardStyleProps(theme, theme.palette.accentA);
        s.addShape("roundRect", {
          x: x + 0.06, y: itemY,
          w: colInnerW - 0.12, h: itemH,
          rectRadius: mcp.rectRadius,
          fill: theme.cardStyle === "filled"
            ? { color: bulletIdx % 2 === 0 ? theme.palette.accentA : theme.palette.accentB }
            : mcp.fill as any,
          line: mcp.line as any,
        });
        // flat/glass ではサイドアクセントのみ
        if (theme.cardStyle !== "filled") {
          s.addShape("roundRect", {
            x: x + 0.06, y: itemY, w: 0.18, h: itemH,
            rectRadius: 0.04,
            fill: { color: theme.palette.accentA, transparency: 28 },
            line: { color: theme.palette.accentA, width: 0 },
          });
        }
        s.addText(item, {
          x: x + (theme.cardStyle === "filled" ? 0.14 : 0.28),
          y: itemY + 0.04,
          w: colInnerW - (theme.cardStyle === "filled" ? 0.28 : 0.46),
          h: itemH - 0.08,
          fontSize: Math.max(theme.bodyFontSize - 3, 10),
          fontFace: theme.fontFace,
          color: getCardBodyColor(theme),
          valign: "middle",
          fit: "shrink",
        });
      });
    } else if (column.bullets.length > 0) {
      const bulletItems = column.bullets.map((item) => ({
        text: item,
        options: {
          bullet: { indent: 10 },
          breakLine: true,
          fontSize: theme.bodyFontSize - 3,
          fontFace: theme.fontFace,
          color: theme.palette.bodyText,
          paraSpaceAfter: 6,
        },
      }));
      s.addText(bulletItems, {
        x: x + 0.06,
        y: colY + headerH + 0.08,
        w: colInnerW - 0.12,
        h: contentH - 0.1,
        margin: 0.05,
        valign: "top",
      });
    }
  });

  addChrome(s, theme);
}

// ─── company-overview レイアウト ───────────────────────────────────────────

function buildCompanyOverviewSlide(pptx: PptxGenJS, slide: PptxSlide, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);

  const { density, textTreatment, visualIntent } = slide;
  const renderMode = parseVisualIntent(visualIntent);
  const fontAdj   = densityFontAdj(density);
  const valFontAdj = intentValueFontAdj(renderMode);
  const noteExpand = intentNoteExpand(renderMode);
  const leadBoost  = leadTextHeightBoost(textTreatment);

  const contentY = HEADER_H + 0.22;
  const contentH = H - contentY - 0.22;
  const LEFT_W = 5.1;
  const RIGHT_X = 5.72;
  const RIGHT_W = W - RIGHT_X - 0.42;

  // ── 左パネル ──
  const leadText = slide.leadText ?? slide.bullets.join("　");
  const hasCallout = Boolean(slide.callout);
  // textTreatment=explanatory → leadText領域を広げ、calloutY を下げる
  const leadH = hasCallout ? Math.max(2.6 + leadBoost, 2.0) : contentH;

  // calloutなしのとき: 背景カードでテキストエリアを視覚的に区切る
  if (!hasCallout) {
    s.addShape("roundRect", {
      x: 0.48, y: contentY,
      w: LEFT_W, h: contentH,
      rectRadius: 0.06,
      fill: { color: theme.palette.surface },
      line: { color: theme.palette.border, width: 0.8 },
    });
    s.addShape("roundRect", {
      x: 0.48, y: contentY,
      w: 0.22, h: contentH,
      rectRadius: 0.06,
      fill: { color: theme.palette.accentA },
      line: { color: theme.palette.accentA, width: 0 },
    });
  }

  s.addText(leadText, {
    x: hasCallout ? 0.48 : 0.82,
    y: contentY + (hasCallout ? 0 : 0.24),
    w: hasCallout ? LEFT_W : LEFT_W - 0.48,
    h: hasCallout ? leadH : contentH - 0.48,
    fontSize: Math.max(theme.bodyFontSize - 1 + fontAdj, 9),
    fontFace: theme.fontFace,
    color: theme.palette.bodyText,
    valign: hasCallout ? "top" : "middle",
    fit: "shrink",
  });

  if (hasCallout && slide.callout) {
    const calloutY = contentY + 2.75;
    const calloutH = contentH - 2.75 - 0.1;
    s.addShape("roundRect", {
      x: 0.48, y: calloutY,
      w: LEFT_W, h: Math.max(calloutH, 1.8),
      rectRadius: 0.06,
      fill: { color: theme.palette.surface },
      line: { color: theme.palette.accentA, width: 1.2 },
    });
    s.addShape("rect", {
      x: 0.48, y: calloutY,
      w: 0.22, h: Math.max(calloutH, 1.8),
      fill: { color: theme.palette.accentA },
      line: { color: theme.palette.accentA, width: 0 },
    });
    s.addText(slide.callout.title, {
      x: 0.82, y: calloutY + 0.14,
      w: LEFT_W - 0.48, h: 0.32,
      fontSize: theme.bodyFontSize - 1,
      fontFace: theme.fontFace,
      bold: true,
      color: theme.palette.accentA,
    });
    s.addText(slide.callout.body, {
      x: 0.82, y: calloutY + 0.5,
      w: LEFT_W - 0.48, h: Math.max(calloutH, 1.8) - 0.68,
      fontSize: theme.bodyFontSize - 2,
      fontFace: theme.fontFace,
      color: theme.palette.bodyText,
      valign: "top",
      fit: "shrink",
    });
  }

  // ── 縦区切り線 ──
  s.addShape("rect", {
    x: 5.62, y: contentY + 0.1,
    w: 0.04, h: contentH - 0.2,
    fill: { color: theme.palette.border },
    line: { color: theme.palette.border, width: 0 },
  });

  // ── 右パネル: メトリクスカード（density で枚数・高さを調整） ──
  const allMetrics = slide.metrics ?? [];
  const maxM = densityMaxItems(density, 4);
  const metrics = allMetrics.slice(0, maxM);
  const COLS = density === "low" && metrics.length <= 2 ? 1 : 2;
  const ROWS = Math.ceil(metrics.length / COLS);
  const CARD_GAP = 0.14;
  const cardW = (RIGHT_W - CARD_GAP * (COLS - 1)) / COLS;
  const cardH = (contentH - CARD_GAP * (ROWS - 1)) / ROWS;

  const OVERVIEW_STRIP_CYCLE = [
    theme.palette.accentA, theme.palette.accentB,
    theme.palette.accentB, theme.palette.accentA,
  ];
  const ICON_D   = 0.42;
  const ICON_PAD = 0.18;
  const NOTE_H   = Math.min(0.42 * noteExpand, 0.58);
  const INNER_W  = cardW - ICON_PAD * 2;

  metrics.forEach((metric, idx) => {
    const col = idx % COLS;
    const row = Math.floor(idx / COLS);
    const cx = RIGHT_X + col * (cardW + CARD_GAP);
    const cy = contentY + row * (cardH + CARD_GAP);

    // ── LLMの colorRole を尊重してアイコン色を決定 ──
    const rawIconColor =
      metric.colorRole === "accent"  ? theme.palette.accentB :
      metric.colorRole === "neutral" ? theme.palette.headerBg :
      OVERVIEW_STRIP_CYCLE[idx % OVERVIEW_STRIP_CYCLE.length];
    const iconColor = rawIconColor;

    // ── カード背景（cardStyle 対応）──
    const ocp = getCardStyleProps(theme, iconColor);
    s.addShape("roundRect", {
      x: cx, y: cy, w: cardW, h: cardH,
      rectRadius: ocp.rectRadius,
      fill: ocp.fill as any,
      line: ocp.line as any,
      ...(ocp.shadow ? { shadow: ocp.shadow } : {}),
      shadow: { type: "outer", color: "C0CCDA", blur: 1, angle: 45, opacity: 0.1 },
    });

    // ── アイコン円 ──
    const iconX = cx + ICON_PAD;
    const iconY = cy + ICON_PAD;
    s.addShape("ellipse", {
      x: iconX, y: iconY, w: ICON_D, h: ICON_D,
      fill: { color: iconColor },
      line: { color: iconColor, width: 0 },
    });
    const iconKey = resolveMetricIconKey(metric.label, metric.iconKey);
    if (iconKey && STEP_ICON_URIS[iconKey]) {
      const pad = ICON_D * 0.2;
      s.addImage({
        data: STEP_ICON_URIS[iconKey],
        x: iconX + pad, y: iconY + pad,
        w: ICON_D - pad * 2, h: ICON_D - pad * 2,
      });
    }

    // ── ラベル（アイコン右・小） ──
    s.addText(metric.label, {
      x: iconX + ICON_D + 0.1, y: iconY,
      w: cardW - ICON_PAD - ICON_D - 0.2, h: ICON_D,
      fontSize: theme.smallFontSize,
      fontFace: theme.fontFace,
      color: theme.palette.mutedText,
      valign: "middle",
    });

    // ── 表示値: LLMが設定した displayValue → value の順で使用 ──
    const displayText = String(metric.displayValue ?? metric.value ?? "").trim();
    const valueY = cy + ICON_PAD + ICON_D + 0.1;
    const noteY  = cy + cardH - NOTE_H - 0.06;
    const valueH = Math.max(noteY - valueY - 0.06, 0.4);

    s.addText(displayText, {
      x: cx + ICON_PAD, y: valueY,
      w: INNER_W, h: valueH,
      fontSize: Math.min(theme.bodyFontSize + 6 + fontAdj + valFontAdj, 24),
      fontFace: theme.fontFace,
      bold: true,
      color: theme.palette.bodyText,
      valign: "middle",
      align: "left",
      fit: "shrink",
    });

    if (metric.note) {
      const noteText = metric.note;
      s.addText(noteText, {
        x: cx + ICON_PAD, y: noteY,
        w: INNER_W, h: NOTE_H,
        fontSize: theme.smallFontSize - 1,
        fontFace: theme.fontFace,
        color: theme.palette.mutedText,
        valign: "top",
        fit: "shrink",
      });
    }
  });

  addChrome(s, theme);
}

// ─── process-cards レイアウト ──────────────────────────────────────────────

function buildProcessCardsSlide(pptx: PptxGenJS, slide: PptxSlide, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);

  let contentY = HEADER_H + 0.18;

  // サブタイトル
  const subtitle = slide.subtitle ?? (slide.bullets.length > 0 ? slide.bullets[0] : "");
  if (subtitle) {
    s.addText(subtitle, {
      x: 0.48, y: contentY,
      w: W - 0.96, h: 0.38,
      fontSize: theme.bodyFontSize - 1,
      fontFace: theme.fontFace,
      color: theme.palette.mutedText,
      valign: "middle",
      fit: "shrink",
    });
    contentY += 0.44;
  }

  const { density: pDensity, textTreatment: pText, visualIntent: pIntent } = slide;
  const pRenderMode = parseVisualIntent(pIntent);
  const pFontAdj   = densityFontAdj(pDensity) + (pRenderMode === "process" ? 1 : 0);
  const pSpaceMult = densitySpacingMult(pDensity) * (pRenderMode === "trust" ? 1.15 : 1.0);

  // bodyが空のstepは描画しない（LLMが生成した空ステップによる空白カードを防ぐ）
  const steps = (slide.steps ?? []).filter((s) => s.title?.trim() && s.body?.trim());
  const benefits = slide.benefits ?? [];
  const BENEFITS_H = benefits.length > 0 ? 0.52 : 0;
  const cardsH = H - contentY - BENEFITS_H - 0.26;
  // density=low → max 3ステップ（より広くゆったり）、density=high → 4ステップまで
  const count = densityMaxItems(pDensity, Math.min(steps.length, 4));
  const GAP = 0.14 * pSpaceMult;
  const totalW = W - 0.84;
  const cardW = (totalW - GAP * (count - 1)) / count;
  const ICON_D = 0.66;
  const ICON_TOP_PAD = 0.2;

  const STEP_ACCENT_CYCLE = [
    theme.palette.accentA,
    theme.palette.accentB,
    theme.palette.headerBg,
    theme.palette.accentB,
  ];
  steps.slice(0, count).forEach((step, idx) => {
    const cx = 0.42 + idx * (cardW + GAP);
    const numStr = String(idx + 1).padStart(2, "0");
    const accentFill = STEP_ACCENT_CYCLE[idx % STEP_ACCENT_CYCLE.length];

    const pcp = getCardStyleProps(theme, accentFill);
    s.addShape("roundRect", {
      x: cx, y: contentY, w: cardW, h: cardsH,
      rectRadius: pcp.rectRadius,
      fill: pcp.fill as any,
      line: pcp.line as any,
      ...(pcp.shadow ? { shadow: pcp.shadow } : {}),
    });
    // ウォーターマーク数字（glass/filled では低コントラストになるため透明度調整）
    s.addText(numStr, {
      x: cx + cardW - 0.98, y: contentY + 0.04,
      w: 0.88, h: 0.78,
      fontSize: theme.bodyFontSize + 6,
      fontFace: theme.fontFace, bold: true,
      color: theme.cardStyle === "filled" ? theme.palette.headerText : theme.palette.accentA,
      transparency: theme.cardStyle === "filled" ? 55 : 70,
      align: "right",
    });
    // アイコン円 + SVGアイコン（またはフォールバック数字）
    const iconX = cx + cardW / 2 - ICON_D / 2;
    const iconY = contentY + ICON_TOP_PAD;
    s.addShape("ellipse", {
      x: iconX, y: iconY, w: ICON_D, h: ICON_D,
      fill: { color: accentFill },
      line: { color: accentFill, width: 0 },
    });
    const iconKey = resolveStepIconKey(step.iconKey ?? step.title);
    if (iconKey) {
      const pad = ICON_D * 0.18;
      s.addImage({
        data: STEP_ICON_URIS[iconKey],
        x: iconX + pad, y: iconY + pad,
        w: ICON_D - pad * 2, h: ICON_D - pad * 2,
      });
    } else {
      s.addText(numStr, {
        x: iconX + 0.04, y: iconY + 0.04,
        w: ICON_D - 0.08, h: ICON_D - 0.08,
        fontSize: theme.bodyFontSize + 1, fontFace: theme.fontFace, bold: true,
        color: "FFFFFF", align: "center", valign: "middle",
      });
    }
    // ステップタイトル
    const titleY = contentY + ICON_TOP_PAD + ICON_D + 0.14;
    s.addText(step.title, {
      x: cx + 0.1, y: titleY, w: cardW - 0.2, h: 0.44,
      fontSize: theme.bodyFontSize + 1, fontFace: theme.fontFace, bold: true,
      color: theme.palette.bodyText, align: "center", valign: "middle", fit: "shrink",
    });
    // ステップ本文
    s.addText(step.body, {
      x: cx + 0.14, y: titleY + 0.48,
      w: cardW - 0.28, h: cardsH - (titleY - contentY) - 0.56,
      fontSize: Math.max(theme.bodyFontSize - 2 + pFontAdj, 9), fontFace: theme.fontFace,
      color: theme.palette.mutedText, valign: "top", fit: "shrink",
    });
    // 矢印
    if (idx < count - 1) {
      const arrowH = cardsH * 0.12;
      s.addShape("chevron", {
        x: cx + cardW + GAP * 0.12, y: contentY + cardsH / 2 - arrowH / 2,
        w: GAP * 0.68, h: arrowH,
        rotate: 0,
        fill: { color: theme.palette.accentA },
        line: { color: theme.palette.accentA, width: 0 },
      });
    }
  });

  // ── 下部メリット行（ピル型バッジ） ──
  if (benefits.length > 0) {
    const barY = H - BENEFITS_H - 0.18;
    s.addShape("rect", {
      x: 0.42, y: barY - 0.06, w: W - 0.84, h: 0.04,
      fill: { color: theme.palette.accentA, transparency: 65 },
      line: { color: theme.palette.accentA, width: 0 },
    });
    const pillW = (W - 1.6) / benefits.length;
    const pillH = BENEFITS_H - 0.08;
    benefits.forEach((benefit, idx) => {
      const px = 0.8 + idx * pillW;
      s.addShape("roundRect", {
        x: px, y: barY, w: pillW - 0.12, h: pillH,
        rectRadius: 0.14,
        fill: { color: theme.palette.accentA, transparency: 88 },
        line: { color: theme.palette.accentA, transparency: 55, width: 0.8 },
      });
      s.addText(`✓ ${benefit}`, {
        x: px + 0.06, y: barY + 0.02,
        w: pillW - 0.24, h: pillH - 0.04,
        fontSize: theme.smallFontSize, fontFace: theme.fontFace,
        bold: true, color: theme.palette.headerBg, valign: "middle", align: "center",
        fit: "shrink",
      });
    });
  }

  addChrome(s, theme);
}

// ─── closing レイアウト ────────────────────────────────────────────────────

function buildClosingSlide(pptx: PptxGenJS, slide: PptxSlide, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.titleBg };
  addTitleDecorativeCircles(s, theme);
  s.addText(slide.title, {
    x: 0.72, y: 1.8, w: W - 4.0, h: 1.4,
    fontSize: theme.titleFontSize + 8,
    fontFace: theme.fontFace, bold: true,
    color: theme.palette.headerText, align: "left", valign: "middle", fit: "shrink",
  });
  if (slide.bullets.length > 0) {
    const spaceMult = getSpacingMult(theme);
    slide.bullets.slice(0, 4).forEach((bullet, idx) => {
      const itemY = 3.4 + idx * 0.78 * spaceMult;
      // closing は暗背景なので cardStyle に関わらず accentA 半透明を維持
      s.addShape("roundRect", {
        x: 0.72, y: itemY, w: W - 4.5, h: 0.66,
        rectRadius: theme.cardStyle === "glass" ? 0.12 : 0.05,
        fill: { color: theme.palette.accentA, transparency: theme.cardStyle === "glass" ? 70 : 85 },
        line: { color: theme.palette.accentB, transparency: 40, width: 0.8 },
      });
      s.addText(bullet, {
        x: 0.94, y: itemY + 0.08, w: W - 4.9, h: 0.5,
        fontSize: theme.bodyFontSize, fontFace: theme.fontFace,
        color: theme.palette.headerText, valign: "middle", fit: "shrink",
      });
    });
  }
}

// ─── metric-cards レイアウト ───────────────────────────────────────────────

function buildMetricCardsSlide(pptx: PptxGenJS, slide: PptxSlide, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);

  const { density, textTreatment, visualIntent } = slide;
  const renderMode    = parseVisualIntent(visualIntent);
  const mFontAdj      = densityFontAdj(density) + intentValueFontAdj(renderMode);
  const mNoteExpand   = intentNoteExpand(renderMode);
  const fontScaleMult = slide.fontScale ?? 1.0;

  const contentY = HEADER_H + 0.24;
  const contentH = H - contentY - 0.36;
  const contentX = 0.42;
  const contentW = W - 0.84;

  const maxM = densityMaxItems(density, 4);
  const metrics = (slide.metrics ?? []).slice(0, maxM);
  if (metrics.length === 0) { addChrome(s, theme); return; }

  // density=low かつ 2枚以下 → 1列表示でゆったり
  const COLS = (density === "low" && metrics.length <= 2) ? 1 : 2;
  const ROWS = Math.ceil(metrics.length / COLS);
  const CARD_GAP = 0.2;
  const cardW = (contentW - CARD_GAP * (COLS - 1)) / COLS;
  const cardH = (contentH - CARD_GAP * (ROWS - 1)) / ROWS;

  const ICON_D   = Math.min(0.46 + (density === "low" ? 0.06 : 0), 0.54);
  const ICON_PAD = 0.22;
  const NOTE_H   = Math.min(0.40 * mNoteExpand, 0.56);
  const INNER_W  = cardW - ICON_PAD * 2;

  metrics.forEach((metric, idx) => {
    const col = idx % COLS;
    const row = Math.floor(idx / COLS);
    const cx = contentX + col * (cardW + CARD_GAP);
    const cy = contentY + row * (cardH + CARD_GAP);

    // colorRole 優先 → フォールバックは accentB(暖色)/accentA(主色) 交互
    const rawColor =
      metric.colorRole === "accent"  ? theme.palette.accentB :
      metric.colorRole === "neutral" ? theme.palette.headerBg :
      metric.colorRole === "primary" ? theme.palette.accentA :
      idx % 2 === 0 ? theme.palette.accentB : theme.palette.accentA;  // accentB(orange) を先頭に
    const topColor = rawColor;

    // カード背景（cardStyle に応じて変化）
    const cp = getCardStyleProps(theme, topColor);
    s.addShape("roundRect", {
      x: cx, y: cy, w: cardW, h: cardH,
      rectRadius: cp.rectRadius,
      fill: cp.fill as any,
      line: cp.line as any,
      ...(cp.shadow ? { shadow: cp.shadow } : {}),
    });
    // top accent strip（flat では非表示、filled では不要）
    if (theme.cardStyle !== "flat" && theme.cardStyle !== "filled") {
      s.addShape("rect", {
        x: cx, y: cy, w: cardW, h: 0.22,
        fill: { color: topColor },
        line: { color: topColor, width: 0 },
      });
    }

    // アイコン円（strip直下）
    const iconX = cx + ICON_PAD;
    const iconY = cy + 0.32;
    s.addShape("ellipse", {
      x: iconX, y: iconY, w: ICON_D, h: ICON_D,
      fill: { color: topColor },
      line: { color: topColor, width: 0 },
    });
    const iconKey = resolveMetricIconKey(metric.label, metric.iconKey);
    if (iconKey && STEP_ICON_URIS[iconKey]) {
      const pad = ICON_D * 0.2;
      s.addImage({
        data: STEP_ICON_URIS[iconKey],
        x: iconX + pad, y: iconY + pad,
        w: ICON_D - pad * 2, h: ICON_D - pad * 2,
      });
    }

    // ラベル（アイコン右横）
    s.addText(metric.label, {
      x: iconX + ICON_D + 0.1, y: iconY,
      w: cardW - ICON_PAD - ICON_D - 0.16, h: ICON_D,
      fontSize: theme.smallFontSize,
      fontFace: theme.fontFace,
      color: theme.palette.mutedText,
      valign: "middle",
    });

    // 表示値: displayValue → value の順
    const displayText = String(metric.displayValue ?? metric.value ?? "").trim();
    const valueY = iconY + ICON_D + 0.12;
    const noteY  = cy + cardH - NOTE_H - 0.08;
    const valueH = Math.max(noteY - valueY - 0.06, 0.36);

    // filled 時: 背景が topColor なので値は白系で表示
    const valueColor = theme.cardStyle === "filled" ? theme.palette.headerText : topColor;
    s.addText(displayText, {
      x: cx + ICON_PAD, y: valueY,
      w: INNER_W, h: valueH,
      fontSize: Math.min(theme.titleFontSize + 4 + mFontAdj, 26),
      fontFace: theme.fontFace,
      bold: true,
      color: valueColor,
      align: "center",
      valign: "middle",
      fit: "shrink",
    });

    if (metric.note) {
      const noteFontSize = textTreatment === "short"
        ? Math.max(8, Math.round((theme.smallFontSize - 2) * fontScaleMult))
        : Math.max(8, Math.round((theme.smallFontSize - 1) * fontScaleMult));
      const noteText = metric.note;
      s.addText(noteText, {
        x: cx + ICON_PAD, y: noteY,
        w: INNER_W, h: NOTE_H,
        fontSize: Math.max(noteFontSize, 8),
        fontFace: theme.fontFace,
        color: getCardMutedColor(theme),
        align: "center",
        valign: "top",
        fit: "shrink",
      });
    }
  });

  addChrome(s, theme);
}

// ─── timeline レイアウト ───────────────────────────────────────────────────

function buildTimelineSlide(pptx: PptxGenJS, slide: PptxSlide, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);

  const { density: tDensity, textTreatment: tText, visualIntent: tIntent } = slide;
  const tRenderMode = parseVisualIntent(tIntent);
  const tFontAdj    = densityFontAdj(tDensity);
  const tMaxSteps   = densityMaxItems(tDensity, 5);
  const tSpaceMult  = densitySpacingMult(tDensity);

  const validSteps = (slide.steps ?? []).filter((s) => s.title?.trim());
  const allFallback = slide.bullets.slice(0, tMaxSteps).map((b, i) => ({
    title: `フェーズ ${i + 1}`, body: b,
  } as PptxStep));
  const steps: PptxStep[] = (validSteps.length > 0 ? validSteps : allFallback).slice(0, tMaxSteps);

  if (steps.length === 0) { addChrome(s, theme); return; }

  const contentY = HEADER_H + 0.28;
  const contentH = H - contentY - 0.38;
  const startX = 0.56;
  const endX = W - 0.56;
  const lineW = endX - startX;
  // density=low → ラインをやや下に（上部タイトル領域を広く）
  const lineYRatio = tDensity === "low" ? 0.50 : tDensity === "high" ? 0.40 : 0.44;
  const lineY = contentY + contentH * lineYRatio;

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: startX, y: contentY, w: lineW, h: 0.38,
      fontSize: theme.bodyFontSize - 1, fontFace: theme.fontFace,
      color: theme.palette.mutedText, align: "center", italic: true,
    });
  }

  // ロードマップ段階色: 短期=info(ティール) / 中期=accent(橙) / 長期=primary(紺)
  const tok = resolveSemanticTokens(theme.palette);
  const ROADMAP_DOT_COLORS = [tok.info, tok.accent, tok.primary];

  // horizontal axis line（blueGray で控えめに）
  s.addShape("line", {
    x: startX, y: lineY, w: lineW, h: 0,
    line: { color: tok.blueGray, width: 2 },
  });

  const stepW = lineW / steps.length;

  steps.forEach((step, idx) => {
    const cx = startX + stepW * idx + stepW / 2;
    const dotR = 0.28;
    const dotColor = ROADMAP_DOT_COLORS[idx % ROADMAP_DOT_COLORS.length];

    // dot（glass/flat は枠を accentA で強調）
    s.addShape("ellipse", {
      x: cx - dotR, y: lineY - dotR, w: dotR * 2, h: dotR * 2,
      fill: { color: dotColor },
      line: {
        color: theme.cardStyle === "glass" ? theme.palette.accentA : theme.palette.surface,
        width: theme.cardStyle === "glass" ? 1.5 : 2,
      },
    });

    // icon inside dot
    const iconKey = resolveStepIconKey(step.iconKey ?? step.title);
    if (iconKey) {
      const pad = dotR * 0.28;
      s.addImage({
        data: STEP_ICON_URIS[iconKey],
        x: cx - dotR + pad, y: lineY - dotR + pad,
        w: dotR * 2 - pad * 2, h: dotR * 2 - pad * 2,
      });
    } else {
      s.addText(String(idx + 1), {
        x: cx - dotR, y: lineY - dotR,
        w: dotR * 2, h: dotR * 2,
        fontSize: theme.smallFontSize - 1, fontFace: theme.fontFace,
        bold: true, color: "FFFFFF",
        align: "center", valign: "middle",
      });
    }

    // title above line
    s.addText(step.title, {
      x: cx - stepW * 0.42, y: lineY - dotR - 0.98,
      w: stepW * 0.84, h: 0.52,
      fontSize: Math.max(theme.bodyFontSize - 1 + tFontAdj, 9),
      fontFace: theme.fontFace,
      bold: true, color: theme.palette.bodyText,
      align: "center", valign: "bottom", fit: "shrink",
    });

    // body below line
    const bodyY = lineY + dotR + 0.16;
    const bodyH = H - bodyY - 0.4;
    // textTreatment=explanatory → bodyを大きめフォント、short → 省略ぎみ
    const bodyFontSize = tText === "explanatory"
      ? Math.max(theme.smallFontSize + tFontAdj, 9)
      : tText === "short"
        ? Math.max(theme.smallFontSize - 1 + tFontAdj, 8)
        : Math.max(theme.smallFontSize + tFontAdj, 8);
    s.addText(step.body, {
      x: cx - stepW * 0.42, y: bodyY,
      w: stepW * 0.84, h: bodyH,
      fontSize: bodyFontSize, fontFace: theme.fontFace,
      color: theme.palette.bodyText, align: "center", valign: "top", fit: "shrink",
    });
  });

  if (slide.benefits?.length) {
    // フッターバンド禁止 — テキストのみ（背景なし）でスライド下部に表示
    s.addText(slide.benefits.slice(0, 4).map((b) => `✓ ${b}`).join("　　"), {
      x: 0.42, y: H - 0.56, w: W - 0.84, h: 0.38,
      fontSize: theme.smallFontSize - 1, fontFace: theme.fontFace,
      color: theme.palette.mutedText, bold: false, align: "center",
    });
  }

  addChrome(s, theme);
}

// ─── テキストのみスライド検出・アップグレード ────────────────────────────────────
// bullets + 視覚データなし → icon_rows に自動変換してビジュアルを必ず確保する

const UPGRADE_ICON_CYCLE = ["gear", "lightbulb", "rocket", "chart", "star", "verified", "shield", "network"] as const;

/** bullet 1本を heading / body に分割する共通ロジック。
 *  body が助詞（は/が/を…）で始まる場合は全文を body に戻して不自然な分割を防ぐ。 */
function splitBulletForCard(b: string): { heading: string; body: string } {
  const full = b.trim();
  const colonIdx = full.indexOf("：") >= 0 ? full.indexOf("：") : full.indexOf(":");
  if (colonIdx > 0 && colonIdx <= 24) {
    const bd = full.slice(colonIdx + 1).trim();
    return { heading: full.slice(0, colonIdx).trim(), body: LEADING_PARTICLE_RE.test(bd) ? full : bd };
  }
  const parenIdx = full.indexOf("（");
  if (parenIdx > 0 && parenIdx <= 15) {
    const bd = full.slice(parenIdx).trim();
    return { heading: full.slice(0, parenIdx).trim(), body: LEADING_PARTICLE_RE.test(bd) ? full : bd };
  }
  const totenIdx = full.indexOf("、");
  if (totenIdx > 0 && totenIdx <= 15) {
    const bd = full.slice(totenIdx + 1).trim();
    return { heading: full.slice(0, totenIdx).trim(), body: LEADING_PARTICLE_RE.test(bd) ? full : bd };
  }
  const spIdx = full.indexOf("　") >= 0 ? full.indexOf("　") : full.indexOf(" ");
  if (spIdx > 0 && spIdx <= 10) {
    const bd = full.slice(spIdx + 1).trim();
    return { heading: full.slice(0, spIdx).trim(), body: LEADING_PARTICLE_RE.test(bd) ? full : bd };
  }
  // は/が: heading = 助詞前、body = 全文（「は完了しており...」という body を生成しない）
  for (let j = 2; j <= Math.min(22, full.length - 2); j++) {
    if (full[j] === "は" || full[j] === "が") {
      return { heading: full.slice(0, j).trim(), body: full };
    }
  }
  return { heading: "", body: full };
}

function autoCardsFromBullets(bullets: string[]): PptxCard[] {
  return bullets.slice(0, 4).map((b, i) => {
    const iconKey = UPGRADE_ICON_CYCLE[i % UPGRADE_ICON_CYCLE.length] as string;
    const { heading, body } = splitBulletForCard(b);
    return { iconKey, heading, body };
  });
}

function upgradeTextOnlySlide(slide: PptxSlide): PptxSlide {
  const lt = slide.layoutType;
  if (lt && lt !== "bullets") return slide;
  const bullets = (slide.bullets ?? []).filter((b: string) => b?.trim());
  const hasCards   = (slide.cards   ?? []).some((c: PptxCard) => c.heading?.trim());
  const hasSteps   = (slide.steps   ?? []).some((s: any) => s.title?.trim());
  const hasMetrics = (slide.metrics ?? []).some((m: any) => m.label?.trim());
  const hasTable   = (slide.tableRows ?? []).length >= 2;
  // 5件以上は bullets レイアウトのまま維持（カード化すると5件目以降が落ちる）
  if (bullets.length >= 2 && bullets.length <= 4 && !hasCards && !hasSteps && !hasMetrics && !hasTable) {
    console.log(`[gen-pptx] upgradeTextOnly: "${slide.title}" bullets→icon_rows`);
    return { ...slide, layoutType: "icon_rows" as const, cards: autoCardsFromBullets(bullets) };
  }
  return slide;
}

// ─── stat_callouts レイアウト ─────────────────────────────────────────────────
// 大きな数字3つ（KPI/実績）+ バーチャート。数値を必ずビジュアル化する。

function statValueFontSize(value: string): number {
  const v = value.trim();
  const hasComplex = /[ぁ-んァ-ヶ一-龥（）【】「」¥￥]/.test(v);
  if (v.length >= 9 || (v.length >= 6 && hasComplex)) return 26;
  if (v.length >= 5) return 32;
  return 42;
}

function buildStatCalloutsSlide(pptx: PptxGenJS, slide: PptxSlide, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  const fontScaleMult = slide.fontScale ?? 1.0;

  // statCallouts → なければ metrics を代用
  const rawCallouts: { value: string; unit: string; label: string }[] = (slide.statCallouts ?? []).length > 0
    ? slide.statCallouts!
    : (slide.metrics ?? []).map((m) => ({ value: m.displayValue ?? m.value, unit: "", label: m.label }));
  const callouts = rawCallouts.filter((c) => c.value?.trim()).slice(0, 3);

  // KPIカード面: surfaceBlue/surfaceOrange 交互で意味の階層を作る（上端帯廃止）
  const tokKpi = resolveSemanticTokens(theme.palette);
  const CARD_FILLS = [tokKpi.surfaceBlue, tokKpi.surfaceOrange, tokKpi.surfaceBlue];
  const CARD_H = 1.72;
  const CARD_Y = HEADER_H + 0.28;
  const count = Math.min(callouts.length, 3);

  if (count > 0) {
    const GAP = 0.18;
    const CARD_W = (W - 0.84 - GAP * (count - 1)) / count;

    callouts.forEach((item, idx) => {
      const cx = 0.42 + idx * (CARD_W + GAP);
      const cardFill = CARD_FILLS[idx % CARD_FILLS.length];

      // カード背景（淡色面 — 上端ストリップ廃止・面の淡色+枠で差別化）
      s.addShape("roundRect", {
        x: cx, y: CARD_Y, w: CARD_W, h: CARD_H,
        rectRadius: 0.08,
        fill: { color: cardFill },
        line: { color: tokKpi.borderGray, width: 0.8 },
        shadow: { type: "outer", color: "000000", blur: 4, angle: 270, opacity: 0.06 },
      });
      // 数値（primaryDark で統一 — 淡色面上でコントラスト確保）— 長い値・日本語を含む場合は自動縮小
      s.addText(item.value, {
        x: cx + 0.18, y: CARD_Y + 0.2,
        w: CARD_W * 0.7, h: 0.86,
        fontSize: statValueFontSize(item.value),
        fontFace: theme.fontFace, bold: true,
        color: tokKpi.primaryDark,
        valign: "middle", fit: "shrink",
      });
      // 単位（数値右横、小さめ）
      if (item.unit) {
        s.addText(item.unit, {
          x: cx + CARD_W * 0.72, y: CARD_Y + 0.55,
          w: CARD_W * 0.22, h: 0.38,
          fontSize: Math.max(8, Math.round(theme.bodyFontSize * fontScaleMult)),
          fontFace: theme.fontFace,
          color: theme.palette.mutedText,
          valign: "bottom", fit: "shrink",
        });
      }
      // ラベル（下部）
      s.addText(item.label, {
        x: cx + 0.18, y: CARD_Y + 1.14,
        w: CARD_W - 0.36, h: 0.44,
        fontSize: theme.smallFontSize,
        fontFace: theme.fontFace,
        color: theme.palette.mutedText,
        valign: "top", fit: "shrink",
      });
    });
  }

  // バーチャート: 数値化可能なcalloutsがあれば bar chart で可視化
  const CHART_Y = CARD_Y + CARD_H + 0.2;
  const CHART_H = H - CHART_Y - 0.28;
  const chartValues: number[] = [];
  const chartLabels: string[] = [];
  callouts.forEach((c) => {
    const clean = c.value.replace(/[約,，\s]/g, "").replace(/[〜～].+$/, "").match(/[\d.]+/)?.[0];
    const n = clean ? parseFloat(clean) : NaN;
    if (!isNaN(n)) { chartValues.push(n); chartLabels.push(c.label.slice(0, 12)); }
  });
  if (chartValues.length >= 2) {
    s.addChart("bar" as any, [{ name: "KPI", labels: chartLabels, values: chartValues }], {
      x: 0.42, y: CHART_Y, w: W - 0.84, h: CHART_H,
      barDir: "bar",
      chartColors: [tokKpi.primaryDark, tokKpi.accent, tokKpi.info],
      showValue: true,
      showLegend: false,
      valAxisLabelFontSize: 9,
      catAxisLabelFontSize: 9,
      dataLabelFontSize: 9,
      dataLabelColor: "363636",
    });
  } else {
    // チャート不可の場合のみ bullets フォールバック
    const bullets = (slide.bullets ?? []).filter((b) => b?.trim()).slice(0, 4);
    if (bullets.length > 0) {
      const bulletItems = bullets.map((b) => ({
        text: b,
        options: { bullet: { indent: 14 }, breakLine: true, fontSize: Math.max(9, Math.round(theme.bodyFontSize * fontScaleMult)), fontFace: theme.fontFace, color: theme.palette.bodyText, paraSpaceAfter: 8 },
      }));
      s.addText(bulletItems, { x: 0.52, y: CHART_Y, w: W - 1.04, h: CHART_H, margin: 0.1, valign: "top" });
    }
  }
}

// ─── card_grid レイアウト ─────────────────────────────────────────────────────
// アイコン付きカードを2〜3列グリッドで並べる。features / benefits 向け。

function buildCardGridSlide(pptx: PptxGenJS, slide: PptxSlide, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  const fontScaleMult = slide.fontScale ?? 1.0;

  // cards → なければ steps を代用
  const rawCards: PptxCard[] = (slide.cards ?? []).length > 0
    ? slide.cards!
    : (slide.steps ?? []).map((st) => ({ iconKey: st.iconKey, heading: st.title, body: st.body }));
  const cards = rawCards.filter((c) => c.heading?.trim()).slice(0, 6);

  if (cards.length === 0) {
    // フォールバック: bullets で描画
    const bulletItems = (slide.bullets ?? []).slice(0, 5).map((b) => ({
      text: b,
      options: { bullet: { indent: 14 }, breakLine: true, fontSize: theme.bodyFontSize, fontFace: theme.fontFace, color: theme.palette.bodyText, paraSpaceAfter: 9 },
    }));
    if (bulletItems.length > 0) s.addText(bulletItems, { x: 0.52, y: HEADER_H + 0.24, w: W - 1.04, h: H - HEADER_H - 0.5, margin: 0.1, valign: "top" });
    return;
  }

  const count = cards.length;
  const cols = count <= 3 ? count : count <= 4 ? 2 : 3;
  const rows = Math.ceil(count / cols);

  const MARGIN_X = 0.42;
  const START_Y = HEADER_H + 0.28;
  const AVAIL_W = W - MARGIN_X * 2;
  const AVAIL_H = H - START_Y - 0.22;
  const GAP_X = 0.18;
  const GAP_Y = 0.18;
  const CARD_W = (AVAIL_W - GAP_X * (cols - 1)) / cols;
  const CARD_H = (AVAIL_H - GAP_Y * (rows - 1)) / rows;
  const ICON_D = Math.min(0.68, CARD_H * 0.28);

  const ICON_COLORS = [theme.palette.accentA, theme.palette.headerBg, theme.palette.accentB,
                       theme.palette.accentA, theme.palette.accentB, theme.palette.headerBg];

  cards.forEach((card, idx) => {
    const col = idx % cols;
    const row = Math.floor(idx / cols);
    const cx = MARGIN_X + col * (CARD_W + GAP_X);
    const cy = START_Y + row * (CARD_H + GAP_Y);
    const iconColor = ICON_COLORS[idx % ICON_COLORS.length];

    // カード背景（白 + 薄い影）
    s.addShape("roundRect", {
      x: cx, y: cy, w: CARD_W, h: CARD_H,
      rectRadius: 0.08,
      fill: { color: "FFFFFF" },
      line: { color: theme.palette.border, width: 0.8 },
      shadow: { type: "outer", color: "000000", blur: 5, angle: 270, opacity: 0.07 },
    });

    // アイコン円（左上）
    const ICON_PAD_X = 0.20;
    const ICON_PAD_Y = CARD_H * 0.14;
    const iconX = cx + ICON_PAD_X;
    const iconY = cy + ICON_PAD_Y;
    s.addShape("ellipse", {
      x: iconX, y: iconY, w: ICON_D, h: ICON_D,
      fill: { color: iconColor },
      line: { color: iconColor, width: 0 },
    });
    const iconSvg = card.iconKey ? STEP_ICON_URIS[card.iconKey as keyof typeof STEP_ICON_URIS] : null;
    if (iconSvg) {
      const pad = ICON_D * 0.18;
      s.addImage({ data: iconSvg, x: iconX + pad, y: iconY + pad, w: ICON_D - pad * 2, h: ICON_D - pad * 2 });
    } else {
      s.addText(String(idx + 1), {
        x: iconX, y: iconY, w: ICON_D, h: ICON_D,
        fontSize: theme.bodyFontSize + 1, fontFace: theme.fontFace, bold: true,
        color: "FFFFFF", align: "center", valign: "middle",
      });
    }

    // 見出し（アイコン右横）
    const textX = cx + ICON_PAD_X + ICON_D + 0.14;
    const textW = CARD_W - ICON_PAD_X - ICON_D - 0.22;
    s.addText(card.heading, {
      x: textX, y: iconY, w: textW, h: ICON_D,
      fontSize: Math.max(9, Math.round((theme.bodyFontSize + 1) * fontScaleMult)), fontFace: theme.fontFace, bold: true,
      color: theme.palette.bodyText, valign: "middle", fit: "shrink",
    });

    // 本文（アイコン下）
    if (card.body) {
      const bodyY = cy + ICON_PAD_Y + ICON_D + 0.1;
      s.addText(card.body, {
        x: cx + ICON_PAD_X, y: bodyY,
        w: CARD_W - ICON_PAD_X * 2,
        h: CARD_H - (bodyY - cy) - 0.14,
        fontSize: Math.max(8, Math.round((theme.bodyFontSize - 2) * fontScaleMult)),
        fontFace: theme.fontFace,
        color: theme.palette.mutedText,
        valign: "top", fit: "shrink",
      });
    }
  });
}

// ─── icon_rows レイアウト ─────────────────────────────────────────────────────
// アイコン + 見出し + 本文の行形式。ステータスピル付き対応。プロセス/機能一覧向け。

function buildIconRowsSlide(pptx: PptxGenJS, slide: PptxSlide, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  const fontScaleMult = slide.fontScale ?? 1.0;

  const rawRows: PptxCard[] = (slide.cards ?? []).length > 0
    ? slide.cards!
    : (slide.steps ?? []).map((st) => ({ iconKey: st.iconKey, heading: st.title, body: st.body }));
  const rows = rawRows.filter((r) => r.heading?.trim()).slice(0, 4);

  if (rows.length === 0) {
    const bulletItems = (slide.bullets ?? []).slice(0, 5).map((b) => ({
      text: b,
      options: { bullet: { indent: 14 }, breakLine: true, fontSize: theme.bodyFontSize, fontFace: theme.fontFace, color: theme.palette.bodyText, paraSpaceAfter: 9 },
    }));
    if (bulletItems.length > 0) s.addText(bulletItems, { x: 0.52, y: HEADER_H + 0.24, w: W - 1.04, h: H - HEADER_H - 0.5, margin: 0.1, valign: "top" });
    return;
  }

  const ICON_D = 0.58;
  const START_Y = HEADER_H + 0.26;
  const AVAIL_H = H - START_Y - 0.22;
  const ROW_SLOT = AVAIL_H / rows.length;
  const ROW_H = Math.min(ROW_SLOT - 0.08, 1.0);

  // サマリ行: アイコン=info 統一・行ごとに surfaceGray/surfaceBlue を交互に
  const tokRow = resolveSemanticTokens(theme.palette);
  const ROW_SURFACES = [tokRow.surfaceGray, tokRow.surfaceBlue];

  rows.forEach((row, idx) => {
    const cy = START_Y + idx * ROW_SLOT + (ROW_SLOT - ROW_H) / 2;
    const iconColor = tokRow.info;  // 機能アイコンは info で統一（視認性と意味の一貫性）
    const rowSurface = ROW_SURFACES[idx % ROW_SURFACES.length];

    // 全行に淡色面（交互）
    s.addShape("roundRect", {
      x: 0.38, y: cy - 0.04, w: W - 0.76, h: ROW_H + 0.08,
      rectRadius: 0.06,
      fill: { color: rowSurface },
      line: { color: tokRow.borderGray, width: 0.5 },
    });

    // アイコン円
    const iconX = 0.52;
    const iconY = cy + ROW_H / 2 - ICON_D / 2;
    s.addShape("ellipse", {
      x: iconX, y: iconY, w: ICON_D, h: ICON_D,
      fill: { color: iconColor },
      line: { color: iconColor, width: 0 },
    });
    const iconSvg = row.iconKey ? STEP_ICON_URIS[row.iconKey as keyof typeof STEP_ICON_URIS] : null;
    if (iconSvg) {
      const pad = ICON_D * 0.18;
      s.addImage({ data: iconSvg, x: iconX + pad, y: iconY + pad, w: ICON_D - pad * 2, h: ICON_D - pad * 2 });
    } else {
      s.addText(String(idx + 1), {
        x: iconX, y: iconY, w: ICON_D, h: ICON_D,
        fontSize: theme.bodyFontSize, fontFace: theme.fontFace, bold: true,
        color: "FFFFFF", align: "center", valign: "middle",
      });
    }

    // テキスト領域
    const PILL_W = row.statusLabel ? 1.9 : 0;
    const textX = iconX + ICON_D + 0.18;
    const textW = W - textX - PILL_W - 0.5;

    s.addText(row.heading, {
      x: textX, y: cy,
      w: textW, h: ROW_H * 0.44,
      fontSize: Math.max(9, Math.round((theme.bodyFontSize + 1) * fontScaleMult)), fontFace: theme.fontFace, bold: true,
      color: theme.palette.bodyText, valign: "bottom", fit: "shrink",
    });
    if (row.body) {
      s.addText(row.body, {
        x: textX, y: cy + ROW_H * 0.44,
        w: textW, h: ROW_H * 0.56,
        fontSize: Math.max(8, Math.round((theme.bodyFontSize - 1) * fontScaleMult)), fontFace: theme.fontFace,
        color: theme.palette.mutedText, valign: "top", fit: "shrink",
      });
    }

    // ステータスピル（セマンティックカラー: 完了=success / 進行中=accent / その他=primaryMid）
    if (row.statusLabel) {
      const PILL_H = 0.30;
      const pillX = W - PILL_W - 0.42;
      const pillY = cy + ROW_H / 2 - PILL_H / 2;
      const pillColor = /完了|done|finished|済/i.test(row.statusLabel)
        ? tokRow.success
        : /進行中|in.?progress|開発中|実装中|対応中/i.test(row.statusLabel)
          ? tokRow.accent
          : tokRow.primaryMid;
      s.addShape("roundRect", {
        x: pillX, y: pillY, w: PILL_W, h: PILL_H,
        rectRadius: PILL_H / 2,
        fill: { color: pillColor, transparency: 85 },
        line: { color: pillColor, width: 0.8 },
      });
      s.addText(row.statusLabel, {
        x: pillX, y: pillY, w: PILL_W, h: PILL_H,
        fontSize: theme.smallFontSize - 1, fontFace: theme.fontFace, bold: true,
        color: pillColor, align: "center", valign: "middle",
      });
    }
  });
}

export async function POST(req: NextRequest) {
  try {
    const body: GenPptxRequest = await req.json();
    const { title, slides, threadId, fontFace, designInstruction, deckPreferences, mode, fileBaseName, promptIntent, palette: requestedPalette } = body;
    if (!title || !slides || slides.length === 0) {
      return NextResponse.json({ error: "title and slides are required" }, { status: 400 });
    }

    const faithfulMode = mode === "faithful";

    const instructionText = [designInstruction, deckPreferences?.designInstruction, ...(deckPreferences?.recentDesignNotes ?? [])]
      .filter(Boolean)
      .join(" / ");

    // PromptIntent ログ
    if (promptIntent) {
      const ld = promptIntent.layoutDirectives;
      const cd = promptIntent.colorDirectives;
      console.log(
        `[PromptIntent] purpose=${promptIntent.documentPurpose} audience=${promptIntent.audience} freedom=${promptIntent.designFreedom}` +
        ` twoColumn=${!!ld.preferTwoColumn} tables=${!!ld.includeTables} metrics=${!!ld.preferMetrics} process=${!!ld.preferProcess}` +
        (cd ? ` colors=${cd.primary ?? "?"}/${cd.accent ?? "?"}` : "")
      );
    }

    // faithfulモード: デザインAIをスキップしてフォールバックブリーフを直接使用
    const designBrief = faithfulMode
      ? createFallbackBrief(title, slides, instructionText, deckPreferences)
      : await generateDesignBrief(title, slides, instructionText, deckPreferences, undefined, promptIntent);

    // 固定パレット適用: リクエストで palette 名が指定された場合は優先、なければキーワード選択
    const strictKey = selectStrictPaletteKey(instructionText, promptIntent ?? undefined);
    const namedPalette = typeof requestedPalette === "string" && requestedPalette in PPTX_PALETTES
      ? requestedPalette : null;
    designBrief.palette = namedPalette
      ? buildPaletteFromName(namedPalette)
      : buildStrictPalette(strictKey);
    // タイトル帯は常に塗り潰しバンド（濃色背景＋白文字）に固定
    designBrief.styleSpec.headerStyle = STRICT_HEADER_STYLE;
    console.log(`[palette] ${namedPalette ? `named=${namedPalette}` : `strict=${strictKey}`} header=${STRICT_HEADER_STYLE}`);

    const theme = resolveTheme(designBrief, instructionText, deckPreferences, fontFace);
    // faithfulモード: イラスト生成をスキップ
    const coverIllustration = faithfulMode ? null : await generateCoverIllustration(title, instructionText, theme);

    console.log("[gen-pptx] theme:", {
      instruction: truncateText(instructionText, 80),
      accentColor: deckPreferences?.accentColor,
      language: deckPreferences?.language,
      recentNotes: deckPreferences?.recentDesignNotes,
      execMode: theme.execMode,
      fontFace: theme.fontFace,
      coverIllustration: Boolean(coverIllustration),
    });

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.theme  = { headFontFace: "Meiryo", bodyFontFace: "Meiryo" };
    pptx.author = "azurechat";
    pptx.subject = title;
    pptx.title = title;

    // faithfulモード: 自動タイトルスライドを追加しない（元ページ数を維持）
    if (!faithfulMode) {
      buildTitleSlide(pptx, title, designBrief, theme, slides.length + 1, coverIllustration);
    }
    let illustrationPlaced = false;
    // スライドタイトルから誤って付いた [H] プレフィックスを除去
    // 非faithfulモードでは自動表紙と重複するAI生成「表紙」スライドを除去
    const COVER_TITLES = new Set(["表紙", "タイトル", "cover", "title slide", "title"]);
    const rawCleaned = slides
      .map((slide) => ({
        ...slide,
        title: slide.title.replace(/^\[H\]\s*/i, "").trim(),
      }))
      .filter((slide, index) => {
        if (!faithfulMode && index === 0 && COVER_TITLES.has(slide.title.toLowerCase().trim())) {
          console.log("[gen-pptx] skipping duplicate cover slide:", slide.title);
          return false;
        }
        return true;
      });

    // PromptIntent の layoutDirectives を先に適用してから品質ゲートを通す
    const intentApplied = promptIntent && !faithfulMode
      ? applyPromptIntentToSlides(rawCleaned, promptIntent)
      : rawCleaned;
    const sanitizedSlides = validateAndRepairSlides(normalizeSlidesForPptx(intentApplied));

    // ── guaranteed chart: stat_callouts がなければ最も数値が多いスライドを強制変換 ──
    // 年 (2000〜2099) のみの値は有効なKPIとみなさない（既存 stat_callouts でも適用）
    const isYearOnlyValue = (val: string, unit: string) => {
      const u = String(unit ?? "").trim();
      if (u && !/^年度?$/.test(u)) return false;
      const n = parseFloat(String(val).replace(/[,，〜～].+$/, "").replace(/[,，]/g, ""));
      return n >= 2000 && n <= 2099;
    };
    const hasParseableChart = sanitizedSlides.some((sl) => {
      if (sl.layoutType !== "stat_callouts") return false;
      const parseable = (sl.statCallouts ?? []).filter((c: any) => {
        const clean = String(c.value ?? "").replace(/[約,，\s]/g, "").replace(/[〜～].+$/, "").match(/[\d.]+/)?.[0];
        if (!clean || isNaN(parseFloat(clean))) return false;
        return !isYearOnlyValue(clean, String(c.unit ?? ""));
      });
      return parseable.length >= 2;
    });
    if (!hasParseableChart && !faithfulMode) {
      // 年・年度をKPI値として採用しないための抽出ヘルパー（金額・割合・件数を年より優先）
      const extractKpiValue = (text: string): { value: string; unit: string } | null => {
        // 年・年度・日付範囲 (2024/10〜2026/3 等) を除去してからKPI数値を探す
        const t = text.replace(/\b20\d{2}(?:\/\d{1,2}(?:[〜～]20\d{2}\/\d{1,2})?|年度?[のはがを]?)?/g, "▪");
        // 優先0: 通貨記号付き (¥1.1M, ¥763千, ¥110千)
        const c0 = t.match(/[¥￥]\s*[\d,，.]+(?:\.\d+)?\s*(?:万|億|千|兆|[MBKk])/);
        if (c0) {
          const raw = c0[0].replace(/^[¥￥]\s*/, "");
          const u = raw.match(/(?:万|億|千|兆|[MBKk])$/)?.[0] ?? "";
          const num = raw.replace(/(?:万|億|千|兆|[MBKk])$/, "").replace(/[,，\s]/g, "");
          return { value: "¥" + num, unit: u };
        }
        // 優先1: 範囲値 + 有意単位 (9〜12万円, 43〜46%)
        const r = t.match(/(?:約\s*)?[\d,，.]+[〜～][\d,，.]+\s*(?:万|億|千)?[円%件人回倍社名台個本枚兆]/);
        if (r) {
          const raw = r[0].replace(/^約\s*/, "");
          const u = raw.match(/(?:万|億|千)?[円%件人回倍社名台個本枚兆]$/)?.[0] ?? "";
          return { value: raw.replace(/(?:万|億|千)?[円%件人回倍社名台個本枚兆]$/, "").replace(/[,，\s]/g, ""), unit: u };
        }
        // 優先2: 数値 + 有意単位 (240万円, 99.9%, 746件, 1,100千円)
        const m = t.match(/(?:約\s*)?[\d,，.]+(?:\.\d+)?\s*(?:万|億|千)?[円%件人回倍社名台個本枚兆]/);
        if (m) {
          const raw = m[0].replace(/^約\s*/, "");
          const u = raw.match(/(?:万|億|千)?[円%件人回倍社名台個本枚兆]$/)?.[0] ?? "";
          return { value: raw.replace(/(?:万|億|千)?[円%件人回倍社名台個本枚兆]$/, "").replace(/[,，\s]/g, ""), unit: u };
        }
        return null; // 年のみ、またはKPI数値なし → statCallouts 候補にしない
      };
      const hasKpi = (s: string) => extractKpiValue(s) !== null;

      let bestIdx = -1;
      let bestCount = 0;
      sanitizedSlides.forEach((sl, i) => {
        const lt = sl.layoutType;
        if (lt === "closing" || lt === "title" || lt === "stat_callouts") return;
        // bullets と cards.body/heading の両方をカウント（年のみ bullet は除外）
        const bulletCount = (sl.bullets ?? []).filter((b: string) => hasKpi(b)).length;
        const cardCount = (sl.cards ?? []).filter((c: any) =>
          hasKpi(c.body ?? "") || hasKpi(c.heading ?? "")
        ).length;
        const count = bulletCount + cardCount;
        if (count > bestCount) { bestCount = count; bestIdx = i; }
      });
      if (bestIdx >= 0 && bestCount >= 2) {
        const src = sanitizedSlides[bestIdx];
        const numTexts: string[] = [
          ...(src.bullets ?? []).filter((b: string) => hasKpi(b)),
          ...(src.cards ?? [])
            .filter((c: any) => hasKpi(c.body ?? "") || hasKpi(c.heading ?? ""))
            .map((c: any) => (hasKpi(c.body ?? "") ? c.body : c.heading) as string),
        ].slice(0, 4);
        const statCallouts = numTexts.map((b: string) => {
          const kpi = extractKpiValue(b)!;
          // ラベル: 年・KPI値パターンをすべて除去して残った文脈テキスト（カンマ付き元文も除去）
          const label = b
            .replace(/\b20\d{2}(?:\/\d{1,2}(?:[〜～]20\d{2}\/\d{1,2})?|年度?[のはがをも]?)/g, "")
            .replace(/[¥￥]\s*[\d,，.]+(?:\.\d+)?\s*(?:万|億|千|兆|[MBKk])/g, "")
            .replace(/(?:約\s*)?[\d,，.]+(?:[〜～][\d,，.]+)?\s*(?:万|億|千)?[円%件人回倍社名台個本枚兆]/g, "")
            .replace(/[：:。、・]\s*/g, " ")
            .trim()
            .slice(0, 18) || b.slice(0, 18);
          return { value: kpi.value, unit: kpi.unit, label };
        });
        console.log(`[gen-pptx] guaranteedChart: slide[${bestIdx}] "${src.title}" (${src.layoutType}) → stat_callouts`);
        (sanitizedSlides as any[])[bestIdx] = { ...src, layoutType: "stat_callouts", statCallouts };
      }
    }

    sanitizedSlides.forEach((slideRaw, index) => {
      const slide = upgradeTextOnlySlide(slideRaw);
      const resolvedLt = resolveLayoutType(slide);
      const visual = designBrief.visualHints[index] ?? {
        title: slide.title,
        visualType:
          resolvedLt === "table"         ? "table" :
          resolvedLt === "multi-column"  ? "comparison" :
          resolvedLt === "diagram" || resolvedLt === "process-cards" ? "process" :
          resolvedLt === "timeline"      ? "timeline" :
          resolvedLt === "company-overview" || resolvedLt === "metric-cards" ? "cards" :
          resolvedLt === "closing"       ? "spotlight" : "editorial",
        emphasis: (Array.isArray(slide.bullets) && slide.bullets[0]) || slide.title,
      };
      const slideIllustration =
        !illustrationPlaced &&
        coverIllustration &&
        resolvedLt === "bullets"
          ? coverIllustration
          : null;
      switch (resolvedLt) {
        case "title":
          buildSectionSlide(pptx, slide.title, theme);
          break;
        case "table":
          buildTableSlide(pptx, slide, theme, visual, faithfulMode);
          break;
        case "multi-column":
          buildMultiColumnSlide(pptx, slide, theme, visual, faithfulMode);
          break;
        case "diagram":
          buildDiagramSlide(pptx, slide, theme, visual, faithfulMode);
          break;
        case "conversation":
          buildConversationSlide(pptx, slide, theme);
          break;
        case "company-overview":
          buildCompanyOverviewSlide(pptx, slide, theme);
          break;
        case "process-cards":
          buildProcessCardsSlide(pptx, slide, theme);
          break;
        case "closing":
          buildClosingSlide(pptx, slide, theme);
          break;
        case "metric-cards":
          buildMetricCardsSlide(pptx, slide, theme);
          break;
        case "timeline":
          buildTimelineSlide(pptx, slide, theme);
          break;
        case "stat_callouts":
          buildStatCalloutsSlide(pptx, slide, theme);
          break;
        case "card_grid":
          buildCardGridSlide(pptx, slide, theme);
          break;
        case "icon_rows":
          buildIconRowsSlide(pptx, slide, theme);
          break;
        case "roadmap":
          buildTimelineSlide(pptx, slide, theme);
          break;
        default:
          buildBulletsSlide(pptx, slide, theme, visual, slideIllustration, faithfulMode);
          if (slideIllustration) illustrationPlaced = true;
          break;
      }
    });

    let buffer = await patchEastAsianFont((await pptx.write({ outputType: "nodebuffer" })) as Buffer);
    // PPTX バリデーション: ZIP magic bytes + スライドXML存在確認
    {
      const isZip = buffer && buffer.length >= 100 && buffer[0] === 0x50 && buffer[1] === 0x4B;
      const hasSlides = isZip && buffer.includes(Buffer.from("ppt/slides/slide"));
      if (!isZip) throw new Error("[gen-pptx] Generated PPTX is not a valid ZIP file — aborting");
      if (!hasSlides) throw new Error("[gen-pptx] Generated PPTX has no slide content — aborting");
    }

    // ── Phase 2: Vision レビュー → パッチ適用 → 再生成 ──────────────────
    if (!faithfulMode && process.env.PPTX_VISION_REVIEW_ENABLED === "true") {
      try {
        const reviewForm = new FormData();
        const pptxBytes = new Uint8Array(buffer.byteLength);
        pptxBytes.set(buffer);
        reviewForm.append("pptx", new Blob([pptxBytes], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
        reviewForm.append("title", title);
        if (promptIntent) {
          reviewForm.append("promptIntent", JSON.stringify(promptIntent));
        }

        const baseUrl = (
          process.env.NEXTAUTH_URL ||
          (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
        ).replace(/\/+$/, "");

        const reviewRes = await fetch(`${baseUrl}/api/vision-review-pptx`, {
          method: "POST",
          body: reviewForm,
        });

        if (reviewRes.ok) {
          const review = await reviewRes.json() as { deckScore: number; fixes: Array<{ slideIndex: number; field: string; value: string }> };
          console.log(`[gen-pptx] Vision review: deckScore=${review.deckScore} fixes=${review.fixes.length}`);

          if (review.fixes.length > 0) {
            const SLIDE_FIELD_ALLOWLIST = new Set(["density", "textTreatment", "layoutType", "subtitle"]);

            // deleteSlide 対象インデックスを先に収集
            const deleteIndices = new Set<number>(
              (review.fixes as any[])
                .filter((f) => f.field === "deleteSlide" && String(f.value) === "true")
                .map((f) => f.slideIndex as number)
            );
            if (deleteIndices.size > 0) {
              console.log(`[gen-pptx] Vision deleteSlide: indices=[${Array.from(deleteIndices).join(",")}]`);
            }

            // 装飾ありレイアウト: bullets への直落ちは装飾保全ガードで禁止（対応 C+D）
            const DECORATED_LAYOUTS = new Set([
              "icon_rows", "card_grid", "stat_callouts", "metric-cards", "metric_cards",
              "company-overview", "process-cards", "timeline", "roadmap",
            ]);

            const patchedSlidesRaw = sanitizedSlides
              .map((slide, idx) => {
                if (deleteIndices.has(idx)) {
                  console.log(`[gen-pptx] Vision deleted slide[${idx}]: "${slide.title}"`);
                  return null; // 後でfilter
                }
                const slideFixes = (review.fixes as any[]).filter((f) => f.slideIndex === idx);
                if (slideFixes.length === 0) return slide;
                const patch: Record<string, unknown> = {};
                let updatedMetrics = slide.metrics ? [...slide.metrics] : undefined;

                for (const fix of slideFixes) {
                  if (fix.field === "metrics.colorRole" && typeof fix.itemIndex === "number" && updatedMetrics) {
                    if (fix.itemIndex < updatedMetrics.length) {
                      updatedMetrics = updatedMetrics.map((m: any, mi: number) =>
                        mi === fix.itemIndex ? { ...m, colorRole: fix.value } : m
                      );
                      console.log(`[gen-pptx] patch slide[${idx}].metrics[${fix.itemIndex}].colorRole = ${fix.value}`);
                    }
                  } else if (fix.field === "bullets" && typeof fix.value === "string") {
                    // パイプ区切りで bullets を上書き
                    const newBullets = fix.value.split("|").map((b: string) => b.trim()).filter(Boolean);
                    if (newBullets.length > 0) {
                      patch.bullets = newBullets;
                      console.log(`[gen-pptx] patch slide[${idx}].bullets (${newBullets.length}件)`);
                    }
                  } else if (fix.field === "steps" && typeof fix.value === "string") {
                    // "タイトル1:説明1|タイトル2:説明2" 形式でstepsを上書き
                    const newSteps = fix.value.split("|")
                      .map((s: string) => {
                        const colonIdx = s.indexOf(":");
                        return colonIdx > 0
                          ? { title: s.slice(0, colonIdx).trim(), body: s.slice(colonIdx + 1).trim() }
                          : { title: s.trim(), body: "" };
                      })
                      .filter((s: { title: string; body: string }) => s.title);
                    if (newSteps.length > 0) {
                      patch.steps = newSteps;
                      console.log(`[gen-pptx] patch slide[${idx}].steps (${newSteps.length}件)`);
                    }
                  } else if (fix.field === "layoutType") {
                    const newLt = fix.value as string;
                    const currentLt = (slide.layoutType as string) ?? "bullets";
                    // 装飾ありレイアウトから bullets への直落ちを阻止（装飾保全ガード D）
                    if (DECORATED_LAYOUTS.has(currentLt) && newLt === "bullets") {
                      console.log(`[gen-pptx] layoutType fix BLOCKED slide[${idx}]: ${currentLt} → bullets (decoration guard)`);
                    } else {
                      patch.layoutType = fix.value;
                      console.log(`[gen-pptx] patch slide[${idx}].layoutType = ${fix.value}`);
                      // layoutType 変更時: 必要なデータが欠如していれば bullets から自動生成
                      const existingBullets = (slide.bullets ?? []).filter((b: string) => b?.trim());
                      if (fix.value === "multi-column" && !hasUsableColumns(slide.columns)) {
                        if (existingBullets.length >= 2) {
                          const half = Math.ceil(existingBullets.length / 2);
                          patch.columns = [
                            { header: `${slide.title.slice(0, 8)}①`, bullets: existingBullets.slice(0, half) },
                            { header: `${slide.title.slice(0, 8)}②`, bullets: existingBullets.slice(half) },
                          ];
                          console.log(`[gen-pptx] auto-generated columns from bullets for slide[${idx}]`);
                        }
                      }
                      if (fix.value === "table" && !hasUsableTableRows(slide.tableRows)) {
                        if (existingBullets.length >= 2) {
                          const hasColon = existingBullets.some((b: string) => b.includes("：") || b.includes(":"));
                          patch.tableRows = hasColon
                            ? [["項目", "内容"], ...existingBullets.map((b: string) => {
                                const ci = b.indexOf("：") >= 0 ? b.indexOf("：") : b.indexOf(":");
                                return ci > 0 ? [b.slice(0, ci).trim(), b.slice(ci + 1).trim()] : [b, ""];
                              })]
                            : [["No.", slide.title], ...existingBullets.map((b: string, i: number) => [String(i + 1), b])];
                          console.log(`[gen-pptx] auto-generated tableRows from bullets for slide[${idx}]`);
                        }
                      }
                      if ((fix.value === "icon_rows" || fix.value === "card_grid") &&
                          !(slide.cards ?? []).some((c: { heading?: string }) => c.heading?.trim()) &&
                          !(slide.steps ?? []).some((s: { title?: string }) => s.title?.trim())) {
                        if (existingBullets.length >= 2) {
                          const ICON_CYCLE = ["gear", "lightbulb", "rocket", "chart", "star", "verified"] as const;
                          // 5件以上でも bullets に戻さず先頭4件でカード化（装飾保全: bullets revert 禁止）
                          patch.cards = existingBullets.slice(0, 4).map((b: string, bi: number) => {
                            const iconKey = ICON_CYCLE[bi % ICON_CYCLE.length];
                            const { heading, body } = splitBulletForCard(b);
                            return { iconKey, heading, body };
                          });
                          console.log(`[gen-pptx] auto-generated cards from bullets for slide[${idx}] (${fix.value}, capped at 4)`);
                        }
                      }
                      if (fix.value === "stat_callouts" &&
                          !(slide.statCallouts ?? []).some((c: { value?: string }) => c.value?.trim()) &&
                          !(slide.metrics ?? []).some((m: { label?: string }) => m.label?.trim())) {
                        // Vision の reason から "value|unit|label" トリプレットを抽出
                        const reasonStr = (fix as { reason?: string }).reason ?? "";
                        const tripletRe = /([^\s|,，]+)\|([^|,]*)\|([^|,\n，]+)/g;
                        const reasonCallouts: Array<{ value: string; unit: string; label: string }> = [];
                        let rm;
                        while ((rm = tripletRe.exec(reasonStr)) !== null && reasonCallouts.length < 3) {
                          const val = rm[1].trim(), unit = rm[2].trim(), label = rm[3].trim().replace(/[|｜]$/, "");
                          if (val && label) reasonCallouts.push({ value: val, unit, label });
                        }
                        if (reasonCallouts.length >= 1) {
                          patch.statCallouts = reasonCallouts;
                          console.log(`[gen-pptx] auto-generated statCallouts from Vision reason for slide[${idx}] (${reasonCallouts.length}件)`);
                        } else if (existingBullets.length >= 1) {
                          // fallback: bullets 内の数値パターンを抽出
                          const numRe = /([\d,，.]+(?:[.．]\d+)?)\s*(万円|百万|千円|億円|%|％|人|件|回|個|点|倍|ms|GB|MB|KB)?/;
                          const autoCallouts = existingBullets.slice(0, 3).flatMap((b: string) => {
                            const nm = b.match(numRe);
                            if (!nm) return [];
                            const label = b.replace(nm[0], "").replace(/[：:=＝\-－]/g, "").trim().slice(0, 14);
                            return label ? [{ value: nm[1], unit: nm[2] ?? "", label }] : [];
                          });
                          if (autoCallouts.length >= 1) {
                            patch.statCallouts = autoCallouts;
                            console.log(`[gen-pptx] auto-generated statCallouts from bullets for slide[${idx}] (${autoCallouts.length}件)`);
                          }
                        }
                      }
                    }
                  } else if (fix.field === "fontScaleDown") {
                    // 実フォント縮小: bullets切り詰めではなく fontScale を設定（build 関数が反映）
                    const ratio = parseFloat(fix.value);
                    if (!isNaN(ratio) && ratio > 0 && ratio < 1) {
                      patch.fontScale = Math.max(0.70, ratio);
                      console.log(`[gen-pptx] fontScaleDown slide[${idx}] fontScale=${patch.fontScale}`);
                    }
                  } else if (fix.field === "trimText" && typeof fix.value === "string" && fix.value.trim().length > 0) {
                    // 指定テキストに置き換え（bullets[0] を短縮版に差し替え）
                    const newBullets = fix.value.split("|").map((b: string) => b.trim()).filter(Boolean);
                    if (newBullets.length > 0) {
                      patch.bullets = newBullets;
                      console.log(`[gen-pptx] trimText slide[${idx}] bullets replaced (${newBullets.length}件)`);
                    }
                  } else if (fix.field === "fallbackLayout") {
                    // 装飾ありレイアウトから bullets への直落ちを阻止（装飾保全ガード C）
                    const requestedLt = (fix.value || "card_grid").trim();
                    const currentLtFb = (slide.layoutType as string) ?? "bullets";
                    if (DECORATED_LAYOUTS.has(currentLtFb) && requestedLt === "bullets") {
                      patch.layoutType = "card_grid";
                      console.log(`[gen-pptx] fallbackLayout GUARDED slide[${idx}]: ${currentLtFb} → card_grid (bullets blocked from decorated layout)`);
                    } else {
                      patch.layoutType = requestedLt;
                      console.log(`[gen-pptx] fallbackLayout slide[${idx}] → ${requestedLt}`);
                    }
                  } else if (fix.field === "fitTextToShape") {
                    // テキストボックス内 fit フラグ: build 関数でテキスト shrink を強制
                    patch.fitTextToShape = true;
                    console.log(`[gen-pptx] fitTextToShape slide[${idx}] → marked`);
                  } else if (["syncItemDecorations", "copyItemDecoration", "alignItemGroup"].includes(fix.field)) {
                    // itemGroup系は python-pptx 直接操作が必要なためログのみ
                    console.log(`[gen-pptx] Vision fix ${fix.field} slide[${idx}] — logged (requires post-gen python apply)`);
                  } else if (SLIDE_FIELD_ALLOWLIST.has(fix.field)) {
                    patch[fix.field] = fix.value;
                    console.log(`[gen-pptx] patch slide[${idx}].${fix.field} = ${fix.value}`);
                  }
                }
                return { ...slide, ...patch, ...(updatedMetrics ? { metrics: updatedMetrics } : {}) };
              })
              .filter((s): s is PptxSlide => s !== null);
            const patchedSlides = validateAndRepairSlides(normalizeSlidesForPptx(patchedSlidesRaw));

            // coverSubtitle パッチ
            const coverFix = review.fixes.find((f: any) => f.slideIndex === -1 && f.field === "coverSubtitle");
            if (coverFix) designBrief.coverSubtitle = String(coverFix.value);

            // regenerateStyle: Vision LLM がスタイル全面差し替えを要求した場合
            const styleFix = review.fixes.find((f: any) => f.slideIndex === -1 && f.field === "regenerateStyle");
            if (styleFix) {
              console.log(`[gen-pptx] Vision regenerateStyle: "${styleFix.value}" → re-running full design brief`);
              try {
                const newBrief = await generateDesignBrief(title, patchedSlides, instructionText, deckPreferences, String(styleFix.value), promptIntent);
                // palette / styleSpec / mood / visualHints / coverSubtitle を全て差し替え
                // named palette 指定時はそれを維持、未指定時は strictKey から再構築
                designBrief.palette = namedPalette
                  ? buildPaletteFromName(namedPalette)
                  : buildStrictPalette(strictKey);
                designBrief.styleSpec     = newBrief.styleSpec;
                designBrief.mood          = newBrief.mood;
                designBrief.visualHints   = newBrief.visualHints;
                if (newBrief.coverSubtitle) designBrief.coverSubtitle = newBrief.coverSubtitle;
                // theme を styleSpec ベースで全面再構築
                Object.assign(theme, resolveTheme(designBrief, instructionText, deckPreferences, fontFace));
              } catch (e) {
                console.warn("[gen-pptx] regenerateStyle failed:", e);
              }
            }

            // 再生成
            const pptx2 = new PptxGenJS();
            pptx2.layout = "LAYOUT_WIDE";
            pptx2.theme  = { headFontFace: "Meiryo", bodyFontFace: "Meiryo" };
            pptx2.author = "azurechat";
            pptx2.subject = title;
            pptx2.title = title;
            if (!faithfulMode) {
              buildTitleSlide(pptx2, title, designBrief, theme, patchedSlides.length + 1, coverIllustration);
            }
            let regeneratedIllustrationPlaced = false;
            patchedSlides.forEach((slideRaw, index) => {
              const slide = upgradeTextOnlySlide(slideRaw);
              const resolvedLt = resolveLayoutType(slide);
              const visual = designBrief.visualHints[index] ?? {
                title: slide.title,
                visualType:
                  resolvedLt === "table"         ? "table" :
                  resolvedLt === "multi-column"  ? "comparison" :
                  resolvedLt === "diagram" || resolvedLt === "process-cards" ? "process" :
                  resolvedLt === "timeline"      ? "timeline" :
                  resolvedLt === "company-overview" || resolvedLt === "metric-cards" ? "cards" :
                  resolvedLt === "closing"       ? "spotlight" : "editorial",
              };
              const slideIllustration =
                !regeneratedIllustrationPlaced &&
                coverIllustration &&
                resolvedLt === "bullets"
                  ? coverIllustration
                  : null;
              switch (resolvedLt) {
                case "title":          buildSectionSlide(pptx2, slide.title, theme); break;
                case "table":          buildTableSlide(pptx2, slide, theme, visual, faithfulMode); break;
                case "multi-column":   buildMultiColumnSlide(pptx2, slide, theme, visual, faithfulMode); break;
                case "diagram":        buildDiagramSlide(pptx2, slide, theme, visual, faithfulMode); break;
                case "conversation":   buildConversationSlide(pptx2, slide, theme); break;
                case "company-overview": buildCompanyOverviewSlide(pptx2, slide, theme); break;
                case "process-cards":  buildProcessCardsSlide(pptx2, slide, theme); break;
                case "closing":        buildClosingSlide(pptx2, slide, theme); break;
                case "metric-cards":   buildMetricCardsSlide(pptx2, slide, theme); break;
                case "timeline":       buildTimelineSlide(pptx2, slide, theme); break;
                case "stat_callouts":  buildStatCalloutsSlide(pptx2, slide, theme); break;
                case "card_grid":      buildCardGridSlide(pptx2, slide, theme); break;
                case "icon_rows":      buildIconRowsSlide(pptx2, slide, theme); break;
                case "roadmap":        buildTimelineSlide(pptx2, slide, theme); break;
                default:
                  buildBulletsSlide(pptx2, slide, theme, visual, slideIllustration, faithfulMode);
                  if (slideIllustration) regeneratedIllustrationPlaced = true;
                  break;
              }
            });
            buffer = await patchEastAsianFont((await pptx2.write({ outputType: "nodebuffer" })) as Buffer);
            {
              const isZip2 = buffer && buffer.length >= 100 && buffer[0] === 0x50 && buffer[1] === 0x4B;
              const hasSlides2 = isZip2 && buffer.includes(Buffer.from("ppt/slides/slide"));
              if (!isZip2 || !hasSlides2) {
                console.warn("[gen-pptx] Re-generated PPTX invalid — keeping original buffer");
                buffer = await patchEastAsianFont((await pptx.write({ outputType: "nodebuffer" })) as Buffer);
              }
            }
            if (buffer && buffer.length >= 100 && buffer[0] === 0x50 && buffer[1] === 0x4B) {
              console.log(`[gen-pptx] Re-generated after Vision review`);
            }
          }
        }
      } catch (visionErr) {
        console.warn("[gen-pptx] Vision review failed (non-fatal):", visionErr);
      }
    }
    // ─────────────────────────────────────────────────────────────────────

    const shortId = uniqueId().slice(0, 8);
    const safeBase = fileBaseName
      ? fileBaseName.replace(/\.pptx$/i, "").replace(/[\\/:*?"<>|]/g, "").trim().slice(0, 40)
      : (threadId ?? uniqueId());
    const displayFileName = `${safeBase}.pptx`;          // 日本語名：リンク表示・DL名
    const blobKey = `pptx_${shortId}.pptx`;              // ASCII のみ：Blob key（URL短縮）
    const downloadUrl = await uploadPptxToBlob(buffer, blobKey, displayFileName);
    return NextResponse.json({ ok: true, downloadUrl, fileName: displayFileName });
  } catch (e: any) {
    console.error("[gen-pptx] error:", e);
    return NextResponse.json({ ok: false, error: String(e?.message ?? e) }, { status: 500 });
  }
}
