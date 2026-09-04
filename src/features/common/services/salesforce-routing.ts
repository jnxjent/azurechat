export type SalesforceIntent =
  | "salesforce_data"
  | "knowledge_search"
  | "normal";

export type SalesforceRoute =
  | "salesforce"
  | "knowledge"
  | "normal"
  | "denied";

export interface SalesforceRoutingInput {
  message: string;
  isSalesforceAllowed: boolean;
  hasSalesforceExtension: boolean;
}

export interface SalesforceRoutingResult {
  intent: SalesforceIntent;
  route: SalesforceRoute;
}

const KNOWLEDGE_SOURCE_PATTERN =
  /(?:share\s*point|シェアポイント|社内(?:資料|文書|情報|データ|ナレッジ)|マニュアル|手順書|Q\s*&\s*A|\bQA\b|個人(?:ファイル|フォルダ(?:ー)?)|部署共通|部門共通|全社共通|全社共有)/i;

const SALESFORCE_KNOWLEDGE_PATTERN =
  /(?:使い方|操作方法|利用方法|何ができ|機能|概要|とは|マニュアル|手順|申請|障害|不具合|エラー|トラブル|Q\s*&\s*A|\bQA\b|ヘルプ|社内資料)/i;

const SALESFORCE_DATA_ACTION_PATTERN =
  /(?:調べ|検索|探し|探す|確認|取得|参照|見せ|表示|一覧|抽出|教えて|何件|件数|集計)/i;

const SALESFORCE_RECORD_PATTERN =
  /(?:取引先|顧客|商談|案件|日報|活動報告|売上|担当者|責任者|レコード|データ|情報)/i;

function normalizeMessage(message: string): string {
  return String(message ?? "").normalize("NFKC").trim();
}

/**
 * Matches SF only when it is used as a Salesforce source marker. In
 * particular, Japanese genre words such as "SF映画" and "SF小説" must not
 * enable Salesforce.
 */
function containsSalesforceMarker(message: string): boolean {
  if (/salesforce|セールスフォース/i.test(message)) return true;
  return /(?:^|[\s、。,.!！?？「」『』（）()])SF(?=$|[\s、。,.!！?？「」『』（）()]|によ|で|から|上|内|連携)/i.test(
    message
  );
}

export function classifySalesforceIntent(message: string): SalesforceIntent {
  const normalized = normalizeMessage(message);

  // An explicitly named knowledge source always wins over Salesforce data.
  if (KNOWLEDGE_SOURCE_PATTERN.test(normalized)) {
    return "knowledge_search";
  }

  if (!containsSalesforceMarker(normalized)) {
    return "normal";
  }

  // Salesforce help, operations and incident questions belong to the
  // knowledge route even when SharePoint is not named explicitly.
  if (SALESFORCE_KNOWLEDGE_PATTERN.test(normalized)) {
    return "knowledge_search";
  }

  const citesSalesforceAsSource =
    /(?:salesforce|セールスフォース)(?:によれば)/i.test(normalized) ||
    /(?:^|[\s、。,.!！?？「」『』（）()])SF(?:によれば)/i.test(normalized);

  const explicitlyNamesSalesforceAsSource =
    /(?:salesforce|セールスフォース)(?:で|から|上|内|連携で)/i.test(
      normalized
    ) ||
    /(?:^|[\s、。,.!！?？「」『』（）()])SF(?:で|から|上|内|連携で)/i.test(
      normalized
    );

  const namesARecordUnderSalesforce =
    /(?:salesforce|セールスフォース)(?:上|内)?の/i.test(normalized) &&
    SALESFORCE_RECORD_PATTERN.test(normalized) &&
    SALESFORCE_DATA_ACTION_PATTERN.test(normalized);
  const asksExplicitSourceQuestion =
    explicitlyNamesSalesforceAsSource && /[?？]\s*$/.test(normalized);

  if (
    citesSalesforceAsSource ||
    (explicitlyNamesSalesforceAsSource &&
      SALESFORCE_DATA_ACTION_PATTERN.test(normalized)) ||
    asksExplicitSourceQuestion ||
    namesARecordUnderSalesforce
  ) {
    return "salesforce_data";
  }

  return "normal";
}

export function resolveSalesforceRoute(
  input: SalesforceRoutingInput
): SalesforceRoutingResult {
  const intent = classifySalesforceIntent(input.message);

  if (intent === "knowledge_search") {
    return { intent, route: "knowledge" };
  }

  if (intent === "salesforce_data" && !input.isSalesforceAllowed) {
    return { intent, route: "denied" };
  }

  if (
    intent === "salesforce_data" &&
    input.isSalesforceAllowed &&
    input.hasSalesforceExtension
  ) {
    return { intent, route: "salesforce" };
  }

  return { intent, route: "normal" };
}

const EXPLICIT_SALESFORCE_OBJECT_PATTERN =
  /(?:取引先|顧客|商談|案件|日報|活動報告|売上|担当者|責任者|連絡先|住所|所在地|電話|メール|与信|請求|レコード|Account|Opportunity|Contact)/i;
const COMPANY_NAME_SUFFIX_PATTERN =
  /(?:株式会社|有限会社|合同会社|ホールディングス|産業|工業|建設|商事|興産|製作所|運輸|物流|銀行|信用金庫|病院|大学)$/;

/**
 * The Gateway defaults an otherwise ambiguous company-name lookup to an
 * Opportunity. Add an Account hint only for a bare lookup whose target looks
 * like a company name; never rewrite requests that already name an object or
 * field.
 */
export function buildSalesforceGatewayQuery(message: string): string {
  const normalized = normalizeMessage(message);
  const withoutSourcePrefix = normalized.replace(
    /^(?:salesforce|セールスフォース|SF)(?:によれば|連携で|で|から|上(?:の|で)|内(?:の|で)|の)?[\s、,:：]*/i,
    ""
  );
  if (EXPLICIT_SALESFORCE_OBJECT_PATTERN.test(withoutSourcePrefix)) {
    return withoutSourcePrefix;
  }
  const match = withoutSourcePrefix.match(
    /^(.{2,60}?)を(?:調べて|検索して|探して|確認して)(?:ください|下さい)?[。.!！?？]*$/
  );
  const companyName = match?.[1]?.trim() ?? "";
  if (!companyName || !COMPANY_NAME_SUFFIX_PATTERN.test(companyName)) {
    return withoutSourcePrefix;
  }

  return `取引先の${companyName}について教えて`;
}
