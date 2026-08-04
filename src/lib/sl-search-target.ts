export type SlSearchScope =
  | "all"
  | "personal"
  | "dept_common"
  | "global_common";

export type SlSearchTarget = {
  scope?: SlSearchScope;
  folder?: string;
  folderUncertain?: boolean;
};

/** チャット本文に検索先が明記されている場合だけ検索条件へ変換する。 */
export function inferSlSearchTarget(text: string): SlSearchTarget {
  const source = (text ?? "").trim();
  const target: SlSearchTarget = {};

  if (/(?:全社共通|全社共有|global[_\s-]?common|\bcommon\b)/i.test(source)) {
    target.scope = "global_common";
  } else if (/(?:部署共通|部門共通|部署共有|部門共有)/.test(source)) {
    target.scope = "dept_common";
  } else if (
    /(?:個人(?:ファイル|資料|フォルダ(?:ー)?)|自分の(?:ファイル|資料|フォルダ(?:ー)?)|私の(?:ファイル|資料|フォルダ(?:ー)?))/.test(
      source
    )
  ) {
    target.scope = "personal";
  }

  const describedFolder = source.match(
    /[「『"]([^」』"]+?)[」』"]\s*(?:という|と呼ばれる)\s*フォルダ(?:ー)?/
  );
  const authoredPlainFolder = source.match(
    /(?:ユーザー|利用者|担当者|社員|私|自分)(?:が|の)?\s*(?:作成|用意|保存)した\s*([^\s、。「」]{1,40}?)\s*(?:という|と呼ばれる)\s*フォルダ(?:ー)?/
  );
  const describedPlainFolder = source.match(
    /(?:^|[\s、。])([^\s、。「」]{1,40}?)\s*(?:という|と呼ばれる)\s*フォルダ(?:ー)?/
  );
  const quotedFolder = source.match(
    /[「『"]([^」』"]+?)[」』"]フォルダ(?:ー)?(?:の中)?から/
  );
  const quotedFolderWithSuffix = source.match(
    /[「『"]([^」』"]+?フォルダ(?:ー)?)[」』"](?:の中)?から/
  );
  const plainFolder = source.match(
    /(?:^|[\s、。])([^\s、。]{1,80}?)(?:フォルダ(?:ー)?)(?:の中)?から/
  );
  const plainFolderCandidate = (plainFolder?.[1] ?? "").trim();
  const safePlainFolder =
    plainFolderCandidate &&
    !/(?:部署共通|部門共通|全社共通|個人|の中|にある|作成した|用意した|保存した)/.test(
      plainFolderCandidate
    )
      ? plainFolderCandidate
      : "";
  const folder = (
    describedFolder?.[1] ??
    authoredPlainFolder?.[1] ??
    describedPlainFolder?.[1] ??
    quotedFolder?.[1] ??
    quotedFolderWithSuffix?.[1]?.replace(/フォルダ(?:ー)?$/, "") ??
    safePlainFolder ??
    ""
  ).trim();
  const genericFolderNames = new Set([
    "個人",
    "自分の",
    "私の",
    "全社共通",
    "全社共有",
    "部署共通",
    "部門共通",
    "部署共有",
    "部門共有",
    "common",
  ]);
  if (folder && !genericFolderNames.has(folder.toLowerCase())) {
    target.folder = folder;
  }

  const withoutGenericScopeFolders = source.replace(
    /(?:個人|自分の|私の|全社共通|全社共有|部署共通|部門共通|部署共有|部門共有)(?:資料|ファイル|フォルダ(?:ー)?)/g,
    " "
  );
  if (!target.folder && /フォルダ(?:ー)?/.test(withoutGenericScopeFolders)) {
    target.folderUncertain = true;
  }

  return target;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** 検索対象を示す言葉を除き、文書名・検索内容だけをAzure Searchへ渡す。 */
export function stripSlSearchTargetTerms(
  query: string,
  target: SlSearchTarget = {}
): string {
  const original = (query ?? "").trim();
  let cleaned = original
    .replace(
      /(?:個人(?:ファイル|資料|フォルダ(?:ー)?)|自分の(?:ファイル|資料|フォルダ(?:ー)?)|私の(?:ファイル|資料|フォルダ(?:ー)?))(?:の中)?(?:にある|から)?/g,
      " "
    )
    .replace(
      /(?:全社共通|全社共有|部署共通|部門共通|部署共有|部門共有)(?:資料|ファイル|フォルダ(?:ー)?)?(?:の中)?(?:にある|から)?/g,
      " "
    );

  const folder = (target.folder ?? "").trim();
  if (folder) {
    cleaned = cleaned.replace(
      new RegExp(
        `[「『"]?${escapeRegExp(folder)}[」』"]?\\s*(?:という|と呼ばれる)?\\s*フォルダ(?:ー)?(?:の中)?(?:にある|から)?`,
        "g"
      ),
      " "
    );
    cleaned = cleaned.replace(
      /(?:ユーザー|利用者|担当者|社員|私|自分)(?:が|の)?\s*(?:作成|用意|保存)した/g,
      " "
    );
  }

  cleaned = cleaned.replace(/[、,]\s*/g, " ").replace(/\s+/g, " ").trim();
  return cleaned || original;
}

function normalizeFolderQuery(value?: string): string {
  return (value ?? "")
    .trim()
    .replace(/[+|!(){}\[\]^"~*?:\\/\-]+/g, " ")
    .replace(/\s+/g, " ")
    .replace(/'/g, "''");
}

/**
 * ユーザーが明示したSharePointの検索先をAzure AI Searchのfilterへ変換する。
 * ACLは呼び出し先で別途AND結合されるため、ここでは検索範囲だけを扱う。
 */
export function buildSlSearchTargetFilter(params: {
  scope?: SlSearchScope;
  folder?: string;
}): string | undefined {
  const filters: string[] = [];

  if (params.scope && params.scope !== "all") {
    filters.push(`isSlDoc eq true and slScope eq '${params.scope}'`);
  }

  const folderQuery = normalizeFolderQuery(params.folder);
  if (folderQuery) {
    filters.push(
      `isSlDoc eq true and search.ismatch('${folderQuery}', 'relativePath', 'simple', 'all')`
    );
  }

  if (filters.length === 0) return undefined;
  return filters.map((filter) => `(${filter})`).join(" and ");
}
