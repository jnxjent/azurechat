"use server";
import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";
import {
  AzureSearchDocumentIndex,
  DocumentSearchResponse,
  SearchAllSharePointDocumentChunks,
  SimpleSearch,
} from "./azure-ai-search/azure-ai-search";

const DEFAULT_MAX_PAGES = 300;
const MAP_MAX_CHARS = 18_000;
const MAP_MAX_CHUNKS = 10;
const MAP_CONCURRENCY = 4;

type SummaryBatch = {
  text: string;
  pageStart: number;
  pageEnd: number;
};

export type SharePointPdfSummaryResult = {
  fileName: string;
  fileUrl: string;
  pageCount: number;
  chunkCount: number;
  summary: string;
  citationSource: DocumentSearchResponse;
};

function normalizedFileName(value: string): string {
  return value
    .trim()
    .normalize("NFKC")
    .replace(/^['"「『]|['"」』]$/g, "")
    .replace(/\s+/g, " ")
    .toLocaleLowerCase("ja-JP");
}

function fileNameFromDocument(document: AzureSearchDocumentIndex): string {
  const metadata = String(document.metadata ?? "").trim();
  if (metadata && !metadata.startsWith("{")) return metadata;
  const rawUrl = document.effectiveFileUrl || document.fileUrl || "";
  try {
    return decodeURIComponent(rawUrl.split(/[?#]/)[0].split("/").pop() ?? "");
  } catch {
    return rawUrl.split(/[?#]/)[0].split("/").pop() ?? "";
  }
}

function selectPdfCandidate(
  results: DocumentSearchResponse[],
  fileQuery: string
): DocumentSearchResponse {
  const byItem = new Map<string, DocumentSearchResponse>();
  for (const result of results) {
    const id = result.document.spItemId;
    if (id && !byItem.has(id)) byItem.set(id, result);
  }
  const candidates = Array.from(byItem.values()).filter((result) =>
    fileNameFromDocument(result.document).toLowerCase().endsWith(".pdf")
  );
  const query = normalizedFileName(fileQuery);
  const exact = candidates.filter(
    (result) => normalizedFileName(fileNameFromDocument(result.document)) === query
  );
  const queryWithoutExtension = query.replace(/\.pdf$/i, "");
  const stemExact = candidates.filter(
    (result) =>
      normalizedFileName(fileNameFromDocument(result.document)).replace(/\.pdf$/i, "") ===
      queryWithoutExtension
  );
  const contains = candidates.filter((result) =>
    normalizedFileName(fileNameFromDocument(result.document)).includes(queryWithoutExtension)
  );
  const matched = exact.length > 0 ? exact : stemExact.length > 0 ? stemExact : contains;

  if (matched.length === 0) {
    throw new Error(`対象PDF「${fileQuery}」が検索可能なSharePoint文書に見つかりません。`);
  }
  if (matched.length > 1) {
    const names = matched.slice(0, 8).map((item) => fileNameFromDocument(item.document));
    throw new Error(
      `対象PDFを一意に特定できません。ファイル名またはフォルダーを指定してください: ${names.join("、")}`
    );
  }
  return matched[0];
}

function validatePageAwareChunks(
  results: DocumentSearchResponse[],
  maxPages: number
): { documents: AzureSearchDocumentIndex[]; pageCount: number; chunkCount: number } {
  if (results.length === 0) {
    throw new Error(
      "対象PDFにはページ情報付きチャンクがありません。SharePoint同期で再インデックスしてください。"
    );
  }
  const documents = results.map((result) => result.document);
  const expectedChunks = documents[0].documentChunkCount;
  const pageCount = documents[0].documentPageCount;
  if (!Number.isInteger(expectedChunks) || !Number.isInteger(pageCount)) {
    throw new Error("対象PDFのページ情報が不完全です。再インデックスしてください。");
  }
  if ((pageCount as number) > maxPages) {
    throw new Error(
      `この同期PoCの上限は${maxPages}ページです（対象PDF: ${pageCount}ページ）。`
    );
  }
  if (documents.length !== expectedChunks) {
    throw new Error(
      `チャンクが不足しています（取得${documents.length}件／予定${expectedChunks}件）。要約を中止しました。`
    );
  }
  for (let index = 0; index < documents.length; index++) {
    const document = documents[index];
    if (
      document.chunkIndex !== index ||
      document.documentChunkCount !== expectedChunks ||
      document.documentPageCount !== pageCount ||
      !Number.isInteger(document.pageStart) ||
      !Number.isInteger(document.pageEnd) ||
      (document.pageStart as number) < 1 ||
      (document.pageEnd as number) > (pageCount as number)
    ) {
      throw new Error(
        `チャンク順序またはページ情報が不整合です（chunkIndex=${document.chunkIndex ?? "null"}）。要約を中止しました。`
      );
    }
  }
  return {
    documents,
    pageCount: pageCount as number,
    chunkCount: expectedChunks as number,
  };
}

function createSummaryBatches(documents: AzureSearchDocumentIndex[]): SummaryBatch[] {
  const batches: SummaryBatch[] = [];
  let parts: string[] = [];
  let chars = 0;
  let start = 0;
  let end = 0;

  const flush = () => {
    if (parts.length === 0) return;
    batches.push({ text: parts.join("\n\n"), pageStart: start, pageEnd: end });
    parts = [];
    chars = 0;
    start = 0;
    end = 0;
  };

  for (const document of documents) {
    const pageStart = document.pageStart as number;
    const pageEnd = document.pageEnd as number;
    const part = `[原文 p.${pageStart}${pageEnd === pageStart ? "" : `-${pageEnd}`} ]\n${document.pageContent}`;
    if (parts.length > 0 && (parts.length >= MAP_MAX_CHUNKS || chars + part.length > MAP_MAX_CHARS)) {
      flush();
    }
    if (parts.length === 0) start = pageStart;
    end = pageEnd;
    parts.push(part);
    chars += part.length;
  }
  flush();
  return batches;
}

async function withRetry<T>(work: () => Promise<T>, attempts = 3): Promise<T> {
  let lastError: unknown;
  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      return await work();
    } catch (error) {
      lastError = error;
      if (attempt < attempts) {
        await new Promise((resolve) => setTimeout(resolve, 500 * 2 ** (attempt - 1)));
      }
    }
  }
  throw lastError;
}

async function mapWithConcurrency<T, R>(
  values: T[],
  concurrency: number,
  mapper: (value: T, index: number) => Promise<R>
): Promise<R[]> {
  const output = new Array<R>(values.length);
  let cursor = 0;
  const workers = Array.from({ length: Math.min(concurrency, values.length) }, async () => {
    while (cursor < values.length) {
      const index = cursor++;
      output[index] = await mapper(values[index], index);
    }
  });
  await Promise.all(workers);
  return output;
}

function responseText(content: string | null): string {
  const text = (content ?? "").trim();
  if (!text) throw new Error("要約モデルから空の応答が返りました。");
  return text;
}

function removeInvalidPageReferences(text: string, pageCount: number): string {
  return text.replace(/\[p\.(\d+)(?:\s*[-–]\s*(\d+))?\]/gi, (whole, startText, endText) => {
    const start = Number(startText);
    const end = Number(endText ?? startText);
    return start >= 1 && end >= start && end <= pageCount ? whole : "";
  });
}

function citedPages(text: string, pageCount: number): Set<number> {
  const pages = new Set<number>();
  const pattern = /\[p\.(\d+)(?:\s*[-–]\s*(\d+))?\]/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text)) !== null) {
    const start = Number(match[1]);
    const end = Number(match[2] ?? match[1]);
    if (start < 1 || end < start || end > pageCount) continue;
    for (let page = start; page <= end; page++) pages.add(page);
  }
  return pages;
}


export async function summarizeSharePointPdf(props: {
  fileQuery: string;
  filter?: string;
  deptLower?: string | null;
  userHash?: string;
  signal?: AbortSignal;
  targetPages?: number;
  targetCharsLow?: number;
  targetCharsHigh?: number;
}): Promise<SharePointPdfSummaryResult> {
  const summaryStartedAt = Date.now();
  const maxPages = Math.max(
    1,
    Number.parseInt(process.env.SL_FULL_SUMMARY_MAX_PAGES ?? "", 10) || DEFAULT_MAX_PAGES
  );
  // Azure Search and the index metadata may distinguish full-width digits
  // from ASCII digits (for example 令和７年度 vs 令和7年度).
  const normalizedQuery = props.fileQuery.normalize("NFKC").trim();
  const discoveryQuery = normalizedQuery.replace(/\.pdf$/i, "").trim();
  const discovery = await SimpleSearch(
    discoveryQuery || normalizedQuery,
    props.filter ? `isSlDoc eq true and (${props.filter})` : "isSlDoc eq true",
    props.deptLower,
    200,
    props.userHash
  );
  if (discovery.status !== "OK") {
    throw new Error(`PDF検索に失敗しました: ${discovery.errors[0]?.message ?? "unknown"}`);
  }
  const candidate = selectPdfCandidate(discovery.response, normalizedQuery);
  const spItemId = candidate.document.spItemId!;
  const allChunks = await SearchAllSharePointDocumentChunks({
    spItemId,
    filter: props.filter,
    deptLower: props.deptLower,
    userHash: props.userHash,
  });
  if (allChunks.status !== "OK") {
    throw new Error(`全文チャンク取得に失敗しました: ${allChunks.errors[0]?.message ?? "unknown"}`);
  }
  const validated = validatePageAwareChunks(allChunks.response, maxPages);
  const batches = createSummaryBatches(validated.documents);
  const targetPages = Number.isFinite(props.targetPages)
    ? Math.max(3, Math.min(20, Math.round(props.targetPages as number)))
    : 10;
  const requestedCharsLow = Number.isFinite(props.targetCharsLow)
    ? Math.max(1_000, Math.round(props.targetCharsLow as number))
    : null;
  const requestedCharsHigh = Number.isFinite(props.targetCharsHigh)
    ? Math.max(1_000, Math.round(props.targetCharsHigh as number))
    : null;
  const hasValidRequestedRange =
    requestedCharsLow !== null &&
    requestedCharsHigh !== null &&
    requestedCharsLow <= requestedCharsHigh;
  const targetCharsLow = hasValidRequestedRange
    ? requestedCharsLow
    : Math.round(targetPages * 750 * 0.85);
  const targetCharsHigh = hasValidRequestedRange
    ? requestedCharsHigh
    : Math.round(targetPages * 750 * 1.15);
  console.log(
    `[SP PDF summary] file=${fileNameFromDocument(candidate.document)} pages=${validated.pageCount} chunks=${validated.chunkCount} mapBatches=${batches.length}`
  );
  const openai = OpenAIInstance();
  const mapStartedAt = Date.now();
  const partials = await mapWithConcurrency(batches, MAP_CONCURRENCY, async (batch) =>
    withRetry(async () => {
      const completion = await openai.chat.completions.create(
          {
            model: "",
            messages: [
            {
              role: "system",
              content:
                `あなたは文書要約者です。与えられた原文だけを要約し、重要事項・数値・決定・例外を落とさないでください。各箇条書きの末尾に必ず原文のページを [p.12] または [p.12-14] の形式で付けてください。入力範囲外のページや推測は書かないでください。空ページは無視してください。最大12項目の日本語Markdownで返してください。`,
            },
            { role: "user", content: batch.text },
          ],
        },
        { signal: props.signal }
      );
      console.log(
        `[SP PDF summary] map completed pages=${batch.pageStart}-${batch.pageEnd}`
      );
      return `## p.${batch.pageStart}-${batch.pageEnd}\n${responseText(completion.choices[0]?.message?.content ?? null)}`;
    })
  );
  console.log(`[SP PDF summary] map elapsedMs=${Date.now() - mapStartedAt}`);

  const reduceInput = partials.join("\n\n");
  const finalPageText = validated.documents
    .filter((document) => document.pageStart === validated.pageCount)
    .map((document) => document.pageContent)
    .filter((content) => content.trim() && content.trim() !== "[PAGE_EMPTY]")
    .join("\n\n")
    .trim();
  const reduceStartedAt = Date.now();
  const [finalCompletion, finalPageSummary] = await Promise.all([
    withRetry(() =>
      openai.chat.completions.create(
        {
          model: "",
          max_completion_tokens: 16000,
          messages: [
            {
              role: "system",
              content:
                `あなたは長文PDFの最終要約者です。部分要約だけを根拠に、重複を統合して日本語Markdownの全体要約を作ってください。構成は「概要」「主要ポイント」「重要な数値・決定事項」「章・論点別要約」「結論・留意点」とします。入力の先頭から末尾まで全ページ範囲を均等にカバーし、後半ページ・終盤の内容も必ず含めてください。すべての実質的な箇条書きに、入力に存在する [p.N] または [p.N-M] のページ参照を最低1つ残してください。ページ番号を推測・変更しないでください。Wordで約${targetPages}ページになる情報量を想定し、約${targetCharsLow}〜${targetCharsHigh}文字にしてください。${targetCharsLow}文字未満で終了してはいけません。最終出力前に情報量を確認し、不足する場合は根拠のある背景・数値・決定事項・章別論点を補ってください。水増しや同じ内容の繰り返しは禁止です。`,
            },
            { role: "user", content: reduceInput },
          ],
        },
        { signal: props.signal }
      )
    ),
    !finalPageText
      ? Promise.resolve(
          `PDFの最終ページは確認済みですが、抽出可能な本文はありませんでした。[p.${validated.pageCount}]`
        )
      : withRetry(() =>
          openai.chat.completions.create(
            {
              model: "",
              messages: [
                {
                  role: "system",
                  content:
                    "与えられたPDF最終ページだけを根拠に、記載内容を1〜3項目の簡潔な日本語で要約してください。推測は禁止です。ページ参照は出力しないでください。",
                },
                { role: "user", content: finalPageText },
              ],
            },
            { signal: props.signal }
          )
        ).then(
          (completion) =>
            `${responseText(completion.choices[0]?.message?.content ?? null)
              .replace(/\[p\.\d+(?:\s*[-–]\s*\d+)?\]/gi, "")
              .trim()} [p.${validated.pageCount}]`
        ),
  ]);
  console.log(`[SP PDF summary] reduce+finalPage elapsedMs=${Date.now() - reduceStartedAt}`);
  let summary = removeInvalidPageReferences(
    responseText(finalCompletion.choices[0]?.message?.content ?? null),
    validated.pageCount
  );
  if (summary.length < targetCharsLow) {
    const refineStartedAt = Date.now();
    console.log(
      `[SP PDF summary] draft below target: chars=${summary.length} target=${targetCharsLow}-${targetCharsHigh}; refining once`
    );
    const refinedCompletion = await withRetry(() =>
      openai.chat.completions.create(
        {
          model: "",
          max_completion_tokens: 16000,
          messages: [
            {
              role: "system",
              content:
                `あなたは長文PDF要約の編集者です。部分要約だけを根拠に、下書きを約${targetCharsLow}〜${targetCharsHigh}文字へ具体化してください。構成と [p.N] / [p.N-M] の参照を維持し、重要な背景・数値・決定事項・章別論点を補ってください。入力にない事実やページ番号の追加、同じ内容の水増しは禁止です。完成した要約本文だけを返してください。`,
            },
            {
              role: "user",
              content: `# 部分要約\n\n${reduceInput}\n\n# 現在の下書き\n\n${summary}`,
            },
          ],
        },
        { signal: props.signal }
      )
    );
    summary = removeInvalidPageReferences(
      responseText(refinedCompletion.choices[0]?.message?.content ?? null),
      validated.pageCount
    );
    console.log(
      `[SP PDF summary] refined chars=${summary.length} elapsedMs=${Date.now() - refineStartedAt}`
    );
  }
  // Always inspect and cite the physical final page explicitly. This proves
  // terminal-page coverage independently of the Reduce model's selection.
  summary += `\n\n# 最終ページの確認\n\n${finalPageSummary}`;
  console.log(
    `[SP PDF summary] final page explicitly summarized page=${validated.pageCount}`
  );

  if (!/\[p\.\d+/i.test(summary)) {
    summary += `\n\n参照範囲: [p.1-${validated.pageCount}]`;
  }
  console.log(
    `[SP PDF summary] reduce completed file=${fileNameFromDocument(candidate.document)} totalElapsedMs=${Date.now() - summaryStartedAt}`
  );

  return {
    fileName: fileNameFromDocument(candidate.document),
    fileUrl: candidate.document.effectiveFileUrl || candidate.document.fileUrl || "",
    pageCount: validated.pageCount,
    chunkCount: validated.chunkCount,
    summary,
    citationSource: candidate,
  };
}
