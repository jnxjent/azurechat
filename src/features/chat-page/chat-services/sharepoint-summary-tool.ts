import "server-only";

import { randomUUID } from "crypto";
import { UploadBlob } from "@/features/common/services/azure-storage";
import { buildSlSearchTargetFilter } from "@/lib/sl-search-target";
import type { SlSearchScope } from "@/lib/sl-search-target";
import { RunnableToolFunction } from "openai/lib/RunnableFunction";
import { summarizeSharePointPdf } from "./sharepoint-summary-service";

export function createSharePointPdfSummaryTool(props: {
  deptLower: string;
  userHash?: string;
  signal?: AbortSignal;
  userMessage?: string;
  chatThreadId: string;
}): RunnableToolFunction<any> {
  return {
    type: "function",
    function: {
      name: "summarize_sp_pdf",
      description:
        "SharePointにインデックス済みの単一PDFを、先頭から末尾まで全チャンク取得して要約します。ファイル名を指定したPDFの要約、全文要約、全体要約、全ページ要約を求められた場合に使います。特定事項だけの検索には使いません。Word出力も求められた場合は、このツールの完了後に、返された要約本文とページ参照をcreate_wordへ渡します。対象はページ情報付きで再インデックス済み、かつ既定300ページ以下のPDFです。",
      parameters: {
        type: "object",
        properties: {
          fileQuery: {
            type: "string",
            description: "対象PDFのファイル名。可能なら .pdf を含む完全なファイル名を指定します。",
          },
          scope: {
            type: "string",
            enum: ["all", "personal", "dept_common", "global_common"],
            description: "SharePoint上の検索範囲。指定がなければ閲覧可能な全範囲です。",
          },
          folder: {
            type: "string",
            description: "ユーザーが指定したSharePointフォルダー名。曖昧検索を避けるために使用します。",
          },
          targetPages: {
            type: "number",
            description: "Word出力などで希望する要約の概算ページ数。指定された場合のみ設定します（3〜20ページ）。例: 10",
          },
        },
        required: ["fileQuery"],
      },
      function: async (args: {
        fileQuery: string;
        scope?: SlSearchScope;
        folder?: string;
        targetPages?: number;
      }) => {
        try {
          const targetFilter = buildSlSearchTargetFilter(args);
          const normalizedUserMessage = (props.userMessage ?? "").replace(
            /[０-９]/g,
            (digit) => String(digit.charCodeAt(0) - "０".charCodeAt(0))
          );
          const requestedPagesMatch = normalizedUserMessage.match(
            /(\d{1,2})\s*(?:ページ|pages?|頁)/i
          );
          const inferredTargetPages = requestedPagesMatch
            ? Number.parseInt(requestedPagesMatch[1], 10)
            : undefined;
          const normalizedLengthMessage = normalizedUserMessage.replace(/[,，]/g, "");
          const requestedCharsMatch = normalizedLengthMessage.match(
            /(\d{3,6})\s*(?:[～〜~\-–—]|から)\s*(\d{3,6})\s*(?:文字|字)/i
          );
          const inferredTargetCharsLow = requestedCharsMatch
            ? Number.parseInt(requestedCharsMatch[1], 10)
            : undefined;
          const inferredTargetCharsHigh = requestedCharsMatch
            ? Number.parseInt(requestedCharsMatch[2], 10)
            : undefined;
          const result = await summarizeSharePointPdf({
            fileQuery: args.fileQuery,
            filter: targetFilter,
            deptLower: props.deptLower,
            userHash: props.userHash,
            signal: props.signal,
            targetPages: args.targetPages ?? inferredTargetPages,
            targetCharsLow: inferredTargetCharsLow,
            targetCharsHigh: inferredTargetCharsHigh,
          });
          const wantsWord = /(?:Word|ワード|docx)/i.test(props.userMessage ?? "");
          let summaryRef = "";
          if (wantsWord) {
            summaryRef = `sp-summary-cache/${props.chatThreadId}/${randomUUID()}.json`;
            const cached = await UploadBlob(
              "dl-link",
              summaryRef,
              Buffer.from(
                JSON.stringify({
                  summary: result.summary,
                  characters: result.summary.length,
                  createdAt: new Date().toISOString(),
                }),
                "utf8"
              )
            );
            if (cached.status !== "OK") {
              throw new Error(
                `Word用要約の一時保存に失敗しました: ${cached.errors[0]?.message ?? "unknown"}`
              );
            }
            console.log(
              `[SP PDF summary] cached for Word chars=${result.summary.length} ref=${summaryRef}`
            );
          }
          const summaryFileName = `${result.fileName.replace(/\.pdf$/i, "")}_要約.docx`;
          return [
            `全文要約完了: ${result.fileName}（${result.pageCount}ページ、${result.chunkCount}チャンク）`,
            result.fileUrl ? `原文: ${result.fileUrl}` : "",
            `Word出力時のファイル名: ${summaryFileName}`,
            summaryRef ? `Word出力用summaryRef: ${summaryRef}` : "",
            summaryRef
              ? "Word/docx出力では本文をコピーせず、create_word.content=\"[summaryRef]\"、create_word.summaryRef=上記の値、create_word.formatMode=\"markdown\"を指定してください。"
              : "以下の要約をページ参照を削らずにユーザーへ提示してください。",
            "",
            result.summary,
          ]
            .filter(Boolean)
            .join("\n");
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          return `全文要約を実行できませんでした: ${message}`;
        }
      },
      parse: (input: string) => JSON.parse(input),
    },
  };
}
