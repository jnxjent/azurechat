import "server-only";

import { createHash } from "crypto";
import {
  DownloadBlobAsText,
  GenerateSasUrl,
  UploadBlob,
} from "@/features/common/services/azure-storage";
import {
  pptxPaletteListText,
  resolvePptxPaletteInstruction,
} from "@/features/pptx/palette";
import {
  findTeamsOfficeFileCandidates,
  searchTeamsKnowledge,
  searchTeamsKnowledgeByFileFamily,
  searchTeamsKnowledgeByOfficeFileName,
  type TeamsOfficeFileCandidate,
  type TeamsSearchSource,
} from "./teams-search-service";
import {
  createTeamsPptCardEdits,
  createTeamsPptPlan,
  type TeamsPptExtractedSlide,
} from "./teams-ppt-plan-service";
import {
  buildExplicitWordReplacementInstruction,
  findTeamsWordProofreadingReplacements,
} from "./teams-word-proofread-service";

export type TeamsOfficeRequest =
  | {
      action: "pdf_to_excel";
      fileQuery: string;
    }
  | {
      action: "pdf_to_word";
      fileQuery: string;
      mode: "layout" | "editable";
    }
  | {
      action: "pdf_to_ppt";
      fileQuery: string;
      mode: "faithful" | "redesign";
    }
  | {
      action: "refine_excel_sheets";
      targetSheets: string[];
    }
  | {
      action: "create_excel";
      prompt: string;
      title: string;
    }
  | {
      action: "create_word";
      prompt: string;
      title: string;
    }
  | {
      action: "create_ppt";
      prompt: string;
      title: string;
    }
  | {
      action: "create_ppt_from_sharepoint";
      prompt: string;
      title: string;
      referenceQuery: string;
      fileFamily: string;
    }
  | {
      action: "edit_latest_ppt";
      instruction: string;
      targetPages: number[];
      targetItemCount?: number;
      cardLayout: boolean;
      referencePage?: number;
    }
  | {
      action: "edit_latest_excel";
      instruction: string;
    }
  | {
      action: "edit_latest_word";
      instruction: string;
    }
  | {
      action: "proofread_sp_word";
      fileQuery: string;
      instruction: string;
    }
  | {
      action: "ppt_color_help";
    }
  | {
      action: "edit_latest_ppt_color";
      instruction: string;
    };

export function parseTeamsOfficeRequest(
  message: string
): TeamsOfficeRequest | null {
  const normalized = message.trim().normalize("NFKC");
  const targetExcelSheets = extractTargetSheetNames(normalized);
  const hasExcelRefinementIntent =
    /(再変換|再抽出|精度|読み直|再読込|もう一度変換)/i.test(normalized);
  const asksForExcelRefinement =
    hasExcelRefinementIntent &&
    (/(excel|エクセル|xlsx|シート)/i.test(normalized) ||
      (targetExcelSheets.length > 0 &&
        !/(ppt|pptx|powerpoint|パワーポイント)/i.test(normalized)));

  if (asksForExcelRefinement) {
    return {
      action: "refine_excel_sheets",
      targetSheets: targetExcelSheets,
    };
  }

  const asksForPptColorHelp =
    /(ppt|pptx|powerpoint|パワーポイント)/i.test(normalized) &&
    /(色味|配色|カラー|色)/i.test(normalized) &&
    /(どういう|どんな|何色|候補|一覧|種類|できる|教えて)/i.test(normalized);
  if (asksForPptColorHelp) {
    return { action: "ppt_color_help" };
  }

  const asksForPptColorEdit =
    /(ppt|pptx|powerpoint|パワーポイント)/i.test(normalized) &&
    /(色味|配色|カラー|色)/i.test(normalized) &&
    /(変更|変え|替え|にして|統一|基調)/i.test(normalized);
  const selectsListedPptColor =
    /^(?:(?:では|じゃあ|それでは|ok)[、,，\s]*)?[1-6]\s*(?:番)?(?:で|に)(?:お願いします|お願い|変更して|変更|して|します)?[。.!！]?$/i.test(
      normalized
    );
  const selectsNamedPptColor =
    /(ネイビー|深緑|バーガンディ|ティール|チャコール|テラコッタ|コーラル|アンバー|ゴールド).{0,12}(?:で|に|変更|変えて|お願い)/i.test(
      normalized
    );
  if (
    asksForPptColorEdit ||
    selectsListedPptColor ||
    selectsNamedPptColor
  ) {
    return { action: "edit_latest_ppt_color", instruction: normalized };
  }

  const targetPptPages = extractPptTargetPages(normalized);
  const pptReferencePage = extractPptReferencePage(normalized);
  const pptEditTargetPages = pptReferencePage
    ? targetPptPages.filter((page) => page !== pptReferencePage)
    : targetPptPages;
  if (
    pptEditTargetPages.length > 0 &&
    /(変更|修正|編集|変えて|にして|増や|減ら|カード|レイアウト|項目数)/i.test(
      normalized
    )
  ) {
    return {
      action: "edit_latest_ppt",
      instruction: normalized,
      targetPages: pptEditTargetPages,
      targetItemCount: extractPptTargetItemCount(normalized) ?? undefined,
      cardLayout: /(カード|card)/i.test(normalized),
      referencePage: pptReferencePage ?? undefined,
    };
  }

  const asksForWord = /(word|ワード|docx)/i.test(normalized);
  const asksForPowerPoint =
    /(powerpoint|パワーポイント|pptx|ppt)/i.test(normalized);
  const asksForExcel = /(excel|エクセル|xlsx)/i.test(normalized);
  const asksForFileEdit =
    /(変更|修正|編集|追加|削除|置換|書き換|差し替|グラフ|チャート|書式|フォント|色|列|行|セル|シート|段落|見出し)/i.test(
      normalized
    );
  const isExplicitConversion = /変換/i.test(normalized);
  const asksForProofreading = /(誤字|誤記|誤変換|校正)/i.test(normalized);
  const explicitlyReferencesLatestWord =
    /(今|直前|先ほど|さっき).{0,16}(出力|作成|変換|編集).{0,16}(word|ワード|docx)/i.test(
      normalized
    );
  if (
    asksForFileEdit &&
    asksForWord &&
    asksForProofreading &&
    !isExplicitConversion &&
    !explicitlyReferencesLatestWord
  ) {
    return {
      action: "proofread_sp_word",
      fileQuery: normalized,
      instruction: normalized,
    };
  }
  if (asksForFileEdit && asksForExcel && !isExplicitConversion) {
    return { action: "edit_latest_excel", instruction: normalized };
  }
  if (asksForFileEdit && asksForWord && !isExplicitConversion) {
    return { action: "edit_latest_word", instruction: normalized };
  }
  const asksForConversion = /(変換|出力|作成|にして|して)/i.test(normalized);
  const hasPdfSource =
    /(sharepoint|\bsp\b|\bsl\b|pdf)/i.test(normalized);
  const asksForKnowledgeBasedPpt =
    asksForPowerPoint &&
    /(作成|生成|作って|まとめて)/i.test(normalized) &&
    /(sharepoint|\bsp\b|\bsl\b)/i.test(normalized) &&
    /(参考|基に|もとに|踏まえ|把握|検索|資料)/i.test(normalized);

  if (asksForKnowledgeBasedPpt) {
    const referenceQuery = extractSharePointReferenceQuery(normalized);
    if (!referenceQuery) return null;
    return {
      action: "create_ppt_from_sharepoint",
      prompt: normalized,
      title: extractDirectOutputTitle(normalized, "create_ppt"),
      referenceQuery,
      fileFamily: extractReferenceFileFamily(referenceQuery),
    };
  }

  if (asksForPowerPoint && asksForConversion && hasPdfSource) {
    const fileQuery =
      extractQuotedFileQuery(normalized) ||
      extractUnquotedFileQuery(normalized, "ppt");
    if (!fileQuery) return null;

    return {
      action: "pdf_to_ppt",
      fileQuery,
      mode: /(そのまま|忠実|原本|レイアウト維持)/i.test(normalized)
        ? "faithful"
        : "redesign",
    };
  }

  if (asksForWord && asksForConversion && hasPdfSource) {
    const fileQuery =
      extractQuotedFileQuery(normalized) ||
      extractUnquotedFileQuery(normalized, "word");
    if (!fileQuery) return null;

    return {
      action: "pdf_to_word",
      fileQuery,
      mode: /(編集可能|編集しやす|テキスト中心)/i.test(normalized)
        ? "editable"
        : "layout",
    };
  }

  const hasDocumentSource =
    /(sharepoint|\bsp\b|\bsl\b|pdf|word|docx)/i.test(normalized);

  if (asksForExcel && asksForConversion && hasDocumentSource) {
    const fileQuery =
      extractQuotedFileQuery(normalized) ||
      extractUnquotedFileQuery(normalized, "excel");
    if (!fileQuery) return null;
    return { action: "pdf_to_excel", fileQuery };
  }

  const asksForCreation = /(作成|生成|出力|作って|まとめて|にして)/i.test(
    normalized
  );
  if (!asksForCreation) return null;

  const action = asksForPowerPoint
    ? "create_ppt"
    : asksForWord
    ? "create_word"
    : asksForExcel
    ? "create_excel"
    : null;
  if (!action) return null;

  return {
    action,
    prompt: normalized,
    title: extractDirectOutputTitle(normalized, action),
  };
}

export async function executeTeamsOfficeRequest(props: {
  request: TeamsOfficeRequest;
  conversationId: string;
  userEmail?: string | null;
}): Promise<string> {
  const teamsThreadId = buildTeamsThreadId(props.conversationId);

  if (props.request.action === "edit_latest_ppt") {
    return editLatestTeamsPowerPoint({
      request: props.request,
      threadId: teamsThreadId,
    });
  }

  if (props.request.action === "ppt_color_help") {
    return [
      "PowerPointでは、次の6種類の色パレットを選べます。",
      "",
      pptxPaletteListText(),
      "",
      "基本色として、赤・青・緑・紺・紫・オレンジ・黄・ピンク・グレーも指定できます。",
      "",
      "例: 「1でお願いします」「PPTをティール×コーラルに変更して」",
    ].join("\n");
  }

  if (props.request.action === "edit_latest_ppt_color") {
    return editLatestTeamsPptColor({
      instruction: props.request.instruction,
      threadId: teamsThreadId,
    });
  }

  if (
    props.request.action === "edit_latest_excel" ||
    props.request.action === "edit_latest_word"
  ) {
    return editLatestTeamsOfficeFile({
      action: props.request.action,
      instruction: props.request.instruction,
      threadId: teamsThreadId,
    });
  }

  if (props.request.action === "proofread_sp_word") {
    return proofreadTeamsSharePointWord({
      fileQuery: props.request.fileQuery,
      instruction: props.request.instruction,
      threadId: teamsThreadId,
      userEmail: props.userEmail,
    });
  }

  if (props.request.action === "create_ppt_from_sharepoint") {
    const knowledge = await searchTeamsKnowledgeByFileFamily({
      query: props.request.referenceQuery,
      searchQueries: buildReferenceSearchQueries(
        props.request.referenceQuery,
        props.request.fileFamily
      ),
      fileFamily: props.request.fileFamily,
      userEmail: props.userEmail,
    });
    if (!knowledge.context.trim()) {
      return `「${props.request.referenceQuery}」に一致する参照可能なSharePoint資料が見つからなかったため、PowerPointの作成を停止しました。`;
    }

    const result = await createDirectOfficeFile({
      action: "create_ppt",
      prompt: props.request.prompt,
      title: props.request.title,
      threadId: teamsThreadId,
      referenceContext: knowledge.context,
    });
    if ("error" in result) {
      return `PowerPointの作成に失敗しました。\n\n${String(result.error)}`;
    }
    if (typeof result.downloadUrl !== "string") {
      return "PowerPointは作成されましたが、ダウンロードリンクを取得できませんでした。";
    }

    const outputName =
      typeof result.fileName === "string"
        ? result.fileName
        : `${props.request.title}.pptx`;
    await savePptxResult(teamsThreadId, result, outputName);
    return `SharePoint資料を基にPowerPointを作成しました。\n\n📊 [${escapeMarkdownLinkText(
      outputName
    )}](${result.downloadUrl})${formatOfficeSources(knowledge.sources)}`;
  }

  if (
    props.request.action === "create_excel" ||
    props.request.action === "create_word" ||
    props.request.action === "create_ppt"
  ) {
    const result = await createDirectOfficeFile({
      action: props.request.action,
      prompt: props.request.prompt,
      title: props.request.title,
      threadId: teamsThreadId,
    });

    if ("error" in result) {
      return `Officeファイルの作成に失敗しました。\n\n${String(
        result.error
      )}`;
    }
    if (typeof result.downloadUrl !== "string") {
      return "Officeファイルは作成されましたが、ダウンロードリンクを取得できませんでした。";
    }

    const extension =
      props.request.action === "create_excel"
        ? ".xlsx"
        : props.request.action === "create_word"
        ? ".docx"
        : ".pptx";
    const icon = props.request.action === "create_word" ? "📄" : "📊";
    const outputName =
      typeof result.fileName === "string"
        ? result.fileName
        : `${props.request.title}${extension}`;
    if (props.request.action === "create_ppt") {
      await savePptxResult(teamsThreadId, result, outputName);
    } else if (props.request.action === "create_excel") {
      await saveExcelPointer(teamsThreadId, {
        url: result.downloadUrl,
        fileName: outputName,
        savedAt: Date.now(),
      });
    } else {
      await saveWordPointer(teamsThreadId, {
        url: result.downloadUrl,
        fileName: outputName,
        savedAt: Date.now(),
      });
    }
    return `Officeファイルを作成しました。\n\n${icon} [${escapeMarkdownLinkText(
      outputName
    )}](${result.downloadUrl})`;
  }

  if (props.request.action === "refine_excel_sheets") {
    if (props.request.targetSheets.length === 0) {
      return "再変換するシート名を指定してください。例: 「P2のシートだけ再変換して」";
    }

    const result = await refineExcelSheets({
      threadId: teamsThreadId,
      targetSheets: props.request.targetSheets,
    });
    return formatExcelRefinementResult(result);
  }

  if (props.request.action === "pdf_to_word") {
    const search = await findTeamsOfficeFileCandidates({
      query: props.request.fileQuery,
      userEmail: props.userEmail,
      extensions: ["pdf"],
    });

    if (search.exactMatches.length === 0) {
      return buildNotFoundMessage(props.request.fileQuery, search.suggestions);
    }
    if (search.exactMatches.length > 1) {
      return buildMultipleFilesMessage(search.exactMatches);
    }

    const file = search.exactMatches[0];
    const result = await convertPdfToWord({
      fileUrl: file.url,
      fileName: file.name,
      threadId: teamsThreadId,
      mode: props.request.mode,
    });

    if ("error" in result) {
      return `Wordへの変換に失敗しました。\n\n${String(result.error)}`;
    }
    if (typeof result.downloadUrl !== "string") {
      return "Wordへの変換は完了しましたが、ダウンロードリンクを取得できませんでした。";
    }

    const outputName =
      typeof result.fileName === "string"
        ? result.fileName
        : file.name.replace(/\.pdf$/i, ".docx");
    await saveWordPointer(teamsThreadId, {
      url: result.downloadUrl,
      fileName: outputName,
      savedAt: Date.now(),
    });
    return `Wordへの変換が完了しました。\n\n📄 [${escapeMarkdownLinkText(
      outputName
    )}](${result.downloadUrl})`;
  }

  if (props.request.action === "pdf_to_ppt") {
    const search = await findTeamsOfficeFileCandidates({
      query: props.request.fileQuery,
      userEmail: props.userEmail,
      extensions: ["pdf"],
    });

    if (search.exactMatches.length === 0) {
      return buildNotFoundMessage(props.request.fileQuery, search.suggestions);
    }
    if (search.exactMatches.length > 1) {
      return buildMultipleFilesMessage(search.exactMatches);
    }

    const file = search.exactMatches[0];
    const result = await convertPdfToPowerPoint({
      fileUrl: file.url,
      fileName: file.name,
      threadId: teamsThreadId,
      mode: props.request.mode,
    });

    if ("error" in result) {
      return `PowerPointへの変換に失敗しました。\n\n${String(
        result.error
      )}`;
    }
    if (typeof result.downloadUrl !== "string") {
      return "PowerPointへの変換は完了しましたが、ダウンロードリンクを取得できませんでした。";
    }

    const outputName =
      typeof result.fileName === "string"
        ? result.fileName
        : file.name.replace(/\.pdf$/i, ".pptx");
    await savePptxResult(teamsThreadId, result, outputName);
    const totalPages = Number(result.totalPages ?? 0);
    const detail = totalPages > 0 ? `（${totalPages}ページを解析）` : "";
    return `PowerPointへの変換が完了しました。${detail}\n\n📊 [${escapeMarkdownLinkText(
      outputName
    )}](${result.downloadUrl})`;
  }

  const search = await findTeamsOfficeFileCandidates({
    query: props.request.fileQuery,
    userEmail: props.userEmail,
    extensions: ["pdf", "docx"],
  });

  if (search.exactMatches.length === 0) {
    return buildNotFoundMessage(props.request.fileQuery, search.suggestions);
  }
  if (search.exactMatches.length > 1) {
    return buildMultipleFilesMessage(search.exactMatches);
  }

  const file = search.exactMatches[0];
  const result = await convertDocumentToExcel({
    fileUrl: file.url,
    fileName: file.name,
    threadId: teamsThreadId,
  });

  if (result && typeof result === "object" && "error" in result) {
    return `Excelへの変換に失敗しました。\n\n${String(result.error)}`;
  }
  if (
    !result ||
    typeof result !== "object" ||
    !("downloadUrl" in result) ||
    typeof result.downloadUrl !== "string"
  ) {
    return "Excelへの変換は完了しましたが、ダウンロードリンクを取得できませんでした。";
  }

  const outputName =
    "fileName" in result && typeof result.fileName === "string"
      ? result.fileName
      : file.name.replace(/\.(pdf|docx)$/i, ".xlsx");
  const detail =
    "message" in result && typeof result.message === "string"
      ? `\n\n${result.message}`
      : "";

  return `Excelへの変換が完了しました。${detail}\n\n📊 [${escapeMarkdownLinkText(
    outputName
  )}](${result.downloadUrl})`;
}

type TeamsExcelPointer = {
  url: string;
  fileName: string;
  savedAt: number;
  sheetNames?: string[];
};

type TeamsPptxPointer = {
  url: string;
  fileName: string;
  savedAt: number;
};

type TeamsWordPointer = {
  url: string;
  fileName: string;
  savedAt: number;
};

const wordPointerBlobName = (threadId: string) =>
  `thread-${threadId}-word-latest.json`;

const pptxPointerBlobName = (threadId: string) =>
  `thread-${threadId}-pptx-latest.json`;

async function editLatestTeamsOfficeFile(props: {
  action: "edit_latest_excel" | "edit_latest_word";
  instruction: string;
  threadId: string;
}): Promise<string> {
  const isExcel = props.action === "edit_latest_excel";
  const pointer = isExcel
    ? await readExcelPointer(props.threadId)
    : await readWordPointer(props.threadId);
  const label = isExcel ? "Excel" : "Word";
  if (!pointer?.url) {
    return `このTeams会話で作成した${label}ファイルが見つかりません。先に${label}ファイルを作成してください。`;
  }

  const response = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileUrl: pointer.url,
      instruction: props.instruction,
      threadId: props.threadId,
      originalFileName: pointer.fileName,
      outputBaseName: buildEditedOfficeBaseName(pointer.fileName),
      ...(!isExcel && /(修正履歴|変更履歴)/.test(props.instruction)
        ? { trackChanges: true }
        : {}),
    }),
  });
  const result = (await response.json().catch(() => ({}))) as Record<
    string,
    unknown
  >;
  if (
    !response.ok ||
    result.ok === false ||
    typeof result.downloadUrl !== "string"
  ) {
    return `${label}編集に失敗しました。\n\n${String(
      result.error ?? `HTTP ${response.status}`
    )}`;
  }

  const extension = isExcel ? ".xlsx" : ".docx";
  const outputName =
    typeof result.fileName === "string"
      ? result.fileName
      : `${buildEditedOfficeBaseName(pointer.fileName)}${extension}`;
  if (isExcel) {
    await saveExcelPointer(props.threadId, {
      ...pointer,
      url: result.downloadUrl,
      fileName: outputName,
      savedAt: Date.now(),
    });
  } else {
    await saveWordPointer(props.threadId, {
      url: result.downloadUrl,
      fileName: outputName,
      savedAt: Date.now(),
    });
  }

  const icon = isExcel ? "📊" : "📄";
  return `${label}ファイルを編集しました。\n\n${icon} [${escapeMarkdownLinkText(
    outputName
  )}](${result.downloadUrl})`;
}

async function proofreadTeamsSharePointWord(props: {
  fileQuery: string;
  instruction: string;
  threadId: string;
  userEmail?: string | null;
}): Promise<string> {
  const search = await findTeamsOfficeFileCandidates({
    query: props.fileQuery,
    userEmail: props.userEmail,
    extensions: ["docx"],
  });
  if (search.exactMatches.length === 0) {
    return buildNotFoundMessage(props.fileQuery, search.suggestions);
  }
  const file = selectMostSpecificOfficeFileMatch(
    props.fileQuery,
    search.exactMatches
  );
  if (!file) {
    return buildMultipleFilesMessage(search.exactMatches);
  }
  console.log("[teams-word-proofread] selected source file", {
    requested: props.fileQuery,
    selected: file.name,
    candidates: search.exactMatches.map((candidate) => candidate.name),
  });
  const knowledge = await searchTeamsKnowledgeByOfficeFileName({
    fileName: file.name,
    userEmail: props.userEmail,
  });
  if (!knowledge.context.trim()) {
    return `「${file.name}」は見つかりましたが、校正する本文をAI Searchから取得できませんでした。`;
  }

  const replacements = await findTeamsWordProofreadingReplacements({
    documentContext: knowledge.context,
    instruction: props.instruction,
  });
  console.log("[teams-word-proofread] corrections identified", {
    fileName: file.name,
    corrections: replacements.length,
  });
  if (replacements.length === 0) {
    return `「${file.name}」を確認しましたが、確実に修正できる誤字・誤記は見つかりませんでした。`;
  }

  const resolvedUrl = await resolveTeamsOfficeSourceUrl({
    fileUrl: file.url,
    fileName: file.name,
    threadId: props.threadId,
  });
  const response = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileUrl: resolvedUrl,
      instruction: buildExplicitWordReplacementInstruction(replacements),
      threadId: props.threadId,
      originalFileName: file.name,
      outputBaseName: buildEditedOfficeBaseName(file.name),
      trackChanges: true,
    }),
  });
  const result = (await response.json().catch(() => ({}))) as Record<
    string,
    unknown
  >;
  if (
    !response.ok ||
    result.ok === false ||
    typeof result.downloadUrl !== "string"
  ) {
    return `Word校正に失敗しました。\n\n${String(
      result.error ?? `HTTP ${response.status}`
    )}`;
  }

  const outputName =
    typeof result.fileName === "string"
      ? result.fileName
      : `${buildEditedOfficeBaseName(file.name)}.docx`;
  await saveWordPointer(props.threadId, {
    url: result.downloadUrl,
    fileName: outputName,
    savedAt: Date.now(),
  });

  return `SharePointのWordを校正し、${replacements.length}件を変更履歴付きで修正しました。\n\n📄 [${escapeMarkdownLinkText(
    outputName
  )}](${result.downloadUrl})`;
}

function selectMostSpecificOfficeFileMatch(
  query: string,
  candidates: TeamsOfficeFileCandidate[]
): TeamsOfficeFileCandidate | null {
  if (candidates.length === 0) return null;

  const uniqueByName = new Map<string, TeamsOfficeFileCandidate>();
  for (const candidate of candidates) {
    const key = normalizeOfficeFileMatchText(candidate.name);
    if (!uniqueByName.has(key)) uniqueByName.set(key, candidate);
  }
  const unique = Array.from(uniqueByName.values());
  if (unique.length === 1) return unique[0];

  const normalizedQuery = normalizeOfficeFileMatchText(query);
  const embedded = unique
    .filter((candidate) =>
      normalizedQuery.includes(normalizeOfficeFileMatchText(candidate.name))
    )
    .sort(
      (a, b) =>
        normalizeOfficeFileMatchText(b.name).length -
        normalizeOfficeFileMatchText(a.name).length
    );
  if (embedded.length === 0) return null;

  const firstLength = normalizeOfficeFileMatchText(embedded[0].name).length;
  const secondLength = embedded[1]
    ? normalizeOfficeFileMatchText(embedded[1].name).length
    : -1;
  return firstLength > secondLength ? embedded[0] : null;
}

function normalizeOfficeFileMatchText(value: string): string {
  return value
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\.(pdf|docx|xlsx)$/i, "")
    .replace(/[\s\u3000「」『』【】()（）・_\-]/g, "");
}

function buildEditedOfficeBaseName(fileName: string): string {
  return (
    fileName
      .replace(/\.(xlsx|xlsm|xls|docx)$/i, "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim()
      .slice(0, 40) + "_編集済み"
  );
}

async function editLatestTeamsPowerPoint(props: {
  request: Extract<TeamsOfficeRequest, { action: "edit_latest_ppt" }>;
  threadId: string;
}): Promise<string> {
  const pointer = await readPptxPointer(props.threadId);
  if (!pointer?.url) {
    return "このTeams会話で作成したPowerPointが見つかりません。先にPowerPointを作成してください。";
  }

  let result: Record<string, unknown>;
  let editLabel: "編集済み" | "カード型" | "レイアウト変更" = "編集済み";
  if (props.request.referencePage) {
    editLabel = "レイアウト変更";
    const slideEdits = props.request.targetPages.map((page) => ({
      slideIndex: page - 1,
      copySlideLayoutFromReference: {
        referenceSlideIndex: props.request.referencePage! - 1,
        preserveTargetText: true,
      },
    }));
    const editResponse = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fileUrl: pointer.url,
        instruction: props.request.instruction,
        threadId: props.threadId,
        action: "apply_pptx_plan",
        outputBaseName: buildEditedPptxBaseName(pointer.fileName, editLabel),
        plan: { slideEdits },
      }),
    });
    result = (await editResponse.json().catch(() => ({}))) as Record<
      string,
      unknown
    >;
    if (!editResponse.ok) {
      return `PowerPoint編集に失敗しました。\n\n${String(
        result.error ?? `HTTP ${editResponse.status}`
      )}`;
    }
  } else if (props.request.cardLayout) {
    editLabel = "カード型";
    const summaryResponse = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fileUrl: pointer.url,
        instruction: "",
        threadId: props.threadId,
        action: "extract_pptx_summary",
      }),
    });
    const summary = (await summaryResponse.json().catch(() => ({}))) as Record<
      string,
      unknown
    >;
    if (!summaryResponse.ok || !Array.isArray(summary.slides)) {
      return `PowerPointのスライド情報を取得できませんでした。${
        typeof summary.error === "string" ? `\n\n${summary.error}` : ""
      }`;
    }

    const pageSet = new Set(props.request.targetPages.map((page) => page - 1));
    const targetSlides = summary.slides
      .map(normalizeExtractedPptSlide)
      .filter(
        (slide): slide is TeamsPptExtractedSlide =>
          slide !== null && pageSet.has(slide.slideIndex)
      );
    if (targetSlides.length !== pageSet.size) {
      return "指定されたページの一部がPowerPoint内に見つかりません。ページ番号を確認してください。";
    }

    const targetItemCount = props.request.targetItemCount ?? 4;
    const slideEdits = await createTeamsPptCardEdits({
      slides: targetSlides,
      targetItemCount,
      instruction: props.request.instruction,
    });
    const editResponse = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fileUrl: pointer.url,
        instruction: props.request.instruction,
        threadId: props.threadId,
        action: "apply_pptx_plan",
        outputBaseName: buildEditedPptxBaseName(pointer.fileName, editLabel),
        plan: { slideEdits },
      }),
    });
    result = (await editResponse.json().catch(() => ({}))) as Record<
      string,
      unknown
    >;
    if (!editResponse.ok) {
      return `PowerPoint編集に失敗しました。\n\n${String(
        result.error ?? `HTTP ${editResponse.status}`
      )}`;
    }
  } else {
    const editResponse = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fileUrl: pointer.url,
        instruction: props.request.instruction,
        threadId: props.threadId,
        outputBaseName: buildEditedPptxBaseName(pointer.fileName, editLabel),
      }),
    });
    result = (await editResponse.json().catch(() => ({}))) as Record<
      string,
      unknown
    >;
    if (!editResponse.ok) {
      return `PowerPoint編集に失敗しました。\n\n${String(
        result.error ?? `HTTP ${editResponse.status}`
      )}`;
    }
  }

  if (result.ok === false || typeof result.downloadUrl !== "string") {
    return `PowerPoint編集に失敗しました。\n\n${String(
      result.error ?? "ダウンロードURLを取得できませんでした。"
    )}`;
  }
  const outputName =
    typeof result.fileName === "string"
      ? result.fileName
      : `${buildEditedPptxBaseName(pointer.fileName, editLabel)}.pptx`;
  await savePptxResult(props.threadId, result, outputName);
  return `PowerPointを編集しました（${props.request.targetPages
    .map((page) => `P${page}`)
    .join("、")}）。\n\n📊 [${escapeMarkdownLinkText(outputName)}](${
    result.downloadUrl
  })`;
}

async function editLatestTeamsPptColor(props: {
  instruction: string;
  threadId: string;
}): Promise<string> {
  const pointer = await readPptxPointer(props.threadId);
  if (!pointer?.url) {
    return "このTeams会話で作成したPowerPointが見つかりません。先にPowerPointを作成してください。";
  }

  const palette = resolvePptxPaletteInstruction(props.instruction);
  if (!palette) {
    return `色パレットを特定できませんでした。\n\n${pptxPaletteListText()}\n\n例: 「4でお願いします」`;
  }

  const response = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileUrl: pointer.url,
      instruction: props.instruction,
      threadId: props.threadId,
      action: "apply_pptx_plan",
      outputBaseName: buildEditedPptxBaseName(pointer.fileName, "色変更"),
      plan: {
        slideEdits: [],
        deckEdits: {
          accentColor: palette.accentColor,
          preserveTextColors: true,
          ...(palette.paletteKey ? { paletteKey: palette.paletteKey } : {}),
        },
      },
    }),
  });
  const result = (await response.json().catch(() => ({}))) as Record<
    string,
    unknown
  >;
  if (
    !response.ok ||
    result.ok === false ||
    typeof result.downloadUrl !== "string"
  ) {
    return `PowerPointの色変更に失敗しました。\n\n${String(
      result.error ?? `HTTP ${response.status}`
    )}`;
  }

  const outputName =
    typeof result.fileName === "string"
      ? result.fileName
      : `${buildEditedPptxBaseName(pointer.fileName, "色変更")}.pptx`;
  await savePptxResult(props.threadId, result, outputName);
  return `PowerPointの色味を変更しました。\n\n📊 [${escapeMarkdownLinkText(
    outputName
  )}](${result.downloadUrl})`;
}

function normalizeExtractedPptSlide(value: unknown): TeamsPptExtractedSlide | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  const slideIndex = Number(source.slideIndex);
  if (!Number.isInteger(slideIndex)) return null;
  return {
    slideIndex,
    title: typeof source.title === "string" ? source.title : "",
    bullets: Array.isArray(source.bullets)
      ? source.bullets.filter(
          (bullet): bullet is string => typeof bullet === "string"
        )
      : [],
  };
}

async function savePptxResult(
  threadId: string,
  result: Record<string, unknown>,
  fallbackFileName: string
): Promise<void> {
  if (typeof result.downloadUrl !== "string") return;
  const pointer: TeamsPptxPointer = {
    url: result.downloadUrl,
    fileName:
      typeof result.fileName === "string" ? result.fileName : fallbackFileName,
    savedAt: Date.now(),
  };
  const response = await UploadBlob(
    "dl-link",
    pptxPointerBlobName(threadId),
    Buffer.from(JSON.stringify(pointer))
  );
  if (response.status !== "OK") {
    throw new Error("Failed to save Teams PowerPoint pointer.");
  }
}

async function readPptxPointer(
  threadId: string
): Promise<TeamsPptxPointer | null> {
  const response = await DownloadBlobAsText(
    "dl-link",
    pptxPointerBlobName(threadId)
  );
  if (response.status === "OK") {
    try {
      return JSON.parse(response.response) as TeamsPptxPointer;
    } catch {
      // 既存のAzureChat PPTポインターへフォールバックする。
    }
  }

  const existing = await DownloadBlobAsText(
    "pptx",
    `thread-${threadId}-pptx-pointer.json`
  );
  if (existing.status !== "OK") return null;
  try {
    const pointer = JSON.parse(existing.response) as {
      containerName?: string;
      blobName?: string;
      fileName?: string;
      savedAt?: string;
    };
    const accountName = process.env.AZURE_STORAGE_ACCOUNT_NAME?.trim();
    if (!accountName || !pointer.blobName) return null;
    const containerName = pointer.containerName?.trim() || "pptx";
    return {
      url: `https://${accountName}.blob.core.windows.net/${containerName}/${encodeURIComponent(
        pointer.blobName
      )}`,
      fileName: pointer.fileName?.trim() || "PowerPoint.pptx",
      savedAt: pointer.savedAt ? Date.parse(pointer.savedAt) || Date.now() : Date.now(),
    };
  } catch {
    return null;
  }
}

function buildEditedPptxBaseName(
  fileName: string,
  label: "編集済み" | "カード型" | "レイアウト変更" | "色変更" = "編集済み"
): string {
  const base = fileName
    .replace(/\.pptx$/i, "")
    .replace(/_(?:rev\d+|編集済み|カード型|レイアウト変更|色変更)$/i, "")
    .replace(/[\\/:*?"<>|]/g, "")
    .trim()
    .slice(0, 40);
  return `${base || "PowerPoint"}_${label}`;
}

const excelPointerBlobName = (threadId: string) =>
  `thread-${threadId}-excel-latest.json`;

async function convertDocumentToExcel(props: {
  fileUrl: string;
  fileName: string;
  threadId: string;
}): Promise<Record<string, unknown>> {
  const sourceUrl = await resolveTeamsOfficeSourceUrl(props);
  const response = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileUrl: sourceUrl,
      instruction: "",
      threadId: props.threadId,
      action: "pdf_to_excel",
      outputBaseName: props.fileName,
    }),
  });

  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    return {
      error: `PDF→Excel変換に失敗しました: HTTP ${response.status}${
        detail ? ` ${detail.slice(0, 200)}` : ""
      }`,
    };
  }

  const result = (await response.json()) as Record<string, unknown>;
  if (result.engine === "none") {
    return { error: "このファイルから表データを抽出できませんでした。" };
  }
  if (typeof result.downloadUrl !== "string") {
    return { error: "ダウンロードURLが取得できませんでした。" };
  }

  const fileName =
    typeof result.fileName === "string"
      ? result.fileName
      : props.fileName.replace(/\.(pdf|docx)$/i, ".xlsx");
  const sheetNames = Array.isArray(result.sheetNames)
    ? result.sheetNames.filter(
        (item): item is string => typeof item === "string"
      )
    : undefined;
  await saveExcelPointer(props.threadId, {
    url: result.downloadUrl,
    fileName,
    savedAt: Date.now(),
    sheetNames,
  });

  const tables = Number(result.tables ?? 0);
  const sheets = Number(result.sheets ?? 0);
  const pages = Number(result.pages ?? 0);
  return {
    ...result,
    fileName,
    message:
      tables > 0
        ? `${pages}ページを変換しました（テーブル${tables}個、${sheets}シート）。`
        : `${pages}ページを変換しました。`,
  };
}

async function convertPdfToWord(props: {
  fileUrl: string;
  fileName: string;
  threadId: string;
  mode: "layout" | "editable";
}): Promise<Record<string, unknown>> {
  const sourceUrl = await resolveTeamsOfficeSourceUrl(props);
  const response = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileUrl: sourceUrl,
      instruction: "",
      threadId: props.threadId,
      action: "pdf_to_word",
      mode: props.mode,
      outputBaseName: props.fileName,
    }),
  });

  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    return {
      error: `PDF→Word変換に失敗しました: HTTP ${response.status}${
        detail ? ` ${detail.slice(0, 200)}` : ""
      }`,
    };
  }

  const result = (await response.json()) as Record<string, unknown>;
  if (result.engine === "none") {
    return { error: "このPDFをWordへ変換できませんでした。" };
  }
  if (typeof result.downloadUrl !== "string") {
    return { error: "ダウンロードURLが取得できませんでした。" };
  }
  return result;
}

async function convertPdfToPowerPoint(props: {
  fileUrl: string;
  fileName: string;
  threadId: string;
  mode: "faithful" | "redesign";
}): Promise<Record<string, unknown>> {
  const sourceUrl = await resolveTeamsOfficeSourceUrl(props);
  const analyzeResponse = await fetch(
    `${getOfficeApiBaseUrl()}/api/analyze-doc-vision`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fileUrl: sourceUrl,
        maxPages: 30,
        mode: props.mode,
      }),
    }
  );

  if (!analyzeResponse.ok) {
    const detail = await analyzeResponse.text().catch(() => "");
    return {
      error: `PDF解析に失敗しました: HTTP ${analyzeResponse.status}${
        detail ? ` ${detail.slice(0, 200)}` : ""
      }`,
    };
  }

  const analysis = (await analyzeResponse.json()) as Record<string, unknown>;
  const slides = Array.isArray(analysis.slides) ? analysis.slides : [];
  if (analysis.ok !== true || slides.length === 0) {
    return {
      error:
        typeof analysis.error === "string"
          ? analysis.error
          : "PDFの解析結果が空でした。",
    };
  }

  const firstSlide = slides[0] as Record<string, unknown> | undefined;
  const title =
    (typeof firstSlide?.title === "string" && firstSlide.title.trim()) ||
    props.fileName.replace(/\.pdf$/i, "") ||
    "プレゼンテーション";
  const fileBaseName = props.fileName
    .replace(/\.pdf$/i, "")
    .replace(/[\\/:*?"<>|]/g, "")
    .trim()
    .slice(0, 40);

  const generateResponse = await fetch(`${getOfficeApiBaseUrl()}/api/gen-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      title,
      slides,
      threadId: props.threadId,
      deckPreferences: {},
      mode: props.mode,
      fileBaseName,
    }),
  });

  if (!generateResponse.ok) {
    const detail = await generateResponse.text().catch(() => "");
    return {
      error: `PowerPoint生成に失敗しました: HTTP ${generateResponse.status}${
        detail ? ` ${detail.slice(0, 200)}` : ""
      }`,
    };
  }

  const result = (await generateResponse.json()) as Record<string, unknown>;
  if (typeof result.downloadUrl !== "string") {
    return { error: "ダウンロードURLが取得できませんでした。" };
  }

  return {
    ...result,
    totalPages: Number(analysis.totalPages ?? slides.length),
  };
}

async function createDirectOfficeFile(props: {
  action: "create_excel" | "create_word" | "create_ppt";
  prompt: string;
  title: string;
  threadId: string;
  referenceContext?: string;
}): Promise<Record<string, unknown>> {
  if (props.action === "create_ppt") {
    const plan = await createTeamsPptPlan({
      prompt: props.prompt,
      title: props.title,
      referenceContext: props.referenceContext,
    });
    return postOfficeGenerationApi("/api/gen-pptx", {
      title: plan.title,
      slides: plan.slides,
      threadId: props.threadId,
      deckPreferences: {},
      fileBaseName: sanitizeOfficeBaseName(plan.title),
    });
  }

  const endpoint =
    props.action === "create_excel" ? "/api/gen-excel" : "/api/gen-word";
  return postOfficeGenerationApi(endpoint, {
    content: props.prompt,
    instruction: props.prompt,
    title: props.title,
    threadId: props.threadId,
    ...(props.action === "create_word" ? { fontFace: "Meiryo" } : {}),
  });
}

function formatOfficeSources(sources: TeamsSearchSource[]): string {
  if (sources.length === 0) return "";
  const lines = [...sources]
    .sort((left, right) =>
      left.name.localeCompare(right.name, "ja", {
        numeric: true,
        sensitivity: "base",
      })
    )
    .slice(0, 8)
    .map((source) => {
      const name = escapeMarkdownLinkText(source.name);
      return source.url ? `- [${name}](${source.url})` : `- ${name}`;
    });
  // ゼロ幅スペースの段落を挟み、成果物リンクと参照資料の間だけを1行空ける。
  // 末尾には改行を追加しないため、参照資料の下は省スペースのままになる。
  return `\n\n\u200b\n\n参照資料\n${lines.join("\n")}`;
}

async function postOfficeGenerationApi(
  endpoint: string,
  body: Record<string, unknown>
): Promise<Record<string, unknown>> {
  const response = await fetch(`${getOfficeApiBaseUrl()}${endpoint}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    return {
      error: `${endpoint} の呼び出しに失敗しました: HTTP ${response.status}${
        detail ? ` ${detail.slice(0, 200)}` : ""
      }`,
    };
  }

  const result = (await response.json()) as Record<string, unknown>;
  if (typeof result.downloadUrl !== "string") {
    return { error: "ダウンロードURLが取得できませんでした。" };
  }
  return result;
}

function sanitizeOfficeBaseName(value: string): string {
  return (
    value.replace(/[\\/:*?"<>|]/g, "").trim().slice(0, 40) ||
    "presentation"
  );
}

async function resolveTeamsOfficeSourceUrl(props: {
  fileUrl: string;
  fileName: string;
  threadId: string;
}): Promise<string> {
  const url = new URL(props.fileUrl);
  if (!url.hostname.endsWith(".sharepoint.com")) return props.fileUrl;

  const tenantId = requiredOfficeEnv("AZURE_AD_TENANT_ID");
  const clientId = requiredOfficeEnv("AZURE_AD_CLIENT_ID");
  const clientSecret = requiredOfficeEnv("AZURE_AD_CLIENT_SECRET");
  const tokenResponse = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: "client_credentials",
        scope: "https://graph.microsoft.com/.default",
      }),
    }
  );
  const tokenData = (await tokenResponse.json().catch(() => ({}))) as {
    access_token?: string;
    error_description?: string;
  };
  if (!tokenResponse.ok || !tokenData.access_token) {
    throw new Error(
      `SharePoint認証に失敗しました: ${
        tokenData.error_description ?? `HTTP ${tokenResponse.status}`
      }`
    );
  }

  const pathParts = decodeURIComponent(url.pathname).split("/").filter(Boolean);
  const sitesIndex = pathParts.indexOf("sites");
  if (sitesIndex < 0 || !pathParts[sitesIndex + 1]) {
    throw new Error("SharePointサイトをURLから特定できませんでした。");
  }
  const siteName = pathParts[sitesIndex + 1];
  const graphHeaders = { Authorization: `Bearer ${tokenData.access_token}` };
  const siteResponse = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${url.hostname}:/sites/${encodeURIComponent(
      siteName
    )}:`,
    { headers: graphHeaders }
  );
  const siteData = (await siteResponse.json().catch(() => ({}))) as {
    id?: string;
  };
  if (!siteResponse.ok || !siteData.id) {
    throw new Error(`SharePointサイトの取得に失敗しました: HTTP ${siteResponse.status}`);
  }

  const contentResponse = url.pathname.includes("/_layouts/")
    ? await downloadSharePointLayoutFile({
        siteId: siteData.id,
        fileName: url.searchParams.get("file") ?? props.fileName,
        headers: graphHeaders,
      })
    : await downloadSharePointPathFile({
        siteId: siteData.id,
        librarySegment: pathParts[sitesIndex + 2] ?? "",
        filePath: pathParts.slice(sitesIndex + 3).join("/"),
        headers: graphHeaders,
      });

  if (!contentResponse?.ok) {
    throw new Error(
      `SharePointファイルの取得に失敗しました: HTTP ${
        contentResponse?.status ?? "unknown"
      }`
    );
  }

  const safeFileName = props.fileName.replace(/[\\/:*?"<>|]/g, "_");
  const blobPath = `${props.threadId}/${safeFileName}`;
  const upload = await UploadBlob(
    "dl-link",
    blobPath,
    Buffer.from(await contentResponse.arrayBuffer())
  );
  if (upload.status !== "OK") {
    throw new Error("SharePointファイルの一時保存に失敗しました。");
  }
  const sas = await GenerateSasUrl("dl-link", blobPath);
  if (sas.status !== "OK") {
    throw new Error("SharePointファイルの一時URLを生成できませんでした。");
  }
  return sas.response;
}

async function downloadSharePointPathFile(props: {
  siteId: string;
  librarySegment: string;
  filePath: string;
  headers: { Authorization: string };
}): Promise<Response | null> {
  const drivesResponse = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${props.siteId}/drives`,
    { headers: props.headers }
  );
  const drivesData = (await drivesResponse.json().catch(() => ({}))) as {
    value?: Array<{ id?: string; name?: string; webUrl?: string }>;
  };
  const drive = drivesData.value?.find((item) => {
    const webSegment = decodeURIComponent(
      String(item.webUrl ?? "").split("/").pop() ?? ""
    );
    return item.name === props.librarySegment || webSegment === props.librarySegment;
  });
  if (!drive?.id || !props.filePath) return null;

  return fetch(
    `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${props.filePath
      .split("/")
      .map(encodeURIComponent)
      .join("/")}:/content`,
    { headers: props.headers }
  );
}

async function downloadSharePointLayoutFile(props: {
  siteId: string;
  fileName: string;
  headers: { Authorization: string };
}): Promise<Response | null> {
  const drivesResponse = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${props.siteId}/drives`,
    { headers: props.headers }
  );
  const drivesData = (await drivesResponse.json().catch(() => ({}))) as {
    value?: Array<{ id?: string }>;
  };

  for (const drive of drivesData.value ?? []) {
    if (!drive.id) continue;
    const searchResponse = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${drive.id}/root/search(q='${encodeURIComponent(
        props.fileName
      )}')`,
      { headers: props.headers }
    );
    const searchData = (await searchResponse.json().catch(() => ({}))) as {
      value?: Array<{ id?: string; name?: string; file?: unknown }>;
    };
    const item = searchData.value?.find(
      (candidate) =>
        candidate.file &&
        candidate.name?.toLowerCase() === props.fileName.toLowerCase()
    );
    if (item?.id) {
      return fetch(
        `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${item.id}/content`,
        { headers: props.headers }
      );
    }
  }
  return null;
}

async function refineExcelSheets(props: {
  threadId: string;
  targetSheets: string[];
}): Promise<Record<string, unknown>> {
  const pointer = await readExcelPointer(props.threadId);
  if (!pointer?.url) {
    return {
      error:
        "このTeams会話で変換したExcelが見つかりません。先にPDF／WordをExcelへ変換してください。",
    };
  }

  const sheetToProcess = props.targetSheets[0];
  const outputFileName = buildRefinedFileName(pointer.fileName);
  const response = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      excelFileUrl: pointer.url,
      targetSheets: [sheetToProcess],
      outputFileName,
      instruction: "",
      threadId: props.threadId,
      action: "refine_excel_pages",
    }),
  });

  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    return {
      error: `Excel再変換に失敗しました: HTTP ${response.status}${
        detail ? ` ${detail.slice(0, 200)}` : ""
      }`,
    };
  }

  const result = (await response.json()) as Record<string, unknown>;
  const didRefine = Number(result.refined ?? 0) > 0;
  if (!didRefine || typeof result.downloadUrl !== "string") {
    return {
      error: `「${sheetToProcess}」は再変換できませんでした。最新Excelは更新していません。`,
    };
  }

  const fileName =
    typeof result.fileName === "string" ? result.fileName : outputFileName;
  await saveExcelPointer(props.threadId, {
    ...pointer,
    url: result.downloadUrl,
    fileName,
    savedAt: Date.now(),
  });

  return {
    ...result,
    fileName,
    processedSheet: sheetToProcess,
    message: `「${sheetToProcess}」をVision AIで再変換しました。`,
  };
}

async function saveExcelPointer(
  threadId: string,
  pointer: TeamsExcelPointer
): Promise<void> {
  const response = await UploadBlob(
    "dl-link",
    excelPointerBlobName(threadId),
    Buffer.from(JSON.stringify(pointer))
  );
  if (response.status !== "OK") {
    throw new Error("Failed to save Teams Excel pointer.");
  }
}

async function saveWordPointer(
  threadId: string,
  pointer: TeamsWordPointer
): Promise<void> {
  const response = await UploadBlob(
    "dl-link",
    wordPointerBlobName(threadId),
    Buffer.from(JSON.stringify(pointer))
  );
  if (response.status !== "OK") {
    throw new Error("Failed to save Teams Word pointer.");
  }
}

async function readWordPointer(
  threadId: string
): Promise<TeamsWordPointer | null> {
  const response = await DownloadBlobAsText(
    "dl-link",
    wordPointerBlobName(threadId)
  );
  if (response.status === "OK") {
    try {
      return JSON.parse(response.response) as TeamsWordPointer;
    } catch {
      // 既存Word編集APIのポインターへフォールバックする。
    }
  }

  const existing = await DownloadBlobAsText(
    "docx",
    `thread-${threadId}-word-pointer.json`
  );
  if (existing.status !== "OK") return null;
  try {
    const pointer = JSON.parse(existing.response) as {
      blobName?: string;
      fileName?: string;
      savedAt?: number;
    };
    if (!pointer.blobName) return null;
    const sas = await GenerateSasUrl("docx", pointer.blobName);
    if (sas.status !== "OK") return null;
    return {
      url: sas.response,
      fileName: pointer.fileName?.trim() || "Word.docx",
      savedAt: pointer.savedAt ?? Date.now(),
    };
  } catch {
    return null;
  }
}

async function readExcelPointer(
  threadId: string
): Promise<TeamsExcelPointer | null> {
  const response = await DownloadBlobAsText(
    "dl-link",
    excelPointerBlobName(threadId)
  );
  if (response.status !== "OK") return null;

  try {
    return JSON.parse(response.response) as TeamsExcelPointer;
  } catch {
    return null;
  }
}

function buildRefinedFileName(fileName: string): string {
  const base = fileName
    .replace(/\.xlsx$/i, "")
    .replace(/(_精度向上後)+$/, "")
    .replace(/[　 ﻿<>:"/\\|?*\x00-\x1f]/g, "_");
  const revision = base.match(/^(.*?)_rev(\d+)$/i);
  return revision
    ? `${revision[1]}_rev${Number(revision[2]) + 1}.xlsx`
    : `${base}_rev1.xlsx`;
}

function getOfficeApiBaseUrl(): string {
  return (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME
      ? `https://${process.env.WEBSITE_HOSTNAME}`
      : "http://localhost:3000")
  ).replace(/\/+$/, "");
}

function requiredOfficeEnv(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`Missing environment variable: ${name}`);
  return value;
}

function extractTargetSheetNames(message: string): string[] {
  const names = new Set<string>();
  const quotedPattern = /[「『"“]([^」』"”]+)[」』"”]/g;
  let match: RegExpExecArray | null;
  while ((match = quotedPattern.exec(message)) !== null) {
    const value = match[1]?.trim();
    if (value && !/再変換|再抽出/.test(value)) names.add(value);
  }

  // ハイフンを含むシート名（例: P3-T2）は分割せず、一つの名前として扱う。
  // 複合名を先に取り除くことで、後段の単一名検索が P3 / T2 を重複追加しないようにする。
  const compoundSheetPattern = /\b([a-z]+\d+(?:-[a-z]+\d+)+)\b/gi;
  const messageWithoutCompoundNames = message.replace(
    compoundSheetPattern,
    (value) => {
      names.add(value.toUpperCase());
      return " ".repeat(value.length);
    }
  );

  const cellStylePattern = /\b([a-z]+\d+)\b/gi;
  while ((match = cellStylePattern.exec(messageWithoutCompoundNames)) !== null) {
    if (match[1]) names.add(match[1].toUpperCase());
  }

  return Array.from(names);
}

function formatExcelRefinementResult(result: unknown): string {
  if (result && typeof result === "object" && "error" in result) {
    return `Excelシートの再変換に失敗しました。\n\n${String(result.error)}`;
  }
  if (
    !result ||
    typeof result !== "object" ||
    !("downloadUrl" in result) ||
    typeof result.downloadUrl !== "string"
  ) {
    return "Excelシートの再変換は完了しましたが、ダウンロードリンクを取得できませんでした。";
  }

  const fileName =
    "fileName" in result && typeof result.fileName === "string"
      ? result.fileName
      : "refined.xlsx";
  const message =
    "message" in result && typeof result.message === "string"
      ? result.message
      : "指定シートを再変換しました。";

  return `${message}\n\n📊 [${escapeMarkdownLinkText(fileName)}](${result.downloadUrl})`;
}

function extractQuotedFileQuery(message: string): string | null {
  return (
    message.match(/「([^」]+)」/)?.[1]?.trim() ??
    message.match(/『([^』]+)』/)?.[1]?.trim() ??
    message.match(/["“]([^"”]+)["”]/)?.[1]?.trim() ??
    null
  );
}

function extractDirectOutputTitle(
  message: string,
  action: "create_excel" | "create_word" | "create_ppt"
): string {
  const quoted = extractQuotedFileQuery(message);
  if (quoted) return quoted.slice(0, 60);

  const about = message.match(/(.{1,60}?)について(?:の)?/i)?.[1]?.trim();
  if (about) {
    return about
      .replace(/^(?:excel|エクセル|word|ワード|powerpoint|パワーポイント)で?/i, "")
      .trim()
      .slice(0, 60);
  }

  if (action === "create_excel") return "Teams作成Excel";
  if (action === "create_word") return "Teams作成Word";
  return "Teams作成PowerPoint";
}

function extractPptTargetPages(message: string): number[] {
  const pages = new Set<number>();
  const compact = message.match(
    /P\s*([0-9]+(?:\s*[,、，]\s*(?:P\s*)?[0-9]+)+)/i
  )?.[1];
  if (compact) {
    for (const value of compact.split(/[,、，]/)) {
      const page = Number(value.replace(/P/gi, "").trim());
      if (Number.isInteger(page) && page > 0) pages.add(page);
    }
  }

  const explicit = /P\s*([0-9]+)/gi;
  let match: RegExpExecArray | null;
  while ((match = explicit.exec(message)) !== null) {
    const page = Number(match[1]);
    if (Number.isInteger(page) && page > 0) pages.add(page);
  }
  return Array.from(pages).sort((a, b) => a - b);
}

function extractPptReferencePage(message: string): number | null {
  const raw = message.match(
    /(?:デザイン|レイアウト|形式|型)(?:を|は)?\s*P\s*([0-9]+)\s*(?:と|の)(?:同じ|同様)/i
  )?.[1];
  const page = Number(raw);
  return Number.isInteger(page) && page > 0 ? page : null;
}

function extractPptTargetItemCount(message: string): number | null {
  const raw = message.match(/項目数(?:を)?\s*([0-9]+)(?:\s*(?:つ|個|項目))?/i)?.[1];
  const count = Number(raw);
  return Number.isInteger(count) && count >= 2 && count <= 6 ? count : null;
}

function extractSharePointReferenceQuery(message: string): string | null {
  const match = message.match(
    /(?:sharepoint|\bsp\b|\bsl\b)(?:上|内)?(?:にある|の)?\s*(.+?)\s*を(?:参考に|基に|もとに|踏まえ)/i
  )?.[1];
  return match?.trim().replace(/[。．]+$/g, "") || null;
}

function extractReferenceFileFamily(referenceQuery: string): string {
  const beforeYear = referenceQuery.split(/(?:の)?20\d{2}年度?/i)[0]?.trim();
  return beforeYear || referenceQuery.trim();
}

function buildReferenceSearchQueries(
  referenceQuery: string,
  fileFamily: string
): string[] {
  const year = referenceQuery.match(/20\d{2}(?:年度)?/)?.[0] ?? "";
  const quarters = Array.from(
    new Set(referenceQuery.match(/[1-4]\s*Q/gi)?.map((q) => q.replace(/\s/g, "")) ?? [])
  );
  if (quarters.length === 0) return [referenceQuery];
  return quarters.map((quarter) =>
    [fileFamily, year, quarter].filter(Boolean).join(" ")
  );
}

function extractUnquotedFileQuery(
  message: string,
  output: "excel" | "word" | "ppt"
): string | null {
  const outputPattern =
    output === "excel"
      ? "(?:excel|エクセル|xlsx)"
      : output === "word"
      ? "(?:word|ワード|docx)"
      : "(?:powerpoint|パワーポイント|pptx|ppt)";
  const patterns = [
    new RegExp(
      `(?:sharepoint|\\bsp\\b)(?:上|内)?(?:にある|の)?\\s*(?:pdf|word|docx)?\\s*(?:ファイル)?\\s*(.+?)\\s*を\\s*${outputPattern}`,
      "i"
    ),
    new RegExp(
      `\\bsl\\b(?:上|内)?(?:にある|の)?\\s*(?:pdf|word|docx)?\\s*(?:ファイル)?\\s*(.+?)\\s*を\\s*${outputPattern}`,
      "i"
    ),
    new RegExp(
      `(?:pdf|word|docx)(?:ファイル)?\\s*(.+?)\\s*を\\s*${outputPattern}`,
      "i"
    ),
  ];

  for (const pattern of patterns) {
    const match = message.match(pattern)?.[1]?.trim();
    if (match) return match;
  }
  return null;
}

function buildNotFoundMessage(
  query: string,
  suggestions: TeamsOfficeFileCandidate[]
): string {
  if (suggestions.length === 0) {
    return `「${query}」に一致する、アクセス可能なPDF／Wordファイルが見つかりませんでした。`;
  }

  const list = suggestions.map((file) => `- ${file.name}`).join("\n");
  return `「${query}」に完全一致するファイルが見つかりませんでした。候補は次のとおりです。\n\n${list}\n\nファイル名を「」で囲んで再度指定してください。`;
}

function buildMultipleFilesMessage(files: TeamsOfficeFileCandidate[]): string {
  const list = files.slice(0, 10).map((file) => `- ${file.name}`).join("\n");
  return `複数のファイルが見つかりました。変換するファイル名を「」で囲んで指定してください。\n\n${list}`;
}

function buildTeamsThreadId(conversationId: string): string {
  return `teams-${createHash("sha256")
    .update(conversationId)
    .digest("hex")
    .slice(0, 32)}`;
}

function escapeMarkdownLinkText(value: string): string {
  return value.replace(/([\\[\]])/g, "\\$1");
}
