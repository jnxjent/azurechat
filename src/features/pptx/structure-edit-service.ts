import "server-only";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, writeFile, rm } from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { requiresCompanyCards, hasCompanyRows } from "./visual-composition";
import { OpenAIPptInstance } from "@/features/common/services/openai";
import { PPT_VISUAL_COMPOSITION_RULES } from "./visual-composition";
import { FindTopChatMessagesForCurrentUser } from "@/features/chat-page/chat-services/chat-message-service";
import {
  type PptxStructureContent,
  type PptxStructureIntent,
  type PptxStructureOp,
  MAX_ADDED_SLIDES,
  verifyPptxStructureContent,
} from "./structure-edit-intent";

type EditResult = { integration: "pptx_structure_edit"; layoutKinds: string[]; downloadUrl: string; fileName: string; displayName: string; message: string } | { error: string };
const SOURCE_CHARS = 16_000;

/** Review a disposable copy containing only changed slides. Never auto-edit it. */
async function reviewAddedSlides(downloadUrl: string, inputUrl: string, baseUrl: string, indices: number[], title: string): Promise<string[]> {
  let dir: string | undefined;
  try {
    if (!indices.length || new URL(downloadUrl).origin !== new URL(inputUrl).origin) throw new Error("review_source_unavailable");
    const response = await fetch(downloadUrl, { signal: AbortSignal.timeout(20_000), cache: "no-store" });
    if (!response.ok) throw new Error("review_download_failed");
    const bytes = Buffer.from(await response.arrayBuffer());
    if (bytes.length > 30_000_000) throw new Error("review_size_limit");
    dir = await mkdtemp(path.join(os.tmpdir(), "pptx-added-review-"));
    const input = path.join(dir, "input.pptx"), output = path.join(dir, "review.pptx");
    await writeFile(input, bytes);
    const script = [path.join(process.cwd(), "scripts/pptx_structure_edit.py"), path.join(process.cwd(), "src/scripts/pptx_structure_edit.py")].find(existsSync);
    if (!script) throw new Error("review_script_missing");
    await promisify(execFile)(process.platform === "win32" ? "python" : "python3", [script, "--input", input, "--output", output, "--review-slides", indices.join(",")], {
      timeout: 20_000, windowsHide: true,
      env: { ...process.env, ...(process.platform !== "win32" ? { PYTHONPATH: `/home/site/python-packages${process.env.PYTHONPATH ? `:${process.env.PYTHONPATH}` : ""}` } : {}) },
    });
    const form = new FormData();
    form.set("pptx", new Blob([new Uint8Array(await readFile(output))]), "added-slides.pptx");
    form.set("title", title);
    const reviewed = await fetch(`${baseUrl}/api/vision-review-pptx`, { method: "POST", body: form, signal: AbortSignal.timeout(60_000) });
    const result = await reviewed.json();
    if (!reviewed.ok || !Number.isFinite(result.deckScore) || result.deckScore <= 0) throw new Error("review_incomplete");
    const fixes: any[] = Array.isArray(result.fixes) ? result.fixes : [];
    console.log("[pptx_structure_edit] visual review", { slides: indices.length, score: result.deckScore, issues: fixes.length });
    return fixes.length ? ["追加スライドの見た目チェックで改善候補が見つかりました。自動修正はしていません。", ...fixes.slice(0, 5).map(fix => {
      const page = Number.isInteger(fix.slideIndex) ? indices[fix.slideIndex] : undefined;
      return `${page !== undefined ? `P${page + 1}: ` : ""}${String(fix.reason ?? fix.description ?? fix.type ?? "配置・読みやすさを確認してください").slice(0, 200)}`;
    })] : [];
  } catch {
    return ["追加スライドの画像による見た目チェックを完了できませんでした。PowerPointで表示をご確認ください。"];
  } finally {
    if (dir) await rm(dir, { recursive: true, force: true });
  }
}

/** Recent assistant answers and tool results in this thread, newest first. */
async function loadThreadSource(threadId: string): Promise<string> {
  const found = await FindTopChatMessagesForCurrentUser(threadId, 12);
  if (found.status !== "OK") return "";
  const parts: string[] = [];
  let assistants = 0;
  for (const message of found.response) {
    if (message.role === "assistant" && message.content?.trim()) {
      if (assistants >= 3) continue;
      assistants++;
      parts.push(message.content);
    } else if (message.role === "tool" && message.content?.trim()) {
      parts.push(message.content);
    }
  }
  return parts.join("\n\n---\n\n").slice(0, SOURCE_CHARS);
}

async function buildSlideContents(instruction: string, source: string, count: number, retry = false): Promise<{ slides: PptxStructureContent[]; error?: string }> {
  const model = process.env.AZURE_OPENAI_PPT_DEPLOYMENT_NAME?.trim() || process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME?.trim() || "";
  const completion = await OpenAIPptInstance().chat.completions.create({
    model,
    max_completion_tokens: 4000,
    response_format: { type: "json_object" } as const,
    messages: [
      {
        role: "system",
        content:
          "You turn chat content into slide content for an existing PowerPoint deck. Output JSON only: " +
          '{"slides":[{"title":string,"kind":"bullets"|"table"|"stats"|"two_column","keyMessage"?:string,"stats"?:[{"label":string,"value":string,"period"?:string,"unit"?:string}],"bullets"?:string[],"table"?:{"columns":string[],"rows":string[][]},"sources"?:[{"label":string,"url":string}],"caveats"?:string[]}],"error":string|null}. ' +
          `Return at most ${count} slide(s). Use ONLY facts in SOURCE or in the user's request. Copy company names, numbers and units exactly as written; never estimate, convert or add facts. ` +
          "Use kind=table as structured data for diagrams as well as comparisons (max 5 columns, 6 rows); otherwise bullets (max 6, each under 60 Japanese characters). Title under 30 characters, in Japanese. " +
          PPT_VISUAL_COMPOSITION_RULES + "\n" +
          (requiresCompanyCards(instruction) ? "MANDATORY: This request is a company overview. Return kind=table with columns 企業 and 特徴, one row per named company. Never return bullets instead. Do not invent a ranking or financial data. Put ranking/metric qualifications in caveats (max 2 short strings), not in company rows. If SOURCE has no company profiles, return an error. " : "") +
          "keyMessage: optional conclusion under 30 Japanese characters, copied from SOURCE verbatim; no inferred claims. Titles must not end with 。. Use stats only for 2–4 explicitly sourced label/value/period/unit items; keep differing periods visible. Use two_column with 3–4 short bullets copied verbatim from SOURCE on the left and a table of source facts on the right. Choose layout based on the relationship, not arbitrary variation. If SOURCE includes previous tool metadata layoutKinds, prefer a different suitable composition when facts support it; do not force a new layout at the cost of clarity or evidence. Never invent quantities for design. " +
          "sources: only URLs that appear verbatim in SOURCE and support the slide. " +
          "If the request does not identify which content to use, or SOURCE lacks it, return slides:[] and error explaining what is missing.",
      },
      { role: "user", content: `REQUEST:\n${instruction}\n\nSOURCE (recent chat answers and search results, newest first):\n${source || "(none)"}` },
    ],
  });
  const raw = completion.choices[0]?.message?.content ?? "";
  try {
    const parsed = JSON.parse(raw) as { slides?: unknown; error?: unknown };
    const slides = Array.isArray(parsed.slides) ? (parsed.slides as PptxStructureContent[]).slice(0, count) : [];
    if (requiresCompanyCards(instruction) && slides.length && !hasCompanyRows(slides)) {
      console.warn("[pptx_structure_edit] layout retry=company-rows-required");
      if (!retry) return buildSlideContents(instruction, source, count, true);
      return { slides: [], error: "企業別の情報をカード用に整理できませんでした。もう一度お試しください。" };
    }
    return { slides, ...(typeof parsed.error === "string" && parsed.error.trim() ? { error: parsed.error.trim() } : {}) };
  } catch {
    return { slides: [], error: "スライド内容を作成できませんでした。" };
  }
}

export async function executePptxStructureEdit(props: {
  intent: PptxStructureIntent;
  instruction: string;
  fileUrl: string;
  threadId: string;
  baseUrl: string;
  outputBaseName: string;
}): Promise<EditResult> {
  const { intent } = props;
  const notes: string[] = [];
  let ops: PptxStructureOp[] = [];

  if (intent.kind === "font_size") {
    ops = [{ type: "font_size", scope: intent.scope, mode: intent.mode, value: intent.value, ...(intent.slides ? { slides: intent.slides } : {}) }];
  } else if (intent.kind === "add_slide" && intent.blank) {
    ops = Array.from({ length: intent.count }, (_, i) => ({
      type: "add_slide" as const, blank: true, title: "",
      ...(intent.position !== undefined ? { position: intent.position >= 0 ? intent.position + i : intent.position } : {}),
    }));
  } else {
    const count = intent.kind === "add_slide" ? intent.count : 1;
    const source = await loadThreadSource(props.threadId).catch(() => "");
    const built = await buildSlideContents(props.instruction, source, count);
    const verified = verifyPptxStructureContent(built.slides, `${props.instruction}\n${source}`);
    console.log("[pptx_structure_edit] layouts=", verified.slides.map(slide => ({
      contentKind: slide.kind,
      layout: slide.composition?.kind ?? (slide.kind === "table" && hasCompanyRows([slide]) ? "company-cards" : slide.kind),
      rows: slide.table?.rows.length ?? 0,
    })));
    notes.push(...verified.warnings);
    if (!verified.slides.length) {
      return { error: built.error ? `スライドを追加できませんでした: ${built.error}` : "どの回答の内容を入れるか特定できませんでした。入れたい内容や対象の回答を指定してください。PPTXは変更していません。" };
    }
    ops = intent.kind === "fill_slide"
      ? [{ type: "fill_slide", slideIndex: intent.slideIndex, ...verified.slides[0] }]
      : verified.slides.map((content, i) => ({
          type: "add_slide" as const, ...content,
          ...(intent.position !== undefined ? { position: intent.position >= 0 ? intent.position + i : intent.position } : {}),
        }));
  }
  if (intent.kind === "add_slide" && intent.requestedCount > MAX_ADDED_SLIDES) {
    notes.push(`1回の依頼で追加できるのは${MAX_ADDED_SLIDES}枚までのため、${MAX_ADDED_SLIDES}枚を追加しました。`);
  }

  const res = await fetch(`${props.baseUrl}/api/edit-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      action: "pptx_structure_edit",
      fileUrl: props.fileUrl,
      instruction: props.instruction,
      threadId: props.threadId,
      outputBaseName: props.outputBaseName,
      structureOps: ops,
    }),
  });
  const json = (await res.json().catch(() => ({}))) as Record<string, unknown>;
  if (!res.ok || json.ok !== true || typeof json.downloadUrl !== "string") {
    return { error: `PPTの編集に失敗しました: ${String(json.error ?? `HTTP ${res.status}`)}。PPTXは変更していません。` };
  }
  if (intent.kind !== "font_size" && !(intent.kind === "add_slide" && intent.blank)) {
    const indices = Array.isArray(json.changedSlideIndices) ? json.changedSlideIndices.filter((n): n is number => Number.isInteger(n) && Number(n) >= 0) : [];
    notes.push(...await reviewAddedSlides(json.downloadUrl, props.fileUrl, props.baseUrl, indices, props.outputBaseName));
  }
  const changed = Array.isArray(json.changedSlideIndices) ? (json.changedSlideIndices as number[]).map((i) => `P${i + 1}`).join("、") : "";
  const summary = intent.kind === "font_size"
    ? `フォントサイズを変更しました${changed ? `（${changed}）` : ""}。`
    : intent.kind === "fill_slide"
    ? `P${intent.slideIndex + 1}に本文を書き込みました。`
    : `スライドを追加しました${changed ? `（${changed}）` : ""}。全${json.totalSlides}枚です。`;
  const warnings = [
    ...notes,
    ...(Array.isArray(json.layoutWarnings) ? (json.layoutWarnings as string[]) : []),
    ...(Array.isArray(json.visualWarnings) ? (json.visualWarnings as string[]) : []),
  ];
  const unsupported = intent.unsupported.length ? `\n⚠ 未対応のため実行していません: ${intent.unsupported.join("、")}` : "";
  const displayName = `${props.outputBaseName}.pptx`;
  return {
    integration: "pptx_structure_edit",
    layoutKinds: ops.filter(op => op.type !== "font_size").map(op => "kind" in op ? op.kind ?? "blank" : "blank"),
    downloadUrl: json.downloadUrl,
    fileName: typeof json.fileName === "string" ? json.fileName : displayName,
    displayName,
    message: summary + warnings.map((w) => `\n⚠ ${w}`).join("") + unsupported,
  };
}
