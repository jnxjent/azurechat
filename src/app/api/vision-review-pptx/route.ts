export const runtime = "nodejs";
export const maxDuration = 120;

import { NextRequest, NextResponse } from "next/server";
import { execFile } from "child_process";
import { promisify } from "util";
import fs from "fs";
import os from "os";
import path from "path";
import {
  OpenAIInstance,
  OpenAIPptVisionInstance,
} from "@/features/common/services/openai";

const execFileAsync = promisify(execFile);

const ALLOWED_SLIDE_FIELDS = new Set([
  "density",
  "textTreatment",
  "layoutType",
  "coverSubtitle",
  "subtitle",
  "bullets",
  "steps",
  "deleteSlide",
  "regenerateStyle",
  "title",
  "callout",
  // テキスト収まり系 (対応2)
  "fitTextToShape",
  "fontScaleDown",
  "trimText",
  // アイテムグループ整合系 (対応2)
  "syncItemDecorations",
  "copyItemDecoration",
  "alignItemGroup",
  "fallbackLayout",
]);

export type VisionFix =
  | {
      slideIndex: number;
      field: "density" | "textTreatment" | "layoutType" | "coverSubtitle" | "subtitle";
      value: string;
      reason?: string;
    }
  | {
      slideIndex: number;
      field: "bullets";
      value: string; // "bullet1|bullet2|bullet3" パイプ区切り
      reason?: string;
    }
  | {
      slideIndex: number;
      field: "steps";
      value: string; // "タイトル1:説明1|タイトル2:説明2" パイプ区切り
      reason?: string;
    }
  | {
      slideIndex: number;
      field: "deleteSlide";
      value: "true";
      reason?: string;
    }
  | {
      slideIndex: -1;
      field: "regenerateStyle";
      value: string; // "modern-dark recruitment" など自由記述のスタイルヒント
      reason?: string;
    }
  | {
      slideIndex: number;
      field: "metrics.colorRole";
      itemIndex: number;
      value: "primary" | "accent" | "neutral";
      reason?: string;
    }
  | {
      slideIndex: number;
      /** テキスト overflow / poorFit 系の修正アクション */
      field: "fitTextToShape" | "fontScaleDown" | "trimText" | "fallbackLayout";
      value: string; // fontScaleDown="0.85", trimText="<短縮版テキスト>", fallbackLayout="bullets"
      reason?: string;
    }
  | {
      slideIndex: number;
      /** アイテムグループ（カード+アイコン+テキスト）整合系 */
      field: "syncItemDecorations" | "copyItemDecoration" | "alignItemGroup";
      value: string; // "true" or 対象グループの説明
      reason?: string;
    };

type SlideIndexMode = "rendered-zero-based" | "cover-separated";

export type VisionReviewResult = {
  deckScore: number;
  fixes: VisionFix[];
};

/** PromptIntent から Vision LLM へのインテントチェック指示を生成 */
function buildIntentChecks(intent: Record<string, unknown>): string[] {
  const checks: string[] = [];
  const ld = (intent.layoutDirectives ?? {}) as Record<string, unknown>;
  const cd = (intent.colorDirectives ?? {}) as Record<string, unknown>;
  const purpose = String(intent.documentPurpose ?? "");
  const audience = String(intent.audience ?? "");
  const freedom = String(intent.designFreedom ?? "");

  checks.push("\n=== INTENT COMPLIANCE CHECKS (evaluate these too) ===");

  if (ld.preferTwoColumn) checks.push("- User requested TWO-COLUMN layout. If no multi-column slide is visible, suggest layoutType='multi-column' for a suitable slide.");
  if (ld.includeTables)   checks.push("- User requested TABLES. If no table slide is visible, suggest layoutType='table' for a data-heavy slide.");
  if (ld.preferMetrics)   checks.push("- User requested METRICS/KPI emphasis. Flag if no metric-cards layout is present.");
  if (ld.preferProcess)   checks.push("- User requested PROCESS/FLOW emphasis. Flag if no process-cards or timeline is present.");
  if (ld.avoidBulletOnly) checks.push("- User wants to AVOID bullet-only slides. If 3+ consecutive bullet slides exist, suggest varying layouts.");

  if (cd.primary || cd.accent) {
    const colors = [cd.primary ? `primary=#${cd.primary}` : "", cd.accent ? `accent=#${cd.accent}` : ""].filter(Boolean).join(", ");
    checks.push(`- User specified COLORS: ${colors}. If the deck colors look mismatched, suggest regenerateStyle with these color names.`);
  }

  if ((purpose === "proposal" || purpose === "ir" || audience === "executive") && freedom !== "expressive") {
    checks.push("- This is a CONSERVATIVE deck (proposal/IR/executive). If it looks too flashy, playful, or low-contrast, flag it.");
  }
  if (purpose === "recruitment" || purpose === "campaign") {
    checks.push("- This is an EXPRESSIVE deck (recruitment/campaign). If it looks too plain/corporate/green, suggest regenerateStyle.");
  }

  return checks;
}

async function pptxToPngs(pptxBuffer: Buffer, maxSlides = 12): Promise<string[]> {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "pptx-review-"));
  const pptxPath = path.join(tmpDir, "input.pptx");
  const pngDir = path.join(tmpDir, "pngs");

  try {
    fs.writeFileSync(pptxPath, pptxBuffer);
    fs.mkdirSync(pngDir, { recursive: true });

    const scriptCandidates = [
      path.join(process.cwd(), "src/scripts/pptx_to_png.py"),
      path.join(process.cwd(), "scripts/pptx_to_png.py"),
      "/home/site/wwwroot/src/scripts/pptx_to_png.py",
    ];
    const scriptPath = scriptCandidates.find((p) => fs.existsSync(p)) ?? scriptCandidates[0];
    const pythonPath = process.env.PYTHONPATH ?? "/home/site/python-packages";

    const { stdout, stderr } = await execFileAsync(
      "python3",
      [scriptPath, pptxPath, pngDir, String(maxSlides)],
      { env: { ...process.env, PYTHONPATH: pythonPath }, timeout: 90_000 }
    );

    if (stderr) {
      console.log(`[vision-review] pptx_to_png stderr: ${stderr.slice(0, 300)}`);
    }

    const pngList = stdout.replace(/\r/g, "").trim().split("\n").filter(Boolean);
    if (pngList.length === 0) {
      console.warn("[vision-review] No PNGs generated — converter not available");
    }
    return pngList;
  } catch (e) {
    console.error("[vision-review] pptxToPngs failed:", e);
    return [];
  }
}

async function reviewWithVision(
  pngPaths: string[],
  title: string,
  promptIntent?: Record<string, unknown>,
  slideIndexMode: SlideIndexMode = "rendered-zero-based",
  b3mode = false,
  b3pass = 1,
  b4mode = false
): Promise<VisionReviewResult> {
  const primaryModel =
    process.env.AZURE_OPENAI_PPT_VISION_DEPLOYMENT_NAME?.trim() ||
    process.env.AZURE_OPENAI_VISION_API_DEPLOYMENT_NAME?.trim() ||
    "";
  const fallbackModel =
    process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME?.trim() || "";

  // 全スライドをレビュー対象にする（上限 12枚 — 10枚資料は全ページカバー）
  const reviewPaths = pngPaths.slice(0, 12);

  const imageContents = reviewPaths.map((p) => ({
    type: "image_url" as const,
    image_url: {
      url: `data:image/png;base64,${fs.readFileSync(p).toString("base64")}`,
      detail: "auto" as const,  // モデルが解像度を自動判断（low固定より崩れ検出精度が高い）
    },
  }));

  const slideIndexRule = slideIndexMode === "cover-separated"
    ? "- slideIndex is the zero-based RENDERED PAGE index: 0 = cover, 1 = first content page, 2 = second content page, etc. The server normalizes it after review."
    : "- slideIndex is the zero-based RENDERED PAGE index: 0 = first page, 1 = second page, etc.";

  const targetScore = b3mode
    ? Math.max(72, parseInt(process.env.PPTX_VISION_TARGET_SCORE ?? "90", 10) || 90)
    : 72;
  const maxFixesNote = b3mode
    ? "- Each slide can receive at most 2 fixes. Prioritize the highest-impact improvement."
    : "- NEVER suggest fixes for slides that look acceptable — fewer fixes is better.";

  const systemPrompt = [
    "You are a QA reviewer for B2B Japanese corporate presentations rendered with Meiryo font.",
    b3mode && b3pass === 1
      ? "Your job is to detect VISIBLE DEFECTS and suggest HIGH-IMPACT quality improvements. Each improvement should have a clear expected benefit."
      : "Your ONLY job is to detect VISIBLE DEFECTS that users would complain about.",
    b3mode && b3pass === 1
      ? "Evaluate in this order: (1) text overflow/overlap/poor fit, (2) information priority and executive readability, (3) layout-content fit, (4) data visualization opportunities, (5) slide monotony and design consistency, (6) cover-to-conclusion story arc."
      : "Do NOT suggest micro-adjustments to slides that look acceptable.",
    "",
    "DEFECT CHECKLIST — check in priority order:",
    "== TIER 1: OVERFLOW / OVERLAP (highest priority — always fix) ==",
    "1. TEXT OVERFLOW: Text visibly extends beyond its containing box, card, or slide edge.",
    "   Note: Meiryo is wide — full-width Japanese chars overflow boxes that look OK with other fonts.",
    "2. TEXT/SHAPE OVERLAP: Text or shapes visibly collide with another element or icon.",
    "3. POOR FIT (poorFit): Long text crammed into a small card/KPI box — text is illegible or truncated.",
    "   Common pattern: summary/stat layout where right-side KPI cards have multi-line explanations.",
    "4. IMBALANCED LAYOUT: One side of a two-column layout is empty/sparse while the other is over-packed.",
    "",
    "== TIER 2: MISSING ITEM DECORATIONS (after bullet/item additions) ==",
    "5. ITEM DECORATION MISMATCH: Text items were added without corresponding background cards or icons,",
    "   while other items in the same slide have full card+icon+text groups.",
    "   This looks like: 3 items with cards/icons, 1 item with text only — clearly inconsistent.",
    "",
    "== TIER 3: STRUCTURAL / LAYOUT DEFECTS ==",
    "6. EMPTY BOTTOM HALF: Significant whitespace in the bottom 40% with content only at top.",
    "7. TEXT-ONLY SLIDE: No icons, shapes, charts, or visual elements — pure text bullets. CRITICAL.",
    "8. NUMBERS IN TEXT: Numeric data (%, counts, dates) as plain text instead of a visual.",
    "9. LAYOUT TYPE MISMATCH: 3+ items clearly better as cards/icons/chart.",
    "",
    "FIX MAPPING:",
    "- Tier 1 overflow/poorFit → PRIORITY ORDER (follow this order strictly):",
    "  1st. fitTextToShape (value='true'): auto-shrinks text into its box — NO content loss, NO layout change.",
    "  2nd. fontScaleDown (value='0.85'): reduces font size slightly — preserves ALL decorations/layout.",
    "  3rd. trimText (value='<shortened text>'): only if content is genuinely too long to fit even at 0.75x.",
    "  LAST RESORT: fallbackLayout — only if none of the above can fix the overflow.",
    "    · If slide has icon/card layout (icon_rows, card_grid, etc.), use 'icon_rows' or 'card_grid' as fallback value.",
    "    · 'fallbackLayout=bullets' is FORBIDDEN for slides that currently have icons, cards, or background shapes.",
    "    · A slide becoming plain text with no icons/cards is a quality DEFECT, not an improvement.",
    "- Tier 1 imbalanced → use 'trimText' for over-packed side",
    "- Tier 2 decoration mismatch → use 'syncItemDecorations' (value='true') or 'copyItemDecoration'",
    "- Tier 3 layout → use 'layoutType' fix (NEVER change to 'bullets' if slide has icons/cards)",
    "- Tier 3 stat/bullets → use 'bullets' or 'steps' fix",
    "- Whitespace only → use 'density' fix",
    "",
    "DECORATION PRESERVATION (critical — applies to ALL fix types):",
    "- NEVER suggest any fix that would remove icon images, icon circles (ellipse), or background card shapes.",
    "- If overflow exists on a slide with icon+card layout, use fitTextToShape or fontScaleDown first.",
    "- A slide going from 'icon_rows'/'card_grid' to plain text bullets is a quality REGRESSION.",
    "- 'layoutType=bullets' and 'fallbackLayout=bullets' are both FORBIDDEN for decorated slides.",
    "",
    "CRITICAL RULE — density prohibition:",
    "If overflow / overlap / poorFit is present, DO NOT return 'density'.",
    "'density' is for whitespace adjustment ONLY — it does NOT fix text overflow or poor fit.",
    "Returning density for overflow is a false fix that will not resolve the visual problem.",
    "",
    "When you detect a layout-type defect (items 7, 8, 9), output a layoutType fix:",
    "- Bullet-only with numbers → layoutType='stat_callouts' (provide statCallouts data in bullets as 'value|unit|label' triplets)",
    b4mode
      ? "- 2-6 narrative or parallel items without essential metrics/process structure → layoutType='asymmetric_list'"
      : "- 3-6 parallel items without visuals → layoutType='card_grid'",
    ...(b4mode ? [
      "- 1-3 strong narrative points → layoutType='editorial_statement'",
      "- Card-heavy slides should not exceed 40% of the deck or appear more than twice consecutively.",
      "- editorial_statement and asymmetric_list are intentionally decorated non-card layouts; do not mark them as text-only defects.",
      "- Do not propose card_grid/icon_rows merely to fill whitespace. Preserve meaningful whitespace and hierarchy.",
    ] : []),
    "- Process/capability list → layoutType='icon_rows'",
    "- Balanced text plus an existing visual/data panel → layoutType='split_visual'",
    "- Existing 2-3 option comparison grid → layoutType='comparison_matrix'",
    "- Explicit conclusion + evidence + next actions → layoutType='decision_summary'",
    "",
    "Return ONLY valid JSON in this exact shape:",
    '{"deckScore": <0-100>, "fixes": [<fix>, ...]}',
    "",
    "Each <fix> must be EXACTLY one of these forms:",
    '{"slideIndex":<n>, "field":"fitTextToShape", "value":"true", "reason":"..."}',
    '{"slideIndex":<n>, "field":"fontScaleDown", "value":"0.85", "reason":"..."}',
    '{"slideIndex":<n>, "field":"trimText", "value":"<shortened text>", "reason":"..."}',
    '{"slideIndex":<n>, "field":"fallbackLayout", "value":"bullets", "reason":"..."}',
    '{"slideIndex":<n>, "field":"syncItemDecorations", "value":"true", "reason":"..."}',
    '{"slideIndex":<n>, "field":"copyItemDecoration", "value":"true", "reason":"..."}',
    '{"slideIndex":<n>, "field":"alignItemGroup", "value":"true", "reason":"..."}',
    b4mode
      ? '{"slideIndex":<n>, "field":"layoutType", "value":"stat_callouts"|"card_grid"|"icon_rows"|"bullets"|"process-cards"|"table"|"multi-column"|"split_visual"|"comparison_matrix"|"decision_summary"|"editorial_statement"|"asymmetric_list", "reason":"..."}'
      : '{"slideIndex":<n>, "field":"layoutType", "value":"stat_callouts"|"card_grid"|"icon_rows"|"bullets"|"process-cards"|"table"|"multi-column"|"split_visual"|"comparison_matrix"|"decision_summary", "reason":"..."}',
    '{"slideIndex":<n>, "field":"bullets", "value":"bullet1|bullet2|bullet3", "reason":"..."}',
    '{"slideIndex":<n>, "field":"steps", "value":"タイトル1:説明1|タイトル2:説明2", "reason":"..."}',
    '{"slideIndex":<n>, "field":"deleteSlide", "value":"true", "reason":"..."}',
    '{"slideIndex":<n>, "field":"density", "value":"low"|"medium"|"high", "reason":"..."}',
    '{"slideIndex":<n>, "field":"coverSubtitle", "value":"<60-char viewer-facing tagline>", "reason":"..."}',
    '{"slideIndex":-1, "field":"regenerateStyle", "value":"<style direction 3-10 words>", "reason":"..."}',
    "",
    ...(promptIntent ? buildIntentChecks(promptIntent) : []),
    "RULES:",
    slideIndexRule,
    "- slideIndex -1 is reserved for the deck-wide regenerateStyle action only.",
    "- Use regenerateStyle ONLY when overall color scheme fundamentally mismatches deck purpose.",
    "- Use deleteSlide ONLY for genuinely empty slides.",
    maxFixesNote,
    b3mode
      ? `- If deckScore >= ${targetScore} and no critical defects, return empty fixes array.`
      : "- If deckScore >= 72 and no critical defects, return empty fixes array.",
    b3mode && b3pass === 2 ? "- PASS 2 RESTRICTION: Do NOT suggest regenerateStyle or layoutType changes. Only suggest fitTextToShape, fontScaleDown, or trimText for remaining overflow/overlap defects." : "",
    "Return no other text — JSON only.",
  ].join("\n");

  const requestReview = (
    client: ReturnType<typeof OpenAIInstance>,
    deploymentName: string
  ) =>
    client.chat.completions.create({
      model: deploymentName,
      messages: [
        { role: "system", content: systemPrompt },
        {
          role: "user",
          content: [
            {
              type: "text",
              text: `Presentation title: "${title}"\nReview the first ${reviewPaths.length} slides:`,
            },
            ...imageContents,
          ],
        },
      ],
      max_completion_tokens: 16000,
    });

  let model = primaryModel;
  let res: Awaited<ReturnType<typeof requestReview>>;
  console.log(
    `[vision-review] model=${model} fallback=${fallbackModel || "none"} slides=${reviewPaths.length}`
  );
  try {
    res = await requestReview(OpenAIPptVisionInstance(), model);
  } catch (error) {
    if (!fallbackModel || fallbackModel === model) throw error;
    console.warn(
      `[vision-review] primary model failed; retrying with ${fallbackModel}: ${
        error instanceof Error ? error.message.slice(0, 200) : String(error).slice(0, 200)
      }`
    );
    model = fallbackModel;
    res = await requestReview(OpenAIInstance(), model);
  }

  let choice = res.choices[0];
  console.log(`[vision-review] finish_reason=${choice?.finish_reason} usage=${JSON.stringify(res.usage)}`);

  let raw = choice?.message?.content ?? "";
  if (!raw.trim() && fallbackModel && fallbackModel !== model) {
    console.warn(
      `[vision-review] empty response from ${model}; retrying with ${fallbackModel}`
    );
    model = fallbackModel;
    res = await requestReview(OpenAIInstance(), model);
    choice = res.choices[0];
    raw = choice?.message?.content ?? "";
    console.log(
      `[vision-review] fallback finish_reason=${choice?.finish_reason} usage=${JSON.stringify(res.usage)}`
    );
  }
  console.log(`[vision-review] raw response: ${raw.slice(0, 300)}`);

  if (!raw.trim()) {
    console.warn("[vision-review] empty response — model may not support image input");
    return { deckScore: 0, fixes: [] };
  }

  const parseReviewJson = (text: string): any => {
    const jsonMatch =
      text.match(/```json\s*([\s\S]*?)```/) ?? text.match(/(\{[\s\S]*\})/);
    const jsonStr = jsonMatch
      ? (jsonMatch[1] ?? jsonMatch[0]).trim()
      : text.trim();
    return JSON.parse(jsonStr);
  };

  let parsed: any;
  try {
    parsed = parseReviewJson(raw);
  } catch {
    if (!fallbackModel || fallbackModel === model) {
      console.warn(`[vision-review] JSON parse failed, raw: ${raw.slice(0, 200)}`);
      return { deckScore: 0, fixes: [] };
    }
    console.warn(
      `[vision-review] JSON parse failed for ${model}; retrying with ${fallbackModel}`
    );
    model = fallbackModel;
    res = await requestReview(OpenAIInstance(), model);
    raw = res.choices[0]?.message?.content ?? "";
    try {
      parsed = parseReviewJson(raw);
    } catch {
      console.warn(
        `[vision-review] fallback JSON parse failed, raw: ${raw.slice(0, 200)}`
      );
      return { deckScore: 0, fixes: [] };
    }
  }

  const rawFixes: any[] = Array.isArray(parsed.fixes) ? parsed.fixes : [];
  // overflow系 field が存在するスライドでは density を除外（density 禁止ルールの後段実施）
  const overflowFields = new Set(["fitTextToShape", "fontScaleDown", "trimText", "syncItemDecorations", "copyItemDecoration", "alignItemGroup", "fallbackLayout"]);
  const slidesWithOverflowFix = new Set<number>(
    rawFixes.filter((f) => overflowFields.has(f.field) && typeof f.slideIndex === "number").map((f) => f.slideIndex as number)
  );
  const validFixes: VisionFix[] = rawFixes.filter((f) => {
    if (typeof f.slideIndex !== "number" || typeof f.field !== "string") return false;
    // density 禁止: overflow fix があるスライドでは density を除外
    if (f.field === "density" && slidesWithOverflowFix.has(f.slideIndex)) {
      console.log(`[vision-review] density suppressed for slide ${f.slideIndex} (overflow fix present)`);
      return false;
    }
    if (f.field === "metrics.colorRole") {
      return typeof f.itemIndex === "number" && ["primary", "accent", "neutral"].includes(f.value);
    }
    if (f.field === "deleteSlide") {
      return f.value === "true";
    }
    if (f.field === "regenerateStyle") {
      return f.slideIndex === -1 && typeof f.value === "string" && f.value.trim().length > 0;
    }
    if (f.field === "bullets" || f.field === "steps") {
      return typeof f.value === "string" && f.value.trim().length > 0;
    }
    return ALLOWED_SLIDE_FIELDS.has(f.field) && typeof f.value === "string";
  });

  const normalizedFixes = validFixes.flatMap((fix): VisionFix[] => {
    // -1 is a deck-wide action (regenerateStyle), not a rendered page number.
    if (fix.slideIndex === -1) return [fix];
    if (fix.slideIndex < 0 || fix.slideIndex >= reviewPaths.length) {
      console.warn(
        `[vision-review] ignored out-of-range rendered page index=${fix.slideIndex} pages=${reviewPaths.length}`
      );
      return [];
    }
    if (slideIndexMode !== "cover-separated") return [fix];

    const contentSlideIndex = fix.slideIndex - 1;
    console.log(
      `[vision-review] normalized rendered page[${fix.slideIndex}] -> content slide[${contentSlideIndex}]`
    );
    return [{ ...fix, slideIndex: contentSlideIndex } as VisionFix];
  });

  console.log(`[vision-review] fixes accepted=${normalizedFixes.length} / raw=${rawFixes.length}`);

  return {
    deckScore: typeof parsed.deckScore === "number" ? parsed.deckScore : 0,
    fixes: normalizedFixes,
  };
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const pptxBlob = formData.get("pptx") as Blob | null;
    const title = String(formData.get("title") ?? "");
    const slideIndexMode: SlideIndexMode =
      formData.get("slideIndexMode") === "cover-separated"
        ? "cover-separated"
        : "rendered-zero-based";

    if (!pptxBlob) {
      return NextResponse.json({ error: "pptx is required" }, { status: 400 });
    }

    const pptxBuffer = Buffer.from(await pptxBlob.arrayBuffer());
    console.log(`[vision-review] pptx size=${pptxBuffer.length} title="${title}"`);

    // PromptIntent を取得（gen-pptx から JSON 文字列で渡される）
    const intentRaw = formData.get("promptIntent");
    let promptIntent: Record<string, unknown> | undefined;
    if (typeof intentRaw === "string") {
      try { promptIntent = JSON.parse(intentRaw); } catch {}
    }
    if (promptIntent) {
      console.log(`[vision-review] intentCheck purpose=${promptIntent.documentPurpose} freedom=${promptIntent.designFreedom}`);
    }

    const b3mode = formData.get("b3mode") === "true";
    const b3pass = parseInt(String(formData.get("b3pass") ?? "1"), 10) || 1;
    const b4mode = formData.get("b4mode") === "true";
    if (b3mode) {
      console.log(`[vision-review] b3mode=true pass=${b3pass}`);
    }
    if (b4mode) {
      console.log("[vision-review] b4mode=true composition-review=enabled");
    }

    const pngPaths = await pptxToPngs(pptxBuffer, 12);
    if (pngPaths.length === 0) {
      return NextResponse.json({ deckScore: 0, fixes: [] });
    }

    console.log(
      `[vision-review] Reviewing ${pngPaths.length} slides with Vision LLM indexMode=${slideIndexMode}`
    );
    const result = await reviewWithVision(pngPaths, title, promptIntent, slideIndexMode, b3mode, b3pass, b4mode);
    console.log(`[vision-review] deckScore=${result.deckScore} fixes=${result.fixes.length}`);

    pngPaths.forEach((p) => { try { fs.unlinkSync(p); } catch {} });

    return NextResponse.json(result);
  } catch (e: any) {
    console.error("[vision-review] error (non-fatal):", String(e?.message ?? e).slice(0, 200));
    return NextResponse.json({ deckScore: 0, fixes: [] });
  }
}
