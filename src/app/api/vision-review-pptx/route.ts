export const runtime = "nodejs";
export const maxDuration = 120;

import { NextRequest, NextResponse } from "next/server";
import { execFile } from "child_process";
import { promisify } from "util";
import fs from "fs";
import os from "os";
import path from "path";
import { OpenAIVisionInstance } from "@/features/common/services/openai";

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
    };

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
  promptIntent?: Record<string, unknown>
): Promise<VisionReviewResult> {
  const openai = OpenAIVisionInstance();
  const model = process.env.AZURE_OPENAI_VISION_API_DEPLOYMENT_NAME!;

  // 全スライドをレビュー対象にする（上限 12枚 — 10枚資料は全ページカバー）
  const reviewPaths = pngPaths.slice(0, 12);

  const imageContents = reviewPaths.map((p) => ({
    type: "image_url" as const,
    image_url: {
      url: `data:image/png;base64,${fs.readFileSync(p).toString("base64")}`,
      detail: "auto" as const,  // モデルが解像度を自動判断（low固定より崩れ検出精度が高い）
    },
  }));

  const systemPrompt = [
    "You are a QA reviewer for B2B Japanese corporate presentations.",
    "Your ONLY job is to detect VISIBLE DEFECTS that users would complain about.",
    "Do NOT suggest micro-adjustments to slides that look acceptable.",
    "",
    "DEFECT CHECKLIST — only flag these:",
    "1. TEXT/SHAPE OVERLAP: Any text box or shape visibly colliding with another element.",
    "2. OVERFLOW/CLIPPING: Text or shapes cut off at slide edges.",
    "3. EMPTY BOTTOM HALF: Slide has significant whitespace in the bottom 40% with content only at top.",
    "4. TEXT-ONLY SLIDE: No icons, shapes, charts, or visual elements — pure text bullets. CRITICAL defect.",
    "5. CENTERED BODY TEXT: Body text is center-aligned (only titles should be centered).",
    "6. FULL-WIDTH COLOR BAND: A horizontal color bar spanning the full slide width (header or footer).",
    "7. NUMBERS IN TEXT: Numeric data (%, counts, dates) presented as plain text instead of a visual.",
    "8. LAYOUT TYPE MISMATCH: 3+ items that would clearly work better as cards/icons/chart.",
    "",
    "When you detect a layout-type defect (items 4, 7, 8), output a layoutType fix:",
    "- Bullet-only with numbers → layoutType='stat_callouts' (provide statCallouts data in bullets as 'value|unit|label' triplets)",
    "- 3-6 parallel items without visuals → layoutType='card_grid'",
    "- Process/capability list → layoutType='icon_rows'",
    "",
    "Return ONLY valid JSON in this exact shape:",
    '{"deckScore": <0-100>, "fixes": [<fix>, ...]}',
    "",
    "Each <fix> must be EXACTLY one of these forms:",
    '{"slideIndex":<n>, "field":"layoutType", "value":"stat_callouts"|"card_grid"|"icon_rows"|"bullets"|"process-cards"|"table"|"multi-column", "reason":"..."}',
    '{"slideIndex":<n>, "field":"bullets", "value":"bullet1|bullet2|bullet3", "reason":"..."}',
    '{"slideIndex":<n>, "field":"steps", "value":"タイトル1:説明1|タイトル2:説明2", "reason":"..."}',
    '{"slideIndex":<n>, "field":"deleteSlide", "value":"true", "reason":"..."}',
    '{"slideIndex":<n>, "field":"density", "value":"low"|"medium"|"high", "reason":"..."}',
    '{"slideIndex":<n>, "field":"coverSubtitle", "value":"<60-char viewer-facing tagline>", "reason":"..."}',
    '{"slideIndex":-1, "field":"regenerateStyle", "value":"<style direction 3-10 words>", "reason":"..."}',
    "",
    ...(promptIntent ? buildIntentChecks(promptIntent) : []),
    "RULES:",
    "- slideIndex -1 = cover slide. slideIndex 0 = first content slide.",
    "- Use regenerateStyle ONLY when overall color scheme fundamentally mismatches deck purpose.",
    "- Use deleteSlide ONLY for genuinely empty slides.",
    "- NEVER suggest fixes for slides that look acceptable — fewer fixes is better.",
    "- If deckScore >= 72 and no critical defects, return empty fixes array.",
    "Return no other text — JSON only.",
  ].join("\n");

  console.log(`[vision-review] model=${model} slides=${reviewPaths.length}`);

  const res = await openai.chat.completions.create({
    model,
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

  const choice = res.choices[0];
  console.log(`[vision-review] finish_reason=${choice?.finish_reason} usage=${JSON.stringify(res.usage)}`);

  const raw = choice?.message?.content ?? "";
  console.log(`[vision-review] raw response: ${raw.slice(0, 300)}`);

  if (!raw.trim()) {
    console.warn("[vision-review] empty response — model may not support image input");
    return { deckScore: 0, fixes: [] };
  }

  const jsonMatch = raw.match(/```json\s*([\s\S]*?)```/) ?? raw.match(/(\{[\s\S]*\})/);
  const jsonStr = jsonMatch ? (jsonMatch[1] ?? jsonMatch[0]).trim() : raw.trim();

  let parsed: any;
  try {
    parsed = JSON.parse(jsonStr);
  } catch {
    console.warn(`[vision-review] JSON parse failed, raw: ${raw.slice(0, 200)}`);
    return { deckScore: 0, fixes: [] };
  }

  const rawFixes: any[] = Array.isArray(parsed.fixes) ? parsed.fixes : [];
  const validFixes: VisionFix[] = rawFixes.filter((f) => {
    if (typeof f.slideIndex !== "number" || typeof f.field !== "string") return false;
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
      // パイプ区切り文字列で中身があること
      return typeof f.value === "string" && f.value.trim().length > 0;
    }
    return ALLOWED_SLIDE_FIELDS.has(f.field) && typeof f.value === "string";
  });

  console.log(`[vision-review] fixes accepted=${validFixes.length} / raw=${rawFixes.length}`);

  return {
    deckScore: typeof parsed.deckScore === "number" ? parsed.deckScore : 0,
    fixes: validFixes,
  };
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const pptxBlob = formData.get("pptx") as Blob | null;
    const title = String(formData.get("title") ?? "");

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

    const pngPaths = await pptxToPngs(pptxBuffer, 12);
    if (pngPaths.length === 0) {
      return NextResponse.json({ deckScore: 0, fixes: [] });
    }

    console.log(`[vision-review] Reviewing ${pngPaths.length} slides with Vision LLM`);
    const result = await reviewWithVision(pngPaths, title, promptIntent);
    console.log(`[vision-review] deckScore=${result.deckScore} fixes=${result.fixes.length}`);

    pngPaths.forEach((p) => { try { fs.unlinkSync(p); } catch {} });

    return NextResponse.json(result);
  } catch (e: any) {
    console.error("[vision-review] error (non-fatal):", String(e?.message ?? e).slice(0, 200));
    return NextResponse.json({ deckScore: 0, fixes: [] });
  }
}
