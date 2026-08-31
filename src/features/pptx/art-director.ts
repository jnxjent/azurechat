/**
 * B4 Art Director — スライドごとのレイアウト計画を構造化JSONで生成する。
 * Art DirectorはTerraを「アートディレクター」として使い、既存レイアウトの選択、
 * 視覚的役割の分類、density/contentBudgetの指定を行う。
 * 内容生成・数値生成・レイアウト追加は行わない。
 */

import { OpenAIPptInstance } from "@/features/common/services/openai";

// 循環importを避けるため gen-pptx/route から import せず最小型定義を使用
type SlideInput = {
  title: string;
  bullets?: string[];
  layoutType?: string;
  tableRows?: unknown[][];
  columns?: unknown[];
  visualBlocks?: unknown[];
  metrics?: unknown[];
  statCallouts?: unknown[];
  cards?: unknown[];
  narrativeRole?: "opening" | "context" | "problem" | "value" | "evidence"
                  | "comparison" | "process" | "risk" | "decision" | "closing";
  narrativeImportance?: "hero" | "primary" | "support";
  keyTakeaway?: string;
  narrativeTransition?: string;
  storyClaim?: string;
  storyEvidenceQuotes?: string[];
  storyPlanApplied?: boolean;
};
type PromptIntentInput = {
  documentPurpose?: string;
  audience?: string;
  designFreedom?: string;
};

// ─── 型定義 ───────────────────────────────────────────────────────────────────

export type ArtDirectorMode = "off" | "audit" | "apply";

export type SlideVisualRole =
  | "opening"
  | "context"
  | "evidence"
  | "comparison"
  | "process"
  | "decision"
  | "closing";

export type DominantElement =
  | "headline"
  | "metric"
  | "chart"
  | "comparison"
  | "process"
  | "cards"
  | "table"
  | "illustration";

export type VisualDensity = "low" | "medium" | "high";

export type ArtDirectedSlide = {
  slideIndex: number;
  visualRole: SlideVisualRole;
  layoutType:
    | "bullets"
    | "table"
    | "multi-column"
    | "diagram"
    | "company-overview"
    | "process-cards"
    | "metric-cards"
    | "timeline"
    | "stat_callouts"
    | "card_grid"
    | "icon_rows"
    | "split_visual"
    | "comparison_matrix"
    | "decision_summary"
    | "editorial_statement"
    | "asymmetric_list"
    | "closing";
  dominantElement: DominantElement;
  density: VisualDensity;
  contentBudget: {
    maxItems: number;
    maxCharsPerItem: number;
  };
  sourceBulletIndices: number[];
  visualReason: string;
};

export type DeckArtDirection = {
  version: 1;
  designThesis: string;
  motif: "rounded-cards" | "icon-circles" | "editorial-type" | "data-first";
  dominantTone: "light" | "dark" | "mixed";
  slides: ArtDirectedSlide[];
};

// ─── モード判定 ────────────────────────────────────────────────────────────────

export function getArtDirectorMode(): ArtDirectorMode {
  const raw = (process.env.PPTX_ART_DIRECTOR_MODE ?? "off").trim().toLowerCase();
  if (raw === "audit" || raw === "apply") return raw;
  return "off";
}

export function isArtDirectorActive(): boolean {
  return getArtDirectorMode() !== "off";
}

// ─── 許可レイアウト一覧（Preflight共通） ────────────────────────────────────────

export const ALLOWED_LAYOUT_TYPES: ReadonlySet<ArtDirectedSlide["layoutType"]> = new Set<ArtDirectedSlide["layoutType"]>([
  "bullets", "table", "multi-column", "diagram", "company-overview",
  "process-cards", "metric-cards", "timeline", "stat_callouts",
  "card_grid", "icon_rows", "split_visual", "comparison_matrix",
  "decision_summary", "editorial_statement", "asymmetric_list", "closing",
]);

// ─── Art Direction生成 ────────────────────────────────────────────────────────

const MAX_STREAK = (): number => {
  const v = parseInt(process.env.PPTX_ART_DIRECTOR_MAX_LAYOUT_STREAK ?? "2", 10);
  return isNaN(v) || v < 1 ? 2 : v;
};

const MAX_CARD_RATIO = (): number => {
  const value = Number.parseFloat(process.env.PPTX_ART_DIRECTOR_MAX_CARD_RATIO ?? "0.4");
  return Number.isFinite(value) && value >= 0.2 && value <= 0.6 ? value : 0.4;
};

const MAX_CARD_FAMILY_STREAK = (): number => {
  const value = Number.parseInt(process.env.PPTX_ART_DIRECTOR_MAX_CARD_STREAK ?? "2", 10);
  return Number.isFinite(value) && value >= 1 && value <= 3 ? value : 2;
};

const CARD_HEAVY_LAYOUTS: ReadonlySet<ArtDirectedSlide["layoutType"]> = new Set<ArtDirectedSlide["layoutType"]>([
  "company-overview", "process-cards", "metric-cards", "card_grid", "icon_rows",
]);

function hasExplicitDecisionStructure(slide: SlideInput): boolean {
  let conclusionCount = 0;
  let evidenceCount = 0;
  let actionCount = 0;
  for (const rawBullet of slide.bullets ?? []) {
    if (!rawBullet?.trim()) continue;
    const match = rawBullet.trim().match(/^(結論|判断|推奨|根拠|理由|次のアクション|アクション|対応|conclusion|decision|evidence|reason|next action|action)\s*[：:]\s*(.+)$/i);
    if (!match?.[2]?.trim()) return false;
    if (/^(結論|判断|推奨|conclusion|decision)$/i.test(match[1])) conclusionCount += 1;
    else if (/^(根拠|理由|evidence|reason)$/i.test(match[1])) evidenceCount += 1;
    else if (/^(次のアクション|アクション|対応|next action|action)$/i.test(match[1])) actionCount += 1;
  }
  return conclusionCount === 1 && evidenceCount >= 1 && evidenceCount <= 3 && actionCount >= 1 && actionCount <= 3;
}

function hasComparisonStructure(slide: SlideInput): boolean {
  const sourceBullets = (slide.bullets ?? []).map((bullet) => bullet.trim()).filter(Boolean);
  if (Array.isArray(slide.columns) && slide.columns.length >= 2 && slide.columns.length <= 3) {
    const columnItems = slide.columns.map((rawColumn) => {
      if (!rawColumn || typeof rawColumn !== "object") return [];
      const column = rawColumn as { bullets?: unknown[] };
      return (column.bullets ?? []).map((value) => String(value).trim()).filter(Boolean);
    });
    return columnItems.every((items) => items.length <= 6) &&
      sourceBullets.every((bullet) => columnItems.some((items) => items.includes(bullet)));
  }
  if (!Array.isArray(slide.tableRows) || slide.tableRows.length < 2 || slide.tableRows.length > 7) return false;
  const columnCount = Math.max(
    ...slide.tableRows.map((row) => Array.isArray(row) ? row.length : 0),
    0
  );
  if (columnCount < 2 || columnCount > 3) return false;
  return sourceBullets.every((bullet) => {
    const colonIndex = bullet.search(/[：:]/);
    return slide.tableRows!.some((rawRow) => {
      const cells = (Array.isArray(rawRow) ? rawRow : []).map((cell) => String(cell).trim());
      if (cells.includes(bullet)) return true;
      if (colonIndex <= 0) return false;
      const key = bullet.slice(0, colonIndex).trim();
      const value = bullet.slice(colonIndex + 1).trim();
      return cells.includes(key) && cells.includes(value);
    });
  });
}

function getPhase3EligibleLayouts(slide: SlideInput): ArtDirectedSlide["layoutType"][] {
  const eligible: ArtDirectedSlide["layoutType"][] = [];
  const bulletCount = (slide.bullets ?? []).filter((bullet) => bullet?.trim()).length;
  const visualCount = Math.max(
    Array.isArray(slide.metrics) ? slide.metrics.length : 0,
    Array.isArray(slide.statCallouts) ? slide.statCallouts.length : 0,
    Array.isArray(slide.cards) ? slide.cards.length : 0,
    Array.isArray(slide.visualBlocks) ? slide.visualBlocks.length : 0
  );
  if (bulletCount >= 1 && bulletCount <= 5 && visualCount >= 1 && visualCount <= 4) {
    eligible.push("split_visual");
  }
  if (hasComparisonStructure(slide)) eligible.push("comparison_matrix");
  if (hasExplicitDecisionStructure(slide)) eligible.push("decision_summary");
  if (bulletCount >= 1 && bulletCount <= 3) eligible.push("editorial_statement");
  if (bulletCount >= 2 && bulletCount <= 6) eligible.push("asymmetric_list");
  return eligible;
}

function ensurePhase3Selection(
  direction: DeckArtDirection,
  slides: SlideInput[]
): DeckArtDirection {
  const phase3Layouts: ReadonlySet<ArtDirectedSlide["layoutType"]> = new Set<ArtDirectedSlide["layoutType"]>([
    "split_visual", "comparison_matrix", "decision_summary", "editorial_statement", "asymmetric_list",
  ]);
  if (direction.slides.some((slide) => phase3Layouts.has(slide.layoutType))) return direction;

  const priority: ArtDirectedSlide["layoutType"][] = [
    "decision_summary", "comparison_matrix", "editorial_statement", "asymmetric_list", "split_visual",
  ];
  for (const layoutType of priority) {
    const slideIndex = slides.findIndex((slide) => getPhase3EligibleLayouts(slide).includes(layoutType));
    if (slideIndex < 0) continue;
    const sourceBulletCount = (slides[slideIndex].bullets ?? []).filter((bullet) => bullet?.trim()).length;
    const directedSlides = direction.slides.map((slide) =>
      slide.slideIndex === slideIndex
        ? {
            ...slide,
            layoutType,
            dominantElement:
              layoutType === "decision_summary" ? "headline" as const :
              layoutType === "comparison_matrix" ? "comparison" as const :
              layoutType === "split_visual" ? "cards" as const :
              "headline" as const,
            contentBudget: {
              ...slide.contentBudget,
              maxItems: Math.max(slide.contentBudget.maxItems, sourceBulletCount),
            },
          }
        : slide
    );
    console.log(`[ppt-b4] phase3 selected-by-guard=${slideIndex}:${layoutType}`);
    return { ...direction, slides: directedSlides };
  }
  return direction;
}

function applyNarrativeSignals(
  direction: DeckArtDirection,
  sourceSlides: SlideInput[]
): DeckArtDirection {
  const roleMap: Record<NonNullable<SlideInput["narrativeRole"]>, SlideVisualRole> = {
    opening: "opening",
    context: "context",
    problem: "context",
    value: "evidence",
    evidence: "evidence",
    comparison: "comparison",
    process: "process",
    risk: "evidence",
    decision: "decision",
    closing: "closing",
  };

  const directedSlides = direction.slides.map((directedSlide) => {
    const source = sourceSlides[directedSlide.slideIndex];
    if (!source?.narrativeRole && !source?.narrativeImportance) return directedSlide;

    const visualRole = source.narrativeRole
      ? roleMap[source.narrativeRole]
      : directedSlide.visualRole;
    let layoutType = directedSlide.layoutType;
    let dominantElement = directedSlide.dominantElement;
    let density = directedSlide.density;

    // A hero page without an essential table/KPI/process structure should read
    // as one strong message, not as a collection of equal-weight cards.
    const hasStructuredContent =
      (source.tableRows?.length ?? 0) >= 2 ||
      (source.columns?.length ?? 0) >= 2 ||
      (source.metrics?.length ?? 0) > 0 ||
      (source.statCallouts?.length ?? 0) > 0 ||
      (source.cards?.length ?? 0) > 0 ||
      (source.visualBlocks?.length ?? 0) > 0;
    if (
      source.narrativeImportance === "hero" &&
      !hasStructuredContent &&
      visualRole !== "closing" &&
      getPhase3EligibleLayouts(source).includes("editorial_statement")
    ) {
      layoutType = "editorial_statement";
      dominantElement = "headline";
      density = "low";
    } else if (source.narrativeImportance === "hero") {
      density = "low";
    } else if (source.narrativeImportance === "support" && density === "low") {
      density = "medium";
    }

    console.log(
      `[ppt-b4] narrative slide=${directedSlide.slideIndex} role=${source.narrativeRole ?? "unchanged"} ` +
      `importance=${source.narrativeImportance ?? "unchanged"} layout=${layoutType}`
    );
    return {
      ...directedSlide,
      visualRole,
      layoutType,
      dominantElement,
      density,
      visualReason: `${directedSlide.visualReason}; narrative-${source.narrativeRole ?? "role-unchanged"}-${source.narrativeImportance ?? "importance-unchanged"}`,
    };
  });
  return { ...direction, slides: directedSlides };
}

function pickNonCardLayout(slide: SlideInput): ArtDirectedSlide["layoutType"] | null {
  const bulletCount = (slide.bullets ?? []).filter((bullet) => bullet?.trim()).length;
  if (bulletCount >= 1 && bulletCount <= 3) return "editorial_statement";
  if (bulletCount >= 2 && bulletCount <= 6) return "asymmetric_list";
  if (bulletCount > 0) return "bullets";
  return null;
}

function replaceCardLayout(
  directedSlide: ArtDirectedSlide,
  sourceSlide: SlideInput | undefined,
  reason: "family-streak" | "ratio"
): ArtDirectedSlide | null {
  if (!sourceSlide) return null;
  const replacement = pickNonCardLayout(sourceSlide);
  if (!replacement) return null;
  const bulletCount = (sourceSlide.bullets ?? []).filter((bullet) => bullet?.trim()).length;
  console.log(
    `[ppt-b4] composition-repair slide=${directedSlide.slideIndex} ` +
    `from=${directedSlide.layoutType} to=${replacement} reason=${reason}`
  );
  return {
    ...directedSlide,
    layoutType: replacement,
    dominantElement: "headline",
    contentBudget: {
      ...directedSlide.contentBudget,
      maxItems: Math.max(directedSlide.contentBudget.maxItems, bulletCount),
    },
    visualReason: `${directedSlide.visualReason}; composition-${reason}`,
  };
}

function enforceCompositionDiversity(
  direction: DeckArtDirection,
  sourceSlides: SlideInput[]
): DeckArtDirection {
  const directedSlides = direction.slides.map((slide) => ({ ...slide }));
  const maxFamilyStreak = MAX_CARD_FAMILY_STREAK();
  let familyStreak = 0;

  // Different card renderer names still produce the same visual rhythm. Break
  // that family-level streak before enforcing the deck-wide ratio.
  for (let index = 0; index < directedSlides.length; index++) {
    const directedSlide = directedSlides[index];
    if (!CARD_HEAVY_LAYOUTS.has(directedSlide.layoutType)) {
      familyStreak = 0;
      continue;
    }
    familyStreak += 1;
    if (familyStreak <= maxFamilyStreak) continue;
    const replacement = replaceCardLayout(
      directedSlide,
      sourceSlides[directedSlide.slideIndex],
      "family-streak"
    );
    if (replacement) {
      directedSlides[index] = replacement;
      familyStreak = 0;
    }
  }

  const maxCardSlides = Math.max(1, Math.floor(directedSlides.length * MAX_CARD_RATIO()));
  let cardCount = directedSlides.filter((slide) => CARD_HEAVY_LAYOUTS.has(slide.layoutType)).length;
  if (cardCount > maxCardSlides) {
    const conversionPriority: Partial<Record<ArtDirectedSlide["layoutType"], number>> = {
      card_grid: 0,
      icon_rows: 1,
      "company-overview": 2,
      "process-cards": 3,
      "metric-cards": 4,
    };
    const candidates = directedSlides
      .map((slide, index) => ({ slide, index }))
      .filter(({ slide }) => CARD_HEAVY_LAYOUTS.has(slide.layoutType))
      .sort((left, right) =>
        (conversionPriority[left.slide.layoutType] ?? 9) -
          (conversionPriority[right.slide.layoutType] ?? 9) ||
        right.slide.slideIndex - left.slide.slideIndex
      );
    for (const candidate of candidates) {
      if (cardCount <= maxCardSlides) break;
      const replacement = replaceCardLayout(
        candidate.slide,
        sourceSlides[candidate.slide.slideIndex],
        "ratio"
      );
      if (!replacement) continue;
      directedSlides[candidate.index] = replacement;
      cardCount -= 1;
    }
  }

  console.log(
    `[ppt-b4] composition-summary cardHeavy=${cardCount}/${directedSlides.length} ` +
    `max=${maxCardSlides} maxStreak=${maxFamilyStreak}`
  );
  return { ...direction, slides: directedSlides };
}

export async function generateDeckArtDirection(props: {
  title: string;
  slides: SlideInput[];
  promptIntent?: PromptIntentInput | null;
  designThesis?: string;
  signal?: AbortSignal;
}): Promise<DeckArtDirection | null> {
  const { title, slides, promptIntent, designThesis, signal } = props;
  if (!slides || slides.length === 0) return null;

  const pptModel = process.env.AZURE_OPENAI_PPT_DEPLOYMENT_NAME?.trim() || "";
  if (!pptModel) {
    console.warn("[ppt-b4] art-director: AZURE_OPENAI_PPT_DEPLOYMENT_NAME not set — skipping Art Director");
    return null;
  }

  const openai = OpenAIPptInstance();
  const maxStreak = MAX_STREAK();
  const phase3Eligibility = slides.flatMap((slide, slideIndex) =>
    getPhase3EligibleLayouts(slide).map((layoutType) => `${slideIndex}:${layoutType}`)
  );
  console.log(
    `[ppt-b4] phase3 eligible=${phase3Eligibility.length > 0 ? phase3Eligibility.join(",") : "none"}`
  );

  const slidesSummary = slides
    .map((s, i) => {
      const bullets = (s.bullets ?? []).slice(0, 6).map((b, bi) => `  [${bi}] ${b}`).join("\n");
      const currentLayout = s.layoutType ?? "bullets";
      const structures = [
        Array.isArray(s.columns) && s.columns.length >= 2 ? `columns:${s.columns.length}` : "",
        Array.isArray(s.tableRows) && s.tableRows.length >= 2 ? `tableRows:${s.tableRows.length}` : "",
        Array.isArray(s.metrics) && s.metrics.length > 0 ? `metrics:${s.metrics.length}` : "",
        Array.isArray(s.statCallouts) && s.statCallouts.length > 0 ? `statCallouts:${s.statCallouts.length}` : "",
        Array.isArray(s.cards) && s.cards.length > 0 ? `cards:${s.cards.length}` : "",
        Array.isArray(s.visualBlocks) && s.visualBlocks.length > 0 ? `visualBlocks:${s.visualBlocks.length}` : "",
      ].filter(Boolean).join(", ") || "none";
      const phase3Eligible = getPhase3EligibleLayouts(s).join(",") || "none";
      const narrative = [
        s.narrativeRole ? `role:${s.narrativeRole}` : "",
        s.narrativeImportance ? `importance:${s.narrativeImportance}` : "",
        s.keyTakeaway ? `takeaway:${s.keyTakeaway}` : "",
        s.storyPlanApplied ? `storyPlanned:true` : "",
        s.storyEvidenceQuotes?.length ? `evidenceCount:${s.storyEvidenceQuotes.length}` : "",
      ].filter(Boolean).join(", ") || "none";
      return `Slide ${i} (current: ${currentLayout}, structures: ${structures}, phase3Eligible: ${phase3Eligible}, narrative: ${narrative}): "${s.title}"\n${bullets}`;
    })
    .join("\n\n");

  const intentHint = promptIntent
    ? `Purpose: ${promptIntent.documentPurpose}, Audience: ${promptIntent.audience}, Freedom: ${promptIntent.designFreedom}`
    : "";

  const systemPrompt = `You are an Art Director for a Japanese business presentation tool.
Your task: analyze slide content and assign the best visual layout to each slide.

RULES:
- Return valid JSON ONLY. No markdown, no explanation.
- Output "slides" array MUST have exactly ${slides.length} entries (index 0 to ${slides.length - 1}).
- Same layout must NOT appear more than ${maxStreak} slides in a row.
- Card-heavy layouts (company-overview, process-cards, metric-cards, card_grid, icon_rows) must not exceed 40% of the deck or appear more than 2 slides consecutively.
- Prefer editorial_statement for 1-3 narrative points and asymmetric_list for 2-6 narrative points when no table, comparison, process, or KPI structure is essential.
- Do not use a card layout merely because a slide has several parallel bullets. Reserve cards for genuinely grouped entities, capabilities, or metrics.
- For slides with ≥2 numeric KPI values: prefer metric-cards, stat_callouts, or table.
- For comparisons (before/after, options): prefer comparison_matrix when a 2-3 column/table structure is reported; otherwise prefer multi-column.
- For processes/timelines: prefer process-cards or timeline.
- phase3Eligible is computed by application code. Never choose a Phase 3 layout that is not listed for that slide.
- When one or more slides have a Phase 3 eligible layout, prefer using an eligible Phase 3 layout on at least one suitable slide instead of mechanically retaining an older layout.
- Use split_visual only when the structure summary reports metrics, statCallouts, cards, or visualBlocks and the slide also has 1-5 bullets.
- Use comparison_matrix only when the structure summary reports 2-3 columns or a 2-3 column table.
- Use decision_summary only when bullets explicitly label a conclusion/decision, evidence/reasons, and next actions. Do not infer these roles from unlabeled text.
- Use editorial_statement only for 1-3 non-empty bullets. Use asymmetric_list only for 2-6 non-empty bullets.
- Do NOT invent new content, numbers, or names. Use only what is in the bullets.
- When narrative metadata is present, treat its role and importance as upstream content strategy: hero is visually dominant, primary carries the argument, and support is visually subordinate.
- Keep visualRole consistent with the reported narrative role. Do not turn a support/context page into the deck's visual climax.
- keyTakeaway is planning metadata only. Do not add it as new visible body text.
- When storyPlanned:true is reported, the current layout is an upstream semantic choice tied to evidence. Preserve it unless it violates an explicit structure or composition rule.
- Use diagram only when visualBlocks is reported. Do not request a structure that the summary does not report and that cannot be derived from bullets.
- Do NOT specify colors, coordinates, or font sizes.
- layoutType must be one of: bullets, table, multi-column, diagram, company-overview, process-cards, metric-cards, timeline, stat_callouts, card_grid, icon_rows, split_visual, comparison_matrix, decision_summary, editorial_statement, asymmetric_list, closing.
- sourceBulletIndices: list all bullet indices ([0],[1],...) used on this slide. Must not be empty.
- contentBudget.maxItems: how many items this layout can display clearly (3-8 for most layouts).
- contentBudget.maxItems must be at least the number of source bullets so application safety checks do not reject the layout.
- contentBudget.maxCharsPerItem: max Japanese characters per item (30-90).
- Do NOT assign illustration as dominantElement unless the slide truly benefits from one.

JSON schema:
{
  "version": 1,
  "designThesis": "one sentence describing the deck's visual strategy",
  "motif": "rounded-cards" | "icon-circles" | "editorial-type" | "data-first",
  "dominantTone": "light" | "dark" | "mixed",
  "slides": [
    {
      "slideIndex": 0,
      "visualRole": "opening" | "context" | "evidence" | "comparison" | "process" | "decision" | "closing",
      "layoutType": "bullets" | "table" | "multi-column" | "diagram" | "company-overview" | "process-cards" | "metric-cards" | "timeline" | "stat_callouts" | "card_grid" | "icon_rows" | "split_visual" | "comparison_matrix" | "decision_summary" | "editorial_statement" | "asymmetric_list" | "closing",
      "dominantElement": "headline" | "metric" | "chart" | "comparison" | "process" | "cards" | "table" | "illustration",
      "density": "low" | "medium" | "high",
      "contentBudget": { "maxItems": 5, "maxCharsPerItem": 60 },
      "sourceBulletIndices": [0, 1, 2],
      "visualReason": "brief reason in English"
    }
  ]
}`;

  const userPrompt = `Deck title: "${title}"
${intentHint}
${designThesis ? `Design context: ${designThesis}` : ""}

Slides to direct (${slides.length} total):
${slidesSummary}

Return JSON only.`;

  try {
    const startedAt = Date.now();
    const res = await openai.chat.completions.create(
      {
        model: pptModel,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userPrompt },
        ],
        max_completion_tokens: 4000,
        response_format: { type: "json_object" },
      },
      { signal }
    );

    const raw = res.choices?.[0]?.message?.content ?? "";
    const durationMs = Date.now() - startedAt;

    if (!raw.trim()) {
      console.warn(`[ppt-b4] art-director: empty response durationMs=${durationMs}`);
      return null;
    }

    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch {
      // Try to extract JSON from markdown fence
      const fenceMatch = raw.match(/```(?:json)?\s*([\s\S]+?)```/);
      if (fenceMatch) {
        try { parsed = JSON.parse(fenceMatch[1]); } catch { parsed = null; }
      }
    }

    if (!parsed || typeof parsed !== "object") {
      console.warn(`[ppt-b4] art-director: invalid JSON durationMs=${durationMs}`);
      return null;
    }

    const validatedDirection = validateArtDirection(parsed as Record<string, unknown>, slides.length);
    if (!validatedDirection) {
      console.warn(`[ppt-b4] art-director: validation failed durationMs=${durationMs}`);
      return null;
    }
    const direction = enforceCompositionDiversity(
      applyNarrativeSignals(ensurePhase3Selection(validatedDirection, slides), slides),
      slides
    );

    console.log(`[ppt-b4] direction generated durationMs=${durationMs} motif=${direction.motif} slides=${direction.slides.length}`);
    const phase3Selected = direction.slides
      .filter((slide) => [
        "split_visual", "comparison_matrix", "decision_summary",
        "editorial_statement", "asymmetric_list",
      ].includes(slide.layoutType))
      .map((slide) => `${slide.slideIndex}:${slide.layoutType}`);
    console.log(`[ppt-b4] phase3 selected=${phase3Selected.length > 0 ? phase3Selected.join(",") : "none"}`);
    return direction;
  } catch (err: unknown) {
    console.warn("[ppt-b4] art-director: API error:", (err as Error)?.message ?? err);
    return null;
  }
}

// ─── JSON検証 ────────────────────────────────────────────────────────────────

function validateArtDirection(
  raw: Record<string, unknown>,
  expectedSlideCount: number
): DeckArtDirection | null {
  if (raw.version !== 1) return null;
  if (typeof raw.designThesis !== "string" || !raw.designThesis) return null;
  if (!["rounded-cards", "icon-circles", "editorial-type", "data-first"].includes(raw.motif as string)) return null;
  if (!["light", "dark", "mixed"].includes(raw.dominantTone as string)) return null;
  if (!Array.isArray(raw.slides)) return null;
  if (raw.slides.length !== expectedSlideCount) return null;

  const VALID_ROLES: ReadonlySet<string> = new Set(["opening", "context", "evidence", "comparison", "process", "decision", "closing"]);
  const VALID_DOMINANT: ReadonlySet<string> = new Set(["headline", "metric", "chart", "comparison", "process", "cards", "table", "illustration"]);
  const VALID_DENSITY: ReadonlySet<string> = new Set(["low", "medium", "high"]);

  const seenIndices = new Set<number>();
  const slides: ArtDirectedSlide[] = [];

  for (const item of raw.slides as unknown[]) {
    if (!item || typeof item !== "object") return null;
    const s = item as Record<string, unknown>;

    const slideIndex = typeof s.slideIndex === "number" ? s.slideIndex : -1;
    if (slideIndex < 0 || slideIndex >= expectedSlideCount) return null;
    if (seenIndices.has(slideIndex)) return null;
    seenIndices.add(slideIndex);

    if (!VALID_ROLES.has(s.visualRole as string)) return null;
    if (!ALLOWED_LAYOUT_TYPES.has(s.layoutType as ArtDirectedSlide["layoutType"])) return null;
    if (!VALID_DOMINANT.has(s.dominantElement as string)) return null;
    if (!VALID_DENSITY.has(s.density as string)) return null;

    const budget = s.contentBudget as Record<string, unknown> | undefined;
    if (
      !budget ||
      typeof budget.maxItems !== "number" || !Number.isFinite(budget.maxItems) || !Number.isInteger(budget.maxItems) ||
      budget.maxItems < 1 || budget.maxItems > 20 ||
      typeof budget.maxCharsPerItem !== "number" || !Number.isFinite(budget.maxCharsPerItem) || !Number.isInteger(budget.maxCharsPerItem) ||
      budget.maxCharsPerItem < 10 || budget.maxCharsPerItem > 200
    ) return null;

    const indices = s.sourceBulletIndices;
    if (!Array.isArray(indices) || indices.length === 0) return null;
    if (!indices.every((idx) => typeof idx === "number" && Number.isInteger(idx) && idx >= 0)) return null;

    if (typeof s.visualReason !== "string") return null;

    slides.push({
      slideIndex,
      visualRole: s.visualRole as SlideVisualRole,
      layoutType: s.layoutType as ArtDirectedSlide["layoutType"],
      dominantElement: s.dominantElement as DominantElement,
      density: s.density as VisualDensity,
      contentBudget: { maxItems: budget.maxItems as number, maxCharsPerItem: budget.maxCharsPerItem as number },
      sourceBulletIndices: indices as number[],
      visualReason: s.visualReason as string,
    });
  }

  // Verify all indices 0..N-1 are covered
  for (let i = 0; i < expectedSlideCount; i++) {
    if (!seenIndices.has(i)) return null;
  }

  return {
    version: 1,
    designThesis: raw.designThesis as string,
    motif: raw.motif as DeckArtDirection["motif"],
    dominantTone: raw.dominantTone as DeckArtDirection["dominantTone"],
    slides: slides.sort((a, b) => a.slideIndex - b.slideIndex),
  };
}

