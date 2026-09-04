/**
 * B4 Preflight — Art DirectionのJSONをAIとは独立して検査・修復する。
 * AIの文字列理由から構造データを抽出しない。
 * 元スライドの既存フィールドから決定的に修復できる場合のみ補完し、
 * それ以外は元レイアウトへ戻す。
 */

import { ALLOWED_LAYOUT_TYPES } from "./art-director";
import type { DeckArtDirection, ArtDirectedSlide } from "./art-director";

// 循環importを避けるため gen-pptx/route から import せず最小型定義を使用
type SlideInput = {
  layoutType?: string;
  tableRows?: unknown[][];
  columns?: unknown[];
  visualBlocks?: unknown[];
  connectors?: unknown[];
  leadText?: string;
  metrics?: unknown[];
  steps?: unknown[];
  statCallouts?: unknown[];
  cards?: unknown[];
  bullets?: string[];
  decisionConclusion?: string;
  decisionEvidence?: string[];
  decisionActions?: string[];
};

// ─── 型定義 ───────────────────────────────────────────────────────────────────

export type PreflightSeverity = "info" | "warning" | "critical";

export type PreflightIssue = {
  code: string;
  severity: PreflightSeverity;
  slideIndex?: number;
  message: string;
};

export type PreflightResult = {
  accepted: boolean;
  direction: DeckArtDirection;
  issues: PreflightIssue[];
  repaired: boolean;
};

// ─── レイアウト必須データ ────────────────────────────────────────────────────

type LayoutRequirement = {
  check: (slide: SlideInput) => boolean;
  safeFallback: ArtDirectedSlide["layoutType"];
};

function resolveSafeFallback(
  sourceSlide: SlideInput,
  requestedLayout: ArtDirectedSlide["layoutType"],
  defaultFallback: ArtDirectedSlide["layoutType"]
): ArtDirectedSlide["layoutType"] {
  const originalLayout = sourceSlide.layoutType as ArtDirectedSlide["layoutType"] | undefined;
  if (
    originalLayout &&
    originalLayout !== requestedLayout &&
    ALLOWED_LAYOUT_TYPES.has(originalLayout)
  ) {
    const originalRequirement = LAYOUT_REQUIREMENTS[originalLayout];
    if (!originalRequirement || originalRequirement.check(sourceSlide)) return originalLayout;
  }
  return defaultFallback;
}

function hasExplicitDecisionStructure(slide: SlideInput): boolean {
  if (
    (slide.bullets ?? []).filter((v) => v?.trim()).length === 0 &&
    slide.decisionConclusion?.trim() &&
    (slide.decisionEvidence ?? []).filter((v) => v?.trim()).length >= 1 &&
    (slide.decisionEvidence ?? []).filter((v) => v?.trim()).length <= 3 &&
    (slide.decisionActions ?? []).filter((v) => v?.trim()).length >= 1 &&
    (slide.decisionActions ?? []).filter((v) => v?.trim()).length <= 3
  ) return true;

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

function hasBoundedSplitVisualStructure(slide: SlideInput): boolean {
  const bulletCount = (slide.bullets ?? []).filter((b) => b?.trim()).length;
  const visualCount = Math.max(
    Array.isArray(slide.metrics) ? slide.metrics.length : 0,
    Array.isArray(slide.statCallouts) ? slide.statCallouts.length : 0,
    Array.isArray(slide.cards) ? slide.cards.length : 0,
    Array.isArray(slide.visualBlocks) ? slide.visualBlocks.length : 0
  );
  return bulletCount >= 1 && bulletCount <= 5 && visualCount >= 1 && visualCount <= 4;
}

function hasComparisonStructure(slide: SlideInput): boolean {
  const sourceBullets = (slide.bullets ?? []).map((b) => b.trim()).filter(Boolean);
  if (Array.isArray(slide.columns) && slide.columns.length >= 2 && slide.columns.length <= 3) {
    const represented = slide.columns.flatMap((rawColumn) => {
      if (!rawColumn || typeof rawColumn !== "object") return [];
      const column = rawColumn as { bullets?: unknown[] };
      return (column.bullets ?? []).map((value) => String(value).trim());
    });
    const itemCountsAreSafe = slide.columns.every((rawColumn) => {
      if (!rawColumn || typeof rawColumn !== "object") return false;
      const column = rawColumn as { bullets?: unknown[] };
      return (column.bullets ?? []).filter((value) => String(value).trim()).length <= 6;
    });
    return itemCountsAreSafe && sourceBullets.every((bullet) => represented.includes(bullet));
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

const LAYOUT_REQUIREMENTS: Partial<Record<ArtDirectedSlide["layoutType"], LayoutRequirement>> = {
  table: {
    check: (s) => Array.isArray(s.tableRows) && s.tableRows.length >= 2 && s.tableRows[0]?.length > 0,
    safeFallback: "bullets",
  },
  "multi-column": {
    // bullets が2件以上あれば派生可能
    check: (s) =>
      (Array.isArray(s.columns) && s.columns.length >= 2) ||
      (Array.isArray(s.bullets) && s.bullets.length >= 2),
    safeFallback: "bullets",
  },
  diagram: {
    check: (s) => Array.isArray(s.visualBlocks) && s.visualBlocks.length > 0 && Array.isArray(s.connectors),
    safeFallback: "bullets",
  },
  "company-overview": {
    check: (s) => !!(s.leadText || (Array.isArray(s.metrics) && s.metrics.length > 0)),
    safeFallback: "bullets",
  },
  "process-cards": {
    // bullets が2件以上あれば派生可能
    check: (s) =>
      (Array.isArray(s.steps) && s.steps.length >= 2) ||
      (Array.isArray(s.bullets) && s.bullets.length >= 2),
    safeFallback: "bullets",
  },
  timeline: {
    // bullets が2件以上あれば派生可能
    check: (s) =>
      (Array.isArray(s.steps) && s.steps.length >= 2) ||
      (Array.isArray(s.bullets) && s.bullets.length >= 2),
    safeFallback: "bullets",
  },
  "metric-cards": {
    check: (s) => Array.isArray(s.metrics) && s.metrics.length >= 1,
    safeFallback: "bullets",
  },
  stat_callouts: {
    check: (s) =>
      (Array.isArray(s.statCallouts) && s.statCallouts.length >= 1) ||
      (Array.isArray(s.metrics) && s.metrics.length >= 1),
    safeFallback: "bullets",
  },
  card_grid: {
    // bullets が2件以上あれば派生可能
    check: (s) =>
      (Array.isArray(s.cards) && s.cards.length >= 2) ||
      (Array.isArray(s.steps) && s.steps.length >= 2) ||
      (Array.isArray(s.bullets) && s.bullets.length >= 2),
    safeFallback: "bullets",
  },
  icon_rows: {
    // bullets が2件以上あれば派生可能
    check: (s) =>
      (Array.isArray(s.cards) && s.cards.length >= 2) ||
      (Array.isArray(s.steps) && s.steps.length >= 2) ||
      (Array.isArray(s.bullets) && s.bullets.length >= 2),
    safeFallback: "bullets",
  },
  split_visual: {
    check: hasBoundedSplitVisualStructure,
    safeFallback: "bullets",
  },
  comparison_matrix: {
    check: hasComparisonStructure,
    safeFallback: "bullets",
  },
  decision_summary: {
    check: hasExplicitDecisionStructure,
    safeFallback: "bullets",
  },
  editorial_statement: {
    check: (s) => {
      const count = (s.bullets ?? []).filter((bullet) => bullet?.trim()).length;
      return count >= 1 && count <= 3;
    },
    safeFallback: "bullets",
  },
  asymmetric_list: {
    check: (s) => {
      const count = (s.bullets ?? []).filter((bullet) => bullet?.trim()).length;
      return count >= 2 && count <= 6;
    },
    safeFallback: "bullets",
  },
};

// ─── Preflight本体 ────────────────────────────────────────────────────────────

export function runPreflight(
  direction: DeckArtDirection,
  slides: SlideInput[]
): PreflightResult {
  const issues: PreflightIssue[] = [];
  let repaired = false;

  // 1. スライド数一致
  if (direction.slides.length !== slides.length) {
    issues.push({
      code: "slide-count-mismatch",
      severity: "critical",
      message: `Art Direction has ${direction.slides.length} slides, expected ${slides.length}`,
    });
    return { accepted: false, direction, issues, repaired: false };
  }

  // 2. index 欠番・重複・範囲外
  const seenIdx = new Set<number>();
  for (const ds of direction.slides) {
    if (ds.slideIndex < 0 || ds.slideIndex >= slides.length) {
      issues.push({ code: "index-mismatch", severity: "critical", slideIndex: ds.slideIndex, message: `slideIndex ${ds.slideIndex} out of range` });
      return { accepted: false, direction, issues, repaired: false };
    }
    if (seenIdx.has(ds.slideIndex)) {
      issues.push({ code: "index-mismatch", severity: "critical", slideIndex: ds.slideIndex, message: `duplicate slideIndex ${ds.slideIndex}` });
      return { accepted: false, direction, issues, repaired: false };
    }
    seenIdx.add(ds.slideIndex);
  }

  // 3. source bullet 参照カバレッジ（0..bulletCount-1 を全て参照・重複なし・範囲内整数のみ）
  for (const ds of direction.slides) {
    const src = slides[ds.slideIndex];
    const bulletCount = (src?.bullets ?? []).length;
    if (bulletCount === 0) continue;

    if (!ds.sourceBulletIndices.every((i) => Number.isInteger(i) && i >= 0)) {
      issues.push({ code: "content-loss", severity: "critical", slideIndex: ds.slideIndex, message: `slide ${ds.slideIndex}: non-integer or negative sourceBulletIndices` });
      return { accepted: false, direction, issues, repaired: false };
    }
    if (ds.sourceBulletIndices.some((i) => i >= bulletCount)) {
      issues.push({ code: "content-loss", severity: "critical", slideIndex: ds.slideIndex, message: `slide ${ds.slideIndex}: sourceBulletIndices out of range (bulletCount=${bulletCount})` });
      return { accepted: false, direction, issues, repaired: false };
    }
    const seenBullets = new Set(ds.sourceBulletIndices);
    if (seenBullets.size !== ds.sourceBulletIndices.length) {
      issues.push({ code: "content-loss", severity: "critical", slideIndex: ds.slideIndex, message: `slide ${ds.slideIndex}: duplicate entries in sourceBulletIndices` });
      return { accepted: false, direction, issues, repaired: false };
    }
    for (let bi = 0; bi < bulletCount; bi++) {
      if (!seenBullets.has(bi)) {
        issues.push({ code: "content-loss", severity: "critical", slideIndex: ds.slideIndex, message: `slide ${ds.slideIndex}: bullet [${bi}] not referenced in sourceBulletIndices` });
        return { accepted: false, direction, issues, repaired: false };
      }
    }
  }

  // 4. 各スライドの個別検査（修復可能）
  const repairedSlides: ArtDirectedSlide[] = direction.slides.map((ds) => {
    let current = { ...ds };

    // 4a. 許可外レイアウト → 元のレイアウトに戻す
    if (!ALLOWED_LAYOUT_TYPES.has(current.layoutType)) {
      const originalLayout = (slides[current.slideIndex]?.layoutType ?? "bullets") as ArtDirectedSlide["layoutType"];
      const fallback = ALLOWED_LAYOUT_TYPES.has(originalLayout) ? originalLayout : "bullets";
      issues.push({
        code: "unsupported-layout",
        severity: "warning",
        slideIndex: current.slideIndex,
        message: `Layout "${current.layoutType}" not allowed → fallback to "${fallback}"`,
      });
      current = { ...current, layoutType: fallback };
      repaired = true;
    }

    // 4b. 必須データ不足 → safeFallback
    const req = LAYOUT_REQUIREMENTS[current.layoutType];
    if (req) {
      const srcSlide = slides[current.slideIndex];
      if (!req.check(srcSlide)) {
        const fallback = resolveSafeFallback(srcSlide, current.layoutType, req.safeFallback);
        issues.push({
          code: "missing-structure",
          severity: "warning",
          slideIndex: current.slideIndex,
          message: `Layout "${current.layoutType}" lacks required data → fallback to "${fallback}"`,
        });
        current = { ...current, layoutType: fallback };
        repaired = true;
      }
    }

    // 4c. dominantElement=illustration で根拠なし → cards に戻す
    if (current.dominantElement === "illustration") {
      issues.push({
        code: "decorative-only",
        severity: "warning",
        slideIndex: current.slideIndex,
        message: "illustration dominantElement rejected — using cards",
      });
      current = { ...current, dominantElement: "cards" };
      repaired = true;
    }

    return current;
  });

  // 5. 同一レイアウト連続制限（修復後に構造検査・連続検査を再実行してループ収束）
  const maxStreak = parseInt(process.env.PPTX_ART_DIRECTOR_MAX_LAYOUT_STREAK ?? "2", 10) || 2;
  let streakSlides = repairedSlides;
  for (let streakIter = 0; streakIter < 5; streakIter++) {
    const streakResult = breakLayoutStreak(streakSlides, maxStreak, issues, slides);
    if (!streakResult.changed) break;
    streakSlides = streakResult.slides;
    repaired = true;
  }
  const finalSlides = streakSlides;

  // 6. 1-1マッピング検証：各slideIndexは必ず1件のArtDirectedSlideにのみ対応（step 2で保証済み）
  // 同一sourceSlideを複数のArtDirectedSlideが参照することは構造上起きない。

  const repairedDirection: DeckArtDirection = { ...direction, slides: finalSlides };
  const critical = issues.filter((i) => i.severity === "critical").length;

  console.log(
    `[ppt-b4] preflight accepted=${critical === 0} repaired=${repaired} ` +
    `warnings=${issues.filter((i) => i.severity === "warning").length} critical=${critical}`
  );
  for (const issue of issues) {
    if (issue.slideIndex !== undefined) {
      console.log(`[ppt-b4] slide=${issue.slideIndex} rejected=${issue.code} reason=${issue.message}`);
    }
  }

  return {
    accepted: critical === 0,
    direction: repairedDirection,
    issues,
    repaired,
  };
}

// ─── レイアウト連続修復 ────────────────────────────────────────────────────────

function breakLayoutStreak(
  slides: ArtDirectedSlide[],
  maxStreak: number,
  issues: PreflightIssue[],
  srcSlides: SlideInput[]
): { slides: ArtDirectedSlide[]; changed: boolean } {
  if (slides.length === 0 || maxStreak < 1) return { slides, changed: false };
  const result = [...slides];
  let changed = false;

  let streakStart = 0;
  for (let i = 1; i <= result.length; i++) {
    const sameLayout = i < result.length && result[i].layoutType === result[streakStart].layoutType;
    if (!sameLayout || i === result.length) {
      const streakLen = i - streakStart;
      if (streakLen > maxStreak) {
        for (let j = streakStart + maxStreak; j < i; j++) {
          const original = result[j].layoutType;
          const srcSlide = srcSlides[result[j].slideIndex];
          const alt = pickAlternativeLayout(original, srcSlide);
          result[j] = { ...result[j], layoutType: alt };
          issues.push({
            code: "layout-streak",
            severity: "warning",
            slideIndex: result[j].slideIndex,
            message: `layout-streak: "${original}" streak=${streakLen} → changed to "${alt}"`,
          });
          changed = true;
        }
      }
      streakStart = i;
    }
  }

  return { slides: result, changed };
}

function pickAlternativeLayout(
  blocked: ArtDirectedSlide["layoutType"],
  srcSlide: SlideInput
): ArtDirectedSlide["layoutType"] {
  // 構造要件を満たすレイアウトのみ候補にする
  const candidates: ArtDirectedSlide["layoutType"][] = [
    "editorial_statement", "asymmetric_list", "bullets", "multi-column", "stat_callouts", "timeline",
    "split_visual", "comparison_matrix", "decision_summary", "card_grid", "icon_rows",
  ];
  for (const alt of candidates) {
    if (alt === blocked) continue;
    const req = LAYOUT_REQUIREMENTS[alt];
    if (!req || req.check(srcSlide)) return alt;
  }
  return "bullets";
}
