import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";

export type TeamsPptSlide = {
  title: string;
  bullets: string[];
  layoutType?: "title" | "bullets" | "multi-column" | "closing";
  columns?: Array<{ header: string; bullets: string[] }>;
};

export type TeamsPptExtractedSlide = {
  slideIndex: number;
  title: string;
  bullets: string[];
};

export type TeamsPptCardEdit = {
  slideIndex: number;
  convertToCards: {
    cards: Array<{ iconKey: string; heading: string; body: string }>;
  };
};

export async function createTeamsPptPlan(props: {
  prompt: string;
  title: string;
  referenceContext?: string;
}): Promise<{ title: string; slides: TeamsPptSlide[] }> {
  const openai = OpenAIInstance();
  const response = await openai.chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      {
        role: "system",
        content: [
          "You create concise Japanese presentation outlines.",
          "Return JSON only with this schema:",
          '{"title":"資料タイトル","slides":[{"title":"スライドタイトル","bullets":["要点"],"layoutType":"title|bullets|multi-column|closing","columns":[{"header":"列見出し","bullets":["要点"]}]}]}',
          "Create 5 to 8 slides unless the user specifies otherwise.",
          "When referenceContext is provided, use it as the factual source and do not supplement it with general web knowledge.",
          "Use only information supplied by the user or referenceContext; clearly label any proposed ideas as suggestions.",
          "The first slide should be a title slide and the last may be a closing slide.",
          "For each non-title slide, write 4 to 7 substantive bullets whenever the referenceContext supports them.",
          "Each substantive Japanese bullet should normally contain 35 to 70 characters and include concrete facts, activities, results, issues, or next actions.",
          "Do not reduce a source-rich section to only one or two short phrases.",
          "When quarterly materials are provided, cover every requested quarter and make changes across quarters understandable.",
          "Include specific names, figures, systems, and outcomes found in referenceContext; never invent missing facts.",
        ].join("\n"),
      },
      {
        role: "user",
        content: JSON.stringify({
          title: props.title,
          request: props.prompt,
          referenceContext: props.referenceContext,
        }),
      },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 8000,
  });

  const raw = response.choices[0]?.message?.content ?? "{}";
  const parsed = JSON.parse(raw) as {
    title?: unknown;
    slides?: unknown;
  };
  const slides = Array.isArray(parsed.slides)
    ? parsed.slides
        .map(normalizeSlide)
        .filter((slide): slide is TeamsPptSlide => slide !== null)
    : [];

  if (slides.length === 0) {
    throw new Error("PowerPointの構成を作成できませんでした。");
  }

  return {
    title:
      typeof parsed.title === "string" && parsed.title.trim()
        ? parsed.title.trim()
        : props.title,
    slides,
  };
}

export async function createTeamsPptCardEdits(props: {
  slides: TeamsPptExtractedSlide[];
  targetItemCount: number;
  instruction: string;
}): Promise<TeamsPptCardEdit[]> {
  const openai = OpenAIInstance();
  const response = await openai.chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      {
        role: "system",
        content: [
          "Convert existing PowerPoint slide bullets into business card items.",
          "Return JSON only: {\"slides\":[{\"slideIndex\":2,\"cards\":[{\"iconKey\":\"gear\",\"heading\":\"見出し\",\"body\":\"本文\"}]}]}",
          `Return exactly ${props.targetItemCount} cards for every supplied slide.`,
          "Preserve all important facts by consolidating or splitting the existing bullets.",
          "Do not introduce facts that do not appear in the supplied slide.",
          "Use concise Japanese headings and substantive Japanese bodies.",
          "Keep slideIndex unchanged and return every supplied slide exactly once.",
        ].join("\n"),
      },
      {
        role: "user",
        content: JSON.stringify({
          instruction: props.instruction,
          slides: props.slides,
        }),
      },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 5000,
  });

  const raw = response.choices[0]?.message?.content ?? "{}";
  const parsed = JSON.parse(raw) as { slides?: unknown };
  const returnedSlides = Array.isArray(parsed.slides) ? parsed.slides : [];
  const byIndex = new Map<number, TeamsPptCardEdit>();

  for (const value of returnedSlides) {
    if (!value || typeof value !== "object") continue;
    const source = value as Record<string, unknown>;
    const slideIndex = Number(source.slideIndex);
    if (!Number.isInteger(slideIndex)) continue;
    const cards = Array.isArray(source.cards)
      ? source.cards
          .map(normalizeCard)
          .filter(
            (
              card
            ): card is { iconKey: string; heading: string; body: string } =>
              card !== null
          )
          .slice(0, props.targetItemCount)
      : [];
    if (cards.length === props.targetItemCount) {
      byIndex.set(slideIndex, {
        slideIndex,
        convertToCards: { cards },
      });
    }
  }

  const edits = props.slides
    .map((slide) => byIndex.get(slide.slideIndex))
    .filter((edit): edit is TeamsPptCardEdit => Boolean(edit));
  if (edits.length !== props.slides.length) {
    throw new Error(
      `指定した全スライドを${props.targetItemCount}項目のカードに変換できませんでした。`
    );
  }
  return edits;
}

function normalizeSlide(value: unknown): TeamsPptSlide | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  if (typeof source.title !== "string" || !source.title.trim()) return null;

  const bullets = Array.isArray(source.bullets)
    ? source.bullets.filter((item): item is string => typeof item === "string")
    : [];
  const columns = Array.isArray(source.columns)
    ? source.columns
        .map((column) => {
          if (!column || typeof column !== "object") return null;
          const item = column as Record<string, unknown>;
          if (typeof item.header !== "string") return null;
          return {
            header: item.header,
            bullets: Array.isArray(item.bullets)
              ? item.bullets.filter(
                  (bullet): bullet is string => typeof bullet === "string"
                )
              : [],
          };
        })
        .filter(
          (
            column
          ): column is { header: string; bullets: string[] } => column !== null
        )
    : undefined;
  const layoutType = ["title", "bullets", "multi-column", "closing"].includes(
    String(source.layoutType)
  )
    ? (source.layoutType as TeamsPptSlide["layoutType"])
    : undefined;

  return {
    title: source.title.trim(),
    bullets,
    layoutType,
    ...(columns?.length ? { columns } : {}),
  };
}

function normalizeCard(
  value: unknown
): { iconKey: string; heading: string; body: string } | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  const heading = typeof source.heading === "string" ? source.heading.trim() : "";
  const body = typeof source.body === "string" ? source.body.trim() : "";
  if (!heading && !body) return null;
  return {
    iconKey:
      typeof source.iconKey === "string" && source.iconKey.trim()
        ? source.iconKey.trim().slice(0, 24)
        : "gear",
    heading: heading.slice(0, 40),
    body: body.slice(0, 180),
  };
}
