type CompanyProfileSlideLike = {
  title?: string;
  layoutType?: string;
  bullets?: unknown[];
  columns?: unknown[];
  tableRows?: unknown[];
  metrics?: unknown[];
  steps?: unknown[];
  benefits?: unknown[];
  cards?: unknown[];
  statCallouts?: unknown[];
  leadText?: string;
  callout?: unknown;
};

export function fitCompanyProfileSlideCount<T extends CompanyProfileSlideLike>(
  slides: readonly T[],
  deckTitle: string,
  targetContentSlides: number
): T[] {
  if (!Number.isInteger(targetContentSlides) || targetContentSlides < 1) {
    return [];
  }

  const contentSlides = slides.filter(
    (slide, index) => index !== 0 || !isCoverSlide(slide, deckTitle)
  );
  if (contentSlides.length <= targetContentSlides) return [...contentSlides];
  if (targetContentSlides === 1) return contentSlides.slice(0, 1);

  // Keep the final closing/next-step slide when the model over-produces.
  return [
    ...contentSlides.slice(0, targetContentSlides - 1),
    contentSlides[contentSlides.length - 1],
  ];
}

function isCoverSlide(
  slide: CompanyProfileSlideLike,
  deckTitle: string
): boolean {
  if (slide.layoutType === "title") return true;
  if (/^(?:表紙|タイトル|cover|title slide)$/i.test(slide.title?.trim() ?? "")) {
    return true;
  }

  const normalizedSlideTitle = normalizeTitle(slide.title ?? "");
  const normalizedDeckTitle = normalizeTitle(deckTitle);
  return (
    normalizedSlideTitle.length > 0 &&
    normalizedSlideTitle === normalizedDeckTitle &&
    !hasBody(slide)
  );
}

function hasBody(slide: CompanyProfileSlideLike): boolean {
  return (
    Boolean(slide.bullets?.length) ||
    Boolean(slide.columns?.length) ||
    Boolean(slide.tableRows?.length) ||
    Boolean(slide.metrics?.length) ||
    Boolean(slide.steps?.length) ||
    Boolean(slide.benefits?.length) ||
    Boolean(slide.cards?.length) ||
    Boolean(slide.statCallouts?.length) ||
    Boolean(slide.leadText?.trim()) ||
    Boolean(slide.callout)
  );
}

function normalizeTitle(value: string): string {
  return value.normalize("NFKC").replace(/[\s　]/g, "").toLowerCase();
}
