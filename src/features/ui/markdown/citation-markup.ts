export type CitationMarkupItem = {
  name: string;
  id: string;
};

// The closing `}` is optional because some models occasionally emit `/%`.
// Whitespace around `items`, `=`, `/`, and `%` is accepted as well.
export const CITATION_MARKUP_RE =
  /\{%\s*citation\s+items\s*=\s*\[([\s\S]*?)\]\s*\/\s*%\}?/gi;

function decodeQuotedValue(value: string): string {
  try {
    return JSON.parse(`"${value.replace(/"/g, '\\"')}"`);
  } catch {
    return value.replace(/\\([\\"'])/g, "$1");
  }
}

function normalizeItem(value: unknown): CitationMarkupItem | null {
  if (!value || typeof value !== "object") return null;
  const item = value as Record<string, unknown>;
  const name = typeof item.name === "string" ? item.name.trim() : "";
  const id = typeof item.id === "string" ? item.id.trim() : "";
  return name && id ? { name, id } : null;
}

/** Parse both strict JSON and the legacy Markdoc object syntax. */
export function parseCitationItems(inner: string): CitationMarkupItem[] {
  const source = String(inner ?? "").trim();
  if (!source) return [];

  try {
    const parsed = JSON.parse(`[${source}]`);
    if (Array.isArray(parsed)) {
      const items = parsed
        .map(normalizeItem)
        .filter((item): item is CitationMarkupItem => Boolean(item));
      if (items.length) return items;
    }
  } catch {
    // Legacy syntax has unquoted keys, so fall through to the tolerant parser.
  }

  const items: CitationMarkupItem[] = [];
  const objectRe = /\{([^{}]*)\}/g;
  let objectMatch: RegExpExecArray | null;
  while ((objectMatch = objectRe.exec(source)) !== null) {
    const values: Partial<CitationMarkupItem> = {};
    const fieldRe =
      /["']?(name|id)["']?\s*:\s*(?:"((?:\\.|[^"\\])*)"|'((?:\\.|[^'\\])*)')/gi;
    let fieldMatch: RegExpExecArray | null;
    while ((fieldMatch = fieldRe.exec(objectMatch[1])) !== null) {
      const key = fieldMatch[1].toLowerCase() as keyof CitationMarkupItem;
      values[key] = decodeQuotedValue(fieldMatch[2] ?? fieldMatch[3] ?? "");
    }
    const normalized = normalizeItem(values);
    if (normalized) items.push(normalized);
  }

  return items;
}

export function extractCitationItems(text: string): CitationMarkupItem[] {
  const items: CitationMarkupItem[] = [];
  const re = new RegExp(CITATION_MARKUP_RE.source, CITATION_MARKUP_RE.flags);
  let match: RegExpExecArray | null;
  while ((match = re.exec(text)) !== null) {
    items.push(...parseCitationItems(match[1]));
  }
  return items;
}

function escapeCitationValue(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

export function formatCitationMarkup(items: CitationMarkupItem[]): string {
  const unique = Array.from(
    new Map(
      items
        .filter((item) => item.name.trim() && item.id.trim())
        .map((item) => [item.id.trim(), { name: item.name.trim(), id: item.id.trim() }])
    ).values()
  );
  const inner = unique
    .map(
      ({ name, id }) =>
        `{name:"${escapeCitationValue(name)}",id:"${escapeCitationValue(id)}"}`
    )
    .join(", ");
  return inner ? `{% citation items=[${inner}] /%}` : "";
}

export function removeCitationMarkup(text: string): string {
  const re = new RegExp(CITATION_MARKUP_RE.source, CITATION_MARKUP_RE.flags);
  return text
    .replace(re, "")
    // Citation is required at the end of an answer. If a model truncated the
    // closing `] /%}`, remove that incomplete tail before appending canonical markup.
    .replace(/\{%\s*citation\b[\s\S]*$/gi, "")
    .replace(/\n{3,}/g, "\n\n")
    .trimEnd();
}

