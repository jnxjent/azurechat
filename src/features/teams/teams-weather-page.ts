const TRUSTED_WEATHER_HOSTS = [
  "jma.go.jp",
  "yahoo.co.jp",
  "tenki.jp",
  "toshin.com",
  "weathernews.jp",
] as const;

const WEATHER_KEYWORDS = [
  "明日",
  "あす",
  "天気",
  "最高気温",
  "最低気温",
  "降水確率",
  "℃",
] as const;

export function isTrustedWeatherUrl(value: string): boolean {
  try {
    const url = new URL(value);
    if (url.protocol !== "https:") return false;
    const host = url.hostname.toLowerCase();
    return TRUSTED_WEATHER_HOSTS.some(
      (allowed) => host === allowed || host.endsWith(`.${allowed}`)
    );
  } catch {
    return false;
  }
}

export function extractWeatherPageText(html: string): string {
  const visibleText = decodeHtmlEntities(
    html
      .replace(/<(script|style|noscript|svg)\b[^>]*>[\s\S]*?<\/\1>/gi, " ")
      .replace(/<(?:br|\/p|\/div|\/li|\/tr|\/h[1-6])\s*>/gi, "\n")
      .replace(/<[^>]+>/g, " ")
  )
    .replace(/[\t\r ]+/g, " ")
    .replace(/\n\s+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  if (!visibleText) return "";

  const windows: string[] = [];
  for (const keyword of WEATHER_KEYWORDS) {
    let offset = 0;
    while (windows.length < 8) {
      const index = visibleText.indexOf(keyword, offset);
      if (index < 0) break;
      const start = Math.max(0, index - 220);
      const end = Math.min(visibleText.length, index + 700);
      const excerpt = visibleText.slice(start, end).trim();
      if (
        excerpt &&
        !windows.some((existing) =>
          existing.includes(excerpt.slice(0, Math.min(120, excerpt.length)))
        )
      ) {
        windows.push(excerpt);
      }
      offset = index + keyword.length;
    }
    if (windows.length >= 8) break;
  }

  return (windows.length > 0 ? windows.join("\n...\n") : visibleText).slice(
    0,
    3_000
  );
}

function decodeHtmlEntities(value: string): string {
  const named: Record<string, string> = {
    amp: "&",
    apos: "'",
    gt: ">",
    lt: "<",
    nbsp: " ",
    quot: '"',
  };
  return value.replace(
    /&(#x[0-9a-f]+|#\d+|[a-z]+);/gi,
    (entity, code: string) => {
      if (code[0] !== "#") return named[code.toLowerCase()] ?? entity;
      const radix = code[1]?.toLowerCase() === "x" ? 16 : 10;
      const digits = radix === 16 ? code.slice(2) : code.slice(1);
      const point = Number.parseInt(digits, radix);
      return Number.isFinite(point) ? String.fromCodePoint(point) : entity;
    }
  );
}
