export type TeamsWebQuery = {
  query: string;
  enriched: boolean;
};

/** Enrich weather queries with an unambiguous JST date and forecast fields. */
export function buildTeamsWebQuery(
  message: string,
  now: Date = new Date()
): TeamsWebQuery {
  const trimmed = message.trim();
  if (!/(?:天気|天候|気温|降水|雨|雪|台風)/i.test(trimmed)) {
    return { query: trimmed, enriched: false };
  }

  const date = resolveRequestedWeatherDate(trimmed, now);
  const dateText = date ? ` ${date}` : "";
  return {
    query: `${trimmed}${dateText} 天気予報 最高気温 最低気温 降水確率`,
    enriched: true,
  };
}

function resolveRequestedWeatherDate(message: string, now: Date): string | null {
  if (!/(?:今日|本日|明日|あす|明後日|あさって)/.test(message)) {
    return null;
  }

  const jstParts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Tokyo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(now);
  const values = Object.fromEntries(
    jstParts.map((part) => [part.type, part.value])
  );
  const base = new Date(
    Date.UTC(Number(values.year), Number(values.month) - 1, Number(values.day))
  );
  const offset = /(?:明後日|あさって)/.test(message)
    ? 2
    : /(?:明日|あす)/.test(message)
      ? 1
      : 0;
  base.setUTCDate(base.getUTCDate() + offset);
  return base.toISOString().slice(0, 10);
}
