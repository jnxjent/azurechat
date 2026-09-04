export function normalizeSalesforceEmail(email: unknown): string {
  return typeof email === "string" ? email.trim().toLowerCase() : "";
}

export function parseSalesforceWhitelist(raw: unknown): string[] {
  if (typeof raw !== "string") return [];
  return Array.from(
    new Set(
      raw
        .split(/[,\n\r]/)
        .map(normalizeSalesforceEmail)
        .filter(Boolean)
    )
  );
}

export function isSalesforceAllowedEmail(
  email: unknown,
  rawWhitelist: unknown = process.env.SF_WHITELIST_EMAILS
): boolean {
  const normalized = normalizeSalesforceEmail(email);
  return (
    normalized.length > 0 &&
    parseSalesforceWhitelist(rawWhitelist).includes(normalized)
  );
}
