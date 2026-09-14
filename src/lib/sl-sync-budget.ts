export type SlSyncAttempt = { count: number; succeeded: boolean; updatedAt: string };
export type SlSyncLedger = {
  version: 1;
  day: string;
  month: string;
  dailyPages: number;
  monthlyPages: number;
  attempts: Record<string, SlSyncAttempt>;
};

export type SlSyncLimits = { daily: number; monthly: number; perFile: number; attempts: number };

export function rolloverSlSyncLedger(ledger: SlSyncLedger, day: string, month: string, now: Date): void {
  if (ledger.month !== month) {
    ledger.month = month;
    ledger.monthlyPages = 0;
    const cutoff = now.getTime() - 90 * 86400000;
    for (const [id, attempt] of Object.entries(ledger.attempts)) {
      if (Date.parse(attempt.updatedAt) < cutoff) delete ledger.attempts[id];
    }
  }
  if (ledger.day !== day) {
    ledger.day = day;
    ledger.dailyPages = 0;
  }
}

export function reserveSlSyncBudget(
  ledger: SlSyncLedger,
  id: string,
  pages: number,
  max: SlSyncLimits,
  now: Date
): string | null {
  if (!Number.isSafeInteger(pages) || pages < 0) return "page_count_unknown";
  if (pages > max.perFile) return "file_page_limit";
  const previous = ledger.attempts[id];
  if (previous?.succeeded) return "already_indexed_this_version";
  if ((previous?.count ?? 0) >= max.attempts) return "attempt_limit";
  if (ledger.dailyPages + pages > max.daily) return "daily_page_limit";
  if (ledger.monthlyPages + pages > max.monthly) return "monthly_page_limit";
  ledger.dailyPages += pages;
  ledger.monthlyPages += pages;
  ledger.attempts[id] = {
    count: (previous?.count ?? 0) + 1,
    succeeded: false,
    updatedAt: now.toISOString(),
  };
  return null;
}
