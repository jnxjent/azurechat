export const SL_SYNC_CONTINUING_RETRY_MAX_PAGES = 200;

export type SlSyncAttempt = {
  count: number;
  totalCount?: number;
  succeeded: boolean;
  updatedAt: string;
  pages?: number;
  day?: string;
  blockedPageLimit?: number;
  deferredBudget?: {
    reason: "daily_page_limit" | "monthly_page_limit";
    period: string;
    limit: number;
  };
};
export type SlSyncLedger = {
  version: 1;
  day: string;
  month: string;
  dailyPages: number;
  monthlyPages: number;
  attempts: Record<string, SlSyncAttempt>;
};

export type SlSyncLimits = {
  daily: number;
  monthly: number;
  perFile: number;
  attempts: number;
  totalAttempts: number;
};

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
  const previous = ledger.attempts[id];
  const totalCount = previous?.totalCount ?? previous?.count ?? 0;
  if (previous?.succeeded) return "already_indexed_this_version";
  if (pages > max.perFile) {
    // Remember oversized file versions so they cannot fill the first batch on
    // every run. Raising the configured limit makes them eligible again.
    ledger.attempts[id] = {
      count: previous?.count ?? 0,
      totalCount,
      succeeded: false,
      updatedAt: now.toISOString(),
      pages,
      day: ledger.day,
      blockedPageLimit: max.perFile,
    };
    return "file_page_limit";
  }
  const continuing = pages <= SL_SYNC_CONTINUING_RETRY_MAX_PAGES;
  const priorCount = continuing && previous?.day !== ledger.day ? 0 : (previous?.count ?? 0);
  if (totalCount >= max.totalAttempts) {
    if (previous) {
      previous.totalCount = totalCount;
      previous.pages = pages;
    }
    return "total_attempt_limit";
  }
  if (priorCount >= max.attempts) {
    // Fill in page metadata for ledgers created before the retry policy was
    // added. The caller persists this even when the claim is denied.
    if (previous) {
      previous.pages = pages;
      previous.totalCount = totalCount;
      previous.day ??= ledger.day;
      previous.blockedPageLimit = undefined;
    }
    return continuing ? "daily_attempt_limit" : "attempt_limit";
  }
  if (ledger.dailyPages + pages > max.daily || ledger.monthlyPages + pages > max.monthly) {
    const daily = ledger.dailyPages + pages > max.daily;
    const reason = daily ? "daily_page_limit" : "monthly_page_limit";
    ledger.attempts[id] = {
      count: previous?.count ?? 0,
      totalCount,
      succeeded: false,
      updatedAt: now.toISOString(),
      pages,
      day: previous?.day,
      deferredBudget: {
        reason,
        period: daily ? ledger.day : ledger.month,
        limit: daily ? max.daily : max.monthly,
      },
    };
    return reason;
  }
  ledger.dailyPages += pages;
  ledger.monthlyPages += pages;
  ledger.attempts[id] = {
    count: priorCount + 1,
    totalCount: totalCount + 1,
    succeeded: false,
    updatedAt: now.toISOString(),
    pages,
    day: ledger.day,
  };
  return null;
}

// Download and page-count failures happen before an OCR reservation. Count
// them toward the file-version retry ceiling without charging OCR pages or
// incrementing the daily OCR-attempt counter.
export function recordSlSyncPreOcrFailure(
  ledger: SlSyncLedger,
  id: string,
  max: SlSyncLimits,
  now: Date
): number {
  const previous = ledger.attempts[id];
  const totalCount = previous?.totalCount ?? previous?.count ?? 0;
  if (previous?.succeeded || totalCount >= max.totalAttempts) return totalCount;
  ledger.attempts[id] = {
    count: previous?.count ?? 0,
    totalCount: totalCount + 1,
    succeeded: false,
    updatedAt: now.toISOString(),
    pages: previous?.pages,
    day: previous?.day,
  };
  return totalCount + 1;
}
