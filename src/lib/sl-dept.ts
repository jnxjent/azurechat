// src/lib/sl-dept.ts
// SharePoint連携(SL)の「メール → 部署判定」＋「部署設定取得」ユーティリティ

export type SlDeptConfig = {
  dept: string; // lower-case ("cp" | "ss" | "others" ...)
  siteUrl: string;
  driveName: string;
  folder: string;
};

function parseCsv(value?: string): string[] {
  if (!value) return [];
  return value
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

function parseCsvEmails(value?: string): Set<string> {
  return new Set(parseCsv(value).map((s) => s.toLowerCase()));
}

export function getAllowedDepts(): string[] {
  // 例: SL_DEPTS=cp,ss,others,common
  const raw = process.env.SL_DEPTS ?? "cp";
  return parseCsv(raw).map((s) => s.toLowerCase());
}

export function isAllowedDept(dept: string): boolean {
  const d = dept.trim().toLowerCase();
  return getAllowedDepts().includes(d);
}

export function detectDeptByEmail(email: string): string | null {
  const emailLc = email.trim().toLowerCase();
  const allowed = getAllowedDepts();

  for (const dept of allowed) {
    const key = `SL_DEPT_BY_EMAIL_${dept.toUpperCase()}`;
    const set = parseCsvEmails(process.env[key]);
    if (set.has(emailLc)) return dept;
  }
  return null;
}

export function decideDept(params: {
  requestedDept?: string;
  userEmail?: string | null;
}): string {
  const defaultDept = (process.env.SL_DEPT_DEFAULT ?? "cp").toLowerCase();

  // 1) email → dept（メール登録リストを最優先）
  if (params.userEmail) {
    const hit = detectDeptByEmail(params.userEmail);
    if (hit) return hit;
  }

  // 2) requested dept（メール未登録時のみ使用）
  if (params.requestedDept) {
    const d = params.requestedDept.trim().toLowerCase();
    if (isAllowedDept(d)) return d;
    console.warn(`[SL] Requested dept not allowed: "${params.requestedDept}" -> fallback`);
  }

  // 3) default
  return defaultDept;
}

export function getDeptConfig(deptLower: string): SlDeptConfig {
  const dept = deptLower.trim().toLowerCase();
  if (!isAllowedDept(dept)) {
    throw new Error(`Dept "${dept}" is not allowed (check SL_DEPTS).`);
  }

  const key = dept.toUpperCase();
  const siteUrl = process.env[`SL_${key}_SITE_URL`];
  const driveName = process.env[`SL_${key}_DRIVE_NAME`];
  const folder = process.env[`SL_${key}_FOLDER`] ?? "SL";

  if (!siteUrl || !driveName) {
    throw new Error(
      `Missing env vars for dept: ${key} (SL_${key}_SITE_URL / SL_${key}_DRIVE_NAME)`
    );
  }

  return { dept, siteUrl, driveName, folder };
}

/**
 * next-auth/jwt token から user email を取り出す（環境によりキーが揺れる）
 */
export function getUserEmailFromJwtToken(token: any): string | null {
  const candidates = [
    token?.email,
    token?.preferred_username,
    token?.upn,
    token?.unique_name,
  ];
  for (const c of candidates) {
    if (typeof c === "string" && c.trim()) return c.trim();
  }
  return null;
}