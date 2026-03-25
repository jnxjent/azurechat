// src/lib/sl-dept.ts
// SharePoint連携(SL)の「メール → 部署判定」＋「部署設定取得」ユーティリティ
// 新仕様対応版
//
// ポイント:
// 1) 「部署」と「アップロード種別(common/personal)」を分離
// 2) SharePoint対応可否は dept から判定
// 3) 個人フォルダー名は email の @ 前から生成
//
// 想定:
// - 部署: cp / ss / others ...
// - uploadScope: common / personal
// - SP非対応部署: 例 others
//
// 主な環境変数:
// - SL_DEPTS=cp,ss,others
// - SL_DEPT_DEFAULT=cp
// - SL_DEPT_NON_SP=others
//   （複数あるなら カンマ区切り: others,xx ）
//
// - SL_CP_SITE_URL=...
// - SL_CP_DRIVE_NAME=...
// - SL_CP_FOLDER=SL
//
// - SL_SS_SITE_URL=...
// - SL_SS_DRIVE_NAME=...
// - SL_SS_FOLDER=SL
//
// - SL_OTHERS_SITE_URL=...   // 非SP部署なら通常は不要
// - SL_OTHERS_DRIVE_NAME=... // 非SP部署なら通常は不要
//
// - SL_DEPT_BY_EMAIL_CP=a@x.com,b@x.com
// - SL_DEPT_BY_EMAIL_SS=c@x.com,d@x.com
// - SL_DEPT_BY_EMAIL_OTHERS=e@x.com

export type SlDeptConfig = {
  dept: string; // lower-case ("cp" | "ss" | "others" ...)
  siteUrl: string;
  driveName: string;
  folder: string; // ベースフォルダー。例: "SL"
};

export type UploadScope = "common" | "personal";

const RESERVED_UPLOAD_SCOPES = new Set(["common", "personal"]);

function parseCsv(value?: string): string[] {
  if (!value) return [];
  return value
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

function parseCsvLower(value?: string): string[] {
  return parseCsv(value).map((s) => s.toLowerCase());
}

function parseCsvEmails(value?: string): Set<string> {
  return new Set(parseCsv(value).map((s) => s.toLowerCase()));
}

/**
 * 許可された部署一覧
 * 例: SL_DEPTS=cp,ss,others
 *
 * 念のため common / personal が混入していても除外する。
 */
export function getAllowedDepts(): string[] {
  const raw = process.env.SL_DEPTS ?? "cp";
  return parseCsvLower(raw).filter((d) => !RESERVED_UPLOAD_SCOPES.has(d));
}

export function isAllowedDept(dept: string): boolean {
  const d = dept.trim().toLowerCase();
  return getAllowedDepts().includes(d);
}

/**
 * user email から部署を判定
 * 例:
 *   SL_DEPT_BY_EMAIL_CP=a@x.com,b@x.com
 *   SL_DEPT_BY_EMAIL_SS=c@x.com,d@x.com
 */
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

/**
 * 部署の最終決定
 * 優先順位:
 * 1) userEmail による部署判定
 * 2) requestedDept（ただし許可済み dept のみ）
 * 3) default
 *
 * 注意:
 * requestedDept に common / personal が来ても dept ではないので採用しない。
 */
export function decideDept(params: {
  requestedDept?: string;
  userEmail?: string | null;
}): string {
  const defaultDept = (process.env.SL_DEPT_DEFAULT ?? "cp").toLowerCase();

  // 1) email → dept（最優先）
  if (params.userEmail) {
    const hit = detectDeptByEmail(params.userEmail);
    if (hit) return hit;
  }

  // 2) requested dept（メール未登録時のみ使用）
  if (params.requestedDept) {
    const d = params.requestedDept.trim().toLowerCase();

    if (RESERVED_UPLOAD_SCOPES.has(d)) {
      console.warn(`[SL] requestedDept="${d}" is uploadScope, not dept. ignored.`);
    } else if (isAllowedDept(d)) {
      return d;
    } else {
      console.warn(`[SL] Requested dept not allowed: "${params.requestedDept}" -> fallback`);
    }
  }

  // 3) default
  if (!isAllowedDept(defaultDept)) {
    throw new Error(
      `SL_DEPT_DEFAULT="${defaultDept}" is not included in SL_DEPTS.`
    );
  }

  return defaultDept;
}

/**
 * SharePoint対応していない部署一覧
 * 例:
 *   SL_DEPT_NON_SP=others
 *   SL_DEPT_NON_SP=others,xx
 */
export function getNonSharePointDepts(): string[] {
  return parseCsvLower(process.env.SL_DEPT_NON_SP ?? "others");
}

/**
 * その部署が SharePoint 対応か
 */
export function isSharePointEnabledDept(dept: string): boolean {
  const d = dept.trim().toLowerCase();

  if (!isAllowedDept(d)) {
    throw new Error(`Dept "${dept}" is not allowed (check SL_DEPTS).`);
  }

  return !getNonSharePointDepts().includes(d);
}

/**
 * uploadScope の正規化
 * - common / personal のみ許可
 * - それ以外は personal 扱い
 *
 * 旧値互換:
 * - cp -> personal
 */
export function normalizeUploadScope(value?: string | null): UploadScope {
  const v = (value ?? "").trim().toLowerCase();

  if (v === "common") return "common";
  if (v === "personal") return "personal";

  // 旧実装の互換
  if (v === "cp") return "personal";

  return "personal";
}

/**
 * email から個人フォルダー名を生成
 * 例:
 *   nomoto@midac.jp -> nomoto
 */
export function getPersonalFolderNameFromEmail(email: string): string {
  const emailLc = email.trim().toLowerCase();
  const at = emailLc.indexOf("@");
  const localPart = at >= 0 ? emailLc.slice(0, at) : emailLc;

  const sanitized = localPart
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, "_")
    .replace(/\.+$/g, "")
    .trim();

  if (!sanitized) {
    throw new Error(`Failed to build personal folder name from email: "${email}"`);
  }

  return sanitized;
}

/**
 * 部署設定取得
 * folder は「ベースフォルダー」として扱う
 * 例:
 *   SL_CP_FOLDER=SL
 *   → 個人アップロード先は route 側で "SL/<mailPrefix>" を組み立てる
 */
export function getDeptConfig(deptLower: string): SlDeptConfig {
  const dept = deptLower.trim().toLowerCase();

  // ★ "common" は RESERVED だが getDeptConfig では許可（SLCommon用）
  if (dept !== "common" && !isAllowedDept(dept)) {
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
 * ベースフォルダー配下の最終アップロード先を組み立てる
 *
 * 例:
 *   baseFolder = "SL"
 *   uploadScope = "common"   -> "SL/common"
 *   uploadScope = "personal" -> "SL/nomoto"
 *
 * 共通フォルダー名は必要なら env で変更可能:
 *   SL_COMMON_SUBFOLDER=common
 */
export function buildUploadFolder(params: {
  baseFolder: string;
  uploadScope: UploadScope;
  userEmail: string;
}): string {
  const baseFolder = params.baseFolder.trim().replace(/^\/+|\/+$/g, "");
  if (!baseFolder) {
    throw new Error("baseFolder is empty.");
  }

  if (params.uploadScope === "common") {
    const commonSubfolder =
      (process.env.SL_COMMON_SUBFOLDER ?? "common").trim() || "common";
    return `${baseFolder}/${commonSubfolder}`;
  }

  const personalFolder = getPersonalFolderNameFromEmail(params.userEmail);
  return `${baseFolder}/${personalFolder}`;
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
    if (typeof c === "string" && c.trim()) {
      return c.trim().toLowerCase();
    }
  }

  return null;
}
// ============================================================
// ★ DeptAdmin / SlRole（2026-03追加）
// ============================================================

export type SlRole = "global_admin" | "dept_admin" | "dept_member";

/**
 * メールアドレスが指定部署のDeptAdminか判定
 * 環境変数: SL_DEPT_ADMIN_EMAILS_CP / SL_DEPT_ADMIN_EMAILS_SS
 */
export function isDeptAdmin(email: string, dept: string): boolean {
  const emailLc = email.trim().toLowerCase();
  const key = `SL_DEPT_ADMIN_EMAILS_${dept.toUpperCase()}`;
  const set = parseCsvEmails(process.env[key]); // 既存関数を再利用
  return set.has(emailLc);
}

/**
 * SlRole解決（優先順位: global_admin > dept_admin > dept_member）
 * 環境変数:
 *   SL_ADMIN_EMAILS          → global_admin
 *   SL_DEPT_ADMIN_EMAILS_CP  → dept_admin（cp）
 *   SL_DEPT_ADMIN_EMAILS_SS  → dept_admin（ss）
 */
export function resolveSlRole(
  email: string | null | undefined,
  dept: string
): SlRole {
  if (!email) return "dept_member";
  const emailLc = email.trim().toLowerCase();

  // 1) GlobalAdmin
  const globalAdminSet = parseCsvEmails(process.env.SL_ADMIN_EMAILS);
  if (globalAdminSet.has(emailLc)) return "global_admin";

  // 2) DeptAdmin
  if (isDeptAdmin(emailLc, dept)) return "dept_admin";

  // 3) Member
  return "dept_member";
}