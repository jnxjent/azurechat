// src/lib/sl-dept.ts

export type SlDeptConfig = {
  dept: string;
  siteUrl: string;
  driveName: string;
  folder: string;
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

function normalizeOptionalEmail(email: string | null | undefined): string | null {
  const normalized = (email ?? "").trim().toLowerCase();
  return normalized || null;
}

export function getAllowedDepts(): string[] {
  const raw = process.env.SL_DEPTS ?? "cp";
  return parseCsvLower(raw).filter((d) => !RESERVED_UPLOAD_SCOPES.has(d));
}

export function isAllowedDept(dept: string): boolean {
  const d = dept.trim().toLowerCase();
  return getAllowedDepts().includes(d);
}

export function detectDeptByEmail(email: string): string | null {
  return detectAllDeptsByEmail(email)[0] ?? null;
}

export function detectAllDeptsByEmail(email: string): string[] {
  const emailLc = email.trim().toLowerCase();

  return getAllowedDepts().filter((dept) => {
    const key = `SL_DEPT_BY_EMAIL_${dept.toUpperCase()}`;
    const set = parseCsvEmails(process.env[key]);
    return set.has(emailLc);
  });
}

function getSafeDefaultDept(): string {
  const allowed = getAllowedDepts();
  const defaultDept = (process.env.SL_DEPT_DEFAULT ?? "cp").toLowerCase();

  if (allowed.includes(defaultDept)) {
    return defaultDept;
  }

  if (allowed.length > 0) {
    console.warn(
      `[SL] SL_DEPT_DEFAULT="${defaultDept}" is not in SL_DEPTS. Fallback to "${allowed[0]}".`
    );
    return allowed[0];
  }

  throw new Error(`No allowed departments found. Check SL_DEPTS.`);
}

/**
 * メールが指定されているがどの SL_DEPT_BY_EMAIL_* にもマッチしない場合の
 * フォールバック部署を返す。SL_DEPT_NON_SP で設定された非SP部署を優先し、
 * 見つからなければ SL_DEPT_DEFAULT に落ちる。
 */
function getUnknownEmailFallbackDept(): string {
  const nonSpDepts = getNonSharePointDepts();
  const allowed = getAllowedDepts();
  return nonSpDepts.find((d) => allowed.includes(d)) ?? getSafeDefaultDept();
}

export function decideDept(params: {
  requestedDept?: string;
  userEmail?: string | null;
}): string {
  const defaultDept = getSafeDefaultDept();

  if (params.userEmail) {
    const hit = detectDeptByEmail(params.userEmail);
    if (hit) return hit;
    // メールはあるが SL_DEPT_BY_EMAIL_* にマッチなし → 非SP部署にフォールバック
    return getUnknownEmailFallbackDept();
  }

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

  return defaultDept;
}

export function getNonSharePointDepts(): string[] {
  return parseCsvLower(process.env.SL_DEPT_NON_SP ?? "others");
}

export function isSharePointEnabledDept(dept: string): boolean {
  const d = dept.trim().toLowerCase();

  if (!isAllowedDept(d)) {
    throw new Error(`Dept "${dept}" is not allowed (check SL_DEPTS).`);
  }

  return !getNonSharePointDepts().includes(d);
}

export function normalizeUploadScope(value?: string | null): UploadScope {
  const v = (value ?? "").trim().toLowerCase();

  if (v === "common") return "common";
  if (v === "personal") return "personal";
  if (v === "cp") return "personal";

  return "personal";
}

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

export function getDeptConfig(deptLower: string): SlDeptConfig {
  const dept = deptLower.trim().toLowerCase();

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

export function buildUploadFolder(params: {
  baseFolder: string;
  uploadScope: UploadScope;
  userEmail: string;
}): string {
  const baseFolder = params.baseFolder.trim().replace(/^\/+|\/+$/g, "");

  if (params.uploadScope === "common") {
    const commonSubfolder =
      (process.env.SL_COMMON_SUBFOLDER ?? "common").trim() || "common";
    return baseFolder ? `${baseFolder}/${commonSubfolder}` : commonSubfolder;
  }

  const personalFolder = getPersonalFolderNameFromEmail(params.userEmail);
  return baseFolder ? `${baseFolder}/${personalFolder}` : personalFolder;
}

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

export function getEffectiveSlUserEmail(
  email: string | null | undefined
): string | null {
  const localOverride = normalizeOptionalEmail(process.env.SL_LOCAL_DEFAULT_EMAIL);
  const actualEmail = normalizeOptionalEmail(email);

  if (process.env.NODE_ENV === "development" && localOverride) {
    if (actualEmail) {
      const globalAdminSet = parseCsvEmails(process.env.SL_ADMIN_EMAILS);
      if (globalAdminSet.has(actualEmail)) {
        return actualEmail;
      }

      if (getDeptAdminDepts(actualEmail).length > 0) {
        return actualEmail;
      }

      if (detectAllDeptsByEmail(actualEmail).length > 0) {
        return actualEmail;
      }
    }

    return localOverride;
  }

  return actualEmail;
}

export type SlRole = "global_admin" | "dept_admin" | "dept_member";

export type SlAccess = {
  role: SlRole;
  dept: string;
};

export type SlUploadTarget = {
  dept: string;
  access: SlAccess;
};

export function isDeptAdmin(email: string, dept: string): boolean {
  const emailLc = email.trim().toLowerCase();
  const key = `SL_DEPT_ADMIN_EMAILS_${dept.toUpperCase()}`;
  const set = parseCsvEmails(process.env[key]);
  return set.has(emailLc);
}

export function getDeptAdminDepts(email: string): string[] {
  const emailLc = email.trim().toLowerCase();
  return getAllowedDepts().filter((dept) => isDeptAdmin(emailLc, dept));
}

function pickPreferredDept(
  preferredDept: string | null | undefined,
  candidates: string[]
): string {
  const preferred = (preferredDept ?? "").trim().toLowerCase();
  if (preferred && candidates.includes(preferred)) {
    return preferred;
  }

  return candidates[0];
}

export function resolveSlAccess(
  email: string | null | undefined,
  preferredDept?: string | null
): SlAccess {
  const fallbackDept =
    preferredDept && isAllowedDept(preferredDept)
      ? preferredDept.trim().toLowerCase()
      : getSafeDefaultDept();

  if (!email) {
    return { role: "dept_member", dept: fallbackDept };
  }

  const emailLc = email.trim().toLowerCase();

  const globalAdminSet = parseCsvEmails(process.env.SL_ADMIN_EMAILS);
  if (globalAdminSet.has(emailLc)) {
    return {
      role: "global_admin",
      dept: decideDept({
        requestedDept: preferredDept ?? undefined,
        userEmail: emailLc,
      }),
    };
  }

  const adminDepts = getDeptAdminDepts(emailLc);
  if (adminDepts.length > 0) {
    return {
      role: "dept_admin",
      dept: pickPreferredDept(preferredDept, adminDepts),
    };
  }

  const memberDepts = detectAllDeptsByEmail(emailLc);
  if (memberDepts.length > 0) {
    return {
      role: "dept_member",
      dept: pickPreferredDept(preferredDept, memberDepts),
    };
  }

  // メール指定あるがどの SL_DEPT_BY_EMAIL_* にもマッチしないユーザー → 非SP部署
  return {
    role: "dept_member",
    dept: getUnknownEmailFallbackDept(),
  };
}

export function resolveSlRole(
  email: string | null | undefined,
  dept: string
): SlRole {
  return resolveSlAccess(email, dept).role;
}

export function resolveSlUploadTarget(
  email: string | null | undefined,
  uploadScope: UploadScope,
  preferredDept?: string | null
): SlUploadTarget {
  const access = resolveSlAccess(email, preferredDept);
  const emailDept = email ? detectDeptByEmail(email) : null;

  if (uploadScope === "common") {
    if (access.role === "global_admin") {
      return {
        dept: "common",
        access,
      };
    }

    return {
      dept: access.dept,
      access,
    };
  }

  if (emailDept) {
    return {
      dept: emailDept,
      access,
    };
  }

  if (access.role === "dept_admin") {
    return {
      dept: access.dept,
      access,
    };
  }

  return {
    dept: decideDept({
      requestedDept: preferredDept ?? undefined,
      userEmail: email ?? undefined,
    }),
    access,
  };
}
