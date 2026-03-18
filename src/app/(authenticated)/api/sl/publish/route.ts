// app/(authenticated)/api/sl/publish/route.ts

export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import { getToken } from "next-auth/jwt";
import {
  buildUploadFolder,
  decideDept,
  getDeptConfig,
  getUserEmailFromJwtToken,
  isSharePointEnabledDept,
  normalizeUploadScope,
  type UploadScope,
} from "@/lib/sl-dept";

// -------------------------------------------------------
// Token refresh helper
// -------------------------------------------------------
async function getValidAccessToken(req: NextRequest): Promise<string | null> {
  const token = await getToken({ req });
  if (!token) return null;

  const accessToken = (token as any).accessToken as string | undefined;
  const expiresAt = (token as any).accessTokenExpiresAt as number | undefined;
  const refreshToken = (token as any).refreshToken as string | undefined;

  if (!accessToken) return null;

  const nowSec = Math.floor(Date.now() / 1000);
  if (expiresAt && nowSec < expiresAt - 60) {
    return accessToken;
  }

  if (!refreshToken) return accessToken;

  const tenantId = process.env.AZURE_AD_TENANT_ID;
  const clientId = process.env.AZURE_AD_CLIENT_ID;
  const clientSecret = process.env.AZURE_AD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) return accessToken;

  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          grant_type: "refresh_token",
          refresh_token: refreshToken,
          scope: "openid profile email offline_access User.Read Files.ReadWrite",
        }),
      }
    );

    const data: any = await res.json().catch(() => ({}));
    if (!res.ok) {
      console.error("[SL publish] Token refresh failed:", data);
      return accessToken;
    }

    console.log("[SL publish] Token refreshed successfully");
    return data.access_token as string;
  } catch (e) {
    console.error("[SL publish] Token refresh error:", e);
    return accessToken;
  }
}

// -------------------------------------------------------
// Admin helpers
// -------------------------------------------------------
function parseCsvLower(value?: string | null) {
  return (value ?? "")
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);
}

function isAdminEmail(email: string | null) {
  if (!email) return false;
  const admins = parseCsvLower(process.env.SL_ADMIN_EMAILS);
  return admins.includes(email.toLowerCase());
}

// -------------------------------------------------------
// Graph API helpers
// -------------------------------------------------------
async function resolveSiteAndDrive(
  accessToken: string,
  siteUrl: string,
  driveName: string
): Promise<{ siteId: string; driveId: string }> {
  const url = new URL(siteUrl);
  const host = url.hostname;
  const path = url.pathname;

  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${host}:${path}`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!siteRes.ok) {
    const err = await siteRes.text();
    throw new Error(`Failed to get site: ${err}`);
  }
  const siteJson = await siteRes.json();
  const siteId: string = siteJson.id;

  const drivesRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!drivesRes.ok) {
    const err = await drivesRes.text();
    throw new Error(`Failed to get drives: ${err}`);
  }
  const drivesJson = await drivesRes.json();
  const drive = drivesJson.value.find((d: any) => d.name === driveName);
  if (!drive) {
    throw new Error(
      `Drive "${driveName}" not found. Available: ${drivesJson.value
        .map((d: any) => d.name)
        .join(", ")}`
    );
  }

  return { siteId, driveId: drive.id };
}

async function graphPutBinary(
  uploadUrl: string,
  accessToken: string,
  buffer: Buffer,
  mimeType: string
): Promise<any> {
  const res = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": mimeType,
    },
    body: new Uint8Array(buffer),  // ← Uint8Array は BodyInit として有効
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Upload failed (${res.status}): ${err}`);
  }
  return res.json();
}

// -------------------------------------------------------
// Utils
// -------------------------------------------------------
function getMimeType(fileName: string): string {
  const ext = fileName.split(".").pop()?.toLowerCase() ?? "";

  const mimeMap: Record<string, string> = {
    pdf: "application/pdf",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    txt: "text/plain",
  };

  return mimeMap[ext] ?? "application/octet-stream";
}

/**
 * 互換対応:
 * - 新: body.uploadScope
 * - 旧: body.dept に common / personal / cp が来る場合
 */
function resolveRequestedUploadScope(body: any): UploadScope {
  return normalizeUploadScope(body?.uploadScope ?? body?.dept);
}

// -------------------------------------------------------
// Route handlers
// -------------------------------------------------------
export async function OPTIONS() {
  return NextResponse.json({}, { status: 200 });
}

export async function POST(req: NextRequest) {
  try {
    const accessToken = await getValidAccessToken(req);
    if (!accessToken) {
      return NextResponse.json(
        { ok: false, error: "No access token. Please re-login." },
        { status: 401 }
      );
    }

    const body = await req.json();
    const { fileName, fileBase64 } = body;

    if (!fileName || !fileBase64) {
      return NextResponse.json(
        { ok: false, error: "Missing file data" },
        { status: 400 }
      );
    }

    const token = await getToken({ req });
    const userEmail = token ? getUserEmailFromJwtToken(token) : null;

    if (!userEmail) {
      return NextResponse.json(
        { ok: false, error: "Failed to identify current user email." },
        { status: 401 }
      );
    }

    const requestedUploadScope = resolveRequestedUploadScope(body);

    // 部署は userEmail から確定
    // requestedDept は「本当に部署を明示したい互換用途」がある場合だけ補助的に使う
    const deptLower = decideDept({
      requestedDept:
        typeof body?.requestedDept === "string" ? body.requestedDept : undefined,
      userEmail,
    });

    const isSpDept = isSharePointEnabledDept(deptLower);
    const admin = isAdminEmail(userEmail);

    // SP非対応部署はこのAPI対象外
    if (!isSpDept) {
      return NextResponse.json(
        {
          ok: false,
          error:
            "Your department is not SharePoint-enabled. Use Blob upload flow instead.",
          dept: deptLower,
          isSharePointEnabled: false,
        },
        { status: 400 }
      );
    }

    // 一般メンバーは personal 固定
    const actualUploadScope: UploadScope = admin
      ? requestedUploadScope
      : "personal";

    // 部署ごとの SP 設定
    const { siteUrl, driveName, folder: baseFolder } = getDeptConfig(deptLower);

    // 最終アップ先フォルダー
    // common   -> SL/common
    // personal -> SL/<mailPrefix>
    const uploadFolder = buildUploadFolder({
      baseFolder,
      uploadScope: actualUploadScope,
      userEmail,
    });

    console.log(
      `[SL publish] dept=${deptLower} user=${userEmail} admin=${admin} requestedScope=${requestedUploadScope} actualScope=${actualUploadScope} site=${siteUrl} drive=${driveName} folder=${uploadFolder}`
    );

    const fileBuffer = Buffer.from(fileBase64, "base64");
    const mimeType = getMimeType(fileName);

    const { driveId } = await resolveSiteAndDrive(accessToken, siteUrl, driveName);

    const uploadUrl =
      `https://graph.microsoft.com/v1.0/drives/${driveId}` +
      `/root:/${uploadFolder}/${fileName}:/content`;

    console.log(`[SL publish] Uploading to: ${uploadUrl}`);

    const result = await graphPutBinary(
      uploadUrl,
      accessToken,
      fileBuffer,
      mimeType
    );

    console.log(`[SL publish] Upload success: ${result.webUrl}`);

    return NextResponse.json({
      ok: true,
      dept: deptLower,
      uploadScope: actualUploadScope,
      isSharePointEnabled: true,
      name: result.name,
      webUrl: result.webUrl,
    });
  } catch (e: any) {
    console.error("[SL publish] Error:", e);
    return NextResponse.json(
      { ok: false, error: String(e?.message ?? e) },
      { status: 500 }
    );
  }
}