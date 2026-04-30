// app/(authenticated)/api/sl/publish/route.ts

export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import { getToken } from "next-auth/jwt";
import {
  buildUploadFolder,
  decideDept,
  getDeptConfig,
  getEffectiveSlUserEmail,
  getUserEmailFromJwtToken,
  isSharePointEnabledDept,
  normalizeUploadScope,
  resolveSlAccess,
  resolveSlUploadTarget,
  type UploadScope,
} from "@/lib/sl-dept";

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

    return data.access_token as string;
  } catch (e) {
    console.error("[SL publish] Token refresh error:", e);
    return accessToken;
  }
}

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
    throw new Error(`Failed to get site: ${await siteRes.text()}`);
  }
  const siteJson = await siteRes.json();
  const siteId: string = siteJson.id;

  const drivesRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!drivesRes.ok) {
    throw new Error(`Failed to get drives: ${await drivesRes.text()}`);
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
    body: new Uint8Array(buffer),
  });

  if (!res.ok) {
    throw new Error(`Upload failed (${res.status}): ${await res.text()}`);
  }
  return res.json();
}

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

function resolveRequestedUploadScope(body: any): UploadScope {
  return normalizeUploadScope(body?.uploadScope ?? body?.dept);
}

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
    const fileName = String(body.fileName ?? "").replace(/^[\s\u3000]+|[\s\u3000]+$/g, "");
    const { fileBase64 } = body;

    if (!fileName || !fileBase64) {
      return NextResponse.json(
        { ok: false, error: "Missing file data" },
        { status: 400 }
      );
    }

    const token = await getToken({ req });
    const rawUserEmail = token ? getUserEmailFromJwtToken(token) : null;
    const userEmail = getEffectiveSlUserEmail(rawUserEmail);

    if (!userEmail) {
      return NextResponse.json(
        { ok: false, error: "Failed to identify current user email." },
        { status: 401 }
      );
    }

    const requestedUploadScope = resolveRequestedUploadScope(body);
    const preferredDept = decideDept({
      requestedDept:
        typeof body?.requestedDept === "string" ? body.requestedDept : undefined,
      userEmail,
    });

    const access = resolveSlAccess(userEmail, preferredDept);
    const admin =
      access.role === "global_admin" || access.role === "dept_admin";
    const actualUploadScope: UploadScope = admin
      ? requestedUploadScope
      : "personal";

    const uploadTarget = resolveSlUploadTarget(
      userEmail,
      actualUploadScope,
      preferredDept
    );
    const targetDept = uploadTarget.dept;

    const isSpDept =
      targetDept === "common" ? true : isSharePointEnabledDept(targetDept);

    if (!isSpDept) {
      return NextResponse.json(
        {
          ok: false,
          error:
            "Your department is not SharePoint-enabled. Use Blob upload flow instead.",
          dept: targetDept,
          isSharePointEnabled: false,
        },
        { status: 400 }
      );
    }

    const { siteUrl, driveName, folder: baseFolder } = getDeptConfig(targetDept);
    const uploadFolder =
      targetDept === "common" && actualUploadScope === "common"
        ? baseFolder
        : buildUploadFolder({
            baseFolder,
            uploadScope: actualUploadScope,
            userEmail,
          });

    console.log(
      `[SL publish] preferredDept=${preferredDept} targetDept=${targetDept} user=${userEmail} role=${access.role} requestedScope=${requestedUploadScope} actualScope=${actualUploadScope} site=${siteUrl} drive=${driveName} folder=${uploadFolder}`
    );

    const fileBuffer = Buffer.from(fileBase64, "base64");
    const mimeType = getMimeType(fileName);
    const { driveId } = await resolveSiteAndDrive(accessToken, siteUrl, driveName);

    const uploadUrl =
      `https://graph.microsoft.com/v1.0/drives/${driveId}` +
      `/root:/${uploadFolder}/${fileName}:/content`;

    const result = await graphPutBinary(
      uploadUrl,
      accessToken,
      fileBuffer,
      mimeType
    );

    return NextResponse.json({
      ok: true,
      dept: targetDept,
      uploadScope: actualUploadScope,
      isSharePointEnabled: true,
      name: result.name,
      webUrl: result.webUrl,
      spItemId: result.id ?? null,
    });
  } catch (e: any) {
    console.error("[SL publish] Error:", e);
    return NextResponse.json(
      { ok: false, error: String(e?.message ?? e) },
      { status: 500 }
    );
  }
}
