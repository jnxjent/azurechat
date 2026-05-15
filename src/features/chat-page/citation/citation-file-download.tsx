"use server";

import { FindCitationByID } from "../chat-services/citation-service";
import { getDeptConfig, getAllowedDepts } from "@/lib/sl-dept";

async function getAppOnlyToken(): Promise<string | null> {
  const tenantId = process.env.AZURE_AD_TENANT_ID?.trim();
  const clientId = process.env.AZURE_AD_CLIENT_ID?.trim();
  const clientSecret = process.env.AZURE_AD_CLIENT_SECRET?.trim();
  if (!tenantId || !clientId || !clientSecret) return null;

  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        cache: "no-store",
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          grant_type: "client_credentials",
          scope: "https://graph.microsoft.com/.default",
        }),
      }
    );
    if (!res.ok) return null;
    const data = await res.json();
    return data.access_token ?? null;
  } catch {
    return null;
  }
}

async function resolveDriveId(token: string, siteUrl: string, driveName: string): Promise<string | null> {
  try {
    const u = new URL(siteUrl);
    const siteRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${u.hostname}:${u.pathname}`,
      { headers: { Authorization: `Bearer ${token}` }, cache: "no-store" }
    );
    if (!siteRes.ok) return null;
    const siteId = (await siteRes.json()).id;

    const drivesRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      { headers: { Authorization: `Bearer ${token}` }, cache: "no-store" }
    );
    if (!drivesRes.ok) return null;
    const drives = (await drivesRes.json()).value ?? [];
    const drive = drives.find((d: any) => d.name === driveName);
    return drive?.id ?? null;
  } catch {
    return null;
  }
}

async function getWebUrlBySpItemId(dept: string, spItemId: string): Promise<string | null> {
  const token = await getAppOnlyToken();
  if (!token) return null;

  const depts = dept ? [dept] : getAllowedDepts();

  for (const d of depts) {
    try {
      const config = getDeptConfig(d);
      const driveId = await resolveDriveId(token, config.siteUrl, config.driveName);
      if (!driveId) continue;

      const res = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${spItemId}?$select=webUrl,deleted`,
        { headers: { Authorization: `Bearer ${token}` }, cache: "no-store" }
      );
      if (!res.ok) continue;
      const item = await res.json();
      if (item?.deleted || !item?.webUrl) continue;
      return item.webUrl as string;
    } catch {
      continue;
    }
  }
  return null;
}

export const CitationFileDownload = async (formData: FormData) => {
  console.log("[DL] CitationFileDownload called, id=", formData.get("id"));
  const searchResponse = await FindCitationByID(formData.get("id") as string);
  if (searchResponse.status === "OK") {
    const { document } = searchResponse.response.content;
    console.log("[DL] spItemId=", document.spItemId, "dept=", document.dept, "effectiveFileUrl=", document.effectiveFileUrl);

    if (document.spItemId) {
      const freshUrl = await getWebUrlBySpItemId(document.dept ?? "", document.spItemId);
      if (freshUrl) {
        console.log("[DL] fresh URL from Graph API:", freshUrl);
        return freshUrl;
      }
      console.warn("[DL] Graph API lookup failed, falling back to stored URL");
    }

    return document.effectiveFileUrl || document.fileUrl;
  }
  return null;
};
