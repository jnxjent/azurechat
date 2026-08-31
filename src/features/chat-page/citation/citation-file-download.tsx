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

type FreshSharePointLink = {
  url: string;
  kind: "download" | "web";
};

function validHttpsUrl(value: unknown): value is string {
  if (typeof value !== "string" || !value.trim()) return false;
  try {
    return new URL(value).protocol === "https:";
  } catch {
    return false;
  }
}

async function getFreshLinkBySpItemId(
  dept: string,
  spItemId: string
): Promise<FreshSharePointLink | null> {
  const token = await getAppOnlyToken();
  if (!token) return null;

  const depts = dept ? [dept] : getAllowedDepts();

  for (const d of depts) {
    try {
      const config = getDeptConfig(d);
      const driveId = await resolveDriveId(token, config.siteUrl, config.driveName);
      if (!driveId) continue;

      const res = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${encodeURIComponent(spItemId)}`,
        { headers: { Authorization: `Bearer ${token}` }, cache: "no-store" }
      );
      if (!res.ok) continue;
      const item = await res.json();
      if (item?.deleted) continue;

      const downloadUrl = item?.["@microsoft.graph.downloadUrl"];
      if (validHttpsUrl(downloadUrl)) {
        return { url: downloadUrl, kind: "download" };
      }
      if (validHttpsUrl(item?.webUrl)) {
        return { url: item.webUrl, kind: "web" };
      }
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
    console.log("[DL] spItemId=", document.spItemId, "dept=", document.dept);

    if (document.spItemId) {
      const freshLink = await getFreshLinkBySpItemId(
        document.dept ?? "",
        document.spItemId
      );
      if (freshLink) {
        // Do not log the actual download URL because Graph embeds a short-lived token.
        console.log(`[DL] Graph link resolved kind=${freshLink.kind}`);
        return freshLink.url;
      }
      console.warn(
        "[DL] Graph API lookup failed; refusing stale stored SharePoint URL"
      );
      return null;
    }

    if (document.isSlDoc === true) {
      console.warn("[DL] SharePoint citation has no spItemId; link unavailable");
      return null;
    }

    return document.effectiveFileUrl || document.fileUrl;
  }
  return null;
};
