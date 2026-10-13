import "server-only";
import { createHmac, timingSafeEqual } from "crypto";

const MAX_AGE_SECONDS = 30 * 24 * 60 * 60;
function secret(): string {
  const value = process.env.NEXTAUTH_SECRET || process.env.AZURE_STORAGE_ACCOUNT_KEY;
  if (!value) throw new Error("Teams image signing is not configured");
  return value.trim();
}
function validPath(threadId: string, fileName: string): boolean {
  return /^teams-[a-f0-9]{32}$/.test(threadId) && /^[A-Za-z0-9-]{1,80}\.png\.preview\.jpg$/.test(fileName);
}
function sign(threadId: string, fileName: string, expires: string): string {
  return createHmac("sha256", secret()).update(`teams-image-v1\n${threadId}\n${fileName}\n${expires}`).digest("hex");
}
export function buildTeamsImagePreviewUrl(threadId: string, fileName: string): string {
  if (!validPath(threadId, fileName)) throw new Error("Invalid Teams preview path");
  const base = process.env.TEAMS_PUBLIC_BASE_URL || process.env.NEXTAUTH_URL;
  if (!base) throw new Error("Teams public URL is not configured");
  const url = new URL("/api/teams/images", base);
  const local = process.env.NODE_ENV === "development" && /^(localhost|127\.0\.0\.1)$/.test(url.hostname);
  if (url.protocol !== "https:" && !(local && url.protocol === "http:")) throw new Error("Teams public URL must use HTTPS");
  const expires = String(Math.floor(Date.now() / 1000) + MAX_AGE_SECONDS);
  url.search = new URLSearchParams({ thread: threadId, file: fileName, expires, sig: sign(threadId, fileName, expires) }).toString();
  return url.toString();
}
export function validateTeamsImagePreviewUrl(params: URLSearchParams): { threadId: string; fileName: string } | null {
  const threadId = params.get("thread") || "", fileName = params.get("file") || "";
  const expires = params.get("expires") || "", sig = params.get("sig") || "";
  const now = Math.floor(Date.now() / 1000);
  if (!validPath(threadId, fileName) || !/^\d{10}$/.test(expires) || !/^[a-f0-9]{64}$/.test(sig) ||
      Number(expires) <= now || Number(expires) > now + MAX_AGE_SECONDS) return null;
  const expected = sign(threadId, fileName, expires);
  return timingSafeEqual(Buffer.from(expected, "hex"), Buffer.from(sig, "hex")) ? { threadId, fileName } : null;
}
