import { GetImageFromStore } from "@/features/chat-page/chat-services/chat-image-service";
import { validateTeamsImagePreviewUrl } from "@/features/teams/teams-image-url";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

/** A bounded signed preview, served through the app's allowed storage network. */
export async function GET(request: Request) {
  try {
    const image = validateTeamsImagePreviewUrl(new URL(request.url).searchParams);
    if (!image) return new Response(null, { status: 403 });
    const stored = await GetImageFromStore(image.threadId, image.fileName);
    if (stored.status !== "OK") return new Response(null, { status: 404 });
    const chunks: Buffer[] = [];
    let length = 0;
    for await (const chunk of stored.response as unknown as AsyncIterable<Uint8Array>) {
      length += chunk.length;
      if (length > 1024 * 1024) return new Response(null, { status: 413 });
      chunks.push(Buffer.from(chunk));
    }
    return new Response(Buffer.concat(chunks), { headers: {
      "Content-Type": "image/jpeg", "Cache-Control": "private, max-age=3600", "X-Content-Type-Options": "nosniff",
    } });
  } catch {
    return new Response(null, { status: 503 });
  }
}
