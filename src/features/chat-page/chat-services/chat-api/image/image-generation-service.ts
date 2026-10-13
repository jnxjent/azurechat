import "server-only";
import { OpenAIDALLEInstance } from "@/features/common/services/openai";
import { uniqueId } from "@/features/common/util";
import { UploadImageToStore } from "../../chat-image-service";
import { buildFaithfulImagePrompt, normalizeGptImageSize } from "./image-intent";
import { normalizeGptImageQuality } from "./image-quality";

/** Shared server-side generation; never expose this as a client server action. */
export async function generateStoredChatImage(props: {
  threadId: string;
  userMessage: string;
  prompt: string;
  size?: string;
  quality?: string;
  signal?: AbortSignal;
}): Promise<{ fileName: string; revised_prompt: string } | { error: string }> {
  const prompt = buildFaithfulImagePrompt(props.userMessage, props.prompt, "generate");
  if (!prompt || prompt.length > 32000) return { error: "Invalid image prompt length" };
  try {
    const response = await OpenAIDALLEInstance().images.generate({
      model: process.env.AZURE_OPENAI_DALLE_API_DEPLOYMENT_NAME!,
      prompt,
      size: normalizeGptImageSize(props.size),
      quality: normalizeGptImageQuality(props.quality),
    }, { signal: props.signal });
    if (!response.data?.[0]?.b64_json) return { error: "Invalid image API response" };
    const fileName = `${uniqueId()}.png`;
    const buffer = Buffer.from(response.data[0].b64_json, "base64");
    for (const name of [fileName, "__base__.png", "__latest__.png"]) {
      const stored = await UploadImageToStore(props.threadId, name, buffer);
      if (stored.status !== "OK") return { error: "Generated image storage failed" };
    }
    return { fileName, revised_prompt: prompt };
  } catch {
    console.error("[image-generation] generation or storage failed");
    return { error: "There was an error creating or storing the image." };
  }
}
