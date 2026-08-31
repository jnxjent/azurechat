"use server";
import "server-only";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { GetBlob, UploadBlob } from "@/features/common/services/azure-storage";

const IMAGE_CONTAINER_NAME = "images";
const PENDING_PPTX_EDIT_BLOB_NAME = "__pending_pptx_edit__.json";
const DEFAULT_PENDING_EDIT_TTL_MS = 30 * 60 * 1000;

export type PendingPptxEdit = {
  version: 1;
  fileUrl: string;
  instruction: string;
  targetPages?: number[];
  targetItemCount?: number;
  whiteBase: true;
  requiresImage: boolean;
  imageFileName?: string;
  imageSavedAt?: string;
  waitingFor: "accentColor";
  createdAt: string;
  consumedAt?: string;
};

function blobPath(threadId: string): string {
  return `${threadId}/${PENDING_PPTX_EDIT_BLOB_NAME}`;
}

function pendingEditTtlMs(): number {
  const configured = Number(process.env.PPTX_PENDING_EDIT_TTL_MS);
  return Number.isFinite(configured) && configured > 0
    ? configured
    : DEFAULT_PENDING_EDIT_TTL_MS;
}

export async function SavePendingPptxEdit(
  threadId: string,
  edit: Omit<PendingPptxEdit, "version" | "createdAt" | "consumedAt">
): Promise<ServerActionResponse<PendingPptxEdit>> {
  const state: PendingPptxEdit = {
    ...edit,
    version: 1,
    createdAt: new Date().toISOString(),
  };
  const saved = await UploadBlob(
    IMAGE_CONTAINER_NAME,
    blobPath(threadId),
    Buffer.from(JSON.stringify(state), "utf-8")
  );
  if (saved.status !== "OK") return saved as ServerActionResponse<PendingPptxEdit>;
  return { status: "OK", response: state };
}

export async function LoadPendingPptxEdit(
  threadId: string
): Promise<PendingPptxEdit | null> {
  const result = await GetBlob(IMAGE_CONTAINER_NAME, blobPath(threadId));
  if (result.status !== "OK" || !result.response) return null;

  try {
    const state = JSON.parse(
      await new Response(result.response as any).text()
    ) as PendingPptxEdit;
    const createdAtMs = Date.parse(state.createdAt);
    if (
      state.version !== 1 ||
      state.waitingFor !== "accentColor" ||
      state.whiteBase !== true ||
      state.consumedAt ||
      !Number.isFinite(createdAtMs) ||
      Date.now() - createdAtMs > pendingEditTtlMs()
    ) {
      return null;
    }
    return state;
  } catch (error) {
    console.warn("[pptx-pending-edit] Failed to load state:", error);
    return null;
  }
}

export async function ConsumePendingPptxEdit(threadId: string): Promise<void> {
  const state = await LoadPendingPptxEdit(threadId);
  if (!state) return;
  state.consumedAt = new Date().toISOString();
  const saved = await UploadBlob(
    IMAGE_CONTAINER_NAME,
    blobPath(threadId),
    Buffer.from(JSON.stringify(state), "utf-8")
  );
  if (saved.status !== "OK") {
    console.warn("[pptx-pending-edit] Failed to consume state");
  }
}

