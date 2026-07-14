/**
 * DeckSpec Blob保存/取得ユーティリティ
 * コンテナ: "pptx-specs"（非公開 — 生成スライドデータ・designInstructionを含むため）
 * ブロブ名: "{blobKey}.spec.json"
 */
import { BlobServiceClient } from "@azure/storage-blob";
import type { DeckSpec } from "@/types/deck-spec";

const SPEC_CONTAINER = "pptx-specs";
const PPTX_CONTAINER = "pptx";

function pptxBlobSvc(): BlobServiceClient | null {
  const acc = (process.env.AZURE_STORAGE_ACCOUNT_NAME ?? "").trim();
  const key = (process.env.AZURE_STORAGE_ACCOUNT_KEY ?? "").trim();
  if (!acc || !key) return null;
  return BlobServiceClient.fromConnectionString(
    `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
}

/** @returns 保存成功なら true、失敗なら false（呼び出し元でポインター更新前に確認可能） */
export async function saveDeckSpec(blobKey: string, spec: DeckSpec): Promise<boolean> {
  const svc = pptxBlobSvc();
  if (!svc || !blobKey?.trim()) return false;
  try {
    const cc = svc.getContainerClient(SPEC_CONTAINER);
    // private アクセス（access 指定なし = private）
    await cc.createIfNotExists();
    const sidecarKey = `${blobKey}.spec.json`;
    await cc.getBlockBlobClient(sidecarKey).uploadData(
      Buffer.from(JSON.stringify({ ...spec, savedAt: new Date().toISOString() })),
      { blobHTTPHeaders: { blobContentType: "application/json" } }
    );
    console.log(`[deckSpec] saved ${sidecarKey} deckId=${spec.deckId} rev=${spec.revision}`);
    return true;
  } catch (e) {
    console.error("[deckSpec] saveDeckSpec failed:", String((e as any)?.message ?? e).slice(0, 160));
    return false;
  }
}

export async function loadDeckSpec(blobKey: string): Promise<DeckSpec | null> {
  const svc = pptxBlobSvc();
  if (!svc || !blobKey?.trim()) return null;
  try {
    const cc = svc.getContainerClient(SPEC_CONTAINER);
    const buf = await cc.getBlockBlobClient(`${blobKey}.spec.json`).downloadToBuffer();
    const spec = JSON.parse(buf.toString()) as DeckSpec;
    console.log(`[deckSpec] loaded deckId=${spec.deckId} rev=${spec.revision} slides=${spec.slides.length}`);
    return spec;
  } catch {
    return null;
  }
}

/**
 * PPTX の公開URLからblobKeyを抽出してDeckSpecを読み込む。
 * URLが自社AzureStorageアカウント＋pptxコンテナであることを検証する。
 */
/**
 * 生成イラスト（data URI）を pptx-specs コンテナに別Blobとして保存する。
 * DeckSpec JSON への直接埋め込みを避けてサイズ肥大化を防ぐ。
 * @returns 保存した Blob キー（"{pptxBlobKey}.illustration"）、失敗時は null
 */
export async function saveIllustrationBlob(pptxBlobKey: string, dataUri: string): Promise<string | null> {
  const svc = pptxBlobSvc();
  if (!svc || !pptxBlobKey?.trim() || !dataUri) return null;
  try {
    const cc = svc.getContainerClient(SPEC_CONTAINER);
    await cc.createIfNotExists();
    const illustKey = `${pptxBlobKey}.illustration`;
    await cc.getBlockBlobClient(illustKey).uploadData(
      Buffer.from(dataUri, "utf-8"),
      { blobHTTPHeaders: { blobContentType: "text/plain" } }
    );
    console.log(`[deckSpec] illustration saved ${illustKey}`);
    return illustKey;
  } catch (e) {
    console.warn("[deckSpec] saveIllustrationBlob failed:", String((e as any)?.message ?? e).slice(0, 160));
    return null;
  }
}

/** イラスト Blob を読み込んで data URI を返す。存在しない場合や失敗時は null。 */
export async function loadIllustrationBlob(illustrationKey: string): Promise<string | null> {
  const svc = pptxBlobSvc();
  if (!svc || !illustrationKey?.trim()) return null;
  try {
    const cc = svc.getContainerClient(SPEC_CONTAINER);
    const buf = await cc.getBlockBlobClient(illustrationKey).downloadToBuffer();
    const dataUri = buf.toString("utf-8");
    if (!dataUri.startsWith("data:")) return null;
    return dataUri;
  } catch {
    return null;
  }
}

/**
 * 自システム生成PPTXのBlobに「DeckSpec必須マーカー」メタデータを付与する。
 * DeckSpec保存成功後に呼ぶ。失敗時はポインター更新をスキップするため boolean を返す。
 */
export async function markPptxAsOurs(blobKey: string, deckId: string): Promise<boolean> {
  const svc = pptxBlobSvc();
  if (!svc || !blobKey?.trim()) return false;
  try {
    const cc = svc.getContainerClient(PPTX_CONTAINER);
    await cc.getBlockBlobClient(blobKey).setMetadata({
      azurechat_deckspec_required: "true",
      deck_id: deckId,
    });
    return true;
  } catch (e) {
    console.error("[deckSpec] markPptxAsOurs failed:", String((e as any)?.message ?? e).slice(0, 160));
    return false;
  }
}

/**
 * PPTXのBlobメタデータを確認して自システム生成PPTXかどうかを返す。
 * マーカーがない場合（外部PPTX・旧バージョン生成）は false → Python互換モードを許可。
 */
export async function checkPptxIsOurs(pptxUrl: string): Promise<boolean> {
  if (!pptxUrl?.trim()) return false;
  const acc = (process.env.AZURE_STORAGE_ACCOUNT_NAME ?? "").trim();
  if (!acc) return false;
  try {
    const parsed = new URL(pptxUrl);
    const expectedHost = `${acc}.blob.core.windows.net`;
    if (parsed.hostname !== expectedHost) return false;
    const pathParts = parsed.pathname.split("/").filter(Boolean);
    if (pathParts[0] !== PPTX_CONTAINER) return false;
    const blobKey = decodeURIComponent(pathParts[pathParts.length - 1] ?? "");
    if (!blobKey) return false;
    const svc = pptxBlobSvc();
    if (!svc) return false;
    const cc = svc.getContainerClient(PPTX_CONTAINER);
    const props = await cc.getBlockBlobClient(blobKey).getProperties();
    return props.metadata?.azurechat_deckspec_required === "true";
  } catch {
    return false;
  }
}

export async function loadDeckSpecForUrl(pptxUrl: string): Promise<DeckSpec | null> {
  if (!pptxUrl?.trim()) return null;
  const acc = (process.env.AZURE_STORAGE_ACCOUNT_NAME ?? "").trim();
  if (!acc) return null;
  try {
    const parsed = new URL(pptxUrl);
    // ホスト名が自社ストレージアカウントでなければスキップ
    const expectedHost = `${acc}.blob.core.windows.net`;
    if (parsed.hostname !== expectedHost) return null;
    // コンテナが "pptx" であることを確認
    const pathParts = parsed.pathname.split("/").filter(Boolean);
    if (pathParts[0] !== "pptx") return null;
    const blobKey = decodeURIComponent(pathParts[pathParts.length - 1] ?? "");
    if (!blobKey || !/\.pptx$/i.test(blobKey)) return null;
    return await loadDeckSpec(blobKey);
  } catch {
    return null;
  }
}
