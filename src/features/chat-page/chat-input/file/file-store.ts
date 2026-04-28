"use client";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { showError, showSuccess } from "@/features/globals/global-message-store";
import { proxy, useSnapshot } from "valtio";
import { IndexDocuments } from "../../chat-services/azure-ai-search/azure-ai-search";
import {
  CrackDocument,
  CreateChatDocument,
  UploadDocument,
} from "../../chat-services/chat-document-service";
import { chatStore } from "../../chat-store";

// File -> base64（data:... を除いた純base64）
async function fileToBase64(file: File): Promise<string> {
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.onload = () => {
      const result = String(reader.result || "");
      const comma = result.indexOf(",");
      resolve(comma >= 0 ? result.slice(comma + 1) : result);
    };
    reader.readAsDataURL(file);
  });
}

type UploadScope = "common" | "personal";

/**
 * uploadScope の正規化
 * - common / personal を正式値とする
 * - 旧値 cp は personal として吸収
 * - 未指定も personal 扱い
 */
function normalizeUploadScope(value?: string | null): UploadScope {
  const v = (value ?? "").toLowerCase().trim();

  if (v === "common") return "common";
  if (v === "personal") return "personal";
  if (v === "cp") return "personal";

  return "personal";
}

/**
 * フロント側の希望 uploadScope
 * 優先順位:
 * 1) props.uploadScope
 * 2) 旧 env 値（互換）
 * 3) personal
 */
function resolveRequestedUploadScope(propsUploadScope?: string): UploadScope {
  const fromProps = normalizeUploadScope(propsUploadScope);
  if (propsUploadScope) return fromProps;

  const fromEnv = normalizeUploadScope(process.env.NEXT_PUBLIC_SL_DEPT);
  return fromEnv;
}

// SharePointへ同期コピー（失敗したらthrow）
async function publishToSharePoint(file: File, uploadScope: UploadScope) {
  const fileBase64 = await fileToBase64(file);

  const r = await fetch("/api/sl/publish", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileName: file.name,
      fileBase64,
      uploadScope,
    }),
  });

  const json = await r.json().catch(() => ({}));

  if (!r.ok || !json?.ok) {
    const msg =
      json?.error || `SharePoint publish failed (status=${r.status})`;
    throw new Error(msg);
  }

  return json as {
    ok: true;
    dept?: string;
    uploadScope?: UploadScope;
    isSharePointEnabled?: boolean;
    webUrl?: string;
    name?: string;
    spItemId?: string | null;
  };
}

class FileStore {
  public uploadButtonLabel: string = "";

  public async onFileChange(props: {
    formData: FormData;
    chatThreadId: string;
    uploadScope?: string; // トグルから渡す
  }) {
    const { formData, chatThreadId, uploadScope: requestedScopeRaw } = props;

    try {
      chatStore.updateLoading("file upload");

      formData.append("id", chatThreadId);
      const file = formData.get("file") as unknown as File | null;

      if (!file) {
        showError("No file selected.");
        return;
      }

      // フロント側の希望値（最終決定はサーバ側）
      const requestedUploadScope = resolveRequestedUploadScope(requestedScopeRaw);

      this.uploadButtonLabel = "Processing document";
      formData.append("fileName", file.name);

      const uploadResponse = await UploadDocument(formData);
      const crackingResponse = await CrackDocument(formData);

      if (crackingResponse.status === "OK" && uploadResponse.status === "OK") {
        let actualDept = "";
        let actualIsSlDoc = false;
        let actualUploadScope: UploadScope = requestedUploadScope;
        // ★ SP webUrl を保持する変数（SPアップ成功時のみ設定される）
        let spWebUrl: string | undefined;
        // ★ SP item ID（移動後もindexと紐付けるための不変ID）
        let spItemId: string | null = null;

        try {
          this.uploadButtonLabel = "Syncing to SharePoint";

          const sp = await publishToSharePoint(file, requestedUploadScope);

          // ★ SP webUrl を取得（Index の fileUrl に使う）
          if (sp.isSharePointEnabled === true && !sp.webUrl) {
            throw new Error("SharePoint publish succeeded but webUrl was empty.");
          }
          spWebUrl = sp.webUrl;
          spItemId = sp.spItemId ?? null;
          actualDept = String(sp.dept ?? "").toLowerCase().trim();
          actualUploadScope = normalizeUploadScope(
            sp.uploadScope ?? requestedUploadScope
          );
          actualIsSlDoc = sp.isSharePointEnabled === true;

          showSuccess({
            title: "SharePoint sync",
            description: sp.webUrl
              ? `Synced to SharePoint (${actualDept || "unknown"} / ${actualUploadScope}): ${sp.webUrl}`
              : `Synced to SharePoint (${actualDept || "unknown"} / ${actualUploadScope}): ${sp.name || file.name}`,
          });
        } catch (e: any) {
          const msg = String(e?.message ?? e);

          // SP非対応部署は新仕様上ありうる
          if (
            msg.includes("not SharePoint-enabled") ||
            msg.includes("Use Blob upload flow instead")
          ) {
            showSuccess({
              title: "Blob upload",
              description:
                `${file.name} uploaded to Blob. ` +
                `This department is not SharePoint-enabled, so SharePoint sync was skipped.`,
            });

            // 従来どおり Blob インデックス用
            actualDept = (
              process.env.NEXT_PUBLIC_SL_DEPT_NON_SP ?? "others"
            ).toLowerCase().trim();
            actualIsSlDoc = false;
            actualUploadScope = "personal";
          } else {
            showError(`SharePoint sync failed (index skipped): ${msg}`);
            return;
          }
        }

        // SharePoint / Blob の最終状態確定後に Index 作成
        let index = 0;
        const documentIndexResponses: Array<ServerActionResponse<boolean>> = [];
        const searchableFileUrl = spWebUrl ?? uploadResponse.response;
        const effectiveFileUrl = spWebUrl ?? uploadResponse.response;

        for (const doc of crackingResponse.response) {
          this.uploadButtonLabel = `Indexing document [${index + 1}]/[${
            crackingResponse.response.length
          }]`;

          // ★ SPアップ成功時は SP webUrl を使う。それ以外は従来通り Blob URL。
          // SL documents use SP webUrl as fileUrl so that deleted SP files
          // are no longer reachable, avoiding stale Blob links in search results.
          const indexResponses = await IndexDocuments(
            file.name,
            searchableFileUrl,
            [doc],
            chatThreadId,
            actualDept,
            actualIsSlDoc,
            actualUploadScope,
            effectiveFileUrl,
            spItemId
          );

          documentIndexResponses.push(...indexResponses);
          index++;
        }

        const allDocumentsIndexed = documentIndexResponses.every(
          (r) => r.status === "OK"
        );

        if (allDocumentsIndexed) {
          this.uploadButtonLabel = file.name + " loaded";

          const response = await CreateChatDocument(file.name, chatThreadId);

          if (response.status === "OK") {
            showSuccess({
              title: "File upload",
              description: `${file.name} uploaded successfully.`,
            });
          } else {
            showError(response.errors.map((e) => e).join("\n"));
          }
        } else {
          const errors: Array<string> = [];
          documentIndexResponses.forEach((r) => {
            if (r.status === "ERROR") {
              errors.push(...r.errors.map((e) => e.message));
            }
          });

          showError(
            "Looks like not all documents were indexed\n" + errors.join("\n")
          );
        }
      } else {
        const crackingErrors =
          (crackingResponse as any)?.errors
            ?.map((e: any) => e.message)
            .join("\n") || "Failed to process document.";
        showError(crackingErrors);
      }
    } catch (error) {
      showError(String(error));
    } finally {
      this.uploadButtonLabel = "";
      chatStore.updateLoading("idle");
    }
  }
}

export const fileStore = proxy(new FileStore());

export function useFileStore() {
  return useSnapshot(fileStore);
}
