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

function normalizeDept(value?: string | null): string {
  const d = (value ?? "").toLowerCase().trim();
  return d || "cp";
}

/**
 * アップロード先 dept の決定
 * 優先順位: props.dept（トグル選択） > NEXT_PUBLIC_SL_DEPT（従来） > "cp"
 */
function resolveUploadDept(propsDept?: string): string {
  const fromProps = normalizeDept(propsDept);
  if (propsDept && fromProps) return fromProps;

  const fromEnv = normalizeDept(process.env.NEXT_PUBLIC_SL_DEPT);
  return fromEnv || "cp";
}

/**
 * SL文書かどうかの判定
 * NEXT_PUBLIC_SL_DEPT_NON_SP に設定された dept は SP未対応 → isSlDoc=false
 * それ以外は SP対応部署 → isSlDoc=true
 */
function resolveIsSlDoc(uploadDept: string): boolean {
  const nonSpDept = (process.env.NEXT_PUBLIC_SL_DEPT_NON_SP ?? "others")
    .toLowerCase()
    .trim();
  return uploadDept !== nonSpDept;
}

// SharePointへ同期コピー
// 戻り値: null = SL無効（スキップ）, オブジェクト = 成功, throw = エラー
async function publishToSharePoint(
  file: File,
  dept: string
): Promise<{ ok: true; dept?: string; webUrl?: string; name?: string } | null> {
  const fileBase64 = await fileToBase64(file);

  const r = await fetch("/api/sl/publish", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileName: file.name,
      fileBase64,
      dept,
    }),
  });

  // ★ SL無効時（NEXT_PUBLIC_SL_ENABLED=false）はサーバが404を返す
  // → エラーにせずnullを返してスキップ
  if (r.status === 404) {
    console.log("[SL] publish disabled, skipping SharePoint sync");
    return null;
  }

  const json = await r.json().catch(() => ({}));

  if (!r.ok || !json?.ok) {
    const msg = json?.error || `SharePoint publish failed (status=${r.status})`;
    throw new Error(msg);
  }

  return json as {
    ok: true;
    dept?: string;
    webUrl?: string;
    name?: string;
  };
}

class FileStore {
  public uploadButtonLabel: string = "";

  public async onFileChange(props: {
    formData: FormData;
    chatThreadId: string;
    dept?: string;
  }) {
    const { formData, chatThreadId, dept: requestedDept } = props;

    try {
      chatStore.updateLoading("file upload");

      formData.append("id", chatThreadId);
      const file = formData.get("file") as unknown as File | null;

      if (!file) {
        showError("No file selected.");
        return;
      }

      const requestedUploadDept = resolveUploadDept(requestedDept);
      const requestedIsSlDoc = resolveIsSlDoc(requestedUploadDept);

      this.uploadButtonLabel = "Processing document";
      formData.append("fileName", file.name);

      const uploadResponse = await UploadDocument(formData);
      const crackingResponse = await CrackDocument(formData);

      if (crackingResponse.status === "OK" && uploadResponse.status === "OK") {
        let actualDept = requestedUploadDept;
        let actualIsSlDoc = requestedIsSlDoc;
        let sp:
          | { ok: true; dept?: string; webUrl?: string; name?: string }
          | null
          | undefined;

        if (requestedIsSlDoc) {
          try {
            this.uploadButtonLabel = "Syncing to SharePoint";
            sp = await publishToSharePoint(file, requestedUploadDept);

            if (sp === null) {
              // ★ SL無効 → SP syncスキップ、通常indexingへ
              actualDept = requestedUploadDept;
              actualIsSlDoc = false;
              console.log("[SL] SP sync skipped (disabled). Proceeding with normal indexing.");
            } else {
              // SP sync成功
              actualDept = normalizeDept(sp.dept || requestedUploadDept);
              actualIsSlDoc = resolveIsSlDoc(actualDept);

              showSuccess({
                title: "SharePoint sync",
                description: sp.webUrl
                  ? `Synced to SharePoint (${actualDept}): ${sp.webUrl}`
                  : `Synced to SharePoint (${actualDept}): ${sp.name || file.name}`,
              });
            }
          } catch (e: any) {
            showError(
              `SharePoint sync failed (index skipped): ${String(
                e?.message ?? e
              )}`
            );
            return;
          }
        }

        // SharePoint 確定後の dept で Index 作成
        let index = 0;
        const documentIndexResponses: Array<ServerActionResponse<boolean>> = [];

        for (const doc of crackingResponse.response) {
          this.uploadButtonLabel = `Indexing document [${index + 1}]/[${
            crackingResponse.response.length
          }]`;

          const indexResponses = await IndexDocuments(
            file.name,
            uploadResponse.response,
            [doc],
            chatThreadId,
            actualDept,
            actualIsSlDoc
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