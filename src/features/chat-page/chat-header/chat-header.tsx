"use client";
import { ExtensionModel } from "@/features/extensions-page/extension-services/models";
import { CHAT_DEFAULT_PERSONA } from "@/features/theme/theme-config";
import { VenetianMask, RefreshCw } from "lucide-react";
import { FC, useState } from "react";
import { ChatDocumentModel, ChatThreadModel } from "../chat-services/models";
import { DocumentDetail } from "./document-detail";
import { ExtensionDetail } from "./extension-detail";
import { PersonaDetail } from "./persona-detail";

interface Props {
  chatThread: ChatThreadModel;
  chatDocuments: Array<ChatDocumentModel>;
  extensions: Array<ExtensionModel>;
  isAdmin?: boolean; // 管理者のみ同期ボタンを表示
}

export const ChatHeader: FC<Props> = (props) => {
  const [syncing, setSyncing] = useState(false);
  const [syncResult, setSyncResult] = useState<string | null>(null);

  const persona =
    props.chatThread.personaMessageTitle === "" ||
    props.chatThread.personaMessageTitle === undefined
      ? CHAT_DEFAULT_PERSONA
      : props.chatThread.personaMessageTitle;

  const handleSync = async () => {
    setSyncing(true);
    setSyncResult(null);
    try {
      const res = await fetch("/api/sl/sync-check", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: "{}",
      });
      const data = await res.json();
      if (data.ok) {
        // 削除件数の合計を集計
        const total = Object.values(data.results as Record<string, any>)
          .filter((r) => !r.error)
          .reduce((sum, r) => sum + (r.deleted ?? 0), 0);
        setSyncResult(`✅ 同期完了（${total}件削除）`);
      } else {
        setSyncResult(`❌ エラー: ${data.error}`);
      }
    } catch (e: any) {
      setSyncResult(`❌ エラー: ${e.message}`);
    } finally {
      setSyncing(false);
      // 3秒後にメッセージを消す
      setTimeout(() => setSyncResult(null), 3000);
    }
  };

  return (
    <div className="bg-background border-b flex items-center py-2">
      <div className="container max-w-3xl flex justify-between items-center">
        <div className="flex flex-col">
          <span>{props.chatThread.name}</span>
          <span className="text-sm text-muted-foreground flex gap-1 items-center">
            <VenetianMask size={18} />
            {persona}
          </span>
        </div>
        <div className="flex gap-2 items-center">
          {/* 管理者のみIndex同期ボタンを表示 */}
          {props.isAdmin && (
            <div className="flex items-center gap-2">
              {syncResult && (
                <span className="text-xs text-muted-foreground">{syncResult}</span>
              )}
              <button
                onClick={handleSync}
                disabled={syncing}
                title="SharePoint → Index同期"
                className="flex items-center gap-1 text-xs px-2 py-1 rounded border border-input bg-background hover:bg-accent disabled:opacity-50"
              >
                <RefreshCw size={14} className={syncing ? "animate-spin" : ""} />
                {syncing ? "同期中..." : "Index同期"}
              </button>
            </div>
          )}
          <PersonaDetail chatThread={props.chatThread} />
          <DocumentDetail chatDocuments={props.chatDocuments} />
          <ExtensionDetail
            disabled={props.chatDocuments.length !== 0}
            extensions={props.extensions}
            installedExtensionIds={props.chatThread.extension}
            chatThreadId={props.chatThread.id}
          />
        </div>
      </div>
    </div>
  );
};
