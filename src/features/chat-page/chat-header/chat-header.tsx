"use client";

import { ExtensionModel } from "@/features/extensions-page/extension-services/models";
import { CHAT_DEFAULT_PERSONA } from "@/features/theme/theme-config";
import { RefreshCw, VenetianMask } from "lucide-react";
import { FC, useState } from "react";
import { ChatDocumentModel, ChatThreadModel } from "../chat-services/models";
import { DocumentDetail } from "./document-detail";
import { ExtensionDetail } from "./extension-detail";
import { PersonaDetail } from "./persona-detail";

interface Props {
  chatThread: ChatThreadModel;
  chatDocuments: Array<ChatDocumentModel>;
  extensions: Array<ExtensionModel>;
  isAdmin?: boolean;
}

type SyncRow = {
  deleted?: number;
  error?: string;
  skipped?: string;
  urlUpdated?: number;
};

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
      const res = await fetch("/api/sl/sync-check?apply=true", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: "{}",
      });

      const data = await res.json();
      if (data.ok) {
        const allRows = Object.entries(data.results as Record<string, SyncRow>);
        const errorDepts = allRows
          .filter(([, row]) => row.error)
          .map(([dept]) => dept);
        const skippedDepts = allRows
          .filter(([, row]) => row.skipped)
          .map(([dept, row]) => `${dept}:${row.skipped}`);
        const okRows = allRows
          .filter(([, row]) => !row.error)
          .map(([, row]) => row);

        const updated = okRows.reduce(
          (sum, row) => sum + (row.urlUpdated ?? 0),
          0
        );
        const deleted = okRows.reduce(
          (sum, row) => sum + (row.deleted ?? 0),
          0
        );

        const parts = [`更新:${updated}件`, `削除:${deleted}件`];
        if (errorDepts.length > 0) {
          parts.push(`エラー:${errorDepts.join(",")}`);
        }
        if (skippedDepts.length > 0) {
          parts.push(`スキップ:${skippedDepts.join(",")}`);
        }

        setSyncResult(parts.join(" "));
      } else {
        setSyncResult(`エラー: ${data.error}`);
      }
    } catch (e: any) {
      setSyncResult(`エラー: ${e.message}`);
    } finally {
      setSyncing(false);
      setTimeout(() => setSyncResult(null), 5000);
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
