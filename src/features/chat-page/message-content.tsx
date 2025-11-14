import { Markdown } from "@/features/ui/markdown/markdown";
import { FunctionSquare } from "lucide-react";
import React from "react";
import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "../ui/accordion";
import { RecursiveUI } from "../ui/recursive-ui";
import { CitationAction } from "./citation/citation-action";

interface MessageContentProps {
  message: {
    role: string;
    content: string;
    name: string;
    message?: string;         // BEの .message
    message_format?: string;  // "markdown" 等
    multiModalImage?: string;
  };
}

// 1) URL/Markdownっぽさの自動判定（保険）
const looksLikeMarkdownOrLinks = (s: string) =>
  /\[[^\]]+\]\([^)]+\)/.test(s) || /(https?:\/\/|www\.)\S+/.test(s);

// 2) 「名前：URL」→ Markdownリンク化、残りの生URLは <URL> で保険
const normalizeToMarkdown = (src: string): string => {
  if (!src) return "";

  const mdLinkPattern = /\[[^\]]+\]\([^)]+\)/; // 既に markdown リンクなら触らない

  const lines = src.split(/\r?\n/).map((line) => {
    // 既に Markdown リンクが含まれる行はそのまま返す（副作用防止）
    if (mdLinkPattern.test(line)) return line;

    // 「- ラベル：URL」/「ラベル：URL」/ 全角・半角コロン両対応 → 正しく Markdown 化
    const m = line.match(/^\s*(-\s*)?([^\n:：]+?)\s*[：:]\s*(https?:\/\/\S+)\s*$/);
    if (m) {
      const dash = m[1] ? "- " : "";
      const label = m[2].trim();
      const url = m[3].trim(); // ★ 末尾に / を付けない
      return `${dash}[${label}](${url})`;
    }

    // 残った生URLだけ <URL> 化。ただし markdown のリンク先URLは除外したいので
    // 直前文字が '(' または '<' の場合はスキップ
    return line.replace(/https?:\/\/\S+/g, (u, offset) => {
      const prev = line[offset - 1] || "";
      if (prev === "(" || prev === "<") return u; // リンク先 or 既に <URL> の一部
      return `<${u}>`;
    });
  });

  return lines.join("\n");
};

const MessageContent: React.FC<MessageContentProps> = ({ message }) => {
  // デバッグ出力
  console.log("[DEBUG] message", {
    message: (message as any).message,
    content: message.content,
    message_format: (message as any).message_format,
    role: message.role,
  });

  if (message.role === "assistant" || message.role === "user") {
    const rawText = (message as any).message ?? message.content ?? "";
    // 先に正規化して Markdown に寄せる
    const mdText = normalizeToMarkdown(rawText);

    const isMd =
      (message.message_format || "").toLowerCase() === "markdown" ||
      looksLikeMarkdownOrLinks(mdText); // ← 正規化後テキストで判定

    return (
      <>
        {isMd ? (
          <Markdown content={mdText} onCitationClick={CitationAction} />
        ) : (
          <div className="whitespace-pre-wrap break-words">{rawText}</div>
        )}
        {message.multiModalImage && <img src={message.multiModalImage} alt="" />}
      </>
    );
  }

  if (message.role === "tool" || message.role === "function") {
    return (
      <div className="py-3">
        <Accordion type="multiple" className="bg-background rounded-md border p-2">
          <AccordionItem value="item-1" className="">
            <AccordionTrigger className="text-sm py-1 items-center gap-2">
              <div className="flex gap-2 items-center">
                <FunctionSquare size={18} strokeWidth={1.4} className="text-muted-foreground" />
                {" "}Show {message.name}{" "}{message.name === "tool" ? "output" : "function"}
              </div>
            </AccordionTrigger>
            <AccordionContent>
              <RecursiveUI documentField={toJson(message.content)} />
            </AccordionContent>
          </AccordionItem>
        </Accordion>
      </div>
    );
  }

  return null;
};

const toJson = (value: string) => {
  try { return JSON.parse(value); } catch { return value; }
};

export default MessageContent;
