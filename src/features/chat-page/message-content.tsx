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
    multiModalImage?: string;
  };
}

/**
 * SF連携の「レコードURL: https://...」をクリック可能なMarkdownに変換する。
 * - citation 行（{% citation ... %}）は一切触らないようにしてある。
 */
const normalizeContent = (src: string): string => {
  if (!src) return "";

  // 行ごとに処理
  const lines = src.split(/\r?\n/).map((line) => {
    // 引用タグは絶対に触らない
    if (line.includes("{% citation")) return line;

    // 「レコードURL: https://...」/「URL: https://...」などを検出
    const m = line.match(
      /^(.*?(レコードURL|URL))\s*[:：]\s*(https?:\/\/\S+)\s*$/
    );
    if (m) {
      const labelPart = m[1].trim(); // "レコードURL" など
      const url = m[3].trim();
      // ラベル行 + URLを別行でMarkdownリンク化
      return `${labelPart}:\n[${url}](${url})`;
    }

    return line;
  });

  return lines.join("\n");
};

const MessageContent: React.FC<MessageContentProps> = ({ message }) => {
  if (message.role === "assistant" || message.role === "user") {
    // ★ SFリンクだけMarkdownに寄せる（citationはそのまま）
    const normalized = normalizeContent(message.content);

    return (
      <>
        <Markdown
          content={normalized}
          onCitationClick={CitationAction}
        />
        {message.multiModalImage && <img src={message.multiModalImage} alt="" />}
      </>
    );
  }

  if (message.role === "tool" || message.role === "function") {
    return (
      <div className="py-3">
        <Accordion
          type="multiple"
          className="bg-background rounded-md border p-2"
        >
          <AccordionItem value="item-1" className="">
            <AccordionTrigger className="text-sm py-1 items-center gap-2">
              <div className="flex gap-2 items-center">
                <FunctionSquare
                  size={18}
                  strokeWidth={1.4}
                  className="text-muted-foreground"
                />{" "}
                Show {message.name}{" "}
                {message.name === "tool" ? "output" : "function"}
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
  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
};

export default MessageContent;
