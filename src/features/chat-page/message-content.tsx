import { Markdown } from "@/features/ui/markdown/markdown";
import { normalizePhoneForTel, splitTextWithPhones } from "@/lib/linkifyPhone";
import { FunctionSquare } from "lucide-react";
import React, { useEffect, useRef } from "react";
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
 * 1行のプレーンテキスト内の電話番号を Markdown リンク [xxx](tel:xxx) に変換する
 * - 既存の Markdown リンク行っぽいものは極力そのまま
 * - URL を含む行は誤爆回避のためそのまま
 */
const linkifyPhoneLine = (line: string): string => {
  if (!line) return line;

  if (/https?:\/\//i.test(line)) return line;
  if (/\[[^\]]+\]\([^)]+\)/.test(line)) return line;

  const parts = splitTextWithPhones(line);

  return parts
    .map((part) => {
      if (part.type === "phone") {
        const tel = normalizePhoneForTel(part.value);
        return `[${part.value}](tel:${tel})`;
      }
      return part.value;
    })
    .join("");
};

/**
 * GFMテーブルの区切り行か判定
 * 例:
 * | --- | --- |
 * |:---|---:|
 */
const isTableSeparatorRow = (trimmed: string): boolean => {
  if (!trimmed.startsWith("|")) return false;
  return /^\|[\s:\-|\u3000]+\|?$/.test(trimmed);
};

/**
 * テーブル行の各セルだけを安全に処理する
 * - 既存リンク入りセルは触らない
 * - URL入りセルは触らない
 * - 電話番号だけ [xxx](tel:xxx) にする
 */
const linkifyPhoneInTableRow = (line: string): string => {
  const trimmed = line.trim();

  if (!trimmed.startsWith("|")) return line;
  if (isTableSeparatorRow(trimmed)) return line;

  const hasLeadingPipe = trimmed.startsWith("|");
  const hasTrailingPipe = trimmed.endsWith("|");

  let inner = trimmed;
  if (hasLeadingPipe) inner = inner.slice(1);
  if (hasTrailingPipe) inner = inner.slice(0, -1);

  const cells = inner.split("|");

  const newCells = cells.map((cell) => {
    const rawCell = cell;
    const cellTrimmed = rawCell.trim();

    if (!cellTrimmed) return rawCell;
    if (/https?:\/\//i.test(cellTrimmed)) return rawCell;
    if (/\[[^\]]+\]\([^)]+\)/.test(cellTrimmed)) return rawCell;

    const linked = linkifyPhoneLine(cellTrimmed);

    // 元の左右余白をなるべく維持
    const leftSpace = rawCell.match(/^\s*/)?.[0] ?? "";
    const rightSpace = rawCell.match(/\s*$/)?.[0] ?? "";

    return `${leftSpace}${linked}${rightSpace}`;
  });

  return `|${newCells.join("|")}|`;
};

/**
 * SF連携の「レコードURL: https://...」をクリック可能なMarkdownに変換する。
 * さらに、通常行や表セルに含まれる電話番号を tel: リンク化する。
 *
 * - citation 行（{% citation ... %}）は一切触らない
 * - fenced code block 内は触らない
 * - URL末尾に混入しがちな記号（| ) ] . , 等）を除去
 */
const normalizeContent = (src: string): string => {
  if (!src) return "";

  let inCodeBlock = false;

  const lines = src.split(/\r?\n/).map((line) => {
    const trimmed = line.trim();

    if (trimmed.startsWith("```")) {
      inCodeBlock = !inCodeBlock;
      return line;
    }

    if (inCodeBlock) return line;
    if (line.includes("{% citation")) return line;

    const looksLikeTableRow =
      trimmed.startsWith("|") && trimmed.includes("|") && !trimmed.startsWith("|-");

    if (looksLikeTableRow) {
      return linkifyPhoneInTableRow(line);
    }

    const m = line.match(
      /^(.*?(?:レコードURL|URL|画像URL))\s*[:：]\s*(https?:\/\/[^\s|)\]]+)\s*$/i
    );
    if (m) {
      const labelPart = m[1].trim();
      let url = (m[2] || "").trim();

      while (/[|)\].,}。、】【]$/.test(url)) {
        url = url.slice(0, -1);
      }

      if (!url) return line;

      return `${labelPart}:\n[${url}](${url})`;
    }

    return linkifyPhoneLine(line);
  });

  return lines.join("\n");
};

/* ------------------------------------------------------------------ */
/* Canvas 用ユーティリティ                                            */
/* ------------------------------------------------------------------ */

type CanvasStyle = {
  fontFamily: string;
  fontSize: number;
  color: string;
  strokeColor: string;
  strokeWidth: number;
  textAlign: CanvasTextAlign;
};

const parseCanvasStyle = (hint: string): CanvasStyle => {
  const base: CanvasStyle = {
    fontFamily: "Yu Gothic",
    fontSize: 48,
    color: "#ffffff",
    strokeColor: "rgba(0,0,0,0.6)",
    strokeWidth: 4,
    textAlign: "center",
  };

  if (!hint) return base;

  const s = hint.replace(/\s+/g, "").toLowerCase();

  if (s.includes("メイリオ")) base.fontFamily = "Meiryo";
  else if (s.includes("游ゴシック") || s.includes("游ｺﾞｼｯｸ"))
    base.fontFamily = "Yu Gothic";
  else if (s.includes("ゴシック")) base.fontFamily = "Yu Gothic";
  else if (s.includes("明朝")) base.fontFamily = "Yu Mincho";

  if (s.includes("特大") || s.includes("めちゃ大") || s.includes("ドーン")) {
    base.fontSize = 72;
  } else if (s.includes("大きめ") || s.includes("大きく") || s.includes("大きい")) {
    base.fontSize = 60;
  } else if (s.includes("小さめ") || s.includes("小さい") || s.includes("控えめ")) {
    base.fontSize = 36;
  } else if (s.includes("普通") || s.includes("標準")) {
    base.fontSize = 48;
  }

  if (s.includes("白文字") || s.includes("白")) base.color = "#ffffff";
  if (s.includes("黒文字") || s.includes("黒")) base.color = "#000000";
  if (s.includes("赤文字") || s.includes("赤")) base.color = "red";
  if (s.includes("青文字") || s.includes("青")) base.color = "blue";
  if (s.includes("黄色") || s.includes("黄")) base.color = "yellow";

  if (s.includes("左寄せ") || s.includes("左揃え") || s.includes("左端")) {
    base.textAlign = "left";
  } else if (s.includes("右寄せ") || s.includes("右揃え") || s.includes("右端")) {
    base.textAlign = "right";
  } else if (
    s.includes("中央") ||
    s.includes("真ん中") ||
    s.includes("センター")
  ) {
    base.textAlign = "center";
  }

  return base;
};

const extractOverlayText = (content: string): string => {
  if (!content) return "";

  const m1 = content.match(/『([^』]+)』/);
  if (m1) return m1[1].trim();

  const m2 = content.match(/「([^」]+)」/);
  if (m2) return m2[1].trim();

  const lines = content.split(/\r?\n/).map((l) => l.trim());
  for (let i = lines.length - 1; i >= 0; i--) {
    if (lines[i]) return lines[i];
  }

  return content.trim();
};

type ImageWithCanvasOverlayProps = {
  src: string;
  alt?: string;
  overlayText?: string;
  styleHint?: string;
};

const ImageWithCanvasOverlay: React.FC<ImageWithCanvasOverlayProps> = ({
  src,
  alt,
  overlayText,
  styleHint,
}) => {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);

  useEffect(() => {
    if (!src || !canvasRef.current) return;

    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    const img = new Image();
    img.crossOrigin = "anonymous";

    img.onload = () => {
      const w = img.naturalWidth || img.width;
      const h = img.naturalHeight || img.height;
      if (!w || !h) return;

      canvas.width = w;
      canvas.height = h;

      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

      const rawText = (overlayText || "").trim();
      if (!rawText) return;

      const text = rawText.slice(0, 80);
      const style = parseCanvasStyle(styleHint || "");

      ctx.font = `${style.fontSize}px "${style.fontFamily}", sans-serif`;
      ctx.textAlign = style.textAlign;
      ctx.textBaseline = "middle";

      ctx.lineWidth = style.strokeWidth;
      ctx.strokeStyle = style.strokeColor;
      ctx.fillStyle = style.color;

      let x = canvas.width / 2;
      if (style.textAlign === "left") x = 40;
      if (style.textAlign === "right") x = canvas.width - 40;
      const y = canvas.height - 60;

      ctx.strokeText(text, x, y);
      ctx.fillText(text, x, y);
    };

    img.onerror = () => {
      console.error("ImageWithCanvasOverlay: failed to load image:", src);
    };

    img.src = src;
  }, [src, overlayText, styleHint]);

  return (
    <div className="mt-3 flex justify-center">
      <canvas
        ref={canvasRef}
        aria-label={alt}
        style={{ maxWidth: "100%", height: "auto", borderRadius: 8 }}
      />
    </div>
  );
};

/* ------------------------------------------------------------------ */

const MessageContent: React.FC<MessageContentProps> = ({ message }) => {
  if (message.role === "assistant" || message.role === "user") {
    const normalized = normalizeContent(message.content);

    const hasImage = !!message.multiModalImage;
    const overlayText = hasImage ? extractOverlayText(message.content) : "";

    return (
      <>
        <Markdown content={normalized} onCitationClick={CitationAction} />
        {hasImage && (
          <ImageWithCanvasOverlay
            key={`${message.multiModalImage}-${overlayText}`}
            src={message.multiModalImage!}
            alt=""
            overlayText={overlayText}
            styleHint={message.content}
          />
        )}
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
                />
                Show {message.name} {message.name === "tool" ? "output" : "function"}
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