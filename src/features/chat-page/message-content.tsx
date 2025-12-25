import { Markdown } from "@/features/ui/markdown/markdown";
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
 * SF連携の「レコードURL: https://...」をクリック可能なMarkdownに変換する。
 * - citation 行（{% citation ... %}）は一切触らない
 * - ★FIX: GFMテーブル行は触らない（黒い箱の主因）
 * - ★FIX: URL末尾に混入しがちな記号（| ) ] . , 等）を除去
 */
const normalizeContent = (src: string): string => {
  if (!src) return "";

  const lines = src.split(/\r?\n/).map((line) => {
    // 引用タグは絶対に触らない
    if (line.includes("{% citation")) return line;

    // ★FIX(1): GFMテーブルっぽい行は触らない（最重要）
    // 例: "| 商談名 | ... | [開く](https://...) |"
    const trimmed = line.trim();
    const looksLikeTableRow =
      trimmed.startsWith("|") && trimmed.includes("|") && !trimmed.startsWith("|-");
    if (looksLikeTableRow) return line;

    // 「レコードURL: https://...」/「URL: https://...」などを検出
    // ★FIX(2): URL は「空白/|/)/]」などで止める（\S+ をやめる）
    const m = line.match(
      /^(.*?(?:レコードURL|URL|画像URL))\s*[:：]\s*(https?:\/\/[^\s|)\]]+)\s*$/i
    );
    if (m) {
      const labelPart = m[1].trim(); // "レコードURL" など
      let url = (m[2] || "").trim();

      // ★FIX(3): 念のため末尾の句読点・区切りを落とす
      while (/[|)\].,}。、】【]$/.test(url)) {
        url = url.slice(0, -1);
      }

      // 空になったら元に戻す
      if (!url) return line;

      // ラベル行 + URLを別行でMarkdownリンク化
      return `${labelPart}:\n[${url}](${url})`;
    }

    return line;
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

/**
 * メッセージ本文から「どんなフォント・色・サイズか」をゆるく推定
 * 例: 「メイリオ」「ゴシック」「明朝」「大きめ」「小さめ」「白文字」「赤文字」など
 */
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

  // フォント
  if (s.includes("メイリオ")) base.fontFamily = "Meiryo";
  else if (s.includes("游ゴシック") || s.includes("游ｺﾞｼｯｸ"))
    base.fontFamily = "Yu Gothic";
  else if (s.includes("ゴシック")) base.fontFamily = "Yu Gothic";
  else if (s.includes("明朝")) base.fontFamily = "Yu Mincho";

  // サイズ
  if (s.includes("特大") || s.includes("めちゃ大") || s.includes("ドーン")) {
    base.fontSize = 72;
  } else if (s.includes("大きめ") || s.includes("大きく") || s.includes("大きい")) {
    base.fontSize = 60;
  } else if (s.includes("小さめ") || s.includes("小さい") || s.includes("控えめ")) {
    base.fontSize = 36;
  } else if (s.includes("普通") || s.includes("標準")) {
    base.fontSize = 48;
  }

  // 色
  if (s.includes("白文字") || s.includes("白")) base.color = "#ffffff";
  if (s.includes("黒文字") || s.includes("黒")) base.color = "#000000";
  if (s.includes("赤文字") || s.includes("赤")) base.color = "red";
  if (s.includes("青文字") || s.includes("青")) base.color = "blue";
  if (s.includes("黄色") || s.includes("黄")) base.color = "yellow";

  // 揃え
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

/**
 * メッセージ本文から「実際に画像に載せる文字」をざっくり抽出。
 * - 「『…』」 or 「「…」」があればその中身を優先
 * - なければ全文ではなく、末尾行を使う などゆるめに
 */
const extractOverlayText = (content: string): string => {
  if (!content) return "";

  // 『…』優先
  const m1 = content.match(/『([^』]+)』/);
  if (m1) return m1[1].trim();

  // 次に「…」
  const m2 = content.match(/「([^」]+)」/);
  if (m2) return m2[1].trim();

  // フォールバック: 最後の非空行
  const lines = content.split(/\r?\n/).map((l) => l.trim());
  for (let i = lines.length - 1; i >= 0; i--) {
    if (lines[i]) return lines[i];
  }

  return content.trim();
};

type ImageWithCanvasOverlayProps = {
  src: string;
  alt?: string;
  /** 画像に乗せるテキスト（空なら純粋なプレビューとして描画だけ） */
  overlayText?: string;
  /** 自然言語ヒント（フォント・色・サイズ・揃えなど） */
  styleHint?: string;
};

/**
 * クライアント側 Canvas で画像＋日本語テキストを描画するコンポーネント。
 * - サーバー側フォントに依存しないので「豆腐」問題を回避できる
 * - overlayText が空なら単なる画像プレビューとして動く
 */
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
      // キャンバスサイズを画像に合わせる
      const w = img.naturalWidth || img.width;
      const h = img.naturalHeight || img.height;
      if (!w || !h) return;

      canvas.width = w;
      canvas.height = h;

      // まずはキャンバス全体をクリア
      ctx.clearRect(0, 0, canvas.width, canvas.height);

      // 画像を描画（毎回ここからスタート）
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

      // テキストが無ければここで終了（純粋プレビューとして）
      const rawText = (overlayText || "").trim();
      if (!rawText) return;

      // 過度な長文を防ぐため、ある程度で切る
      const text = rawText.slice(0, 80);

      const style = parseCanvasStyle(styleHint || "");

      ctx.font = `${style.fontSize}px "${style.fontFamily}", sans-serif`;
      ctx.textAlign = style.textAlign;
      ctx.textBaseline = "middle";

      ctx.lineWidth = style.strokeWidth;
      ctx.strokeStyle = style.strokeColor;
      ctx.fillStyle = style.color;

      // 基本は下部付近に描画（揃えは textAlign で制御）
      let x = canvas.width / 2;
      if (style.textAlign === "left") x = 40;
      if (style.textAlign === "right") x = canvas.width - 40;
      const y = canvas.height - 60;

      // 縁取り → 塗り
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
    // ★ SFリンクだけMarkdownに寄せる（citationはそのまま）
    const normalized = normalizeContent(message.content);

    // multimodalImage がある場合だけ Canvas でプレビュー＋文字入れ
    const hasImage = !!message.multiModalImage;
    const overlayText = hasImage ? extractOverlayText(message.content) : "";

    return (
      <>
        <Markdown content={normalized} onCitationClick={CitationAction} />
        {hasImage && (
          <ImageWithCanvasOverlay
            // ★ ここを追加：overlayText が変わるたびに Canvas を作り直す
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
