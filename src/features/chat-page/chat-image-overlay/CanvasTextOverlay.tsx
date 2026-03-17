// src/features/chat-page/chat-image-overlay/CanvasTextOverlay.tsx
"use client";

import React, { useEffect, useRef, useState } from "react";

type AlignH = "left" | "center" | "right";
type AlignV = "top" | "middle" | "bottom";

/**
 * NL から指定できるスタイルプリセット
 * - "formal": ビジネス年賀状など（白＋薄い影）
 * - "pop":    ポップな目立つ書体（黄色＋影）
 * - "handwriting": 手書き風（擬似）
 * - "simple": 小さめ、控えめ
 */
export type OverlayStylePreset =
  | "formal"
  | "pop"
  | "handwriting"
  | "simple";

export interface CanvasTextOverlayProps {
  /** もとの画像URL（/api/images?... や /generated/...） */
  imageUrl: string;

  /** 重ねるテキスト（日本語OK） */
  text: string;

  /** プリセット名（NL から渡してくる想定） */
  stylePreset?: OverlayStylePreset;

  /** 追加の細かい指定（あれば上書き） */
  fontFamily?: string;
  fontSizePx?: number;
  fillColor?: string;
  strokeColor?: string;
  strokeWidth?: number;

  align?: AlignH; // left / center / right
  vAlign?: AlignV; // top / middle / bottom

  /** 下部マージン（vAlign = bottom のときのみ有効） */
  bottomMargin?: number;

  /** 描画完了後の dataURL を外に返したい場合（任意） */
  onRenderedDataUrl?: (dataUrl: string) => void;

  /** className など、ラッパー <div> に付けたいもの（任意） */
  className?: string;
}

/**
 * NL で指定された stylePreset → Canvas 用のスタイルにマッピング
 */
function resolvePresetStyle(
  preset: OverlayStylePreset | undefined
): {
  fontFamily: string;
  fontSizePx: number;
  fillColor: string;
  strokeColor: string;
  strokeWidth: number;
} {
  switch (preset) {
    case "formal":
      return {
        fontFamily: `"Noto Sans JP", "Meiryo", system-ui, sans-serif`,
        fontSizePx: 52,
        fillColor: "white",
        strokeColor: "rgba(0,0,0,0.6)",
        strokeWidth: 4,
      };
    case "pop":
      return {
        fontFamily: `"Noto Sans JP", "Meiryo", system-ui, sans-serif`,
        fontSizePx: 56,
        fillColor: "#ffeb3b",
        strokeColor: "rgba(0,0,0,0.8)",
        strokeWidth: 5,
      };
    case "handwriting":
      return {
        fontFamily: `"Yu Mincho", "Hiragino Mincho ProN", "Noto Serif JP", serif`,
        fontSizePx: 54,
        fillColor: "white",
        strokeColor: "rgba(0,0,0,0.6)",
        strokeWidth: 4,
      };
    case "simple":
    default:
      return {
        fontFamily: `"Noto Sans JP", "Meiryo", system-ui, sans-serif`,
        fontSizePx: 40,
        fillColor: "white",
        strokeColor: "rgba(0,0,0,0.5)",
        strokeWidth: 3,
      };
  }
}

/**
 * 画像の上に Canvas で日本語テキストを重ねて表示するコンポーネント。
 * - サーバー側の sharp / フォント問題を回避し、ブラウザのフォントを利用。
 */
export const CanvasTextOverlay: React.FC<CanvasTextOverlayProps> = (props) => {
  const {
    imageUrl,
    text,
    stylePreset = "formal",
    fontFamily,
    fontSizePx,
    fillColor,
    strokeColor,
    strokeWidth,
    align = "center",
    vAlign = "bottom",
    bottomMargin = 80,
    onRenderedDataUrl,
    className,
  } = props;

  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [renderedDataUrl, setRenderedDataUrl] = useState<string | null>(null);
  const [isDrawing, setIsDrawing] = useState(false);

  useEffect(() => {
    if (!imageUrl || !text) return;
    const canvas = canvasRef.current;
    if (!canvas) return;

    let cancelled = false;
    const img = new Image();

    // CORS 対策：同一オリジン想定だが一応
    img.crossOrigin = "anonymous";
    img.src = imageUrl;

    setIsDrawing(true);

    img.onload = () => {
      if (cancelled) return;

      const ctx = canvas.getContext("2d");
      if (!ctx) return;

      const dpr = window.devicePixelRatio || 1;
      const displayWidth = img.width;
      const displayHeight = img.height;

      canvas.width = displayWidth * dpr;
      canvas.height = displayHeight * dpr;
      canvas.style.width = `${displayWidth}px`;
      canvas.style.height = `${displayHeight}px`;

      ctx.setTransform(dpr, 0, 0, dpr, 0, 0); // 高DPI対応
      ctx.clearRect(0, 0, displayWidth, displayHeight);

      // ベース画像描画
      ctx.drawImage(img, 0, 0, displayWidth, displayHeight);

      // スタイル決定（プリセット＋オーバーライド）
      const baseStyle = resolvePresetStyle(stylePreset);
      const usedFontFamily = fontFamily || baseStyle.fontFamily;
      const usedFontSize = fontSizePx ?? baseStyle.fontSizePx;
      const usedFillColor = fillColor || baseStyle.fillColor;
      const usedStrokeColor = strokeColor || baseStyle.strokeColor;
      const usedStrokeWidth = strokeWidth ?? baseStyle.strokeWidth;

      ctx.font = `${usedFontSize}px ${usedFontFamily}`;
      ctx.textBaseline = "middle";
      ctx.fillStyle = usedFillColor;
      ctx.strokeStyle = usedStrokeColor;
      ctx.lineWidth = usedStrokeWidth;

      // 横位置
      let x: number;
      if (align === "left") {
        ctx.textAlign = "left";
        x = 40;
      } else if (align === "right") {
        ctx.textAlign = "right";
        x = displayWidth - 40;
      } else {
        ctx.textAlign = "center";
        x = displayWidth / 2;
      }

      // 縦位置
      let y: number;
      if (vAlign === "top") {
        y = 40 + usedFontSize / 2;
      } else if (vAlign === "middle") {
        y = displayHeight / 2;
      } else {
        y = displayHeight - bottomMargin;
      }

      // 軽い影
      ctx.shadowColor = "rgba(0,0,0,0.5)";
      ctx.shadowBlur = 6;
      ctx.shadowOffsetX = 2;
      ctx.shadowOffsetY = 2;

      // アウトライン → 塗り
      if (usedStrokeWidth > 0) {
        ctx.strokeText(text, x, y);
      }
      ctx.fillText(text, x, y);

      // dataURL 取得
      const url = canvas.toDataURL("image/png");
      setRenderedDataUrl(url);
      onRenderedDataUrl?.(url);
      setIsDrawing(false);
    };

    img.onerror = (err) => {
      console.error("[CanvasTextOverlay] image load error:", err);
      setIsDrawing(false);
    };

    return () => {
      cancelled = true;
    };
  }, [
    imageUrl,
    text,
    stylePreset,
    fontFamily,
    fontSizePx,
    fillColor,
    strokeColor,
    strokeWidth,
    align,
    vAlign,
    bottomMargin,
    onRenderedDataUrl,
  ]);

  return (
    <div className={className}>
      {/* Canvas 自身は hidden でも良いが、デバッグのため一旦表示 */}
      <canvas
        ref={canvasRef}
        style={{ maxWidth: "100%", display: renderedDataUrl ? "none" : "block" }}
      />
      {/* ユーザーには dataURL を <img> で見せる */}
      {renderedDataUrl && (
        <img
          src={renderedDataUrl}
          alt={text}
          style={{ maxWidth: "100%", display: "block" }}
        />
      )}
      {isDrawing && !renderedDataUrl && (
        <div style={{ fontSize: 12, color: "#888" }}>画像をレンダリング中…</div>
      )}
    </div>
  );
};
