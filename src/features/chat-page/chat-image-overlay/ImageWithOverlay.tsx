// src/features/chat-page/chat-image-overlay/ImageWithOverlay.tsx
"use client";

import React from "react";
import {
  CanvasTextOverlay,
  type OverlayStylePreset,
} from "./CanvasTextOverlay";

type AlignH = "left" | "center" | "right";
type AlignV = "top" | "middle" | "bottom";

export interface ImageWithOverlayProps {
  /** ベース画像の URL（/api/images や /generated/...） */
  imageUrl: string;

  /** 普通の alt テキスト */
  alt?: string;

  /** 重ねるテキスト（なければ単なる <img> として表示） */
  overlayText?: string;

  /**
   * NL っぽいスタイル指定
   * 例：
   *   - "フォーマルに"
   *   - "ポップに大きく"
   *   - "手書き風で"
   *   - "控えめに小さめで"
   */
  overlayStyleHint?: string;

  /** 明示的な位置指定があれば使う（なければヒントから推定） */
  align?: AlignH;
  vAlign?: AlignV;
  bottomMargin?: number;

  className?: string;
}

/**
 * NL のヒント文字列から、プリセットや配置をざっくり決める
 */
function mapHintToStyle(
  hint: string | undefined
): {
  preset: OverlayStylePreset;
  align: AlignH;
  vAlign: AlignV;
  bottomMargin: number;
} {
  const h = (hint || "").toLowerCase();

  // かな・漢字もざっくり見る
  const isFormal =
    h.includes("フォーマル") ||
    h.includes("ビジネス") ||
    h.includes("年賀状") ||
    h.includes("かたい");
  const isPop =
    h.includes("ポップ") ||
    h.includes("カジュアル") ||
    h.includes("目立つ") ||
    h.includes("派手");
  const isHand =
    h.includes("手書き") || h.includes("手書き風") || h.includes("筆");
  const isSmall = h.includes("小さめ") || h.includes("控えめ");

  let preset: OverlayStylePreset = "formal";
  if (isPop) preset = "pop";
  else if (isHand) preset = "handwriting";
  else if (isSmall) preset = "simple";

  // 位置のヒント
  let align: AlignH = "center";
  let vAlign: AlignV = "bottom";
  let bottomMargin = 80;

  if (h.includes("左上")) {
    align = "left";
    vAlign = "top";
  } else if (h.includes("右上")) {
    align = "right";
    vAlign = "top";
  } else if (h.includes("左下")) {
    align = "left";
    vAlign = "bottom";
  } else if (h.includes("右下")) {
    align = "right";
    vAlign = "bottom";
  } else if (h.includes("中央")) {
    align = "center";
    if (h.includes("中央下") || h.includes("下部中央")) {
      vAlign = "bottom";
    } else if (h.includes("中央上") || h.includes("上部中央")) {
      vAlign = "top";
    } else {
      vAlign = "middle";
    }
  }

  // 「かなり下に」とか「少し上に」などを見てもいいが、いったん簡単に
  if (h.includes("かなり下") || h.includes("余白多め")) {
    bottomMargin = 120;
  } else if (h.includes("少し上")) {
    bottomMargin = 60;
  }

  return { preset, align, vAlign, bottomMargin };
}

/**
 * 既存の <img> を置き換える入り口。
 * - overlayText がない → 旧来どおり <img>
 * - overlayText がある → CanvasTextOverlay で日本語テキストを合成して表示
 */
export const ImageWithOverlay: React.FC<ImageWithOverlayProps> = (props) => {
  const {
    imageUrl,
    alt,
    overlayText,
    overlayStyleHint,
    align,
    vAlign,
    bottomMargin,
    className,
  } = props;

  // テキスト指定なしなら、単純に <img> で返す
  if (!overlayText) {
    return (
      <img
        src={imageUrl}
        alt={alt || ""}
        className={className}
        style={{ maxWidth: "100%", display: "block" }}
      />
    );
  }

  // NL ヒントからスタイル推定
  const mapped = mapHintToStyle(overlayStyleHint);
  const usedAlign = align || mapped.align;
  const usedVAlign = vAlign || mapped.vAlign;
  const usedBottomMargin = bottomMargin ?? mapped.bottomMargin;

  return (
    <CanvasTextOverlay
      imageUrl={imageUrl}
      text={overlayText}
      stylePreset={mapped.preset}
      align={usedAlign}
      vAlign={usedVAlign}
      bottomMargin={usedBottomMargin}
      className={className}
    />
  );
};
