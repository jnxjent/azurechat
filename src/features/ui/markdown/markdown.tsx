// src/features/ui/markdown/markdown.tsx
import Markdoc from "@markdoc/markdoc";
import React, { FC } from "react";
import { Citation } from "./citation";
import { CodeBlock } from "./code-block";
import { citationConfig } from "./config";
import { MarkdownProvider } from "./markdown-context";
import { Paragraph } from "./paragraph";

interface Props {
  content: string;
  onCitationClick: (
    previousState: any,
    formData: FormData
  ) => Promise<JSX.Element>;
}

/**
 * 1. GPT が吐いた HTML/Markdown リンクを「素の URL 一個」に潰す。
 *    ただし Markdown 画像 `![](URL)` はそのまま残す。
 */
function simplifyLinks(raw: string): string {
  if (!raw) return "";

  let s = raw;

  // (A) HTML アンカー: <a href="URL">何でも</a> → URL
  s = s.replace(
    /<a\s+[^>]*href="(https?:\/\/[^"]+)"[^>]*>[\s\S]*?<\/a>/gi,
    "$1"
  );

  // (B) Markdown リンク: [何でも](URL) → URL
  //     ただし直前が "!" の場合 (= 画像) は対象外にする
  s = s.replace(
    /(?<!!)\[[^\]]*\]\((https?:\/\/[^\s)]+)\)/g,
    "$1"
  );

  return s;
}

/**
 * 2. 素の http / https を Markdown リンクや画像に変換。
 *    - /api/images/?t=...&img=...png/jpg/... → 画像: ![](URL)
 *    - それ以外 → リンク: [URL](URL)
 *    なお、既存の Markdown 画像 `![](URL)` は保護する。
 */
function autoLinkUrls(raw: string): string {
  if (!raw) return "";

  // まず Markdown 画像 `![](URL)` を一時退避して保護する
  const imagePattern = /!\[[^\]]*\]\((https?:\/\/[^\s)]+)\)/g;
  const placeholders: string[] = [];

  let tmp = raw.replace(imagePattern, (match) => {
    const key = `__MARKDOWN_IMAGE_PLACEHOLDER_${placeholders.length}__`;
    placeholders.push(match);
    return key;
  });

  // 素の URL をリンク化 or 画像化
  const urlRegex = /(?<!\]\()https?:\/\/[^\s)]+/g;
  tmp = tmp.replace(urlRegex, (url) => {
    // AzureChat の画像エンドポイントだけは「画像として埋め込み」
    const isImageApi =
      /\/api\/images\/\?t=[^&]+&img=[^&]+\.(png|jpg|jpeg|gif|webp)/i.test(url);

    if (isImageApi) {
      // 画像として埋め込み
      return `\n![](${url})`;
    }

    // 通常のリンクとして扱う
    return `\n[${url}](${url})`;
  });

  // 一時退避していた画像 Markdown を元に戻す
  placeholders.forEach((img, i) => {
    const key = `__MARKDOWN_IMAGE_PLACEHOLDER_${i}__`;
    tmp = tmp.replace(key, img);
  });

  return tmp;
}

export const Markdown: FC<Props> = (props) => {
  const simplified = simplifyLinks(props.content);
  const source = autoLinkUrls(simplified);

  const ast = Markdoc.parse(source);

  const baseNodes = (citationConfig as any).nodes || {};

  // ★ href/title を含めて link ノードを定義しつつ、target/_blank を追加
  const mergedConfig: any = {
    ...citationConfig,
    nodes: {
      ...baseNodes,
      link: {
        ...(baseNodes.link || {}),
        render: "a",
        attributes: {
          ...(baseNodes.link?.attributes || {}),
          href: { type: String }, // ← これがないと href が落ちる
          title: { type: String }, // 任意
          target: { type: String, default: "_blank" },
          rel: { type: String, default: "noopener noreferrer" },
        },
      },
    },
  };

  const content = Markdoc.transform(ast, mergedConfig);

  const WithContext = () => (
    <MarkdownProvider onCitationClick={props.onCitationClick}>
      {Markdoc.renderers.react(content, React, {
        components: { Citation, Paragraph, CodeBlock },
      })}
    </MarkdownProvider>
  );

  return <WithContext />;
};
