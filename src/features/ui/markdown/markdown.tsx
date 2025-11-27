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
 *
 *   - <a href="https://...">任意のテキスト</a> → https://...
 *   - [任意のテキスト](https://...)         → https://...
 *
 *   テキスト側が URL でも日本語でも、とにかく
 *   「https://...」だけ残す。
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
  s = s.replace(
    /\[[^\]]*\]\((https?:\/\/[^\s)]+)\)/g,
    "$1"
  );

  return s;
}

/**
 * 2. 素の http / https を Markdown リンク [URL](URL) に変換しつつ、
 *    できるだけ URL を単独行に出す。
 *
 *   - 例: 「画像URL: https://...」 → 「画像URL:\n[https://...](https://...)」
 */
function autoLinkUrls(raw: string): string {
  if (!raw) return "";

  // 素の URL（既に [text](url) にはなっていないもの）を検知
  const urlRegex = /(?<!\]\()https?:\/\/[^\s)]+/g;

  return raw.replace(urlRegex, (url) => `\n[${url}](${url})`);
}

export const Markdown: FC<Props> = (props) => {
  // 1) まず HTML/Markdown のリンク表現を「URL だけ」に潰す
  const simplified = simplifyLinks(props.content);

  // 2) その URL を [URL](URL) に変換（ついでに前に改行を入れる）
  const source = autoLinkUrls(simplified);

  const ast = Markdoc.parse(source);

  // citationConfig に image ノードだけ追加する
  const mergedConfig: any = {
    ...citationConfig,
    nodes: {
      ...(citationConfig as any).nodes,
      image: {
        render: "img",
        attributes: {
          src: { type: String, required: true },
          alt: { type: String },
          title: { type: String },
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
