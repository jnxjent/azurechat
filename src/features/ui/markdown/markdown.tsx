// src/features/ui/markdown/markdown.tsx
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import React, { FC } from "react";
import { Citation } from "./citation";
import { CodeBlock } from "./code-block";
import { MarkdownProvider } from "./markdown-context";
import { Paragraph } from "./paragraph";

interface Props {
  content: string;
  onCitationClick: (
    previousState: any,
    formData: FormData
  ) => Promise<JSX.Element>;
}

/* ------------------------------------------------------------
 * utils
 * ------------------------------------------------------------ */

function isImageUrl(url: string): boolean {
  const u = (url || "").toLowerCase();
  if (!u) return false;
  if (u.includes("/api/images")) return true;
  return (
    u.endsWith(".png") ||
    u.endsWith(".jpg") ||
    u.endsWith(".jpeg") ||
    u.endsWith(".webp") ||
    u.endsWith(".gif")
  );
}

function isPureTextChildren(children: any): boolean {
  if (children == null) return true;
  if (typeof children === "string" || typeof children === "number") return true;
  if (Array.isArray(children)) {
    return children.every(
      (c) => c == null || typeof c === "string" || typeof c === "number"
    );
  }
  return false;
}

/**
 * Citation表現を react-markdown 用に正規化
 */
function preprocessCitations(src: string): string {
  if (!src) return src;

  // Markdoc citation
  const MARKDOC_RE = /\{%\s*citation\s+items=\[([\s\S]*?)\]\s*\/%\}/gi;
  src = src.replace(MARKDOC_RE, (_all, inner) => {
    const payload = encodeURIComponent(String(inner ?? "").trim());
    return `[引用](citation:${payload})`;
  });

  // 〔Name, Id〕
  const BRACKET_RE = /〔\s*([^,\]\n]+?)\s*,\s*([A-Za-z0-9_-]{10,})\s*〕/g;
  src = src.replace(BRACKET_RE, (_all, name, id) => {
    const safeName = String(name).trim().replace(/"/g, '\\"');
    const safeId = String(id).trim().replace(/"/g, '\\"');
    const inner = `{name:"${safeName}",id:"${safeId}"}`;
    const payload = encodeURIComponent(inner);
    return `[引用](citation:${payload})`;
  });

  return src;
}

function decodeCitationItemsFromHref(
  href: string
): Array<{ name: string; id: string }> | null {
  if (!href || !href.startsWith("citation:")) return null;

  const inner = decodeURIComponent(href.slice("citation:".length)).trim();
  if (!inner) return null;

  const items: Array<{ name: string; id: string }> = [];
  const objRe =
    /name\s*:\s*["']([^"']+)["']\s*,\s*id\s*:\s*["']([^"']+)["']/gi;

  let m: RegExpExecArray | null;
  while ((m = objRe.exec(inner)) !== null) {
    if (m[1] && m[2]) {
      items.push({ name: m[1].trim(), id: m[2].trim() });
    }
  }

  return items.length ? items : null;
}

/**
 * 重要:
 * LLMが「表」を ``` で囲って返してしまうと、react-markdownは <pre> 扱いにして黒い箱になる。
 * そこで、``` フェンスの中身が "GFMテーブルっぽい" 場合だけフェンスを剥がして表に戻す。
 *
 * ★安定化:
 * - CRLF/末尾改行なし でもマッチするようにする
 * - 閉じフェンス直前に改行が無いケースも拾う
 */
function unwrapFencedTables(md: string): string {
  if (!md) return md;

  // ```lang?\n ... \n``` だけでなく、最後が ``` で終わる(末尾改行なし)も拾う
  return md.replace(
    /```[a-zA-Z0-9_-]*\r?\n([\s\S]*?)\r?\n?```/g,
    (all, inner) => {
      const s = String(inner ?? "").trim();
      if (!s) return all;

      const lines = s.split(/\r?\n/).map((l) => l.trim());
      if (lines.length < 2) return all;

      const l1 = lines[0];
      const l2 = lines[1];

      // 2行目が区切りっぽいか（| --- | --- | / |:---|---:| 等）
      const looksLikeTable =
        l1.startsWith("|") &&
        l1.includes("|") &&
        l2.startsWith("|") &&
        /^\|(?:\s*:?-{3,}:?\s*\|)+\s*$/.test(l2);

      if (!looksLikeTable) return all;

      // フェンスを剥がして中身だけ返す（= 表として解釈される）
      return s + "\n";
    }
  );
}

/**
 * 裸の画像URLだけを Markdown画像に変換（最小・安全）
 * ※ 表( GFM )の中身は壊さない
 *
 * ★安定化:
 * - 変換対象は「行がURLだけ」のケースに限定（文中URLまで触らない）
 *   → テーブル・リンク周りへの副作用をさらに減らす
 */
function embedNakedImageUrls(src: string): string {
  if (!src) return src;

  // 壊れた Markdown を正規化
  src = src.replace(
    /!\[[^\]]*\]\(\s*<img[^>]*src=["']([^"']+)["'][^>]*>\s*\)/gi,
    "![]($1)"
  );

  // 行が URL だけの場合のみ画像にする（最小化）
  src = src.replace(/^\s*(https?:\/\/[^\s]+)\s*$/gim, (m, url) => {
    // 末尾の句読点などを除去
    let u = String(url || "").trim();
    while (/[)\],.}。、】【]/.test(u.slice(-1))) u = u.slice(0, -1);
    return isImageUrl(u) ? `![](${u})` : m;
  });

  return src;
}

/* ------------------------------------------------------------
 * Component
 * ------------------------------------------------------------ */

export const Markdown: FC<Props> = (props) => {
  // ★順序重要：
  // 1) citation正規化
  // 2) フェンス内の「表」を剥がして表に戻す（黒い箱対策の本丸）
  // 3) 画像URL正規化（副作用を最小化）
  const step1 = preprocessCitations(props.content);
  const step2 = unwrapFencedTables(step1);
  const content = embedNakedImageUrls(step2);

  return (
    <MarkdownProvider onCitationClick={props.onCitationClick}>
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        urlTransform={(url) => url}
        components={{
          // a はリンク専用（img化しない）
          a: ({ ...linkProps }) => {
            const href = String((linkProps as any).href || "");
            const items = decodeCitationItemsFromHref(href);

            if (items) {
              return <Citation items={items as any} />;
            }

            return (
              <a
                {...(linkProps as any)}
                href={href}
                target="_blank"
                rel="noopener noreferrer"
              />
            );
          },

          p: ({ ...pProps }) => {
            const children = (pProps as any).children;

            if (isPureTextChildren(children)) {
              const { children: _ignored, ...rest } = pProps as any;
              return <Paragraph {...rest}>{children}</Paragraph>;
            }

            return <p {...(pProps as any)} />;
          },

          img: ({ ...imgProps }) => (
            <img
              {...(imgProps as any)}
              loading="lazy"
              style={{ maxWidth: "100%", height: "auto" }}
            />
          ),

          code: ({ className, children, ...codeProps }) => {
            const match = /language-(\w+)/.exec(className || "");
            const language = match ? match[1] : "";

            if (!language) {
              return (
                <code className={className} {...(codeProps as any)}>
                  {children}
                </code>
              );
            }

            const codeText = String(children).replace(/\n$/, "");
            return (
              <CodeBlock language={language} {...({} as any)}>
                {codeText}
              </CodeBlock>
            );
          },

          table: ({ ...tableProps }) => (
            <table className="markdown-table" {...(tableProps as any)} />
          ),
        }}
      >
        {content}
      </ReactMarkdown>
    </MarkdownProvider>
  );
};
