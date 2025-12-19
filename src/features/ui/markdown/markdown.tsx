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

/** 画像っぽいURL判定（拡張子 or /api/images） */
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

/**
 * 裸の画像URLを Markdown画像に変換（はめ込み用）
 * - 行単独のURLも
 * - 文中のURLも（ただし誤爆を避けて画像っぽいものだけ）
 */
function embedNakedImageUrls(src: string): string {
  if (!src) return src;

  // 1) 行がURLだけの場合（前後空白OK）
  src = src.replace(
    /^\s*(https?:\/\/[^\s]+)\s*$/gim,
    (m, url) => (isImageUrl(url) ? `![](${url})` : m)
  );

  // 2) 文中のURL（スペース/改行区切り想定）
  // 末尾の ) ] } . , などが付くケースを軽くケア
  src = src.replace(/(https?:\/\/[^\s<>"']+)/g, (m) => {
    let url = m;
    // よくある末尾句読点や括弧を落とす
    while (/[)\],.}。、】【]/.test(url.slice(-1))) url = url.slice(0, -1);
    if (!isImageUrl(url)) return m;
    return `![](${url})`;
  });

  return src;
}

/**
 * Citation表現を “react-markdownで扱える形” に正規化する（最小・安全）
 * - Markdoc: {% citation items=[ ... ] /%}
 * - Bracket: 〔Name, 0AbC...〕
 *
 * NOTE: dotAll(s) フラグは es2018+ が必要なため使用しない。
 *       [\s\S]*? で改行を含む非貪欲マッチを実現する。
 */
function preprocessCitations(src: string): string {
  if (!src) return src;

  // Markdoc citation → [引用](citation:...)
  const MARKDOC_RE = /\{%\s*citation\s+items=\[([\s\S]*?)\]\s*\/%\}/gi;
  src = src.replace(MARKDOC_RE, (_all, inner) => {
    const payload = encodeURIComponent(String(inner ?? "").trim());
    return `[引用](citation:${payload})`;
  });

  // 〔Name, Id〕 → [引用](citation:...)
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

  const encoded = href.slice("citation:".length);
  const inner = decodeURIComponent(encoded || "").trim();
  if (!inner) return null;

  const items: Array<{ name: string; id: string }> = [];
  const objRe =
    /name\s*:\s*["']([^"']+)["']\s*,\s*id\s*:\s*["']([^"']+)["']/gi;

  let m: RegExpExecArray | null;
  while ((m = objRe.exec(inner)) !== null) {
    const name = (m[1] || "").trim();
    const id = (m[2] || "").trim();
    if (name && id) items.push({ name, id });
  }

  return items.length ? items : null;
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

export const Markdown: FC<Props> = (props) => {
  // 1) citationを正規化
  // 2) 裸の画像URLを ![]() に変換して “はめ込み” 寄せ
  const content = embedNakedImageUrls(preprocessCitations(props.content));

  return (
    <MarkdownProvider onCitationClick={props.onCitationClick}>
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        urlTransform={(url) => url}
        components={{
          a: ({ node, ...linkProps }) => {
            const href = String((linkProps as any).href || "");
            const items = decodeCitationItemsFromHref(href);

            if (items) {
              return <Citation items={items as any} />;
            }

            // ★リンク先が画像なら「リンク」ではなく「画像として表示」
            if (href && isImageUrl(href)) {
              return (
                <img
                  src={href}
                  alt={String((linkProps as any).children || "")}
                  loading="lazy"
                  style={{ maxWidth: "100%", height: "auto" }}
                />
              );
            }

            return (
              <a {...linkProps} target="_blank" rel="noopener noreferrer" />
            );
          },

          p: ({ node, ...pProps }) => {
            const children = (pProps as any).children;

            if (isPureTextChildren(children)) {
              // Paragraph は children 必須なので明示的に渡す
              const { children: _ignored, ...rest } = pProps as any;
              return <Paragraph {...rest}>{children}</Paragraph>;
            }

            return <p {...pProps} />;
          },

          img: ({ node, ...imgProps }) => (
            <img
              {...imgProps}
              loading="lazy"
              style={{ maxWidth: "100%", height: "auto" }}
            />
          ),

          code: ({ node, className, children, ...codeProps }) => {
            // react-markdown の型では `inline` が無い場合があるため、
            // className の language- でブロック/インラインを判定する。
            const match = /language-(\w+)/.exec(className || "");
            const language = match ? match[1] : "";

            // language 指定が無ければインライン扱い
            if (!language) {
              return (
                <code className={className} {...(codeProps as any)}>
                  {children}
                </code>
              );
            }

            // language 指定があればブロック扱い（CodeBlockは children で本文を渡す）
            const codeText = String(children).replace(/\n$/, "");
            return (
              <CodeBlock language={language} {...({} as any)}>
                {codeText}
              </CodeBlock>
            );
          },

          table: ({ node, ...tableProps }) => (
            <table className="markdown-table" {...tableProps} />
          ),
        }}
      >
        {content}
      </ReactMarkdown>
    </MarkdownProvider>
  );
};
