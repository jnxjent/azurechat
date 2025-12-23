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
 * 裸の画像URLを Markdown画像に変換（最小・安全）
 *
 * ★FIX(A)
 *   壊れた `![](<img src="...">)` を `![](URL)` に戻す
 *
 * ★FIX(B)
 *   既に Markdown 記法内のURL（![](...), ](...)）は二重変換しない
 */
function embedNakedImageUrls(src: string): string {
  if (!src) return src;

  // ------------------------------------------------------------
  // ★FIX(A): 壊れた Markdown を正規化
  // ------------------------------------------------------------
  // ![](<img src="URL" ...>) → ![](URL)
  src = src.replace(
    /!\[[^\]]*\]\(\s*<img[^>]*src=["']([^"']+)["'][^>]*>\s*\)/gi,
    "![]($1)"
  );

  // ------------------------------------------------------------
  // 1) 行が URL だけの場合
  // ------------------------------------------------------------
  src = src.replace(
    /^\s*(https?:\/\/[^\s]+)\s*$/gim,
    (m, url) => (isImageUrl(url) ? `![](${url})` : m)
  );

  // ------------------------------------------------------------
  // 2) 文中の裸URL（ただし Markdown 内は除外）
  // ------------------------------------------------------------
  src = src.replace(/https?:\/\/[^\s<>"']+/g, function (match) {
    const offset = arguments[arguments.length - 2] as number;
    const whole = arguments[arguments.length - 1] as string;

    // 直前が ]( or ![]( なら Markdown 内なので触らない
    const before = whole.slice(Math.max(0, offset - 4), offset);
    if (before.includes("](") || before.includes("![](")) {
      return match;
    }

    let url = match;
    // 末尾の句読点などを除去
    while (/[)\],.}。、】【]/.test(url.slice(-1))) {
      url = url.slice(0, -1);
    }

    if (!isImageUrl(url)) return match;

    return `![](${url})`;
  });

  return src;
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
  // ★順序重要：citation → 画像正規化
  const content = embedNakedImageUrls(
    preprocessCitations(props.content)
  );

  return (
    <MarkdownProvider onCitationClick={props.onCitationClick}>
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        urlTransform={(url) => url}
        components={{
          a: ({ ...linkProps }) => {
            const href = String((linkProps as any).href || "");
            const items = decodeCitationItemsFromHref(href);

            if (items) {
              return <Citation items={items as any} />;
            }

            // ★画像URLはリンクではなく画像として表示
            if (href && isImageUrl(href)) {
              return (
                <img
                  src={href}
                  loading="lazy"
                  style={{ maxWidth: "100%", height: "auto" }}
                />
              );
            }

            return (
              <a {...linkProps} target="_blank" rel="noopener noreferrer" />
            );
          },

          p: ({ ...pProps }) => {
            const children = (pProps as any).children;

            if (isPureTextChildren(children)) {
              const { children: _ignored, ...rest } = pProps as any;
              return <Paragraph {...rest}>{children}</Paragraph>;
            }

            return <p {...pProps} />;
          },

          img: ({ ...imgProps }) => (
            <img
              {...imgProps}
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
            <table className="markdown-table" {...tableProps} />
          ),
        }}
      >
        {content}
      </ReactMarkdown>
    </MarkdownProvider>
  );
};
