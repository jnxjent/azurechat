"use client";
import { FC, useMemo, useState, useEffect } from "react";
import { CitationSlider } from "./citation-slider";
import { CitationFileDownload } from "@/features/chat-page/citation/citation-file-download";

interface Citation {
  name: string;
  id: string;
  fileUrl?: string; // ← 追加（RAGから来る場合に拾う）
}

interface Props {
  items: Citation[];
}

export const citation = {
  render: "Citation",
  selfClosing: true,
  attributes: {
    items: { type: Array },
  },
};

export const Citation: FC<Props> = (props: Props) => {
  // 同名ファイルでグルーピング
  const citations = useMemo(() => {
    return props.items.reduce((acc, c) => {
      (acc[c.name] ||= []).push(c);
      return acc;
    }, {} as Record<string, Citation[]>);
  }, [props.items]);

  // id→URL 解決（fileUrl が無ければ API で取得）
  const [resolved, setResolved] = useState<Record<string, string>>({}); // name -> url

  useEffect(() => {
    (async () => {
      const entries = Object.entries(citations);
      const next: Record<string, string> = { ...resolved };

      for (const [name, items] of entries) {
        // 1) どれかに fileUrl があればそれを採用
        const withUrl = items.find((i) => !!i.fileUrl)?.fileUrl;
        if (withUrl) {
          next[name] = withUrl;
          continue;
        }
        // 2) なければ最初の id でダウンロードURLを解決
        if (!next[name]) {
          const form = new FormData();
          form.append("id", items[0].id);
          const url = await CitationFileDownload(form);
          if (url) next[name] = url;
        }
      }
      setResolved(next);
    })();
  }, [citations]); // props.items が変わったら再解決

  const onClickFileName = async (e: React.MouseEvent, fileName: string) => {
    // 既に href を埋めているので、href があれば JS での遷移は不要
    if (!resolved[fileName]) {
      e.preventDefault();
      const formData = new FormData();
      formData.append("id", citations[fileName][0].id);
      const url = await CitationFileDownload(formData);
      if (url) {
        setResolved((s) => ({ ...s, [fileName]: url }));
        window.location.href = url;
      }
    }
  };

  return (
    <div className="interactive-citation p-4 border mt-4 flex flex-col rounded-md gap-2">
      {Object.entries(citations).map(([name, items], groupIndex) => {
        const url = resolved[name]; // 解決済みの URL（or undefined）
        return (
          <div key={groupIndex} className="flex flex-col gap-2">
            <div className="font-semibold text-sm">
              {/* ✅ URL が解決済みなら「本物のリンク」にする（中クリック対応） */}
              <a
                href={url ?? ""}
                onClick={(e) => onClickFileName(e, name)}
                target={url ? "_blank" : undefined}
                rel={url ? "noopener noreferrer" : undefined}
              >
                {name}
              </a>
            </div>

            <div className="flex gap-2">
              {items.map((item, index) => (
                <div key={index}>
                  <CitationSlider index={index + 1} name={item.name} id={item.id} />
                </div>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
};
