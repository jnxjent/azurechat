"use client";
import { FC, useCallback, useState } from "react";
import { CitationSlider } from "./citation-slider";
import { CitationFileDownload } from "@/features/chat-page/citation/citation-file-download";

interface Citation {
  name: string;
  id: string;
}

interface Props {
  items: Citation[];
}

export const citation = {
  render: "Citation",
  selfClosing: true,
  attributes: {
    items: {
      type: Array,
    },
  },
};

export const Citation: FC<Props> = (props: Props) => {
  const citations = props.items.reduce((acc, citation) => {
    const { name } = citation;
    if (!acc[name]) {
      acc[name] = [];
    }
    acc[name].push(citation);
    return acc;
  }, {} as Record<string, Citation[]>);

  const [urlCache, setUrlCache] = useState<Record<string, string>>({});

  const resolveUrl = useCallback(
    async (fileName: string): Promise<string | null> => {
      if (urlCache[fileName]) return urlCache[fileName];
      const formData = new FormData();
      formData.append("id", citations[fileName][0].id);
      const url = await CitationFileDownload(formData);
      if (url) {
        setUrlCache((prev) => ({ ...prev, [fileName]: url }));
      }
      return url ?? null;
    },
    [citations, urlCache]
  );

  const handleMouseEnter = useCallback(
    async (fileName: string) => {
      if (!urlCache[fileName]) {
        await resolveUrl(fileName);
      }
    },
    [urlCache, resolveUrl]
  );

  const handleClick = useCallback(
    async (e: React.MouseEvent, fileName: string) => {
      e.preventDefault();
      // キャッシュ済みなら直接 noopener で開く（ポップアップブロック回避 + 安全）
      const cached = urlCache[fileName];
      if (cached) {
        window.open(cached, "_blank", "noopener,noreferrer");
        return;
      }
      // 未キャッシュ: noopener なしで空タブを開き、取得後に URL をセット
      const newTab = window.open("", "_blank");
      const url = await resolveUrl(fileName);
      if (url && newTab) {
        newTab.location.href = url;
      } else if (newTab) {
        newTab.close();
      }
    },
    [urlCache, resolveUrl]
  );

  return (
    <div className="interactive-citation p-4 border mt-4 flex flex-col rounded-md gap-2">
      {Object.entries(citations).map(([name, items], index: number) => {
        return (
          <div key={index} className="flex flex-col gap-2">
            <div className="font-semibold text-sm">
              <a
                className="text-primary underline hover:opacity-80 cursor-pointer"
                href={urlCache[name] ?? "#"}
                target="_blank"
                rel="noopener noreferrer"
                onMouseEnter={() => handleMouseEnter(name)}
                onClick={(e) => handleClick(e, name)}
              >
                {name}
              </a>
            </div>
            <div className="flex gap-2">
              {items.map((item, index: number) => {
                return (
                  <div key={index}>
                    <CitationSlider
                      index={index + 1}
                      name={item.name}
                      id={item.id}
                    />
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
};
