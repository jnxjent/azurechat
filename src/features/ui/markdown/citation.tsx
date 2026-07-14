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

const isBlobCitationId = (id: string): boolean =>
  /^https?:\/\//i.test(id) ||
  /blob\.core\.windows\.net/i.test(id) ||
  /[?&]sig=/i.test(id);

export const Citation: FC<Props> = (props: Props) => {
  // Filter out invalid citations (blob URLs, SAS URLs, empty ids)
  const validItems = props.items.filter(
    (c) => !!c.id && !isBlobCitationId(c.id) && !isBlobCitationId(c.name)
  );

  const citations = validItems.reduce(
    (acc, cit) => {
      const { name } = cit;
      if (!acc[name]) acc[name] = [];
      acc[name].push(cit);
      return acc;
    },
    {} as Record<string, Citation[]>
  );

  // All hooks must be declared before early return
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
      const cached = urlCache[fileName];
      if (cached) {
        window.open(cached, "_blank", "noopener,noreferrer");
        return;
      }
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

  if (validItems.length === 0) return null;

  return (
    <div className="interactive-citation p-4 border mt-4 flex flex-col rounded-md gap-2">
      {Object.entries(citations).map(([name, items], index: number) => (
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
            {items.map((item, idx: number) => (
              <div key={idx}>
                <CitationSlider index={idx + 1} name={item.name} id={item.id} />
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
};
