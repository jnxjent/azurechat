"use client";
import { FC, useEffect, useMemo } from "react";
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

  
  const onClickFileName = async (e: any, fileName: string) => {
    e.preventDefault()
    const formData = new FormData()
    formData.append("id", citations[fileName][0].id)
    const resp = await CitationFileDownload(formData)
    if (resp) {
      window.location.href = resp
    }
  }
  return (
    <div className="interactive-citation p-4 border mt-4 flex flex-col rounded-md gap-2">
      {Object.entries(citations).map(([name, items], index: number) => {
        return (
          <div key={index} className="flex flex-col gap-2">
            <div className="font-semibold text-sm">
              <a href="" onClick={(e) => onClickFileName(e, name)}>{name}</a>
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
