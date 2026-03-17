import { FC } from "react";
import { ScrollArea } from "../ui/scroll-area";
import { AddExtension } from "./add-extension/add-new-extension";
import { ExtensionCard } from "./extension-card/extension-card";
import { ExtensionHero } from "./extension-hero/extension-hero";
import { ExtensionModel } from "./extension-services/models";

interface Props {
  extensions: ExtensionModel[];
  /** WhiteList 判定（ChatHome と同じ） */
  canUseSalesforce: boolean;
}

/** Salesforce 連携 Extension の ID（環境変数） */
const SF_EXTENSION_ID = process.env.SF_EXTENSION_ID || "";

export const ExtensionPage: FC<Props> = (props) => {
  const filteredExtensions = (props.extensions ?? []).filter((extension) => {
    if (extension.id === SF_EXTENSION_ID) {
      return props.canUseSalesforce; // ★ WhiteList 制御
    }
    return true;
  });

  return (
    <ScrollArea className="flex-1">
      <main className="flex flex-1 flex-col">
        <ExtensionHero />
        <div className="container max-w-4xl py-3">
          <div className="grid grid-cols-3 gap-3">
            {filteredExtensions.map((extension) => {
              return (
                <ExtensionCard
                  extension={extension}
                  key={extension.id}
                  showContextMenu
                />
              );
            })}
          </div>
        </div>
        <AddExtension />
      </main>
    </ScrollArea>
  );
};
