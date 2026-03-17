// src/features/chat-home-page/chat-home.tsx

import { AddExtension } from "@/features/extensions-page/add-extension/add-new-extension";
import { ExtensionCard } from "@/features/extensions-page/extension-card/extension-card";
import { ExtensionModel } from "@/features/extensions-page/extension-services/models";
import { PersonaCard } from "@/features/persona-page/persona-card/persona-card";
import { PersonaModel } from "@/features/persona-page/persona-services/models";
import { AI_DESCRIPTION, AI_NAME } from "@/features/theme/theme-config";
import { Hero } from "@/features/ui/hero";
import { ScrollArea } from "@/features/ui/scroll-area";
import Image from "next/image";
import { FC } from "react";

interface ChatPersonaProps {
  personas: PersonaModel[];
  extensions: ExtensionModel[];

  /** このユーザーが SF 連携ボタンを見てよいか */
  canUseSalesforce: boolean;
}

/** Salesforce 連携 Extension の ID（環境変数から取得） */
const SF_EXTENSION_ID = process.env.SF_EXTENSION_ID || "";

export const ChatHome: FC<ChatPersonaProps> = (props) => {
  console.log("SF canUseSalesforce =", props.canUseSalesforce);
  console.log(
    "Extension IDs =",
    (props.extensions ?? []).map((e) => e.id)
  );

  const filteredExtensions = (props.extensions ?? []).filter((extension) => {
    if (extension.id === SF_EXTENSION_ID) {
      return props.canUseSalesforce;
    }
    return true;
  });

  return (
    <ScrollArea className="flex-1">
      <main className="flex flex-1 flex-col gap-6 pb-6">
        <Hero
          title={
            <div className="flex items-center gap-3">
              <Image
                src={"/ai-icon.png"}
                width={60}
                height={60}
                quality={100}
                alt="ai-icon"
                className="h-12 w-12 sm:h-[60px] sm:w-[60px]"
              />
              <span className="text-2xl font-bold sm:text-3xl">{AI_NAME}</span>
            </div>
          }
          description={AI_DESCRIPTION}
        />

        <div className="container flex max-w-4xl flex-col gap-10 px-4 sm:px-6">
          <div>
            <h2 className="mb-3 text-xl font-bold sm:text-2xl">Extensions</h2>

            {filteredExtensions && filteredExtensions.length > 0 ? (
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
                {filteredExtensions.map((extension) => {
                  return (
                    <ExtensionCard
                      extension={extension}
                      key={extension.id}
                      showContextMenu={false}
                    />
                  );
                })}
              </div>
            ) : (
              <p className="max-w-xl text-sm text-muted-foreground sm:text-base">
                No extentions created
              </p>
            )}
          </div>

          <div>
            <h2 className="mb-3 text-xl font-bold sm:text-2xl">Personas</h2>

            {props.personas && props.personas.length > 0 ? (
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
                {props.personas.map((persona) => {
                  return (
                    <PersonaCard
                      persona={persona}
                      key={persona.id}
                      showContextMenu={false}
                    />
                  );
                })}
              </div>
            ) : (
              <p className="max-w-xl text-sm text-muted-foreground sm:text-base">
                No personas created
              </p>
            )}
          </div>
        </div>

        <AddExtension />
      </main>
    </ScrollArea>
  );
};