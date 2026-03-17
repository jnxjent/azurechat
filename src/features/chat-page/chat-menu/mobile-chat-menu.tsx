"use client";

import Link from "next/link";
import { CreateChatAndRedirect } from "../chat-services/chat-thread-service";
import { ChevronLeft, Menu, X } from "lucide-react";
import React, { useState, useTransition } from "react";
import { NewChat } from "./new-chat";
import { ChatMenu } from "./chat-menu";
import type { ChatThreadModel } from "../chat-services/models";

interface MobileChatMenuProps {
  menuItems: ChatThreadModel[];
}

type MobileMenuView = "root" | "history";

export const MobileChatMenu: React.FC<MobileChatMenuProps> = ({
  menuItems,
}) => {
  const [open, setOpen] = useState(false);
  const [view, setView] = useState<MobileMenuView>("root");
  const [, startTransition] = useTransition();

  const closeMenu = () => {
    setOpen(false);
    setView("root");
  };

  const openMenu = () => {
    setOpen(true);
    setView("root");
  };

  return (
    <>
      <div className="relative z-[120] flex items-center md:hidden">
        <button
          type="button"
          aria-label="Open menu"
          onClick={openMenu}
          className="pointer-events-auto inline-flex h-10 w-10 items-center justify-center rounded-md border border-border bg-background text-foreground shadow-sm"
        >
          <Menu size={18} />
        </button>
      </div>

      {open && (
        <div className="fixed inset-0 z-[200] md:hidden">
          <div
            className="absolute inset-0 bg-black/50"
            onClick={closeMenu}
          />

          <div className="absolute left-0 top-0 flex h-full w-[85%] max-w-[320px] flex-col border-r border-border bg-background shadow-xl">
            <div className="flex items-center justify-between border-b px-4 py-3">
              <div className="text-sm font-semibold">
                {view === "root" ? "Menu" : "過去スレッド"}
              </div>

              <div className="flex items-center gap-2">
                {view !== "root" && (
                  <button
                    type="button"
                    aria-label="Back"
                    onClick={() => setView("root")}
                    className="inline-flex h-9 w-9 items-center justify-center rounded-md border border-border bg-background"
                  >
                    <ChevronLeft size={18} />
                  </button>
                )}

                <button
                  type="button"
                  aria-label="Close menu"
                  onClick={closeMenu}
                  className="inline-flex h-9 w-9 items-center justify-center rounded-md border border-border bg-background"
                >
                  <X size={18} />
                </button>
              </div>
            </div>

            <div className="flex-1 overflow-y-auto">
              {view === "root" ? (
                <div className="p-4">
                  <div className="space-y-3">
                    <div>
                      <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                        New Chat
                      </div>
                      <form
                        action={() => {
                          closeMenu();
                          startTransition(() => {
                            CreateChatAndRedirect();
                          });
                        }}
                        className="w-full"
                      >
                        <NewChat />
                      </form>
                    </div>

                    <div>
                      <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                        Extensions
                      </div>
                      <Link
                        href="/chat"
                        onClick={closeMenu}
                        className="block rounded-md border border-border px-3 py-3 text-sm hover:bg-muted"
                      >
                        Extensions 一覧へ
                      </Link>
                    </div>

                    <div>
                      <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                        過去スレッド
                      </div>
                      <button
                        type="button"
                        onClick={() => setView("history")}
                        className="block w-full rounded-md border border-border px-3 py-3 text-left text-sm hover:bg-muted"
                      >
                        過去スレッドを開く
                      </button>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="p-4">
                  <div
                    onClick={() => {
                      closeMenu();
                    }}
                  >
                    <ChatMenu menuItems={menuItems} />
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default MobileChatMenu;