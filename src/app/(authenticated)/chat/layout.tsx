import { ChatMenu } from "@/features/chat-page/chat-menu/chat-menu";
import { ChatMenuHeader } from "@/features/chat-page/chat-menu/chat-menu-header";
import { MobileChatMenu } from "@/features/chat-page/chat-menu/mobile-chat-menu";
import { FindAllChatThreadForCurrentUser } from "@/features/chat-page/chat-services/chat-thread-service";
import { MenuTray } from "@/features/main-menu/menu-tray";
import { cn } from "@/ui/lib";

import { AI_NAME } from "@/features/theme/theme-config";
import { DisplayError } from "@/features/ui/error/display-error";
import { ScrollArea } from "@/features/ui/scroll-area";

export const dynamic = "force-dynamic";

export const metadata = {
  title: AI_NAME,
  description: AI_NAME,
};

export default async function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  const chatHistoryResponse = await FindAllChatThreadForCurrentUser();

  if (chatHistoryResponse.status !== "OK") {
    return <DisplayError errors={chatHistoryResponse.errors} />;
  }

  return (
    <div className={cn("flex flex-1 items-stretch")}>
      {/* PC / タブレット: 旧レイアウトをそのまま維持 */}
      <div className="hidden flex-1 md:flex">
        <MenuTray>
          <ChatMenuHeader />
          <ScrollArea>
            <ChatMenu menuItems={chatHistoryResponse.response} />
          </ScrollArea>
        </MenuTray>
        {children}
      </div>

      {/* モバイル: children はそのまま、☰だけ浮かせる */}
      <div className="relative flex flex-1 md:hidden">
        {children}

        <div className="pointer-events-none absolute left-3 top-20 z-[400]">
        <div className="pointer-events-auto">
          <MobileChatMenu menuItems={chatHistoryResponse.response} />
        </div>
      </div>
      </div>
    </div>
  );
}