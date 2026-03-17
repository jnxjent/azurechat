// src/app/(authenticated)/chat/page.tsx

import { ChatHome } from "@/features/chat-home-page/chat-home";
import { FindAllExtensionForCurrentUser } from "@/features/extensions-page/extension-services/extension-service";
import { FindAllPersonaForCurrentUser } from "@/features/persona-page/persona-services/persona-service";
import { DisplayError } from "@/features/ui/error/display-error";

// 実際のログインユーザー取得
import { getCurrentUser } from "@/features/auth-page/helpers";

export default async function Home() {
  const [personaResponse, extensionResponse] = await Promise.all([
    FindAllPersonaForCurrentUser(),
    FindAllExtensionForCurrentUser(),
  ]);

  if (personaResponse.status !== "OK") {
    return <DisplayError errors={personaResponse.errors} />;
  }

  if (extensionResponse.status !== "OK") {
    return <DisplayError errors={extensionResponse.errors} />;
  }

  // 実際のログインユーザー（Local/Remote 共通）
  const user = await getCurrentUser();
  const email = (user?.email || "").toLowerCase().trim();

  // ホワイトリスト（カンマ or 改行 区切り対応）
  const raw = (process.env.SF_WHITELIST_EMAILS || "").trim();
  const allowList = raw
    .split(/[,\n]/)
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);

  const canUseSalesforce = !!email && allowList.includes(email);

  return (
    <ChatHome
      personas={personaResponse.response}
      extensions={extensionResponse.response}
      canUseSalesforce={canUseSalesforce}
    />
  );
}
