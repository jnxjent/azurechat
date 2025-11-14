// src/app/(authenticated)/chat/page.tsx

import { ChatHome } from "@/features/chat-home-page/chat-home";
import { FindAllExtensionForCurrentUser } from "@/features/extensions-page/extension-services/extension-service";
import { FindAllPersonaForCurrentUser } from "@/features/persona-page/persona-services/persona-service";
import { DisplayError } from "@/features/ui/error/display-error";

export default async function Home() {
  // Persona + Extension をまとめて取得（元の動き）
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

  // -------------------------------
  // ★ ローカル専用：擬似「現在ユーザーのメール」
  //   - .env.local に NEXT_PUBLIC_DEV_USER_EMAIL を定義してテスト
  //   - 例: NEXT_PUBLIC_DEV_USER_EMAIL=xxxx@xxxx.co.jp
  // -------------------------------
  const emailRaw = process.env.NEXT_PUBLIC_DEV_USER_EMAIL || "";
  const email = String(emailRaw || "").toLowerCase();

  // ★ .env.local からホワイトリストを読み込む（カンマ区切り）
  //   例: SF_WHITELIST_EMAILS=foo@bar.com,xxxx@xxxx.co.jp
  const rawWhitelist = process.env.SF_WHITELIST_EMAILS || "";
  const whitelist = rawWhitelist
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);

  // ★ この「擬似ユーザー」が SF 連携を使ってよいか判定
  const canUseSalesforce = email !== "" && whitelist.includes(email);

  return (
    <ChatHome
      personas={personaResponse.response}
      extensions={extensionResponse.response}
      canUseSalesforce={canUseSalesforce}
    />
  );
}
