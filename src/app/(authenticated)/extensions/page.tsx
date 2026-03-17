// src/app/(authenticated)/extensions/page.tsx
import { ExtensionPage } from "@/features/extensions-page/extension-page";
import { FindAllExtensionForCurrentUser } from "@/features/extensions-page/extension-services/extension-service";
import { DisplayError } from "@/features/ui/error/display-error";

// ★追加: ログインユーザー取得（既存で使っているのと同じヘルパー）
import { getCurrentUser } from "@/features/auth-page/helpers";

export default async function Home() {
  const extensionResponse = await FindAllExtensionForCurrentUser();

  if (extensionResponse.status !== "OK") {
    return <DisplayError errors={extensionResponse.errors} />;
  }

  // ★追加: whitelist 判定（空なら誰も許可しない）
  const user = await getCurrentUser().catch(() => null as any);
  const email = ((user as any)?.email || "").toLowerCase().trim();

  const raw = (process.env.SF_WHITELIST_EMAILS || "").trim();
  const allowList = raw
    .split(/[,\n]/)
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);

  const canUseSalesforce =
    !!email && allowList.length > 0 && allowList.includes(email);

  return (
    <ExtensionPage
      extensions={extensionResponse.response}
      canUseSalesforce={canUseSalesforce}
    />
  );
}
