// src/app/(authenticated)/chat/[id]/page.tsx
export const dynamic = "force-dynamic";
export const revalidate = 0;

import "@/lib/no-store-fetch";
import { getServerSession } from "next-auth";
import { options as authOptions } from "@/features/auth-page/auth-api";
import { ChatPage } from "@/features/chat-page/chat-page";
import { FindAllChatDocuments } from "@/features/chat-page/chat-services/chat-document-service";
import { FindAllChatMessagesForCurrentUser } from "@/features/chat-page/chat-services/chat-message-service";
import { FindChatThreadForCurrentUser } from "@/features/chat-page/chat-services/chat-thread-service";
import { FindAllExtensionForCurrentUser } from "@/features/extensions-page/extension-services/extension-service";
import { AI_NAME } from "@/features/theme/theme-config";
import { DisplayError } from "@/features/ui/error/display-error";
import { resolveSlAccess } from "@/lib/sl-dept";

export const metadata = {
  title: AI_NAME,
  description: AI_NAME,
};

interface HomeParams {
  params: {
    id: string;
  };
}

export default async function Home(props: HomeParams) {
  const { id } = props.params;

  const [chatResponse, chatThreadResponse, docsResponse, extensionResponse, session] =
    await Promise.all([
      FindAllChatMessagesForCurrentUser(id),
      FindChatThreadForCurrentUser(id),
      FindAllChatDocuments(id),
      FindAllExtensionForCurrentUser(),
      getServerSession(authOptions),
    ]);

  if (docsResponse.status !== "OK") {
    return <DisplayError errors={docsResponse.errors} />;
  }
  if (chatResponse.status !== "OK") {
    return <DisplayError errors={chatResponse.errors} />;
  }
  if (extensionResponse.status !== "OK") {
    return <DisplayError errors={extensionResponse.errors} />;
  }
  if (chatThreadResponse.status !== "OK") {
    return <DisplayError errors={chatThreadResponse.errors} />;
  }

  const userEmail = session?.user?.email;
  const sessionRole = session?.user?.slRole;
  const sessionDept = session?.user?.slDept;
  const resolvedAccess =
    sessionRole && sessionDept
      ? { role: sessionRole, dept: sessionDept }
      : resolveSlAccess(userEmail);

  const isAdmin =
    resolvedAccess.role === "global_admin" ||
    resolvedAccess.role === "dept_admin";

  console.log(
    `[ROLE] email=${userEmail} dept=${resolvedAccess.dept} role=${resolvedAccess.role}`
  );

  return (
    <ChatPage
      messages={chatResponse.response}
      chatThread={chatThreadResponse.response}
      chatDocuments={docsResponse.response}
      extensions={extensionResponse.response}
      isAdmin={isAdmin}
    />
  );
}
