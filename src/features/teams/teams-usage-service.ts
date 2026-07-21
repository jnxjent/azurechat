import "server-only";

import { createHash } from "crypto";
import { HistoryContainer } from "@/features/common/services/cosmos";

const TEAMS_CHANNEL = "teams" as const;
const TEAMS_THREAD_TYPE = "TEAMS_CHAT_THREAD" as const;

type TeamsThreadUsageRecord = {
  id: string;
  userId: string;
  createdAt: string;
  lastMessageAt: string;
  isDeleted: false;
  type: typeof TEAMS_THREAD_TYPE;
  channel: typeof TEAMS_CHANNEL;
  conversationIdHash: string;
};

/**
 * Records one Cosmos DB item per Teams conversation.
 *
 * The deterministic item ID keeps repeated messages in the same conversation
 * from increasing the thread count. User and conversation identifiers are
 * hashed before storage.
 */
export async function recordTeamsThreadUsage(props: {
  conversationId: string;
  teamsUserId: string;
}): Promise<void> {
  const conversationIdHash = hashValue(props.conversationId);
  const userId = hashValue(props.teamsUserId);
  const id = `teams-thread-${conversationIdHash}`;
  const container = HistoryContainer();
  const now = new Date().toISOString();

  try {
    const { resources } = await container.items
      .query<TeamsThreadUsageRecord>(
        {
          query: "SELECT * FROM c WHERE c.id = @id",
          parameters: [{ name: "@id", value: id }],
        },
        { partitionKey: userId }
      )
      .fetchAll();
    const existing = resources[0];

    const record: TeamsThreadUsageRecord = {
      id,
      userId,
      createdAt: existing?.createdAt ?? now,
      lastMessageAt: now,
      isDeleted: false,
      type: TEAMS_THREAD_TYPE,
      channel: TEAMS_CHANNEL,
      conversationIdHash,
    };

    await container.items.upsert(record);
  } catch (error) {
    // Usage statistics must never prevent the bot from replying.
    console.warn("[teams] Failed to record usage statistics", error);
  }
}

function hashValue(value: string): string {
  return createHash("sha256")
    .update(value.trim())
    .digest("hex");
}
