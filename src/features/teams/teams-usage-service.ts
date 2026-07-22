import "server-only";

import { createHash } from "crypto";
import { HistoryContainer } from "@/features/common/services/cosmos";

const TEAMS_CHANNEL = "teams" as const;
const TEAMS_THREAD_TYPE = "TEAMS_CHAT_THREAD" as const;
const TEAMS_TURN_TYPE = "TEAMS_CHAT_TURN" as const;

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

type TeamsTurnUsageRecord = {
  id: string;
  userId: string;
  createdAt: string;
  isDeleted: false;
  type: typeof TEAMS_TURN_TYPE;
  channel: typeof TEAMS_CHANNEL;
  conversationIdHash: string;
  activityIdHash: string;
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

/**
 * Records one completed user-question/bot-answer turn.
 *
 * Teams can redeliver the same activity. A deterministic ID based on the
 * conversation and activity IDs makes the write idempotent without storing
 * either identifier in plain text.
 */
export async function recordTeamsChatTurn(props: {
  conversationId: string;
  activityId: string;
  teamsUserId: string;
}): Promise<void> {
  const conversationIdHash = hashValue(props.conversationId);
  const activityIdHash = hashValue(
    `${props.conversationId}\n${props.activityId}`
  );
  const userId = hashValue(props.teamsUserId);
  const record: TeamsTurnUsageRecord = {
    id: `teams-turn-${activityIdHash}`,
    userId,
    createdAt: new Date().toISOString(),
    isDeleted: false,
    type: TEAMS_TURN_TYPE,
    channel: TEAMS_CHANNEL,
    conversationIdHash,
    activityIdHash,
  };

  try {
    await HistoryContainer().items.create(record);
  } catch (error) {
    const statusCode = Number(
      (error as { code?: unknown; statusCode?: unknown })?.statusCode ??
        (error as { code?: unknown })?.code
    );
    if (statusCode === 409) {
      // The same Teams activity was redelivered; the original turn is kept.
      return;
    }
    // Usage statistics must never prevent or replace the bot reply.
    console.warn("[teams] Failed to record completed chat turn", error);
  }
}

function hashValue(value: string): string {
  return createHash("sha256")
    .update(value.trim())
    .digest("hex");
}
