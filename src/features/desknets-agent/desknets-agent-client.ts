import "server-only";

import { userHashedId, userSession } from "@/features/auth-page/helpers";
import type {
  DeskNetsAgentRunRequest,
  DeskNetsAgentRunResponse,
} from "./desknets-agent-types";

const POLL_INTERVAL_MS = 500;
// Browser Agent allows runs up to 300 seconds. Keep a small transport margin.
const MAX_POLL_DURATION_MS = 330_000;
const TERMINAL_STATUSES = new Set([
  "completed",
  "failed",
  "cancelled",
  "awaiting_user_input",
  "awaiting_approval",
]);

function responseMessage(
  body: DeskNetsAgentRunResponse,
  fallback: string
): string {
  const message = body.message?.trim();
  if (message) return message;
  const error = body.error?.trim();
  return error || fallback;
}

function getAgentBaseUrl(): string {
  return (process.env.DESKNETS_AGENT_API_URL || "").replace(/\/+$/, "");
}

export function isDeskNetsAgentEnabled(): boolean {
  const enabled = Boolean(getAgentBaseUrl()) && process.env.DESKNETS_AGENT_ENABLED !== "false";
  if (process.env.NODE_ENV !== "production") {
    console.log("[DeskNetsAgent] enabled check", {
      enabled,
      flag: process.env.DESKNETS_AGENT_ENABLED,
      hasApiUrl: Boolean(getAgentBaseUrl()),
    });
  }
  return enabled;
}

export async function runDeskNetsAgent(
  prompt: string,
  chatThreadId: string
): Promise<DeskNetsAgentRunResponse> {
  const baseUrl = getAgentBaseUrl();
  if (!baseUrl) {
    return {
      status: "failed",
      message: "DeskNet's Agent is not configured.",
    };
  }

  const [hashedUserId, currentUser] = await Promise.all([
    userHashedId(),
    userSession().catch(() => null),
  ]);

  const payload: DeskNetsAgentRunRequest = {
    userId: hashedUserId,
    userEmail: currentUser?.email || undefined,
    threadId: chatThreadId,
    site: "desknets",
    mode: "read",
    prompt,
  };

  const headers: Record<string, string> = {
    "content-type": "application/json",
  };
  const apiKey = process.env.DESKNETS_AGENT_API_KEY?.trim();
  if (apiKey) headers.authorization = `Bearer ${apiKey}`;
  if (currentUser?.email) headers["x-user-email"] = currentUser.email;
  headers["x-user-id"] = hashedUserId;
  headers["x-chat-thread-id"] = chatThreadId;

  try {
    console.log("[DeskNetsAgent] POST /browser-agent/runs", {
      chatThreadId,
      baseUrl,
    });
    const response = await fetch(`${baseUrl}/browser-agent/runs`, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
      cache: "no-store",
    });

    const raw = await response.text();
    let body: DeskNetsAgentRunResponse;
    try {
      body = JSON.parse(raw) as DeskNetsAgentRunResponse;
    } catch {
      body = { status: "failed", message: raw.slice(0, 1000) };
    }

    if (!response.ok) {
      return {
        ...body,
        status: "failed",
        message: responseMessage(
          body,
          `DeskNet's Agent returned HTTP ${response.status}.`
        ),
      };
    }

    const runId = body.id || body.runId;
    if (!runId || TERMINAL_STATUSES.has(body.status)) {
      return body;
    }

    const pollDeadline = Date.now() + MAX_POLL_DURATION_MS;
    while (Date.now() < pollDeadline) {
      await new Promise((resolve) => setTimeout(resolve, POLL_INTERVAL_MS));
      const statusResponse = await fetch(
        `${baseUrl}/browser-agent/runs/${encodeURIComponent(runId)}`,
        { method: "GET", headers, cache: "no-store" }
      );
      const statusRaw = await statusResponse.text();
      let statusBody: DeskNetsAgentRunResponse;
      try {
        statusBody = JSON.parse(statusRaw) as DeskNetsAgentRunResponse;
      } catch {
        statusBody = { id: runId, status: "failed", message: statusRaw.slice(0, 1000) };
      }
      if (!statusResponse.ok) {
        return {
          ...statusBody,
          id: runId,
          status: "failed",
          message: responseMessage(
            statusBody,
            `DeskNet's Agent status returned HTTP ${statusResponse.status}.`
          ),
        };
      }
      if (TERMINAL_STATUSES.has(statusBody.status)) {
        return statusBody;
      }
      if (!["queued", "running"].includes(statusBody.status)) {
        return {
          ...statusBody,
          id: runId,
          status: "failed",
          message: `DeskNet's Agent returned unknown status: ${statusBody.status}`,
        };
      }
    }

    return {
      id: runId,
      status: "failed",
      message: "DeskNet's Agent timed out while waiting for the browser run.",
    };
  } catch (error) {
    console.error("[DeskNetsAgent] request failed:", error);
    return {
      status: "failed",
      message: `DeskNet's Agent request failed: ${String(error)}`,
    };
  }
}

