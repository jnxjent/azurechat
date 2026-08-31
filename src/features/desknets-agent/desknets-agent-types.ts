export type DeskNetsAgentMode = "read" | "write";

export type DeskNetsAgentRunStatus =
  | "queued"
  | "running"
  | "awaiting_user_input"
  | "awaiting_approval"
  | "completed"
  | "failed"
  | "cancelled";

export type DeskNetsAgentRunRequest = {
  userId: string;
  userEmail?: string;
  threadId: string;
  site: "desknets";
  mode: DeskNetsAgentMode;
  prompt: string;
};

export type DeskNetsAgentRunResponse = {
  id?: string;
  runId?: string;
  status: DeskNetsAgentRunStatus | string;
  message?: string;
  error?: string;
  currentUrl?: string;
  screenshotUrl?: string;
  sessionId?: string;
  candidates?: unknown[];
  result?: {
    assistantMessage?: string;
    summary?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

