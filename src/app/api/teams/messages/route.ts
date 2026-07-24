import { NextResponse } from "next/server";
import {
  getTeamsRuntime,
  isTeamsAiSearchConfigured,
  isTeamsConfigured,
  isTeamsEnabled,
  isTeamsWebSearchConfigured,
  TEAMS_MESSAGING_ENDPOINT,
} from "@/features/teams/teams-app";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const revalidate = 0;

export async function GET() {
  return NextResponse.json({
    service: "azurechat-teams",
    enabled: isTeamsEnabled(),
    configured: isTeamsConfigured(),
    aiSearchConfigured: isTeamsAiSearchConfigured(),
    webSearchConfigured: isTeamsWebSearchConfigured(),
  });
}

export async function POST(request: Request) {
  if (!isTeamsEnabled()) {
    return NextResponse.json(
      { error: "teams_disabled" },
      { status: 503 }
    );
  }

  try {
    const { adapter } = await getTeamsRuntime();
    const response = await adapter.handle(
      "POST",
      TEAMS_MESSAGING_ENDPOINT,
      request
    );

    if (response.body === undefined) {
      return new Response(null, { status: response.status });
    }

    if (typeof response.body === "string") {
      return new Response(response.body, {
        status: response.status,
        headers: { "content-type": "text/plain; charset=utf-8" },
      });
    }

    return NextResponse.json(response.body, { status: response.status });
  } catch (error) {
    console.error("[teams] Failed to process activity", error);
    return NextResponse.json(
      { error: "teams_activity_failed" },
      { status: 500 }
    );
  }
}
