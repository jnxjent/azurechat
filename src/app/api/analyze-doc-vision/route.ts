export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import {
  analyzeDocVision,
  AnalyzeDocVisionRequest,
  AnalyzeDocVisionResponse,
} from "./handler";

export async function POST(
  req: NextRequest
): Promise<NextResponse<AnalyzeDocVisionResponse>> {
  try {
    const body: AnalyzeDocVisionRequest = await req.json();
    const { fileUrl, maxPages = 30, mode } = body;
    const result = await analyzeDocVision(fileUrl, maxPages, mode);
    return NextResponse.json(result, {
      status: result.ok ? 200 : (result.error === "fileUrl is required" ? 400 : 500),
    });
  } catch (e: any) {
    console.error("[analyze-doc-vision] error:", e);
    return NextResponse.json(
      { ok: false, error: String(e?.message ?? e) },
      { status: 500 }
    );
  }
}
