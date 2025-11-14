// app/api/document/route.ts
import { NextResponse } from "next/server";
import { SearchAzureAISimilarDocuments } from "@/features/chat-page/chat-services/chat-api/chat-api-rag-extension";

export async function POST(req: Request) {
  try {
    // 実処理（RAG 検索）
    const results = await SearchAzureAISimilarDocuments(req);

    // 返り値が string（JSON文字列）でも Response でも 素のオブジェクトでも吸収
    let data: any;
    if (results instanceof Response) {
      const ct = results.headers.get("content-type") || "";
      const text = await results.text();
      data = ct.includes("json") ? JSON.parse(text) : safeParse(text);
    } else if (typeof results === "string") {
      data = safeParse(results);
    } else {
      data = results;
    }

    // ✅ 必ず application/json で返す
    return NextResponse.json(data, {
      headers: { "Content-Type": "application/json; charset=utf-8" },
    });
  } catch (err: any) {
    return NextResponse.json(
      { error: true, message: err?.message ?? "Internal error" },
      { status: 500, headers: { "Content-Type": "application/json; charset=utf-8" } }
    );
  }
}

function safeParse(t: string) {
  try { return JSON.parse(t); } catch { return { raw: t }; }
}
