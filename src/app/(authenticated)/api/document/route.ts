// app/api/document/route.ts
import { NextRequest, NextResponse } from "next/server";
import { getToken } from "next-auth/jwt";
import { decideDept, getUserEmailFromJwtToken } from "@/lib/sl-dept";
import { SearchAzureAISimilarDocuments } from "@/features/chat-page/chat-services/chat-api/chat-api-rag-extension";

export async function POST(req: NextRequest) {
  try {
    const token = await getToken({ req });
    let email = token ? getUserEmailFromJwtToken(token) : null;

    if (!email && process.env.SL_LOCAL_EMAIL) {
      email = process.env.SL_LOCAL_EMAIL;
    }

    const deptLower = decideDept({
      requestedDept: undefined,
      userEmail: email,
    });

    console.log("[DOC] email =", email);
    console.log("[DOC] deptLower =", deptLower);

    const results = await SearchAzureAISimilarDocuments(req, deptLower);

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

    return NextResponse.json(data, {
      headers: { "Content-Type": "application/json; charset=utf-8" },
    });
  } catch (err: any) {
    return NextResponse.json(
      { error: true, message: err?.message ?? "Internal error" },
      {
        status: 500,
        headers: { "Content-Type": "application/json; charset=utf-8" },
      }
    );
  }
}

function safeParse(t: string) {
  try {
    return JSON.parse(t);
  } catch {
    return { raw: t };
  }
}