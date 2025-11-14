// App Router版（route.ts の例）
// import { openai } from "@/server/openai"; 等、既存のクライアントを利用
export async function POST(req: Request) {
  const { messages, mode } = await req.json() as { messages: any[]; mode?: "normal"|"thinking" };
  const thinking = mode === "thinking";

  const sys = thinking
    ? "You may reason in up to 4 compact steps before answering concisely. Avoid repeated tool calls."
    : "Answer directly and briefly. Do not call the same tool twice in one turn.";

  const params: any = {
    model: "gpt-5",
    temperature: thinking ? 0.5 : 0.3,
    max_output_tokens: thinking ? 1800 : 900,
    parallel_tool_calls: false,
  };

  const finalMessages = [{ role: "system", content: sys }, ...messages];

  const completion = await openai.chat.completions.create({
    ...params,
    messages: finalMessages,
  });

  return new Response(JSON.stringify(completion), { status: 200, headers: { "content-type": "application/json" } });
}
