// File: src/features/image-text/overlay-llm.ts
// 日本語のレイアウト指示 → OverlayLayout(JSON) を返す LLM モジュール

const AZ_ENDPOINT = process.env.AZURE_OPENAI_ENDPOINT!;
const AZ_KEY = process.env.AZURE_OPENAI_API_KEY!;
const API_VERSION =
  process.env.AZURE_OPENAI_API_VERSION || "2024-12-01-preview";

// 文字レイアウト用 LLM デプロイ名（例: gpt-5.1-mini 等）
// 例: AZURE_OPENAI_LLM_DEPLOYMENT=gpt-5.1-mini
const LLM_DEPLOYMENT = process.env.AZURE_OPENAI_LLM_DEPLOYMENT;

export type FontFamilyKey = "gothic" | "mincho" | "meiryo";

export type OverlayLayout = {
  text?: string;
  fontFamily?: FontFamilyKey;
  bold?: boolean;
  italic?: boolean;
  fontSizePx?: number;
  x?: number; // 0〜1（左→右）
  y?: number; // 0〜1（上→下）
  fillColor?: string;
  strokeColor?: string;
  align?: "left" | "center" | "right";
  vAlign?: "top" | "middle" | "bottom";
};

export async function decideOverlayLayout(params: {
  instruction: string;
  baseText: string;
  width: number;
  height: number;
}): Promise<OverlayLayout> {
  if (!LLM_DEPLOYMENT) {
    throw new Error("AZURE_OPENAI_LLM_DEPLOYMENT is not set.");
  }

  const url = `${AZ_ENDPOINT.replace(
    /\/+$/,
    ""
  )}/openai/deployments/${LLM_DEPLOYMENT}/chat/completions?api-version=${API_VERSION}`;

  const systemPrompt = `
あなたは画像の文字レイアウトエンジンです。
ユーザーの日本語指示を読み取り、次の TypeScript 型 OverlayLayout に従う JSON オブジェクト1個だけを返してください。

type OverlayLayout = {
  text?: string;                   // 実際に描く文字列（未指定なら baseText を使う）
  fontFamily?: "gothic" | "mincho" | "meiryo";
  bold?: boolean;
  italic?: boolean;
  fontSizePx?: number;             // 例: 32〜96
  x?: number;                      // 0〜1 の相対座標（左端0, 右端1）
  y?: number;                      // 0〜1 の相対座標（上端0, 下端1）
  fillColor?: string;              // "#ffffff" など
  strokeColor?: string;            // "#000000" など
  align?: "left" | "center" | "right";
  vAlign?: "top" | "middle" | "bottom";
};

制約:
- 必ず JSON オブジェクトのみを返し、説明文・コメント・コードブロック記法( \`\`\` )は一切付けないこと。
- x, y を指定する場合は 0〜1 の範囲にすること（例: 0.5 は中央）。
- 座標が「下に」「中央」などの指示の場合、y は 0.7〜0.9 など、指示にふさわしい相対値にすること。
- 明朝は fontFamily "mincho"、ゴシックは "gothic"、メイリオ系は "meiryo" として指定してください。
- 太字のときは bold: true を返してください。
`;

  const userPrompt = `
【テキスト】
${params.baseText}

【レイアウト指示】
${params.instruction}

【画像サイズ】
width=${params.width}, height=${params.height}
`;

  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "api-key": AZ_KEY,
    },
    body: JSON.stringify({
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt },
      ],
      temperature: 0.2,
      response_format: { type: "json_object" },
    }),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`layout LLM error ${res.status}: ${text}`);
  }

  const json = await res.json();
  const content = json?.choices?.[0]?.message?.content;
  if (!content || typeof content !== "string") {
    throw new Error("layout LLM: empty content");
  }

  try {
    const parsed = JSON.parse(content);
    return parsed as OverlayLayout;
  } catch (e) {
    console.error("layout JSON parse failed:", e, content);
    throw new Error("layout LLM returned invalid JSON");
  }
}
