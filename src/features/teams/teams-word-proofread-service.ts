import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";

export type TeamsWordReplacement = {
  find: string;
  replace: string;
};

export async function findTeamsWordProofreadingReplacements(props: {
  documentContext: string;
  instruction: string;
}): Promise<TeamsWordReplacement[]> {
  if (!props.documentContext.trim()) return [];

  const response = await OpenAIInstance().chat.completions.create({
    model:
      process.env.AZURE_OPENAI_WORD_EDIT_DEPLOYMENT_NAME ||
      process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      {
        role: "system",
        content: `あなたは日本語文書の校正担当です。文書本文から、明確な誤字・誤記・誤変換だけを特定してください。
JSONのみを返してください: {"replaceText":[{"find":"誤った原文","replace":"正しい表記"}]}

規則:
- find は提供された本文にそのまま存在する文字列にする。
- 明確な誤字、誤記、誤変換だけを対象にする。
- 文体改善、言い換え、要約、語調変更はしない。
- 人名・会社名・製品名などの固有名詞は、本文だけから誤りと断定できない場合は変更しない。
- find と replace が同じ項目は出力しない。
- 確信できない項目は省略する。
- 該当がなければ {"replaceText":[]} を返す。`,
      },
      {
        role: "user",
        content: `依頼: ${props.instruction}\n\n対象文書:\n${props.documentContext}`,
      },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 4000,
  });

  const raw = response.choices[0]?.message?.content ?? "{}";
  try {
    const parsed = JSON.parse(raw) as {
      replaceText?: Array<{ find?: unknown; replace?: unknown }>;
    };
    const seen = new Set<string>();
    return (Array.isArray(parsed.replaceText) ? parsed.replaceText : [])
      .filter(
        (item): item is { find: string; replace: string } =>
          typeof item.find === "string" &&
          typeof item.replace === "string" &&
          item.find.trim().length > 0 &&
          item.replace.trim().length > 0 &&
          item.find !== item.replace &&
          props.documentContext.includes(item.find)
      )
      .map((item) => ({ find: item.find.trim(), replace: item.replace.trim() }))
      .filter((item) => {
        const key = `${item.find}\u0000${item.replace}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  } catch {
    console.warn("[teams-word-proofread] response JSON could not be parsed");
    return [];
  }
}

export function buildExplicitWordReplacementInstruction(
  replacements: TeamsWordReplacement[]
): string {
  return replacements
    .map(
      ({ find, replace }) =>
        `「${escapeInstructionText(find)}」を「${escapeInstructionText(
          replace
        )}」に置換`
    )
    .join("、");
}

function escapeInstructionText(value: string): string {
  return value.replace(/[「」]/g, "");
}
