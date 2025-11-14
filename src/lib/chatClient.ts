// 既存の送信関数に追記
export type ChatMode = "normal" | "thinking";

function hasAttachment() {
  // 既存の添付管理に合わせて判定（暫定は false でOK）
  return false;
}

function autoDecide(prompt: string, currentUI: ChatMode, manual: boolean): ChatMode {
  if (manual) return currentUI;
  const complex =
    prompt.length > 300 ||
    /なぜ|分析|比較|傾向|要因|理由|まとめ|要約/.test(prompt) ||
    hasAttachment();
  return complex ? "thinking" : "normal";
}

export async function sendChat(messages: any[], latestUserText: string) {
  const uiMode =
    (typeof window !== "undefined" &&
      (localStorage.getItem("chat_mode") as ChatMode)) ||
    "normal";
  const manual =
    typeof window !== "undefined" &&
    localStorage.getItem("chat_manual") === "1";
  const decided = autoDecide(latestUserText ?? "", uiMode, manual);

  const res = await fetch("/api/chat", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ messages, mode: decided, manualOverride: manual }),
    // ▼▼ キャッシュ抑止（2MBフェッチキャッシュ問題の止血）
    cache: "no-store",
    next: { revalidate: 0 },
  });

  return res.json();
}
