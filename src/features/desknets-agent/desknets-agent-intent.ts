/** Lightweight routing for the built-in DeskNet's Agent tool. */
export function isDeskNetsAgentRequest(message: string): boolean {
  const text = (message || "").trim();
  if (!text) return false;

  const matched = [
    /desknet'?s/i,
    /\u30c7\u30b9\u30af\u30cd\u30c3\u30c4/i,
    /(?:schedule|meeting|appointment|availability|calendar).*(?:find|check|book|add|available|candidate)/i,
    /(?:find|check|book|add|available|candidate).*(?:schedule|meeting|appointment|availability|calendar)/i,
    /\u7a7a\u304d\u6642\u9593|\u7a7a\u304d\u72b6\u6cc1|\u65e5\u7a0b\u8abf\u6574|\u4e88\u5b9a\u8ffd\u52a0/i,
    /(?:\u6253\u3061\u5408\u308f\u305b|\u4f1a\u8b70).*(?:\u53ef\u80fd|\u5019\u88dc|\u7a7a\u304d|\u4e88\u5b9a)/i,
    /(?:\u53ef\u80fd|\u5019\u88dc|\u7a7a\u304d).*(?:\u65e5\u6642|\u6642\u9593|\u65e5\u7a0b)/i,
    /\u53c2\u52a0\u8005.*(?:\u4e88\u5b9a|\u7a7a\u304d|\u65e5\u7a0b)/i,
  ].some((pattern) => pattern.test(text));

  if (process.env.NODE_ENV !== "production") {
    console.log("[DeskNetsAgent] intent check", {
      matched,
      messagePreview: text.slice(0, 80),
    });
  }
  return matched;
}

export function isDeskNetsAgentFollowUpRequest(message: string): boolean {
  const text = (message || "").normalize("NFKC").trim();
  if (!text) return false;

  return [
    /\u5019\u88dc\s*\d+|\d+\s*\u756a/,
    /(?:\d+\s*\u5206|\d+\s*\u6642\u9593).*(?:\u5909\u66f4|\u306b\u3057\u3066|\u3067)/,
    /(?:\u4f1a\u8b70\u5ba4|\u30eb\u30fc\u30e0).*(?:\u7a7a\u304d|\u5019\u88dc|\u5909\u66f4|\u306b\u3057\u3066)/,
    // \u5019\u88dc ("\u5019\u88dc") + \u623b/\u3084\u3081/\u898b\u305b/\u8868\u793a/\u4e00\u89a7 ("\u5019\u88dc\u306b\u623b\u3057\u3066" etc.) \u2014 the user changed their mind about
    // a booking in progress and wants the previous candidate list again.
    /\u5019\u88dc.*(?:\u623b|\u3084\u3081|\u898b\u305b|\u8868\u793a|\u4e00\u89a7)|(?:\u623b|\u3084\u3081).*\u5019\u88dc/,
    /\u30e1\u30fc\u30eb.*(?:\u9001\u4fe1|\u4e0d\u8981|\u306a\u3057|\u3042\u308a)/,
    /^(?:\u306f\u3044|\u3044\u3044\u3048|\u304a\u9858\u3044\u3057\u307e\u3059|\u9001\u4fe1\u3057\u307e\u3059|\u9001\u4fe1\u3057\u307e\u305b\u3093)[\u3002\uff01!]?$/,
    /(?:\u305d\u308c|\u305d\u306e\u5019\u88dc|\u5148\u307b\u3069).*(?:\u5909\u66f4|\u4e88\u7d04|\u8ffd\u52a0|\u9032\u3081)/,
  ].some((pattern) => pattern.test(text));
}

// Scans the last few messages backward for `marker`, but stops (returns
// false) the moment it hits a user-role message first — that means a user
// reply has already come and gone since the question was asked, so the
// question is stale/answered and must not keep triggering routing on every
// later message. Without this, a short-lived question (e.g. "会議室の場所は
// どこにしますか") lingering in the last-4 window would misroute an
// unrelated later message like "売上を教えて" straight to DeskNet's.
function isAwaitingReplyTo(messages: unknown[], marker: string): boolean {
  if (!Array.isArray(messages)) return false;

  const recent = messages.slice(-4);
  for (let index = recent.length - 1; index >= 0; index -= 1) {
    const message: any = recent[index];
    const role = typeof message?.role === "string" ? message.role : "";
    if (role === "user") return false;
    const content = typeof message?.content === "string" ? message.content : "";
    if (content.includes(marker)) return true;
  }
  return false;
}

// The literal question text the Agent API asks when it can't resolve a
// meeting room automatically (see azure-browser-agent's
// server.ts:formatFacilityChoiceMessage). A reply like "有玉" or "アクト" on
// its own matches none of isDeskNetsAgentFollowUpRequest's patterns, so
// without this check the answer would be routed to normal chat instead of
// back to the Agent API and the booking flow would stall.
const FACILITY_CHOICE_QUESTION_MARKER = "会議室の場所はどこにしますか";

export function isAwaitingFacilityChoiceReply(messages: unknown[]): boolean {
  return isAwaitingReplyTo(messages, FACILITY_CHOICE_QUESTION_MARKER);
}

// The literal (name-independent) fragment of the question the Agent API
// asks when a requested participant's name matches people in more than one
// organization (see azure-browser-agent's desknets-worker.ts, the
// AmbiguousParticipantError handling in executeAvailabilityRun). A reply
// like "経営企画部" or "三晃" on its own matches none of
// isDeskNetsAgentFollowUpRequest's patterns either, so without this check
// the answer would stall the same way the facility-choice reply did.
const PARTICIPANT_CHOICE_QUESTION_MARKER = "さんが複数見つかりました";

export function isAwaitingParticipantChoiceReply(messages: unknown[]): boolean {
  return isAwaitingReplyTo(messages, PARTICIPANT_CHOICE_QUESTION_MARKER);
}

export function hasDeskNetsAgentContext(messages: unknown[]): boolean {
  if (!Array.isArray(messages)) return false;

  return messages.slice(-8).some((message: any) => {
    const content = typeof message?.content === "string" ? message.content : "";
    const name = typeof message?.name === "string" ? message.name : "";
    return /desknets_schedule_agent|desknet'?s|\u30c7\u30b9\u30af\u30cd\u30c3\u30c4|\u4e88\u5b9a\u8ffd\u52a0\u753b\u9762/i.test(
      `${name}\n${content}`
    );
  });
}

export function shouldRouteToDeskNetsAgent(
  message: string,
  history: unknown[]
): boolean {
  return (
    isDeskNetsAgentRequest(message) ||
    isAwaitingFacilityChoiceReply(history) ||
    isAwaitingParticipantChoiceReply(history) ||
    (isDeskNetsAgentFollowUpRequest(message) &&
      hasDeskNetsAgentContext(history))
  );
}

