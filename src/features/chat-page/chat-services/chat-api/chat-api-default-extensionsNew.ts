// ---------------- 既存画像への文字追加（EDIT 用・Vision 不使用） ----------------
async function executeAddTextToExistingImage(
  args: {
    imageUrl: string;
    text: string;
    styleHint?: string;
    font?: string;
    color?: string;
    size?: string;
    offsetX?: number;
    offsetY?: number;
  },
  chatThread: ChatThreadModel,
  userMessage: string,
  signal: AbortSignal,
  modeOpts?: {
    reasoning_effort?: "low" | "medium" | "high";
    temperature?: number;
  }
) {
  // LLM から渡された URL はログ用に保持（実際のベースには使わない）
  const explicitUrl = (args?.imageUrl || "").trim();
  const text = (args?.text || "").trim();
  const styleHint = (args?.styleHint || "").trim();

  // ★ ベース画像は常に「threadId/__base__.png」
  const baseImageUrl = buildExternalImageUrl(chatThread.id, "__base__.png");

  console.log("🖋 add_text_to_existing_image (simple) called:", {
    passedImageUrl: explicitUrl,
    usedBaseImageUrl: baseImageUrl,
    text,
    styleHint,
    offsetX: args?.offsetX,
    offsetY: args?.offsetY,
  });

  if (!text) {
    return {
      error: "text is required for add_text_to_existing_image.",
    };
  }

  // ★ styleHint + userMessage からスタイルを推定
  const hintSource = styleHint || userMessage || "";
  const parsed = parseStyleHint(hintSource);

  // ---- 位置・サイズ・色 ----
  const align: "left" | "center" | "right" =
    (parsed.align as any) ?? "center";
  const vAlign: "top" | "middle" | "bottom" =
    (parsed.vAlign as any) ?? "bottom";
  const size: "small" | "medium" | "large" | "xlarge" =
    (args.size as any) ?? parsed.size ?? "large";
  const color = args.color ?? parsed.color ?? "white";

  // ---- フォント種別（ゴシック / 明朝 / メイリオ） ----
  const fontHint = (
    (styleHint || "") +
    " " +
    (args.font || "") +
    " " +
    (parsed.font || "")
  ).toLowerCase();

  let fontFamily: "gothic" | "mincho" | "meiryo" = "gothic";

  if (
    fontHint.includes("明朝") ||
    fontHint.includes("mincho") ||
    fontHint.includes("serif")
  ) {
    fontFamily = "mincho";
  } else if (fontHint.includes("メイリオ") || fontHint.includes("meiryo")) {
    fontFamily = "meiryo";
  } else {
    // 特に指定がなければ「ゴシック系」
    fontFamily = "gothic";
  }

  // ---- 太字 / イタリック ----
  const lowerHint = hintSource.toLowerCase();
  const bold =
    hintSource.includes("太字") ||
    hintSource.includes("ボールド") ||
    lowerHint.includes("bold");
  const italic =
    hintSource.includes("イタリック") ||
    hintSource.includes("斜体") ||
    lowerHint.includes("italic");

  // ★ 累積移動：args の offset をベースに、styleHint 由来の増分を足す
  const baseOffsetX =
    typeof args.offsetX === "number" ? args.offsetX : 0;
  const baseOffsetY =
    typeof args.offsetY === "number" ? args.offsetY : 0;

  const offsetX = baseOffsetX + (parsed.offsetX ?? 0);
  const offsetY = baseOffsetY + (parsed.offsetY ?? 0);

  const bottomMargin = parsed.bottomMargin; // route.ts 側で undefined ならデフォルト 80

  const baseUrl =
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME
      ? `https://${process.env.WEBSITE_HOSTNAME}`
      : "http://localhost:3000");

  const genImageBase = baseUrl.replace(/\/+$/, "");
  console.log("[gen-image] base URL for overlay:", genImageBase);
  console.log("[gen-image] resolved style params:", {
    align,
    vAlign,
    size,
    color,
    fontFamily,
    bold,
    italic,
    offsetX,
    offsetY,
    bottomMargin,
  });

  try {
    const resp = await fetch(`${genImageBase}/api/gen-image`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        imageUrl: baseImageUrl, // ← ★ 毎回 __base__.png を元絵として使う
        text,
        align,
        vAlign,
        size, // small/medium/large/xlarge を route.ts 側で fontSize にマップ
        color,
        offsetX,
        offsetY,
        bottomMargin,
        autoDetectPlacard: false, // プラカード自動検出はここではOFF
        // ★ フォント指定（ここが新しく増えた）
        fontFamily, // "gothic" | "mincho" | "meiryo"
        bold,
        italic,
      }),
    });

    if (!resp.ok) {
      const t = await resp.text().catch(() => "");
      console.error(
        "🔴 /api/gen-image failed in edit:",
        resp.status,
        t
      );
      return {
        error: `Text overlay failed: HTTP ${resp.status}`,
      };
    }

    const result = await resp.json();
    const generatedPath = result?.imageUrl as string | undefined;

    if (!generatedPath) {
      console.error("🔴 gen-image edit returned no imageUrl");
      return { error: "gen-image edit returned no imageUrl" };
    }

    // /generated/xxx.png を Azure Storage の images コンテナに保存し直す
    const fs = require("fs");
    const path = require("path");
    const finalImageName = `${uniqueId()}.png`;
    const finalImagePath = path.join(
      process.cwd(),
      "public",
      generatedPath.startsWith("/")
        ? generatedPath.slice(1)
        : generatedPath
    );
    const finalImageBuffer = fs.readFileSync(finalImagePath);

    await UploadImageToStore(
      chatThread.id,
      finalImageName,
      finalImageBuffer
    );

    const finalImageUrl = buildExternalImageUrl(
      chatThread.id,
      finalImageName
    );

    return {
      revised_prompt: text,
      url: finalImageUrl,
    };
  } catch (err) {
    console.error("🔴 error in executeAddTextToExistingImage (simple):", err);
    return {
      error: "There was an error adding text to the existing image: " + err,
    };
  }
}
