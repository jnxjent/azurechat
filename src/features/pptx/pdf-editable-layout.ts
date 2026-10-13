import sharp from "sharp";
import { DocumentIntelligenceInstance } from "@/features/common/services/document-intelligence";
import { OpenAIVisionInstance, OpenAIPptVisionInstance } from "@/features/common/services/openai";
import { sourceFontFace } from "@/features/pptx/source-typography";
import { matchSourceFonts } from "@/features/pptx/source-font-match";

export type SourceBox = { x: number; y: number; w: number; h: number };
export type EditablePdfText = SourceBox & { text: string; color: string; fontSize: number; fontFace?: string; bold?: boolean; italic?: boolean; align?: "left" | "center" | "right" };
export type EditablePdfLayout = {
  backgroundColor: string;
  texts: EditablePdfText[];
};
export type NormalizedEditablePdfLayout = { layout: EditablePdfLayout; issues: string[] };

const MAX_ATTEMPTS = 3;
const MAX_TRANSIENT_RETRIES = 3;
/** Share of detected text boxes that may be discarded before asking Vision again. */
const MAX_DROPPED_RATIO = 0.2;
/** Contact-sheet panels are small; OCR is far more accurate on an enlarged copy. */
const MIN_VISION_LONG_SIDE = 1600;
/** Share of off-colour ink above which a Latin-only line is treated as a logo (see sampleColors). */
const MULTI_COLOR_LOGO = 0.12;

/**
 * Validates each text box independently. A broken box is dropped and reported
 * in `issues` instead of discarding the whole page, because Vision coordinates
 * are routinely a little off.
 */
export function normalizeEditablePdfLayout(raw: any): NormalizedEditablePdfLayout {
  if (raw?.texts != null && !Array.isArray(raw.texts)) throw new Error("textsが配列ではありません");
  const issues: string[] = [];
  const color = (v: unknown, fallback: string) => {
    const s = String(v ?? "").replace(/^#/, "");
    return /^[0-9a-f]{6}$/i.test(s) ? s : fallback;
  };
  const texts: EditablePdfText[] = [];
  (raw?.texts ?? []).forEach((t: any, i: number) => {
    if (typeof t?.text !== "string" || !t.text.trim()) return;
    const values = [t.x, t.y, t.w, t.h].map(v => v === null || v === undefined || v === "" ? NaN : Number(v));
    if (!values.every(Number.isFinite)) { issues.push(`文字領域${i}: 座標が不足しています`); return; }
    const [x, y, w, h] = values;
    const left = Math.max(0, x), top = Math.max(0, y);
    const right = Math.min(1, x + w), bottom = Math.min(1, y + h);
    // Clamp small overflows; drop boxes that are mostly outside the image.
    if (w <= 0 || h <= 0 || right - left < w * .5 || bottom - top < h * .5) {
      issues.push(`文字領域${i}: 座標が画像の範囲外です`);
      return;
    }
    const bounds = { x: left, y: top, w: right - left, h: bottom - top };
    const size = Number(t.fontSize);
    // Missing font size can be inferred from the detected text box and line count.
    const fontSize = Number.isFinite(size) && size > 0 && size <= .3 ? size : bounds.h / Math.max(1, t.text.split("\n").length) * .7;
    texts.push({ ...bounds, text: t.text, color: color(t.color, "222222"), fontSize, fontFace: sourceFontFace(t.fontFace ?? t.fontStyle), bold: t.bold === true, italic: t.italic === true,
      align: (["left", "center", "right"].includes(t.align) ? t.align : "left") as "left" | "center" | "right" });
  });
  return { layout: { backgroundColor: color(raw?.backgroundColor, "FFFFFF"), texts }, issues };
}

function isTransient(error: any): boolean {
  const status = error?.status ?? error?.response?.status;
  return [408, 409, 429, 500, 502, 503, 504].includes(status) ||
    /ECONNRESET|ETIMEDOUT|EAI_AGAIN|socket hang up|timeout|fetch failed/i.test(String(error?.message ?? error));
}

async function createWithRetry(params: any, usePptVision = false) {
  for (let retry = 1; ; retry++) {
    try {
      return await (usePptVision ? OpenAIPptVisionInstance() : OpenAIVisionInstance()).chat.completions.create(params);
    } catch (error) {
      if (!isTransient(error) || retry >= MAX_TRANSIENT_RETRIES) throw error;
      console.warn("[PDF editable layout] transient Vision error, retrying", { retry, status: (error as any)?.status });
      await new Promise(resolve => setTimeout(resolve, 1500 * retry));
    }
  }
}

async function enlargeForVision(dataUrl: string): Promise<string> {
  const bytes = Buffer.from(dataUrl.split(",")[1] ?? "", "base64");
  const { width = 0, height = 0 } = await sharp(bytes).metadata();
  const longSide = Math.max(width, height);
  if (!longSide || longSide >= MIN_VISION_LONG_SIDE) return dataUrl;
  const scale = MIN_VISION_LONG_SIDE / longSide;
  const enlarged = await sharp(bytes).resize(Math.round(width * scale), Math.round(height * scale), { kernel: "lanczos3" }).png().toBuffer();
  return `data:image/png;base64,${enlarged.toString("base64")}`;
}

async function recognizeWithVisionLayout(visionUrl: string): Promise<EditablePdfLayout> {
  let feedback = "";
  let lastError = "";
  let best: { layout: EditablePdfLayout; dropped: number } | null = null;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const response = await createWithRetry({
      model: process.env.AZURE_OPENAI_VISION_API_DEPLOYMENT_NAME!,
      response_format: { type: "json_object" },
      max_completion_tokens: 16000,
      messages: [{ role: "user", content: [
        { type: "text", text: `画像を編集可能なPowerPointへ復元するためのOCRと文字位置の解析をしてください。要約・加筆・言い換えは禁止。読める文字を正確に転記してください。JSONだけ返す:
{"backgroundColor":"FFFFFF","texts":[{"text":"原文","x":0.1,"y":0.1,"w":0.8,"h":0.1,"color":"123456","fontSize":0.04,"fontStyle":"gothic","bold":false,"italic":false,"align":"left"}]}
各行の元の書体・太さを保持する。fontStyleはgothic（日本語ゴシック）、mincho（明朝）、rounded（丸ゴシック）、sans（欧文サンセリフ）、serif（欧文セリフ）。重い見出しはbold=true、斜体はitalic=true。
座標と寸法は画像全体を1とした比率（小数3桁まで）。x,y,w,hは文字のインクをちょうど囲む矩形。fontSizeは1行の文字の高さ/画像高さ。textsは見出し・段落・表のセル・図のラベルを別々にし、改行を保存。写真内やロゴマーク自体の文字は転記しない（画像として残す）。ページ最下部も必ず確認し、会社名・担当者名・運行管理者・電話・内線などの連絡先欄は、ロゴの隣でも通常の文字として転記する。文字が1つもなければtextsは空配列。backgroundColorはページ余白に近い単色。${feedback ? `前回の解析の問題: ${feedback}。原画像を再確認し、指定したJSON形式で修正してください。` : ""}` },
        { type: "image_url", image_url: { url: visionUrl, detail: "high" } },
      ] }],
    });
    if (response.choices[0]?.finish_reason === "length") {
      lastError = feedback = "出力が長さ制限に達しました。座標は小数3桁までにし、JSON以外を出力しないでください";
      console.warn("[PDF editable layout] response truncated", { attempt });
      continue;
    }
    let result: NormalizedEditablePdfLayout;
    try {
      result = normalizeEditablePdfLayout(JSON.parse(response.choices[0]?.message?.content || "{}"));
    } catch (error) {
      lastError = feedback = error instanceof SyntaxError ? "JSON形式が不正です" : (error as Error).message;
      console.warn("[PDF editable layout] invalid response", { attempt, reason: feedback });
      continue;
    }
    const { layout, issues } = result;
    const total = layout.texts.length + issues.length;
    // Keep the attempt with the most usable text so a later empty or worse answer cannot replace it.
    if (!best || layout.texts.length > best.layout.texts.length || (layout.texts.length === best.layout.texts.length && issues.length < best.dropped)) {
      best = { layout, dropped: issues.length };
    }
    // An empty page is legitimate (photo-only page), but confirm it once, and never
    // accept it after an earlier attempt did find text.
    const needsConfirmation = layout.texts.length === 0 && (attempt === 1 || best.layout.texts.length > 0);
    if (!needsConfirmation && issues.length <= Math.max(1, total * MAX_DROPPED_RATIO)) return layout;
    feedback = needsConfirmation
      ? "文字領域が0件でした。画像内に読める文字が本当にないか確認してください"
      : issues.slice(0, 5).join("、");
    console.warn("[PDF editable layout] retrying with feedback", { attempt, dropped: issues.length, total, feedback });
  }
  if (best) {
    console.warn("[PDF editable layout] using best attempt", { texts: best.layout.texts.length, dropped: best.dropped });
    return best.layout;
  }
  throw new Error(`PDFの文字配置を解析できませんでした（${lastError}）。`);
}

type OcrLine = SourceBox & { text: string };

/**
 * True for a fix of a few misread glyphs (e.g. 従業員敗→従業員数). Rewrites, drops and
 * additions change more characters than the edit budget allows and are refused.
 */
export function isMisreadCorrection(ocr: string, corrected: string): boolean {
  // Full-/half-width variants (（ vs (, ・ vs ·) are not edits.
  const canonical = (s: string) => Array.from(s.normalize("NFKC").replace(/[·･]/g, "・"));
  const a = canonical(ocr), b = canonical(corrected);
  let previous = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const current = [i];
    for (let j = 1; j <= b.length; j++) {
      current[j] = Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    previous = current;
  }
  const budget = Math.max(1, Math.min(a.length - 1, Math.max(2, Math.round(a.length * .2))));
  // Misreads are substitutions; a length change beyond one glyph means words were dropped or added.
  return Math.abs(a.length - b.length) <= 1 && previous[b.length] <= budget;
}

async function readLinesWithDocumentIntelligence(png: Buffer): Promise<OcrLine[]> {
  for (let retry = 1; ; retry++) {
    try {
      const poller = await DocumentIntelligenceInstance().beginAnalyzeDocument("prebuilt-read", png);
      const page = (await poller.pollUntilDone()).pages?.[0];
      if (!page?.width || !page.height) return [];
      const pw = page.width, ph = page.height;
      const toBox = (polygon: Array<{ x: number; y: number }>) => {
        const xs = polygon.map(p => p.x), ys = polygon.map(p => p.y);
        const x = Math.max(0, Math.min(...xs) / pw), y = Math.max(0, Math.min(...ys) / ph);
        return { x, y, w: Math.min(1, Math.max(...xs) / pw) - x, h: Math.min(1, Math.max(...ys) / ph) - y };
      };
      const result: OcrLine[] = [];
      for (const line of page.lines ?? []) {
        if (!line.content.trim() || !line.polygon?.length) continue;
        const box = toBox(line.polygon);
        // Letter spacing in contact footers must not split a role/name into
        // separate glyphs that the logo filter can mistake for a brand mark.
        if (box.y >= .8 && /運行管理者|管理責任者|担当者|連絡先|内線|電話|株式会社|ホールディングス|[（(]株[）)]/.test(line.content.replace(/\s/g, ""))) {
          result.push({ text: line.content.trim(), ...box });
          continue;
        }
        // DI sometimes joins a heading and a logo on the same baseline into one line.
        // Split only at column-sized gaps. Japanese tracking and OCR word
        // polygons can leave gaps as wide as a glyph; splitting there drops
        // unrecognised characters between word spans and changes fonts mid-line.
        const lineOffset = line.spans?.[0]?.offset;
        const words = typeof line.words === "function" && lineOffset !== undefined
          ? Array.from(line.words()).filter(word => word.polygon?.length).map(word => ({ ...toBox(word.polygon!), offset: word.span.offset, length: word.span.length }))
          : [];
        const segments: Array<typeof words> = [];
        // Median word height, not the line height: a tall logo inflates the line box.
        const wordHeight = words.map(w => w.h).sort((a, b) => a - b)[words.length >> 1] ?? box.h;
        for (const word of words.sort((a, b) => a.x - b.x)) {
          const last = segments[segments.length - 1]?.slice(-1)[0];
          if (last && (word.x - (last.x + last.w)) * pw > wordHeight * ph * 2) segments.push([word]);
          else if (last) segments[segments.length - 1].push(word);
          else segments.push([word]);
        }
        if (segments.length < 2) {
          result.push({ text: line.content.trim(), ...box });
          continue;
        }
        for (const segment of segments) {
          const first = segment[0], end = segment[segment.length - 1];
          const text = line.content.slice(first.offset - lineOffset!, end.offset + end.length - lineOffset!).trim();
          const x = first.x, y = Math.min(...segment.map(w => w.y));
          if (text) result.push({ text, x, y, w: end.x + end.w - x, h: Math.max(...segment.map(w => w.y + w.h)) - y });
        }
      }
      return result.filter(line => line.w > 0 && line.h > 0);
    } catch (error) {
      if (!isTransient(error) || retry >= MAX_TRANSIENT_RETRIES) throw error;
      console.warn("[PDF editable layout] transient Document Intelligence error, retrying", { retry });
      await new Promise(resolve => setTimeout(resolve, 1500 * retry));
    }
  }
}

/**
 * OCR positions come from Document Intelligence; Vision only fixes misread
 * characters and marks text that belongs to a logo or photo. Corrections that
 * replace most of the characters are rejected so the model cannot rewrite text.
 */
async function correctOcrLines(visionUrl: string, lines: OcrLine[]): Promise<Array<{ text: string; keep: boolean; bold: boolean; italic: boolean; fontFace: string }>> {
  const result = lines.map(line => ({ text: line.text, keep: true, bold: false, italic: false, fontFace: "Meiryo" }));
  const typographyDeployment = process.env.AZURE_OPENAI_PPT_VISION_DEPLOYMENT_NAME?.trim();
  // Enlarged per-line references prevent tiny body type from being mistaken for
  // regular weight when the full page is resized by Vision.
  const source = Buffer.from(visionUrl.split(",")[1], "base64");
  const { width = 1, height = 1 } = await sharp(source).metadata();
  const fontSheets: string[] = [];
  for (let start = 0; start < Math.min(lines.length, 72); start += 12) {
    const group = lines.slice(start, start + 12);
    const overlays: Array<{ input: Buffer; top: number; left: number }> = [];
    for (let index = 0; index < group.length; index++) {
      const line = group[index];
      const left = Math.max(0, Math.floor(line.x * width)), top = Math.max(0, Math.floor(line.y * height));
      const cropW = Math.max(1, Math.min(width - left, Math.ceil(line.w * width)));
      const cropH = Math.max(1, Math.min(height - top, Math.ceil(line.h * height)));
      const crop = await sharp(source).extract({ left, top, width: cropW, height: cropH })
        .resize(1480, 64, { fit: "inside" }).png().toBuffer();
      overlays.push({ input: crop, left: 100, top: index * 80 + 8 });
      const label = Buffer.from(`<svg width="90" height="80"><text x="8" y="48" font-size="28">${start + index}</text></svg>`);
      overlays.push({ input: label, left: 0, top: index * 80 });
    }
    const sheet = await sharp({ create: { width: 1600, height: group.length * 80, channels: 3, background: "white" } }).composite(overlays).png().toBuffer();
    fontSheets.push(`data:image/png;base64,${sheet.toString("base64")}`);
  }
  for (let attempt = 1; attempt <= 2; attempt++) try {
    const response = await createWithRetry({
      model: typographyDeployment || process.env.AZURE_OPENAI_VISION_API_DEPLOYMENT_NAME!,
      response_format: { type: "json_object" },
      reasoning_effort: "low",
      max_completion_tokens: 8000,
      messages: [{ role: "user", content: [
        { type: "text", text: `画像のOCR結果を画像と照合して校正してください。各行について、画像と違う誤認識の文字だけを直してください。要約・加筆・言い換え・行の統合や分割は禁止。keep=falseはロゴマークの文字と、写真に写り込んだ看板・ラベル等の文字だけ。写真や図に重ねたキャプション・見出し・ページ番号は通常の文字なのでkeep=true。太字ならbold=true。JSONだけ返す:
{"lines":[{"i":0,"text":"校正後の文字","keep":true,"bold":false,"italic":false,"fontStyle":"gothic"}]}
各行の元画像の書体と太さも判定する。fontStyleはgothic（日本語ゴシック）、mincho（明朝）、rounded（丸ゴシック）、sans（欧文サンセリフ）、serif（欧文セリフ）。大見出しの重い字形はbold=true。斜体ならitalic=true。OCRの行ごとに判定し、全行を同じ書体にしない。
2枚目以降は行番号付きの拡大画像です。書体・太さの判定はこの拡大画像の実際の線を優先してください。本文・フッターも見出しと同じように太いなら必ずbold=true。本文だから通常の太さと推測しない。行番号は原文に含めない。
OCR結果:
${JSON.stringify(lines.map((line, i) => ({ i, text: line.text })))}` },
        { type: "image_url", image_url: { url: visionUrl, detail: "high" } },
        ...fontSheets.map(url => ({ type: "image_url" as const, image_url: { url, detail: "high" as const } })),
      ] }],
    }, Boolean(typographyDeployment));
    const parsed = JSON.parse(response.choices[0]?.message?.content || "{}");
    if (!Array.isArray(parsed?.lines)) throw new Error("linesがありません");
    for (const item of parsed.lines) {
      const i = Number(item?.i);
      if (!Number.isInteger(i) || !result[i]) continue;
      if (item.keep === false) result[i].keep = false;
      result[i].bold = item.bold === true;
      result[i].italic = item.italic === true;
      result[i].fontFace = sourceFontFace(item.fontStyle);
      const text = typeof item.text === "string" ? item.text.trim() : "";
      if (text && text !== lines[i].text) {
        if (isMisreadCorrection(lines[i].text, text)) result[i].text = text;
        else console.warn("[PDF editable layout] rejected correction", { ocr: lines[i].text, corrected: text });
      }
    }
    return result;
  } catch (error) {
    // The OCR text is still usable as is, so a failed correction never fails the conversion.
    console.warn("[PDF editable layout] OCR correction failed", { attempt, reason: (error as Error).message });
  }
  return result;
}

/**
 * Text colour = the pixels inside the box that differ most from its surrounding ring.
 * `multiColor` is the share of ink pixels that are not a blend of that colour and the
 * background: plain text stays near 0 (≤0.07 on real pages), a multi-hue logo is ~0.17.
 */
async function sampleColors(png: Buffer, boxes: SourceBox[]): Promise<{ background: string; colors: string[]; multiColor: number[] }> {
  const { data, info } = await sharp(png).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  const { width: W, height: H, channels: C } = info;
  const hex = (rgb: number[]) => rgb.map(v => Math.round(v).toString(16).padStart(2, "0")).join("").toUpperCase();
  const median = (pixels: number[][]) => [0, 1, 2].map(c => pixels.map(p => p[c]).sort((a, b) => a - b)[pixels.length >> 1] ?? 255);
  const at = (x: number, y: number) => { const p = (y * W + x) * C; return [data[p], data[p + 1], data[p + 2]]; };
  const border: number[][] = [];
  for (let x = 0; x < W; x += 2) border.push(at(x, 0), at(x, H - 1));
  for (let y = 0; y < H; y += 2) border.push(at(0, y), at(W - 1, y));
  const measured = boxes.map(box => {
    const x0 = Math.max(0, Math.floor(box.x * W)), x1 = Math.min(W - 1, Math.ceil((box.x + box.w) * W));
    const y0 = Math.max(0, Math.floor(box.y * H)), y1 = Math.min(H - 1, Math.ceil((box.y + box.h) * H));
    const ring: number[][] = [];
    for (let x = x0; x <= x1; x++) ring.push(at(x, Math.max(0, y0 - 2)), at(x, Math.min(H - 1, y1 + 2)));
    for (let y = y0; y <= y1; y++) ring.push(at(Math.max(0, x0 - 2), y), at(Math.min(W - 1, x1 + 2), y));
    const bg = median(ring);
    const inside: Array<{ d: number; rgb: number[] }> = [];
    for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) {
      const rgb = at(x, y);
      inside.push({ d: Math.hypot(rgb[0] - bg[0], rgb[1] - bg[1], rgb[2] - bg[2]), rgb });
    }
    const ink = inside.filter(p => p.d > 40);
    const strongest = [...ink].sort((a, b) => b.d - a.d).slice(0, Math.max(1, Math.round(inside.length * .08)));
    if (!strongest.length) return { color: "222222", multiColor: 0 };
    const fg = [0, 1, 2].map(c => strongest.reduce((s, p) => s + p.rgb[c], 0) / strongest.length);
    // Distance of each ink pixel from the background→text colour blend line.
    const axis = fg.map((v, c) => v - bg[c]);
    const axisLength2 = axis.reduce((s, v) => s + v * v, 0) || 1;
    const offAxis = ink.filter(({ rgb }) => {
      // Not clamped at 1: ink darker than the averaged text colour is still on the same line.
      const u = Math.max(0, rgb.reduce((s, v, c) => s + (v - bg[c]) * axis[c], 0) / axisLength2);
      return Math.hypot(...rgb.map((v, c) => v - (bg[c] + u * axis[c]))) > 45;
    }).length;
    return { color: hex(fg), multiColor: offAxis / ink.length };
  });
  return { background: hex(median(border)), colors: measured.map(m => m.color), multiColor: measured.map(m => m.multiColor) };
}

export function isContactFooterText(line: { text: string; y: number }, lines: Array<{ text: string; y: number }>): boolean {
  const contactLabel = /運行管理者|管理責任者|担当者|連絡先|内線|電話|TEL\s*[:：]/i;
  return line.y >= .8 && (contactLabel.test(line.text.replace(/\s/g, "")) ||
    (/[一-鿿ぁ-ゖァ-ヺ]/.test(line.text) && lines.some(other => other.y >= .8 && contactLabel.test(other.text.replace(/\s/g, "")))));
}

async function recognizeWithDocumentIntelligence(dataUrl: string, visionUrl: string): Promise<EditablePdfLayout> {
  const source = Buffer.from(dataUrl.split(",")[1] ?? "", "base64");
  const lines = await readLinesWithDocumentIntelligence(Buffer.from(visionUrl.split(",")[1] ?? "", "base64"));
  const corrected = lines.length ? await correctOcrLines(visionUrl, lines) : [];
  const { background, colors, multiColor } = await sampleColors(source, lines);
  const merged = lines.map((line, i) => ({
    ...line, ...corrected[i], color: colors[i],
    // Vision's logo judgement varies between runs, so multi-hue Latin-only marks (logos) are also kept as image.
    keep: isContactFooterText(line, lines) || (corrected[i].keep && !(multiColor[i] > MULTI_COLOR_LOGO && !/[　-鿿＀-￯]/.test(line.text))),
  }));
  const excluded = merged.filter(line => !line.keep).map(line => line.text);
  if (excluded.length) console.info("[PDF editable layout] kept as image (logo/photo text)", { excluded });
  return {
    backgroundColor: background,
    texts: merged.filter(line => line.keep).map(line => {
      // Font size is fitted from actual glyph metrics when writing the slide.
      return { x: line.x, y: line.y, w: line.w, h: line.h, text: line.text, color: line.color,
        fontSize: line.h, fontFace: line.fontFace, bold: line.bold, italic: line.italic, align: "left" as const };
    }),
  };
}

export async function recognizeEditablePdfLayout(dataUrl: string): Promise<EditablePdfLayout> {
  const visionUrl = await enlargeForVision(dataUrl);
  if (process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT && process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY) {
    try {
      const layout = await recognizeWithDocumentIntelligence(dataUrl, visionUrl);
      return { ...layout, texts: await matchSourceFonts(Buffer.from(dataUrl.split(",")[1], "base64"), layout.texts) };
    } catch (error) {
      console.warn("[PDF editable layout] Document Intelligence failed; using Vision layout", { reason: (error as Error).message });
    }
  }
  const layout = await recognizeWithVisionLayout(visionUrl);
  return { ...layout, texts: await matchSourceFonts(Buffer.from(dataUrl.split(",")[1], "base64"), layout.texts) };
}
