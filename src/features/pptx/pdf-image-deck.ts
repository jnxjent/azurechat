import PptxGenJS from "pptxgenjs";
import sharp from "sharp";
import { createCanvas } from "@napi-rs/canvas";
import { sourceFontFace, sourceMeasurementFont } from "@/features/pptx/source-typography";
import type { EditablePdfLayout, EditablePdfText } from "./pdf-editable-layout";

export type PdfImageSlide = { sourceImageDataUrl: string; sourceWidth: number; sourceHeight: number; editableLayout?: EditablePdfLayout };

/** Fit the actual font's ink to OCR bounds; avoid reducing every line with an em heuristic. */
export function fitSourceTypography(box: EditablePdfText, slideW: number, slideH: number) {
  const fontFace = sourceFontFace(box.fontFace);
  const context = createCanvas(1, 1).getContext("2d");
  const sampleSize = 100;
  context.font = `${box.italic ? "italic " : ""}${box.bold ? "bold " : ""}${sampleSize}px "${sourceMeasurementFont(fontFace)}"`;
  const lines = box.text.split("\n");
  const metrics = lines.map(line => context.measureText(line));
  const inkH = Math.max(...metrics.map(m => m.actualBoundingBoxAscent + m.actualBoundingBoxDescent), 1);
  const inkW = Math.max(...metrics.map(m => m.width), 1);
  const targetH = box.h * slideH * 72 / lines.length;
  const targetW = box.w * slideW * 72;
  const sizeByHeight = targetH * sampleSize / inkH;
  const sizeByWidth = targetW * sampleSize / inkW;
  const fontSize = Math.max(5, Math.min(sizeByHeight, sizeByWidth));
  const chars = Math.max(...lines.map(line => Array.from(line).length));
  // Widely spaced headings and contact fields should span their source line.
  const charSpacing = chars > 1 && lines.length === 1
    ? Math.max(0, Math.min(fontSize * .5, (targetW - inkW * fontSize / sampleSize) / (chars - 1))) : 0;
  return { fontFace, fontSize, charSpacing };
}

/**
 * Remove source glyphs before adding editable text. For solid fills, sample
 * background inside the text area and replace only foreground ink. For complex
 * backgrounds, interpolate masked pixels from nearby unmasked rows/columns.
 */
export async function eraseTextRegions(png: Buffer, texts: (Pick<EditablePdfText, "x" | "y" | "w" | "h" | "fontSize"> & { color?: string })[]): Promise<Buffer> {
  const { data, info } = await sharp(png).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  const { width: W, height: H, channels: C } = info;
  const mask = new Uint8Array(W * H);
  const solidRepairs: Array<{ pixels: number[]; background: number[] }> = [];
  for (const t of texts) {
    // Pad by a fraction of the line height so anti-aliased glyph edges are covered too.
    const pad = Math.max(2, Math.round(Math.max(t.fontSize * H, t.h * H / 3) * .3));
    const x0 = Math.max(0, Math.floor(t.x * W) - pad), x1 = Math.min(W, Math.ceil((t.x + t.w) * W) + pad);
    const y0 = Math.max(0, Math.floor(t.y * H) - pad), y1 = Math.min(H, Math.ceil((t.y + t.h) * H) + pad);
    if (/^[0-9a-f]{6}$/i.test(t.color || "")) {
      const fg = [0, 2, 4].map(i => parseInt(t.color!.slice(i, i + 2), 16));
      const left = Math.max(0, Math.floor(t.x * W)), right = Math.min(W, Math.ceil((t.x + t.w) * W));
      const top = Math.max(0, Math.floor(t.y * H)), bottom = Math.min(H, Math.ceil((t.y + t.h) * H));
      const buckets = new Map<number, { count: number; sum: number[] }>();
      let backgroundCount = 0;
      for (let y = top; y < bottom; y++) for (let x = left; x < right; x++) {
        const p = (y * W + x) * C, rgb = [data[p], data[p + 1], data[p + 2]];
        // Small rasterised PDF text can have a light sampled foreground colour.
        // A broad exclusion radius would discard the white background too.
        if (Math.hypot(...rgb.map((v, c) => v - fg[c])) < 30) continue;
        const key = (rgb[0] >> 4) * 256 + (rgb[1] >> 4) * 16 + (rgb[2] >> 4);
        const bucket = buckets.get(key) || { count: 0, sum: [0, 0, 0] };
        bucket.count++; rgb.forEach((v, c) => { bucket.sum[c] += v; }); buckets.set(key, bucket); backgroundCount++;
      }
      const ranked = Array.from(buckets.values()).sort((a, b) => b.count - a.count);
      // A shaded orange label can straddle several quantisation buckets. Treat
      // nearby shades as one background instead of falling back to frame colours.
      let dominant: { count: number; sum: number[] } | undefined;
      for (const seed of ranked.slice(0, 8)) {
        const centre = seed.sum.map(v => v / seed.count);
        const cluster = { count: 0, sum: [0, 0, 0] };
        for (const bucket of ranked) {
          if (Math.hypot(...bucket.sum.map((v, c) => v / bucket.count - centre[c])) > 48) continue;
          cluster.count += bucket.count;
          bucket.sum.forEach((v, c) => { cluster.sum[c] += v; });
        }
        if (!dominant || cluster.count > dominant.count) dominant = cluster;
      }
      if (dominant && dominant.count >= backgroundCount * .6 && backgroundCount > (right - left) * (bottom - top) * .08) {
        const bg = dominant.sum.map(v => v / dominant.count), axis = fg.map((v, c) => v - bg[c]);
        const norm = axis.reduce((sum, v) => sum + v * v, 0);
        if (norm > 30 * 30) {
          const pixels: number[] = [];
          // Remove ink rather than the whole padded rectangle: neighbouring
          // frames and coloured bands must not become interpolation samples.
          for (let y = Math.max(0, top - 2); y < Math.min(H, bottom + 2); y++) for (let x = Math.max(0, left - 2); x < Math.min(W, right + 2); x++) {
            const p = (y * W + x) * C, rgb = [data[p], data[p + 1], data[p + 2]];
            const alpha = rgb.reduce((sum, v, c) => sum + (v - bg[c]) * axis[c], 0) / norm;
            const error = Math.hypot(...rgb.map((v, c) => v - bg[c] - alpha * axis[c]));
            // The sampled foreground averages antialiasing. Actual black ink
            // can be much darker (alpha > 1); it still belongs to the same ink.
            if (alpha > .025 && error < 45) pixels.push(y * W + x);
          }
          solidRepairs.push({ pixels, background: bg });
          continue;
        }
      }
    }
    for (let y = y0; y < y1; y++) mask.fill(1, y * W + x0, y * W + x1);
  }
  // Solid repairs are applied later. Their original glyphs must never be
  // sampled as background for an adjacent line using interpolation.
  const sampleMask = mask.slice();
  for (const repair of solidRepairs) for (const p of repair.pixels) sampleMask[p] = 1;
  const sum = new Float64Array(W * H * C), weight = new Float64Array(W * H);
  // Per-channel median of up to 5 unmasked pixels walking outward from a run boundary,
  // so a stray glyph edge next to the box does not bleed into the fill.
  const sample = (start: number, step: number, limit: number, index: (i: number) => number): number[] | null => {
    const picked: number[] = [];
    for (let i = start; i >= 0 && i < limit && picked.length < 5; i += step) {
      const p = index(i);
      if (!sampleMask[p]) picked.push(p);
    }
    if (!picked.length) return null;
    return Array.from({ length: C }, (_, c) => picked.map(p => data[p * C + c]).sort((a, b) => a - b)[picked.length >> 1]);
  };
  const pass = (outer: number, inner: number, index: (o: number, i: number) => number) => {
    for (let o = 0; o < outer; o++) {
      for (let i = 0; i < inner; i++) {
        if (!mask[index(o, i)]) continue;
        let end = i;
        while (end < inner && mask[index(o, end)]) end++;
        const before = sample(i - 1, -1, inner, k => index(o, k));
        const after = sample(end, 1, inner, k => index(o, k));
        const span = end - i;
        for (let k = i; k < end; k++) {
          const u = (k - i + 1) / (span + 1);
          const color = before && after ? before.map((v, c) => v * (1 - u) + after[c] * u) : before ?? after;
          if (!color) continue;
          // Shorter spans are more reliable: text boxes are wide, so the vertical pass usually dominates.
          const p = index(o, k), w = 1 / span;
          for (let c = 0; c < C; c++) sum[p * C + c] += color[c] * w;
          weight[p] += w;
        }
        i = end;
      }
    }
  };
  pass(H, W, (y, x) => y * W + x);
  pass(W, H, (x, y) => y * W + x);
  for (let p = 0; p < W * H; p++) {
    if (!mask[p] || !weight[p]) continue;
    for (let c = 0; c < C; c++) data[p * C + c] = Math.round(sum[p * C + c] / weight[p]);
  }
  // Column interpolation leaves faint vertical streaks; a short horizontal mean inside the mask removes them.
  const RADIUS = 4;
  for (let y = 0; y < H; y++) {
    const row = data.slice(y * W * C, (y + 1) * W * C);
    for (let x = 0; x < W; x++) {
      if (!mask[y * W + x]) continue;
      const acc = new Array(C).fill(0);
      let n = 0;
      for (let k = Math.max(0, x - RADIUS); k <= Math.min(W - 1, x + RADIUS); k++) {
        if (!mask[y * W + k]) continue;
        for (let c = 0; c < C; c++) acc[c] += row[k * C + c];
        n++;
      }
      for (let c = 0; c < C; c++) data[(y * W + x) * C + c] = Math.round(acc[c] / n);
    }
  }
  for (const repair of solidRepairs) for (const p of repair.pixels) {
    repair.background.forEach((v, c) => { data[p * C + c] = Math.round(v); });
  }
  return await sharp(data, { raw: { width: W, height: H, channels: C } }).png().toBuffer();
}

/** Preserve source pixels, aspect ratio, page order; never add a synthetic cover. */
export async function createPdfImageDeck(slides: PdfImageSlide[]): Promise<Buffer> {
  if (!slides.length || slides.some(s => !/^data:image\/png;base64,/.test(s.sourceImageDataUrl) || !(s.sourceWidth > 0 && s.sourceHeight > 0))) {
    throw new Error("PDFページ画像または寸法が不正です。");
  }
  const pptx = new PptxGenJS();
  const width = 10, height = width * slides[0].sourceHeight / slides[0].sourceWidth;
  pptx.defineLayout({ name: "SOURCE", width, height });
  pptx.layout = "SOURCE";
  for (const source of slides) {
    const scale = Math.min(width / source.sourceWidth, height / source.sourceHeight);
    const w = source.sourceWidth * scale, h = source.sourceHeight * scale;
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    const ox = (width - w) / 2, oy = (height - h) / 2;
    const layout = source.editableLayout;
    if (layout?.texts.length) {
      slide.background = { color: /^[0-9a-f]{6}$/i.test(layout.backgroundColor) ? layout.backgroundColor : "FFFFFF" };
      const erased = await eraseTextRegions(Buffer.from(source.sourceImageDataUrl.split(",")[1], "base64"), layout.texts);
      slide.addImage({ data: `data:image/png;base64,${erased.toString("base64")}`, x: ox, y: oy, w, h, altText: "Source figures" });
      for (const box of layout.texts) {
        const typography = fitSourceTypography(box, w, h);
        const fontInches = typography.fontSize / 72;
        const lines = box.text.split("\n").length;
        // OCR boxes hug the glyphs; give the shape room for the font's line height, centred on the detected line.
        const boxH = Math.max(box.h * h, fontInches * 1.35 * lines);
        const centerY = oy + (box.y + box.h / 2) * h;
        slide.addText(box.text, { x: ox + box.x * w, y: centerY - boxH / 2, w: box.w * w, h: boxH,
          ...typography, color: /^[0-9a-f]{6}$/i.test(box.color) ? box.color : "222222",
          bold: box.bold === true, italic: box.italic === true, align: ["left", "center", "right"].includes(box.align || "") ? box.align : "left",
          margin: 0, breakLine: false, valign: "middle", wrap: lines > 1 });
      }
    } else {
      // No recognized text (photo-only page, or no layout requested): the source page is kept as is.
      slide.addImage({ data: source.sourceImageDataUrl, x: ox, y: oy, w, h });
    }
  }
  return await pptx.write({ outputType: "nodebuffer" }) as Buffer;
}
