import PptxGenJS from "pptxgenjs";
import sharp from "sharp";
import type { EditablePdfLayout, EditablePdfText } from "./pdf-editable-layout";

export type PdfImageSlide = { sourceImageDataUrl: string; sourceWidth: number; sourceHeight: number; editableLayout?: EditablePdfLayout };

/**
 * Paints over the recognized text so the editable text boxes do not sit on top
 * of the original glyphs. Every masked pixel is interpolated from the nearest
 * unmasked pixels on its row and column, which keeps cell fills, gradients and
 * photo edges around the text intact. All other pixels stay untouched.
 */
export async function eraseTextRegions(png: Buffer, texts: Pick<EditablePdfText, "x" | "y" | "w" | "h" | "fontSize">[]): Promise<Buffer> {
  const { data, info } = await sharp(png).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  const { width: W, height: H, channels: C } = info;
  const mask = new Uint8Array(W * H);
  for (const t of texts) {
    // Pad by a fraction of the line height so anti-aliased glyph edges are covered too.
    const pad = Math.max(2, Math.round(Math.max(t.fontSize * H, t.h * H / 3) * .3));
    const x0 = Math.max(0, Math.floor(t.x * W) - pad), x1 = Math.min(W, Math.ceil((t.x + t.w) * W) + pad);
    const y0 = Math.max(0, Math.floor(t.y * H) - pad), y1 = Math.min(H, Math.ceil((t.y + t.h) * H) + pad);
    for (let y = y0; y < y1; y++) mask.fill(1, y * W + x0, y * W + x1);
  }
  const sum = new Float64Array(W * H * C), weight = new Float64Array(W * H);
  // Per-channel median of up to 5 unmasked pixels walking outward from a run boundary,
  // so a stray glyph edge next to the box does not bleed into the fill.
  const sample = (start: number, step: number, limit: number, index: (i: number) => number): number[] | null => {
    const picked: number[] = [];
    for (let i = start; i >= 0 && i < limit && picked.length < 5; i += step) {
      const p = index(i);
      if (!mask[p]) picked.push(p);
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
        const fontInches = box.fontSize * h;
        const lines = box.text.split("\n").length;
        // OCR boxes hug the glyphs; give the shape room for the font's line height, centred on the detected line.
        const boxH = Math.max(box.h * h, fontInches * 1.35 * lines);
        const centerY = oy + (box.y + box.h / 2) * h;
        slide.addText(box.text, { x: ox + box.x * w, y: centerY - boxH / 2, w: box.w * w, h: boxH,
          fontFace: "Yu Gothic", fontSize: Math.max(5, fontInches * 72), color: /^[0-9a-f]{6}$/i.test(box.color) ? box.color : "222222",
          bold: box.bold === true, align: ["left", "center", "right"].includes(box.align || "") ? box.align : "left",
          margin: 0, breakLine: false, valign: "middle", wrap: lines > 1 });
      }
    } else {
      // No recognized text (photo-only page, or no layout requested): the source page is kept as is.
      slide.addImage({ data: source.sourceImageDataUrl, x: ox, y: oy, w, h });
    }
  }
  return await pptx.write({ outputType: "nodebuffer" }) as Buffer;
}
