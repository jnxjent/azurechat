import { createCanvas, GlobalFonts } from "@napi-rs/canvas";
import sharp from "sharp";
import type { EditablePdfText } from "./pdf-editable-layout";
import { registerSourceFonts } from "@/features/pptx/source-typography";

// Remove gaps between glyphs before comparing: OCR omits typographic spaces,
// which otherwise misalign every later character and make bold look too light.
function packInkColumns(mask: Buffer, width: number, height: number): { data: Buffer; width: number } {
  const columns: number[] = [];
  for (let x = 0; x < width; x++) {
    let ink = 0;
    for (let y = 0; y < height; y++) if (mask[y * width + x] >= 128) ink++;
    if (ink) columns.push(x);
  }
  const packed = Buffer.alloc(Math.max(1, columns.length) * height);
  for (let y = 0; y < height; y++) columns.forEach((x, i) => { packed[y * columns.length + i] = mask[y * width + x]; });
  return { data: packed, width: Math.max(1, columns.length) };
}

/** Compare editable Office font candidates with the original ink, including weight.
 * Vision identifies broad families; pixel comparison corrects missed bold body text. */
export async function matchSourceFonts(png: Buffer, texts: EditablePdfText[], report?: (match: { text: string; score: number; face: string; bold: boolean }) => void): Promise<EditablePdfText[]> {
  registerSourceFonts();
  const { data, info } = await sharp(png).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  const installed = new Set(GlobalFonts.families.map(f => f.family));
  const output: EditablePdfText[] = [];
  for (const box of texts) {
    if (box.text.includes("\n") || box.text.trim().length < 2 || box.h * info.height < 10) { output.push(box); continue; }
    const x = Math.max(0, Math.floor(box.x * info.width)), y = Math.max(0, Math.floor(box.y * info.height));
    const w = Math.min(info.width - x, Math.ceil(box.w * info.width)), h = Math.min(info.height - y, Math.ceil(box.h * info.height));
    if (w < 2 || h < 2) { output.push(box); continue; }
    const fg = [0, 2, 4].map(i => parseInt(box.color.slice(i, i + 2), 16));
    if (!fg.every(Number.isFinite)) { output.push(box); continue; }
    const mask = Buffer.alloc(w * h);
    let count = 0, left = w, right = 0, top = h, bottom = 0;
    for (let row = 0; row < h; row++) for (let col = 0; col < w; col++) {
      const p = ((y + row) * info.width + x + col) * info.channels;
      if (Math.hypot(data[p] - fg[0], data[p + 1] - fg[1], data[p + 2] - fg[2]) < 90) {
        mask[row * w + col] = 255; count++;
        left = Math.min(left, col); right = Math.max(right, col); top = Math.min(top, row); bottom = Math.max(bottom, row);
      }
    }
    // Insufficient contrast / solid fills cannot reliably identify a font.
    if (count < 8 || count > w * h * .85 || right <= left || bottom <= top) { output.push(box); continue; }
    const tight = await sharp(mask, { raw: { width: w, height: h, channels: 1 } })
      .extract({ left, top, width: right - left + 1, height: bottom - top + 1 })
      .greyscale().raw().toBuffer();
    const packedSource = packInkColumns(tight, right - left + 1, bottom - top + 1);
    const targetW = Math.min(512, packedSource.width), targetH = Math.min(48, bottom - top + 1);
    const sourceMask = await sharp(packedSource.data, { raw: { width: packedSource.width, height: bottom - top + 1, channels: 1 } })
      .resize(targetW, targetH, { fit: "fill" }).greyscale().raw().toBuffer();
    const sourceInk = sourceMask.reduce((sum, v) => sum + (v >= 128 ? 1 : 0), 0);
    const isSerif = /Mincho|Serif|Times/.test(box.fontFace || "");
    const families = isSerif ? [box.fontFace || "Yu Mincho", "Yu Mincho", "Noto Serif JP"]
      : /丸|Maru/.test(box.fontFace || "") ? [box.fontFace!]
      : [box.fontFace || "Meiryo", "Noto Sans JP", "Yu Gothic", "Meiryo", "BIZ UDGothic"];
    let best: { score: number; face: string; bold: boolean } | undefined;
    let bestShape: { score: number; face: string; bold: boolean } | undefined;
    const candidates: Array<{ score: number; face: string; bold: boolean }> = [];
    // Never lighten explicitly recognised bold text merely because another
    // family's regular glyphs have a similar silhouette.
    for (const face of Array.from(new Set(families)).filter(face => installed.has(face))) for (const bold of box.bold ? [true] : [false, true]) {
      const canvas = createCanvas(1, 1), ctx = canvas.getContext("2d");
      const font = `${box.italic ? "italic " : ""}${bold ? "700" : "400"} 64px "${face}"`;
      ctx.font = font;
      const m = ctx.measureText(box.text);
      const rw = Math.ceil(m.actualBoundingBoxLeft + m.actualBoundingBoxRight), rh = Math.ceil(m.actualBoundingBoxAscent + m.actualBoundingBoxDescent);
      if (rw < 1 || rh < 1) continue;
      canvas.width = rw + 4; canvas.height = rh + 4;
      ctx.font = font; ctx.fillStyle = "white";
      ctx.fillText(box.text, m.actualBoundingBoxLeft + 2, m.actualBoundingBoxAscent + 2);
      const rgba = ctx.getImageData(2, 2, rw, rh).data;
      const alpha = Buffer.alloc(rw * rh);
      for (let p = 0; p < alpha.length; p++) alpha[p] = rgba[p * 4 + 3];
      const packedCandidate = packInkColumns(alpha, rw, rh);
      const rendered = await sharp(packedCandidate.data, { raw: { width: packedCandidate.width, height: rh, channels: 1 } })
        .resize(targetW, targetH, { fit: "fill" }).greyscale().raw().toBuffer();
      let candidateInk = 0;
      for (let p = 0; p < rendered.length; p++) if (rendered[p] >= 128) candidateInk++;
      let score = 0;
      // Small OCR-coordinate errors must not determine the family.
      for (const dx of [-1, 0, 1]) for (const dy of [-1, 0, 1]) {
        let intersection = 0, union = 0;
        for (let row = 0; row < targetH; row++) for (let col = 0; col < targetW; col++) {
          const a = sourceMask[row * targetW + col] >= 128;
          const sx = col + dx, sy = row + dy;
          const b = sx >= 0 && sx < targetW && sy >= 0 && sy < targetH && rendered[sy * targetW + sx] >= 128;
          if (a && b) intersection++; if (a || b) union++;
        }
        score = Math.max(score, intersection / Math.max(1, union));
      }
      // Long Japanese lines can have different proportional spacing despite the
      // correct family. Ink density still identifies their stroke weight reliably.
      if (!bestShape || score > bestShape.score) bestShape = { score, face, bold };
      score = score * .3 + Math.min(sourceInk, candidateInk) / Math.max(1, sourceInk, candidateInk) * .7;
      candidates.push({ score, face, bold });
    }
    best = bestShape && bestShape.score >= .65 ? bestShape
      : candidates.filter(c => !box.bold || c.bold).sort((a, b) => b.score - a.score)[0];
    if (best) report?.({ text: box.text, ...best });
    output.push(best && best.score >= .45 ? { ...box, fontFace: best.face, bold: best.bold } : box);
  }
  return output;
}
