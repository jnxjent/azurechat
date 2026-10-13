import { GlobalFonts } from "@napi-rs/canvas";
import { existsSync } from "node:fs";
import path from "node:path";

/** Use the bundled Japanese font on Linux as well as installed Windows fonts. */
export function registerSourceFonts(): void {
  if (GlobalFonts.families.some(f => f.family === "Noto Sans JP")) return;
  for (const root of [process.cwd(), path.resolve(process.cwd(), "..")]) {
    const file = path.join(root, "public", "fonts", "NotoSansJP-Regular.ttf");
    if (existsSync(file) && GlobalFonts.registerFromPath(file, "Noto Sans JP")) return;
  }
}

/** Measure real Japanese glyphs when the requested Office font is unavailable. */
export function sourceMeasurementFont(face: string): string {
  registerSourceFonts();
  const installed = new Set(GlobalFonts.families.map(f => f.family));
  return installed.has(face) ? face : installed.has("Noto Sans JP") ? "Noto Sans JP" : face;
}

/** Supported Office fonts, selected from the source image's broad font family. */
export function sourceFontFace(value: unknown): string {
  const faces: Record<string, string> = { gothic: "Meiryo", mincho: "Yu Mincho", rounded: "HGMaruGothicMPRO", sans: "Arial", serif: "Times New Roman" };
  const additional = ["Yu Gothic", "BIZ UDGothic", "Noto Sans JP", "Noto Serif JP"];
  return faces[String(value)] || ([...Object.values(faces), ...additional].includes(String(value)) ? String(value) : "Meiryo");
}
