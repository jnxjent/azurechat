export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import PptxGenJS from "pptxgenjs";
import {
  BlobSASPermissions,
  BlobServiceClient,
  StorageSharedKeyCredential,
  generateBlobSASQueryParameters,
} from "@azure/storage-blob";
import { uniqueId } from "@/features/common/util";
import { OpenAIDALLEInstance, OpenAIInstance } from "@/features/common/services/openai";

export type PptxColumn = { header: string; bullets: string[] };

export type PptxVisualBlock = {
  kind: "callout" | "node" | "badge" | "figure";
  role?: "primary" | "supporting" | "annotation";
  groupId?: string;
  text: string;
  x: number;
  y: number;
  w: number;
  h: number;
  emphasis?: boolean;
};

export type PptxConnector = {
  from: number;
  to: number;
  label?: string;
  style?: "arrow" | "line";
  relationshipType?: "flow" | "compare" | "annotation" | "support";
};

export type PptxSlide = {
  title: string;
  bullets: string[];
  layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram";
  tableRows?: string[][];
  columns?: PptxColumn[];
  visualBlocks?: PptxVisualBlock[];
  connectors?: PptxConnector[];
};

export type DeckPreferencesInput = {
  designInstruction?: string;
  accentColor?: string;
  fontScale?: "small" | "medium" | "large" | "xlarge";
  avoidEnglishLabels?: boolean;
  language?: "ja" | "en";
  recentDesignNotes?: string[];
};

export type GenPptxRequest = {
  title: string;
  slides: PptxSlide[];
  threadId: string;
  fontFace?: string;
  designInstruction?: string;
  deckPreferences?: DeckPreferencesInput;
};

type Palette = {
  canvas: string;
  surface: string;
  titleBg: string;
  headerBg: string;
  accentA: string;
  accentB: string;
  headerText: string;
  bodyText: string;
  mutedText: string;
  sectionBg: string;
  tableHeaderBg: string;
  tableHeaderText: string;
  tableAltBg: string;
  border: string;
};

type SlideVisualType =
  | "editorial"
  | "process"
  | "comparison"
  | "spotlight"
  | "cards"
  | "timeline"
  | "table";

type SlideVisualHint = {
  title: string;
  visualType: SlideVisualType;
  emphasis?: string;
};

type DeckDesignBrief = {
  palette: Palette;
  coverKicker: string;
  coverSubtitle: string;
  footerNote: string;
  mood: string;
  visualHints: SlideVisualHint[];
};

type GeneratedIllustration = {
  dataUri: string;
  prompt: string;
};

type Theme = {
  palette: Palette;
  fontFace: string;
  titleFontSize: number;
  bodyFontSize: number;
  smallFontSize: number;
  execMode: boolean;
  playfulMode: boolean;
  minimalMode: boolean;
  language: "ja" | "en";
  useJapaneseLabels: boolean;
};

const DEFAULT_PALETTES: Record<string, Palette> = {
  blue: { canvas: "F4F8FF", surface: "FFFFFF", titleBg: "163D77", headerBg: "214C8F", accentA: "3B82F6", accentB: "93C5FD", headerText: "FFFFFF", bodyText: "122033", mutedText: "58667A", sectionBg: "EAF2FF", tableHeaderBg: "3B82F6", tableHeaderText: "FFFFFF", tableAltBg: "EDF5FF", border: "D4E4FF" },
  red: { canvas: "FFF6F6", surface: "FFFFFF", titleBg: "7B1E2B", headerBg: "A3293B", accentA: "E24A5A", accentB: "F7B0B8", headerText: "FFFFFF", bodyText: "35141B", mutedText: "7B5960", sectionBg: "FFF0F1", tableHeaderBg: "E24A5A", tableHeaderText: "FFFFFF", tableAltBg: "FFF4F5", border: "F2D2D6" },
  green: { canvas: "F3FBF7", surface: "FFFFFF", titleBg: "184B3A", headerBg: "21644D", accentA: "35A073", accentB: "A7E1C7", headerText: "FFFFFF", bodyText: "163128", mutedText: "5B7268", sectionBg: "EAF8F0", tableHeaderBg: "35A073", tableHeaderText: "FFFFFF", tableAltBg: "F2FBF6", border: "D2ECDD" },
  gold: { canvas: "FFFBEF", surface: "FFFFFF", titleBg: "5B4212", headerBg: "7B5A18", accentA: "C6922D", accentB: "F0D99A", headerText: "FFF8E5", bodyText: "36280E", mutedText: "746344", sectionBg: "FFF6DD", tableHeaderBg: "C6922D", tableHeaderText: "FFFFFF", tableAltBg: "FFF9EA", border: "EADDAE" },
  pastel: { canvas: "FFF9FD", surface: "FFFFFF", titleBg: "8A6BBE", headerBg: "A085D6", accentA: "F39BCB", accentB: "CBB7F7", headerText: "FFFFFF", bodyText: "3B3150", mutedText: "7B7091", sectionBg: "F7F1FF", tableHeaderBg: "A085D6", tableHeaderText: "FFFFFF", tableAltBg: "FBF7FF", border: "E6DBFB" },
  pop: { canvas: "FFFDF1", surface: "FFFFFF", titleBg: "8B0D57", headerBg: "D61F69", accentA: "FF9F1C", accentB: "FFE066", headerText: "FFFFFF", bodyText: "2D2230", mutedText: "7F6570", sectionBg: "FFF4DB", tableHeaderBg: "FF9F1C", tableHeaderText: "FFFFFF", tableAltBg: "FFF9E6", border: "FFE2A8" },
};

function normalizeHex(input: string, fallback: string): string {
  const value = String(input ?? "").replace("#", "").trim();
  return /^[0-9A-Fa-f]{6}$/.test(value) ? value.toUpperCase() : fallback;
}

function truncateText(value: string, max: number): string {
  const text = String(value ?? "").trim();
  return text.length > max ? `${text.slice(0, max).trim()}...` : text;
}

function containsAny(value: string, tokens: string[]): boolean {
  return tokens.some((token) => value.includes(token));
}

function resolvePaletteKeyFromPrompt(input: string): keyof typeof DEFAULT_PALETTES {
  const hint = input.toLowerCase();
  if (containsAny(hint, ["pastel", "soft", "gentle", "sweet", "やわらか", "パステル"])) return "pastel";
  if (containsAny(hint, ["pop", "playful", "vivid", "colorful", "ポップ", "元気"])) return "pop";
  if (containsAny(hint, ["red", "crimson", "scarlet", "赤"])) return "red";
  if (containsAny(hint, ["green", "eco", "natural", "organic", "緑", "エコ"])) return "green";
  if (containsAny(hint, ["gold", "yellow", "luxury", "premium", "golden", "金", "高級"])) return "gold";
  return "blue";
}

function getFontScaleMultiplier(fontScale?: DeckPreferencesInput["fontScale"]): number {
  switch (fontScale) {
    case "small":
      return 0.92;
    case "large":
      return 1.12;
    case "xlarge":
      return 1.22;
    default:
      return 1;
  }
}

function applyAccentOverride(palette: Palette, accentColor?: string): Palette {
  const key = String(accentColor ?? "").toLowerCase();
  if (containsAny(key, ["blue", "青"])) return DEFAULT_PALETTES.blue;
  if (containsAny(key, ["red", "赤"])) return DEFAULT_PALETTES.red;
  if (containsAny(key, ["green", "eco", "緑", "エコ"])) return DEFAULT_PALETTES.green;
  if (containsAny(key, ["gold", "yellow", "金"])) return DEFAULT_PALETTES.gold;
  return palette;
}

function getFontFace(input: string, explicitFontFace?: string): string {
  if (explicitFontFace?.trim()) return explicitFontFace.trim();
  const hint = input.toLowerCase();
  if (containsAny(hint, ["yu gothic", "游ゴシック"])) return "Yu Gothic";
  if (containsAny(hint, ["yu mincho", "游明朝", "明朝"])) return "Yu Mincho";
  if (containsAny(hint, ["arial"])) return "Arial";
  return "Meiryo";
}

function localizeLabel(
  key: "coverKicker" | "summary" | "points" | "details" | "pages",
  theme: Theme
): string {
  if (!theme.useJapaneseLabels) {
    if (key === "coverKicker") return "AI-ENHANCED PRESENTATION";
    if (key === "summary") return "SUMMARY";
    if (key === "points") return "POINTS";
    if (key === "details") return "DETAILS";
    return "pages";
  }
  if (key === "coverKicker") return "AI生成プレゼンテーション";
  if (key === "summary") return "要約";
  if (key === "points") return "要点";
  if (key === "details") return "詳細";
  return "ページ";
}

function createFallbackBrief(
  title: string,
  slides: PptxSlide[],
  instructionText: string,
  prefs?: DeckPreferencesInput
): DeckDesignBrief {
  const paletteKey = resolvePaletteKeyFromPrompt(instructionText);
  const palette = applyAccentOverride(DEFAULT_PALETTES[paletteKey], prefs?.accentColor);
  const visualCycle: SlideVisualType[] = ["spotlight", "cards", "process", "comparison", "editorial", "timeline"];
  return {
    palette,
    coverKicker: containsAny(instructionText.toLowerCase(), ["proposal", "提案", "企画"]) ? "PROPOSAL DECK" : "AI-ENHANCED PRESENTATION",
    coverSubtitle: instructionText ? truncateText(instructionText, 68) : "Structured from the uploaded source document",
    footerNote: containsAny(instructionText.toLowerCase(), ["executive", "役員", "board"]) ? "Executive-ready summary deck" : "Auto-generated document summary",
    mood: instructionText || "editorial",
    visualHints: slides.map((slide, index) => ({
      title: slide.title,
      visualType:
        slide.layoutType === "table"
          ? "table"
          : slide.layoutType === "multi-column"
            ? "comparison"
            : slide.layoutType === "diagram"
              ? "process"
              : visualCycle[index % visualCycle.length],
      emphasis: slide.bullets[0] ?? slide.title,
    })),
  };
}

async function generateDesignBrief(
  title: string,
  slides: PptxSlide[],
  instructionText: string,
  prefs?: DeckPreferencesInput
): Promise<DeckDesignBrief> {
  const fallback = createFallbackBrief(title, slides, instructionText, prefs);
  try {
    const openai = OpenAIInstance();
    const slideOutline = slides.slice(0, 12).map((slide, index) => ({
      index: index + 1,
      title: slide.title,
      bullets: slide.bullets.slice(0, 4),
      layoutType: slide.layoutType ?? "bullets",
      hasTable: (slide.tableRows?.length ?? 0) > 0,
      hasColumns: (slide.columns?.length ?? 0) > 1,
    }));

    const res = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
      messages: [
        {
          role: "system",
          content:
            "You are a presentation art director. Return JSON only. Prioritize the user's natural-language request over generic defaults. Explicit corrections such as color changes, font changes, Japanese labels, executive tone, pastel, pop, soft, premium, minimal, eco, or bold should be reflected directly in the visual system. Suggested visualType values: editorial, process, comparison, spotlight, cards, timeline, table.",
        },
        {
          role: "user",
          content: JSON.stringify({
            task: "Create a design brief for auto-generating a PowerPoint.",
            title,
            instructionText,
            deckPreferences: prefs ?? {},
            slideOutline,
          }),
        },
      ],
      response_format: { type: "json_object" },
      max_completion_tokens: 900,
    });

    const parsed = JSON.parse(res.choices[0]?.message?.content ?? "{}");
    const palette: Palette = {
      canvas: normalizeHex(parsed?.palette?.canvas, fallback.palette.canvas),
      surface: normalizeHex(parsed?.palette?.surface, fallback.palette.surface),
      titleBg: normalizeHex(parsed?.palette?.titleBg, fallback.palette.titleBg),
      headerBg: normalizeHex(parsed?.palette?.headerBg, fallback.palette.headerBg),
      accentA: normalizeHex(parsed?.palette?.accentA, fallback.palette.accentA),
      accentB: normalizeHex(parsed?.palette?.accentB, fallback.palette.accentB),
      headerText: normalizeHex(parsed?.palette?.headerText, fallback.palette.headerText),
      bodyText: normalizeHex(parsed?.palette?.bodyText, fallback.palette.bodyText),
      mutedText: normalizeHex(parsed?.palette?.mutedText, fallback.palette.mutedText),
      sectionBg: normalizeHex(parsed?.palette?.sectionBg, fallback.palette.sectionBg),
      tableHeaderBg: normalizeHex(parsed?.palette?.tableHeaderBg, fallback.palette.tableHeaderBg),
      tableHeaderText: normalizeHex(parsed?.palette?.tableHeaderText, fallback.palette.tableHeaderText),
      tableAltBg: normalizeHex(parsed?.palette?.tableAltBg, fallback.palette.tableAltBg),
      border: normalizeHex(parsed?.palette?.border, fallback.palette.border),
    };

    const visualHints: SlideVisualHint[] = slides.map((slide, index) => {
      const matched = Array.isArray(parsed?.visualHints) && parsed.visualHints.find((item: any) => item?.title === slide.title);
      const visualType = String(matched?.visualType ?? "").trim();
      const safeVisual: SlideVisualType =
        visualType === "editorial" || visualType === "process" || visualType === "comparison" ||
        visualType === "spotlight" || visualType === "cards" || visualType === "timeline" ||
        visualType === "table"
          ? visualType
          : fallback.visualHints[index]?.visualType ?? "editorial";

      return {
        title: slide.title,
        visualType:
          slide.layoutType === "table"
            ? "table"
            : slide.layoutType === "multi-column"
              ? "comparison"
              : slide.layoutType === "diagram"
                ? "process"
                : safeVisual,
        emphasis: String(matched?.emphasis ?? "").trim() || fallback.visualHints[index]?.emphasis || slide.bullets[0] || slide.title,
      };
    });

    return {
      palette: applyAccentOverride(palette, prefs?.accentColor),
      coverKicker: String(parsed?.coverKicker ?? "").trim() || fallback.coverKicker,
      coverSubtitle: String(parsed?.coverSubtitle ?? "").trim() || fallback.coverSubtitle,
      footerNote: String(parsed?.footerNote ?? "").trim() || fallback.footerNote,
      mood: String(parsed?.mood ?? "").trim() || fallback.mood,
      visualHints,
    };
  } catch (error) {
    console.warn("[gen-pptx] design brief fallback:", error);
    return fallback;
  }
}

function resolveTheme(
  designBrief: DeckDesignBrief,
  instructionText: string,
  prefs?: DeckPreferencesInput,
  explicitFontFace?: string
): Theme {
  const palette = applyAccentOverride(designBrief.palette, prefs?.accentColor);
  const lowered = instructionText.toLowerCase();
  const execMode = containsAny(lowered, ["executive", "board", "役員", "経営", "提案書"]);
  const playfulMode = containsAny(lowered, ["pop", "playful", "fun", "cute", "ポップ", "親しみ", "やわらか", "パステル"]);
  const minimalMode = containsAny(lowered, ["minimal", "simple", "clean", "quiet", "ミニマル", "シンプル"]);
  const scale = getFontScaleMultiplier(prefs?.fontScale ?? (execMode ? "large" : "medium"));
  const useJapaneseLabels = prefs?.language === "ja" || prefs?.avoidEnglishLabels === true;
  return {
    palette,
    fontFace: getFontFace(instructionText, explicitFontFace),
    titleFontSize: Math.round(24 * scale),
    bodyFontSize: Math.round(17 * scale),
    smallFontSize: Math.round(11 * scale),
    execMode,
    playfulMode,
    minimalMode,
    language: prefs?.language === "en" ? "en" : "ja",
    useJapaneseLabels,
  };
}

async function uploadPptxToBlob(buffer: Buffer, fileName: string): Promise<string> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const containerName = "pptx";
  const sharedKeyCredential = new StorageSharedKeyCredential(acc, key);
  const blobServiceClient = BlobServiceClient.fromConnectionString(
    `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
  const containerClient = blobServiceClient.getContainerClient(containerName);
  await containerClient.createIfNotExists({ access: "blob" });
  const blockBlobClient = containerClient.getBlockBlobClient(fileName);
  await blockBlobClient.uploadData(buffer, {
    blobHTTPHeaders: {
      blobContentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      blobContentDisposition: `attachment; filename="${fileName}"`,
    },
  });
  const sasToken = generateBlobSASQueryParameters(
    { containerName, blobName: fileName, expiresOn: new Date(Date.now() + 24 * 60 * 60 * 1000), permissions: BlobSASPermissions.parse("r") },
    sharedKeyCredential
  );
  return `${blockBlobClient.url}?${sasToken}`;
}

const W = 13.33;
const H = 7.5;
const HEADER_H = 1.05;

function addHeaderBand(s: PptxGenJS.Slide, title: string, theme: Theme) {
  s.addShape("rect", {
    x: 0,
    y: 0,
    w: W,
    h: HEADER_H,
    fill: { color: theme.palette.headerBg },
    line: { color: theme.palette.headerBg, width: 0 },
  });
  s.addShape("rect", {
    x: 0,
    y: HEADER_H,
    w: W,
    h: 0.06,
    fill: { color: theme.palette.accentA },
    line: { color: theme.palette.accentA, width: 0 },
  });
  s.addText(title, {
    x: 0.45,
    y: 0.1,
    w: W - 0.9,
    h: HEADER_H - 0.18,
    fontSize: theme.titleFontSize,
    fontFace: theme.fontFace,
    bold: true,
    color: theme.palette.headerText,
    valign: "middle",
  });
}

function addChrome(s: PptxGenJS.Slide, theme: Theme) {
  s.addShape("rect", {
    x: 0,
    y: H - 0.08,
    w: W,
    h: 0.08,
    fill: { color: theme.palette.accentA },
    line: { color: theme.palette.accentA, width: 0 },
  });
  s.addShape("rect", {
    x: W - 0.1,
    y: HEADER_H + 0.06,
    w: 0.1,
    h: H - HEADER_H - 0.14,
    fill: { color: theme.palette.accentB, transparency: 55 },
    line: { color: theme.palette.accentB, width: 0 },
  });
}

function clampPercent(value: number): number {
  return Math.max(0, Math.min(100, value));
}

function toDiagramX(value: number): number {
  return 0.55 + ((W - 1.2) * clampPercent(value)) / 100;
}

function toDiagramY(value: number): number {
  return HEADER_H + 0.25 + ((H - HEADER_H - 0.6) * clampPercent(value)) / 100;
}

function toDiagramW(value: number): number {
  return ((W - 1.2) * Math.max(6, Math.min(100, value))) / 100;
}

function toDiagramH(value: number): number {
  return ((H - HEADER_H - 0.6) * Math.max(6, Math.min(100, value))) / 100;
}

function buildDiagramRowBands(s: PptxGenJS.Slide, blocks: PptxVisualBlock[], theme: Theme) {
  if (blocks.length < 3) return;
  const sorted = [...blocks].sort((a, b) => a.y - b.y);
  const groups: PptxVisualBlock[][] = [];
  for (const block of sorted) {
    const last = groups[groups.length - 1];
    if (!last) {
      groups.push([block]);
      continue;
    }
    const centerY = last.reduce((sum, item) => sum + item.y + item.h / 2, 0) / last.length;
    const blockCenter = block.y + block.h / 2;
    if (Math.abs(blockCenter - centerY) <= 12) {
      last.push(block);
    } else {
      groups.push([block]);
    }
  }

  groups
    .filter((group) => group.length >= 2)
    .slice(0, 4)
    .forEach((group, index) => {
      const minY = Math.min(...group.map((item) => item.y));
      const maxY = Math.max(...group.map((item) => item.y + item.h));
      s.addShape("roundRect", {
        x: 0.62,
        y: Math.max(HEADER_H + 0.26, toDiagramY(minY) - 0.1),
        w: 9.15,
        h: Math.min(H - HEADER_H - 0.7, toDiagramY(maxY) - toDiagramY(minY) + 0.22),
        rectRadius: 0.04,
        fill: { color: index % 2 === 0 ? theme.palette.sectionBg : theme.palette.tableAltBg, transparency: 55 },
        line: { color: theme.palette.border, transparency: 75, width: 0.7 },
      });
    });
}

function buildDiagramGroupCards(s: PptxGenJS.Slide, blocks: PptxVisualBlock[], theme: Theme) {
  const grouped = new Map<string, PptxVisualBlock[]>();
  for (const block of blocks) {
    if (!block.groupId) continue;
    const list = grouped.get(block.groupId) ?? [];
    list.push(block);
    grouped.set(block.groupId, list);
  }

  Array.from(grouped.values())
    .filter((group) => group.length >= 2)
    .slice(0, 5)
    .forEach((group, index) => {
      const minX = Math.min(...group.map((item) => item.x));
      const minY = Math.min(...group.map((item) => item.y));
      const maxX = Math.max(...group.map((item) => item.x + item.w));
      const maxY = Math.max(...group.map((item) => item.y + item.h));
      s.addShape("roundRect", {
        x: Math.max(0.55, toDiagramX(minX) - 0.12),
        y: Math.max(HEADER_H + 0.25, toDiagramY(minY) - 0.1),
        w: Math.min(9.0, toDiagramX(maxX) - toDiagramX(minX) + 0.24),
        h: Math.min(H - HEADER_H - 0.75, toDiagramY(maxY) - toDiagramY(minY) + 0.22),
        rectRadius: 0.04,
        fill: {
          color: index % 2 === 0 ? theme.palette.tableAltBg : theme.palette.sectionBg,
          transparency: 70,
        },
        line: { color: theme.palette.border, transparency: 82, width: 0.8 },
      });
    });
}

function drawDiagramConnector(
  s: PptxGenJS.Slide,
  from: PptxVisualBlock,
  to: PptxVisualBlock,
  connector: PptxConnector,
  theme: Theme
) {
  const fromX = toDiagramX(from.x) + toDiagramW(from.w) / 2;
  const fromY = toDiagramY(from.y) + toDiagramH(from.h) / 2;
  const toX = toDiagramX(to.x) + toDiagramW(to.w) / 2;
  const toY = toDiagramY(to.y) + toDiagramH(to.h) / 2;
  const midX = fromX + (toX - fromX) / 2;
  const midY = fromY + (toY - fromY) / 2;
  const relationship = connector.relationshipType ?? "flow";
  const lineColor =
    relationship === "compare"
      ? theme.palette.accentB
      : relationship === "annotation"
        ? theme.palette.mutedText
        : relationship === "support"
          ? theme.palette.border
          : theme.palette.accentA;
  const lineWidth =
    relationship === "annotation" ? 1 : relationship === "support" ? 0.9 : relationship === "compare" ? 1.6 : 1.4;
  const endArrow =
    relationship === "compare"
      ? "stealth"
      : connector.style === "line" || relationship === "support"
        ? "none"
        : "triangle";

  s.addShape("line", {
    x: fromX,
    y: fromY,
    w: toX - fromX,
    h: toY - fromY,
    line: {
      color: lineColor,
      width: lineWidth,
      transparency: relationship === "annotation" ? 18 : 0,
      beginArrowType: "none",
      endArrowType: endArrow,
    },
  });

  s.addShape("ellipse", {
    x: fromX - 0.04,
    y: fromY - 0.04,
    w: 0.08,
    h: 0.08,
    fill: { color: lineColor },
    line: { color: lineColor, width: 0 },
  });

  if (connector.label) {
    s.addText(connector.label, {
      x: midX - 0.72,
      y: midY - 0.16,
      w: 1.2,
      h: 0.25,
      fontSize: 9,
      fontFace: theme.fontFace,
      align: "center",
      color: relationship === "annotation" ? theme.palette.mutedText : theme.palette.bodyText,
      margin: 0,
      fill: { color: theme.palette.canvas, transparency: 8 },
    });
  }
}

function buildDiagramSlide(
  pptx: PptxGenJS,
  slide: PptxSlide,
  theme: Theme,
  visual: SlideVisualHint
) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  addVisualAccent(s, visual, theme);

  const blocks = slide.visualBlocks ?? [];
  const connectors = slide.connectors ?? [];
  const primaryBlock =
    blocks.find((block) => block.role === "primary") ??
    blocks.find((block) => block.emphasis) ??
    [...blocks].sort((a, b) => b.w * b.h - a.w * a.h)[0] ??
    null;
  const validConnectors = connectors.filter(
    (conn) => blocks[conn.from] && blocks[conn.to] && conn.from !== conn.to
  );

  s.addShape("roundRect", {
    x: 0.42,
    y: HEADER_H + 0.18,
    w: 9.55,
    h: H - HEADER_H - 0.48,
    rectRadius: 0.05,
    fill: { color: theme.palette.surface, transparency: 2 },
    line: { color: theme.palette.border, width: 0.9 },
  });

  buildDiagramGroupCards(s, blocks, theme);
  buildDiagramRowBands(s, blocks, theme);

  s.addShape("roundRect", {
    x: 10.12,
    y: HEADER_H + 0.18,
    w: 2.78,
    h: H - HEADER_H - 0.48,
    rectRadius: 0.05,
    fill: { color: theme.palette.sectionBg },
    line: { color: theme.palette.border, width: 0.9 },
  });

  s.addText(localizeLabel("summary", theme), {
    x: 10.34,
    y: HEADER_H + 0.38,
    w: 2.2,
    h: 0.24,
    fontSize: theme.smallFontSize,
    fontFace: theme.fontFace,
    bold: true,
    color: theme.palette.mutedText,
  });

  if (primaryBlock) {
    s.addShape("roundRect", {
      x: 10.34,
      y: HEADER_H + 0.76,
      w: 2.18,
      h: 0.82,
      rectRadius: 0.05,
      fill: { color: theme.palette.accentA, transparency: theme.execMode ? 8 : 0 },
      line: { color: theme.palette.accentA, width: 0 },
    });
    s.addText(primaryBlock.text, {
      x: 10.52,
      y: HEADER_H + 0.93,
      w: 1.82,
      h: 0.48,
      fontSize: Math.max(theme.bodyFontSize - 1, 12),
      fontFace: theme.fontFace,
      bold: true,
      color: theme.palette.headerText,
      align: "center",
      valign: "middle",
      fit: "shrink",
      margin: 0.04,
    });
  }

  const summaryItems = slide.bullets.slice(0, 4);
  summaryItems.forEach((item, index) => {
    const y = HEADER_H + 1.8 + index * 0.88;
    s.addShape("roundRect", {
      x: 10.34,
      y,
      w: 2.18,
      h: 0.68,
      rectRadius: 0.04,
      fill: { color: "FFFFFF", transparency: 6 },
      line: { color: theme.palette.border, width: 0.8 },
    });
    s.addText(item, {
      x: 10.5,
      y: y + 0.1,
      w: 1.86,
      h: 0.46,
      fontSize: Math.max(theme.smallFontSize - 1, 9),
      fontFace: theme.fontFace,
      color: theme.palette.bodyText,
      valign: "middle",
      fit: "shrink",
      margin: 0.02,
    });
  });

  for (const connector of validConnectors) {
    drawDiagramConnector(s, blocks[connector.from], blocks[connector.to], connector, theme);
  }

  blocks.forEach((block) => {
    const x = toDiagramX(block.x);
    const y = toDiagramY(block.y);
    const w = toDiagramW(block.w);
    const h = toDiagramH(block.h);
    const isBadge = block.kind === "badge";
    const isCallout = block.kind === "callout";
    const isFigure = block.kind === "figure";
    const isAnnotation = block.role === "annotation";
    const isSupporting = block.role === "supporting";
    const isPrimary =
      primaryBlock !== null &&
      block.text === primaryBlock.text &&
      block.x === primaryBlock.x &&
      block.y === primaryBlock.y;
    const fillColor = block.emphasis
      ? theme.palette.accentA
      : isBadge
        ? theme.palette.sectionBg
        : isFigure
          ? theme.palette.tableAltBg
          : isAnnotation
            ? theme.palette.sectionBg
          : theme.palette.surface;
    const textColor = block.emphasis ? theme.palette.headerText : theme.palette.bodyText;

    s.addShape(isCallout ? "roundRect" : isBadge ? "hexagon" : "roundRect", {
      x,
      y,
      w,
      h,
      rectRadius: isCallout ? 0.06 : 0.03,
      fill: { color: fillColor, transparency: isAnnotation ? 8 : 0 },
      line: { color: block.emphasis || isPrimary ? theme.palette.accentA : theme.palette.border, width: isPrimary ? 2 : 1.2 },
      shadow: { type: "outer", color: "AAB7C4", blur: isPrimary ? 2 : 1, angle: 45, opacity: isAnnotation ? 0.08 : isPrimary ? 0.18 : 0.12 },
    });

    if (isCallout) {
      s.addShape("chevron", {
        x: x + Math.max(w / 2 - 0.14, 0.06),
        y: y + h - 0.05,
        w: 0.28,
        h: 0.22,
        rotate: 90,
        fill: { color: fillColor },
        line: { color: block.emphasis || isPrimary ? theme.palette.accentA : theme.palette.border, width: 1 },
      });
    }

    if (isPrimary || isFigure) {
      s.addShape("rect", {
        x: x + 0.08,
        y: y + 0.08,
        w: w - 0.16,
        h: 0.1,
        fill: { color: theme.palette.accentB, transparency: isPrimary ? 0 : 15 },
        line: { color: theme.palette.accentB, width: 0 },
      });
    }

    s.addText(block.text, {
      x: x + 0.12,
      y: y + (isPrimary || isFigure ? 0.2 : 0.08),
      w: w - 0.24,
      h: h - (isPrimary || isFigure ? 0.28 : 0.16),
      fontSize: Math.max(theme.bodyFontSize - (isBadge ? 4 : isPrimary ? -1 : isSupporting ? 1 : 2), 10),
      fontFace: theme.fontFace,
      bold: block.emphasis || isBadge || isPrimary,
      color: textColor,
      align: "center",
      valign: "middle",
      margin: 0.05,
      breakLine: false,
      fit: "shrink",
    });
  });

  addChrome(s, theme);
}

function addAmbientShapes(s: PptxGenJS.Slide, theme: Theme) {
  if (theme.minimalMode) return;
  s.addShape("ellipse", {
    x: W - 1.9,
    y: -0.38,
    w: 2.3,
    h: 2.3,
    fill: { color: theme.palette.accentA, transparency: 78 },
    line: { color: theme.palette.accentA, transparency: 35, width: 1.5 },
  });
  s.addShape("ellipse", {
    x: 0.5,
    y: H - 1.55,
    w: 1.05,
    h: 1.05,
    fill: { color: theme.palette.accentB, transparency: 72 },
    line: { color: theme.palette.accentB, transparency: 45, width: 1 },
  });
  if (theme.playfulMode && !theme.execMode) {
    s.addShape("ellipse", {
      x: W - 3.2,
      y: H - 1.4,
      w: 0.72,
      h: 0.72,
      fill: { color: theme.palette.accentA, transparency: 58 },
      line: { color: theme.palette.accentA, transparency: 25, width: 1 },
    });
    s.addShape("ellipse", {
      x: W - 2.55,
      y: H - 1.12,
      w: 0.34,
      h: 0.34,
      fill: { color: theme.palette.accentB, transparency: 42 },
      line: { color: theme.palette.accentB, transparency: 20, width: 0.8 },
    });
  }
}

function addVisualAccent(s: PptxGenJS.Slide, visual: SlideVisualHint, theme: Theme) {
  const emphasis = truncateText(visual.emphasis || "", 48);
  const x = 9.3;
  const y = 1.45;
  if (!emphasis) return;
  if (theme.execMode && visual.visualType === "editorial") return;

  s.addShape("roundRect", {
    x,
    y,
    w: 2.7,
    h: 1.05,
    rectRadius: 0.06,
    fill: { color: theme.palette.sectionBg },
    line: { color: theme.palette.border, width: 1 },
  });

  if (visual.visualType === "timeline") {
    s.addShape("line", {
      x: x + 0.22,
      y: y + 0.58,
      w: 1.95,
      h: 0,
      line: { color: theme.palette.accentA, width: 1.4 },
    });
    [0, 0.88, 1.76].forEach((offset) => {
      s.addShape("ellipse", {
        x: x + 0.18 + offset,
        y: y + 0.51,
        w: 0.13,
        h: 0.13,
        fill: { color: theme.palette.accentA },
        line: { color: theme.palette.accentA, width: 0 },
      });
    });
  } else if (visual.visualType === "comparison") {
    s.addShape("rect", {
      x: x + 0.16,
      y: y + 0.34,
      w: 0.78,
      h: 0.34,
      fill: { color: theme.palette.accentA },
      line: { color: theme.palette.accentA, width: 0 },
    });
    s.addShape("rect", {
      x: x + 1.12,
      y: y + 0.34,
      w: 0.98,
      h: 0.34,
      fill: { color: theme.palette.accentB },
      line: { color: theme.palette.accentB, width: 0 },
    });
  } else if (visual.visualType === "process") {
    [0, 0.62, 1.24].forEach((offset) => {
      s.addShape("roundRect", {
        x: x + 0.18 + offset,
        y: y + 0.34,
        w: 0.38,
        h: 0.22,
        rectRadius: 0.06,
        fill: { color: theme.palette.accentA },
        line: { color: theme.palette.accentA, width: 0 },
      });
    });
  }

  s.addText(emphasis, {
    x: x + 0.18,
    y: y + 0.74,
    w: 2.25,
    h: 0.16,
    fontFace: theme.fontFace,
    fontSize: theme.smallFontSize - 1,
    color: theme.palette.bodyText,
    italic: visual.visualType === "editorial",
    fit: "shrink",
  });
}

function shouldGenerateIllustration(instructionText: string, theme: Theme): boolean {
  const lowered = instructionText.toLowerCase();
  if (containsAny(lowered, ["イラスト不要", "画像不要", "no illustration", "no image"])) return false;
  if (theme.execMode && !containsAny(lowered, ["イラスト", "挿絵", "絵", "robot", "ロボット", "image", "画像"])) {
    return false;
  }
  return containsAny(lowered, [
    "illustration",
    "illustrated",
    "image",
    "insert image",
    "robot",
    "eco mascot",
    "イラスト",
    "挿絵",
    "絵",
    "画像",
    "ロボット",
    "マスコット",
  ]);
}

async function generateCoverIllustration(
  title: string,
  instructionText: string,
  theme: Theme
): Promise<GeneratedIllustration | null> {
  if (!shouldGenerateIllustration(instructionText, theme)) return null;
  try {
    const openai = OpenAIDALLEInstance();
    const prompt = [
      "Create a clean presentation illustration for a business proposal cover.",
      `Topic: ${title}.`,
      `Design direction: ${instructionText || "professional and polished"}.`,
      "Style: modern editorial illustration, no text, no watermark, transparent or plain clean background, suitable for a PowerPoint cover.",
      theme.execMode
        ? "Tone: trustworthy, restrained, executive-ready, not cartoonish."
        : "Tone: approachable, polished, optimistic, presentation-friendly.",
    ].join(" ");
    const response = await openai.images.generate({
      model: "gpt-image-1.5",
      prompt,
    });
    const b64 = response.data?.[0]?.b64_json;
    if (!b64) return null;
    return {
      dataUri: `data:image/png;base64,${b64}`,
      prompt,
    };
  } catch (error) {
    console.warn("[gen-pptx] cover illustration skipped:", error);
    return null;
  }
}

function buildTitleSlide(
  pptx: PptxGenJS,
  title: string,
  brief: DeckDesignBrief,
  theme: Theme,
  slideCount: number,
  illustration?: GeneratedIllustration | null
) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.titleBg };
  addAmbientShapes(s, theme);
  s.addShape("rect", {
    x: 0,
    y: 0,
    w: 0.28,
    h: H,
    fill: { color: theme.palette.accentA },
    line: { color: theme.palette.accentA, width: 0 },
  });
  s.addShape("rect", {
    x: 0,
    y: H - 0.12,
    w: W,
    h: 0.12,
    fill: { color: theme.palette.accentB },
    line: { color: theme.palette.accentB, width: 0 },
  });
  s.addText(localizeLabel("coverKicker", theme), {
    x: 0.72,
    y: 1.05,
    w: 3.2,
    h: 0.24,
    fontSize: theme.smallFontSize,
    fontFace: theme.fontFace,
    bold: true,
    color: theme.palette.headerText,
  });
  s.addText(title, {
    x: 0.72,
    y: 2.05,
    w: W - 1.45,
    h: 1.9,
    fontSize: theme.titleFontSize + 8,
    fontFace: theme.fontFace,
    bold: true,
    color: theme.palette.headerText,
    align: "center",
    valign: "middle",
    fit: "shrink",
  });
  s.addText(brief.coverSubtitle, {
    x: 1.1,
    y: 4.45,
    w: W - 2.2,
    h: 0.42,
    fontSize: theme.bodyFontSize - 1,
    fontFace: theme.fontFace,
    color: theme.palette.headerText,
    align: "center",
    fit: "shrink",
  });
  if (illustration?.dataUri) {
    s.addShape("roundRect", {
      x: 9.55,
      y: 1.18,
      w: 2.75,
      h: 2.75,
      rectRadius: 0.08,
      fill: { color: theme.palette.surface, transparency: 4 },
      line: { color: theme.palette.accentB, transparency: 35, width: 1.2 },
      shadow: { type: "outer", color: "1B1B1B", blur: 2, angle: 45, opacity: 0.18 },
    });
    s.addImage({
      data: illustration.dataUri,
      x: 9.72,
      y: 1.33,
      w: 2.42,
      h: 2.42,
      sizing: { type: "contain", x: 9.72, y: 1.33, w: 2.42, h: 2.42 },
    });
  }
  s.addText(`${slideCount} ${localizeLabel("pages", theme)}`, {
    x: W - 2.2,
    y: 6.65,
    w: 1.7,
    h: 0.2,
    fontSize: theme.smallFontSize,
    fontFace: theme.fontFace,
    bold: true,
    align: "right",
    color: theme.palette.headerText,
  });
}

type Section = { header: string | null; items: string[] };

function parseSections(bullets: string[]): { hasSections: boolean; sections: Section[] } {
  const sections: Section[] = [];
  let current: Section = { header: null, items: [] };
  for (const bullet of bullets) {
    if (bullet.startsWith("[H]")) {
      if (current.header !== null || current.items.length > 0) sections.push(current);
      current = { header: bullet.slice(3).trim(), items: [] };
    } else {
      current.items.push(bullet);
    }
  }
  if (current.header !== null || current.items.length > 0) sections.push(current);
  return { hasSections: sections.some((section) => section.header !== null), sections };
}

function buildSectionSlide(pptx: PptxGenJS, title: string, theme: Theme) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  s.addShape("rect", {
    x: 0,
    y: 2.82,
    w: W,
    h: 1.45,
    fill: { color: theme.palette.headerBg },
    line: { color: theme.palette.headerBg, width: 0 },
  });
  s.addShape("rect", {
    x: 0,
    y: 4.27,
    w: W,
    h: 0.06,
    fill: { color: theme.palette.accentA },
    line: { color: theme.palette.accentA, width: 0 },
  });
  s.addText(title, {
    x: 0.6,
    y: 2.93,
    w: W - 1.2,
    h: 1.18,
    fontSize: theme.titleFontSize + 4,
    fontFace: theme.fontFace,
    bold: true,
    color: theme.palette.headerText,
    align: "center",
    valign: "middle",
  });
  addChrome(s, theme);
}

function buildBulletsSlide(
  pptx: PptxGenJS,
  slide: PptxSlide,
  theme: Theme,
  visual: SlideVisualHint,
  illustration?: GeneratedIllustration | null
) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  addVisualAccent(s, visual, theme);
  const showIllustration = Boolean(illustration?.dataUri) && !theme.execMode;
  const textWidth = showIllustration ? 6.85 : 8.15;

  const { hasSections, sections } = parseSections(slide.bullets);
  if (theme.execMode) {
    const allBullets = sections.flatMap((section) => section.items);
    const keyMessage = allBullets[0] ?? slide.bullets[0] ?? "";
    const rest = allBullets.slice(1, 4);

    if (keyMessage) {
      s.addShape("roundRect", {
        x: 0.55,
        y: 1.45,
        w: 8.15,
        h: 1.0,
        rectRadius: 0.04,
        fill: { color: theme.palette.sectionBg },
        line: { color: theme.palette.accentA, width: 1.5 },
      });
      s.addText(keyMessage, {
        x: 0.75,
        y: 1.63,
        w: 7.75,
        h: 0.65,
        fontSize: theme.bodyFontSize + 2,
        fontFace: theme.fontFace,
        bold: true,
        color: theme.palette.bodyText,
        valign: "middle",
      });
    }

    if (rest.length > 0) {
      const bulletItems = rest.map((item) => ({
        text: item,
        options: {
          bullet: { indent: 14 },
          breakLine: true,
          fontSize: theme.bodyFontSize,
          fontFace: theme.fontFace,
          color: theme.palette.bodyText,
          paraSpaceAfter: 10,
        },
      }));
      s.addText(bulletItems, {
        x: 0.55,
        y: 2.72,
        w: 8.15,
        h: H - 3.05,
        margin: 0.1,
        valign: "top",
      });
    }
      addChrome(s, theme);
      return;
  }

  if (showIllustration) {
    s.addShape("roundRect", {
      x: 9.05,
      y: 1.72,
      w: 3.25,
      h: 3.25,
      rectRadius: 0.08,
      fill: { color: theme.palette.surface },
      line: { color: theme.palette.border, width: 1 },
    });
    s.addImage({
      data: illustration!.dataUri,
      x: 9.22,
      y: 1.9,
      w: 2.92,
      h: 2.92,
      sizing: { type: "contain", x: 9.22, y: 1.9, w: 2.92, h: 2.92 },
    });
  }

  if (hasSections) {
    let currentY = HEADER_H + 0.24;
    const maxY = H - 0.28;
    for (const section of sections) {
      if (section.header) {
        if (currentY + 0.42 > maxY) break;
        s.addShape("rect", {
          x: 0.42,
          y: currentY,
          w: 8.35,
          h: 0.42,
          fill: { color: theme.palette.sectionBg },
          line: { color: theme.palette.border, width: 0.8 },
        });
        s.addText(section.header, {
          x: 0.56,
          y: currentY + 0.04,
          w: textWidth - 0.15,
          h: 0.28,
          fontSize: theme.bodyFontSize - 1,
          fontFace: theme.fontFace,
          bold: true,
          color: theme.palette.bodyText,
          valign: "middle",
        });
        currentY += 0.5;
      }
      for (const item of section.items) {
        if (currentY + 0.31 > maxY) break;
        s.addText(`• ${item}`, {
          x: 0.72,
          y: currentY,
          w: textWidth - 0.3,
          h: 0.28,
          fontSize: theme.bodyFontSize - 1,
          fontFace: theme.fontFace,
          color: theme.palette.bodyText,
          valign: "middle",
        });
        currentY += 0.33;
      }
      currentY += 0.08;
    }
  } else if (slide.bullets.length > 0) {
    const bulletItems = slide.bullets.map((item) => ({
      text: item,
      options: {
        bullet: { indent: 14 },
        breakLine: true,
        fontSize: theme.bodyFontSize,
        fontFace: theme.fontFace,
        color: theme.palette.bodyText,
        paraSpaceAfter: 9,
      },
    }));
    s.addText(bulletItems, {
      x: 0.55,
      y: HEADER_H + 0.24,
      w: textWidth,
      h: H - HEADER_H - 0.55,
      margin: 0.1,
      valign: "top",
    });
  }

  addChrome(s, theme);
}

function buildTableSlide(
  pptx: PptxGenJS,
  slide: PptxSlide,
  theme: Theme,
  visual: SlideVisualHint
) {
  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  addVisualAccent(s, visual, theme);

  const rows = slide.tableRows ?? [];
  if (rows.length === 0) {
    addChrome(s, theme);
    return;
  }

  const colCount = Math.max(...rows.map((row) => row.length), 1);
  const colW = (W - 1.05) / colCount;
  const tableData = rows.map((row, rowIndex) =>
    Array.from({ length: colCount }, (_, colIndex) => ({
      text: row[colIndex] ?? "",
      options: {
        bold: rowIndex === 0,
        fontSize: rowIndex === 0 ? theme.bodyFontSize - 1 : theme.bodyFontSize - 2,
        fontFace: theme.fontFace,
        color: rowIndex === 0 ? theme.palette.tableHeaderText : theme.palette.bodyText,
        fill: { color: rowIndex === 0 ? theme.palette.tableHeaderBg : rowIndex % 2 === 0 ? theme.palette.tableAltBg : "FFFFFF" },
        align: "center" as const,
        valign: "middle" as const,
        margin: 5,
        border: [
          { type: "solid" as const, pt: 0.5, color: theme.palette.border },
          { type: "solid" as const, pt: 0.5, color: theme.palette.border },
          { type: "solid" as const, pt: 0.5, color: theme.palette.border },
          { type: "solid" as const, pt: 0.5, color: theme.palette.border },
        ] as [any, any, any, any],
      },
    }))
  );

  s.addTable(tableData, {
    x: 0.52,
    y: HEADER_H + 0.24,
    w: 8.25,
    colW: Array(colCount).fill(colW),
  });
  addChrome(s, theme);
}

function buildMultiColumnSlide(
  pptx: PptxGenJS,
  slide: PptxSlide,
  theme: Theme,
  visual: SlideVisualHint
) {
  if (!slide.columns || slide.columns.length === 0) {
    buildBulletsSlide(pptx, slide, theme, visual);
    return;
  }

  const s = pptx.addSlide();
  s.background = { color: theme.palette.canvas };
  addHeaderBand(s, slide.title, theme);
  addVisualAccent(s, visual, theme);

  const columns = slide.columns;
  const totalW = 8.4;
  const colW = totalW / columns.length;
  const colX0 = 0.42;
  const colY = HEADER_H + 0.24;
  const headerH = 0.46;
  const contentH = H - colY - 0.32;

  columns.forEach((column, index) => {
    const x = colX0 + index * colW;
    s.addShape("rect", {
      x,
      y: colY,
      w: colW - 0.08,
      h: headerH,
      fill: { color: theme.palette.sectionBg },
      line: { color: theme.palette.border, width: 0.8 },
    });
    s.addText(column.header, {
      x: x + 0.06,
      y: colY + 0.06,
      w: colW - 0.2,
      h: 0.24,
      fontSize: theme.bodyFontSize - 1,
      fontFace: theme.fontFace,
      bold: true,
      color: theme.palette.bodyText,
      align: "center",
      valign: "middle",
    });
    if (column.bullets.length > 0) {
      const bulletItems = column.bullets.map((item) => ({
        text: item,
        options: {
          bullet: { indent: 10 },
          breakLine: true,
          fontSize: theme.bodyFontSize - 3,
          fontFace: theme.fontFace,
          color: theme.palette.bodyText,
          paraSpaceAfter: 6,
        },
      }));
      s.addText(bulletItems, {
        x: x + 0.06,
        y: colY + headerH + 0.08,
        w: colW - 0.2,
        h: contentH - headerH - 0.1,
        margin: 0.05,
        valign: "top",
      });
    }
  });

  addChrome(s, theme);
}

export async function POST(req: NextRequest) {
  try {
    const body: GenPptxRequest = await req.json();
    const { title, slides, threadId, fontFace, designInstruction, deckPreferences } = body;
    if (!title || !slides || slides.length === 0) {
      return NextResponse.json({ error: "title and slides are required" }, { status: 400 });
    }

    const instructionText = [designInstruction, deckPreferences?.designInstruction, ...(deckPreferences?.recentDesignNotes ?? [])]
      .filter(Boolean)
      .join(" / ");

    const designBrief = await generateDesignBrief(title, slides, instructionText, deckPreferences);
    const theme = resolveTheme(designBrief, instructionText, deckPreferences, fontFace);
    const coverIllustration = await generateCoverIllustration(title, instructionText, theme);

    console.log("[gen-pptx] theme:", {
      instruction: truncateText(instructionText, 80),
      accentColor: deckPreferences?.accentColor,
      language: deckPreferences?.language,
      recentNotes: deckPreferences?.recentDesignNotes,
      execMode: theme.execMode,
      fontFace: theme.fontFace,
      coverIllustration: Boolean(coverIllustration),
    });

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "azurechat";
    pptx.subject = title;
    pptx.title = title;

    buildTitleSlide(pptx, title, designBrief, theme, slides.length, coverIllustration);
    let illustrationPlaced = false;
    slides.forEach((slide, index) => {
      const visual = designBrief.visualHints[index] ?? {
        title: slide.title,
        visualType:
          slide.layoutType === "table"
            ? "table"
            : slide.layoutType === "multi-column"
              ? "comparison"
              : slide.layoutType === "diagram"
                ? "process"
                : "editorial",
        emphasis: slide.bullets[0] ?? slide.title,
      };
      const slideIllustration =
        !illustrationPlaced &&
        coverIllustration &&
        (slide.layoutType ?? "bullets") === "bullets"
          ? coverIllustration
          : null;
      switch (slide.layoutType ?? "bullets") {
        case "title":
          buildSectionSlide(pptx, slide.title, theme);
          break;
        case "table":
          buildTableSlide(pptx, slide, theme, visual);
          break;
        case "multi-column":
          buildMultiColumnSlide(pptx, slide, theme, visual);
          break;
        case "diagram":
          buildDiagramSlide(pptx, slide, theme, visual);
          break;
        default:
          buildBulletsSlide(pptx, slide, theme, visual, slideIllustration);
          if (slideIllustration) illustrationPlaced = true;
          break;
      }
    });

    const buffer = (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
    const fileName = `${threadId ?? uniqueId()}_${uniqueId()}.pptx`;
    const downloadUrl = await uploadPptxToBlob(buffer, fileName);
    return NextResponse.json({ ok: true, downloadUrl, fileName });
  } catch (e: any) {
    console.error("[gen-pptx] error:", e);
    return NextResponse.json({ ok: false, error: String(e?.message ?? e) }, { status: 500 });
  }
}
