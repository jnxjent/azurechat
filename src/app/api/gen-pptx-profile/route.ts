export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import { execFile } from "child_process";
import { promisify } from "util";
import * as fs from "fs/promises";
import * as os from "os";
import * as path from "path";
import {
  BlobSASPermissions,
  BlobServiceClient,
} from "@azure/storage-blob";
import { uniqueId } from "@/features/common/util";

const execFileAsync = promisify(execFile);

// ── パレット定義（gen_pptx_profile.py の PALETTES と同期） ───────────────
const PYTHON_PALETTES = {
  // 経営・IT・テック・AI 系（落ち着いた紺でプロフェッショナル感）
  navy_orange:    { label: "紺×オレンジ（経営・IT・DX・AI）",
    keywords: ["役員","経営","it","dx","ai","azurechat","システム","導入","デジタル","紺","ネイビー","提案","稟議","セキュリティ","クラウド","platform","saas"] },
  // 人材・採用・成長 + 農業・環境・食品 系（緑＝育ち・生命感）
  forest_amber:   { label: "深緑×琥珀（採用・人材・農業・食品・エコ）",
    keywords: ["採用","人材","インターン","募集","リクルート","学生","新卒","就活","活躍","成長","研修","教育",
               "農業","食品","エコ","環境","自然","サステナ","グリーン","森","林"] },
  // 伝統・高級・製造・ものづくり 系（赤の重厚感）
  burgundy_gold:  { label: "深赤×金（伝統・高級・製造・ものづくり）",
    keywords: ["伝統","高級","老舗","製造","工業","クラフト","重厚","酒造","醸造","文化","工場","品質","精密"] },
  // 産廃・廃棄物・リサイクル + 医療・ヘルス + スタートアップ系（動的なブルーグリーン）
  teal_coral:     { label: "青緑×珊瑚（産廃・廃棄物・医療・ヘルス・スタートアップ）",
    keywords: ["産廃","廃棄物","廃棄","リサイクル","収集","運搬","中間処理","最終処分","ごみ","廃",
               "医療","ヘルス","クリニック","病院","薬","スタートアップ","health","clinic"] },
  // 建設・インフラ・重厚産業 系（落ち着いたアース系）
  charcoal_terra: { label: "チャコール×テラ（建設・土木・インフラ・重工）",
    keywords: ["建設","土木","インフラ","アース","重機","工事","施設","建物","鉄道","電力","ガス","プラント"] },
} as const;

type PythonPaletteName = keyof typeof PYTHON_PALETTES;

/**
 * プロンプトのキーワードから最適なパレットを選択する。
 * LLM が明示的に `palette` を渡した場合はそちらを優先。
 */
function selectPythonPalette(instructionText: string): PythonPaletteName {
  const h = instructionText.toLowerCase();
  for (const [key, { keywords }] of Object.entries(PYTHON_PALETTES) as [PythonPaletteName, typeof PYTHON_PALETTES[PythonPaletteName]][]) {
    if (key === "navy_orange") continue; // default は最後に
    if (keywords.some((kw) => h.includes(kw.toLowerCase()))) {
      return key;
    }
  }
  return "navy_orange";
}

// ── Python スクリプトパス解決 ────────────────────────────────────────────
async function resolvePythonScriptPath(): Promise<string> {
  const candidates = [
    path.join(process.cwd(), "src", "scripts", "gen_pptx_profile.py"),
    path.join(process.cwd(), "scripts", "gen_pptx_profile.py"),
    "/home/site/wwwroot/src/scripts/gen_pptx_profile.py",
    "/home/site/wwwroot/scripts/gen_pptx_profile.py",
  ];
  for (const c of candidates) {
    try { await fs.access(c); return c; } catch {}
  }
  throw new Error("gen_pptx_profile.py が見つかりません");
}

// ── Blob アップロード ────────────────────────────────────────────────────
async function uploadPptxToBlob(buffer: Buffer, blobKey: string, displayFileName?: string): Promise<string> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const client = BlobServiceClient.fromConnectionString(
    `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
  const cc = client.getContainerClient("pptx");
  await cc.createIfNotExists({ access: "blob" });
  const bc = cc.getBlockBlobClient(blobKey);
  await bc.uploadData(buffer, {
    blobHTTPHeaders: {
      blobContentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      blobContentDisposition: `attachment; filename*=UTF-8''${encodeURIComponent(displayFileName ?? blobKey)}`,
    },
  });
  return bc.generateSasUrl({
    expiresOn: new Date(Date.now() + 24 * 60 * 60 * 1000),
    permissions: BlobSASPermissions.parse("r"),
  });
}

// ── POST ─────────────────────────────────────────────────────────────────
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      title,
      slides,
      palette: paletteFromClient,
      designInstruction,
      threadId,
      fileBaseName,
    } = body as {
      title: string;
      slides: Array<{ layoutType?: string; [k: string]: unknown }>;
      palette?: string;
      designInstruction?: string;
      threadId?: string;
      fileBaseName?: string;
    };

    if (!title || !slides?.length) {
      return NextResponse.json({ ok: false, error: "title and slides are required" }, { status: 400 });
    }

    // パレット選択: クライアント指定 > キーワード自動選択
    const instructionText = [designInstruction, title].filter(Boolean).join(" ");
    const palette = (paletteFromClient && paletteFromClient in PYTHON_PALETTES)
      ? (paletteFromClient as PythonPaletteName)
      : selectPythonPalette(instructionText);

    console.log(`[gen-pptx-profile] palette=${palette} slides=${slides.length}`);

    const scriptPath = await resolvePythonScriptPath();

    const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-pptx-py-"));
    const inputPath  = path.join(tempDir, "input.json");
    const outputPath = path.join(tempDir, "output.pptx");

    try {
      const inputJson = JSON.stringify({ title, slides, slideCount: slides.length });
      await fs.writeFile(inputPath, inputJson, "utf8");

      const pythonBin = process.platform === "win32" ? "python" : "python3";
      const pyEnv = process.platform !== "win32"
        ? {
            ...process.env,
            PYTHONPATH: `/home/site/python-packages${process.env.PYTHONPATH ? `:${process.env.PYTHONPATH}` : ""}`,
            LD_LIBRARY_PATH: `/home/site/python-packages${process.env.LD_LIBRARY_PATH ? `:${process.env.LD_LIBRARY_PATH}` : ""}`,
          }
        : process.env;

      const { stdout, stderr } = await execFileAsync(
        pythonBin,
        [scriptPath, "--input", inputPath, "--output", outputPath, "--palette", palette],
        { env: pyEnv, timeout: 60000 }
      );

      if (stderr?.trim()) console.warn("[gen-pptx-profile] python stderr:", stderr.trim());

      const pyResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};
      if (!pyResult.ok) throw new Error(pyResult.error ?? "Python script failed");

      const buffer = await fs.readFile(outputPath);
      const safeBase = fileBaseName
        ? fileBaseName.replace(/\.pptx$/i, "").replace(/[\\/:*?"<>|]/g, "").trim().slice(0, 40)
        : (threadId ?? uniqueId());
      const displayFileName = `${safeBase}.pptx`;
      const blobKey = `pptx_${uniqueId().slice(0, 8)}.pptx`;
      const downloadUrl = await uploadPptxToBlob(buffer, blobKey, displayFileName);

      return NextResponse.json({ ok: true, downloadUrl, fileName: displayFileName, palette });
    } finally {
      await fs.rm(tempDir, { recursive: true, force: true });
    }
  } catch (e: any) {
    console.error("[gen-pptx-profile] error:", e);
    return NextResponse.json({ ok: false, error: String(e?.message ?? e) }, { status: 500 });
  }
}
