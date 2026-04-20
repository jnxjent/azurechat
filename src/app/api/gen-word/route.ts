export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import {
  BlobSASPermissions,
  BlobServiceClient,
  StorageSharedKeyCredential,
  generateBlobSASQueryParameters,
} from "@azure/storage-blob";
import { OpenAIInstance } from "@/features/common/services/openai";
import { uniqueId } from "@/features/common/util";

const execFileAsync = promisify(execFile);

// ── 型定義 ────────────────────────────────────────────────────────────────

type WordTableDef = {
  headers: string[];
  rows: string[][];
};

type WordSection = {
  heading?: string;
  level?: number;
  paragraphs?: string[];
  bullets?: string[];
  table?: WordTableDef;
};

type WordDocStyle = {
  fontFace?: string;
  fontSize?: number;
  titleFontSize?: number;
};

type WordDocPlan = {
  title: string;
  sections: WordSection[];
  style?: WordDocStyle;
};

export type GenWordRequest = {
  content: string;
  instruction?: string;
  title?: string;
  threadId: string;
  fontFace?: string;
};

// ── LLM で WordDocPlan を生成 ─────────────────────────────────────────────

async function generateWordPlan(
  content: string,
  instruction: string,
  titleHint: string,
  fontFace: string
): Promise<WordDocPlan> {
  const openai = OpenAIInstance();

  const res = await openai.chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      {
        role: "system",
        content: [
          "You are a Word document formatter. Given raw content and an instruction, return a JSON WordDocPlan.",
          "Schema:",
          '{',
          '  "title": "文書タイトル",',
          '  "sections": [',
          '    {',
          '      "heading": "見出し（省略可）",',
          '      "level": 1,',
          '      "paragraphs": ["本文テキスト"],',
          '      "bullets": ["箇条書きテキスト"],',
          '      "table": { "headers": ["列1","列2"], "rows": [["A","B"]] }',
          '    }',
          '  ],',
          '  "style": { "fontFace": "Meiryo", "fontSize": 11, "titleFontSize": 16 }',
          '}',
          "Rules:",
          "- title: Use the user-provided title or infer from content.",
          "- Split content logically into sections. Use headings where appropriate.",
          "- Use paragraphs for prose, bullets for lists, table for tabular data.",
          "- A section may have heading only, paragraphs only, bullets only, or a mix.",
          "- table is optional; only include when the content is clearly tabular.",
          "- Preserve the exact text from the content — do not paraphrase or summarize.",
          "- Return JSON only, no explanation.",
        ].join("\n"),
      },
      {
        role: "user",
        content: JSON.stringify({ title: titleHint, content, instruction, fontFace }),
      },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 8000,
  });

  const raw = res.choices[0]?.message?.content ?? "{}";
  const parsed = JSON.parse(raw);

  // バリデーション・フォールバック
  const title = String(parsed?.title ?? titleHint ?? "文書").trim() || "文書";
  const sections: WordSection[] = Array.isArray(parsed?.sections) ? parsed.sections : [];

  // sections が空の場合はコンテンツ全体を1段落として扱う
  if (sections.length === 0 && content.trim()) {
    sections.push({ paragraphs: [content.trim()] });
  }

  const style: WordDocStyle = {
    fontFace: String(parsed?.style?.fontFace ?? fontFace ?? "Meiryo"),
    fontSize: Number(parsed?.style?.fontSize ?? 11),
    titleFontSize: Number(parsed?.style?.titleFontSize ?? 16),
  };

  return { title, sections, style };
}

// ── Pythonスクリプトパス解決 ──────────────────────────────────────────────

async function resolveCreateWordScriptPath(): Promise<string> {
  const candidates = [
    path.join(process.cwd(), "src", "scripts", "create_word.py"),
    path.join(process.cwd(), "scripts", "create_word.py"),
    "/home/site/wwwroot/src/scripts/create_word.py",
    "/home/site/wwwroot/scripts/create_word.py",
  ];
  for (const candidate of candidates) {
    try {
      await fs.access(candidate);
      return candidate;
    } catch {
      // not found, try next
    }
  }
  throw new Error(`create_word.py not found. Checked: ${candidates.join(", ")}`);
}

// ── Blob アップロード ─────────────────────────────────────────────────────

async function uploadWordToBlob(buffer: Buffer, fileName: string): Promise<string> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const containerName = "docx";

  const cred = new StorageSharedKeyCredential(acc, key);
  const svc = BlobServiceClient.fromConnectionString(
    `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
  const cc = svc.getContainerClient(containerName);
  await cc.createIfNotExists({ access: "blob" });

  const bbc = cc.getBlockBlobClient(fileName);
  await bbc.uploadData(buffer, {
    blobHTTPHeaders: {
      blobContentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      blobContentDisposition: `attachment; filename="${fileName}"`,
    },
  });

  const sas = generateBlobSASQueryParameters(
    {
      containerName,
      blobName: fileName,
      expiresOn: new Date(Date.now() + 24 * 60 * 60 * 1000),
      permissions: BlobSASPermissions.parse("r"),
    },
    cred
  );
  return `${bbc.url}?${sas}`;
}

// ── Python 実行 ───────────────────────────────────────────────────────────

async function runPythonCreateWord(plan: WordDocPlan, threadId: string) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-docx-"));
  const outputPath = path.join(tempDir, "output.docx");
  const planPath = path.join(tempDir, "plan.json");
  const scriptPath = await resolveCreateWordScriptPath();

  // PYTHONPATH・LD_LIBRARY_PATH を明示的に設定（startup.sh が動いていない環境でも動作させるため）
  const pyEnv = process.platform !== "win32"
    ? {
        ...process.env,
        PYTHONPATH: `/home/site/python-packages${process.env.PYTHONPATH ? `:${process.env.PYTHONPATH}` : ""}`,
        LD_LIBRARY_PATH: `/home/site/python-packages${process.env.LD_LIBRARY_PATH ? `:${process.env.LD_LIBRARY_PATH}` : ""}`,
      }
    : process.env;

  try {
    await fs.writeFile(planPath, JSON.stringify(plan), "utf8");

    const pythonBin = process.platform === "win32" ? "python" : "python3";

    if (process.platform !== "win32") {
      try {
        await execFileAsync(pythonBin, ["-c", "import docx"], { env: pyEnv });
      } catch {
        throw new Error(
          "python-docx がサーバーにインストールされていません。" +
          "startup.sh の設定を確認してください。"
        );
      }
    }

    const { stdout, stderr } = await execFileAsync(pythonBin, [
      scriptPath,
      "--output", outputPath,
      "--plan", planPath,
    ], { env: pyEnv });

    if (stderr?.trim()) {
      console.warn("[gen-word] python stderr:", stderr.trim());
    }

    const outputBuffer = await fs.readFile(outputPath);
    const pythonResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};

    // ASCII word chars only — Japanese/non-ASCII are invalid in HTTP header filenames
    const safeTitle = plan.title
      .replace(/[^\w]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "")
      .slice(0, 40) || "document";
    const fileName = `${threadId || uniqueId()}_${safeTitle}_${uniqueId()}.docx`;
    const downloadUrl = await uploadWordToBlob(outputBuffer, fileName);

    return {
      downloadUrl,
      fileName,
      paragraphs: Number(pythonResult.paragraphs ?? 0),
      tables: Number(pythonResult.tables ?? 0),
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

// ── POST ハンドラ ─────────────────────────────────────────────────────────

export async function POST(req: NextRequest) {
  try {
    const body: GenWordRequest = await req.json();
    const { content, instruction, title, threadId, fontFace } = body;

    if (!content?.trim() && !title?.trim()) {
      return NextResponse.json(
        { error: "content または title を指定してください。" },
        { status: 400 }
      );
    }

    const resolvedFont = fontFace?.trim() || "Meiryo";
    const resolvedTitle = title?.trim() || "";
    const resolvedInstruction = instruction?.trim() || "";

    const plan = await generateWordPlan(
      content ?? "",
      resolvedInstruction,
      resolvedTitle,
      resolvedFont
    );

    const result = await runPythonCreateWord(plan, threadId ?? uniqueId());

    return NextResponse.json(result);
  } catch (error: any) {
    console.error("[gen-word] error:", error);
    return NextResponse.json(
      { error: String(error?.message ?? "Word生成中にエラーが発生しました。") },
      { status: 500 }
    );
  }
}
