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

// ── Types ────────────────────────────────────────────────────────────────

type ExcelSheetDef = {
  name: string;
  rows: (string | number | null)[][];
  headerRowIndex?: number;
  autoWidth?: boolean;
};

type ExcelDocPlan = {
  sheets: ExcelSheetDef[];
};

export type GenExcelRequest = {
  content: string;
  instruction?: string;
  title?: string;
  threadId: string;
};

// ── LLM plan generation ───────────────────────────────────────────────────

async function generateExcelPlan(
  content: string,
  instruction: string,
  titleHint: string
): Promise<ExcelDocPlan> {
  const openai = OpenAIInstance();

  const res = await openai.chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      {
        role: "system",
        content: [
          "You are an Excel document formatter. Given raw content and an instruction, return a JSON ExcelDocPlan.",
          "Schema:",
          '{',
          '  "sheets": [',
          '    {',
          '      "name": "Sheet1",',
          '      "rows": [',
          '        ["", "野元", "山本", "鈴木"],',
          '        ["点数", 100, 60, 40]',
          '      ],',
          '      "headerRowIndex": 0',
          '    }',
          '  ]',
          '}',
          "Rules:",
          "- Identify all tabular data in the content and map each logical table to a sheet.",
          "- rows is an array of arrays. Each inner array is one row of cells.",
          "- Use the actual values from the content — do not invent or alter numbers.",
          "- If a cell is empty, use null or empty string.",
          "- headerRowIndex: set to 0 if the first row is a header row (column labels), otherwise omit.",
          "- name: a short sheet name describing the data (ASCII or Japanese ok, max 31 chars).",
          "- If the content has multiple independent tables, create multiple sheets.",
          "- Number values should be numbers (not strings) so Excel treats them as numeric.",
          "- Return JSON only, no explanation.",
        ].join("\n"),
      },
      {
        role: "user",
        content: JSON.stringify({ title: titleHint, content, instruction }),
      },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 8000,
  });

  const raw = res.choices[0]?.message?.content ?? "{}";
  const parsed = JSON.parse(raw);

  const sheets: ExcelSheetDef[] = Array.isArray(parsed?.sheets) ? parsed.sheets : [];

  if (sheets.length === 0 && content.trim()) {
    // Fallback: treat raw content as single-column text data
    const lines = content.trim().split("\n").filter((l) => l.trim());
    sheets.push({
      name: titleHint || "Sheet1",
      rows: lines.map((line) => [line]),
    });
  }

  return { sheets };
}

// ── Script path resolution ────────────────────────────────────────────────

async function resolveCreateExcelScriptPath(): Promise<string> {
  const candidates = [
    path.join(process.cwd(), "src", "scripts", "create_excel.py"),
    path.join(process.cwd(), "scripts", "create_excel.py"),
    "/home/site/wwwroot/src/scripts/create_excel.py",
    "/home/site/wwwroot/scripts/create_excel.py",
  ];
  for (const candidate of candidates) {
    try {
      await fs.access(candidate);
      return candidate;
    } catch {
      // not found, try next
    }
  }
  throw new Error(`create_excel.py not found. Checked: ${candidates.join(", ")}`);
}

// ── Blob upload ───────────────────────────────────────────────────────────

async function uploadExcelToBlob(buffer: Buffer, fileName: string): Promise<string> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const containerName = "xlsx";

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
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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

// ── Python execution ──────────────────────────────────────────────────────

async function runPythonCreateExcel(plan: ExcelDocPlan, threadId: string) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-xlsx-"));
  const outputPath = path.join(tempDir, "output.xlsx");
  const planPath = path.join(tempDir, "plan.json");
  const scriptPath = await resolveCreateExcelScriptPath();

  // Explicitly set Python paths for App Service (startup.sh installs to /home/site/python-packages)
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
        await execFileAsync(pythonBin, ["-c", "import openpyxl"], { env: pyEnv });
      } catch {
        throw new Error(
          "openpyxl is not installed on the server. Check startup.sh."
        );
      }
    }

    const { stdout, stderr } = await execFileAsync(pythonBin, [
      scriptPath,
      "--output", outputPath,
      "--plan", planPath,
    ], { env: pyEnv });

    if (stderr?.trim()) {
      console.warn("[gen-excel] python stderr:", stderr.trim());
    }

    const outputBuffer = await fs.readFile(outputPath);
    const pythonResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};

    // ASCII-only filename to avoid ERR_INVALID_CHAR in Content-Disposition header
    const safeTitle = (plan.sheets[0]?.name ?? "")
      .replace(/[^\w]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "")
      .slice(0, 40) || "workbook";
    const fileName = `${threadId || uniqueId()}_${safeTitle}_${uniqueId()}.xlsx`;
    const downloadUrl = await uploadExcelToBlob(outputBuffer, fileName);

    return {
      downloadUrl,
      fileName,
      sheets: Number(pythonResult.sheets ?? 0),
      totalRows: Number(pythonResult.totalRows ?? 0),
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

// ── POST handler ──────────────────────────────────────────────────────────

export async function POST(req: NextRequest) {
  try {
    const body: GenExcelRequest = await req.json();
    const { content, instruction, title, threadId } = body;

    if (!content?.trim() && !title?.trim()) {
      return NextResponse.json(
        { error: "content or title is required." },
        { status: 400 }
      );
    }

    const plan = await generateExcelPlan(
      content ?? "",
      instruction?.trim() ?? "",
      title?.trim() ?? ""
    );

    const result = await runPythonCreateExcel(plan, threadId ?? uniqueId());

    return NextResponse.json(result);
  } catch (error: any) {
    console.error("[gen-excel] error:", error);
    return NextResponse.json(
      { error: String(error?.message ?? "Excel generation failed.") },
      { status: 500 }
    );
  }
}
