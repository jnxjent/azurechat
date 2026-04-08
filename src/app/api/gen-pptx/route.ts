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

export type PptxSlide = {
  title: string;
  bullets: string[];
};

export type GenPptxRequest = {
  title: string;
  slides: PptxSlide[];
  threadId: string;
  fontFace?: string;
};

async function uploadPptxToBlob(
  buffer: Buffer,
  fileName: string
): Promise<string> {
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
      blobContentType:
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      blobContentDisposition: `attachment; filename="${fileName}"`,
    },
  });

  const sasToken = generateBlobSASQueryParameters(
    {
      containerName,
      blobName: fileName,
      expiresOn: new Date(Date.now() + 24 * 60 * 60 * 1000),
      permissions: BlobSASPermissions.parse("r"),
    },
    sharedKeyCredential
  );

  return `${blockBlobClient.url}?${sasToken}`;
}

export async function POST(req: NextRequest) {
  try {
    const body: GenPptxRequest = await req.json();
    const { title, slides, threadId, fontFace } = body;
    const resolvedFontFace = fontFace?.trim() || "Arial";

    if (!title || !slides || slides.length === 0) {
      return NextResponse.json(
        { error: "title and slides are required" },
        { status: 400 }
      );
    }

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "azurechat";
    pptx.subject = title;
    pptx.title = title;
    pptx.company = "azurechat";
    pptx.lang = "ja-JP";

    const titleSlide = pptx.addSlide();
    titleSlide.background = { color: "1F3864" };
    titleSlide.addText(title, {
      x: 0.5,
      y: 2.2,
      w: 12.3,
      h: 1.4,
      fontSize: 28,
      fontFace: resolvedFontFace,
      bold: true,
      color: "FFFFFF",
      align: "center",
      valign: "mid",
      margin: 0.1,
    });

    for (const slide of slides) {
      const s = pptx.addSlide();
      s.background = { color: "FFFFFF" };

      s.addText(slide.title, {
        x: 0.5,
        y: 0.3,
        w: 12.3,
        h: 0.7,
        fontSize: 22,
        fontFace: resolvedFontFace,
        bold: true,
        color: "1F3864",
        margin: 0.05,
      });

      s.addShape(pptx.ShapeType.line, {
        x: 0.5,
        y: 1.05,
        w: 12.3,
        h: 0,
        line: { color: "1F3864", width: 1.5 },
      });

      if (slide.bullets.length > 0) {
        const bulletItems = slide.bullets.map((bullet) => ({
          text: bullet,
          options: {
            bullet: { indent: 14 },
            breakLine: true,
            fontSize: 18,
            fontFace: resolvedFontFace,
            color: "333333",
          },
        }));

        s.addText(bulletItems, {
          x: 0.7,
          y: 1.35,
          w: 11.8,
          h: 5.2,
          margin: 0.08,
          valign: "top",
          paraSpaceAfterPt: 10,
          breakLine: false,
        });
      }
    }

    const buffer = (await pptx.write({
      outputType: "nodebuffer",
    })) as Buffer;

    const fileName = `${threadId ?? uniqueId()}_${uniqueId()}.pptx`;
    const downloadUrl = await uploadPptxToBlob(buffer, fileName);

    return NextResponse.json({ ok: true, downloadUrl, fileName });
  } catch (e: any) {
    console.error("[gen-pptx] error:", e);
    return NextResponse.json(
      { ok: false, error: String(e?.message ?? e) },
      { status: 500 }
    );
  }
}
