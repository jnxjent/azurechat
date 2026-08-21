import { BlobSASPermissions, BlobServiceClient, RestError } from "@azure/storage-blob";
import { ServerActionResponse } from "../server-action-response";

// initialize the blobServiceClient
const InitBlobServiceClient = () => {
  // trim() で環境変数末尾の改行・空白を除去（混入すると署名が狂う）
  const acc = (process.env.AZURE_STORAGE_ACCOUNT_NAME ?? "").trim();
  const key = (process.env.AZURE_STORAGE_ACCOUNT_KEY ?? "").trim();

  if (!acc || !key)
    throw new Error(
      "Azure Storage Account not configured correctly, check environment variables."
    );

  const connectionString = `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`;

  return BlobServiceClient.fromConnectionString(connectionString);
};

export const UploadBlob = async (
  containerName: string,
  blobName: string,
  blobData: Buffer,
  returnName: boolean = false
): Promise<ServerActionResponse<string>> => {
  const blobServiceClient = InitBlobServiceClient();

  const containerClient = blobServiceClient.getContainerClient(containerName);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);

  const response = await blockBlobClient.uploadData(blobData);

  // Check for upload success
  if (response.errorCode !== undefined) {
    return {
      status: "ERROR",
      errors: [
        {
          message: `Error uploading blob to storage: ${response.errorCode}`,
        },
      ],
    };
  }
  return {
    status: "OK",
    response: blockBlobClient.url,
  };
};

export const GenerateSasUrl = async (
  containerName: string,
  blobPath: string
): Promise<ServerActionResponse<string>> => {
  try {
    const blockBlobClient = InitBlobServiceClient()
      .getContainerClient(containerName)
      .getBlockBlobClient(blobPath);
    const sasUrl = await blockBlobClient.generateSasUrl({
      expiresOn: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
      permissions: BlobSASPermissions.parse("r"),
    });
    return { status: "OK", response: sasUrl };
  } catch (e) {
    return { status: "ERROR", errors: [{ message: `GenerateSasUrl failed: ${String(e)}` }] };
  }
}
  

export const DownloadBlobAsText = async (
  containerName: string,
  blobPath: string
): Promise<ServerActionResponse<string>> => {
  try {
    const blobServiceClient = InitBlobServiceClient();
    const containerClient = blobServiceClient.getContainerClient(containerName);
    const blockBlobClient = containerClient.getBlockBlobClient(blobPath);
    const downloadResponse = await blockBlobClient.download(0);
    if (!downloadResponse.readableStreamBody) {
      return { status: "ERROR", errors: [{ message: "Empty response body" }] };
    }
    const chunks: Buffer[] = [];
    for await (const chunk of downloadResponse.readableStreamBody as any) {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    return { status: "OK", response: Buffer.concat(chunks).toString("utf-8") };
  } catch (e: any) {
    if (e?.statusCode === 404 || e?.code === "BlobNotFound") {
      return { status: "NOT_FOUND", errors: [{ message: `Blob not found: ${blobPath}` }] };
    }
    return { status: "ERROR", errors: [{ message: String(e) }] };
  }
};

export const DeleteBlob = async (
  containerName: string,
  blobPath: string
): Promise<ServerActionResponse<boolean>> => {
  try {
    const deleted = await InitBlobServiceClient()
      .getContainerClient(containerName)
      .getBlockBlobClient(blobPath)
      .deleteIfExists();
    return { status: "OK", response: deleted.succeeded };
  } catch (e) {
    return { status: "ERROR", errors: [{ message: `DeleteBlob failed: ${String(e)}` }] };
  }
};

export const GetBlob = async (
  containerName: string,
  blobPath: string
): Promise<ServerActionResponse<ReadableStream<any>>> => {
  const blobServiceClient = InitBlobServiceClient();

  const containerClient = blobServiceClient.getContainerClient(containerName);
  const blockBlobClient = containerClient.getBlockBlobClient(blobPath);

  try {
    const downloadBlockBlobResponse = await blockBlobClient.download(0);

    // Passes stream to caller to decide what to do with
    if (!downloadBlockBlobResponse.readableStreamBody) {
      return {
        status: "ERROR",
        errors: [
          {
            message: `Error downloading blob: ${blobPath}`,
          },
        ],
      };
    }

    return {
      status: "OK",
      response:
        downloadBlockBlobResponse.readableStreamBody as unknown as ReadableStream<any>,
    };
  } catch (error) {
    if (error instanceof RestError) {
      if (error.statusCode === 404) {
        return {
          status: "NOT_FOUND",
          errors: [
            {
              message: `Blob not found: ${blobPath}`,
            },
          ],
        };
      }
    }

    return {
      status: "ERROR",
      errors: [
        {
          message: `Error downloading blob: ${blobPath}`,
        },
      ],
    };
  }
};
