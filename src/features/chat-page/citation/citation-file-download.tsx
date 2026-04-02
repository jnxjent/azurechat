"use server";

import { FindCitationByID } from "../chat-services/citation-service";
import { GenerateSasUrl } from "@/features/common/services/azure-storage";

export const CitationFileDownload = async (formData: FormData) => {
  console.log("[DL] CitationFileDownload called, id=", formData.get("id"));
  const searchResponse = await FindCitationByID(formData.get("id") as string);
  if (searchResponse.status === "OK") {
    const response = searchResponse.response;
    const { document } = response.content;
    console.log("[DL] effectiveFileUrl=", document.effectiveFileUrl, "fileUrl=", document.fileUrl);
    return document.effectiveFileUrl || document.fileUrl;
  }
  return null;
};