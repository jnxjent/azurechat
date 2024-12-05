"use server";

import { FindCitationByID } from "../chat-services/citation-service";
import { GenerateSasUrl } from "@/features/common/services/azure-storage";
export const CitationFileDownload = async (
  formData: FormData
) => {
  const searchResponse = await FindCitationByID(formData.get("id") as string);

  if (searchResponse.status === "OK") {
    const response = searchResponse.response;
    const {document} = response.content;
    return document.fileUrl
  }

  return null;
};
