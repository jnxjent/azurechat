const GENERATED_FILE_EXTENSION_RE = /\.(?:docx|xlsx|pptx|pdf|csv)(?:$|[?#])/i;

/** Search results may contain names and source URLs too, so downloadUrl is
 * required to distinguish a successful generated-file tool result. */
export function hasGeneratedFileResult(toolResults: string[]): boolean {
  return toolResults.some((toolResult) => {
    try {
      const parsed = JSON.parse(toolResult) as Record<string, unknown>;
      if (typeof parsed.error === "string" && parsed.error.trim()) return false;

      const fileName =
        typeof parsed.fileName === "string" ? parsed.fileName : "";
      const displayName =
        typeof parsed.displayName === "string" ? parsed.displayName : "";
      const downloadUrl =
        typeof parsed.downloadUrl === "string" ? parsed.downloadUrl : "";
      const message = typeof parsed.message === "string" ? parsed.message : "";
      if (!downloadUrl) return false;

      let decodedDownloadUrl = downloadUrl;
      try {
        decodedDownloadUrl = decodeURIComponent(downloadUrl);
      } catch {
        // Keep the original URL when malformed percent encoding is present.
      }

      return (
        GENERATED_FILE_EXTENSION_RE.test(fileName) ||
        GENERATED_FILE_EXTENSION_RE.test(displayName) ||
        GENERATED_FILE_EXTENSION_RE.test(decodedDownloadUrl) ||
        /(?:Word|Excel|PowerPoint|PDF|ファイル).*(?:作成|生成|編集|変換|完了)/i.test(
          message
        )
      );
    } catch {
      return false;
    }
  });
}
