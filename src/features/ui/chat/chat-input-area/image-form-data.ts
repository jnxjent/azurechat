/** Multipart text fields are truncated at 1 MiB by the server's form parser.
 * Send the data URL as a file part so the complete original image survives. */
export function prepareImageFormData(form: FormData): void {
  const image = form.get("image-base64");
  if (typeof image === "string" && image) {
    form.set("image-base64", new Blob([image], { type: "text/plain;charset=utf-8" }), "image-data-url.txt");
  }
}

export async function readImageFormData(form: FormData): Promise<string> {
  const image = form.get("image-base64");
  return typeof image === "string" ? image : image ? await image.text() : "";
}
