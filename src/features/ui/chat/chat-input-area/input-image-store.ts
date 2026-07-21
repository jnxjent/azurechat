import { proxy, useSnapshot } from "valtio";

const SUPPORTED_CHAT_IMAGE_TYPES = new Set([
  "image/png",
  "image/jpeg",
  "image/webp",
]);

const SUPPORTED_CHAT_IMAGE_EXTENSION = /\.(?:png|jpe?g|webp)$/i;

export function isSupportedChatImageFile(file: File): boolean {
  return (
    SUPPORTED_CHAT_IMAGE_TYPES.has(file.type.toLowerCase()) ||
    (!file.type && SUPPORTED_CHAT_IMAGE_EXTENSION.test(file.name))
  );
}

class InputImageState {
  public previewImage: string = "";
  public base64Image: string = "";
  public fileUrl: string = "";

  get PreViewImage() {
    return this.previewImage;
  }

  public UpdateBase64Image(image: string) {
    this.base64Image = image;
  }

  public Reset() {
    if (this.previewImage) URL.revokeObjectURL(this.previewImage);
    this.previewImage = "";
    this.base64Image = "";
    this.fileUrl = "";
  }

  private fileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        if (typeof reader.result === "string") {
          resolve(reader.result);
        } else {
          reject("Error converting file to base64");
        }
      };
    });
  }

  public async OnFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) await this.SetFile(file);
  }

  public async SetFile(file: File) {
    if (!isSupportedChatImageFile(file)) {
      throw new Error("画像はPNG、JPEG、WebP形式を選択してください。");
    }

    const base64 = await this.fileToBase64(file);
    const url = URL.createObjectURL(file);
    if (this.previewImage) URL.revokeObjectURL(this.previewImage);
    this.previewImage = url;
    this.base64Image = base64;
    this.fileUrl = file.name;
  }
}

export const InputImageStore = proxy(new InputImageState());

export const useInputImage = () => {
  return useSnapshot(InputImageStore, { sync: true });
};
