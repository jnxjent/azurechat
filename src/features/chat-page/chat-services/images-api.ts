import {
  GetImageFromStore,
  GetThreadAndImageFromUrl,
} from "./chat-image-service";

/**
 * imgName ("HcCu2FQ...Ins.png" など) から
 * ベース画像用のファイル名 ("HcCu2FQ...Ins_base.png") を生成する。
 *
 * - 拡張子がなければ ".png" を付与
 * - すでに "_base" が付いている場合はそのまま返す
 */
const toBaseImageName = (imgName: string): string => {
  if (!imgName) return imgName;

  // すでに _base が付いているならそのまま
  if (/_base(\.[A-Za-z0-9]+)?$/.test(imgName)) {
    return imgName;
  }

  const match = imgName.match(/^(.*?)(\.[A-Za-z0-9]+)?$/);
  if (!match) return imgName;

  const stem = match[1]; // 拡張子を除いた部分
  const ext = match[2] || ".png";
  return `${stem}_base${ext}`;
};

export const ImageAPIEntry = async (request: Request): Promise<Response> => {
  const urlPath = request.url;

  const response = GetThreadAndImageFromUrl(urlPath);

  if (response.status !== "OK") {
    return new Response(response.errors[0].message, { status: 404 });
  }

  const { threadId, imgName } = response.response;

  // 1️⃣ まず「ベース画像名」を作る（例: xxx.png → xxx_base.png）
  const baseImgName = toBaseImageName(imgName);

  // 2️⃣ もしベース画像があれば、そちらを優先して返す
  const baseImageData = await GetImageFromStore(threadId, baseImgName);
  if (baseImageData.status === "OK") {
    return new Response(baseImageData.response, {
      headers: { "content-type": "image/png" },
    });
  }

  // 3️⃣ ベース画像がない場合は、従来どおり元の imgName を使う
  const imageData = await GetImageFromStore(threadId, imgName);

  if (imageData.status === "OK") {
    return new Response(imageData.response, {
      headers: { "content-type": "image/png" },
    });
  } else {
    return new Response(imageData.errors[0].message, { status: 404 });
  }
};
