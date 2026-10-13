import JSZip from "jszip";
import sharp from "sharp";

/** Overlay the original logo after all design/vision passes, on every slide. */
export async function addLogoToPptx(buffer: Buffer, dataUrl: string): Promise<Buffer> {
  const match = /^data:image\/(?:png|jpeg|webp);base64,([a-z0-9+/=\s]+)$/i.exec(dataUrl);
  if (!match) throw new Error("ロゴはPNG・JPEG・WEBP画像で指定してください。");
  const source = Buffer.from(match[1], "base64");
  if (!source.length || source.length > 15 * 1024 * 1024) throw new Error("ロゴ画像のサイズが不正です。");
  const png = await sharp(source).rotate().png().toBuffer();
  const meta = await sharp(png).metadata();
  const zip = await JSZip.loadAsync(buffer);
  const presentation = await zip.file("ppt/presentation.xml")!.async("string");
  const size = /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(presentation);
  if (!size || !meta.width || !meta.height) throw new Error("PPTまたはロゴの寸法を取得できません。");
  const sw = Number(size[1]), sh = Number(size[2]);
  const scale = Math.min(sw * .115 / meta.width, sh * .065 / meta.height);
  const w = Math.round(meta.width * scale), h = Math.round(meta.height * scale);
  const x = Math.round(sw - w - sw * .025), y = Math.round(sh * .02);
  zip.file("ppt/media/azurechat-company-logo.png", png);
  let types = await zip.file("[Content_Types].xml")!.async("string");
  if (!/Extension="png"/i.test(types)) types = types.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>');
  zip.file("[Content_Types].xml", types);
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    let xml = await zip.file(name)!.async("string");
    const relName = name.replace("slides/", "slides/_rels/") + ".rels";
    let rels = zip.file(relName) ? await zip.file(relName)!.async("string") : '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    let rid = "rIdCompanyLogo";
    while (rels.includes(`Id="${rid}"`)) rid += "L";
    const id = Math.max(0, ...Array.from(xml.matchAll(/<p:cNvPr[^>]*id="(\d+)"/g), m => Number(m[1]))) + 1;
    rels = rels.replace("</Relationships>", `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/azurechat-company-logo.png"/></Relationships>`);
    const pic = `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Company logo"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    xml = xml.replace("</p:spTree>", pic + "</p:spTree>");
    zip.file(name, xml); zip.file(relName, rels);
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
