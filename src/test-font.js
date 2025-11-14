// test-font.js (CommonJS版)
const fs = require("fs");
const path = require("path");
const sharp = require("sharp");

const fontPath = path.join(process.cwd(), "public", "fonts", "NotoSansJP-Regular.ttf");
const text = "ミダックよろしくね。";
const output = "test-font-output.png";

// フォントファイルを base64 埋め込み
if (!fs.existsSync(fontPath)) {
  console.error("❌ フォントが見つかりません:", fontPath);
  process.exit(1);
}
const fontData = fs.readFileSync(fontPath).toString("base64");
const fontUrl = `data:font/truetype;base64,${fontData}`;

const svg = Buffer.from(`<?xml version="1.0" encoding="UTF-8"?>
<svg width="600" height="200" xmlns="http://www.w3.org/2000/svg">
  <style>
    @font-face {
      font-family: 'NotoJP';
      src: url('${fontUrl}') format('truetype');
      font-weight: 400;
      font-style: normal;
    }
    .jp {
      font-family: 'NotoJP', sans-serif;
      fill: #000;
      font-size: 48px;
    }
  </style>
  <rect width="100%" height="100%" fill="#fff"/>
  <text x="50%" y="50%" text-anchor="middle" alignment-baseline="middle" class="jp">${text}</text>
</svg>`);

(async () => {
  await sharp(svg).png().toFile(output);
  console.log(`✅ 出力完了: ${output}`);
})();
