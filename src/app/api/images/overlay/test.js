// app/api/images/overlay/test.js
// test-font.js 相当のテストスクリプト（構文エラー修正版）
import fs from "fs";
import path from "path";
import sharp from "sharp";

const fontPath = path.join(process.cwd(), "public", "fonts", "NotoSansJP-Regular.ttf");
const text = "ミダックよろしくね。";
const output = path.join(process.cwd(), "test-font-output.png");

// フォントファイルを base64 埋め込み
if (!fs.existsSync(fontPath)) {
  console.error("❌ フォントが見つかりません:", fontPath);
  process.exit(1);
}

const fontData = fs.readFileSync(fontPath).toString("base64");
const fontUrl = `data:font/truetype;base64,${fontData}`;

// SVG 文字列（テンプレートリテラルをちゃんと閉じる）
const svgString = `<?xml version="1.0" encoding="UTF-8"?>
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
      fill: #000000;
      font-size: 48px;
    }
  </style>
  <rect width="100%" height="100%" fill="#ffffff"/>
  <text x="50%" y="50%" text-anchor="middle" dominant-baseline="central" class="jp">${text}</text>
</svg>`;

const svgBuffer = Buffer.from(svgString, "utf8");

// sharp で PNG に変換（Next.js のビルドでは実行されず、構文だけ有効ならOK）
sharp(svgBuffer)
  .png()
  .toFile(output)
  .then(() => {
    console.log("✅ 画像を書き出しました:", output);
  })
  .catch((err) => {
    console.error("❌ 画像生成エラー:", err);
  });
