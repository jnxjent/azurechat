/** @type {import('next').NextConfig} */
const nextConfig = {
  output: "standalone",

  // standalone ビルドに含めるファイルを明示（pdfjs worker と Python スクリプト）
  outputFileTracingIncludes: {
    "/*": [
      "./node_modules/pdfjs-dist/legacy/build/pdf.worker.js",
      "./src/scripts/**/*.py",
      "./scripts/**/*.py",
    ],
  },

  experimental: {
    serverComponentsExternalPackages: [
      "@azure/storage-blob",
      "@napi-rs/canvas",
      "pdfjs-dist",
      "xlsx",
    ],
  },

  images: {
    imageSizes: [16, 32, 48, 64, 96, 128, 256, 384], // ← 追加
    remotePatterns: [
      {
        protocol: "https",
        hostname: "midac19-webapp-yhggrda5qr5ae.azurewebsites.net",
        pathname: "/api/images/**",
      },
      // 必要であれば、test 環境なども後でここに追加できます
      // {
      //   protocol: "https",
      //   hostname: "azurechat-gpt5-test.azurewebsites.net",
      //   pathname: "/api/images/**",
      // },
    ],
  },

  webpack: (config, { isServer }) => {
    // .node ネイティブバイナリ用のローダーを追加
    config.module.rules.push({
      test: /\.node$/,
      use: "node-loader",
    });

    return config;
  },
};

module.exports = nextConfig;
