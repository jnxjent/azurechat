/** @type {import('next').NextConfig} */
const nextConfig = {
  output: "standalone",

  experimental: {
    outputFileTracingIncludes: {
      "/*": [
        "./node_modules/pdfjs-dist/legacy/build/pdf.worker.js",
        "./src/scripts/**/*.py",
        "./scripts/**/*.py",
      ],
    },
    serverComponentsExternalPackages: [
      "@azure/storage-blob",
      "@napi-rs/canvas",
      "pdfjs-dist",
      "xlsx",
    ],
  },

  images: {
    imageSizes: [16, 32, 48, 64, 96, 128, 256, 384],
    remotePatterns: [
      {
        protocol: "https",
        hostname: "midac19-webapp-yhggrda5qr5ae.azurewebsites.net",
        pathname: "/api/images/**",
      },
      // {
      //   protocol: "https",
      //   hostname: "azurechat-gpt5-test.azurewebsites.net",
      //   pathname: "/api/images/**",
      // },
    ],
  },

  webpack: (config, { isServer }) => {
    config.module.rules.push({
      test: /\.node$/,
      use: "node-loader",
    });

    return config;
  },
};

module.exports = nextConfig;
