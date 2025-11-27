/** @type {import('next').NextConfig} */
const nextConfig = {
  output: "standalone",

  experimental: {
    serverComponentsExternalPackages: ["@azure/storage-blob"],
  },

  images: {
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
};

module.exports = nextConfig;
