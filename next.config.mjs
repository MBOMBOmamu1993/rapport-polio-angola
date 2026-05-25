/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  poweredByHeader: false,
  webpack: (config, { webpack: wp, isServer }) => {
    if (!isServer) {
      // pptxgenjs référence des modules Node (node:fs, node:https…) inutiles côté
      // navigateur : on retire le préfixe "node:" puis on les neutralise.
      config.plugins.push(
        new wp.NormalModuleReplacementPlugin(/^node:/, (resource) => {
          resource.request = resource.request.replace(/^node:/, "");
        })
      );
      config.resolve.fallback = {
        ...config.resolve.fallback,
        fs: false,
        https: false,
        os: false,
        path: false,
        "image-size": false,
      };
    }
    return config;
  },
};

export default nextConfig;
