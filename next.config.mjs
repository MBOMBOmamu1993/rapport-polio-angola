/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  poweredByHeader: false,
  images: {
    // Autorise le rendu des logos SVG internes (public/logo/*.svg) via next/image.
    dangerouslyAllowSVG: true,
    contentDispositionType: "attachment",
    contentSecurityPolicy: "default-src 'self'; script-src 'none'; sandbox;",
  },
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
