import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  webpack: (config) => {
    config.watchOptions = {
      ...config.watchOptions,
      ignored: ["**/.git/**", "**/.next/**", "**/node_modules/**", "**/.tools/**", "**/.npm-cache/**"]
    };

    return config;
  }
};

export default nextConfig;
