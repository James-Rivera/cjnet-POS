import type { NextConfig } from "next";
import path from "node:path";

const projectRoot = path.resolve(__dirname);
const isProduction = process.env.NODE_ENV === "production";

const nextConfig: NextConfig = {
  reactCompiler: isProduction,
  outputFileTracingRoot: projectRoot,
  allowedDevOrigins: ["100.88.66.93"],
  turbopack: {
    root: projectRoot,
  },
};

export default nextConfig;
