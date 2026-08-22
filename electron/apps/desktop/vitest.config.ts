import { defineConfig } from "vitest/config";
import { resolve } from "node:path";

export default defineConfig({
  test: {
    include: ["tests/**/*.test.ts"],
    environment: "node",
  },
  resolve: {
    alias: {
      "@rpd/shared": resolve(__dirname, "../../packages/shared/src/index.ts"),
    },
  },
});
