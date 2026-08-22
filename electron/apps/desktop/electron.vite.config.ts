import { defineConfig } from "electron-vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

export default defineConfig({
  main: {
    // Bundle workspace + runtime deps into out/main so the packaged app has no node_modules resolution
    // requirements beyond Electron itself (no native modules yet).
    plugins: [],
    resolve: {
      alias: {
        "@rpd/shared": resolve(__dirname, "../../packages/shared/src/index.ts"),
      },
    },
    build: {
      lib: { entry: "src/main/index.ts" },
      rollupOptions: {
        external: ["electron"],
      },
    },
  },
  preload: {
    plugins: [],
    resolve: {
      alias: {
        "@rpd/shared": resolve(__dirname, "../../packages/shared/src/index.ts"),
      },
    },
    build: {
      lib: { entry: "src/preload/index.ts" },
      rollupOptions: {
        external: ["electron"],
      },
    },
  },
  renderer: {
    root: "src/renderer",
    build: {
      rollupOptions: {
        input: resolve(__dirname, "src/renderer/index.html"),
      },
    },
    plugins: [react()],
    resolve: {
      alias: {
        "@": resolve(__dirname, "src/renderer/src"),
        "@rpd/shared": resolve(__dirname, "../../packages/shared/src/index.ts"),
      },
    },
  },
});
