import { defineConfig } from "vite";

/**
 * Vite config cho E2E test taskpane.
 * Chạy ở port 3003 (khác với dev port 3000 để không xung đột).
 * Bundle file test-taskpane.ts thành JS để Word taskpane iframe load được.
 */
export default defineConfig({
  root: __dirname,
  publicDir: false,
  server: {
    port: 3003,
    https: {
      // Self-signed cert từ office-addin-dev-certs (npm run certs)
      cert: undefined, // sẽ load runtime
      key: undefined,
    },
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      input: "src/test-taskpane.ts",
    },
  },
  resolve: {
    alias: {
      "@": "../public", // cho phép import từ public/modules
    },
  },
});
