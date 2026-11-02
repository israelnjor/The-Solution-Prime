import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  plugins: [react()],
  // Workspace packages are consumed from their built dist/; their CommonJS deps need pre-bundling.
  // pdf.js is an ES module that starts its own worker: used as published, not pre-bundled.
  optimizeDeps: { include: ["@xmldom/xmldom", "fflate"], exclude: ["pdfjs-dist"] },
  // Dev only: allow ?file=<absolute path> to load a document from the repository for testing.
  // The Rust build output (src-tauri/target) must not be watched: it is large and its files are locked while compiling.
  server: { fs: { allow: ["../.."] }, watch: { ignored: ["**/src-tauri/**"] } },
  clearScreen: false,
});
