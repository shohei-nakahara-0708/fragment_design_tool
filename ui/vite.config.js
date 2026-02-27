import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      "/upload": "http://127.0.0.1:8000",
      "/upload-batch": "http://127.0.0.1:8000",
      "/render": "http://127.0.0.1:8000",
      "/jobs": "http://127.0.0.1:8000",
      "/job": "http://127.0.0.1:8000",
      "/preview": "http://127.0.0.1:8000",
      "/download": "http://127.0.0.1:8000",
      "/debug": "http://127.0.0.1:8000",
      "/jobs/export.zip": "http://127.0.0.1:8000",
    },
  },
});
