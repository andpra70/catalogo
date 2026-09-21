import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), "");
  const backendUrl = env.VITE_BACKEND_PROXY_URL || "https://127.0.0.1:8443";
  const backendHost = env.VITE_BACKEND_HOST || "belle.iliadboxos.it";
  const backendProxy = {
    target: backendUrl,
    changeOrigin: true,
    secure: false,
    headers: { host: backendHost },
  };
  return {
    base: env.VITE_APP_BASE || "/",
    plugins: [react()],
    server: {
      host: '0.0.0.0',
      allowedHosts: true,
      proxy: {
        "/auth": { ...backendProxy },
        "/vfs": { ...backendProxy },
        "/gpdr": { ...backendProxy },
      },
    },
    preview: {
      host: '0.0.0.0',
      allowedHosts: true,
    },
  };
});
