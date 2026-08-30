import { defineConfig, loadEnv } from "vite";

export default defineConfig(({ command, mode }) => {
  const env = loadEnv(mode, process.cwd(), "");
  const productionBase = env.VITE_PUBLIC_BASE_PATH || (process.env.GITHUB_ACTIONS ? "/approval-circulation-app/" : "/");
  return {
    base: command === "build" ? productionBase : "/",
    server: {
      host: "127.0.0.1",
      port: 5173,
      proxy: {
        "/api": "http://127.0.0.1:3000",
      },
    },
  };
});
