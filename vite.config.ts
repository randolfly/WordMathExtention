import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";
import { getHttpsServerOptions } from "office-addin-dev-certs";

export default defineConfig(async () => {
  const httpsOptions = await getHttpsServerOptions();

  return {
    plugins: [vue()],
    resolve: {
      alias: {
        "@": fileURLToPath(new URL("./src", import.meta.url))
      }
    },
    server: {
      port: 3000,
      strictPort: true,
      https: httpsOptions
    },
    build: {
      rollupOptions: {
        input: {
          taskpane: "taskpane.html",
          commands: "commands.html"
        }
      }
    }
  };
});
