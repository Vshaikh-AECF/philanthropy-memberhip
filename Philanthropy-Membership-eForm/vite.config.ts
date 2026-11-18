import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { powerApps } from "@microsoft/power-apps-vite/plugin"
import { readFileSync } from "node:fs";

const { version } = JSON.parse(
  readFileSync(new URL("./package.json", import.meta.url), "utf-8"),
) as { version: string };

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), powerApps()],
  define: {
    __APP_VERSION__: JSON.stringify(version),
  },
});
