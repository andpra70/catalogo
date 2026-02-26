import { defineConfig } from "vite";

export default defineConfig({
  server: {
    allowedHosts: ["zanotti.iliadboxos.it"],
  },
  preview: {
    allowedHosts: ["zanotti.iliadboxos.it"],
  },
});
