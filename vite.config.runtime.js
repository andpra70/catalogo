import { defineConfig } from "vite";

export default defineConfig({
  base: "./",
  server: {
    allowedHosts: ["zanotti.iliadboxos.it"],
  },
  preview: {
    allowedHosts: ["zanotti.iliadboxos.it"],
  },
});
