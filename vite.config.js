import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";
import { Proxy } from "@domoinc/ryuu-proxy";
import manifest from "./public/manifest.json";
import tailwindcss from "@tailwindcss/vite";

const config = { manifest };
const proxy = new Proxy(config);

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    react(),
    tailwindcss(),
    {
      name: "ryuu-proxy",
      configureServer(server) {
        server.middlewares.use((req, res, next) => {
          res.status = function (code) {
            this.statusCode = code;
            return this;
          };
          res.send = function (body) {
            this.setHeader("Content-Type", "text/plain");
            this.end(body);
          };
          next();
        });
        // ryuu-proxy grabs every /api/* URL and sends it to Domo (404),
        // so let our own Express routes fall through to server.proxy below.
        const domoProxy = proxy.express();
        server.middlewares.use((req, res, next) =>
          /^\/api\/(requests|send-email)(\/|\?|$)/.test(req.url)
            ? next()
            : domoProxy(req, res, next),
        );
      },
    },
  ],
  // Local dev: forward /api to the Express backend (api/index.js)
  server: {
    proxy: {
      "/api": `http://localhost:${process.env.PORT || 5000}`,
    },
  },
  define: { "process.env": {} },
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
  },
});
