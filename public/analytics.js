import { inject } from "./vendor/vercel-analytics.js";

inject({
  mode: window.location.hostname === "localhost" ? "development" : "production",
});
