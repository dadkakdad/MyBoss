import { createServer } from "node:http";
import { dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { createHandler } from "./server/app.mjs";
import { createPgStore } from "./server/pg-store.mjs";

const connectionString = process.env.DATABASE_URL || "";
if (!connectionString) {
  console.error("DATABASE_URL is required");
  process.exit(1);
}

const store = await createPgStore(connectionString);
const webRoot = dirname(fileURLToPath(import.meta.url));
const port = Number(process.env.PORT || 80);
const server = createServer(createHandler({
  store,
  webRoot,
  cookieSecure: (req) => req.headers["x-forwarded-proto"] === "https",
}));

server.listen(port, "0.0.0.0", () => {
  console.log(`myboss listening on ${port}`);
});
