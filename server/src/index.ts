import { buildApp } from "./app.js";
import { config } from "./config.js";
import { startJobWatcher } from "./jobs.js";
import { initVault } from "./vaultwarden.js";

const app = await buildApp();
await initVault((msg) => app.log.info(msg));
await app.listen({ host: "0.0.0.0", port: config.port });
startJobWatcher(app.log);
