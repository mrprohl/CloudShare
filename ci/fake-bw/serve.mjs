// CI only – mock of the "bw serve" Vault Management API on 127.0.0.1:8087 for the stack smoke test.
// Answers just what server/src/vaultwarden.ts uses: POST /unlock and POST /object/send.
import { createServer } from "node:http";

const reply = (res, status, body) => {
  res.writeHead(status, { "Content-Type": "application/json" });
  res.end(JSON.stringify(body));
};

createServer((req, res) => {
  if (req.method === "POST" && req.url === "/unlock") return reply(res, 200, { success: true, data: {} });
  if (req.method === "POST" && req.url === "/object/send") {
    return reply(res, 200, { success: true, data: { accessUrl: "https://vault.example.com/#/send/smoke-test" } });
  }
  reply(res, 404, { success: false, message: `fake bw serve: ${req.method} ${req.url} not supported` });
}).listen(8087, "127.0.0.1");
