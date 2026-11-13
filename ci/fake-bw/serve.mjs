// CI only – mock of the "bw serve" Vault Management API on 127.0.0.1:8087 for the stack smoke test.
// Answers just what server/src/vaultwarden.ts uses: POST /unlock, POST /object/send, GET/PUT /object/send/:id.
import { createServer } from "node:http";

const reply = (res, status, body) => {
  res.writeHead(status, { "Content-Type": "application/json" });
  res.end(JSON.stringify(body));
};

const sends = new Map();

createServer((req, res) => {
  if (req.method === "POST" && req.url === "/unlock") return reply(res, 200, { success: true, data: {} });
  if (req.method === "POST" && req.url === "/object/send") {
    const id = `smoke-${sends.size + 1}`;
    const send = { id, object: "send", accessUrl: `https://vault.example.com/#/send/${id}`, maxAccessCount: 3 };
    sends.set(id, send);
    return reply(res, 200, { success: true, data: send });
  }
  const match = /^\/object\/send\/([^/]+)$/.exec(req.url ?? "");
  if (match && sends.has(match[1])) {
    if (req.method === "GET") return reply(res, 200, { success: true, data: sends.get(match[1]) });
    if (req.method === "PUT") {
      let body = "";
      req.on("data", (chunk) => (body += chunk));
      req.on("end", () => {
        const send = { ...sends.get(match[1]), ...JSON.parse(body || "{}") };
        sends.set(match[1], send);
        reply(res, 200, { success: true, data: send });
      });
      return;
    }
  }
  reply(res, 404, { success: false, message: `fake bw serve: ${req.method} ${req.url} not supported` });
}).listen(8087, "127.0.0.1");
