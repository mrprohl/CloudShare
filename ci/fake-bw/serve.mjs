// CI only – mock of the "bw serve" Vault Management API on 127.0.0.1:8087 for the stack smoke and load tests.
// Answers just what server/src/vaultwarden.ts uses: POST /unlock, POST /object/send, GET/PUT/DELETE /object/send/:id.
// Like Vaultwarden it refuses a Send whose deletion date is more than 31 days ahead, and – CloudShare's own rule
// (the recipients' access ends at the latest with the Send) – one that expires after its deletion, so a regression
// in the dates fails the stack tests instead of only the real server.
import { createServer } from "node:http";

const DAY = 86_400_000;

const reply = (res, status, body) => {
  res.writeHead(status, { "Content-Type": "application/json" });
  res.end(JSON.stringify(body));
};

/** Why Vaultwarden (or CloudShare's rule) would refuse this Send, or null. */
function refusal(send) {
  const deletion = Date.parse(send.deletionDate ?? "");
  if (!Number.isFinite(deletion)) return "A deletion date is required.";
  if (deletion > Date.now() + 31 * DAY) {
    return "You cannot have a Send with a deletion date that far into the future. Adjust the Deletion Date to a value less than 31 days from now and try again.";
  }
  const expiration = send.expirationDate ? Date.parse(send.expirationDate) : NaN;
  if (Number.isFinite(expiration) && expiration > deletion) return "fake bw serve: expiration date after the deletion date";
  return null;
}

/** Reads the JSON body, then answers with `handle(body)`. */
const withBody = (req, handle) => {
  let body = "";
  req.on("data", (chunk) => (body += chunk));
  req.on("end", () => handle(JSON.parse(body || "{}")));
};

const sends = new Map();

createServer((req, res) => {
  if (req.method === "POST" && req.url === "/unlock") return reply(res, 200, { success: true, data: {} });
  if (req.method === "POST" && req.url === "/object/send") {
    return withBody(req, (body) => {
      const problem = refusal(body);
      if (problem) return reply(res, 400, { success: false, message: problem });
      const id = `smoke-${sends.size + 1}`;
      const send = { ...body, id, object: "send", accessUrl: `https://vault.example.com/#/send/${id}` };
      sends.set(id, send);
      reply(res, 200, { success: true, data: send });
    });
  }
  const match = /^\/object\/send\/([^/]+)$/.exec(req.url ?? "");
  if (match && sends.has(match[1])) {
    if (req.method === "GET") return reply(res, 200, { success: true, data: sends.get(match[1]) });
    if (req.method === "DELETE") {
      sends.delete(match[1]);
      return reply(res, 200, { success: true });
    }
    if (req.method === "PUT") {
      return withBody(req, (body) => {
        const send = { ...sends.get(match[1]), ...body };
        const problem = refusal(send);
        if (problem) return reply(res, 400, { success: false, message: problem });
        sends.set(match[1], send);
        reply(res, 200, { success: true, data: send });
      });
    }
  }
  reply(res, 404, { success: false, message: `fake bw serve: ${req.method} ${req.url} not supported` });
}).listen(8087, "127.0.0.1");
