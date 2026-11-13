// UI tests only – serves the built add-in (addin/dist) like the container does: placeholders replaced, so the
// taskpane talks to this origin. /api/* is answered by the tests (Playwright route mocks), not by this server.
import { readFile } from "node:fs/promises";
import { createServer } from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

const PORT = 4173;
const ORIGIN = `http://127.0.0.1:${PORT}`;
const dist = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../dist");
const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".xml": "application/xml" };

createServer(async (req, res) => {
  const file = path.join(dist, path.normalize(decodeURIComponent(new URL(req.url, ORIGIN).pathname)));
  if (!file.startsWith(dist)) return res.writeHead(403).end();
  try {
    const ext = path.extname(file);
    let body = await readFile(file);
    if ([".html", ".js", ".xml"].includes(ext)) {
      body = body
        .toString()
        .replaceAll("%%PUBLIC_BASE_URL%%", ORIGIN)
        .replaceAll("%%PUBLIC_HOST%%", `127.0.0.1:${PORT}`)
        .replaceAll("%%ENTRA_CLIENT_ID%%", "00000000-0000-0000-0000-000000000002");
    }
    res.writeHead(200, { "Content-Type": types[ext] ?? "application/octet-stream" }).end(body);
  } catch {
    res.writeHead(404).end();
  }
}).listen(PORT, "127.0.0.1");
