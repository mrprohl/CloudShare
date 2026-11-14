// UI tests only – serves the built add-in (addin/dist) like the container does: placeholders replaced and the
// German texts filled in (scripts/localize.mjs), so the taskpane talks to this origin. Below /en/ the same files come
// with the English texts (README pictures). /api/* is answered by the tests (Playwright route mocks), not by this server.
import { readFile } from "node:fs/promises";
import { createServer } from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { loadLocale, localize } from "../scripts/localize.mjs";

const PORT = 4173;
const ORIGIN = `http://127.0.0.1:${PORT}`;
const dist = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../dist");
const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".xml": "application/xml" };
const locales = { de: loadLocale("de"), en: loadLocale("en") };

createServer(async (req, res) => {
  let pathname = decodeURIComponent(new URL(req.url, ORIGIN).pathname);
  let language = "de";
  const prefixed = /^\/(de|en)(\/.*)$/.exec(pathname);
  if (prefixed) [, language, pathname] = prefixed;
  const file = path.join(dist, path.normalize(pathname));
  if (!file.startsWith(dist)) return res.writeHead(403).end();
  try {
    const ext = path.extname(file);
    let body = await readFile(file);
    if ([".html", ".js", ".xml"].includes(ext)) {
      body = localize(
        body
          .toString()
          .replaceAll("%%PUBLIC_BASE_URL%%", ORIGIN)
          .replaceAll("%%PUBLIC_HOST%%", `127.0.0.1:${PORT}`)
          .replaceAll("%%ENTRA_CLIENT_ID%%", "00000000-0000-0000-0000-000000000002"),
        ext,
        locales[language],
        language,
      );
    }
    res.writeHead(200, { "Content-Type": types[ext] ?? "application/octet-stream" }).end(body);
  } catch {
    res.writeHead(404).end();
  }
}).listen(PORT, "127.0.0.1");
