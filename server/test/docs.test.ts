/**
 * The documentation keeps up with the code: every environment variable the server reads is described in the README
 * table and in cloudshare.env.example, every HTTP route is in the API table of docs/architecture.md, every README
 * picture exists, and the help in the taskpane has its texts. A failure names what to document – fix the docs, not
 * this test (only truly internal variables belong in INTERNAL, with the reason).
 */
import { existsSync, readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

const repo = path.resolve(import.meta.dirname, "../..");
const read = (file: string) => readFileSync(path.join(repo, file), "utf8");
const serverSources = readdirSync(path.join(repo, "server/src"))
  .filter((f) => f.endsWith(".ts"))
  .map((f) => read(`server/src/${f}`))
  .join("\n");

/** Variables that are set by the image or the tests, not by the admin – not in the operations manual on purpose. */
const INTERNAL: Record<string, string> = {
  ADDIN_ID: "the manifest <Id>; only for own builds with another add-in ID",
  BW_SERVE_URL: "bw serve inside the container (127.0.0.1)",
  CUSTOM_LOCALES_DIR: "fixed by the image (/config/locales, mounted from env/config/locales)",
  DATA_DIR: "fixed by the image (/data volume)",
  LOCALES_DIR: "fixed by the image (/app/locales)",
  TEMPLATES_FILE: "fixed by the image (/config/templates.json, mounted from env/config)",
  PORT: "internal port behind Caddy",
  TZ: "set by compose.yml",
  ENTRA_JWKS_URL: "load test only – never in production",
  SEVEN_URL: "tests only (seven.io mock)",
};

/** Every environment variable the server reads: process.env.X and the helpers of config.ts. */
function environmentVariables(): string[] {
  const names = new Set<string>();
  for (const m of serverSources.matchAll(/process\.env(?:\.|\[")([A-Z][A-Z0-9_]+)/g)) names.add(m[1]);
  for (const m of serverSources.matchAll(/\b(?:number|envFlag|envRegex|required|sendDays)\(\s*"([A-Z][A-Z0-9_]+)"/g)) names.add(m[1]);
  return [...names].sort();
}

/** Every route: app.get/post/put/patch/delete("…") and app.route({ url: "…" }). */
function routes(): string[] {
  const app = read("server/src/app.ts");
  const found = new Set<string>();
  for (const m of app.matchAll(/app\.(get|post|put|patch|delete)(?:<[^(]*?>)?\(\s*"([^"]+)"/gs)) found.add(`${m[1].toUpperCase()} ${m[2]}`);
  for (const m of app.matchAll(/app\.route(?:<[^(]*?>)?\(\{[^}]*?url:\s*"([^"]+)"/gs)) found.add(`ANY ${m[1]}`);
  return [...found].sort();
}

/**
 * Whether the API table names the path: in full, or shortened as "…<rest>" after the first ones (the table writes
 * `PUT …/:id/chunks/:n`). Parameter names may differ (`:index` / `:n`).
 */
function documented(table: string, route: string): boolean {
  const routePath = route.split(" ")[1];
  const escape = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/:\w+/g, ":\\w+");
  const segments = routePath.split("/");
  for (let i = segments.length; i >= 1; i--) {
    const rest = segments.slice(i).join("/");
    const full = escape(segments.slice(0, i).join("/") + (rest ? `/${rest}` : ""));
    if (new RegExp(`${full}(?![\\w/])`).test(table)) return true;
    if (rest && new RegExp(`…/${escape(rest)}(?![\\w/])`).test(table)) return true;
  }
  return false;
}

describe("documentation", () => {
  it("describes every environment variable in the README table and in cloudshare.env.example", () => {
    const readme = read("README.md");
    const example = read("cloudshare.env.example");
    const missing: string[] = [];
    for (const name of environmentVariables()) {
      if (INTERNAL[name]) continue;
      if (!readme.includes(`\`${name}\``)) missing.push(`${name}: README.md (table "Settings")`);
      if (!new RegExp(`\\b${name}\\b`).test(example)) missing.push(`${name}: cloudshare.env.example`);
    }
    expect(missing, "document these settings").toEqual([]);
    // The list of internal ones stays honest: each is still read by the server.
    expect(Object.keys(INTERNAL).filter((name) => !environmentVariables().includes(name))).toEqual([]);
  });

  it("lists every HTTP route in the API table of docs/architecture.md", () => {
    const architecture = read("docs/architecture.md");
    const table = architecture.slice(architecture.indexOf("## 4. HTTP API"), architecture.indexOf("## 5."));
    expect(table.length, "section 4 of architecture.md").toBeGreaterThan(100);
    expect(routes().filter((route) => !documented(table, route)), "add these routes to the table").toEqual([]);
  });

  it("has every picture the README shows", () => {
    const readme = read("README.md");
    const pictures = [...readme.matchAll(/docs\/images\/([\w.-]+\.png)/g)].map((m) => m[1]);
    expect(pictures.length).toBeGreaterThan(5);
    expect(pictures.filter((file) => !existsSync(path.join(repo, "docs/images", file)))).toEqual([]);
  });

  it("names every README picture of the UI tests in the README and the other way round", () => {
    const tests = read("addin/e2e/taskpane.e2e.ts");
    const shots = [...tests.matchAll(/readmeShot\(page, "([\w-]+)"\)/g)].map((m) => `taskpane-${m[1]}.png`).sort();
    const readme = read("README.md");
    const shown = [...new Set([...readme.matchAll(/docs\/images\/(taskpane-[\w-]+\.png)/g)].map((m) => m[1]))].sort();
    expect(shown).toEqual(shots);
  });

  it("explains every section of the help in every language", () => {
    const html = read("addin/src/taskpane/taskpane.html");
    const helpKeys = [...html.matchAll(/%%t:addin\.(help\.[\w.]+)%%/g)].map((m) => m[1]);
    expect(helpKeys.length).toBeGreaterThan(20);
    for (const file of readdirSync(path.join(repo, "locales")).filter((f) => f.endsWith(".json"))) {
      const texts = (JSON.parse(read(`locales/${file}`)) as { addin: Record<string, string> }).addin;
      expect(helpKeys.filter((key) => !texts[key]?.trim()), `${file}: empty help texts`).toEqual([]);
    }
  });
});
