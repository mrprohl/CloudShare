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
  for (const m of serverSources.matchAll(/\b(?:number|envFlag|envRegex|required|sendDays|oneOf|webUrl)(?:<[^>(]*>)?\(\s*"([A-Z][A-Z0-9_]+)"/g)) {
    names.add(m[1]);
  }
  return [...names].sort();
}

/** Files the image and the stack read settings from besides the server (Bitwarden CLI login, Caddy, Compose). */
const stackSources = ["docker/entrypoint.sh", "compose.yml", "docker-compose.yml", "Dockerfile", "caddy/Caddyfile", "versions.env.example"]
  .filter((file) => existsSync(path.join(repo, file)))
  .map(read)
  .join("\n");

/** Settings the docs name: the first column of the README settings table and every line VAR= / #VAR= of the example. */
function documentedSettings(): Map<string, string> {
  const found = new Map<string, string>();
  for (const m of read("README.md").matchAll(/^\| ((?:`[A-Z][A-Z0-9_]+`(?: \/ )?)+) \|/gm)) {
    for (const name of m[1].matchAll(/`([A-Z][A-Z0-9_]+)`/g)) found.set(name[1], "README.md");
  }
  for (const m of read("cloudshare.env.example").matchAll(/^#?([A-Z][A-Z0-9_]+)=/gm)) found.set(m[1], "cloudshare.env.example");
  return found;
}

/** Defaults the server uses: number("X", 14) and oneOf("X", …, "warn") with a literal default. */
function serverDefaults(): Map<string, string> {
  const config = read("server/src/config.ts");
  const defaults = new Map<string, string>();
  for (const m of config.matchAll(/\bnumber\(\s*"([A-Z][A-Z0-9_]+)",\s*(\d+)[,)]/g)) defaults.set(m[1], m[2]);
  for (const m of config.matchAll(/\boneOf(?:<[^>(]*>)?\(\s*"([A-Z][A-Z0-9_]+)",\s*[\w.]+,\s*"([^"]+)"\)/g)) defaults.set(m[1], m[2]);
  return defaults;
}

/** Every text file of the repository (no dependencies, builds, pictures, local settings). */
function textFiles(dir = ""): string[] {
  const skip = new Set(["node_modules", "dist", ".git", "images", "test-results", "playwright-report", "env"]);
  return readdirSync(path.join(repo, dir), { withFileTypes: true }).flatMap((entry) => {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) return skip.has(entry.name) ? [] : textFiles(file);
    return /\.(ts|mjs|js|json|md|html|css|sh|yml|yaml|example|xml|txt)$|^Dockerfile$|^Caddyfile$/.test(entry.name) ? [file] : [];
  });
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

  it("names no setting that nothing reads any more (a removed variable stays out of the docs)", () => {
    const known = new Set(environmentVariables());
    const stale = [...documentedSettings()]
      .filter(([name]) => !known.has(name) && !new RegExp(`\\b${name}\\b`).test(stackSources))
      .map(([name, file]) => `${name} (${file})`);
    expect(stale, "remove these from the docs, or read them").toEqual([]);
  });

  it("states the defaults the server really uses – in the README table, the example and its comments", () => {
    const defaults = serverDefaults();
    expect(defaults.size).toBeGreaterThan(8);
    const wrong: string[] = [];
    // README: "| `A` / `B` | `14` / `90` |" – the values in the order of the names.
    for (const m of read("README.md").matchAll(/^\| ((?:`[A-Z][A-Z0-9_]+`(?: \/ )?)+) \| ((?:`[^`|]*`(?: \/ )?)+) \|/gm)) {
      const names = [...m[1].matchAll(/`([A-Z][A-Z0-9_]+)`/g)].map((x) => x[1]);
      const values = [...m[2].matchAll(/`([^`]*)`/g)].map((x) => x[1]);
      if (names.length !== values.length) continue;
      names.forEach((name, i) => {
        if (defaults.has(name) && defaults.get(name) !== values[i]) {
          wrong.push(`README.md: ${name} says ${values[i]}, the server uses ${defaults.get(name)}`);
        }
      });
    }
    // Example: a commented line shows the default; a comment right above it that says "default N" must say the same.
    const lines = read("cloudshare.env.example").split(/\r?\n/);
    lines.forEach((line, i) => {
      const m = /^(#?)([A-Z][A-Z0-9_]+)=['"]?([^'"\s]*)/.exec(line);
      if (!m || !defaults.has(m[2])) return;
      const expected = defaults.get(m[2])!;
      if (m[1] === "#" && m[3] && m[3] !== expected) wrong.push(`cloudshare.env.example: #${m[2]}=${m[3]}, the server uses ${expected}`);
      const said = /\bdefault:? (\w+)/i.exec(lines[i - 1] ?? "");
      if (said && /^\d+$/.test(said[1]) && said[1] !== expected) {
        wrong.push(`cloudshare.env.example: the comment above ${m[2]} says default ${said[1]}, the server uses ${expected}`);
      }
    });
    expect(wrong).toEqual([]);
  });

  it("keeps every text file valid UTF-8 without doubly encoded characters", () => {
    const files = textFiles();
    expect(files.length).toBeGreaterThan(50);
    const decoder = new TextDecoder("utf-8", { fatal: true });
    const broken: string[] = [];
    for (const file of files) {
      let text: string;
      try {
        text = decoder.decode(readFileSync(path.join(repo, file)));
      } catch {
        broken.push(`${file}: not UTF-8`);
        continue;
      }
      // An umlaut or a dash written as two or three odd characters: UTF-8 read as Latin-1 and written again.
      if (/\u00c3[\u0080-\u00bf]|\u00e2\u20ac/.test(text)) broken.push(`${file}: doubly encoded characters`);
    }
    expect(broken).toEqual([]);
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
