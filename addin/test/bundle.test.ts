/**
 * The send handler runs in classic Outlook's JavaScript-only runtime – an older engine without DOM. If the bundle
 * uses something it lacks, the handler never starts and Outlook only shows "takes longer than expected".
 * This builds commands.js like build.mjs does and looks for the usual suspects.
 */
import { build } from "esbuild";
import { describe, expect, it } from "vitest";

const { outputFiles } = await build({
  entryPoints: ["src/commands/commands.ts"],
  bundle: true,
  format: "iife",
  target: "es2016",
  minify: false,
  write: false,
});
const code = outputFiles[0].text;

describe("commands.js for the JavaScript-only runtime", () => {
  it.each([
    ["async functions", /\basync\s+(function\b|\()/],
    ["await", /\bawait\s/],
    ["optional chaining", /\?\.[A-Za-z_$[(]/],
    ["nullish coalescing", /\?\?/],
    ["Promise.prototype.finally", /\.finally\(/],
    ["Headers", /\bnew Headers\b/],
    ["DOM access", /\b(document|window|localStorage)\./],
    ["ES module syntax", /^\s*(import|export)\s/m],
  ])("contains no %s", (_name, pattern) => {
    expect(code).not.toMatch(pattern);
  });

  it("keeps the marker the server replaces with the texts of the configured language", () => {
    expect(code).toContain('"%%CLOUDSHARE_I18N%%"');
  });

  it("registers both handlers", () => {
    expect(code).toContain('"onMessageSendHandler"');
    expect(code).toContain('"onAttachmentsChangedHandler"');
  });
});

describe("language injection into the minified bundles", () => {
  it.each(["src/commands/commands.ts", "src/taskpane/taskpane.ts"])("%s keeps a replaceable marker", async (entry) => {
    const { outputFiles: [bundle] } = await build({
      entryPoints: [entry],
      bundle: true,
      format: "iife",
      target: entry.includes("commands") ? "es2016" : "es2017",
      minify: true,
      write: false,
    });
    const { loadLocale, localize } = await import("../scripts/localize.mjs");
    const localized: string = localize(bundle.text, ".js", loadLocale("en"), "en");
    expect(localized).not.toContain("%%CLOUDSHARE_I18N%%");
    expect(localized).toContain("Open CloudShare");
    expect(() => new Function(localized)).not.toThrow(); // still valid JavaScript
  });
});
