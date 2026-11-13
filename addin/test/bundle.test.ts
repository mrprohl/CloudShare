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

  it("registers both handlers", () => {
    expect(code).toContain('"onMessageSendHandler"');
    expect(code).toContain('"onAttachmentsChangedHandler"');
  });
});
