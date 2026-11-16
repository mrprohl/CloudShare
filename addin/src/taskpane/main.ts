/**
 * Entry of taskpane.js: once Office.js is ready (it knows the client's display language), the page is switched to
 * that language (language.ts) and only then taskpane.ts runs – it binds its elements when it is evaluated. esbuild
 * keeps the dynamic import inside the single IIFE bundle and evaluates the module only here.
 */
import { applyLanguage } from "./language";

Office.onReady(async () => {
  await applyLanguage();
  await import("./taskpane");
});
