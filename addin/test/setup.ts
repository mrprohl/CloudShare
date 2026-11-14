/**
 * The bundles get their texts from the server at container start (server/src/i18n.ts). Unit tests import the
 * sources directly, so they load the German texts of locales/de.json here.
 */
import { readFileSync } from "node:fs";
import { setMessages } from "../src/shared/i18n";

const de = JSON.parse(readFileSync(new URL("../../locales/de.json", import.meta.url), "utf8")) as {
  addin: Record<string, string>;
};
setMessages(de.addin, "de-DE");
