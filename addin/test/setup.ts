/**
 * The bundles get their texts from the server at container start (server/src/i18n.ts). Unit tests import the
 * sources directly, so they load the German texts of locales/de.json here – as a JSON module, not with a file read:
 * CodeQL would otherwise follow the texts from that read into the diagnostics request of commands.ts (trace) and
 * report "file data in outbound network request" for production code that never reads a file.
 */
import de from "../../locales/de.json";
import { setMessages } from "../src/shared/i18n";

setMessages((de as { addin: Record<string, string> }).addin, "de-DE");
