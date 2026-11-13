/**
 * Outlook on the web / new Outlook: office.js loads the Outlook API (Office.actions) only after commands.js ran.
 * The handlers must still get registered – otherwise Outlook waits for a send handler that never exists.
 */
import { expect, it } from "vitest";
import { installOfficeMock } from "./officeMock";

it("registers the handlers once office.js is ready, even if Office.actions is missing at load", async () => {
  const office = installOfficeMock();
  const Office = (globalThis as unknown as { Office: Record<string, unknown> }).Office;
  const actions = Office.actions;
  delete Office.actions; // like the browser runtime while office.js is still loading
  let ready: (() => void) | undefined;
  Office.onReady = (callback: () => void) => {
    ready = callback;
    return Promise.resolve();
  };

  await import("../src/commands/commands");
  expect(office.handlers.onMessageSendHandler).toBeUndefined();

  Office.actions = actions; // office.js finished loading
  ready!();
  expect(office.handlers.onMessageSendHandler).toBeTypeOf("function");
  expect(office.handlers.onAttachmentsChangedHandler).toBeTypeOf("function");
});
