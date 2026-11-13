/**
 * Minimal Office.js stand-in for the event handlers in src/commands/commands.ts: a compose item with attachments,
 * session data, custom properties, notification messages and a "From" field, plus requirement sets,
 * roamingSettings and Office.actions.associate (which captures the handlers).
 */
import { vi } from "vitest";

export interface Attachment {
  id: string;
  name: string;
  size: number;
  attachmentType: "file" | "item" | "cloud";
  isInline: boolean;
}

export interface OfficeState {
  attachments: Attachment[];
  session: Record<string, string>;
  customProps: Record<string, unknown>;
  roaming: Record<string, unknown>;
  /** Highest supported version per requirement set, e.g. { Mailbox: "1.14" }. */
  supported: Record<string, string>;
  from: string;
  userEmail: string;
  notifications: { replaced: Array<Record<string, unknown>>; removed: string[] };
}

type Handler = (event: { completed: (options?: unknown) => void }) => Promise<void>;

const ok = (value?: unknown) => ({ status: "succeeded", value });
const failed = () => ({ status: "failed", error: new Error("not found") });

/** "1.12" ≤ "1.14" – compares dotted version numbers. */
function atLeast(available: string | undefined, required: string) {
  if (!available) return false;
  const a = available.split(".").map(Number);
  const r = required.split(".").map(Number);
  for (let i = 0; i < Math.max(a.length, r.length); i++) {
    if ((a[i] ?? 0) !== (r[i] ?? 0)) return (a[i] ?? 0) > (r[i] ?? 0);
  }
  return true;
}

export function installOfficeMock() {
  const state: OfficeState = {
    attachments: [],
    session: {},
    customProps: {},
    roaming: {},
    supported: { Mailbox: "1.14" },
    from: "max@example.test",
    userEmail: "max@example.test",
    notifications: { replaced: [], removed: [] },
  };
  const handlers: Record<string, Handler> = {};

  type Callback = (result: unknown) => void;
  const item = {
    getAttachmentsAsync: (cb: Callback) => cb(ok(state.attachments)),
    sessionData: {
      getAsync: (key: string, cb: Callback) => cb(key in state.session ? ok(state.session[key]) : failed()),
      setAsync: (key: string, value: string, cb: Callback) => {
        state.session[key] = value;
        cb(ok());
      },
    },
    loadCustomPropertiesAsync: (cb: Callback) =>
      cb(
        ok({
          get: (key: string) => state.customProps[key],
          set: (key: string, value: unknown) => {
            state.customProps[key] = value;
          },
          saveAsync: (done: Callback) => done(ok()),
        }),
      ),
    notificationMessages: {
      replaceAsync: (id: string, message: Record<string, unknown>, cb: Callback) => {
        state.notifications.replaced.push({ id, ...message });
        cb(ok());
      },
      removeAsync: (id: string, cb: Callback) => {
        state.notifications.removed.push(id);
        cb(ok());
      },
    },
    from: { getAsync: (cb: Callback) => cb(ok({ emailAddress: state.from })) },
  };

  (globalThis as Record<string, unknown>).Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    MailboxEnums: {
      AttachmentType: { File: "file", Item: "item", Cloud: "cloud" },
      ItemNotificationMessageType: { InsightMessage: "insightMessage" },
      ActionType: { ShowTaskPane: "showTaskPane" },
      SendModeOverride: { PromptUser: "promptUser" },
    },
    context: {
      mailbox: {
        item,
        userProfile: {
          get emailAddress() {
            return state.userEmail;
          },
          displayName: "Max Mustermann",
        },
      },
      requirements: { isSetSupported: (name: string, version: string) => atLeast(state.supported[name], version) },
      roamingSettings: {
        get: (key: string) => state.roaming[key],
        set: (key: string, value: unknown) => {
          state.roaming[key] = value;
        },
        saveAsync: () => undefined,
      },
    },
    auth: { getAccessToken: vi.fn(async () => "sso-token") },
    actions: {
      associate: (name: string, fn: Handler) => {
        handlers[name] = fn;
      },
    },
  };

  /** Back to a draft without attachments, state or cache; Mailbox 1.14. */
  function reset() {
    Object.assign(state, {
      attachments: [],
      session: {},
      customProps: {},
      roaming: {},
      supported: { Mailbox: "1.14" },
      from: "max@example.test",
      userEmail: "max@example.test",
      notifications: { replaced: [], removed: [] },
    });
  }

  return { state, handlers, reset };
}
