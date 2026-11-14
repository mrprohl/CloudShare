import {
  api,
  composeItem,
  getConfigCached,
  getFileAttachments,
  getToken,
  MB,
  API_BASE,
  CONFIG_CACHE_KEY,
  NOTIFICATION_ID,
  officeAsync,
  readDraftState,
  releaseDraftIfLinksGone,
  tagDraft,
  TASKPANE_COMMAND_ID,
  TIMEOUTS,
  withTimeout,
  type ServerConfig,
} from "../shared/office";
import { t } from "../shared/i18n";
import { errorText } from "../shared/sso";

interface LaunchEvent {
  completed(options?: {
    allowEvent?: boolean;
    errorMessage?: string;
    cancelLabel?: string;
    commandId?: string;
    sendModeOverride?: string;
  }): void;
}
type CompletedOptions = NonNullable<Parameters<LaunchEvent["completed"]>[0]>;

/**
 * Outlook's limits for event.completed (Smart Alerts): button text at most 20 characters, message at most 500.
 * Outlook on the web rejects anything longer and leaves the send dialog spinning.
 */
const alertButton = () => t("alert.button");
const MAX_LABEL = 20;
const MAX_MESSAGE = 500;

/** Keeps every completion within Outlook's limits, whatever the handler built. */
function withinLimits(options: CompletedOptions): CompletedOptions {
  const result = { ...options };
  if (result.errorMessage && result.errorMessage.length > MAX_MESSAGE) {
    result.errorMessage = result.errorMessage.slice(0, MAX_MESSAGE - 2) + " …";
  }
  if (result.cancelLabel && result.cancelLabel.length > MAX_LABEL) result.cancelLabel = result.cancelLabel.slice(0, MAX_LABEL);
  return result;
}

// --- Diagnostics (CLIENT_TRACE=true on the server) ----------------------------------------------------------------
// Reports each step of the send handler to POST /api/trace, so the server log shows where a client hangs.
// Fire-and-forget: never waits, never throws. Off unless the server configuration says otherwise.

let traceOn = false;
let traceRun = "";
let traceStart = 0;

function startTrace() {
  traceRun = Math.random().toString(36).slice(2, 10);
  traceStart = Date.now();
  try {
    const cached = Office.context.roamingSettings.get(CONFIG_CACHE_KEY) as ServerConfig | undefined;
    traceOn = cached?.trace === true;
  } catch {
    traceOn = false;
  }
}

function trace(step: string, detail?: string) {
  if (!traceOn) return;
  try {
    const diagnostics = Office.context.diagnostics;
    const host = diagnostics ? `${diagnostics.platform} ${diagnostics.version}` : "?";
    // Origin of the frame the handler runs in ("null" = sandboxed): decides whether POST requests are allowed.
    const origin = typeof location !== "undefined" ? location.origin : "none";
    const query = [
      ["run", traceRun],
      ["step", step],
      ["ms", String(Date.now() - traceStart)],
      ["host", host],
      ["origin", origin],
      ["detail", detail ?? ""],
    ]
      .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
      .join("&");
    // A plain GET without headers: no preflight, so it also gets through where POST requests are blocked.
    fetch(`${API_BASE}/api/trace?${query}`).catch(() => undefined);
  } catch {
    // Diagnostics must never disturb sending.
  }
}

/** Custom Smart Alert buttons and sendModeOverride need Mailbox 1.14. */
const supportsCustomAlert = () => Office.context.requirements.isSetSupported("Mailbox", "1.14");

const largeAttachments = async (config: ServerConfig) =>
  (await getFileAttachments()).filter((a) => a.size > config.thresholdMb * MB);

/**
 * OnMessageAttachmentsChanged: Outlook does not allow add-ins to open the taskpane themselves.
 * Instead an info bar with the button "Mit CloudShare senden" appears above the mail.
 */
async function onAttachmentsChangedHandler(event: LaunchEvent) {
  const messages = composeItem().notificationMessages;
  try {
    const config = await getConfigCached();
    const large = await largeAttachments(config);
    if (large.length > 0) {
      const action = t(config.attachmentPolicy === "enforce" ? "notify.enforce" : "notify.suggest");
      let message = t("notify.large", { count: large.length, name: large[0].name, mb: config.thresholdMb, action });
      // Notification messages are limited to 150 characters.
      if (message.length > 150) message = t("notify.largeShort", { mb: config.thresholdMb, action });
      if (message.length > 150) message = message.slice(0, 149) + "…";
      await officeAsync<void>((cb) =>
        messages.replaceAsync(
          NOTIFICATION_ID,
          {
            type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
            message,
            icon: "Icon.16",
            actions: [
              {
                actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
                actionText: t("notify.action"),
                commandId: TASKPANE_COMMAND_ID,
                contextData: "{}",
              },
            ],
          },
          cb,
        ),
      );
    } else {
      await officeAsync<void>((cb) => messages.removeAsync(NOTIFICATION_ID, cb)).catch(() => undefined);
    }
  } catch {
    // The hint is optional – errors here must not disturb composing.
  }
  event.completed();
}

/**
 * Arms the password mail for the links CloudShare inserted into this draft. Nothing is sent yet: the server waits
 * until the tagged mail shows up in "Sent Items" and then sends the password to its actual recipients.
 * If the user cancels or sending fails, no password mail goes out and the job stays open for further files.
 */
async function armPasswordMail(jobId: string, passwordTo?: string, passwordSms?: string) {
  // Normally already done by the taskpane – make sure the sent copy carries the tag.
  trace("arm:tag");
  await withTimeout(tagDraft(jobId), TIMEOUTS.office, t("step.tagDraft"));
  trace("arm:from");
  const from = await withTimeout(sharedSender(), TIMEOUTS.office, t("step.readFrom"));
  trace("arm:token");
  const token = await withTimeout(getToken(), TIMEOUTS.token, t("step.signIn"));
  trace("arm:server");
  await withTimeout(
    api(
      `/api/jobs/${encodeURIComponent(jobId)}/arm`,
      // passwordTo: normally already on the job (taskpane) – passed again in case saving it failed there.
      {
        method: "POST",
        body: JSON.stringify({
          ...(from ? { from } : {}),
          ...(passwordSms ? { passwordSms } : passwordTo ? { passwordTo } : {}),
        }),
      },
      token,
    ),
    TIMEOUTS.server,
    t("step.server"),
  );
  trace("arm:done");
}

/**
 * Address in the "From" field if it is not the user's own mailbox (shared mailbox, send as / on behalf).
 * The server then also searches that mailbox's "Sent Items" and sends the password mail from the same address.
 */
async function sharedSender(): Promise<string | undefined> {
  try {
    const from = await officeAsync<Office.EmailAddressDetails>((cb) => composeItem().from.getAsync(cb));
    const own = Office.context.mailbox.userProfile.emailAddress;
    const address = from?.emailAddress;
    return address && address.toLowerCase() !== own.toLowerCase() ? address : undefined;
  } catch {
    return undefined; // "From" not readable – treat as the user's own mailbox
  }
}

const shorten = (names: string) => (names.length > 200 ? names.slice(0, 200) + " …" : names);

/**
 * Smart Alert on send. The manifest uses SendMode SoftBlock, so a blocked mail cannot be sent without fixing it.
 *   enforce: large attachments block sending – no "Send anyway".
 *   suggest: SoftBlock is relaxed to PromptUser via sendModeOverride, which adds "Send anyway".
 * If the add-in cannot run at all, SoftBlock lets the mail through (otherwise a server outage would stop all mail).
 */
async function onMessageSendHandler(event: LaunchEvent) {
  // Every step has a time limit: a step that never answers would leave Outlook at "takes longer than expected".
  trace("start");
  const state = await withTimeout(readDraftState(), TIMEOUTS.office, t("step.readState"));
  trace("state", state.jobId ? "job" : state.noPassword ? "no-password" : "none");

  let config: ServerConfig | undefined;
  let large: Office.AttachmentDetailsCompose[] = [];
  try {
    config = await getConfigCached();
    traceOn = config.trace === true;
    trace("config", config.attachmentPolicy);
    large = await withTimeout(largeAttachments(config), TIMEOUTS.office, t("step.readAttachments"));
  } catch {
    // No configuration known (server never reached) → no attachment check.
  }
  trace("attachments", String(large.length));
  const mustAct = config !== undefined && large.length > 0;
  const names = shorten(large.map((a) => a.name).join(", "));

  // enforce: block first – the handler runs again on the next send attempt.
  if (mustAct && config!.attachmentPolicy === "enforce") {
    const options: CompletedOptions = {
      allowEvent: false,
      errorMessage:
        t("alert.enforce", { mb: config!.thresholdMb, names }) +
        " " +
        (supportsCustomAlert() ? t("alert.chooseButton", { button: alertButton() }) : t("alert.useRibbon")),
    };
    if (supportsCustomAlert()) {
      options.cancelLabel = alertButton();
      options.commandId = TASKPANE_COMMAND_ID;
    }
    event.completed(options);
    return;
  }

  // Arm the password mail for links already inserted. Under "suggest" this must happen before the prompt,
  // because "Send anyway" sends the mail without calling this handler again. Arming is harmless if the user
  // cancels: the password mail only goes out once the mail appears in "Sent Items".
  // Links deleted from the mail again: no password mail for a mail without links.
  const released =
    state.jobId !== undefined &&
    (await withTimeout(releaseDraftIfLinksGone(), TIMEOUTS.office, t("step.readBody")).catch(() => false));
  if (state.jobId && !released) {
    try {
      await armPasswordMail(state.jobId, state.passwordTo, state.passwordSms);
    } catch (err) {
      trace("arm:failed", errorText(err));
      event.completed({
        allowEvent: false,
        errorMessage: t("alert.armFailed", { error: errorText(err) }),
      });
      return;
    }
  }

  if (mustAct) {
    // suggest: prompt with "Send anyway". Older clients cannot relax SoftBlock, so they are not blocked.
    if (supportsCustomAlert()) {
      event.completed({
        allowEvent: false,
        errorMessage: t("alert.suggest", { mb: config!.thresholdMb, names }),
        cancelLabel: alertButton(),
        commandId: TASKPANE_COMMAND_ID,
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
      });
      return;
    }
  }

  event.completed({ allowEvent: true });
}

// Anything unexpected still completes the event with a message instead of leaving the send dialog hanging.
function sendHandler(event: LaunchEvent) {
  startTrace();
  const safe: LaunchEvent = {
    completed: (options) => {
      trace("completed", options?.allowEvent ? "allow" : (options?.errorMessage ?? "block").slice(0, 200));
      event.completed(options ? withinLimits(options) : options);
    },
  };
  return onMessageSendHandler(safe).catch((err) =>
    safe.completed({ allowEvent: false, errorMessage: t("alert.checkFailed", { error: errorText(err) }) }),
  );
}

function registerHandlers() {
  Office.actions.associate("onAttachmentsChangedHandler", onAttachmentsChangedHandler);
  Office.actions.associate("onMessageSendHandler", sendHandler);
}

// Classic Outlook (JavaScript-only runtime) has Office.js preloaded: register right away.
// Outlook on the web / new Outlook (browser runtime): office.js loads the Outlook API (Office.actions) only after this
// script ran – registering now fails, and Outlook would wait for a handler that never exists ("CloudShare dauert
// länger als erwartet"). So register again once office.js is ready; associating twice just replaces the mapping.
try {
  registerHandlers();
} catch {
  // Office.actions not there yet – see Office.onReady below.
}
try {
  void Office.onReady(() => {
    try {
      registerHandlers();
    } catch {
      // Nothing more to do – the event cannot be handled in this runtime.
    }
  });
} catch {
  // No Office.onReady in this runtime (classic Outlook ignores it).
}
