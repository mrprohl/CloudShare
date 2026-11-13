import {
  api,
  composeItem,
  getConfigCached,
  getFileAttachments,
  getToken,
  MB,
  NOTIFICATION_ID,
  officeAsync,
  readDraftState,
  tagDraft,
  TASKPANE_COMMAND_ID,
  TIMEOUTS,
  withTimeout,
  type DraftState,
  type ServerConfig,
} from "../shared/office";
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

/** Custom Smart Alert buttons and sendModeOverride need Mailbox 1.14. */
const supportsCustomAlert = () => Office.context.requirements.isSetSupported("Mailbox", "1.14");

const largeAttachments = async (config: ServerConfig) =>
  (await getFileAttachments()).filter((a) => a.size > config.thresholdMb * MB);

/** Whether CloudShare has to step in: always under "enforce", under "suggest" unless the user opted out. */
const applies = (config: ServerConfig, state: DraftState) => config.attachmentPolicy === "enforce" || !state.optOut;

/**
 * OnMessageAttachmentsChanged: Outlook does not allow add-ins to open the taskpane themselves.
 * Instead an info bar with the button "Mit CloudShare senden" appears above the mail (UI text is German).
 */
async function onAttachmentsChangedHandler(event: LaunchEvent) {
  const messages = composeItem().notificationMessages;
  try {
    const [config, state] = await Promise.all([getConfigCached(), readDraftState()]);
    const large = await largeAttachments(config);
    if (large.length > 0 && applies(config, state)) {
      const what = large.length === 1 ? `„${large[0].name}“ ist` : `${large.length} Anhänge sind`;
      const tail =
        config.attachmentPolicy === "enforce"
          ? "kann nur als kennwortgeschützter Nextcloud-Link versendet werden."
          : "Als kennwortgeschützten Nextcloud-Link senden?";
      let message = `${what} größer als ${config.thresholdMb} MB – ${tail}`;
      // Notification messages are limited to 150 characters.
      if (message.length > 150) message = `Anhänge größer als ${config.thresholdMb} MB – ${tail}`;
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
                actionText: "Mit CloudShare senden",
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
async function armPasswordMail(jobId: string) {
  // Normally already done by the taskpane – make sure the sent copy carries the tag.
  await withTimeout(tagDraft(jobId), TIMEOUTS.office, "Outlook (Markieren des Entwurfs)");
  const from = await withTimeout(sharedSender(), TIMEOUTS.office, "Outlook (Absender lesen)");
  const token = await withTimeout(getToken(), TIMEOUTS.token, "Die Anmeldung bei Microsoft (SSO)");
  await withTimeout(
    api(`/api/jobs/${encodeURIComponent(jobId)}/arm`, { method: "POST", body: JSON.stringify(from ? { from } : {}) }, token),
    TIMEOUTS.server,
    "Der CloudShare-Server",
  );
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
  const state = await withTimeout(readDraftState(), TIMEOUTS.office, "Outlook (Entwurfsdaten lesen)");

  let config: ServerConfig | undefined;
  let large: Office.AttachmentDetailsCompose[] = [];
  try {
    config = await getConfigCached();
    large = await withTimeout(largeAttachments(config), TIMEOUTS.office, "Outlook (Anhänge lesen)");
  } catch {
    // No configuration known (server never reached) → no attachment check.
  }
  const mustAct = config !== undefined && large.length > 0 && applies(config, state);
  const names = shorten(large.map((a) => a.name).join(", "));

  // enforce: block first – the handler runs again on the next send attempt.
  if (mustAct && config!.attachmentPolicy === "enforce") {
    const options: CompletedOptions = {
      allowEvent: false,
      errorMessage:
        `Anhänge über ${config!.thresholdMb} MB dürfen nur mit CloudShare als kennwortgeschützter Nextcloud-Link ` +
        `versendet werden: ${names}. ` +
        (supportsCustomAlert()
          ? `Bitte „Mit CloudShare senden“ wählen.`
          : `Bitte über „Dateien oder Nachrichten sicher versenden oder abrufen“ in der Menüleiste versenden.`),
    };
    if (supportsCustomAlert()) {
      options.cancelLabel = "Mit CloudShare senden";
      options.commandId = TASKPANE_COMMAND_ID;
    }
    event.completed(options);
    return;
  }

  // Arm the password mail for links already inserted. Under "suggest" this must happen before the prompt,
  // because "Send anyway" sends the mail without calling this handler again. Arming is harmless if the user
  // cancels: the password mail only goes out once the mail appears in "Sent Items".
  if (state.jobId) {
    try {
      await armPasswordMail(state.jobId);
    } catch (err) {
      event.completed({
        allowEvent: false,
        errorMessage: `CloudShare kann das Kennwort für diese Nachricht nicht vorbereiten: ${errorText(err)}`,
      });
      return;
    }
  }

  if (mustAct) {
    // suggest: prompt with "Send anyway". Older clients cannot relax SoftBlock, so they are not blocked.
    if (supportsCustomAlert()) {
      event.completed({
        allowEvent: false,
        errorMessage:
          `Anhänge über ${config!.thresholdMb} MB: ${names}. ` +
          `Mit CloudShare werden sie als kennwortgeschützter Nextcloud-Link gesendet. ` +
          `Mit „Trotzdem senden“ werden sie als normale Anhänge gesendet.`,
        cancelLabel: "Mit CloudShare senden",
        commandId: TASKPANE_COMMAND_ID,
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
      });
      return;
    }
  }

  event.completed({ allowEvent: true });
}

Office.actions.associate("onAttachmentsChangedHandler", onAttachmentsChangedHandler);
// Anything unexpected still completes the event with a message instead of leaving the send dialog hanging.
Office.actions.associate("onMessageSendHandler", (event: LaunchEvent) =>
  onMessageSendHandler(event).catch((err) =>
    event.completed({ allowEvent: false, errorMessage: `CloudShare konnte die Nachricht nicht prüfen: ${errorText(err)}` }),
  ),
);
