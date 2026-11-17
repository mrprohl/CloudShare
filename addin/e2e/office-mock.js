// UI tests only – stands in for https://appsforoffice.microsoft.com/lib/1/hosted/office.js in the browser.
// The scenario comes from window.__officeConfig (set by the test before the page loads); everything the taskpane
// does to the mail is recorded in window.__officeLog for the assertions.
(function () {
  const cfg = window.__officeConfig || {};
  const log = (window.__officeLog = {
    inserted: [],
    removed: [],
    customProps: {},
    session: {},
    removedNotifications: [],
    openedUrls: [],
    // Mails opened with displayMessageForm ("Verlauf": click on a subject), as EWS IDs.
    openedMails: [],
    tokenRequests: 0,
  });
  let tokenBusy = false;
  const ok = (value) => ({ status: "succeeded", value });
  const failed = (message) => ({ status: "failed", error: { message } });
  const attachments = (cfg.attachments || []).slice();

  window.Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Html: "html", Text: "text" },
    EventType: { AttachmentsChanged: "olkAttachmentsChanged", OfficeThemeChanged: "officeThemeChanged", ItemChanged: "olkItemSelectedChanged" },
    MailboxEnums: {
      AttachmentType: { File: "file", Item: "item", Cloud: "cloud" },
      AttachmentContentFormat: { Base64: "base64", Url: "url", Eml: "eml", ICalendar: "iCalendar" },
    },
    onReady(callback) {
      const info = { host: "Outlook", platform: "OfficeOnline" };
      setTimeout(() => callback && callback(info), 0);
      return Promise.resolve(info);
    },
    // Like Outlook: failures are plain objects { code, name, message }, not Errors, and only one call may run at a
    // time – a second concurrent call is rejected with 13008.
    auth: {
      getAccessToken() {
        if (cfg.tokenError) return Promise.reject(cfg.tokenError);
        if (tokenBusy) return Promise.reject({ code: 13008, name: "Error", message: "getAccessToken already in progress" });
        tokenBusy = true;
        log.tokenRequests++;
        return new Promise((resolve) => setTimeout(() => { tokenBusy = false; resolve("test-sso-token"); }, 50));
      },
    },
    context: {
      // Outlook's display language: German, below /en/ English (README pictures), or the scenario's.
      displayLanguage: cfg.displayLanguage || (location.pathname.indexOf("/en/") === 0 ? "en-GB" : "de-DE"),
      officeTheme: cfg.theme ? { bodyBackgroundColor: cfg.theme } : undefined,
      // Mailbox 1.13 (sensitivity labels) only in a scenario with labels.
      requirements: { isSetSupported: (name) => name === "OpenBrowserWindowApi" || (name === "Mailbox" && Boolean(cfg.labels)) },
      // The labels published to the user (Outlook's catalog), in the policy's order.
      sensitivityLabelsCatalog: cfg.labels
        ? { getIsEnabledAsync(cb) { cb(ok(true)); }, getAsync(cb) { cb(ok(cfg.labels)); } }
        : undefined,
      ui: { openBrowserWindow: (url) => log.openedUrls.push(url) },
      // Kept in sessionStorage like Outlook keeps them in the mailbox: they survive a reload of the pane (language).
      roamingSettings: {
        get(key) {
          try {
            const value = sessionStorage.getItem(`roaming:${key}`);
            return value ? JSON.parse(value) : undefined;
          } catch (e) {
            return undefined;
          }
        },
        set(key, value) {
          try { sessionStorage.setItem(`roaming:${key}`, JSON.stringify(value)); } catch (e) { /* ignore */ }
        },
        saveAsync(cb) { if (cb) cb(ok()); },
      },
      mailbox: {
        userProfile: { displayName: "Max Mustermann", emailAddress: "max.mustermann@example.com" },
        displayMessageForm: (id) => log.openedMails.push(id),
        displayMessageFormAsync(id, cb) { log.openedMails.push(id); if (cb) cb(ok()); },
        // Mailbox events (e.g. ItemChanged of a pinned pane) – the test can fire them via window.__officeHandlers.
        addHandlerAsync(type, handler, cb) { (window.__officeHandlers = window.__officeHandlers || {})[type] = handler; if (cb) cb(ok()); },
        item: {
          getAttachmentsAsync(cb) {
            cb(ok(attachments.map((a) => ({ id: a.id, name: a.name, size: a.size, attachmentType: "file", isInline: false }))));
          },
          getAttachmentContentAsync(id, cb) {
            const attachment = attachments.find((a) => a.id === id);
            cb(attachment ? ok({ format: "base64", content: attachment.content }) : failed("unknown attachment"));
          },
          // The mail's sensitivity label (GUID) – set by the pane when a shared file has a higher one.
          sensitivityLabel: {
            getAsync(cb) { cb(ok(log.mailLabel !== undefined ? log.mailLabel : cfg.mailLabel || "")); },
            setAsync(label, cb) { log.mailLabel = typeof label === "string" ? label : label.id; cb(ok()); },
          },
          // Subject of the draft ("Verlauf" shows it until the mail is sent).
          subject: { getAsync(cb) { cb(ok(cfg.subject === undefined ? "Unterlagen Neubau" : cfg.subject)); } },
          removeAttachmentAsync(id, cb) {
            const index = attachments.findIndex((a) => a.id === id);
            if (index >= 0) attachments.splice(index, 1);
            log.removed.push(id);
            cb(ok());
          },
          addHandlerAsync(_type, _handler, cb) { if (cb) cb(ok()); },
          body: {
            getTypeAsync(cb) { cb(ok(cfg.bodyType || "html")); },
            // New Outlook reports success but inserts nothing when the cursor has never been in the body.
            setSelectedDataAsync(data, options, cb) {
              if (!cfg.cursorOutsideBody && !cfg.bodyRejectsInsert) {
                log.inserted.push({ data, coercionType: options.coercionType, at: "cursor" });
              }
              cb(ok());
            },
            prependAsync(data, options, cb) {
              if (!cfg.bodyRejectsInsert) log.inserted.push({ data, coercionType: options.coercionType, at: "top" });
              cb(ok());
            },
            // A mail being read (cfg.readBody, may change during the test) or the draft with what was inserted.
            getAsync(_coercionType, cb) {
              const read = (window.__officeConfig || {}).readBody;
              cb(ok(typeof read === "string" ? read : log.inserted.map((i) => i.data).join("\n")));
            },
          },
          sessionData: {
            getAsync(key, cb) { cb(key in log.session ? ok(log.session[key]) : failed("not found")); },
            setAsync(key, value, cb) { log.session[key] = value; cb(ok()); },
          },
          loadCustomPropertiesAsync(cb) {
            cb(ok({
              get: (key) => log.customProps[key],
              set: (key, value) => { log.customProps[key] = value; },
              remove: (key) => { delete log.customProps[key]; },
              saveAsync: (done) => done(ok()),
            }));
          },
          notificationMessages: {
            removeAsync(id, cb) { log.removedNotifications.push(id); cb(ok()); },
          },
        },
      },
    },
  };
})();
