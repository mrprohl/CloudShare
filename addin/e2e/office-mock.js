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
  });
  const ok = (value) => ({ status: "succeeded", value });
  const failed = (message) => ({ status: "failed", error: { message } });
  const attachments = (cfg.attachments || []).slice();

  window.Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Html: "html", Text: "text" },
    EventType: { AttachmentsChanged: "olkAttachmentsChanged", OfficeThemeChanged: "officeThemeChanged" },
    MailboxEnums: {
      AttachmentType: { File: "file", Item: "item", Cloud: "cloud" },
      AttachmentContentFormat: { Base64: "base64", Url: "url", Eml: "eml", ICalendar: "iCalendar" },
    },
    onReady(callback) {
      const info = { host: "Outlook", platform: "OfficeOnline" };
      setTimeout(() => callback && callback(info), 0);
      return Promise.resolve(info);
    },
    auth: { getAccessToken: () => Promise.resolve("test-sso-token") },
    context: {
      officeTheme: cfg.theme ? { bodyBackgroundColor: cfg.theme } : undefined,
      requirements: { isSetSupported: (name) => name === "OpenBrowserWindowApi" },
      ui: { openBrowserWindow: (url) => log.openedUrls.push(url) },
      roamingSettings: { get: () => undefined, set() {}, saveAsync(cb) { if (cb) cb(ok()); } },
      mailbox: {
        userProfile: { displayName: "Max Mustermann", emailAddress: "max.mustermann@example.com" },
        addHandlerAsync(_type, _handler, cb) { if (cb) cb(ok()); },
        item: {
          getAttachmentsAsync(cb) {
            cb(ok(attachments.map((a) => ({ id: a.id, name: a.name, size: a.size, attachmentType: "file", isInline: false }))));
          },
          getAttachmentContentAsync(id, cb) {
            const attachment = attachments.find((a) => a.id === id);
            cb(attachment ? ok({ format: "base64", content: attachment.content }) : failed("unknown attachment"));
          },
          removeAttachmentAsync(id, cb) {
            const index = attachments.findIndex((a) => a.id === id);
            if (index >= 0) attachments.splice(index, 1);
            log.removed.push(id);
            cb(ok());
          },
          addHandlerAsync(_type, _handler, cb) { if (cb) cb(ok()); },
          body: {
            getTypeAsync(cb) { cb(ok(cfg.bodyType || "html")); },
            setSelectedDataAsync(data, options, cb) {
              log.inserted.push({ data, coercionType: options.coercionType });
              cb(ok());
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
