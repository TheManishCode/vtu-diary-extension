// UI automation and auto-login logic intentionally removed.
// This extension now relies on authenticated browser session + API calls only.

if (!window.__vtuContentNoopInstalled) {
  window.__vtuContentNoopInstalled = true;

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg?.type === "vtu_auto_login" || msg?.type === "vtu_bulk_upload") {
      chrome.runtime.sendMessage(
        { type: "log", text: "⚠️ UI automation is disabled. Use an active logged-in VTU session." },
        () => {
          void chrome.runtime.lastError;
        }
      );
      sendResponse({ ok: false, disabled: true });
      return true;
    }
    return false;
  });
}
