// content.js
console.log("[NaggyOutlook] content script loaded", { runtimeId: chrome.runtime?.id });

(function injectPageHook() {
  const src = chrome.runtime.getURL("page_hook.js");
  const s = document.createElement("script");
  s.src = src;
  s.onload = () => {
    console.log("[NaggyOutlook] page_hook.js injected");
    s.remove();
  };
  (document.documentElement || document.head || document.body).appendChild(s);
})();

function forwardRemindersToBackground(data, source) {
  console.log("[NaggyOutlook] content forwarding reminders to background from", source, data);

  if (!chrome.runtime?.id) {
    console.log("[NaggyOutlook] Extension context invalidated (runtime.id missing). Not sending.");
    return;
  }

  try {
    chrome.runtime.sendMessage({
      type: "REMINDERS_FROM_OUTLOOK",
      payload: data
    });
  } catch (e) {
    console.error("[NaggyOutlook] sendMessage failed:", e);
  }
}

// Listen for messages coming from page_hook.js
window.addEventListener("message", function onMessage(event) {
  if (event.source !== window) return;
  const msg = event.data;
  if (!msg || msg.type !== "OUTLOOK_GET_REMINDERS") return;

  console.log("[NaggyOutlook] content script received OUTLOOK_GET_REMINDERS via postMessage", msg.data);

  forwardRemindersToBackground(msg.data, 'postMessage');
});

// Also listen for the in-page custom event bridge
window.addEventListener('OUTLOOK_GET_REMINDERS_INTERNAL', function (e) {
  try {
    console.log('[NaggyOutlook] content script received OUTLOOK_GET_REMINDERS_INTERNAL', e.detail);
    forwardRemindersToBackground(e.detail, 'customEvent');
  } catch (err) {
    console.error('[NaggyOutlook] error handling internal reminders event', err);
  }
});

// Listen for messages coming from page_hook.js
window.addEventListener("message", function onMessage(event) {
  if (event.source !== window) return;
  const msg = event.data;
  if (!msg || msg.type !== "OUTLOOK_GET_REMINDERS") return;

  console.log("[NaggyOutlook] content script received OUTLOOK_GET_REMINDERS");

  if (!chrome.runtime?.id) {
    console.log("[NaggyOutlook] Extension context invalidated (runtime.id missing). Removing listener.");
    window.removeEventListener("message", onMessage);
    return;
  }

  try {
    chrome.runtime.sendMessage({
      type: "REMINDERS_FROM_OUTLOOK",
      payload: msg.data
    });
  } catch (e) {
    console.log("[NaggyOutlook] Extension context invalidated (sendMessage failed). Removing listener.");
    window.removeEventListener("message", onMessage);
  }
});
