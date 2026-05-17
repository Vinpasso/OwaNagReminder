// page_hook.js
console.log("[NaggyOutlook] page_hook.js loaded – patching fetch and XHR");

(function () {
  const getUrlString = (input) => {
    if (typeof input === "string") return input;
    if (input && typeof input.url === "string") return input.url;
    return "";
  };

  const isGetRemindersUrl = (url) =>
    typeof url === "string" &&
    url.includes("/owa/service.svc") &&
    url.includes("action=GetReminders");

  const isRemindersPayload = (data) =>
    data && typeof data === "object" &&
    data.Body && Array.isArray(data.Body.Reminders);

  const postReminders = (data, source) => {
    if (!isRemindersPayload(data)) {
      return false;
    }
    console.log("[NaggyOutlook] posting Outlook reminders payload from", source, data);
    try {
      window.postMessage({ type: "OUTLOOK_GET_REMINDERS", data }, "*");
    } catch (e) {
      console.error('[NaggyOutlook] window.postMessage failed', e);
    }
    try {
      window.dispatchEvent(new CustomEvent('OUTLOOK_GET_REMINDERS_INTERNAL', { detail: data }));
    } catch (e) {
      console.error('[NaggyOutlook] dispatchEvent failed', e);
    }
    return true;
  };

  const maybeParseJson = (text, url, source) => {
    try {
      const data = JSON.parse(text);
      if (!postReminders(data, source)) {
        console.log("[NaggyOutlook] JSON parsed, but not reminders payload", source, url);
      }
    } catch (err) {
      console.error("[NaggyOutlook] failed to parse JSON from", source, url, err);
    }
  };

  const origFetch = window.fetch;
  if (origFetch) {
    window.fetch = async function (...args) {
      const url = getUrlString(args[0]);
      const isGetReminders = isGetRemindersUrl(url);

      if (isGetReminders) {
        console.log("[NaggyOutlook] fetch called for GetReminders:", url);
      }

      const res = await origFetch.apply(this, args);

      if (isGetReminders) {
        try {
          const clone = res.clone();
          clone.json().then((data) => {
            if (!postReminders(data, "fetch")) {
              console.log("[NaggyOutlook] fetch response JSON is not reminders payload", url);
            }
          }).catch((err) => {
            console.error("[NaggyOutlook] error parsing GetReminders JSON:", err);
          });
        } catch (e) {
          console.error("[NaggyOutlook] error cloning GetReminders response:", e);
        }
      }

      return res;
    };
  } else {
    console.warn("[NaggyOutlook] window.fetch not found, skipping fetch hook");
  }

  const origXhrOpen = XMLHttpRequest.prototype.open;
  const origXhrSend = XMLHttpRequest.prototype.send;

  XMLHttpRequest.prototype.open = function (method, url, async, user, password) {
    try {
      this.__naggy_url = typeof url === "string" ? url : getUrlString(url);
    } catch (_e) {
      this.__naggy_url = "";
    }
    return origXhrOpen.apply(this, arguments);
  };

  XMLHttpRequest.prototype.send = function (body) {
    const url = this.__naggy_url;
    const isGetReminders = isGetRemindersUrl(url);

    if (isGetReminders) {
      console.log("[NaggyOutlook] XHR send called for GetReminders:", url);
      this.addEventListener("load", function () {
        if (this.status >= 200 && this.status < 300) {
          const responseText = this.responseType === "text" || this.responseType === "" ? this.responseText : null;
          if (responseText) {
            maybeParseJson(responseText, url, "xhr");
          } else if (this.response && typeof this.response === "object") {
            if (!postReminders(this.response, "xhr")) {
              console.log("[NaggyOutlook] XHR response is not reminders payload", url);
            }
          }
        }
      });
    }

    return origXhrSend.apply(this, arguments);
  };
})();
