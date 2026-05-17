// background.js (service worker)

console.log("[NaggyOutlook] background service worker started");

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  console.log("[NaggyOutlook] background got message:", msg?.type, msg);

  if (msg && msg.type === "REMINDERS_FROM_OUTLOOK") {
    console.log("[NaggyOutlook] processing REMINDERS_FROM_OUTLOOK snapshot", msg.payload);
    handleRemindersSnapshot(msg.payload).catch((e) => {
      console.error("[NaggyOutlook] handleRemindersSnapshot threw:", e);
    });
  }
  // No async sendResponse, so we don't need to return true here
});

async function handleRemindersSnapshot(raw) {
  console.log("[NaggyOutlook] handleRemindersSnapshot called with:", raw);

  const body = raw && raw.Body;
  if (!body || !Array.isArray(body.Reminders)) {
    console.warn("[NaggyOutlook] no Reminders array in payload");
    return;
  }

  const now = Date.now();

  const newEvents = {};
  for (const r of body.Reminders) {
    const uid = r.UID || (r.ItemId && r.ItemId.Id);
    if (!uid) continue;

    const start = new Date(r.StartDate);
    if (start.getTime() < now) continue;

    newEvents[uid] = {
      subject: r.Subject || "(no subject)",
      start: start.toISOString()
    };
  }

  console.log("[NaggyOutlook] newEvents snapshot:", newEvents);

  const stored = await chrome.storage.local.get("events");
  const oldEvents = stored.events || {};
  console.log("[NaggyOutlook] oldEvents snapshot:", oldEvents);

  const toClear = [];
  for (const uid of Object.keys(oldEvents)) {
    const oldEvt = oldEvents[uid];
    const newEvt = newEvents[uid];

    const changed = !newEvt ||
      newEvt.start !== oldEvt.start ||
      newEvt.subject !== oldEvt.subject;

    if (changed) {
      toClear.push(`nag-${uid}-2m`, `nag-${uid}-15s`);
    }
  }
  console.log("[NaggyOutlook] alarms to clear:", toClear);

  for (const name of toClear) {
    chrome.alarms.clear(name);
  }

  const now2 = Date.now();
  for (const [uid, evt] of Object.entries(newEvents)) {
    const oldEvt = oldEvents[uid];
    const isNewOrChanged =
      !oldEvt || oldEvt.start !== evt.start || oldEvt.subject !== evt.subject;

    if (!isNewOrChanged) continue;

    const startMs = new Date(evt.start).getTime();
    const tMinus2m = startMs - 2 * 60 * 1000;
    const tMinus15s = startMs - 15 * 1000;

    if (tMinus2m > now2) {
      const alarmName = `nag-${uid}-2m`;
      console.log("[NaggyOutlook] creating alarm:", alarmName, "at", new Date(tMinus2m));
      chrome.alarms.create(alarmName, { when: tMinus2m });
    }
    if (tMinus15s > now2) {
      const alarmName = `nag-${uid}-15s`;
      console.log("[NaggyOutlook] creating alarm:", alarmName, "at", new Date(tMinus15s));
      chrome.alarms.create(alarmName, { when: tMinus15s });
    }
  }

  await chrome.storage.local.set({
    events: newEvents,
    lastUpdated: Date.now()
  });
  console.log("[NaggyOutlook] events snapshot stored, lastUpdated set to", Date.now());
}

let currentNotificationWindowId = null;

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (!alarm.name.startsWith("nag-")) return;

  console.log("[NaggyOutlook] alarm fired:", alarm.name);

  const { events } = await chrome.storage.local.get("events") || {};
  if (!events) {
    console.warn("[NaggyOutlook] no events in storage at alarm time");
    return;
  }

  // Naive parse: assume no '-' in uid for now
  const parts = alarm.name.split("-");
  const uid = parts.slice(1, parts.length - 1).join("-");

  const evt = events[uid];
  if (!evt) {
    console.warn("[NaggyOutlook] no event found for uid", uid);
    return;
  }

  // Close existing window if any
  if (currentNotificationWindowId !== null) {
    try {
      await chrome.windows.remove(currentNotificationWindowId);
    } catch (e) {
      console.log("[NaggyOutlook] could not close existing window (maybe already closed):", e);
    }
    currentNotificationWindowId = null;
  }

  // Open a popup window instead of a notification
  const width = 400;
  const height = 300;

  const url = `reminder.html?subject=${encodeURIComponent(evt.subject || "Upcoming meeting")}&startTime=${encodeURIComponent(evt.start)}`;

  try {
    const win = await chrome.windows.create({
      url: url,
      type: "popup",
      width: width,
      height: height,
      focused: true
    });

    if (win) {
      currentNotificationWindowId = win.id;
    }

    console.log("[NaggyOutlook] popup window created for", alarm.name);
  } catch (e) {
    console.error("[NaggyOutlook] failed to create popup window:", e);
  }
});
