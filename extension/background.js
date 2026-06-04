// background.js — PageCapture Extension
// Handles toolbar button click AND relays content to native host via Native Messaging.

const HOST_NAME = "com.careerquest.pagecapture";
let port = null;

function connectToHost() {
  try {
    console.log("PageCapture: connecting to native host...");
    port = browser.runtime.connectNative(HOST_NAME);
    console.log("PageCapture: connected to native host");

    port.onDisconnect.addListener(() => {
      const err = browser.runtime.lastError;
      console.warn("PageCapture: host disconnected:", err ? err.message : "no reason");
      port = null;
    });
  } catch (e) {
    console.error("PageCapture: connectNative failed:", e.message);
    port = null;
  }
}

function ensureConnected() {
  if (!port) connectToHost();
}

// Toolbar button click — tell content script to capture
browser.browserAction.onClicked.addListener(async (tab) => {
  console.log("PageCapture: toolbar clicked on tab", tab.id);
  try {
    await browser.tabs.sendMessage(tab.id, { action: 'capture' });
  } catch (e) {
    console.log("PageCapture: injecting content script...");
    await browser.tabs.executeScript(tab.id, { file: 'content.js' });
    await browser.tabs.sendMessage(tab.id, { action: 'capture' });
  }
});

// Relay captured page data from content.js to the desktop app
browser.runtime.onMessage.addListener((message, sender) => {
  if (message.action !== 'sendToApp') return;

  console.log("PageCapture: received sendToApp message");
  ensureConnected();

  if (!port) {
    console.error("PageCapture: no port available after ensureConnected");
    return Promise.resolve({
      success: false,
      error: 'PageCapture desktop app is not running. Please open it first.'
    });
  }

  try {
    port.postMessage(message.payload);
    console.log("PageCapture: postMessage sent successfully");
    return Promise.resolve({ success: true });
  } catch (e) {
    console.error("PageCapture: postMessage failed:", e.message);
    port = null;
    return Promise.resolve({ success: false, error: e.message });
  }
});
