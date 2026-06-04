// background.js — PageCapture Extension
// Handles toolbar button click AND relays content to native host via Native Messaging.

const HOST_NAME = "com.careerquest.pagecapture";
let port = null;

function connectToHost() {
  try {
    port = browser.runtime.connectNative(HOST_NAME);
    port.onDisconnect.addListener(() => {
      port = null;
    });
  } catch (e) {
    port = null;
  }
}

function ensureConnected() {
  if (!port) connectToHost();
}

// Toolbar button click — tell content script to capture
browser.browserAction.onClicked.addListener(async (tab) => {
  try {
    await browser.tabs.sendMessage(tab.id, { action: 'capture' });
  } catch (e) {
    await browser.tabs.executeScript(tab.id, { file: 'content.js' });
    await browser.tabs.sendMessage(tab.id, { action: 'capture' });
  }
});

// Relay captured page data from content.js to the desktop app
browser.runtime.onMessage.addListener((message, sender) => {
  if (message.action !== 'sendToApp') return;

  ensureConnected();

  if (!port) {
    return Promise.resolve({
      success: false,
      error: 'PageCapture desktop app is not running. Please open it first.'
    });
  }

  try {
    port.postMessage(message.payload);
    return Promise.resolve({ success: true });
  } catch (e) {
    port = null;
    return Promise.resolve({ success: false, error: e.message });
  }
});
