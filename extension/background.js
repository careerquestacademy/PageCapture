// background.js — PageCapture Extension
// Listens for toolbar button click, tells content script to capture.

browser.browserAction.onClicked.addListener(async (tab) => {
  try {
    await browser.tabs.sendMessage(tab.id, { action: 'capture' });
  } catch (e) {
    // Content script not ready — inject it first
    await browser.tabs.executeScript(tab.id, { file: 'content.js' });
    await browser.tabs.sendMessage(tab.id, { action: 'capture' });
  }
});
