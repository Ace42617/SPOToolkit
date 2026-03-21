// Minimal service worker: open View Manager from Site Contents launcher with tab + list context.
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type !== "SPOToolkitOpenViewManager") return;
  const tabId = sender.tab && sender.tab.id;
  if (!tabId) {
    sendResponse({ ok: false, error: "No tab" });
    return true;
  }
  const lid = String(message.listId || "").replace(/[{}]/g, "").trim();
  const webUrl = String(message.webUrl || "").replace(/\/$/, "");
  if (!lid || !webUrl) {
    sendResponse({ ok: false, error: "Missing list or site" });
    return true;
  }
  const u = new URL(chrome.runtime.getURL("views.html"));
  u.searchParams.set("tabId", String(tabId));
  u.searchParams.set("listId", lid);
  u.searchParams.set("webUrl", webUrl);
  chrome.tabs.create({ url: u.toString() });
  sendResponse({ ok: true });
  return true;
});
