// Service worker: View Manager/Formatter, and second-stage recycle bin (must run here so
// tabs.onUpdated survives after the popup closes).

const RB_WAIT_SESSION_KEY = "__SPO_TOOLKIT_RB_WAIT__";
const RB_WAIT_SESSION_TS_KEY = "__SPO_TOOLKIT_RB_WAIT_TS__";
const ADMIN_WAIT_SESSION_KEY = "__SPO_TOOLKIT_ADMIN_WAIT__";

/** @param {string} url */
function isSharePointTenantAdminUrl(url) {
  try {
    const u = new URL(url);
    if (u.protocol !== "https:") return false;
    return u.hostname.toLowerCase().endsWith("-admin.sharepoint.com");
  } catch (_) {
    return false;
  }
}

/** Prime tab session + inject banner (same pattern as recycle-bin wait). */
function injectAdminCenterWaitOverlay(tabId, done) {
  chrome.scripting.executeScript(
    {
      target: { tabId },
      func: (k) => {
        try {
          sessionStorage.setItem(k, "1");
        } catch (e) {}
      },
      args: [ADMIN_WAIT_SESSION_KEY],
    },
    () => {
      void chrome.runtime.lastError;
      chrome.scripting.executeScript(
        { target: { tabId }, files: ["adminCenterWaitBanner.js"] },
        () => {
          void chrome.runtime.lastError;
          if (typeof done === "function") done();
        }
      );
    }
  );
}

/**
 * Open tenant admin in a new tab and show the wait overlay after the first
 * navigation commits to *-admin.sharepoint.com (handles login redirect chains).
 * @param {string} url
 */
function openTenantAdminUrlWithWaitOverlay(url) {
  chrome.tabs.create({ url }, (tab) => {
    if (chrome.runtime.lastError || tab == null || tab.id == null) return;
    const tabId = tab.id;
    let primed = false;
    let detachTimer = 0;

    function detach() {
      if (detachTimer) clearTimeout(detachTimer);
      detachTimer = 0;
      chrome.webNavigation.onCommitted.removeListener(onCommitted);
    }

    function onCommitted(details) {
      if (primed || details.tabId !== tabId || details.frameId !== 0) return;
      const u = String(details.url || "").toLowerCase();
      if (u.indexOf("-admin.sharepoint.com") === -1) return;
      primed = true;
      detach();
      injectAdminCenterWaitOverlay(tabId);
    }

    chrome.webNavigation.onCommitted.addListener(onCommitted);
    detachTimer = setTimeout(detach, 60000);
  });
}

function clearRecycleBinWaitSession(tabId) {
  chrome.scripting.executeScript(
    {
      target: { tabId },
      func: (k, tsKey) => {
        try {
          sessionStorage.removeItem(k);
          sessionStorage.removeItem(tsKey);
        } catch (e) {}
        document.getElementById("sp-toolkit-rb-wait-root")?.remove();
      },
      args: [RB_WAIT_SESSION_KEY, RB_WAIT_SESSION_TS_KEY],
    },
    () => void chrome.runtime.lastError
  );
}

/** Prime tab session + inject banner script so overlay survives full refresh (document_start handler). */
function injectRecycleBinWaitOverlay(tabId, done) {
  chrome.scripting.executeScript(
    {
      target: { tabId },
      func: (k, tsKey) => {
        try {
          sessionStorage.setItem(k, "1");
          sessionStorage.setItem(tsKey, String(Date.now()));
        } catch (e) {}
      },
      args: [RB_WAIT_SESSION_KEY, RB_WAIT_SESSION_TS_KEY],
    },
    () => {
      void chrome.runtime.lastError;
      chrome.scripting.executeScript(
        { target: { tabId }, files: ["recycleBinWaitBanner.js"] },
        () => {
          void chrome.runtime.lastError;
          if (typeof done === "function") done();
        }
      );
    }
  );
}

/**
 * @param {number} tabId
 * @param {string} firstStageUrl
 * @param {string} secondStageUrl
 */
function openSecondStageRecycleBinSequence(tabId, firstStageUrl, secondStageUrl) {
  let settled = false;
  let sequenceClosed = false;
  let failsafe = 0;
  let fallback = 0;
  /** @type {null | (() => void)} */
  let teardownAdminHashWait = null;

  function shutdownRecycleBinSequence() {
    if (sequenceClosed) return;
    sequenceClosed = true;
    settled = true;
    clearRecycleBinWaitSession(tabId);
    chrome.tabs.onUpdated.removeListener(onUpdated);
    chrome.webNavigation.onCommitted.removeListener(onNavCommitted);
    if (failsafe) clearTimeout(failsafe);
    if (fallback) clearTimeout(fallback);
    failsafe = 0;
    fallback = 0;
    if (typeof teardownAdminHashWait === "function") {
      teardownAdminHashWait();
      teardownAdminHashWait = null;
    }
  }

  function endFirstStageOnly() {
    chrome.tabs.onUpdated.removeListener(onUpdated);
    if (fallback) clearTimeout(fallback);
    fallback = 0;
    settled = true;
  }

  function onNavCommitted(details) {
    if (sequenceClosed || details.tabId !== tabId || details.frameId !== 0) return;
    const u = (details.url || "").toLowerCase();
    if (u.indexOf("sharepoint") === -1) return;
    const isFirst = u.includes("recyclebin.aspx") && !u.includes("adminrecyclebin");
    const isAdmin = u.includes("adminrecyclebin.aspx");
    if (!isFirst && !isAdmin) return;
    injectRecycleBinWaitOverlay(tabId);
  }

  /**
   * chrome.tabs.update ignores URL fragments; SharePoint needs #view=13 in-page.
   * Open ?view=5 first, then location.replace with the full URL from the tab.
   */
  function applySecondStageUrl(fullUrl) {
    const hashPos = fullUrl.indexOf("#");
    if (hashPos < 0) {
      injectRecycleBinWaitOverlay(tabId, () => {
        chrome.tabs.update(tabId, { url: fullUrl });
      });
      return;
    }
    const withoutHash = fullUrl.slice(0, hashPos);
    let hashWaitDone = false;
    let hashWaitTimer = 0;
    let pendingInjectTimer = 0;
    let injected = false;

    function injectFullUrl() {
      if (sequenceClosed || injected) return;
      injected = true;
      try {
        chrome.webNavigation.onCommitted.removeListener(onNavCommitted);
      } catch (e) {}
      chrome.scripting.executeScript(
        {
          target: { tabId },
          func: (u, k, tsKey) => {
            try {
              sessionStorage.removeItem(k);
              sessionStorage.removeItem(tsKey);
            } catch (e) {}
            document.getElementById("sp-toolkit-rb-wait-root")?.remove();
            window.location.replace(u);
          },
          args: [fullUrl, RB_WAIT_SESSION_KEY, RB_WAIT_SESSION_TS_KEY],
        },
        () => {
          void chrome.runtime.lastError;
          shutdownRecycleBinSequence();
        }
      );
    }

    function endHashWait() {
      if (pendingInjectTimer) clearTimeout(pendingInjectTimer);
      pendingInjectTimer = 0;
      if (hashWaitDone) return;
      hashWaitDone = true;
      chrome.tabs.onUpdated.removeListener(onAdminLoaded);
      if (hashWaitTimer) clearTimeout(hashWaitTimer);
    }

    teardownAdminHashWait = endHashWait;

    function onAdminLoaded(id, info) {
      if (id !== tabId || info.status !== "complete") return;
      chrome.tabs.get(tabId, (t) => {
        if (hashWaitDone || chrome.runtime.lastError || !t?.url) return;
        try {
          if (!new URL(t.url).pathname.toLowerCase().includes("adminrecyclebin")) return;
        } catch (_) {
          return;
        }
        endHashWait();
        injectRecycleBinWaitOverlay(tabId, () => {
          if (sequenceClosed) return;
          pendingInjectTimer = setTimeout(() => {
            pendingInjectTimer = 0;
            injectFullUrl();
          }, 250);
        });
      });
    }

    injectRecycleBinWaitOverlay(tabId, () => {
      chrome.tabs.onUpdated.addListener(onAdminLoaded);
      hashWaitTimer = setTimeout(() => {
        hashWaitTimer = 0;
        endHashWait();
        injectRecycleBinWaitOverlay(tabId, () => {
          injectFullUrl();
        });
      }, 6000);

      chrome.tabs.update(tabId, { url: withoutHash }, () => {
        if (chrome.runtime.lastError) endHashWait();
      });
    });
  }

  function proceed() {
    if (settled) return;
    endFirstStageOnly();
    applySecondStageUrl(secondStageUrl);
  }

  function onUpdated(id, info) {
    if (id !== tabId || info.status !== "complete") return;
    chrome.tabs.get(tabId, (t) => {
      if (settled || chrome.runtime.lastError || !t?.url) return;
      let path = "";
      try {
        path = new URL(t.url).pathname;
      } catch (_) {
        return;
      }
      if (path.toLowerCase().includes("adminrecyclebin")) return;
      if (/recyclebin\.aspx/i.test(path) && !/adminrecyclebin/i.test(path)) {
        injectRecycleBinWaitOverlay(tabId, () => {
          setTimeout(() => {
            if (settled) return;
            proceed();
          }, 400);
        });
      }
    });
  }

  injectRecycleBinWaitOverlay(tabId, () => {
    chrome.webNavigation.onCommitted.addListener(onNavCommitted);
    failsafe = setTimeout(() => shutdownRecycleBinSequence(), 20000);
    chrome.tabs.onUpdated.addListener(onUpdated);
    chrome.tabs.update(tabId, { url: firstStageUrl }, () => {
      if (chrome.runtime.lastError) {
        shutdownRecycleBinSequence();
        return;
      }
      fallback = setTimeout(() => {
        fallback = 0;
        if (!settled) proceed();
      }, 4500);
    });
  });
}

/** Keyboard shortcut (chrome://extensions/shortcuts) — opens action popup, then Universal Search. */
chrome.commands.onCommand.addListener((command) => {
  if (command !== "open-universal-search") return;
  chrome.storage.local.get("universalSearchHotkeyEnabled", (r) => {
    if (r.universalSearchHotkeyEnabled === false) return;
    chrome.storage.local.set({ pendingOpenUniversalSearch: true }, () => {
      void chrome.runtime.lastError;
      if (typeof chrome.action !== "undefined" && typeof chrome.action.openPopup === "function") {
        chrome.action.openPopup().catch(() => {});
      }
    });
  });
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "SPOToolkitOpenTenantAdminWithWait") {
    const url = String(message.url || "");
    if (!isSharePointTenantAdminUrl(url)) {
      sendResponse({ ok: false, error: "Invalid admin URL" });
      return false;
    }
    openTenantAdminUrlWithWaitOverlay(url);
    sendResponse({ ok: true });
    return false;
  }
  if (message.type === "SPOToolkitSecondStageRecycleBin") {
    const tabId = message.tabId;
    const firstUrl = String(message.firstUrl || "");
    const secondUrl = String(message.secondUrl || "");
    if (tabId == null || !firstUrl || !secondUrl) {
      sendResponse({ ok: false, error: "Missing tab or URL" });
      return false;
    }
    openSecondStageRecycleBinSequence(tabId, firstUrl, secondUrl);
    sendResponse({ ok: true });
    return false;
  }
  if (message.type === "SPOToolkitOpenViewFormatter") {
    const previewUrl = String(message.previewUrl || (sender.tab && sender.tab.url) || "");
    if (!previewUrl || !previewUrl.includes("sharepoint.com")) {
      sendResponse({ ok: false, error: "Need a SharePoint URL" });
      return true;
    }
    const u = new URL(chrome.runtime.getURL("view-formatter.html"));
    u.searchParams.set("src", previewUrl);
    const tid = sender.tab && sender.tab.id;
    if (tid != null) u.searchParams.set("tabId", String(tid));
    chrome.tabs.create({ url: u.toString() });
    sendResponse({ ok: true });
    return true;
  }
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
