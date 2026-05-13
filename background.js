// Service worker: View Manager/Formatter, and second-stage recycle bin (must run here so
// tabs.onUpdated survives after the popup closes).

const RB_WAIT_SESSION_KEY = "__SPO_TOOLKIT_RB_WAIT__";
const ADMIN_WAIT_SESSION_KEY = "__SPO_TOOLKIT_ADMIN_WAIT__";
const VIEW_FORMATTER_TAB_IDS_KEY = "__SPO_TOOLKIT_VIEW_FORMATTER_TABS__";
const VIEW_FORMATTER_DNR_RULE_IDS = [901, 902, 903];
const VIEW_FORMATTER_DNR_DOMAINS = ["sharepoint.com", "sharepoint.us", "sharepoint.de"];

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

/** @param {string} url */
function isSharePointPreviewUrl(url) {
  try {
    const u = new URL(url);
    if (u.protocol !== "https:") return false;
    const host = u.hostname.toLowerCase();
    return (
      host === "sharepoint.com" ||
      host.endsWith(".sharepoint.com") ||
      host === "sharepoint.us" ||
      host.endsWith(".sharepoint.us") ||
      host === "sharepoint.de" ||
      host.endsWith(".sharepoint.de")
    );
  } catch (_) {
    return false;
  }
}

/** @param {string} url */
function isViewFormatterUrl(url) {
  return String(url || "").startsWith(chrome.runtime.getURL("view-formatter.html"));
}

function getStoredViewFormatterTabIds(done) {
  const storage = chrome.storage && chrome.storage.session;
  if (!storage) {
    done([]);
    return;
  }
  storage.get(VIEW_FORMATTER_TAB_IDS_KEY, (r) => {
    if (chrome.runtime.lastError) {
      done([]);
      return;
    }
    const ids = Array.isArray(r && r[VIEW_FORMATTER_TAB_IDS_KEY]) ? r[VIEW_FORMATTER_TAB_IDS_KEY] : [];
    done(ids.filter((id) => Number.isInteger(id) && id > 0));
  });
}

function setStoredViewFormatterTabIds(tabIds, done) {
  const storage = chrome.storage && chrome.storage.session;
  if (!storage) {
    if (typeof done === "function") done();
    return;
  }
  storage.set({ [VIEW_FORMATTER_TAB_IDS_KEY]: tabIds }, () => {
    void chrome.runtime.lastError;
    if (typeof done === "function") done();
  });
}

function buildViewFormatterFrameRules(tabIds) {
  return VIEW_FORMATTER_DNR_DOMAINS.map((domain, index) => ({
    id: VIEW_FORMATTER_DNR_RULE_IDS[index],
    priority: 1,
    action: {
      type: "modifyHeaders",
      responseHeaders: [
        { header: "x-frame-options", operation: "remove" },
        { header: "content-security-policy", operation: "remove" },
        { header: "content-security-policy-report-only", operation: "remove" },
      ],
    },
    condition: {
      urlFilter: "||" + domain + "^",
      resourceTypes: ["sub_frame"],
      tabIds,
    },
  }));
}

function updateViewFormatterFrameRules(tabIds, done) {
  const dnr = chrome.declarativeNetRequest;
  if (!dnr || typeof dnr.updateSessionRules !== "function") {
    if (typeof done === "function") done({ ok: false, error: "Missing declarativeNetRequest support" });
    return;
  }
  const uniqueTabIds = Array.from(new Set(tabIds)).filter((id) => Number.isInteger(id) && id > 0);
  dnr.updateSessionRules(
    {
      removeRuleIds: VIEW_FORMATTER_DNR_RULE_IDS,
      addRules: uniqueTabIds.length ? buildViewFormatterFrameRules(uniqueTabIds) : [],
    },
    () => {
      const err = chrome.runtime.lastError;
      if (typeof done === "function") done(err ? { ok: false, error: err.message } : { ok: true });
    }
  );
}

function registerViewFormatterTab(tabId, done) {
  getStoredViewFormatterTabIds((existing) => {
    const next = Array.from(new Set(existing.concat([tabId])));
    setStoredViewFormatterTabIds(next, () => updateViewFormatterFrameRules(next, done));
  });
}

function unregisterViewFormatterTab(tabId) {
  getStoredViewFormatterTabIds((existing) => {
    const next = existing.filter((id) => id !== tabId);
    setStoredViewFormatterTabIds(next, () => updateViewFormatterFrameRules(next));
  });
}

function openViewFormatterWithScopedFrameRules(previewUrl, sourceTabId, sendResponse) {
  if (!isSharePointPreviewUrl(previewUrl)) {
    sendResponse({ ok: false, error: "Need a SharePoint URL" });
    return;
  }
  chrome.tabs.create({ url: chrome.runtime.getURL("view-formatter.html") }, (tab) => {
    if (chrome.runtime.lastError || tab == null || tab.id == null) {
      sendResponse({ ok: false, error: chrome.runtime.lastError?.message || "Could not open View formatter" });
      return;
    }
    const formatterTabId = tab.id;
    registerViewFormatterTab(formatterTabId, (ruleResult) => {
      if (!ruleResult || !ruleResult.ok) {
        unregisterViewFormatterTab(formatterTabId);
        sendResponse({ ok: false, error: (ruleResult && ruleResult.error) || "Could not enable formatter preview" });
        return;
      }
      const u = new URL(chrome.runtime.getURL("view-formatter.html"));
      u.searchParams.set("src", previewUrl);
      if (sourceTabId != null) u.searchParams.set("tabId", String(sourceTabId));
      chrome.tabs.update(formatterTabId, { url: u.toString() }, () => {
        const err = chrome.runtime.lastError;
        if (err) unregisterViewFormatterTab(formatterTabId);
        sendResponse(err ? { ok: false, error: err.message } : { ok: true });
      });
    });
  });
}

chrome.tabs.onRemoved.addListener((tabId) => {
  unregisterViewFormatterTab(tabId);
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (changeInfo.url && !isViewFormatterUrl(changeInfo.url)) {
    unregisterViewFormatterTab(tabId);
  }
});

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
      func: (k) => {
        try {
          sessionStorage.removeItem(k);
        } catch (e) {}
        document.getElementById("sp-toolkit-rb-wait-root")?.remove();
      },
      args: [RB_WAIT_SESSION_KEY],
    },
    () => void chrome.runtime.lastError
  );
}

/** Prime tab session + inject banner script so overlay survives full refresh (document_start handler). */
function injectRecycleBinWaitOverlay(tabId, done) {
  chrome.scripting.executeScript(
    {
      target: { tabId },
      func: (k) => {
        try {
          sessionStorage.setItem(k, "1");
        } catch (e) {}
      },
      args: [RB_WAIT_SESSION_KEY],
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
    let injected = false;

    function injectFullUrl() {
      if (injected) return;
      injected = true;
      try {
        chrome.webNavigation.onCommitted.removeListener(onNavCommitted);
      } catch (e) {}
      chrome.scripting.executeScript(
        {
          target: { tabId },
          func: (u, k) => {
            try {
              sessionStorage.removeItem(k);
            } catch (e) {}
            document.getElementById("sp-toolkit-rb-wait-root")?.remove();
            window.location.replace(u);
          },
          args: [fullUrl, RB_WAIT_SESSION_KEY],
        },
        () => {
          void chrome.runtime.lastError;
          shutdownRecycleBinSequence();
        }
      );
    }

    function endHashWait() {
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
          setTimeout(injectFullUrl, 250);
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
    const messageTabId = Number.isInteger(message.tabId) ? message.tabId : null;
    const sourceTabId = messageTabId != null ? messageTabId : sender.tab && sender.tab.id;
    openViewFormatterWithScopedFrameRules(previewUrl, sourceTabId, sendResponse);
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
