// Service worker: View Manager/Formatter, and second-stage recycle bin (must run here so
// tabs.onUpdated survives after the popup closes).

const RB_WAIT_SESSION_KEY = "__SPO_TOOLKIT_RB_WAIT__";
const RB_WAIT_SESSION_TS_KEY = "__SPO_TOOLKIT_RB_WAIT_TS__";
const ADMIN_WAIT_SESSION_KEY = "__SPO_TOOLKIT_ADMIN_WAIT__";
const VIEW_FORMATTER_FRAME_RULE_IDS = [101];
const VIEW_FORMATTER_FRAME_RULE_FILTERS = ["||sharepoint.com^"];

function buildViewFormatterFrameRules(tabIds) {
  return VIEW_FORMATTER_FRAME_RULE_FILTERS.map((urlFilter, index) => ({
    id: VIEW_FORMATTER_FRAME_RULE_IDS[index],
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
      urlFilter,
      resourceTypes: ["sub_frame"],
      tabIds,
    },
  }));
}

function normalizeDnrTabIds(tabIds) {
  const seen = new Set();
  const out = [];
  (tabIds || []).forEach((value) => {
    const n = Number(value);
    if (!Number.isInteger(n) || n < 0 || seen.has(n)) return;
    seen.add(n);
    out.push(n);
  });
  return out;
}

function readViewFormatterFrameRuleTabIds(done) {
  if (!chrome.declarativeNetRequest || !chrome.declarativeNetRequest.getSessionRules) {
    done([]);
    return;
  }
  chrome.declarativeNetRequest.getSessionRules((rules) => {
    if (chrome.runtime.lastError) {
      done([]);
      return;
    }
    const ids = new Set(VIEW_FORMATTER_FRAME_RULE_IDS);
    const tabIds = [];
    (rules || []).forEach((rule) => {
      if (!ids.has(rule.id)) return;
      const ruleTabIds = rule.condition && rule.condition.tabIds;
      if (Array.isArray(ruleTabIds)) tabIds.push(...ruleTabIds);
    });
    done(normalizeDnrTabIds(tabIds));
  });
}

function setViewFormatterFrameRuleTabIds(tabIds, done) {
  if (!chrome.declarativeNetRequest || !chrome.declarativeNetRequest.updateSessionRules) {
    if (typeof done === "function") done(false);
    return;
  }
  const cleanTabIds = normalizeDnrTabIds(tabIds);
  const update = { removeRuleIds: VIEW_FORMATTER_FRAME_RULE_IDS };
  if (cleanTabIds.length) update.addRules = buildViewFormatterFrameRules(cleanTabIds);
  chrome.declarativeNetRequest.updateSessionRules(update, () => {
    const ok = !chrome.runtime.lastError;
    if (typeof done === "function") done(ok);
  });
}

function enableViewFormatterFrameRules(tabId, done) {
  readViewFormatterFrameRuleTabIds((tabIds) => {
    tabIds.push(tabId);
    setViewFormatterFrameRuleTabIds(tabIds, done);
  });
}

function disableViewFormatterFrameRules(tabId) {
  readViewFormatterFrameRuleTabIds((tabIds) => {
    setViewFormatterFrameRuleTabIds(tabIds.filter((id) => id !== tabId));
  });
}

function openViewFormatterTab(previewUrl, sourceTabId) {
  const u = new URL(chrome.runtime.getURL("view-formatter.html"));
  u.searchParams.set("src", previewUrl);
  if (sourceTabId != null) u.searchParams.set("tabId", String(sourceTabId));
  chrome.tabs.create({ url: u.toString() }, (tab) => {
    if (chrome.runtime.lastError || tab == null || tab.id == null) return;
    enableViewFormatterFrameRules(tab.id);
  });
}

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
    return u.hostname.toLowerCase().endsWith(".sharepoint.com");
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
      func: (k, tk) => {
        try {
          sessionStorage.removeItem(k);
          sessionStorage.removeItem(tk);
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
      func: (k, tk) => {
        try {
          sessionStorage.setItem(k, "1");
          sessionStorage.setItem(tk, String(Date.now()));
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

  function armFailsafe(ms) {
    if (failsafe) clearTimeout(failsafe);
    failsafe = setTimeout(() => shutdownRecycleBinSequence(), ms);
  }

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
        chrome.tabs.update(tabId, { url: fullUrl }, () => shutdownRecycleBinSequence());
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
        if (chrome.runtime.lastError) shutdownRecycleBinSequence();
      });
    });
  }

  function proceed() {
    if (settled) return;
    endFirstStageOnly();
    armFailsafe(30000);
    applySecondStageUrl(secondStageUrl);
  }

  function proceedIfFirstStageLoaded() {
    chrome.tabs.get(tabId, (t) => {
      if (settled || chrome.runtime.lastError || !t?.url || t.status !== "complete") return;
      try {
        const path = new URL(t.url).pathname;
        if (/recyclebin\.aspx/i.test(path) && !/adminrecyclebin/i.test(path)) proceed();
      } catch (_) {}
    });
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
    armFailsafe(60000);
    chrome.tabs.onUpdated.addListener(onUpdated);
    chrome.tabs.update(tabId, { url: firstStageUrl }, () => {
      if (chrome.runtime.lastError) {
        shutdownRecycleBinSequence();
        return;
      }
      fallback = setTimeout(() => {
        fallback = 0;
        if (!settled) proceedIfFirstStageLoaded();
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

chrome.tabs.onRemoved.addListener((tabId) => {
  disableViewFormatterFrameRules(tabId);
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "SPOToolkitEnableViewFormatterFrameRules") {
    const tabId = Number(message.tabId);
    if (!Number.isInteger(tabId) || tabId < 0) {
      sendResponse({ ok: false, error: "Missing formatter tab" });
      return false;
    }
    enableViewFormatterFrameRules(tabId, (ok) => {
      sendResponse(ok ? { ok: true } : { ok: false, error: "Could not scope frame rules" });
    });
    return true;
  }
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
    if (!isSharePointPreviewUrl(previewUrl)) {
      sendResponse({ ok: false, error: "Need a SharePoint URL" });
      return true;
    }
    const sourceTabId = message.tabId != null ? message.tabId : sender.tab && sender.tab.id;
    openViewFormatterTab(previewUrl, sourceTabId);
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
