// Service worker: View Manager/Formatter, and second-stage recycle bin (must run here so
// tabs.onUpdated survives after the popup closes).

const RB_WAIT_SESSION_KEY = "__SPO_TOOLKIT_RB_WAIT__";
const ADMIN_WAIT_SESSION_KEY = "__SPO_TOOLKIT_ADMIN_WAIT__";
/** Long-lived port from popup/content keeps the service worker alive for the async RecycleBin → AdminRecycleBin handoff. */
const SECOND_STAGE_RB_PORT = "SPOToolkitSecondStageRecycleBin";

function normalizeTrailingSlashRb(url) {
  if (!url || typeof url !== "string") return "";
  return url.replace(/\/$/, "");
}

const ADMIN_RECYCLE_BIN_SECOND_STAGE_SUFFIX_RB = "/_layouts/15/AdminRecycleBin.aspx?view=5#view=13";

function secondStageRecycleBinUrlRb(siteCollectionRootAbsoluteUrl) {
  const base = normalizeTrailingSlashRb(siteCollectionRootAbsoluteUrl);
  if (!base) return "";
  return base + ADMIN_RECYCLE_BIN_SECOND_STAGE_SUFFIX_RB;
}

function isToolkitSharePointPageForRb(url) {
  if (!url || typeof url !== "string") return false;
  if (!url.includes("sharepoint.com")) return false;
  try {
    return !new URL(url).hostname.toLowerCase().endsWith("-admin.sharepoint.com");
  } catch (_) {
    return true;
  }
}

/** Same path logic as popup.js `parseContextFromUrl` (Quick Links / recycle bin). */
function parseContextFromUrlForRb(pageUrl) {
  if (!pageUrl || !isToolkitSharePointPageForRb(pageUrl)) {
    return { webAbsoluteUrl: "", pageListId: "", viewId: "" };
  }
  try {
    const u = new URL(pageUrl);
    let webAbsoluteUrl = "";
    let pageListId = "";
    let viewId = "";
    const listParam = u.searchParams.get("List") || u.searchParams.get("list");
    if (listParam) pageListId = String(listParam).replace(/^\{|\}$/g, "").replace(/%7B|%7D/gi, "");
    const viewParam = u.searchParams.get("View") || u.searchParams.get("view");
    if (viewParam) viewId = String(viewParam).replace(/^\{|\}$/g, "").replace(/%7B|%7D/gi, "");
    const path = decodeURIComponent(u.pathname.replace(/\/$/, ""));
    const segments = path.split("/").filter(Boolean);
    let siteEnd = -1;
    for (let i = 0; i < segments.length; i++) {
      if ((segments[i] === "sites" || segments[i] === "site" || segments[i] === "teams") && segments[i + 1]) {
        siteEnd = i + 1;
        break;
      }
    }
    webAbsoluteUrl = siteEnd >= 0 ? u.origin + "/" + segments.slice(0, siteEnd + 1).join("/") : u.origin;
    return { webAbsoluteUrl, pageListId, viewId };
  } catch (_) {
    return { webAbsoluteUrl: "", pageListId: "", viewId: "" };
  }
}

function sitePathFromWebAbsoluteUrlRb(webAbsoluteUrl) {
  if (!webAbsoluteUrl) return "";
  try {
    return new URL(webAbsoluteUrl).pathname.replace(/\/$/, "") || "/";
  } catch (_) {
    return "";
  }
}

/**
 * Same URL resolution as popup quick links: getPageContext.siteAbsoluteUrl → REST rootweb → siteBase.
 * @param {number} tabId
 * @param {chrome.runtime.Port | null | undefined} [keepalivePort]
 */
function startSecondStageRecycleBinLikePopup(tabId, keepalivePort) {
  function bail() {
    if (!keepalivePort) return;
    try {
      keepalivePort.disconnect();
    } catch (_) {}
  }
  chrome.tabs.get(tabId, (tab) => {
    if (chrome.runtime.lastError || !tab || !tab.url) {
      bail();
      return;
    }
    const pageUrl = tab.url;
    if (!isToolkitSharePointPageForRb(pageUrl)) {
      bail();
      return;
    }
    const ctx = parseContextFromUrlForRb(pageUrl);
    let siteBase = "";
    try {
      siteBase = (ctx.webAbsoluteUrl || "").replace(/\/$/, "") || new URL(pageUrl).origin;
    } catch (_) {
      bail();
      return;
    }
    if (!siteBase) {
      bail();
      return;
    }

    chrome.tabs.sendMessage(tabId, { action: "getPageContext" }, (pc) => {
      void chrome.runtime.lastError;
      let fromPage = "";
      if (pc && pc.ok && (pc.siteAbsoluteUrl || "").trim()) {
        fromPage = normalizeTrailingSlashRb(pc.siteAbsoluteUrl.trim());
      }
      if (fromPage) {
        finish(fromPage);
        return;
      }
      const sitePath = sitePathFromWebAbsoluteUrlRb(siteBase) || "/";
      chrome.tabs.sendMessage(
        tabId,
        { action: "rest", method: "GET", path: sitePath + "/_api/site/rootweb?$select=Url" },
        (rest) => {
          void chrome.runtime.lastError;
          let fromRest = "";
          if (rest && rest.ok && rest.data) {
            const d = rest.data;
            const u =
              (typeof d.Url === "string" && d.Url.trim()) ||
              (typeof d.url === "string" && d.url.trim()) ||
              "";
            if (u) fromRest = normalizeTrailingSlashRb(u);
          }
          finish(fromRest || normalizeTrailingSlashRb(siteBase));
        }
      );
    });

    function finish(siteCollRoot) {
      const secondUrl = secondStageRecycleBinUrlRb(siteCollRoot);
      const firstUrl = normalizeTrailingSlashRb(siteBase) + "/_layouts/15/RecycleBin.aspx";
      if (!secondUrl || !firstUrl) {
        bail();
        return;
      }
      openSecondStageRecycleBinSequence(tabId, firstUrl, secondUrl, keepalivePort);
    }
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

chrome.runtime.onConnect.addListener((port) => {
  if (port.name !== SECOND_STAGE_RB_PORT) return;
  const onStart = (msg) => {
    if (!msg) return;
    if (msg.type === "startLikePopup") {
      port.onMessage.removeListener(onStart);
      const tabId = msg.tabId != null ? msg.tabId : port.sender && port.sender.tab && port.sender.tab.id;
      if (tabId == null) {
        try {
          port.disconnect();
        } catch (_) {}
        return;
      }
      startSecondStageRecycleBinLikePopup(tabId, port);
      return;
    }
    if (msg.type !== "start") return;
    port.onMessage.removeListener(onStart);
    const tabId = msg.tabId != null ? msg.tabId : port.sender && port.sender.tab && port.sender.tab.id;
    const firstUrl = String(msg.firstUrl || "");
    const secondUrl = String(msg.secondUrl || "");
    if (tabId == null || !firstUrl || !secondUrl) {
      try {
        port.disconnect();
      } catch (_) {}
      return;
    }
    openSecondStageRecycleBinSequence(tabId, firstUrl, secondUrl, port);
  };
  port.onMessage.addListener(onStart);
});

const EXPORT_PROGRESS_KEY = "spcsvExportProgress";
const EXPORT_STALE_MS = 45000;
let exportProgressMemory = { log: [], active: false };
let exportProgressSaveTimer = null;
let exportReconnectInFlight = false;

function scheduleExportProgressSave() {
  if (exportProgressSaveTimer) return;
  exportProgressSaveTimer = setTimeout(() => {
    exportProgressSaveTimer = null;
    chrome.storage.session.set({
      [EXPORT_PROGRESS_KEY]: Object.assign({}, exportProgressMemory, { updatedAt: Date.now() })
    });
  }, 120);
}

function updateExportProgressMemory(patch) {
  if (patch.reset) {
    exportProgressMemory = { log: [], active: false };
    delete patch.reset;
  }
  if (patch.logLine) {
    const msg = String(patch.logLine);
    const log = Array.isArray(exportProgressMemory.log) ? exportProgressMemory.log.slice() : [];
    const last = log.length ? log[log.length - 1].msg : "";
    if (msg && msg !== last) log.push({ t: Date.now(), msg });
    exportProgressMemory.log = log.slice(-400);
    delete patch.logLine;
  }
  if (patch.pulse) {
    exportProgressMemory.pulseAt = Date.now();
    delete patch.pulse;
  }
  exportProgressMemory = Object.assign({}, exportProgressMemory, patch);
  if (patch.percent != null) {
    exportProgressMemory.percent = Math.max(exportProgressMemory.percent || 0, Math.min(100, Math.round(patch.percent)));
  } else if (exportProgressMemory.percent != null) {
    exportProgressMemory.percent = Math.max(0, Math.min(100, Math.round(exportProgressMemory.percent)));
  }
  scheduleExportProgressSave();
}

const matrixExportWorker = {
  params: null
};

function isMatrixSharePointUrl(url) {
  if (!url || typeof url !== "string") return false;
  try {
    const host = new URL(url).hostname.toLowerCase();
    return host.endsWith(".sharepoint.com") && !host.endsWith("-admin.sharepoint.com");
  } catch (_) {
    return false;
  }
}

function cancelMatrixExportInTab(notifyCancelled) {
  matrixExportWorker.params = null;
  const tabId = exportProgressMemory.workerTabId != null ? exportProgressMemory.workerTabId : exportProgressMemory.uiTabId;
  if (tabId != null) {
    try {
      chrome.tabs.sendMessage(tabId, { action: "cancelExportCSV" });
    } catch (_) {}
  }
  if (notifyCancelled && exportProgressMemory.active) {
    updateExportProgressMemory({
      active: false,
      success: false,
      cancelled: true,
      message: "Export cancelled.",
      logLine: "Export cancelled.",
      finishedAt: Date.now(),
      workerReady: false
    });
  }
}

function pingMatrixExportTab(callback) {
  const tabId = exportProgressMemory.workerTabId != null ? exportProgressMemory.workerTabId : exportProgressMemory.uiTabId;
  if (tabId == null) {
    callback(false);
    return;
  }
  chrome.tabs.get(tabId, function (tab) {
    if (chrome.runtime.lastError || !tab) {
      callback(false);
      return;
    }
    chrome.tabs.sendMessage(tabId, { type: "SPCSVExportPing" }, function (resp) {
      callback(!chrome.runtime.lastError && resp && resp.running === true);
    });
  });
}

function touchExportProgressStaleCheck() {
  if (!exportProgressMemory.active) return;
  const lastTouch = Math.max(
    exportProgressMemory.updatedAt || 0,
    exportProgressMemory.pulseAt || 0,
    exportProgressMemory.startedAt || 0
  );
  const ageMs = Date.now() - lastTouch;

  if (!exportProgressMemory.workerReady) {
    if (ageMs > 120000) {
      updateExportProgressMemory({
        active: false,
        success: false,
        message: "Export stopped (failed to start). Start a new export to continue.",
        logLine: "Export stopped (failed to start). Start a new export to continue.",
        finishedAt: Date.now(),
        workerReady: false
      });
    }
    return;
  }

  pingMatrixExportTab(function (alive) {
    if (!alive && exportProgressMemory.active && ageMs > 45000) {
      updateExportProgressMemory({
        active: false,
        success: false,
        message: "Export stopped (page was refreshed or closed). Start a new export to continue.",
        logLine: "Export stopped (page was refreshed or closed). Start a new export to continue.",
        finishedAt: Date.now(),
        workerReady: false
      });
      return;
    }
    if (alive || !matrixExportWorker.params || exportReconnectInFlight) return;
    if (ageMs > EXPORT_STALE_MS) {
      exportReconnectInFlight = true;
      updateExportProgressMemory({ logLine: "Export not responding — restarting worker tab…" });
      const tabId = exportProgressMemory.workerTabId != null ? exportProgressMemory.workerTabId : exportProgressMemory.uiTabId;
      const params = matrixExportWorker.params || exportProgressMemory.exportParams;
      if (tabId != null && params) {
        try {
          chrome.tabs.sendMessage(tabId, { action: "cancelExportCSV" }, function () {
            setTimeout(function () {
              chrome.tabs.sendMessage(tabId, Object.assign({ action: "runExportCSVWorker", _workerFrame: true }, params), function () {
                exportReconnectInFlight = false;
              });
            }, 600);
          });
        } catch (_) {
          exportReconnectInFlight = false;
        }
      } else {
        exportReconnectInFlight = false;
      }
    }
  });
}

setInterval(function () {
  if (exportProgressMemory.active) touchExportProgressStaleCheck();
}, 15000);

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "SPCSVStartMatrixExport") {
    const exportMessage = message.exportMessage || {};
    const siteUrl = String(exportMessage.siteUrl || "").replace(/\/$/, "");
    if (!siteUrl || !isMatrixSharePointUrl(siteUrl)) {
      sendResponse({ ok: false, error: "Invalid SharePoint site URL for export." });
      return false;
    }
    if (exportProgressMemory.active && exportProgressMemory.report === "permissionsMatrix") {
      sendResponse({
        ok: true,
        alreadyRunning: true,
        message: "Permissions matrix export is already running. Progress updates below."
      });
      return false;
    }
    const workerTabId = sender.tab && sender.tab.id;
    if (workerTabId == null) {
      sendResponse({ ok: false, error: "No tab available to start export." });
      return false;
    }
    chrome.tabs.get(workerTabId, function (workerTab) {
      if (chrome.runtime.lastError || !workerTab || !workerTab.url || !isMatrixSharePointUrl(workerTab.url)) {
        sendResponse({ ok: false, error: "Open a SharePoint page in this tab, then try again." });
        return;
      }
      const pageUrl = workerTab.url;
      matrixExportWorker.params = exportMessage;
      chrome.tabs.create({ url: pageUrl, active: true }, function (newTab) {
        if (chrome.runtime.lastError || !newTab || newTab.id == null) {
          sendResponse({ ok: false, error: "Could not open a continuation tab." });
          return;
        }
        updateExportProgressMemory({
          reset: true,
          active: true,
          report: "permissionsMatrix",
          message: "Permissions matrix export started…",
          percent: 0,
          startedAt: Date.now(),
          workerTabId: workerTabId,
          uiTabId: newTab.id,
          exportParams: exportMessage,
          workerReady: true,
          logLine: "Permissions matrix export started…"
        });
        updateExportProgressMemory({
          logLine: "Export running in a dedicated tab — continue working in the new tab."
        });
        chrome.tabs.sendMessage(workerTabId, {
          action: "matrixWorkerLockAndRun",
          exportMessage: exportMessage
        }, function () {
          if (chrome.runtime.lastError) {
            updateExportProgressMemory({
              active: false,
              success: false,
              message: "Export failed to start.",
              logLine: "Export failed to start.",
              finishedAt: Date.now(),
              workerReady: false
            });
            matrixExportWorker.params = null;
            sendResponse({ ok: false, error: "Could not start export on worker tab." });
            return;
          }
          sendResponse({
            ok: true,
            workerLock: true,
            message: "Export started in a dedicated tab. Continue in the new tab — the worker tab closes when finished."
          });
        });
      });
    });
    return true;
  }
  if (message.type === "SPCSVExportCancel") {
    cancelMatrixExportInTab(true);
    sendResponse({ ok: true });
    return false;
  }
  if (message.type === "SPCSVExportEnsureWorker") {
    touchExportProgressStaleCheck();
    sendResponse({ ok: true });
    return false;
  }
  if (message.type === "SPCSVExportStarted") {
    if (exportProgressMemory.active && (message.report || "export") === (exportProgressMemory.report || "export")) {
      if (exportProgressMemory.report !== "permissionsMatrix") {
        updateExportProgressMemory({
          workerReady: true,
          logLine: "Export worker connected."
        });
      }
    } else {
      updateExportProgressMemory({
        reset: true,
        active: true,
        report: message.report || "export",
        message: message.message || "Export started…",
        percent: 0,
        tabId: sender.tab && sender.tab.id,
        startedAt: Date.now(),
        logLine: message.message || "Export started…"
      });
    }
    sendResponse({ ok: true });
    return false;
  }
  if (message.type === "SPCSVExportProgress") {
    var nextPct = message.percent != null ? message.percent : exportProgressMemory.percent;
    if (nextPct != null && exportProgressMemory.percent != null) {
      nextPct = Math.max(exportProgressMemory.percent, nextPct);
    }
    updateExportProgressMemory({
      active: true,
      report: message.report || "export",
      message: message.message || "Export in progress…",
      percent: nextPct,
      tabId: sender.tab && sender.tab.id,
      workerReady: exportProgressMemory.report === "permissionsMatrix" ? true : exportProgressMemory.workerReady,
      logLine: message.pulse ? undefined : (message.logLine || message.message),
      pulse: message.pulse || undefined
    });
    sendResponse({ ok: true });
    return false;
  }
  if (message.type === "SPCSVExportDone") {
    const wasMatrix = exportProgressMemory.report === "permissionsMatrix" || !!matrixExportWorker.params;
    const workerTabId = exportProgressMemory.workerTabId;
    const success = !!message.success;
    const doneMessage = message.message || (success ? "Export complete" : "Export failed");
    updateExportProgressMemory({
      active: false,
      success: success,
      message: doneMessage,
      percent: success ? 100 : exportProgressMemory.percent,
      logLine: doneMessage,
      finishedAt: Date.now(),
      workerReady: false
    });
    if (wasMatrix) {
      matrixExportWorker.params = null;
      if (workerTabId != null) {
        const closeMs = success ? 5000 : 15000;
        try {
          chrome.tabs.sendMessage(workerTabId, {
            action: "matrixWorkerFinish",
            success: success,
            message: doneMessage,
            autoCloseMs: closeMs
          });
        } catch (_) {}
        setTimeout(function () {
          chrome.tabs.remove(workerTabId, function () {
            void chrome.runtime.lastError;
          });
        }, closeMs);
      }
    }
    sendResponse({ ok: true });
    return false;
  }
  if (message.type === "SPCSVExportProgressGet") {
    chrome.storage.session.get(EXPORT_PROGRESS_KEY, (data) => {
      sendResponse(data[EXPORT_PROGRESS_KEY] || null);
    });
    return true;
  }
  if (message.type === "SPCSVExportProgressClear") {
    cancelMatrixExportInTab(false);
    exportProgressMemory = { log: [], active: false };
    if (exportProgressSaveTimer) {
      clearTimeout(exportProgressSaveTimer);
      exportProgressSaveTimer = null;
    }
    chrome.storage.session.remove(EXPORT_PROGRESS_KEY);
    sendResponse({ ok: true });
    return false;
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
  if (message.type === "SPOToolkitSecondStageRecycleBinLikePopup") {
    const tabId =
      message.tabId != null ? message.tabId : sender.tab && sender.tab.id;
    if (tabId == null) {
      sendResponse({ ok: false, error: "No tab" });
      return false;
    }
    startSecondStageRecycleBinLikePopup(tabId, null);
    sendResponse({ ok: true });
    return false;
  }
  if (message.type === "SPOToolkitSecondStageRecycleBin") {
    const tabId =
      message.tabId != null ? message.tabId : sender.tab && sender.tab.id;
    const firstUrl = String(message.firstUrl || "");
    const secondUrl = String(message.secondUrl || "");
    if (tabId == null || !firstUrl || !secondUrl) {
      sendResponse({ ok: false, error: "Missing tab or URL" });
      return false;
    }
    openSecondStageRecycleBinSequence(tabId, firstUrl, secondUrl, null);
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
  if (message.type === "SPOToolkitOpenViewsPage") {
    const u = new URL(chrome.runtime.getURL("views.html"));
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
