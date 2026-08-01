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
    exportProgressMemory.percent = Math.min(100, Math.max(0, Math.round(patch.percent)));
  } else if (exportProgressMemory.percent != null) {
    exportProgressMemory.percent = Math.min(100, Math.max(0, Math.round(exportProgressMemory.percent)));
  }
  scheduleExportProgressSave();
}

const exportWorker = {
  params: null
};

function resolveExportWorkerUrl(exportMessage, fallbackUrl) {
  const msg = exportMessage || {};
  const report = msg.report || "exportCSV";
  const lists = msg.reportSelectedLists;
  if (Array.isArray(lists) && lists.length > 0) {
    const first = lists[0];
    if (first.viewUrl) return String(first.viewUrl);
    if (first.siteUrl) return String(first.siteUrl).replace(/\/$/, "");
  }
  const siteUrl = String(msg.siteUrl || "").replace(/\/$/, "");
  if (report === "permissionsMatrix") {
    const paths = msg.matrixSelectedPaths;
    if (Array.isArray(paths) && paths.length) {
      try {
        const origin = siteUrl
          ? new URL(siteUrl).origin
          : (fallbackUrl && isMatrixSharePointUrl(fallbackUrl) ? new URL(fallbackUrl).origin : "");
        const firstPath = paths[0];
        if (firstPath) {
          const p = String(firstPath).trim();
          if (/^https?:\/\//i.test(p)) return p.replace(/\/$/, "");
          if (origin && p) return origin + (p.charAt(0) === "/" ? p : "/" + p);
        }
      } catch (_) {}
    }
    if (siteUrl && isMatrixSharePointUrl(siteUrl)) return siteUrl;
  }
  if (siteUrl && isMatrixSharePointUrl(siteUrl)) {
    if (fallbackUrl && isMatrixSharePointUrl(fallbackUrl)) return fallbackUrl;
    return siteUrl;
  }
  if (fallbackUrl && isMatrixSharePointUrl(fallbackUrl)) return fallbackUrl;
  return siteUrl || fallbackUrl || "";
}

function exportReportLabel(report) {
  const map = {
    permissionsMatrix: "Permissions matrix export",
    exportCSV: "List / library export",
    folderCount: "Folder count report",
    pathLengths: "Path length report"
  };
  return map[report] || "Report export";
}

function waitForTabReady(tabId, callback, timeoutMs) {
  timeoutMs = timeoutMs || 90000;
  let done = false;
  function finish() {
    if (done) return;
    done = true;
    chrome.tabs.onUpdated.removeListener(onUpdated);
    clearTimeout(timerId);
    setTimeout(callback, 900);
  }
  function onUpdated(id, info) {
    if (id !== tabId || info.status !== "complete") return;
    finish();
  }
  chrome.tabs.onUpdated.addListener(onUpdated);
  chrome.tabs.get(tabId, function (tab) {
    if (chrome.runtime.lastError || !tab) return;
    if (tab.status === "complete") finish();
  });
  const timerId = setTimeout(finish, timeoutMs);
}

function startExportWorker(exportMessage, sender, sendResponse) {
  const siteUrl = String(exportMessage.siteUrl || "").replace(/\/$/, "");
  if (!siteUrl && !exportMessage.reportSelectedLists && !exportMessage.matrixSelectedPaths) {
    sendResponse({ ok: false, error: "Missing site URL for export." });
    return;
  }
  if (exportProgressMemory.active) {
    sendResponse({
      ok: true,
      alreadyRunning: true,
      message: "An export is already running. Progress updates below."
    });
    return;
  }
  const uiTabId = sender.tab && sender.tab.id;
  if (uiTabId == null) {
    sendResponse({ ok: false, error: "No tab available to start export." });
    return;
  }
  chrome.tabs.get(uiTabId, function (uiTab) {
    if (chrome.runtime.lastError || !uiTab || !uiTab.url || !isMatrixSharePointUrl(uiTab.url)) {
      sendResponse({ ok: false, error: "Open a SharePoint page in this tab, then try again." });
      return;
    }
    const workerUrl = resolveExportWorkerUrl(exportMessage, uiTab.url);
    if (!workerUrl || !isMatrixSharePointUrl(workerUrl)) {
      sendResponse({ ok: false, error: "Could not determine a SharePoint URL for the export worker." });
      return;
    }
    const report = exportMessage.report || "exportCSV";
    const reportTitle = exportReportLabel(report);
    exportWorker.params = exportMessage;
    chrome.tabs.create({ url: workerUrl, active: false }, function (workerTab) {
      if (chrome.runtime.lastError || !workerTab || workerTab.id == null) {
        exportWorker.params = null;
        sendResponse({ ok: false, error: "Could not open a background export tab." });
        return;
      }
      updateExportProgressMemory({
        reset: true,
        active: true,
        report: report,
        message: reportTitle + " started…",
        percent: 0,
        startedAt: Date.now(),
        workerTabId: workerTab.id,
        uiTabId: uiTabId,
        exportParams: exportMessage,
        workerReady: false,
        logLine: reportTitle + " started…"
      });
      updateExportProgressMemory({
        logLine: "Export running in a background tab — keep working here; play Snake in the worker tab."
      });
      waitForTabReady(workerTab.id, function () {
        chrome.tabs.sendMessage(workerTab.id, {
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
            exportWorker.params = null;
            try { chrome.tabs.remove(workerTab.id); } catch (_) {}
            sendResponse({ ok: false, error: "Could not start export in background tab." });
            return;
          }
          updateExportProgressMemory({ workerReady: true });
          sendResponse({
            ok: true,
            workerLock: true,
            message: "Export running in a background tab. Keep working here — open the background tab for Snake and progress."
          });
        });
      });
    });
  });
}

function isMatrixSharePointUrl(url) {
  if (!url || typeof url !== "string") return false;
  try {
    const host = new URL(url).hostname.toLowerCase();
    return host.endsWith(".sharepoint.com") && !host.endsWith("-admin.sharepoint.com");
  } catch (_) {
    return false;
  }
}

function cancelExportInWorkerTab(notifyCancelled) {
  exportWorker.params = null;
  const tabId = exportProgressMemory.workerTabId;
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
    if (alive || !exportWorker.params || exportReconnectInFlight) return;
    if (ageMs > EXPORT_STALE_MS) {
      exportReconnectInFlight = true;
      updateExportProgressMemory({ logLine: "Export not responding — restarting background worker…" });
      const tabId = exportProgressMemory.workerTabId;
      const params = exportWorker.params || exportProgressMemory.exportParams;
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

function exportDownloadPayloadBytes(detail) {
  let bytes = 0;
  if (detail && detail.text != null) bytes = String(detail.text).length;
  else if (detail && detail.bufferBase64) bytes = Math.floor(String(detail.bufferBase64).length * 0.75);
  else if (detail && detail.buffer) bytes = detail.buffer.byteLength || 0;
  return bytes;
}

function exportDownloadTimeoutMs(detail) {
  const bytes = exportDownloadPayloadBytes(detail);
  return Math.min(900000, Math.max(60000, 60000 + Math.floor(bytes / (512 * 1024)) * 5000));
}

function normalizeBackgroundDownloadDetail(detail) {
  detail = detail || {};
  if (detail.buffer) return detail;
  if (!detail.bufferBase64) return detail;
  const binary = atob(String(detail.bufferBase64));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return {
    path: detail.path,
    mime: detail.mime,
    text: detail.text,
    buffer: bytes.buffer
  };
}

function bufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunk = 8192;
  const parts = [];
  for (let i = 0; i < bytes.length; i += chunk) {
    parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunk)));
  }
  return btoa(parts.join(""));
}

function triggerAnchorDownloadInTab(tabId, detail, sendResponse) {
  detail = normalizeBackgroundDownloadDetail(detail);
  const filename = String(detail.path || "export.dat").replace(/\\/g, "/");
  const mime = detail.mime || "application/octet-stream";
  let base64 = detail.bufferBase64;
  if (!base64 && detail.buffer) {
    try {
      base64 = bufferToBase64(detail.buffer);
    } catch (err) {
      sendResponse({ ok: false, error: err && err.message ? err.message : String(err) });
      return;
    }
  }
  if (!base64) {
    sendResponse({ ok: false, error: "Empty download payload" });
    return;
  }
  chrome.scripting.executeScript({
    target: { tabId: tabId },
    world: "MAIN",
    func: function (fname, mimeType, encoded) {
      try {
        var binary = atob(encoded);
        var bytes = new Uint8Array(binary.length);
        for (var i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        var blob = new Blob([bytes], { type: mimeType });
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = String(fname || "export.dat").split("/").pop() || "export.dat";
        a.style.display = "none";
        (document.body || document.documentElement).appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () {
          try { URL.revokeObjectURL(url); } catch (_) {}
        }, 120000);
        return true;
      } catch (e) {
        return String(e && e.message ? e.message : e);
      }
    },
    args: [filename, mime, base64]
  }, function (results) {
    if (chrome.runtime.lastError) {
      sendResponse({ ok: false, error: chrome.runtime.lastError.message || "Anchor download failed" });
      return;
    }
    const result = results && results[0] && results[0].result;
    if (result === true) sendResponse({ ok: true });
    else sendResponse({ ok: false, error: result || "Anchor download failed" });
  });
}

function waitForBackgroundDownload(downloadId, timeoutMs) {
  return new Promise(function (resolve, reject) {
    let settled = false;
    function finish(err) {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      clearInterval(pollTimer);
      try { chrome.downloads.onChanged.removeListener(onChanged); } catch (_) {}
      if (err) reject(err);
      else resolve();
    }
    function checkItem(item) {
      if (!item) return;
      if (item.state === "complete") finish(null);
      else if (item.state === "interrupted") finish(new Error(item.error || "Download interrupted"));
    }
    function onChanged(delta) {
      if (delta.id !== downloadId) return;
      if (delta.error && delta.error.current) finish(new Error(String(delta.error.current)));
      if (!delta.state || !delta.state.current) return;
      if (delta.state.current === "complete") finish(null);
      else if (delta.state.current === "interrupted") finish(new Error("Download interrupted"));
    }
    const timer = setTimeout(function () { finish(new Error("Download timed out")); }, timeoutMs);
    const pollTimer = setInterval(function () {
      if (settled) return;
      try {
        chrome.downloads.search({ id: downloadId }, function (items) {
          checkItem(items && items[0]);
        });
      } catch (_) {}
    }, 1500);
    try { chrome.downloads.onChanged.addListener(onChanged); } catch (_) {}
    try {
      chrome.downloads.search({ id: downloadId }, function (items) {
        checkItem(items && items[0]);
      });
    } catch (_) {}
  });
}

function buildExportDownloadUrl(detail) {
  if (detail.text != null) {
    const mime = detail.mime || "text/csv;charset=utf-8";
    return "data:" + mime + "," + encodeURIComponent(String(detail.text));
  }
  if (detail.buffer) {
    const mime = detail.mime || "application/octet-stream";
    const bytes = new Uint8Array(detail.buffer);
    let binary = "";
    const chunk = 8192;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return "data:" + mime + ";base64," + btoa(binary);
  }
  return "";
}

function downloadPayloadInBackground(detail) {
  detail = detail || {};
  if (detail.buffer) {
    return downloadBufferInBackground(detail);
  }
  const url = buildExportDownloadUrl(detail);
  if (!url) return Promise.reject(new Error("Empty download payload"));
  const filename = String(detail.path || "export.dat").replace(/\\/g, "/");
  const timeoutMs = exportDownloadTimeoutMs(detail);
  return new Promise(function (resolve, reject) {
    chrome.downloads.download({
      url: url,
      filename: filename,
      conflictAction: "uniquify",
      saveAs: false
    }, function (downloadId) {
      if (chrome.runtime.lastError || downloadId == null) {
        reject(new Error(chrome.runtime.lastError?.message || "Download failed"));
        return;
      }
      waitForBackgroundDownload(downloadId, timeoutMs).then(resolve, reject);
    });
  });
}

function downloadBufferInBackground(detail) {
  detail = detail || {};
  const buffer = detail.buffer;
  if (!buffer) return Promise.reject(new Error("Empty download payload"));
  const blob = new Blob([buffer], { type: detail.mime || "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const filename = String(detail.path || "export.dat").replace(/\\/g, "/");
  const timeoutMs = exportDownloadTimeoutMs(detail);
  return new Promise(function (resolve, reject) {
    chrome.downloads.download({
      url: url,
      filename: filename,
      conflictAction: "uniquify",
      saveAs: false
    }, function (downloadId) {
      if (chrome.runtime.lastError || downloadId == null) {
        URL.revokeObjectURL(url);
        reject(new Error(chrome.runtime.lastError?.message || "Download failed"));
        return;
      }
      resolve();
      waitForBackgroundDownload(downloadId, timeoutMs)
        .then(function () {
          setTimeout(function () {
            try { URL.revokeObjectURL(url); } catch (_) {}
          }, 5000);
        })
        .catch(function () {
          setTimeout(function () {
            try { URL.revokeObjectURL(url); } catch (_) {}
          }, 15000);
        });
    });
  });
}

function relayExportDownloadToTab(tabId, detail, sendResponse, timeoutMs) {
  timeoutMs = timeoutMs || 20000;
  let settled = false;
  const finish = (resp) => {
    if (settled) return;
    settled = true;
    clearTimeout(timer);
    sendResponse(resp);
  };
  const timer = setTimeout(function () {
    finish({ ok: false, error: "Download tab timed out" });
  }, timeoutMs);
  chrome.tabs.sendMessage(
    tabId,
    { action: "performExportDownload", detail: detail || {}, waitForComplete: false },
    function (response) {
      if (chrome.runtime.lastError) {
        finish({ ok: false, error: chrome.runtime.lastError.message || "Download tab unavailable" });
        return;
      }
      finish(response || { ok: false, error: "Download handler did not respond" });
    }
  );
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "SPCSVExportDownloadPerform") {
    const detail = normalizeBackgroundDownloadDetail(message.detail || {});
    const workerTabId = sender.tab && sender.tab.id;
    const uiTabId = exportProgressMemory.uiTabId;
    const finishOk = function () { sendResponse({ ok: true }); };
    const finishErr = function (err) {
      sendResponse({ ok: false, error: err && err.message ? err.message : String(err) });
    };

    if (detail.buffer) {
      downloadBufferInBackground(detail)
        .then(finishOk)
        .catch(function (swErr) {
          const fallbackTabId = workerTabId || uiTabId;
          if (fallbackTabId == null) {
            finishErr(swErr);
            return;
          }
          triggerAnchorDownloadInTab(fallbackTabId, message.detail || detail, function (resp) {
            if (resp && resp.ok) sendResponse(resp);
            else finishErr(new Error((resp && resp.error) || (swErr && swErr.message) || "Download failed"));
          });
        });
      return true;
    }

    downloadPayloadInBackground(detail)
      .then(finishOk)
      .catch(function (bgErr) {
        finishErr(bgErr);
      });
      return true;
  }
  if (message.type === "SPCSVStartExportWorker" || message.type === "SPCSVStartMatrixExport") {
    startExportWorker(message.exportMessage || {}, sender, sendResponse);
    return true;
  }
  if (message.type === "SPCSVExportCancel") {
    cancelExportInWorkerTab(true);
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
      updateExportProgressMemory({
        workerReady: true,
        logLine: "Export worker connected."
      });
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
    updateExportProgressMemory({
      active: true,
      report: message.report || "export",
      message: message.message || "Export in progress…",
      percent: message.percent != null ? message.percent : exportProgressMemory.percent,
      tabId: sender.tab && sender.tab.id,
      workerReady: exportProgressMemory.workerTabId != null ? true : exportProgressMemory.workerReady,
      logLine: message.pulse ? undefined : (message.logLine || message.message),
      pulse: message.pulse || undefined
    });
    sendResponse({ ok: true });
    return false;
  }
  if (message.type === "SPCSVExportDone") {
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
    exportWorker.params = null;
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
    cancelExportInWorkerTab(false);
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

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || message.type !== "SPOToolkitGetPowerAppsHeaderContext") return;
  const tabId =
    message.tabId != null ? message.tabId : sender.tab && sender.tab.id;
  if (tabId == null) {
    sendResponse({ ok: false, error: "No tab" });
    return false;
  }
  chrome.scripting.executeScript(
    {
      target: { tabId },
      world: "MAIN",
      func: () => {
        try {
          const Xrm = window.Xrm;
          if (!Xrm?.Utility?.getGlobalContext) {
            return { displayName: "", clientUrl: "", orgLabel: "" };
          }
          const ctx = Xrm.Utility.getGlobalContext();
          let clientUrl = "";
          try {
            clientUrl = typeof ctx.getClientUrl === "function" ? ctx.getClientUrl() || "" : "";
          } catch (_) {}
          let orgLabel = "";
          try {
            const org = ctx.organizationSettings || {};
            orgLabel = org.uniqueName || org.friendlyName || "";
          } catch (_) {}

          const finish = (displayName) => ({
            displayName: displayName || "",
            clientUrl,
            orgLabel,
          });

          if (typeof ctx.getCurrentAppProperties === "function") {
            return Promise.resolve(ctx.getCurrentAppProperties())
              .then((app) => finish((app && (app.displayName || app.uniqueName)) || ""))
              .catch(() => {
                if (typeof ctx.getCurrentAppName === "function") {
                  return Promise.resolve(ctx.getCurrentAppName())
                    .then((name) => finish(name || ""))
                    .catch(() => finish(""));
                }
                return finish("");
              });
          }
          if (typeof ctx.getCurrentAppName === "function") {
            return Promise.resolve(ctx.getCurrentAppName())
              .then((name) => finish(name || ""))
              .catch(() => finish(""));
          }
          return finish("");
        } catch (_) {
          return { displayName: "", clientUrl: "", orgLabel: "" };
        }
      },
    },
    (results) => {
      if (chrome.runtime.lastError) {
        sendResponse({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }
      const payload = results && results[0] && results[0].result ? results[0].result : null;
      sendResponse({ ok: true, context: payload || { displayName: "", clientUrl: "", orgLabel: "" } });
    }
  );
  return true;
});
