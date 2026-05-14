import {
  applyQuicklinksFilter,
  listColumnSettingsUrl,
  searchSchemaColumnMatchesFilter,
  searchSchemaListMetaFromResponse,
} from "./lib/popupUi.mjs";
import { secondStageRecycleBinUrl, normalizeTrailingSlash } from "./lib/recycleBinUrls.mjs";
import { normalizeSharePointPreviewUrl } from "./lib/viewFormatterSecurity.mjs";

const btnExport = document.getElementById("btnExport");
const btnChooseColumns = document.getElementById("btnChooseColumns");
const btnSettings = document.getElementById("btnSettings");
const btnSettingsBack = document.getElementById("btnSettingsBack");
const pageSizeSelect = document.getElementById("pageSizeSelect");
const defaultExportFormatSelect = document.getElementById("defaultExportFormatSelect");
const statusEl = document.getElementById("status");
const mainPanel = document.getElementById("downloadMain");
const reportSelect = document.getElementById("reportSelect");
const formatSelect = document.getElementById("formatSelect");
const reportOptions = document.getElementById("reportOptions");
const pickerPanel = document.getElementById("pickerPanel");
const settingsPanel = document.getElementById("settingsPanel");
const columnList = document.getElementById("columnList");
const btnBack = document.getElementById("btnBack");
const btnSelectAll = document.getElementById("btnSelectAll");
const btnSelectNone = document.getElementById("btnSelectNone");
const btnSelectDefaults = document.getElementById("btnSelectDefaults");
const btnExportSelected = document.getElementById("btnExportSelected");
const pickerStatus = document.getElementById("pickerStatus");

let pickerData = null; // { defaultColumns, pickerFields }

const SHAREPOINT_FACTS = window.SHAREPOINT_FACTS || [];

const DEFAULT_PAGE_SIZE = 5000;
const DEFAULT_EXPORT_FORMAT = "xlsx";

// Refinable managed property types and their index range (00 to count-1). Default: RefinableString.
const REFINABLE_TYPES = [
  { id: "RefinableString", label: "RefinableString", count: 200 },
  { id: "RefinableStringFirst", label: "RefinableStringFirst", count: 40 },
  { id: "RefinableStringLn", label: "RefinableStringLn", count: 10 },
  { id: "RefinableStringWbOff", label: "RefinableStringWbOff", count: 50 },
  { id: "RefinableStringWbOffFirst", label: "RefinableStringWbOffFirst", count: 50 },
  { id: "RefinableDate", label: "RefinableDate", count: 20 },
  { id: "RefinableDateFirst", label: "RefinableDateFirst", count: 20 },
  { id: "RefinableDateSingle", label: "RefinableDateSingle", count: 5 },
  { id: "RefinableDateInvariant", label: "RefinableDateInvariant", count: 2 },
  { id: "RefinableInt", label: "RefinableInt", count: 50 },
  { id: "RefinableDecimal", label: "RefinableDecimal", count: 10 },
  { id: "RefinableDouble", label: "RefinableDouble", count: 10 },
  { id: "RefinableYesNo", label: "RefinableYesNo", count: 5 }
];

// --- Dark mode ---
async function loadDarkMode() {
  const st = await chrome.storage.local.get("darkMode");
  const isDark = !!st.darkMode;
  document.documentElement.classList.toggle("dark-mode", isDark);
  updateDarkToggleLabel(isDark);
}
function updateDarkToggleLabel(isDark) {
  const btn = document.getElementById("darkModeToggle");
  if (!btn) return;
  if (isDark) {
    btn.title = "Night Owl";
    btn.setAttribute("aria-label", "Theme: Night Owl. Click to switch to Early Riser.");
  } else {
    btn.title = "Early Riser";
    btn.setAttribute("aria-label", "Theme: Early Riser. Click to switch to Night Owl.");
  }
}
document.getElementById("darkModeToggle")?.addEventListener("click", () => {
  const isDark = document.documentElement.classList.toggle("dark-mode");
  chrome.storage.local.set({ darkMode: isDark });
  updateDarkToggleLabel(isDark);
});
const headerIconBtn = document.getElementById("headerIconBtn");
const headerCompassImg = document.getElementById("headerCompassIcon");
if (headerCompassImg) headerCompassImg.src = chrome.runtime.getURL("icon.png");
if (headerIconBtn) {
  headerIconBtn.addEventListener("click", async () => {
    headerIconBtn.classList.remove("spin");
    headerIconBtn.offsetHeight;
    headerIconBtn.classList.add("spin");
    setTimeout(() => headerIconBtn.classList.remove("spin"), 600);
    if (!SHAREPOINT_FACTS.length) return;
    const fact = SHAREPOINT_FACTS[Math.floor(Math.random() * SHAREPOINT_FACTS.length)];
    const factText = typeof fact === "string" ? fact : (fact.text || "");
    const sourceUrl = typeof fact === "object" && fact.sourceUrl && String(fact.sourceUrl).trim() ? fact.sourceUrl : "";
    if (typeof fact === "object" && typeof fact.openUrl === "string") chrome.tabs.create({ url: fact.openUrl });
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (tab?.id) chrome.tabs.sendMessage(tab.id, { action: "showFact", fact: factText, sourceUrl });
    } catch (_) {}
  });
}
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message && message.type === "SPCSVExportDone") {
    if (message.success) {
      setStatus("Export complete", "success");
    } else {
      setStatus(message.message || "Export failed.", "error");
    }
    sendResponse({});
    return true;
  }
});
loadDarkMode();

// --- Settings (gear: open settings screen; back button: return to main) ---
const chkListsLauncherEnabled = document.getElementById("chkListsLauncherEnabled");
const chkUniversalSearchOpenOnLoad = document.getElementById("chkUniversalSearchOpenOnLoad");
const chkUniversalSearchHotkeyEnabled = document.getElementById("chkUniversalSearchHotkeyEnabled");
const universalSearchShortcutDisplay = document.getElementById("universalSearchShortcutDisplay");

function refreshUniversalSearchShortcutDisplay() {
  if (!universalSearchShortcutDisplay || typeof chrome.commands === "undefined" || !chrome.commands.getAll) return;
  chrome.commands.getAll((cmds) => {
    const c = cmds.find((x) => x.name === "open-universal-search");
    universalSearchShortcutDisplay.textContent =
      c && c.shortcut ? "Current shortcut: " + c.shortcut : "Current shortcut: not set";
  });
}

document.getElementById("btnSettingsGear")?.addEventListener("click", () => {
  document.body.classList.add("popup-settings-visible");
  refreshUniversalSearchShortcutDisplay();
});
document.getElementById("btnSettingsBackToMain")?.addEventListener("click", () => {
  document.body.classList.remove("popup-settings-visible");
});
chrome.storage.local.get(
  ["listsLauncherEnabled", "universalSearchOpenOnLoad", "universalSearchHotkeyEnabled"],
  (r) => {
    if (chkListsLauncherEnabled) chkListsLauncherEnabled.checked = r.listsLauncherEnabled !== false;
    if (chkUniversalSearchOpenOnLoad) chkUniversalSearchOpenOnLoad.checked = !!r.universalSearchOpenOnLoad;
    if (chkUniversalSearchHotkeyEnabled) chkUniversalSearchHotkeyEnabled.checked = r.universalSearchHotkeyEnabled !== false;
  }
);
chkListsLauncherEnabled?.addEventListener("change", () => {
  chrome.storage.local.set({ listsLauncherEnabled: chkListsLauncherEnabled.checked });
});
chkUniversalSearchOpenOnLoad?.addEventListener("change", () => {
  chrome.storage.local.set({ universalSearchOpenOnLoad: chkUniversalSearchOpenOnLoad.checked });
});
chkUniversalSearchHotkeyEnabled?.addEventListener("change", () => {
  chrome.storage.local.set({ universalSearchHotkeyEnabled: chkUniversalSearchHotkeyEnabled.checked });
});

document.getElementById("btnOpenExtensionShortcuts")?.addEventListener("click", () => {
  chrome.tabs.create({ url: "chrome://extensions/shortcuts" });
});

// --- Tabs: Quick links, Page Properties, Reports ---
const TAB_IDS = ["quicklinks", "context", "searchSchema", "reports", "refinableProps", "viewManager"];
function updateTabIndicator() {
  const bar = document.querySelector(".tab-bar");
  const indicator = bar && bar.querySelector(".tab-indicator");
  const active = bar && bar.querySelector(".tab.active");
  if (!bar || !indicator || !active || active.offsetParent === null) return;
  indicator.style.left = active.offsetLeft + "px";
  indicator.style.width = active.offsetWidth + "px";
}
function switchToTab(id) {
  const tabEl = document.querySelector(".tab[data-tab=\"" + id + "\"]");
  if (!tabEl || tabEl.offsetParent === null) return;
  document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));
  tabEl.classList.add("active");
  const panel = document.getElementById(id + "Panel");
  if (panel) panel.classList.add("active");
  document.body.classList.toggle("columns-tab-active", id === "searchSchema");
  document.body.classList.remove("reports-expanded", "reports-settings-visible");
  requestAnimationFrame(function () { updateTabIndicator(); });
  if (id === "quicklinks") renderQuickLinks();
  else if (id === "context") loadContextInfo();
  else if (id === "searchSchema") loadSearchSchema();
  else if (id === "reports") {
    updateExportReportLabel();
    toggleReportOptions();
  }
  else if (id === "refinableProps") initRefinableProps();
}
document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    const id = tab.dataset.tab;
    if (id) {
      chrome.storage.local.set({ popupActiveTab: id });
      switchToTab(id);
    }
  });
});

// --- Universal Search (Finder-style) ---
const universalSearchOverlay = document.getElementById("universalSearchOverlay");
const universalSearchInput = document.getElementById("universalSearchInput");
const universalSearchResults = document.getElementById("universalSearchResults");
const btnUniversalSearch = document.getElementById("btnUniversalSearch");
let universalSearchIndex = [];
let universalSearchFiltered = [];
let universalSearchActive = -1;

function openUniversalSearch() {
  if (!universalSearchOverlay || !universalSearchInput) return;
  document.body.classList.add("universal-search-open");
  universalSearchOverlay.classList.add("open");
  universalSearchOverlay.setAttribute("aria-hidden", "false");
  universalSearchInput.value = "";
  refreshUniversalSearchIndex().then(() => {
    renderUniversalSearchResults("");
    requestAnimationFrame(() => universalSearchInput.focus());
  });
}

function closeUniversalSearch() {
  if (!universalSearchOverlay || !universalSearchInput) return;
  universalSearchOverlay.classList.remove("open");
  universalSearchOverlay.setAttribute("aria-hidden", "true");
  universalSearchInput.value = "";
  document.body.classList.remove("universal-search-open");
}

/** Hotkey sets this flag in the service worker; listen so search opens even if the popup was already open (openPopup does not reload). */
chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== "local" || !changes.pendingOpenUniversalSearch) return;
  if (changes.pendingOpenUniversalSearch.newValue !== true) return;
  chrome.storage.local.remove("pendingOpenUniversalSearch");
  document.body.classList.remove("popup-settings-visible");
  requestAnimationFrame(() => openUniversalSearch());
});

function normalizeSearchText(v) {
  return String(v || "").toLowerCase().trim();
}

function normalizeGuidCopyText(value) {
  const raw = String(value == null ? "" : value).trim();
  const noBraces = raw.replace(/[{}]/g, "");
  return /^[0-9a-fA-F-]{36}$/.test(noBraces) ? noBraces : raw;
}

function isEmptyUniversalSearchPropValue(v) {
  if (v == null) return true;
  return String(v).trim() === "";
}

function addUniversalSearchItem(group, title, meta, keywords, run) {
  universalSearchIndex.push({ group, title, meta, keywords: normalizeSearchText(keywords), run });
}

async function refreshUniversalSearchIndex() {
  universalSearchIndex = [];
  let activeTab = null;
  let siteBase = "";
  let parsedCtx = { webAbsoluteUrl: "", pageListId: "", viewId: "" };
  let pageCtx = null;
  try {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    activeTab = tabs && tabs[0] ? tabs[0] : null;
    if (activeTab && activeTab.url && isToolkitSharePointPage(activeTab.url)) {
      parsedCtx = parseContextFromUrl(activeTab.url);
      siteBase = (parsedCtx.webAbsoluteUrl || "").replace(/\/$/, "") || new URL(activeTab.url).origin;
      if (activeTab.id != null) {
        try {
          pageCtx = await new Promise((resolve) => {
            chrome.tabs.sendMessage(activeTab.id, { action: "getPageContext" }, (res) => {
              if (chrome.runtime.lastError || !res || !res.ok) resolve(null);
              else resolve(res);
            });
          });
        } catch (_) {
          pageCtx = null;
        }
      }
    }
  } catch (_) {}
  let isListPageForSearch = false;
  if (activeTab && activeTab.id != null && isToolkitSharePointPage(activeTab?.url)) {
    try {
      const checkRes = await new Promise((resolve) => {
        chrome.tabs.sendMessage(activeTab.id, { action: "checkListPage" }, (res) => {
          if (chrome.runtime.lastError || !res) resolve({ isListPage: false });
          else resolve(res);
        });
      });
      isListPageForSearch = !!checkRes.isListPage;
    } catch (_) {
      isListPageForSearch = false;
    }
  }
  await waitForRenderQuickLinksIdle();
  try {
    await renderQuickLinks();
  } catch (_) {}
  // Ensure global index includes freshest page props and columns when available.
  try {
    await loadContextInfo();
  } catch (_) {}
  try {
    await loadSearchSchema();
  } catch (_) {}

  document.querySelectorAll(".tab").forEach((tab) => {
    if (tab.offsetParent === null) return;
    const id = tab.dataset.tab || "";
    const title = (tab.textContent || "").trim();
    if (!id || !title) return;
    addUniversalSearchItem(
      "Tabs",
      title,
      "Switch tab",
      title + " " + id,
      () => switchToTab(id)
    );
  });

  if (activeTab && activeTab.url && isToolkitSharePointPage(activeTab.url)) {
    const jumpToContextValue = (valueText, fallbackText) => {
      switchToTab("context");
      const input = document.getElementById("contextFilterInput");
      const byValue = document.getElementById("contextFilterByValue");
      if (byValue) byValue.checked = true;
      if (input) input.value = String(valueText || fallbackText || "");
      renderContextList();
      try {
        if (valueText) navigator.clipboard.writeText(normalizeGuidCopyText(String(valueText)));
      } catch (_) {}
    };
    const urlPath = (() => {
      try {
        return new URL(activeTab.url).pathname || "/";
      } catch (_) {
        return "";
      }
    })();
    const listId =
      (pageCtx && (pageCtx.pageListId || pageCtx.listId || pageCtx.listGuid)) ||
      parsedCtx.pageListId ||
      "";
    const viewId = (pageCtx && (pageCtx.viewId || pageCtx.pageViewId)) || parsedCtx.viewId || "";
    const webUrl = (pageCtx && pageCtx.webAbsoluteUrl) || parsedCtx.webAbsoluteUrl || "";
    const siteUrl = (pageCtx && pageCtx.siteAbsoluteUrl) || "";
    const siteId = (pageCtx && pageCtx.siteId) || "";
    addUniversalSearchItem(
      "Current Page",
      "Page URL",
      urlPath,
      "page url current url location " + activeTab.url + " " + urlPath,
      () => jumpToContextValue(activeTab.url, "url")
    );
    if (webUrl) {
      addUniversalSearchItem("Current Page", "Web URL", webUrl, "web url webabsoluteurl " + webUrl, () => jumpToContextValue(webUrl, "webAbsoluteUrl"));
    }
    if (siteUrl) {
      addUniversalSearchItem("Current Page", "Site URL", siteUrl, "site url siteabsoluteurl " + siteUrl, () => jumpToContextValue(siteUrl, "siteAbsoluteUrl"));
    }
    if (listId) {
      addUniversalSearchItem(
        "Current Page",
        "List ID",
        String(listId),
        "list id listid pagelistid listguid guid " + listId,
        () => jumpToContextValue(listId, "listid")
      );
    }
    if (viewId) {
      addUniversalSearchItem(
        "Current Page",
        "View ID",
        String(viewId),
        "view id viewid pageviewid guid " + viewId,
        () => jumpToContextValue(viewId, "viewid")
      );
    }
    if (siteId) {
      addUniversalSearchItem(
        "Current Page",
        "Site ID",
        String(siteId),
        "site id siteid guid " + siteId,
        () => jumpToContextValue(siteId, "siteid")
      );
    }
  }

  const quickLinksContainer = document.getElementById("quicklinksContent");
  if (quickLinksContainer) {
    quickLinksContainer.querySelectorAll(".ql-section").forEach((section) => {
      const sectionLabel = (section.querySelector(".ql-section-label")?.textContent || "").trim();
      section.querySelectorAll("a[href]").forEach((a) => {
        const label = (a.textContent || "").trim();
        const href = a.getAttribute("href") || "";
        if (!label || !href) return;
        addUniversalSearchItem(
          "Quick Links",
          label,
          sectionLabel || "Quick link",
          label + " " + href + " " + sectionLabel,
          () => {
            if (
              /second stage recycle bin/i.test(label) &&
              /adminrecyclebin\.aspx/i.test(href) &&
              activeTab &&
              activeTab.id != null &&
              siteBase
            ) {
              const firstUrl = normalizeTrailingSlash(siteBase) + "/_layouts/15/RecycleBin.aspx";
              chrome.runtime.sendMessage(
                {
                  type: "SPOToolkitSecondStageRecycleBin",
                  tabId: activeTab.id,
                  firstUrl,
                  secondUrl: href,
                },
                () => {
                  void chrome.runtime.lastError;
                }
              );
            } else if (isTenantAdminQuickLinkHref(href)) {
              chrome.runtime.sendMessage({ type: "SPOToolkitOpenTenantAdminWithWait", url: href }, () => {
                void chrome.runtime.lastError;
              });
            } else {
              chrome.tabs.create({ url: href });
            }
            window.close();
          }
        );
      });
    });
  }

  if (contextData && typeof contextData === "object") {
    Object.entries(contextData).slice(0, 200).forEach(([k, v]) => {
      if (isEmptyUniversalSearchPropValue(v)) return;
      const val = String(v);
      addUniversalSearchItem(
        "Page Props",
        k,
        val.slice(0, 90),
        k + " " + val,
        () => {
          switchToTab("context");
          const input = document.getElementById("contextFilterInput");
          if (input) input.value = k;
          renderContextList();
        }
      );
    });
  }

  if (Array.isArray(searchSchemaColumns) && searchSchemaColumns.length) {
    searchSchemaColumns.slice(0, 300).forEach((c) => {
      const title = c.title || c.internalName || "";
      const internal = c.internalName || "";
      const crawled = c.crawledProperty || "";
      const type = c.type || "";
      addUniversalSearchItem(
        "Columns",
        title,
        internal || type,
        title + " " + internal + " " + crawled + " " + type,
        () => {
          switchToTab("searchSchema");
          const input = document.getElementById("searchSchemaFilter");
          if (input) input.value = title || internal;
          renderSearchSchemaList();
        }
      );
    });
  }

  if (isListPageForSearch) {
    addUniversalSearchItem("Tools", "View Manager", "List & library views", "view manager views manage edit editor view editor", () => {
      void chrome.tabs.query({ active: true, currentWindow: true }).then(([tab]) => {
        if (!tab?.id || !isToolkitSharePointPage(tab.url)) return;
        chrome.tabs.create({ url: chrome.runtime.getURL("views.html?tabId=" + tab.id) });
      });
    });
    addUniversalSearchItem("Tools", "View formatter", "JSON editor & list preview", "view formatter json column formatting format", () => {
      void chrome.tabs.query({ active: true, currentWindow: true }).then(([tab]) => {
        openViewFormatterForTab(tab, false);
      });
    });
    addUniversalSearchItem("Reports", "Run Export", "Reports", "run export reports", () => {
      switchToTab("reports");
      btnExport?.focus();
    });
    addUniversalSearchItem("Reports", "Permissions Matrix", "Report type", "permissions matrix", () => {
      switchToTab("reports");
      if (reportSelect) reportSelect.value = "permissionsMatrix";
      toggleReportOptions();
      btnExport?.focus();
    });
    addUniversalSearchItem("Reports", "Path Length Report", "Report type", "path length", () => {
      switchToTab("reports");
      if (reportSelect) reportSelect.value = "pathLengths";
      toggleReportOptions();
      btnExport?.focus();
    });
    addUniversalSearchItem("Reports", "Folder and Item Counts", "Report type", "folder item count", () => {
      switchToTab("reports");
      if (reportSelect) reportSelect.value = "folderCount";
      toggleReportOptions();
      btnExport?.focus();
    });
  }
}

function renderUniversalSearchResults(queryRaw) {
  if (!universalSearchResults) return;
  const query = normalizeSearchText(queryRaw);
  universalSearchFiltered = query
    ? universalSearchIndex.filter((it) => it.keywords.includes(query))
    : universalSearchIndex;
  universalSearchActive = universalSearchFiltered.length ? 0 : -1;

  if (!universalSearchFiltered.length) {
    universalSearchResults.innerHTML = '<div class="us-empty">No matches yet.</div>';
    return;
  }

  const groups = new Map();
  universalSearchFiltered.forEach((it) => {
    if (!groups.has(it.group)) groups.set(it.group, []);
    groups.get(it.group).push(it);
  });

  universalSearchResults.innerHTML = "";
  let flatIndex = 0;
  groups.forEach((items, group) => {
    const g = document.createElement("div");
    g.className = "us-group";
    const h = document.createElement("div");
    h.className = "us-group-title";
    h.textContent = group;
    g.appendChild(h);
    items.slice(0, 18).forEach((it) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "us-item" + (flatIndex === universalSearchActive ? " active" : "");
      btn.dataset.idx = String(flatIndex);
      btn.innerHTML =
        '<span class="us-item-title">' + escapeHtml(it.title) + "</span>" +
        (it.meta ? '<span class="us-item-meta">' + escapeHtml(it.meta) + "</span>" : "");
      btn.addEventListener("click", () => {
        it.run();
        closeUniversalSearch();
      });
      g.appendChild(btn);
      flatIndex += 1;
    });
    universalSearchResults.appendChild(g);
  });
}

function moveUniversalSearchSelection(dir) {
  if (!universalSearchFiltered.length) return;
  const max = universalSearchFiltered.length - 1;
  universalSearchActive = Math.max(0, Math.min(max, universalSearchActive + dir));
  const buttons = universalSearchResults ? Array.from(universalSearchResults.querySelectorAll(".us-item")) : [];
  buttons.forEach((b, idx) => b.classList.toggle("active", idx === universalSearchActive));
  const activeEl = buttons[universalSearchActive];
  if (activeEl) activeEl.scrollIntoView({ block: "nearest" });
}

btnUniversalSearch?.addEventListener("click", openUniversalSearch);
universalSearchOverlay?.addEventListener("click", (e) => {
  if (e.target === universalSearchOverlay) closeUniversalSearch();
});
universalSearchInput?.addEventListener("input", () => renderUniversalSearchResults(universalSearchInput.value));
universalSearchInput?.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    e.preventDefault();
    closeUniversalSearch();
    return;
  }
  if (e.key === "ArrowDown") {
    e.preventDefault();
    moveUniversalSearchSelection(1);
    return;
  }
  if (e.key === "ArrowUp") {
    e.preventDefault();
    moveUniversalSearchSelection(-1);
    return;
  }
  if (e.key === "Enter" && universalSearchFiltered[universalSearchActive]) {
    e.preventDefault();
    universalSearchFiltered[universalSearchActive].run();
    closeUniversalSearch();
  }
});
document.addEventListener("keydown", (e) => {
  const isCmdK = (e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === "k";
  if (isCmdK) {
    e.preventDefault();
    openUniversalSearch();
    return;
  }
  if (e.key === "Escape" && universalSearchOverlay?.classList.contains("open")) {
    e.preventDefault();
    closeUniversalSearch();
  }
});

const RECENT_SP_KEY = "recentSharePointUrls";
const RECENT_SP_MAX = 25;

/** SharePoint pages the toolkit supports (excludes tenant admin / *-admin.sharepoint.com). */
function isToolkitSharePointPage(url) {
  return normalizeSharePointPreviewUrl(url) !== "";
}

function openViewFormatterForTab(tab, showError) {
  const previewUrl = normalizeSharePointPreviewUrl(tab?.url);
  if (!previewUrl) {
    if (showError) alert("Open a SharePoint list or library page first, then open View formatter.");
    return;
  }

  chrome.runtime.sendMessage(
    { type: "SPOToolkitOpenViewFormatter", previewUrl, tabId: tab?.id },
    (response) => {
      if (!showError) return;
      if (chrome.runtime.lastError) {
        alert(chrome.runtime.lastError.message || "Could not open View formatter.");
        return;
      }
      if (!response || response.ok === false) {
        alert((response && response.error) || "Could not open View formatter.");
      }
    }
  );
}

/** Dedupe recent list by full page path (not only site root). */
function recentSharePointPageKey(url) {
  try {
    const u = new URL(url);
    const path = u.pathname.replace(/\/$/, "") || "";
    return (u.origin + path).toLowerCase();
  } catch (_) {
    return (url.split("?")[0].replace(/\/$/, "") || url).toLowerCase();
  }
}

function migrateRecentSharePointStoredList(list) {
  if (!Array.isArray(list)) return [];
  return list
    .map((item) => {
      if (typeof item === "string") return { url: item, title: "" };
      if (item && typeof item.url === "string") return { url: item.url, title: String(item.title || "") };
      return null;
    })
    .filter(Boolean);
}

function cleanSharePointTabTitle(title) {
  if (!title) return "";
  let t = String(title).trim();
  t = t.replace(/\s+[\-|]\s+SharePoint Online.*$/i, "");
  t = t.replace(/\s+[\-|]\s+SharePoint.*$/i, "");
  t = t.replace(/\s+[\-|]\s*Microsoft 365\s*$/i, "");
  t = t.trim();
  const parts = t.split(/\s*-\s*/).map((p) => p.trim()).filter(Boolean);
  if (parts.length) return parts[0];
  return t;
}

function friendlyNameFromSharePointUrl(url) {
  try {
    const u = new URL(url);
    const segments = u.pathname.split("/").filter(Boolean);
    if (segments.length === 0) return "Home";
    const last = segments[segments.length - 1];
    if (/^viewlsts\.aspx$/i.test(last)) return "Site contents";
    if (/^settings\.aspx$/i.test(last)) return "Site settings";
    if (/^home\.aspx$/i.test(last)) return "Home";
    if (segments.includes("Forms")) {
      const prev = segments[segments.length - 2];
      return (prev && decodeURIComponent(prev)) || "Library";
    }
    const leaf = decodeURIComponent(last.replace(/\.aspx$/i, ""));
    return leaf || "Page";
  } catch (_) {
    return "SharePoint";
  }
}

function formatRecentSharePointEntryLabel(entry) {
  const url = typeof entry === "string" ? entry : entry.url;
  let title = typeof entry === "string" ? "" : String(entry.title || "");
  title = cleanSharePointTabTitle(title);
  if (!title) title = friendlyNameFromSharePointUrl(url);
  try {
    const host = new URL(url).hostname;
    return title + " - " + host;
  } catch (_) {
    return url;
  }
}

async function addCurrentUrlToRecentSharePoint(url, pageTitle) {
  if (!isToolkitSharePointPage(url)) return;
  const normalized = url.split("?")[0].replace(/\/$/, "") || url;
  const title = String(pageTitle || "").trim();
  const { [RECENT_SP_KEY]: list = [] } = await chrome.storage.local.get(RECENT_SP_KEY);
  const migrated = migrateRecentSharePointStoredList(list);
  const pageKey = recentSharePointPageKey(normalized);
  const deduped = [{ url: normalized, title }].concat(
    migrated.filter((e) => recentSharePointPageKey(e.url) !== pageKey)
  );
  const next = deduped.slice(0, RECENT_SP_MAX);
  await chrome.storage.local.set({ [RECENT_SP_KEY]: next });
}

async function populateRecentSharePointDropdown() {
  const sel = document.getElementById("recentSharePointSelect");
  if (!sel) return;
  const { [RECENT_SP_KEY]: stored = [] } = await chrome.storage.local.get(RECENT_SP_KEY);
  let entries = migrateRecentSharePointStoredList(stored).filter((e) => isToolkitSharePointPage(e.url));
  try {
    const historyItems = await chrome.history.search({ text: "sharepoint.com", maxResults: 60, startTime: 0 });
    const seenPageKeys = new Set(entries.map((e) => recentSharePointPageKey(e.url)));
    for (const h of historyItems) {
      const u = h.url;
      if (!u || !isToolkitSharePointPage(u)) continue;
      const n = u.split("?")[0].replace(/\/$/, "") || u;
      const pk = recentSharePointPageKey(n);
      if (seenPageKeys.has(pk)) continue;
      seenPageKeys.add(pk);
      entries.push({ url: n, title: h.title || "" });
    }
  } catch (_) {}
  entries = entries.slice(0, RECENT_SP_MAX);
  sel.innerHTML = '<option value="">Choose a page…</option>';
  entries.forEach((entry) => {
    const opt = document.createElement("option");
    opt.value = entry.url;
    opt.textContent = formatRecentSharePointEntryLabel(entry);
    sel.appendChild(opt);
  });
}

async function goToSelectedSharePointSite() {
  const sel = document.getElementById("recentSharePointSelect");
  const url = sel && sel.value ? sel.value.trim() : "";
  if (!url) return;
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tab?.id) {
      chrome.tabs.update(tab.id, { url });
      window.close();
    }
  } catch (_) {}
}

async function syncPopupToCurrentTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  const onToolkitSp = isToolkitSharePointPage(tab?.url);

  if (!onToolkitSp) {
    document.body.classList.add("not-on-sharepoint");
    await populateRecentSharePointDropdown();
    return;
  }

  document.body.classList.remove("not-on-sharepoint");
  await addCurrentUrlToRecentSharePoint(tab.url, tab.title || "");

  const tabColumns = document.getElementById("tabColumns");
  const tabViewManager = document.getElementById("tabViewManager");
  const tabReports = document.getElementById("tabReports");
  let isListPage = false;
  try {
    const response = await chrome.tabs.sendMessage(tab.id, { action: "checkListPage" });
    isListPage = !!response?.isListPage;
    const show = isListPage ? "" : "none";
    if (tabColumns) tabColumns.style.display = show;
    if (tabViewManager) tabViewManager.style.display = show;
    if (tabReports) tabReports.style.display = show;
  } catch (_) {
    if (tabColumns) tabColumns.style.display = "none";
    if (tabViewManager) tabViewManager.style.display = "none";
    if (tabReports) tabReports.style.display = "none";
  }

  const { popupActiveTab } = await chrome.storage.local.get("popupActiveTab");
  let id = popupActiveTab && TAB_IDS.includes(popupActiveTab) ? popupActiveTab : "quicklinks";
  if ((id === "searchSchema" || id === "viewManager" || id === "reports") && !isListPage) {
    id = "quicklinks";
  }
  const tabEl = document.querySelector(".tab[data-tab=\"" + id + "\"]");
  if (tabEl && tabEl.offsetParent !== null) switchToTab(id);
  else switchToTab("quicklinks");
}

// Restore last active tab on open; when not on SharePoint, show message + recent sites only
(async function restorePopupTab() {
  await syncPopupToCurrentTab();
  try {
    const r = await chrome.storage.local.get(["universalSearchOpenOnLoad", "pendingOpenUniversalSearch"]);
    let fromHotkey = false;
    if (r.pendingOpenUniversalSearch) {
      fromHotkey = true;
      chrome.storage.local.remove("pendingOpenUniversalSearch");
    }
    if (r.universalSearchOpenOnLoad || fromHotkey) {
      requestAnimationFrame(() => openUniversalSearch());
    }
  } catch (_) {}
})();

document.getElementById("recentSharePointSelect")?.addEventListener("change", function () {
  if (this.value) goToSelectedSharePointSite();
});

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") syncPopupToCurrentTab();
});
window.addEventListener("pageshow", () => syncPopupToCurrentTab());
window.addEventListener("resize", () => { requestAnimationFrame(updateTabIndicator); });

function escapeHtml(s) {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}

function addQuickLink(ul, label, href, icon) {
  const li = document.createElement("li");
  const iconAttr = icon ? ` data-ql-icon="${escapeHtml(icon)}"` : "";
  li.innerHTML = `<a href="${escapeHtml(href)}" target="_blank" rel="noopener"${iconAttr}>${escapeHtml(label)}</a>`;
  ul.appendChild(li);
}

function isQuickLinkTenantAdminHost(href, adminHost) {
  if (!adminHost) return false;
  try {
    return new URL(href).hostname.toLowerCase() === String(adminHost).toLowerCase();
  } catch (_) {
    return false;
  }
}

/** Tenant *-admin.sharepoint.com links: new tab + loading overlay (background). */
function addQuickLinkOrTenantAdminWait(ul, label, href, icon, adminHost) {
  if (!isQuickLinkTenantAdminHost(href, adminHost)) {
    addQuickLink(ul, label, href, icon);
    return;
  }
  const li = document.createElement("li");
  const iconAttr = icon ? ` data-ql-icon="${escapeHtml(icon)}"` : "";
  li.innerHTML = `<a href="${escapeHtml(href)}" target="_blank" rel="noopener"${iconAttr}>${escapeHtml(label)}</a>`;
  const a = li.querySelector("a");
  a.addEventListener("click", (e) => {
    e.preventDefault();
    chrome.runtime.sendMessage({ type: "SPOToolkitOpenTenantAdminWithWait", url: href }, () => {
      void chrome.runtime.lastError;
    });
  });
  ul.appendChild(li);
}

function isTenantAdminQuickLinkHref(href) {
  try {
    const u = new URL(href);
    if (u.protocol !== "https:") return false;
    return u.hostname.toLowerCase().endsWith("-admin.sharepoint.com");
  } catch (_) {
    return false;
  }
}

/** Same tab: end-user RecycleBin first, then site-collection AdminRecycleBin (?view=5#view=13). */
function addSecondStageRecycleBinQuickLink(ul, label, tabId, siteBase, secondStageHref, icon) {
  const firstUrl = normalizeTrailingSlash(siteBase) + "/_layouts/15/RecycleBin.aspx";
  const li = document.createElement("li");
  const a = document.createElement("a");
  a.href = secondStageHref;
  a.target = "_blank";
  a.rel = "noopener";
  a.textContent = label;
  if (icon) a.setAttribute("data-ql-icon", icon);
  a.addEventListener("click", (e) => {
    e.preventDefault();
    if (tabId == null || !secondStageHref) return;
    chrome.runtime.sendMessage(
      {
        type: "SPOToolkitSecondStageRecycleBin",
        tabId,
        firstUrl,
        secondUrl: secondStageHref,
      },
      () => {
        void chrome.runtime.lastError;
      }
    );
    window.close();
  });
  li.appendChild(a);
  ul.appendChild(li);
}

function filterQuickLinks() {
  const input = document.getElementById("quicklinksFilter");
  const content = document.getElementById("quicklinksContent");
  if (!input || !content) return;
  const q = String(input.value || "").toLowerCase().trim();
  const sections = content.querySelectorAll(".ql-section");
  if (!q) {
    sections.forEach((section) => {
      section.style.display = "";
      section.querySelectorAll("li").forEach((li) => {
        li.style.display = "";
      });
    });
    return;
  }
  sections.forEach((section) => {
    const labelEl = section.querySelector(".ql-section-label");
    const labelText = labelEl ? (labelEl.textContent || "").toLowerCase() : "";
    const sectionLabelMatch = labelText.indexOf(q) >= 0;
    const lis = section.querySelectorAll("li");
    let any = false;
    lis.forEach((li) => {
      const a = li.querySelector("a");
      const text = a ? (a.textContent || "").toLowerCase() : "";
      const href = a ? (a.getAttribute("href") || "").toLowerCase() : "";
      const match = sectionLabelMatch || text.indexOf(q) >= 0 || href.indexOf(q) >= 0;
      li.style.display = match ? "" : "none";
      if (match) any = true;
    });
    section.style.display = any ? "" : "none";
  });
}

let renderQuickLinksBusy = false;

function waitForRenderQuickLinksIdle(timeoutMs) {
  const max = timeoutMs == null ? 10000 : timeoutMs;
  const start = typeof performance !== "undefined" ? performance.now() : Date.now();
  return new Promise((resolve) => {
    function tick() {
      if (!renderQuickLinksBusy) {
        resolve();
        return;
      }
      const elapsed = (typeof performance !== "undefined" ? performance.now() : Date.now()) - start;
      if (elapsed >= max) {
        resolve();
        return;
      }
      setTimeout(tick, 24);
    }
    tick();
  });
}

async function renderQuickLinks() {
  if (renderQuickLinksBusy) return;
  const hint = document.getElementById("quicklinksHint");
  const content = document.getElementById("quicklinksContent");
  const currentSite = document.getElementById("quicklinksCurrentSite");
  const currentUser = document.getElementById("quicklinksCurrentUser");
  const modes = document.getElementById("quicklinksModes");
  const tenant = document.getElementById("quicklinksTenant");
  if (!content || !currentSite || !currentUser || !modes || !tenant) return;
  renderQuickLinksBusy = true;
  try {
  currentSite.innerHTML = "";
  currentUser.innerHTML = "";
  modes.innerHTML = "";
  tenant.innerHTML = "";
  let tab = null;
  try {
    const [t] = await chrome.tabs.query({ active: true, currentWindow: true });
    tab = t;
    if (!isToolkitSharePointPage(tab?.url)) {
      hint.style.display = "block";
      content.style.display = "none";
      return;
    }
  } catch (_) {
    hint.style.display = "block";
    content.style.display = "none";
    return;
  }
  hint.style.display = "none";
  content.style.display = "block";
  const ctx = parseContextFromUrl(tab.url);
  const siteBase = (ctx.webAbsoluteUrl || "").replace(/\/$/, "") || new URL(tab.url).origin;
  const host = new URL(tab.url).host;
  const adminHost = host.replace(".sharepoint.com", "-admin.sharepoint.com");
  const siteId = await getCurrentSiteIdFromPage(tab.id, siteBase);
  const siteCollRoot = await resolveSiteCollectionRootUrl(tab.id, siteBase);
  const secondStageHref = secondStageRecycleBinUrl(siteCollRoot);

  addQuickLink(currentSite, "Site settings", siteBase + "/_layouts/15/settings.aspx", "settings");
  addQuickLink(currentSite, "Site contents", siteBase + "/_layouts/15/viewlsts.aspx", "folder");
  addQuickLink(currentSite, "Site content types", siteBase + "/_layouts/15/SiteAdmin.aspx#/contentTypes", "layers");
  addQuickLinkOrTenantAdminWait(
    currentSite,
    "Tenant content types",
    `https://${adminHost}/_layouts/15/online/AdminHome.aspx#/contentTypes`,
    "adminHome",
    adminHost
  );
  addQuickLink(currentSite, "Recycle bin", siteBase + "/_layouts/15/RecycleBin.aspx", "trash");
  addSecondStageRecycleBinQuickLink(currentSite, "Second Stage Recycle Bin", tab.id, siteBase, secondStageHref, "trash");
  addQuickLink(currentSite, "All People", siteBase + "/_layouts/15/people.aspx?MembershipGroupId=0", "users");
  addQuickLink(currentSite, "Storage metrics", siteBase + "/_layouts/15/storman.aspx", "chart");
  addQuickLinkOrTenantAdminWait(
    currentSite,
    "SharePoint Admin Settings",
    siteId
      ? `https://${adminHost}/_layouts/15/online/AdminHome.aspx#/siteManagement/:/SiteDetails/${siteId}`
      : `https://${adminHost}/_layouts/15/online/AdminHome.aspx#/siteManagement`,
    "adminHome",
    adminHost
  );

  addQuickLink(currentUser, "Edit user profile", siteBase + "/_layouts/15/me.aspx", "user");
  addQuickLink(currentUser, "Login as another user", siteBase + "/_layouts/15/closeConnection.aspx?loginasanotheruser=1", "logout");

  const pageUrl = tab.url.split("?")[0];
  const q = (param) => pageUrl + (pageUrl.indexOf("?") >= 0 ? "&" : "?") + param;
  addQuickLink(modes, "MaintenanceMode", q("MaintenanceMode=true"), "wrench");
  addQuickLink(modes, "WebView", q("env=WebView"), "monitor");
  addQuickLink(modes, "WebViewList", q("env=WebViewList"), "list");
  addQuickLink(modes, "Disable SPFx code", q("disable3PCode"), "shield");
  addQuickLink(modes, "Web Part Maintenance", q("contents=1"), "layers");

  addQuickLinkOrTenantAdminWait(tenant, "Admin center", `https://${adminHost}`, "shield", adminHost);
  addQuickLinkOrTenantAdminWait(
    tenant,
    "Admin Center Settings",
    siteId
      ? `https://${adminHost}/_layouts/15/online/AdminHome.aspx#/siteManagement/:/SiteDetails/${siteId}/Settings`
      : `https://${adminHost}/_layouts/15/online/AdminHome.aspx#/siteManagement`,
    "settings",
    adminHost
  );
  addQuickLinkOrTenantAdminWait(
    tenant,
    "Tenant site settings",
    `https://${adminHost}/_layouts/15/online/tenantsettings.aspx`,
    "settings",
    adminHost
  );
  addQuickLinkOrTenantAdminWait(
    tenant,
    "User profiles",
    `https://${adminHost}/_layouts/15/TenantProfileAdmin/ManageUserProfileServiceApplication.aspx`,
    "users",
    adminHost
  );
  addQuickLink(tenant, "Term store", `https://${host}/_layouts/15/termstoremanager.aspx`, "tag");
  addQuickLinkOrTenantAdminWait(
    tenant,
    "Search administration",
    `https://${adminHost}/_layouts/15/searchadmin/TA_SearchAdministration.aspx`,
    "search",
    adminHost
  );
  addQuickLinkOrTenantAdminWait(
    tenant,
    "API access",
    `https://${adminHost}/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement`,
    "key",
    adminHost
  );
  addQuickLink(tenant, "Teams admin", "https://admin.teams.microsoft.com/dashboard", "teams");
  addQuickLink(tenant, "App catalog", `https://${host}/sites/AppCatalog/_layouts/15/appStore.aspx`, "package");
  addQuickLink(tenant, "Classic app catalog", `https://${host}/sites/AppCatalog`, "archive");

  } finally {
    renderQuickLinksBusy = false;
  }
}

function normalizeGuidString(raw) {
  if (!raw) return "";
  const s = String(raw).trim().replace(/[{}]/g, "");
  return /^[0-9a-fA-F-]{36}$/.test(s) ? s : "";
}

function findSiteIdDeep(obj, depth) {
  if (!obj || typeof obj !== "object" || depth > 6) return "";
  if (Array.isArray(obj)) {
    for (let i = 0; i < obj.length; i++) {
      const hit = findSiteIdDeep(obj[i], depth + 1);
      if (hit) return hit;
    }
    return "";
  }
  const directKeys = ["siteId", "SiteId", "siteID", "SiteID"];
  for (let i = 0; i < directKeys.length; i++) {
    const hit = normalizeGuidString(obj[directKeys[i]]);
    if (hit) return hit;
  }
  for (const k of Object.keys(obj)) {
    if (/^siteid$/i.test(k)) {
      const hit = normalizeGuidString(obj[k]);
      if (hit) return hit;
    }
  }
  for (const k of Object.keys(obj)) {
    const hit = findSiteIdDeep(obj[k], depth + 1);
    if (hit) return hit;
  }
  return "";
}

function sitePathFromWebAbsoluteUrl(webAbsoluteUrl) {
  if (!webAbsoluteUrl) return "";
  try {
    return new URL(webAbsoluteUrl).pathname.replace(/\/$/, "") || "/";
  } catch (_) {
    return "";
  }
}

/** Site collection root web absolute URL (second-stage recycle bin lives here, not on subsites). */
async function getSiteCollectionRootAbsoluteUrl(tabId, webAbsoluteUrl) {
  const sitePath = sitePathFromWebAbsoluteUrl(webAbsoluteUrl);
  if (!tabId || !sitePath) return "";
  return new Promise((resolve) => {
    try {
      chrome.tabs.sendMessage(
        tabId,
        { action: "rest", method: "GET", path: sitePath + "/_api/site/rootweb?$select=Url" },
        (res) => {
          if (chrome.runtime.lastError || !res || !res.ok || !res.data) {
            resolve("");
            return;
          }
          const d = res.data || {};
          const url =
            (typeof d.Url === "string" && d.Url.trim()) ||
            (typeof d.url === "string" && d.url.trim()) ||
            "";
          resolve(url ? url.replace(/\/$/, "") : "");
        }
      );
    } catch (_) {
      resolve("");
    }
  });
}

/** Prefer _spPageContextInfo.siteAbsoluteUrl from the tab; then REST rootweb; then current web URL. */
async function resolveSiteCollectionRootUrl(tabId, webAbsoluteUrl) {
  const fallback = normalizeTrailingSlash(webAbsoluteUrl);
  if (!tabId) return fallback;
  const fromPage = await new Promise((resolve) => {
    try {
      chrome.tabs.sendMessage(tabId, { action: "getPageContext" }, (res) => {
        if (chrome.runtime.lastError || !res || !res.ok) {
          resolve("");
          return;
        }
        const s = (res.siteAbsoluteUrl || "").trim();
        resolve(s ? normalizeTrailingSlash(s) : "");
      });
    } catch (_) {
      resolve("");
    }
  });
  if (fromPage) return fromPage;
  const fromRest = await getSiteCollectionRootAbsoluteUrl(tabId, webAbsoluteUrl);
  if (fromRest) return fromRest;
  return fallback;
}

async function getCurrentSiteIdFromPage(tabId, webAbsoluteUrl) {
  if (!tabId) return "";
  const fromContext = await new Promise((resolve) => {
    try {
      chrome.tabs.sendMessage(tabId, { action: "getPageContext" }, (res) => {
        if (chrome.runtime.lastError || !res || !res.ok) {
          resolve("");
          return;
        }
        resolve(normalizeGuidString(res.siteId));
      });
    } catch (_) {
      resolve("");
    }
  });
  if (fromContext) return fromContext;

  const fromJson = await new Promise((resolve) => {
    try {
      chrome.tabs.sendMessage(tabId, { action: "getPageContextJson" }, (res) => {
        if (chrome.runtime.lastError || !res || !res.ok || !res.data) {
          resolve("");
          return;
        }
        resolve(findSiteIdDeep(res.data, 0));
      });
    } catch (_) {
      resolve("");
    }
  });
  if (fromJson) return fromJson;

  const sitePath = sitePathFromWebAbsoluteUrl(webAbsoluteUrl);
  if (!sitePath) return "";

  return new Promise((resolve) => {
    try {
      chrome.tabs.sendMessage(tabId, { action: "rest", method: "GET", path: sitePath + "/_api/site/id" }, (res) => {
        if (chrome.runtime.lastError || !res || !res.ok || !res.data) {
          resolve("");
          return;
        }
        const d = res.data || {};
        resolve(normalizeGuidString(d.value || d.Id || (d.d && d.d.Id)));
      });
    } catch (_) {
      resolve("");
    }
  });
}

let contextData = null;

/** Flatten nested object to single-level keys (e.g. Web.WebAbsoluteUrl) for display. */
function flattenContext(obj, prefix) {
  if (obj === null || typeof obj !== "object") return prefix ? { [prefix]: obj } : {};
  const out = {};
  for (const k of Object.keys(obj)) {
    const v = obj[k];
    const key = prefix ? prefix + "." + k : k;
    if (v !== null && typeof v === "object" && !Array.isArray(v) && Object.prototype.toString.call(v) === "[object Object]") {
      Object.assign(out, flattenContext(v, key));
    } else {
      out[key] = v !== null && typeof v === "object" ? JSON.stringify(v) : v;
    }
  }
  return out;
}

async function loadContextInfo() {
  const list = document.getElementById("contextList");
  const countEl = document.getElementById("contextCount");
  list.innerHTML = "";
  countEl.textContent = "ContextInfo properties: …";
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !isToolkitSharePointPage(tab.url)) {
      countEl.textContent = "ContextInfo properties: 0";
      list.innerHTML = "<p class=\"hint\">Open a SharePoint page, then click Refresh to load page properties.</p>";
      return;
    }
    // Try _spPageContextInfo in all frames (classic list/library and some modern iframes)
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id, allFrames: true },
      func: () => (typeof window._spPageContextInfo !== "undefined" ? window._spPageContextInfo : null),
      world: "MAIN"
    });
    let ctx = results && results.length > 0
      ? results.map((r) => r.result).find((r) => r && typeof r === "object" && Object.keys(r).length > 0)
      : null;
    contextData = ctx && typeof ctx === "object" ? ctx : null;
    // Modern pages (e.g. home) often don't set _spPageContextInfo; fetch ?as=json in page context
    if (!contextData || Object.keys(contextData).length === 0) {
      try {
        const res = await new Promise((resolve) => {
          chrome.tabs.sendMessage(tab.id, { action: "getPageContextJson" }, (r) => {
            if (chrome.runtime.lastError) resolve({ ok: false, error: chrome.runtime.lastError.message });
            else resolve(r || { ok: false, error: "No response" });
          });
        });
        if (res.ok && res.data && typeof res.data === "object" && Object.keys(res.data).length > 0) {
          contextData = flattenContext(res.data);
        }
      } catch (_) {}
    }
    renderContextList();
  } catch (e) {
    countEl.textContent = "ContextInfo properties: 0";
    list.innerHTML = "<p class=\"hint\">Error: " + escapeHtml(e.message || String(e)) + ". Refresh the SharePoint tab and try again.</p>";
  }
}

function renderContextList() {
  const list = document.getElementById("contextList");
  const countEl = document.getElementById("contextCount");
  const filterInput = document.getElementById("contextFilterInput");
  const filterByValue = document.getElementById("contextFilterByValue");
  const filter = (filterInput && filterInput.value) ? (filterInput.value || "").toLowerCase() : "";
  const byValue = filterByValue && filterByValue.checked;
  list.innerHTML = "";
  if (!contextData) {
    countEl.textContent = "ContextInfo properties: 0";
    list.innerHTML = "<p class=\"hint\">Click Refresh to load page properties from the current SharePoint page.</p>";
    return;
  }
  const entries = Object.entries(contextData);
  const filtered = filter
    ? entries.filter(([k, v]) =>
        byValue
          ? String(v).toLowerCase().indexOf(filter) >= 0
          : k.toLowerCase().indexOf(filter) >= 0
      )
    : entries;
  countEl.textContent = "ContextInfo properties: " + filtered.length;

  // Data-URI icons (white on blue for copy to match sp-editor / Fluent)
  const COPY_ICON_DATA_URI = "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxNiAxNiI+PHBhdGggZmlsbD0iI2ZmZiIgZD0iTTEzLjUgMWgtOUExLjUgMS41IDAgMCAwIDMgMi41djlBMS41IDEuNSAwIDAgMCA0LjUgMTNoOWExLjUgMS41IDAgMCAwIDEuNS0xLjV2LTlhMS41IDEuNSAwIDAgMC0xLjUtMS41em0tOSAxaDlhLjUuNSAwIDAgMSAuNS41djlhLjUuNSAwIDAgMS0uNS41aC05YS41LjUgMCAwIDEtLjUtLjV2LTlhLjUuNSAwIDAgMSAuNS0uNXoiLz48cGF0aCBmaWxsPSIjZmZmIiBkPSJNMTEuNSA1aC03YS41LjUgMCAwIDAtLjUuNXY3YS41LjUgMCAwIDAgLjUuNWg3YS41LjUgMCAwIDAgLjUtLjV2LTdhLjUuNSAwIDAgMC0uNS0uNXoiLz48L3N2Zz4=";
  const CHECK_ICON_DATA_URI = "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxNiAxNiI+PHBhdGggZmlsbD0iI2ZmZiIgZD0iTTE0LjQzMSAzLjMyM2wtOC40NyAxMC0uNzktLjAzNi0zLjM1LTQuNzcuODE4LS41NzQgMi45NzggNC4yNCA4LjA1MS05LjUwNi43NjQuNjQ2eiIvPjwvc3ZnPg==";

  function createCopyIconImg() {
    const img = document.createElement("img");
    img.src = COPY_ICON_DATA_URI;
    img.alt = "";
    img.setAttribute("aria-hidden", "true");
    return img;
  }

  function createCheckIconImg() {
    const img = document.createElement("img");
    img.src = CHECK_ICON_DATA_URI;
    img.alt = "";
    img.setAttribute("aria-hidden", "true");
    return img;
  }

  filtered.forEach(([name, value]) => {
    const valStr = value === null || value === undefined ? "null" : String(value);
    const row = document.createElement("div");
    row.className = "context-row";
    const nameSpan = document.createElement("span");
    nameSpan.className = "name";
    nameSpan.textContent = name;
    const valueRow = document.createElement("div");
    valueRow.className = "context-value-row";
    const valueBox = document.createElement("div");
    valueBox.className = "value-box";
    valueBox.textContent = valStr;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "copy-btn";
    btn.title = "Copy value";
    btn.appendChild(createCopyIconImg());
    btn.addEventListener("click", () => {
      navigator.clipboard.writeText(normalizeGuidCopyText(valStr)).then(() => {
        btn.textContent = "";
        btn.className = "copy-btn copied";
        btn.appendChild(createCheckIconImg());
        btn.setAttribute("aria-label", "Copied");
        setTimeout(() => {
          btn.textContent = "";
          btn.className = "copy-btn";
          btn.appendChild(createCopyIconImg());
          btn.setAttribute("aria-label", "Copy value");
        }, 800);
      });
    });
    valueRow.appendChild(valueBox);
    valueRow.appendChild(btn);
    row.appendChild(nameSpan);
    row.appendChild(valueRow);
    list.appendChild(row);
  });
}

document.getElementById("btnRefreshContext").addEventListener("click", () => loadContextInfo());
const quicklinksFilterEl = document.getElementById("quicklinksFilter");
if (quicklinksFilterEl) quicklinksFilterEl.addEventListener("input", () => filterQuickLinks());
document.getElementById("contextFilterInput").addEventListener("input", () => renderContextList());
document.getElementById("contextFilterByValue").addEventListener("change", () => renderContextList());

document.getElementById("btnOpenViewManager").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id || !isToolkitSharePointPage(tab.url)) {
    alert("Open a SharePoint list or library page first, then open View Manager.");
    return;
  }
  chrome.tabs.create({ url: chrome.runtime.getURL("views.html?tabId=" + tab.id) });
});

document.getElementById("btnOpenViewFormatter")?.addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  openViewFormatterForTab(tab, true);
});

// --- Columns tab: list fields via REST; display name links to FldEdit.aspx
let searchSchemaColumns = [];
let searchSchemaListMeta = { siteUrl: "", listId: "" };

async function loadSearchSchema() {
  const countEl = document.getElementById("searchSchemaCount");
  const listEl = document.getElementById("searchSchemaList");
  const errEl = document.getElementById("searchSchemaError");
  if (!listEl || !errEl) return;
  errEl.style.display = "none";
  listEl.innerHTML = "";
  searchSchemaListMeta = { siteUrl: "", listId: "" };
  if (countEl) countEl.textContent = "Loading…";
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !isToolkitSharePointPage(tab.url)) {
      if (countEl) countEl.textContent = "Columns: 0";
      errEl.textContent = "Open a list or library page to load columns.";
      errEl.style.display = "block";
      return;
    }
    const response = await chrome.tabs.sendMessage(tab.id, { action: "getSearchSchema" });
    if (!response?.ok) {
      searchSchemaColumns = [];
      if (countEl) countEl.textContent = "Columns: 0";
      errEl.textContent = response?.error || "Failed to load list columns.";
      errEl.style.display = "block";
      return;
    }
    searchSchemaColumns = response.columns || [];
    searchSchemaListMeta = searchSchemaListMetaFromResponse(response);
    renderSearchSchemaList();
  } catch (e) {
    searchSchemaColumns = [];
    searchSchemaListMeta = { siteUrl: "", listId: "" };
    if (countEl) countEl.textContent = "Columns: 0";
    errEl.textContent = "Error: " + (e.message || "Could not reach page. Try refreshing the SharePoint tab.");
    errEl.style.display = "block";
  }
}

function renderSearchSchemaList() {
  const countEl = document.getElementById("searchSchemaCount");
  const listEl = document.getElementById("searchSchemaList");
  const filterInput = document.getElementById("searchSchemaFilter");
  if (!listEl) return;
  const q = (filterInput && filterInput.value) ? String(filterInput.value).toLowerCase().trim() : "";
  const filtered = q ? searchSchemaColumns.filter((c) => searchSchemaColumnMatchesFilter(c, q)) : searchSchemaColumns;
  if (countEl) countEl.textContent = "Columns: " + filtered.length;
  listEl.innerHTML = "";
  filtered.forEach((c) => {
    const row = document.createElement("div");
    row.className = "search-schema-row";
    const displayTitle = c.title || c.internalName || "";
    const settingsHref = listColumnSettingsUrl(
      searchSchemaListMeta.siteUrl,
      searchSchemaListMeta.listId,
      c.internalName || ""
    );
    const nameCell = settingsHref
      ? '<a class="col-name-link" href="' + escapeHtml(settingsHref) + '" target="_blank" rel="noopener" title="Open column settings">' + escapeHtml(displayTitle) + "</a>"
      : escapeHtml(displayTitle);
    row.innerHTML =
      '<span class="col-name">' + nameCell +
      '</span><span class="col-internal">' + escapeHtml(c.internalName || "") +
      '</span><span class="col-crawled">' + escapeHtml(c.crawledProperty || "") +
      '</span><span class="col-type">' + escapeHtml(c.type || "") + '</span>';
    listEl.appendChild(row);
  });
}

document.getElementById("btnRefreshSearchSchema").addEventListener("click", () => loadSearchSchema());
document.getElementById("searchSchemaFilter").addEventListener("input", () => renderSearchSchemaList());

async function getPageSize() {
  const st = await chrome.storage.local.get("pageSize");
  const v = parseInt(st.pageSize, 10);
  return (v === 500 || v === 1000 || v === 5000) ? v : DEFAULT_PAGE_SIZE;
}

async function getDefaultExportFormat() {
  const st = await chrome.storage.local.get("defaultExportFormat");
  return (st.defaultExportFormat === "csv" || st.defaultExportFormat === "xlsx") ? st.defaultExportFormat : DEFAULT_EXPORT_FORMAT;
}

async function loadSettings() {
  const size = await getPageSize();
  pageSizeSelect.value = String(size);
  const format = await getDefaultExportFormat();
  if (defaultExportFormatSelect) defaultExportFormatSelect.value = format;
}

function savePageSize() {
  const v = parseInt(pageSizeSelect.value, 10);
  chrome.storage.local.set({ pageSize: v });
}

function saveDefaultExportFormat() {
  const format = defaultExportFormatSelect && (defaultExportFormatSelect.value === "csv" || defaultExportFormatSelect.value === "xlsx")
    ? defaultExportFormatSelect.value
    : DEFAULT_EXPORT_FORMAT;
  chrome.storage.local.set({ defaultExportFormat: format });
}

function setStatus(message, type) {
  const textEl = statusEl.querySelector(".status-text");
  const spinnerEl = statusEl.querySelector(".status-spinner");
  const successIconEl = statusEl.querySelector(".status-success-icon");
  if (textEl) textEl.textContent = message;
  statusEl.className = type || "info";
  statusEl.style.display = "block";
  if (spinnerEl) spinnerEl.style.display = type === "info" ? "block" : "none";
  if (successIconEl) successIconEl.style.display = type === "success" ? "inline-flex" : "none";
}

function setPickerStatus(message, type) {
  pickerStatus.textContent = message;
  pickerStatus.className = type || "info";
  pickerStatus.style.display = message ? "block" : "none";
}

function clearStatus() {
  statusEl.style.display = "none";
  statusEl.className = "";
}

// --- Refinable Props tab ---
const refinableTypeSelect = document.getElementById("refinableTypeSelect");
const refinableNumberSelect = document.getElementById("refinableNumberSelect");
const btnRefinableConfigure = document.getElementById("btnRefinableConfigure");
const refinableMappingsOut = document.getElementById("refinableMappingsOut");

function initRefinableProps() {
  if (!refinableTypeSelect || !refinableNumberSelect) return;
  if (refinableTypeSelect.options.length === 0) {
    REFINABLE_TYPES.forEach((t) => {
      const opt = document.createElement("option");
      opt.value = t.id;
      opt.textContent = t.label;
      refinableTypeSelect.appendChild(opt);
    });
    refinableTypeSelect.value = "RefinableString";
  }
  fillRefinableNumberOptions();
  loadRefinableMappings();
}

function fillRefinableNumberOptions() {
  if (!refinableTypeSelect || !refinableNumberSelect) return;
  const typeId = refinableTypeSelect.value;
  const type = REFINABLE_TYPES.find((t) => t.id === typeId) || REFINABLE_TYPES[0];
  const count = type.count;
  refinableNumberSelect.innerHTML = "";
  for (let i = 0; i < count; i++) {
    const opt = document.createElement("option");
    const num = String(i).padStart(2, "0");
    opt.value = num;
    opt.textContent = num;
    refinableNumberSelect.appendChild(opt);
  }
}

refinableTypeSelect?.addEventListener("change", function () {
  fillRefinableNumberOptions();
  loadRefinableMappings();
});
refinableNumberSelect?.addEventListener("change", loadRefinableMappings);

function clearRefinableMappingsDisplay() {
  if (refinableMappingsOut) {
    refinableMappingsOut.textContent = "";
    refinableMappingsOut.className = "refinable-mappings-out empty";
  }
}

btnRefinableConfigure?.addEventListener("click", async () => {
  const typeId = refinableTypeSelect?.value || "RefinableString";
  const num = refinableNumberSelect?.value || "00";
  const propertyName = typeId + num;
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!isToolkitSharePointPage(tab?.url)) {
      alert("Open a SharePoint site in the current tab first, then click Configure to open the managed property page for that site.");
      return;
    }
    const ctx = parseContextFromUrl(tab.url);
    const siteUrl = (ctx.webAbsoluteUrl || "").replace(/\/$/, "");
    if (!siteUrl) {
      alert("Could not determine the site URL. Open a SharePoint site page and try again.");
      return;
    }
    const url = siteUrl + "/_layouts/15/managedproperty.aspx?property=" + encodeURIComponent(propertyName) + "&level=sitecol";
    chrome.tabs.create({ url });
  } catch (e) {
    alert("Error: " + (e.message || String(e)));
  }
});

async function loadRefinableMappings() {
  if (!refinableMappingsOut) return;
  const typeId = refinableTypeSelect?.value || "RefinableString";
  const num = refinableNumberSelect?.value || "00";
  const propertyName = typeId + num;
  refinableMappingsOut.textContent = "Loading…";
  refinableMappingsOut.className = "refinable-mappings-out loading";
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !isToolkitSharePointPage(tab.url)) {
      refinableMappingsOut.textContent = "Open a SharePoint site in the current tab to load mappings.";
      refinableMappingsOut.className = "refinable-mappings-out error";
      return;
    }
    const ctx = parseContextFromUrl(tab.url);
    const siteUrl = (ctx.webAbsoluteUrl || "").replace(/\/$/, "");
    if (!siteUrl) {
      refinableMappingsOut.textContent = "Could not determine site URL. Open a SharePoint site page first.";
      refinableMappingsOut.className = "refinable-mappings-out error";
      return;
    }
    const res = await chrome.tabs.sendMessage(tab.id, {
      action: "getRefinableMappings",
      siteUrl,
      propertyName
    });
    if (!res) {
      refinableMappingsOut.textContent = "No response. Refresh the SharePoint tab and try again.";
      refinableMappingsOut.className = "refinable-mappings-out error";
      return;
    }
    if (!res.ok) {
      refinableMappingsOut.textContent = res.error || "Failed to load mappings.";
      refinableMappingsOut.className = "refinable-mappings-out error";
      return;
    }
    if (!res.mappings || res.mappings.length === 0) {
      refinableMappingsOut.innerHTML = `<span class="refinable-mappings-header">Crawled properties</span>${res.alias ? `<p class="refinable-alias">Alias: ${escapeHtml(res.alias)}</p>` : ""}<p>No crawled properties mapped for ${escapeHtml(propertyName)}.</p>`;
      refinableMappingsOut.className = "refinable-mappings-out";
      return;
    }
    const aliasPart = res.alias ? `<p class="refinable-alias">Alias: ${escapeHtml(res.alias)}</p>` : "";
    const listItems = res.mappings.map((m) => `<li>${escapeHtml(m.name)}${m.type ? `<span class="refinable-mapping-type">${escapeHtml(m.type)}</span>` : ""}</li>`).join("");
    refinableMappingsOut.innerHTML = `${aliasPart}<span class="refinable-mappings-header">Crawled properties (${res.mappings.length})</span><ul class="refinable-mappings-list">${listItems}</ul>`;
    refinableMappingsOut.className = "refinable-mappings-out";
  } catch (e) {
    refinableMappingsOut.textContent = "Error: " + (e.message || String(e)) + ". Try refreshing the SharePoint tab.";
    refinableMappingsOut.className = "refinable-mappings-out error";
  }
}

// Run once so number dropdown is populated if user opens Refinable Props first
initRefinableProps();

function parseContextFromUrl(pageUrl) {
  if (!pageUrl || !isToolkitSharePointPage(pageUrl)) return { webAbsoluteUrl: "", pageListId: "", viewId: "" };
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
      if (["sites", "site", "teams"].includes(segments[i]) && segments[i + 1]) {
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

// SharePoint system field names (same on all lists/libraries), used for column picker defaults.
const REQUIRED_FIELDS = ["ID", "Created", "Modified", "Author", "Editor", "_UIVersionString"];

function renderColumnPicker() {
  if (!pickerData?.pickerFields) return;
  columnList.innerHTML = "";
  const defaultSet = new Set(pickerData.defaultColumns || []);
  const requiredSet = new Set(REQUIRED_FIELDS);

  for (const f of pickerData.pickerFields) {
    const iname = f.InternalName || f.Title;
    const title = f.Title || iname;
    const isDefault = defaultSet.has(iname);
    const isRequired = requiredSet.has(iname);

    const div = document.createElement("div");
    div.className = "column-item";
    const id = "col-" + iname.replace(/[^a-zA-Z0-9]/g, "_");
    div.innerHTML =
      '<input type="checkbox" id="' + id + '" data-internal="' + iname.replace(/"/g, "&quot;") + '" ' + (isDefault ? "checked" : "") + '>' +
      '<label for="' + id + '">' + (title !== iname ? title + " <span class=\"required-badge\">(" + iname + ")</span>" : iname) + (isRequired ? " *" : "") + "</label>";
    columnList.appendChild(div);
  }
}

function getSelectedColumns() {
  const checkboxes = columnList.querySelectorAll('input[type="checkbox"]:checked');
  return Array.from(checkboxes).map((cb) => cb.dataset.internal);
}

function getIncludeVersions() {
  const el = document.getElementById("chkIncludeVersionsPicker");
  return el ? el.checked : false;
}

async function runExport(selectedColumns = null) {
  clearStatus();
  setPickerStatus("");
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !isToolkitSharePointPage(tab.url)) {
      setStatus(tab?.id ? "Open a SharePoint list or library page first." : "No active tab.", "error");
      return;
    }

    const ctx = parseContextFromUrl(tab.url);
    const siteUrl = (ctx.webAbsoluteUrl || "").replace(/\/$/, "");
    const listId = ctx.pageListId || "";
    const viewId = ctx.viewId || "";

    setStatus("Starting export…", "info");
    let response;
    try {
      const includeVersions = getIncludeVersions();
      const pageLimit = await getPageSize();
      const reportType = reportSelect && reportSelect.value ? reportSelect.value : null;
      const exportFormat = (formatSelect && formatSelect.value) ? formatSelect.value : "xlsx";
      const chkMatrixWholeSite = document.getElementById("chkMatrixWholeSite");
      const msg = {
        action: "runExportCSV",
        siteUrl,
        listId,
        viewId,
        exportFilename: "", // Extension builds: sitename_listname_datetime.ext
        pageLimit,
        includeVersions,
        selectedColumns: selectedColumns && selectedColumns.length > 0 ? selectedColumns : null,
        report: reportType,
        format: exportFormat,
        permissionsMatrixWholeSite: !!(chkMatrixWholeSite && chkMatrixWholeSite.checked)
      };
      response = await chrome.tabs.sendMessage(tab.id, msg);
    } catch (sendErr) {
      setStatus(
        "Could not connect to page. Try refreshing the SharePoint tab (F5), then click Export again. Error: " + sendErr.message,
        "error"
      );
      return;
    }

    const statusMessage = (response?.message != null && response.message !== "")
      ? response.message
      : (response?.ok ? "Export started. Watch the progress bar at the top of the page for status." : (response?.error || "Export failed."));
    setStatus(statusMessage, response?.ok ? "info" : "error");
    if (response?.ok) {
      mainPanel.classList.remove("hidden");
      pickerPanel.classList.remove("visible");
    }
  } catch (e) {
    setStatus("Error: " + e.message, "error");
  }
}

async function openColumnPicker() {
  clearStatus();
  mainPanel.classList.add("hidden");
  pickerPanel.classList.add("visible");
  document.body.classList.add("reports-expanded");
  setPickerStatus("Loading columns…", "info");
  columnList.innerHTML = "";

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !isToolkitSharePointPage(tab.url)) {
      setPickerStatus("Open a SharePoint list or library page first.", "error");
      return;
    }

    const response = await chrome.tabs.sendMessage(tab.id, { action: "getColumns" });
    if (!response?.ok) {
      setPickerStatus(response?.error || "Failed to load columns.", "error");
      return;
    }

    pickerData = response;
    renderColumnPicker();
    setPickerStatus("");
  } catch (e) {
    setPickerStatus("Error: " + e.message + ". Try refreshing the SharePoint tab.", "error");
  }
}

function closeColumnPicker() {
  pickerPanel.classList.remove("visible");
  mainPanel.classList.remove("hidden");
  document.body.classList.remove("reports-expanded");
  setPickerStatus("");
}

function toggleReportOptions() {
  if (!reportSelect || !reportOptions) return;
  const value = reportSelect.value;
  const showPanel = value === "exportCSV" || value === "folderCount" || value === "pathLengths" || value === "permissionsMatrix";
  reportOptions.style.display = showPanel ? "block" : "none";
  const matrixWholeSiteRow = document.getElementById("matrixWholeSiteRow");
  if (matrixWholeSiteRow) matrixWholeSiteRow.style.display = value === "permissionsMatrix" ? "" : "none";
  const exportOptionsRow = document.getElementById("exportOptionsRow");
  if (exportOptionsRow) exportOptionsRow.style.display = value === "exportCSV" ? "" : "none";
  if (btnChooseColumns) btnChooseColumns.style.display = value === "exportCSV" ? "" : "none";
}
if (reportSelect) reportSelect.addEventListener("change", toggleReportOptions);

/** Updates the "Export List/Library to Excel/CSV" option text and syncs format select to default. */
async function updateExportReportLabel() {
  const optionEl = document.getElementById("reportOptionExportCSV");
  if (!optionEl || !formatSelect) return;
  const format = await getDefaultExportFormat();
  formatSelect.value = format;
  const formatLabel = format === "csv" ? "CSV" : "Excel";
  let listOrLibrary = "List";
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !isToolkitSharePointPage(tab.url)) {
      optionEl.textContent = "Export " + listOrLibrary + " to " + formatLabel;
      return;
    }

    // Prefer content script + getListType.js (same as export flow) – works in iframes and modern pages.
    try {
      const csResult = await new Promise((resolve) => {
        chrome.tabs.sendMessage(tab.id, { action: "getListType" }, (response) => {
          if (chrome.runtime.lastError) resolve(null);
          else resolve(response);
        });
      });
      if (csResult && csResult.isListPage === true) {
        listOrLibrary = csResult.isLibrary ? "Library" : "List";
        optionEl.textContent = "Export " + listOrLibrary + " to " + formatLabel;
        return;
      }
    } catch (_) {}

    const ctx = parseContextFromUrl(tab.url);
    const siteUrl = (ctx.webAbsoluteUrl || "").replace(/\/$/, "");
    const listIdFromUrl = (ctx.pageListId || "").replace(/[{}]/g, "").trim();

    // Strategy 1: When URL has List= param, fetch BaseTemplate directly in main frame (same as export).
    // This works regardless of which frame has _spPageContextInfo.
    if (siteUrl && listIdFromUrl) {
      try {
        const inj = await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: function (siteUrlArg, listIdArg) {
            return fetch(siteUrlArg + "/_api/web/lists(guid'" + listIdArg + "')?$select=BaseTemplate", {
              credentials: "include",
              headers: { "Accept": "application/json;odata=nometadata" }
            })
              .then(function (r) { return r.json(); })
              .then(function (j) {
                var bt = j.BaseTemplate != null ? j.BaseTemplate : (j.d && j.d.BaseTemplate != null ? j.d.BaseTemplate : null);
                var n = parseInt(bt, 10);
                return { isListPage: true, isLibrary: n === 101 };
              })
              .catch(function () { return { isListPage: true, isLibrary: false }; });
          },
          args: [siteUrl, listIdFromUrl],
          world: "MAIN"
        });
        if (inj && inj[0] && inj[0].result && inj[0].result.isListPage === true) {
          listOrLibrary = inj[0].result.isLibrary ? "Library" : "List";
          optionEl.textContent = "Export " + listOrLibrary + " to " + formatLabel;
          return;
        }
      } catch (_) {}
    }

    // Strategy 2: No List= in URL or strategy 1 failed – run in all frames using _spPageContextInfo.
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: function () {
        var spi = window._spPageContextInfo || {};
        var listUrl = spi.listUrl || spi.listServerRelativeUrl;
        if (!listUrl) return Promise.resolve({ isListPage: false, isLibrary: false });
        var siteUrl = (spi.webAbsoluteUrl || spi.siteAbsoluteUrl || "").replace(/\/$/, "");
        var listId = (spi.pageListId || "").replace(/[{}]/g, "").trim();
        var enc = encodeURIComponent("'" + listUrl.replace(/'/g, "''") + "'");
        var opts = { credentials: "include", headers: { "Accept": "application/json;odata=nometadata" } };
        function getListId() {
          if (listId) return Promise.resolve(listId);
          return fetch(siteUrl + "/_api/web/GetList(@u)/Id?@u=" + enc, opts).then(function (r) { return r.json(); })
            .then(function (j) {
              var id = (j.value != null ? j.value : (j.d && j.d.Id) != null ? j.d.Id : j.Id || "");
              return ("" + id).replace(/[{}]/g, "").trim();
            });
        }
        return getListId().then(function (lid) {
          if (!lid) return { isListPage: true, isLibrary: false };
          return fetch(siteUrl + "/_api/web/lists(guid'" + lid + "')?$select=BaseTemplate", opts)
            .then(function (r) { return r.json(); })
            .then(function (j) {
              var bt = j.BaseTemplate != null ? j.BaseTemplate : (j.d && j.d.BaseTemplate != null ? j.d.BaseTemplate : null);
              var n = parseInt(bt, 10);
              return { isListPage: true, isLibrary: n === 101 };
            })
            .catch(function () { return { isListPage: true, isLibrary: false }; });
        }).catch(function () { return { isListPage: true, isLibrary: false }; });
      },
      world: "MAIN",
      allFrames: true
    });
    let res = null;
    if (results && results.length) {
      const main = results[0] && results[0].result;
      if (main && main.isListPage === true) {
        res = main;
      } else {
        for (const frame of results) {
          const r = frame && frame.result;
          if (r && r.isListPage === true) {
            res = r;
            break;
          }
        }
      }
    }
    if (res && res.isListPage) {
      listOrLibrary = res.isLibrary ? "Library" : "List";
    }
  } catch (_) {}
  optionEl.textContent = "Export " + listOrLibrary + " to " + formatLabel;
}

btnExport.addEventListener("click", () => {
  if (reportSelect && (reportSelect.value === "exportCSV" || reportSelect.value === "folderCount" || reportSelect.value === "pathLengths" || reportSelect.value === "permissionsMatrix")) runExport();
});

btnChooseColumns.addEventListener("click", openColumnPicker);

btnSettings.addEventListener("click", () => {
  mainPanel.classList.add("hidden");
  pickerPanel.classList.remove("visible");
  settingsPanel.classList.add("visible");
  document.body.classList.add("reports-settings-visible");
  loadSettings();
});

btnSettingsBack.addEventListener("click", () => {
  settingsPanel.classList.remove("visible");
  mainPanel.classList.remove("hidden");
  document.body.classList.remove("reports-settings-visible");
  savePageSize();
  saveDefaultExportFormat();
  updateExportReportLabel();
});

pageSizeSelect.addEventListener("change", savePageSize);
if (defaultExportFormatSelect) defaultExportFormatSelect.addEventListener("change", saveDefaultExportFormat);

btnBack.addEventListener("click", closeColumnPicker);

btnSelectAll.addEventListener("click", () => {
  columnList.querySelectorAll('input[type="checkbox"]').forEach((cb) => (cb.checked = true));
});

btnSelectNone.addEventListener("click", () => {
  columnList.querySelectorAll('input[type="checkbox"]').forEach((cb) => (cb.checked = false));
});

btnSelectDefaults.addEventListener("click", () => {
  if (!pickerData?.defaultColumns) return;
  const defaultSet = new Set(pickerData.defaultColumns);
  columnList.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
    cb.checked = defaultSet.has(cb.dataset.internal);
  });
});

btnExportSelected.addEventListener("click", () => {
  const cols = getSelectedColumns();
  if (cols.length === 0) {
    setPickerStatus("Select at least one column.", "error");
    return;
  }
  runExport(cols);
});

// Show version at bottom of extension settings
const settingsVersionEl = document.getElementById("settingsVersionPopup");
if (settingsVersionEl) settingsVersionEl.textContent = "v" + (chrome.runtime.getManifest().version || "");

// Notify content script when extension popup opens so Site Contents launcher can roll down
chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
  if (tabs[0]?.id) chrome.tabs.sendMessage(tabs[0].id, { type: "SPOToolkitPopupOpened" }).catch(() => {});
});
window.addEventListener("blur", () => {
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    if (tabs[0]?.id) chrome.tabs.sendMessage(tabs[0].id, { type: "SPOToolkitPopupClosed" }).catch(() => {});
  });
});
// When Site Contents panel is opened on the page, fade out then close the popup so both aren't open at once
chrome.runtime.onMessage.addListener((message) => {
  if (message.type !== "SPOToolkitPanelOpened") return;
  document.body.classList.add("popup-fade-out");
  setTimeout(() => window.close(), 220);
});
