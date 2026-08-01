/* global chrome */
/**
 * In-page compass panel: same tab surfaces as the extension popup (Quick Links, Page Props,
 * Columns, Reports, Refinables, Views, Site Contents). Loaded before content.js; exposes
 * SPOT_initCompassToolkitPanels(panel, invokeToolkitAction, ctx).
 */
(function () {
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

  const DEFAULT_PAGE_SIZE = 5000;
  const DEFAULT_EXPORT_FORMAT = "xlsx";
  const DEFAULT_ANIMATION_SECONDS = 0.22;
  const MIN_ANIMATION_SECONDS = 0.06;
  const MAX_ANIMATION_SECONDS = 0.9;
  const LAUNCHER_ICON_SCALE_KEY = "listsLauncherIconScale";
  const DEFAULT_LAUNCHER_ICON_SCALE = 1.28;
  const MIN_LAUNCHER_ICON_SCALE = 0.75;
  const MAX_LAUNCHER_ICON_SCALE = 1.6;
  const ANIMATION_SECONDS_KEY = "listsLauncherAnimationSeconds";
  const COMPASS_SHORTCUTS_ENABLED_KEY = "compassShortcutMacrosEnabled";
  const COMPASS_SHORTCUTS_KEY = "compassShortcutMacros";
  const COMPASS_ACTIVE_TAB_KEY = "popoutActiveTab";
  const COMPASS_REMEMBER_TAB_KEY = "compassRememberActiveTab";
  const COMPASS_PANE_IDS = ["quicklinks", "context", "columns", "reports", "refinableProps", "viewManager", "siteContents"];
  const COMPASS_DEFAULT_PANE = "siteContents";
  const DEFAULT_COMPASS_SHORTCUTS = {
    openPanel: "alt+shift+s",
    openSearch: "alt+shift+f",
    quicklinks: "alt+shift+1",
    context: "alt+shift+2",
    columns: "alt+shift+3",
    reports: "alt+shift+4",
    refinableProps: "alt+shift+5",
    viewManager: "alt+shift+6",
    siteContents: "alt+shift+7",
    toolkitSettings: "alt+shift+8"
  };
  const REQUIRED_FIELDS = ["ID", "Created", "Modified", "Author", "Editor", "_UIVersionString"];

  function escapeHtml(s) {
    const d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function copyTextToClipboard(text) {
    const s = String(text == null ? "" : text);
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(s).catch(function () {
        fallbackCopyText(s);
      });
    }
    fallbackCopyText(s);
    return Promise.resolve();
  }

  function fallbackCopyText(str) {
    try {
      const ta = document.createElement("textarea");
      ta.value = str;
      ta.setAttribute("readonly", "");
      ta.style.position = "fixed";
      ta.style.left = "-9999px";
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
    } catch (_) {}
  }

  const PAGE_PROP_VALUE_COPY_HINT = "Copy value";

  function flashPagePropValueCopied(el) {
    if (!el) return;
    el.setAttribute("title", "Copied!");
    el.classList.add("compass-page-props-value-copied");
    setTimeout(function () {
      el.setAttribute("title", PAGE_PROP_VALUE_COPY_HINT);
      el.classList.remove("compass-page-props-value-copied");
    }, 1400);
  }

  function listColumnSettingsUrl(siteUrl, listId, internalName) {
    const base = String(siteUrl || "").replace(/\/$/, "");
    const id = String(listId || "").replace(/[{}]/g, "");
    if (!base || !id || !internalName) return "";
    return (
      base +
      "/_layouts/15/FldEdit.aspx?List=" +
      encodeURIComponent("{" + id + "}") +
      "&Field=" +
      encodeURIComponent(internalName)
    );
  }

  function searchSchemaListMetaFromResponse(response) {
    return {
      siteUrl: String(response && response.siteUrl ? response.siteUrl : "").replace(/\/$/, ""),
      listId: String(response && response.listId ? response.listId : "").replace(/[{}]/g, "")
    };
  }

  function searchSchemaColumnMatchesFilter(column, qLower) {
    if (!qLower) return true;
    const parts = [column.internalName, column.title, column.type, column.group, column.crawledProperty];
    return parts.some(function (p) {
      return p && String(p).toLowerCase().indexOf(qLower) >= 0;
    });
  }

  function normalizeTrailingSlash(url) {
    if (!url || typeof url !== "string") return "";
    return url.replace(/\/$/, "");
  }

  const ADMIN_RECYCLE_BIN_SECOND_STAGE_SUFFIX = "/_layouts/15/AdminRecycleBin.aspx?view=5#view=13";

  function secondStageRecycleBinUrl(siteCollectionRootAbsoluteUrl) {
    const base = normalizeTrailingSlash(siteCollectionRootAbsoluteUrl);
    if (!base) return "";
    return base + ADMIN_RECYCLE_BIN_SECOND_STAGE_SUFFIX;
  }

  function parseContextFromUrl(pageUrl) {
    if (!pageUrl) return { webAbsoluteUrl: "", pageListId: "", viewId: "" };
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
        if (["sites", "site", "teams"].indexOf(segments[i]) >= 0 && segments[i + 1]) {
          siteEnd = i + 1;
          break;
        }
      }
      webAbsoluteUrl = siteEnd >= 0 ? u.origin + "/" + segments.slice(0, siteEnd + 1).join("/") : u.origin;
      return { webAbsoluteUrl: webAbsoluteUrl, pageListId: pageListId, viewId: viewId };
    } catch (_) {
      return { webAbsoluteUrl: "", pageListId: "", viewId: "" };
    }
  }

  function normalizeGuidString(s) {
    const t = String(s || "")
      .replace(/[{}]/g, "")
      .replace(/%7B|%7D/gi, "")
      .trim();
    return t ? t.toUpperCase() : "";
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

  const pagePropsCacheByUrl = Object.create(null);

  function pagePropsCacheKey() {
    return location.href;
  }

  function clearPagePropsCacheForUrl(url) {
    if (url) delete pagePropsCacheByUrl[url];
  }

  function invalidatePanelPageProps(panel) {
    try {
      delete panel.dataset.compassContextLoaded;
    } catch (_) {}
    panel._compassContextFlat = null;
  }

  function applyPagePropsCacheToPanel(panel, flat) {
    panel._compassContextFlat = flat;
    panel.dataset.compassContextLoaded = "1";
    renderPagePropsList(panel);
  }

  var COMPASS_TAB_TITLE_SUFFIX = {
    quicklinks: "Quick Links",
    context: "Page Props",
    columns: "Columns",
    reports: "Reports",
    refinableProps: "Refinables",
    viewManager: "Views",
    toolkitSettings: "Settings",
    siteContents: "Contents"
  };
  var LIST_ONLY_TAB_HINT = "Open a list or library view to use this tab.";

  function updateCompassPanelTitle(panel, paneId) {
    const titleEl = panel.querySelector(".sp-toolkit-panel-title-label");
    if (!titleEl) return;
    let id = paneId;
    if (!id) {
      const activeBtn = panel.querySelector(".sp-toolkit-compass-nav-tabs .sp-toolkit-compass-pane-tab.active");
      id = activeBtn ? activeBtn.getAttribute("data-compass-pane") : "siteContents";
    }
    if (panel.dataset.compassTitlePane === id) return;
    panel.dataset.compassTitlePane = id;
    const suffix = COMPASS_TAB_TITLE_SUFFIX[id] || id || "";
    const prefix = panel._compassTitlePrefixHtml;
    if (typeof prefix === "string" && prefix.length && suffix) {
      titleEl.innerHTML =
        prefix +
        "<span class=\"sp-toolkit-compass-title-sep\" aria-hidden=\"true\"> | </span>" +
        "<span class=\"sp-toolkit-compass-title-suffix\">" +
        escapeHtml(suffix) +
        "</span>";
    } else if (typeof prefix === "string" && prefix.length) {
      titleEl.innerHTML = prefix;
    } else {
      titleEl.innerHTML =
        "<span class=\"sp-toolkit-compass-title-suffix\">" + escapeHtml(suffix || "Site Contents") + "</span>";
    }
  }

  function resolveRememberedCompassPane(storedId, onListOrLibraryView) {
    let id = storedId && COMPASS_PANE_IDS.includes(storedId) ? storedId : COMPASS_DEFAULT_PANE;
    if ((id === "columns" || id === "viewManager") && !onListOrLibraryView) {
      id = COMPASS_DEFAULT_PANE;
    }
    return id;
  }

  function persistCompassActiveTab(paneId) {
    if (!paneId || paneId === "toolkitSettings") return;
    chrome.storage.local.get(COMPASS_REMEMBER_TAB_KEY, function (st) {
      if (st[COMPASS_REMEMBER_TAB_KEY] === false) return;
      chrome.storage.local.set({ [COMPASS_ACTIVE_TAB_KEY]: paneId });
    });
  }

  function pagePropsPaneHasRenderedRows(panel) {
    const host = panel.querySelector(".sp-toolkit-compass-context-host");
    return !!(host && host.querySelector(".compass-page-props-row"));
  }

  function columnsPaneHasRenderedRows(panel) {
    const listEl = panel.querySelector("#searchSchemaList");
    return !!(listEl && listEl.querySelector(".search-schema-row"));
  }

  function isCompassPaneContentReady(panel, paneId, onListOrLibraryView) {
    if (!panel || !paneId) return false;
    if (paneId === "quicklinks") {
      const content = panel.querySelector("#quicklinksContent");
      return !!(content && content.querySelector(".sp-toolkit-ql-section"));
    }
    if (paneId === "context") {
      return pagePropsPaneHasRenderedRows(panel);
    }
    if (paneId === "columns") {
      if (!columnsPaneHasRenderedRows(panel)) return false;
      const ctx = parseContextFromUrl(typeof location !== "undefined" ? location.href : "");
      const curList = normalizeGuidString(ctx.pageListId);
      const cacheList = normalizeGuidString(columnsCache.meta && columnsCache.meta.listId);
      if (curList && cacheList && curList !== cacheList) return false;
      return !!(columnsCache.columns && columnsCache.columns.length > 0);
    }
    if (paneId === "reports") {
      const host = panel.querySelector(".sp-toolkit-compass-reports-host");
      return !!(host && host.dataset.built === "10");
    }
    if (paneId === "refinableProps") {
      const out = panel.querySelector(".sp-toolkit-compass-refinable-out");
      const typeSel = panel.querySelector(".sp-toolkit-compass-refinable-type");
      const numSel = panel.querySelector(".sp-toolkit-compass-refinable-num");
      if (!out || !typeSel || !numSel || !typeSel.options.length) return false;
      const propertyName = (typeSel.value || "RefinableString") + (numSel.value || "00");
      if (out.dataset.refinableLoadedKey !== propertyName) return false;
      return !!(
        out.querySelector(".refinable-mappings-list") ||
        out.querySelector(".refinable-mappings-header") ||
        (out.textContent && out.textContent.indexOf("No crawled properties mapped") >= 0)
      );
    }
    if (paneId === "viewManager") {
      const host = panel.querySelector(".sp-toolkit-compass-views-host");
      return !!(host && host.dataset.wired === "1");
    }
    if (paneId === "siteContents") {
      return !!panel.querySelector(".sp-toolkit-contents-table.sp-toolkit-lists-loaded");
    }
    if (paneId === "toolkitSettings") {
      const host = panel.querySelector(".sp-toolkit-compass-settings-host");
      return !!(host && host.dataset.wired === "1");
    }
    return false;
  }

  function wireCompassPaneRegistry(panel) {
    if (!panel) return;
    const panes = {};
    const tabs = {};
    panel.querySelectorAll(".sp-toolkit-compass-pane[data-compass-pane]").forEach(function (p) {
      panes[p.getAttribute("data-compass-pane")] = p;
    });
    const bar = panel.querySelector(".sp-toolkit-compass-nav-tabs");
    if (bar) {
      bar.querySelectorAll(".sp-toolkit-compass-pane-tab[data-compass-pane]").forEach(function (t) {
        tabs[t.getAttribute("data-compass-pane")] = t;
      });
    }
    let activeId = panel.dataset.compassActivePane || "";
    if (!activeId || !panes[activeId]) {
      const activePane = panel.querySelector(".sp-toolkit-compass-pane.sp-toolkit-compass-pane-active, .sp-toolkit-compass-pane.active");
      activeId = activePane ? activePane.getAttribute("data-compass-pane") || "siteContents" : "siteContents";
    }
    panel._compassPaneRegistry = { panes: panes, tabs: tabs, activeId: activeId };
    panel.dataset.compassActivePane = activeId;
  }

  function refreshCompassTabRegistry(panel) {
    if (!panel) return;
    if (!panel._compassPaneRegistry) {
      wireCompassPaneRegistry(panel);
      return;
    }
    const tabs = {};
    const bar = panel.querySelector(".sp-toolkit-compass-nav-tabs");
    if (bar) {
      bar.querySelectorAll(".sp-toolkit-compass-pane-tab[data-compass-pane]").forEach(function (t) {
        tabs[t.getAttribute("data-compass-pane")] = t;
      });
    }
    panel._compassPaneRegistry.tabs = tabs;
  }

  function isCompassPaneLoadInFlight(panel, paneId) {
    if (!panel || !paneId) return false;
    if (paneId === "quicklinks") return panel.dataset.compassQuicklinksLoading === "1";
    if (paneId === "context") return panel.dataset.compassContextLoading === "1";
    if (paneId === "columns") return panel.dataset.compassColumnsLoading === "1";
    if (paneId === "refinableProps") return panel.dataset.compassRefinableLoading === "1";
    if (paneId === "reports") {
      const host = panel.querySelector(".sp-toolkit-compass-reports-host");
      return !!(host && host.dataset.building === "1");
    }
    return false;
  }

  function switchCompassPane(panel, paneId, invoke, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId, options) {
    options = options || {};
    if (!panel || !paneId) return;
    showCompassPane(panel, paneId);
    const ready = isCompassPaneContentReady(panel, paneId, onListOrLibraryView);
    const inFlight = isCompassPaneLoadInFlight(panel, paneId);
    if (ready) {
      if (panel.dataset.compassTitlePane !== paneId) {
        queueMicrotask(function () {
          updateCompassPanelTitle(panel, paneId);
        });
      }
    } else if (!inFlight) {
      requestAnimationFrame(function () {
        updateCompassPanelTitle(panel, paneId);
        loadCompassPaneContent(panel, paneId, invoke, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId);
      });
    } else if (panel.dataset.compassTitlePane !== paneId) {
      queueMicrotask(function () {
        updateCompassPanelTitle(panel, paneId);
      });
    }
    if (!options.skipPersist) {
      queueMicrotask(function () {
        persistCompassActiveTab(paneId);
      });
    }
  }

  function applyReportsPaneLayout(panel) {
    if (!panel) return;
    const host = panel.querySelector(".sp-toolkit-compass-reports-host");
    const reportSelect = host && host.querySelector(".sp-toolkit-compass-report-select");
    if (!host || !reportSelect) return;
    const reportType = reportSelect.value || "exportCSV";
    const showReportLists = reportType === "exportCSV" || reportType === "folderCount" || reportType === "pathLengths";
    host.classList.toggle("sp-toolkit-reports-lists-mode", showReportLists);
    host.classList.toggle("sp-toolkit-reports-matrix-mode", reportType === "permissionsMatrix");
    const reportsScroll = host.closest(".sp-toolkit-compass-scroll");
    if (reportsScroll) {
      reportsScroll.classList.toggle("sp-toolkit-reports-lists-scroll", showReportLists);
      reportsScroll.classList.toggle("sp-toolkit-reports-matrix-scroll", reportType === "permissionsMatrix");
    }
  }

  function showCompassPane(panel, paneId) {
    if (!panel || !paneId) return;
    if (!panel._compassPaneRegistry) {
      wireCompassPaneRegistry(panel);
    }
    const reg = panel._compassPaneRegistry;
    if (!reg || reg.activeId === paneId) {
      return;
    }
    const prevPane = reg.panes[reg.activeId];
    const nextPane = reg.panes[paneId];
    const prevTab = reg.tabs[reg.activeId];
    const nextTab = reg.tabs[paneId];
    if (prevPane) {
      prevPane.classList.remove("sp-toolkit-compass-pane-active", "active");
    }
    if (nextPane) {
      nextPane.classList.add("sp-toolkit-compass-pane-active", "active");
    }
    if (prevTab) {
      prevTab.classList.remove("active");
    }
    if (nextTab) {
      nextTab.classList.add("active");
    }
    reg.activeId = paneId;
    panel.dataset.compassActivePane = paneId;
    if (paneId === "reports") {
      applyReportsPaneLayout(panel);
    }
    scheduleTabIndicatorUpdate(panel);
  }

  function prefetchCompassPaneData(panel, invoke, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId) {
    if (!panel || panel.dataset.compassPrefetchStarted === "1") return;
    panel.dataset.compassPrefetchStarted = "1";
    void fillQuickLinksPane(panel, invoke, siteUrl, onSiteContentsPage);
    void loadPagePropsPane(panel, invoke);
    if (onListOrLibraryView) {
      void loadColumnsPane(panel, invoke);
    }
    requestAnimationFrame(function () {
      fillRefinableControls(panel);
      void loadRefinableMappings(panel, invoke, siteUrl);
      buildReportsPane(panel, invoke, siteUrl, currentListId);
    });
  }

  function runCompassPaneContentLoad(panel, paneId, invoke, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId) {
    if (paneId === "quicklinks") void fillQuickLinksPane(panel, invoke, siteUrl, onSiteContentsPage);
    else if (paneId === "context") void loadPagePropsPane(panel, invoke);
    else if (paneId === "columns") void loadColumnsPane(panel, invoke);
    else if (paneId === "reports") buildReportsPane(panel, invoke, siteUrl, currentListId);
    else if (paneId === "refinableProps") {
      fillRefinableControls(panel);
      void loadRefinableMappings(panel, invoke, siteUrl);
    } else if (paneId === "viewManager") wireViewsPane(panel, onListOrLibraryView, siteUrl, currentListId);
    else if (paneId === "toolkitSettings") wireToolkitSettingsPane(panel);
  }

  function loadCompassPaneContent(panel, paneId, invoke, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId) {
    runCompassPaneContentLoad(panel, paneId, invoke, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId);
  }

  function applyListOnlyCompassTabs(panel, onListOrLibraryView) {
    const nav = panel.querySelector(".sp-toolkit-compass-nav-tabs");
    if (!nav) return;
    ["columns", "viewManager"].forEach(function (pid) {
      const btn = nav.querySelector(".sp-toolkit-compass-pane-tab[data-compass-pane=\"" + pid + "\"]");
      if (!btn) return;
      btn.disabled = !onListOrLibraryView;
      btn.title = onListOrLibraryView ? "" : LIST_ONLY_TAB_HINT;
    });
    const reportsBtn = nav.querySelector(".sp-toolkit-compass-pane-tab[data-compass-pane=\"reports\"]");
    if (reportsBtn) {
      reportsBtn.disabled = false;
      reportsBtn.title = "";
    }
  }

  function activatePane(panel, paneId, options) {
    options = options || {};
    showCompassPane(panel, paneId);
    updateCompassPanelTitle(panel, paneId);
    if (!options.skipPersist) {
      queueMicrotask(function () {
        persistCompassActiveTab(paneId);
      });
    }
    if (!options.skipResize) {
      noteCompassPanelContentChanged(panel);
    }
  }

  function focusCompassPanePrimaryField(panel, paneId) {
    if (!panel) return;
    const overlay = panel.querySelector("#compassUniversalSearchOverlay");
    if (overlay && overlay.classList.contains("open")) return;

    const selectorsByPane = {
      quicklinks: "#quicklinksFilter",
      context: "#contextFilterInput",
      columns: "#searchSchemaFilter",
      siteContents: ".sp-toolkit-site-contents-name-filter input"
    };
    const selector = selectorsByPane[paneId];
    if (!selector) return;

    let attempts = 0;
    const maxAttempts = 3;
    function tryFocus() {
      if (!panel.isConnected) return;
      const el = panel.querySelector(selector);
      if (el && typeof el.focus === "function" && !el.disabled && el.offsetParent !== null) {
        el.focus({ preventScroll: true });
        if (typeof el.select === "function" && (paneId === "quicklinks" || paneId === "context" || paneId === "columns")) {
          try {
            el.select();
          } catch (_) {}
        }
        return;
      }
      attempts += 1;
      if (attempts < maxAttempts) requestAnimationFrame(tryFocus);
    }
    requestAnimationFrame(function () {
      requestAnimationFrame(tryFocus);
    });
  }

  function triggerPaneContentTransition(panel, paneId) {
    const pane = panel.querySelector(".sp-toolkit-compass-pane[data-compass-pane=\"" + paneId + "\"]");
    if (!pane) return;
    const scroll = pane.querySelector(".sp-toolkit-compass-scroll") || pane;
    if (!scroll) return;
    scroll.classList.remove("sp-toolkit-pane-content-enter");
    void scroll.offsetWidth;
    scroll.classList.add("sp-toolkit-pane-content-enter");
  }

  /** Run after the browser has painted tab activation (matches open-panel UX; avoids jank from heavy loaders). */
  function scheduleAfterCompassTabPaint(fn) {
    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(fn);
    } else {
      queueMicrotask(fn);
    }
  }

  function applyFlowListAnimation(container, selector, force) {
    return;
  }

  function ensureCompassTabIndicator(panel) {
    const bar = panel && panel.querySelector(".sp-toolkit-compass-nav-tabs");
    if (!bar) return null;
    let indicator = bar.querySelector(".sp-toolkit-tab-indicator");
    if (!indicator) {
      indicator = document.createElement("div");
      indicator.className = "sp-toolkit-tab-indicator";
      indicator.setAttribute("aria-hidden", "true");
      bar.appendChild(indicator);
    }
    if (!bar.dataset.spotIndicatorScrollBound) {
      bar.dataset.spotIndicatorScrollBound = "1";
      bar.addEventListener(
        "scroll",
        function () {
          updateTabIndicator(panel, true);
        },
        { passive: true }
      );
    }
    return indicator;
  }

  function scheduleTabIndicatorUpdate(panel, instant) {
    if (!panel) return;
    ensureCompassTabIndicator(panel);
    if (instant) {
      updateTabIndicator(panel, true);
      return;
    }
    requestAnimationFrame(function () {
      updateTabIndicator(panel, false);
      requestAnimationFrame(function () {
        updateTabIndicator(panel, false);
      });
    });
  }

  function updateTabIndicator(panel, instant) {
    const bar = panel && panel.querySelector(".sp-toolkit-compass-nav-tabs");
    if (!bar) return;
    const indicator = ensureCompassTabIndicator(panel);
    if (!indicator) return;
    const activeId =
      panel.dataset.compassActivePane ||
      (panel._compassPaneRegistry && panel._compassPaneRegistry.activeId) ||
      "";
    let active = activeId
      ? bar.querySelector('.sp-toolkit-compass-pane-tab[data-compass-pane="' + activeId + '"]')
      : null;
    if (!active) active = bar.querySelector(".sp-toolkit-compass-pane-tab.active");
    if (!active || active.disabled || active.offsetParent === null) {
      indicator.style.opacity = "0";
      return;
    }
    const barRect = bar.getBoundingClientRect();
    const tabRect = active.getBoundingClientRect();
    const left = tabRect.left - barRect.left;
    const width = tabRect.width;
    if (instant) indicator.classList.add("sp-toolkit-tab-indicator-instant");
    indicator.style.opacity = width > 0 ? "1" : "0";
    indicator.style.left = Math.max(0, Math.round(left)) + "px";
    indicator.style.width = Math.max(24, Math.round(width)) + "px";
    if (instant) {
      void indicator.offsetWidth;
      indicator.classList.remove("sp-toolkit-tab-indicator-instant");
    }
  }

  function getCompassPanelSizeCaps() {
    return {
      width: Math.min(Math.floor(window.innerWidth * 0.92), 1000),
      height: Math.floor(window.innerHeight * 0.7)
    };
  }

  function applyCompassPanelSize(panel) {
    const root = panel && (panel.closest("#sp-toolkit-lists-panel") || panel);
    if (!root || !root.isConnected) return;
    const caps = getCompassPanelSizeCaps();
    root.style.width = Math.max(360, caps.width) + "px";
    root.style.height = Math.max(280, caps.height) + "px";
  }

  function resetCompassPanelSize(panel) {
    const root = panel && (panel.closest("#sp-toolkit-lists-panel") || panel);
    if (!root) return;
    root.style.width = "";
    root.style.height = "";
  }

  function noteCompassPanelContentChanged(panel) {
    if (!panel) return;
    const paneId = panel.dataset.compassActivePane || "";
    if (paneId === "siteContents") {
      applyCompassPanelSize(panel);
    }
  }

  function addQuickLink(ul, label, href, icon) {
    const li = document.createElement("li");
    const a = document.createElement("a");
    a.href = href;
    a.target = "_blank";
    a.rel = "noopener noreferrer";
    a.textContent = label;
    if (icon) a.setAttribute("data-ql-icon", icon);
    li.appendChild(a);
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

  function addQuickLinkOrTenantAdminWait(ul, label, href, icon, adminHost) {
    if (!isQuickLinkTenantAdminHost(href, adminHost)) {
      addQuickLink(ul, label, href, icon);
      return;
    }
    const li = document.createElement("li");
    const a = document.createElement("a");
    a.href = href;
    a.target = "_blank";
    a.rel = "noopener noreferrer";
    a.textContent = label;
    if (icon) a.setAttribute("data-ql-icon", icon);
    a.addEventListener("click", function (e) {
      e.preventDefault();
      try {
        chrome.runtime.sendMessage({ type: "SPOToolkitOpenTenantAdminWithWait", url: href }, function () {
          void chrome.runtime.lastError;
        });
      } catch (_) {}
    });
    li.appendChild(a);
    ul.appendChild(li);
  }

  /** Same behavior as extension popup Quick Links: background resolves URLs like popup.js (no drift). */
  function addSecondStageRecycleBinQuickLink(ul, label, siteBase, secondStageHref, icon) {
    const li = document.createElement("li");
    const a = document.createElement("a");
    a.href = secondStageHref || "#";
    a.target = "_blank";
    a.rel = "noopener";
    a.textContent = label;
    if (icon) a.setAttribute("data-ql-icon", icon);
    a.addEventListener("click", function (e) {
      e.preventDefault();
      try {
        var port = chrome.runtime.connect({ name: "SPOToolkitSecondStageRecycleBin" });
        port.postMessage({ type: "startLikePopup" });
      } catch (_) {}
    });
    li.appendChild(a);
    ul.appendChild(li);
  }

  function qlSection(host, label, listId) {
    const sec = document.createElement("div");
    sec.className = "sp-toolkit-ql-section ql-section";
    const lab = document.createElement("div");
    lab.className = "sp-toolkit-ql-section-label ql-section-label";
    lab.textContent = label;
    const ul = document.createElement("ul");
    ul.className = "sp-toolkit-ql-list";
    if (listId) ul.id = listId;
    sec.appendChild(lab);
    sec.appendChild(ul);
    host.appendChild(sec);
    return ul;
  }

  async function fillQuickLinksPane(panel, invoke, siteUrl, onSiteContentsPage) {
    const host = panel.querySelector(".sp-toolkit-compass-quicklinks-host");
    if (!host || !siteUrl) return;
    if (panel.dataset.compassQuicklinksReady === "1" && host.querySelector("#quicklinksContent")) return;
    if (panel.dataset.compassQuicklinksLoading === "1") return;
    panel.dataset.compassQuicklinksLoading = "1";
    try {
    host.innerHTML = "";
    const filterRow = document.createElement("div");
    filterRow.className = "sp-toolkit-ql-filter-row";
    const inp = document.createElement("input");
    inp.id = "quicklinksFilter";
    inp.type = "text";
    inp.className = "sp-toolkit-ql-filter";
    inp.placeholder = "Filter quick links…";
    filterRow.appendChild(inp);
    host.appendChild(filterRow);
    const content = document.createElement("div");
    content.className = "sp-toolkit-ql-content";
    content.id = "quicklinksContent";
    host.appendChild(content);

    const tabUrl = typeof location !== "undefined" ? location.href : "";
    const hostPart = new URL(tabUrl).host;
    const adminHost = hostPart.replace(".sharepoint.com", "-admin.sharepoint.com");
    const ctx = parseContextFromUrl(tabUrl);
    const siteBase = (ctx.webAbsoluteUrl || "").replace(/\/$/, "") || siteUrl;

    let siteId = "";
    let siteCollRoot = "";
    try {
      const full = await invoke({ action: "getPageContextFull" });
      if (full && full.ok && full.data && typeof full.data === "object") {
        const d = full.data;
        siteId = normalizeGuidString(d.siteId || d.siteID || d.SiteId || d.SiteID || "");
        siteCollRoot = normalizeTrailingSlash(d.siteAbsoluteUrl || "");
      }
    } catch (_) {}
    if (!siteId) {
      try {
        const pj = await invoke({ action: "getPageContextJson" });
        if (pj && pj.ok && pj.data) siteId = findSiteIdDeep(pj.data, 0);
      } catch (_) {}
    }
    if (!siteCollRoot && siteBase) {
      const path = (function () {
        try {
          return new URL(siteBase).pathname.replace(/\/$/, "") || "/";
        } catch (_) {
          return "/";
        }
      })();
      try {
        const rest = await invoke({ action: "rest", method: "GET", path: path + "/_api/site/rootweb?$select=Url" });
        if (rest && rest.ok && rest.data) {
          const d = rest.data;
          siteCollRoot = normalizeTrailingSlash(
            (typeof d.Url === "string" && d.Url.trim()) || (typeof d.url === "string" && d.url.trim()) || ""
          );
        }
      } catch (_) {}
    }
    if (!siteCollRoot) siteCollRoot = siteBase;
    const secondStageHref = secondStageRecycleBinUrl(siteCollRoot);

    const currentSite = qlSection(content, "Current site", "quicklinksCurrentSite");
    addQuickLink(currentSite, "Site settings", siteBase + "/_layouts/15/settings.aspx", "settings");
    addQuickLink(currentSite, "Site content types", siteBase + "/_layouts/15/SiteAdmin.aspx#/contentTypes", "layers");
    addQuickLinkOrTenantAdminWait(
      currentSite,
      "Tenant content types",
      "https://" + adminHost + "/_layouts/15/online/AdminHome.aspx#/contentTypes",
      "adminHome",
      adminHost
    );
    addQuickLink(currentSite, "Recycle bin", siteBase + "/_layouts/15/RecycleBin.aspx", "trash");
    if (secondStageHref) addSecondStageRecycleBinQuickLink(currentSite, "Second Stage Recycle Bin", siteBase, secondStageHref, "trash");
    addQuickLink(currentSite, "All People", siteBase + "/_layouts/15/people.aspx?MembershipGroupId=0", "users");
    addQuickLink(currentSite, "Storage metrics", siteBase + "/_layouts/15/storman.aspx", "chart");
    addQuickLinkOrTenantAdminWait(
      currentSite,
      "SharePoint Admin Settings",
      siteId
        ? "https://" + adminHost + "/_layouts/15/online/AdminHome.aspx#/siteManagement/:/SiteDetails/" + siteId
        : "https://" + adminHost + "/_layouts/15/online/AdminHome.aspx#/siteManagement",
      "adminHome",
      adminHost
    );

    const currentUser = qlSection(content, "Current user", "quicklinksCurrentUser");
    addQuickLink(currentUser, "Edit user profile", siteBase + "/_layouts/15/me.aspx", "user");
    addQuickLink(currentUser, "Login as another user", siteBase + "/_layouts/15/closeConnection.aspx?loginasanotheruser=1", "logout");

    const modes = qlSection(content, "Page modes", "quicklinksModes");
    const pageUrl = tabUrl.split("?")[0];
    const q = function (param) {
      return pageUrl + (pageUrl.indexOf("?") >= 0 ? "&" : "?") + param;
    };
    addQuickLink(modes, "MaintenanceMode", q("MaintenanceMode=true"), "wrench");
    addQuickLink(modes, "WebView", q("env=WebView"), "monitor");
    addQuickLink(modes, "WebViewList", q("env=WebViewList"), "list");
    addQuickLink(modes, "Disable SPFx code", q("disable3PCode"), "shield");
    addQuickLink(modes, "Web Part Maintenance", q("contents=1"), "layers");

    const tenant = qlSection(content, "Tenant", "quicklinksTenant");
    addQuickLinkOrTenantAdminWait(tenant, "Admin center", "https://" + adminHost, "shield", adminHost);
    addQuickLinkOrTenantAdminWait(
      tenant,
      "Admin Center Settings",
      siteId
        ? "https://" + adminHost + "/_layouts/15/online/AdminHome.aspx#/siteManagement/:/SiteDetails/" + siteId + "/Settings"
        : "https://" + adminHost + "/_layouts/15/online/AdminHome.aspx#/siteManagement",
      "settings",
      adminHost
    );
    addQuickLinkOrTenantAdminWait(
      tenant,
      "Tenant site settings",
      "https://" + adminHost + "/_layouts/15/online/tenantsettings.aspx",
      "settings",
      adminHost
    );
    addQuickLinkOrTenantAdminWait(
      tenant,
      "User profiles",
      "https://" + adminHost + "/_layouts/15/TenantProfileAdmin/ManageUserProfileServiceApplication.aspx",
      "users",
      adminHost
    );
    addQuickLink(tenant, "Term store", "https://" + hostPart + "/_layouts/15/termstoremanager.aspx", "tag");
    addQuickLinkOrTenantAdminWait(
      tenant,
      "Search administration",
      "https://" + adminHost + "/_layouts/15/searchadmin/TA_SearchAdministration.aspx",
      "search",
      adminHost
    );
    addQuickLinkOrTenantAdminWait(
      tenant,
      "API access",
      "https://" + adminHost + "/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement",
      "key",
      adminHost
    );
    addQuickLink(tenant, "Teams admin", "https://admin.teams.microsoft.com/dashboard", "teams");
    addQuickLink(tenant, "App catalog", "https://" + hostPart + "/sites/AppCatalog/_layouts/15/appStore.aspx", "package");
    addQuickLink(tenant, "Classic app catalog", "https://" + hostPart + "/sites/AppCatalog", "archive");

    if (!onSiteContentsPage) {
      const classic = qlSection(content, "Classic");
      addQuickLink(classic, "Classic Site contents", siteBase + "/_layouts/15/viewlsts.aspx", "grid");
    }

    function applyFilter() {
      const qv = String(inp.value || "").toLowerCase().trim();
      content.querySelectorAll(".sp-toolkit-ql-section").forEach(function (section) {
        const labelText = (section.querySelector(".sp-toolkit-ql-section-label") || {}).textContent || "";
        const labelLower = labelText.toLowerCase();
        const sectionHit = qv && labelLower.indexOf(qv) >= 0;
        let any = false;
        section.querySelectorAll("li").forEach(function (li) {
          const a = li.querySelector("a");
          const text = (a && a.textContent) ? a.textContent.toLowerCase() : "";
          const href = (a && a.getAttribute("href")) ? a.getAttribute("href").toLowerCase() : "";
          const match = !qv || sectionHit || text.indexOf(qv) >= 0 || href.indexOf(qv) >= 0;
          li.style.display = match ? "" : "none";
          if (match) any = true;
        });
        section.style.display = !qv || any ? "" : "none";
      });
      noteCompassPanelContentChanged(panel);
    }
    inp.addEventListener("input", applyFilter);
    applyFlowListAnimation(content, ".sp-toolkit-ql-list li");
    panel.dataset.compassQuicklinksReady = "1";
    noteCompassPanelContentChanged(panel);
    } finally {
      delete panel.dataset.compassQuicklinksLoading;
    }
  }

  async function loadPagePropsPane(panel, invoke) {
    const cacheKey = pagePropsCacheKey();
    const cached = pagePropsCacheByUrl[cacheKey];
    if (cached && Object.keys(cached).length) {
      panel._compassContextFlat = cached;
      panel.dataset.compassContextLoaded = "1";
      if (pagePropsPaneHasRenderedRows(panel)) return;
      applyPagePropsCacheToPanel(panel, cached);
      return;
    }
    if (panel.dataset.compassContextLoaded === "1" && panel._compassContextFlat) {
      if (pagePropsPaneHasRenderedRows(panel)) return;
      renderPagePropsList(panel);
      return;
    }
    const host = panel.querySelector(".sp-toolkit-compass-context-host");
    if (!host) return;
    if (panel.dataset.compassContextLoading === "1") return;
    panel.dataset.compassContextLoading = "1";
    try {
    const countEl = panel.querySelector(".sp-toolkit-compass-context-count");
    if (countEl) countEl.textContent = "Properties: …";
    host.innerHTML = "<div class=\"sp-toolkit-panel-loading\">Loading…</div>";
    let flat = null;
    try {
      const full = await invoke({ action: "getPageContextFull" });
      if (full && full.ok && full.data && typeof full.data === "object" && Object.keys(full.data).length) {
        flat = flattenContext(full.data);
      }
    } catch (_) {}
    if (!flat || Object.keys(flat).length === 0) {
      try {
        const pj = await invoke({ action: "getPageContextJson" });
        if (pj && pj.ok && pj.data && typeof pj.data === "object" && Object.keys(pj.data).length) {
          flat = flattenContext(pj.data);
        } else {
          host.innerHTML =
            "<div class=\"sp-toolkit-panel-loading sp-toolkit-compass-err\">" +
            escapeHtml((pj && pj.error) || "Could not load page properties.") +
            "</div>";
          return;
        }
      } catch (_) {
        host.innerHTML =
          "<div class=\"sp-toolkit-panel-loading sp-toolkit-compass-err\">Could not load page properties.</div>";
        return;
      }
    }
    pagePropsCacheByUrl[cacheKey] = flat;
    applyPagePropsCacheToPanel(panel, flat);
    } finally {
      delete panel.dataset.compassContextLoading;
    }
  }

  function renderPagePropsList(panel) {
    const host = panel.querySelector(".sp-toolkit-compass-context-host");
    const flat = panel._compassContextFlat;
    if (!host || !flat) return;
    const filterInput = panel.querySelector(".sp-toolkit-compass-context-filter");
    const filter = filterInput ? String(filterInput.value || "").toLowerCase() : "";
    const entries = Object.entries(flat);
    const hadRows = pagePropsPaneHasRenderedRows(panel);
    const filtered = filter
      ? entries.filter(function (kv) {
          return kv[0].toLowerCase().indexOf(filter) >= 0 || String(kv[1]).toLowerCase().indexOf(filter) >= 0;
        })
      : entries;
    const countEl = panel.querySelector(".sp-toolkit-compass-context-count");
    if (countEl) countEl.textContent = "Properties: " + filtered.length;
    host.innerHTML = "";
    filtered.forEach(function (kv) {
      const name = kv[0];
      const value = kv[1];
      const valStr = value === null || value === undefined ? "null" : String(value);
      const row = document.createElement("div");
      row.className = "sp-toolkit-compass-colrow search-schema-row compass-page-props-row";
      const nameSpan = document.createElement("span");
      nameSpan.className = "compass-page-props-name";
      nameSpan.textContent = name;
      const valueWrap = document.createElement("span");
      valueWrap.className = "compass-page-props-value-wrap";
      const valueSpan = document.createElement("span");
      valueSpan.className = "compass-page-props-value";
      valueSpan.textContent = valStr;
      valueSpan.setAttribute("title", PAGE_PROP_VALUE_COPY_HINT);
      valueSpan.setAttribute("aria-label", PAGE_PROP_VALUE_COPY_HINT);
      const copyBtn = document.createElement("button");
      copyBtn.type = "button";
      copyBtn.className = "compass-page-props-copy-btn";
      copyBtn.setAttribute("title", PAGE_PROP_VALUE_COPY_HINT);
      copyBtn.setAttribute("aria-label", "Copy value");
      copyBtn.innerHTML =
        "<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\" focusable=\"false\"><path fill=\"currentColor\" d=\"M16 1H6a2 2 0 0 0-2 2v12h2V3h10V1zm3 4H10a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h9a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2zm0 16H10V7h9v14z\"/></svg>";
      copyBtn.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        copyTextToClipboard(valStr).then(function () {
          flashPagePropValueCopied(valueSpan);
          flashPagePropValueCopied(copyBtn);
        });
      });
      valueWrap.appendChild(valueSpan);
      valueWrap.appendChild(copyBtn);
      row.appendChild(nameSpan);
      row.appendChild(valueWrap);
      host.appendChild(row);
    });
    applyFlowListAnimation(host, ".compass-page-props-row");
    if (!hadRows) {
      noteCompassPanelContentChanged(panel);
    }
  }

  let columnsCache = { columns: [], meta: { siteUrl: "", listId: "" } };

  function ensureColumnCreatorHost(panel) {
    const pane = panel.querySelector('[data-compass-pane="columns"] #searchSchemaPanel');
    if (!pane) return null;
    let host = pane.querySelector(".sp-toolkit-compass-column-creator-host");
    if (!host) {
      host = document.createElement("div");
      host.className = "sp-toolkit-compass-column-creator-host";
      const toolbar = pane.querySelector(".toolbar");
      if (toolbar && toolbar.parentNode) {
        if (toolbar.nextSibling) toolbar.parentNode.insertBefore(host, toolbar.nextSibling);
        else toolbar.parentNode.appendChild(host);
      } else {
        pane.appendChild(host);
      }
    }
    return host;
  }

  function mountCompassColumnCreator(panel, invoke, hasList, attempt) {
    const host = ensureColumnCreatorHost(panel);
    if (!host || host.dataset.columnCreatorMounted === "1") return;
    const mount = typeof window.SPOT_mountColumnCreator === "function" ? window.SPOT_mountColumnCreator : null;
    if (!mount) {
      const tries = attempt || 0;
      if (tries < 40) {
        setTimeout(function () {
          mountCompassColumnCreator(panel, invoke, hasList, tries + 1);
        }, 50);
        return;
      }
      host.innerHTML =
        '<div class="column-creator column-creator-load-err">Column creator could not load. Reload the extension, refresh this page, then reopen Columns.</div>';
      return;
    }
    const toggleBtn = panel.querySelector("#btnOpenColumnCreator");
    mount(host, {
      invoke: invoke,
      hasList: !!hasList,
      toggleBtn: toggleBtn || undefined,
      onCreated: function () {
        loadColumnsPane(panel, invoke, true);
      },
    });
  }

  async function loadColumnsPane(panel, invoke, forceRefresh) {
    const countEl = panel.querySelector(".sp-toolkit-compass-columns-count");
    const listEl = panel.querySelector("#searchSchemaList");
    const errEl = panel.querySelector(".sp-toolkit-compass-columns-error");
    if (!listEl || !errEl) return;
    if (
      !forceRefresh &&
      columnsCache.columns &&
      columnsCache.columns.length > 0 &&
      columnsCache.meta &&
      columnsCache.meta.listId
    ) {
      const ctx = parseContextFromUrl(typeof location !== "undefined" ? location.href : "");
      const curList = normalizeGuidString(ctx.pageListId);
      const cacheList = normalizeGuidString(columnsCache.meta.listId);
      if (curList && cacheList && curList === cacheList) {
        errEl.style.display = "none";
        if (!columnsPaneHasRenderedRows(panel)) {
          renderColumnsList(panel);
        }
        mountCompassColumnCreator(panel, invoke, true);
        return;
      }
    }
    errEl.style.display = "none";
    errEl.textContent = "";
    if (panel.dataset.compassColumnsLoading === "1") return;
    panel.dataset.compassColumnsLoading = "1";
    try {
    listEl.innerHTML = "<div class=\"sp-toolkit-panel-loading\">Loading…</div>";
    if (countEl) countEl.textContent = "Columns: …";
    const response = await invoke({ action: "getSearchSchema" });
    if (!response || !response.ok) {
      columnsCache.columns = [];
      columnsCache.meta = { siteUrl: "", listId: "" };
      if (countEl) countEl.textContent = "Columns: 0";
      errEl.textContent = (response && response.error) || "Open a list or library view on this page to load columns.";
      errEl.style.display = "block";
      mountCompassColumnCreator(panel, invoke, false);
      return;
    }
    columnsCache.columns = response.columns || [];
    columnsCache.meta = searchSchemaListMetaFromResponse(response);
    if (countEl) countEl.textContent = "Columns: " + columnsCache.columns.length;
    renderColumnsList(panel);
    mountCompassColumnCreator(panel, invoke, !!(columnsCache.meta && columnsCache.meta.listId));
    noteCompassPanelContentChanged(panel);
    } finally {
      delete panel.dataset.compassColumnsLoading;
    }
  }

  function renderColumnsList(panel) {
    const countEl = panel.querySelector(".sp-toolkit-compass-columns-count");
    const listEl = panel.querySelector("#searchSchemaList");
    const filterInput = panel.querySelector(".sp-toolkit-compass-columns-filter");
    if (!listEl) return;
    const q = filterInput ? String(filterInput.value || "").toLowerCase().trim() : "";
    const filtered = q ? columnsCache.columns.filter(function (c) { return searchSchemaColumnMatchesFilter(c, q); }) : columnsCache.columns;
    if (countEl) countEl.textContent = "Columns: " + filtered.length;
    listEl.innerHTML = "";
    filtered.forEach(function (c) {
      const row = document.createElement("div");
      row.className = "sp-toolkit-compass-colrow search-schema-row";
      const displayTitle = c.title || c.internalName || "";
      const settingsHref = listColumnSettingsUrl(columnsCache.meta.siteUrl, columnsCache.meta.listId, c.internalName || "");
      const nameCell = settingsHref
        ? "<a class=\"col-name-link\" href=\"" + escapeHtml(settingsHref) + "\" target=\"_blank\" rel=\"noopener\" title=\"Open column settings\">" + escapeHtml(displayTitle) + "</a>"
        : escapeHtml(displayTitle);
      row.innerHTML =
        "<span class=\"sp-toolkit-cc-name col-name\">" +
        nameCell +
        "</span><span class=\"col-internal\">" +
        escapeHtml(c.internalName || "") +
        "</span><span class=\"col-crawled\">" +
        escapeHtml(c.crawledProperty || "") +
        "</span><span class=\"col-type\"><span class=\"col-type-pill\">" +
        escapeHtml(c.type || "") +
        "</span>" +
        "</span>";
      listEl.appendChild(row);
    });
    applyFlowListAnimation(listEl, ".sp-toolkit-compass-colrow");
    noteCompassPanelContentChanged(panel);
  }

  function fillRefinableControls(panel) {
    const typeSel = panel.querySelector(".sp-toolkit-compass-refinable-type");
    const numSel = panel.querySelector(".sp-toolkit-compass-refinable-num");
    if (!typeSel || !numSel || typeSel.options.length) return;
    REFINABLE_TYPES.forEach(function (t) {
      const opt = document.createElement("option");
      opt.value = t.id;
      opt.textContent = t.label;
      typeSel.appendChild(opt);
    });
    typeSel.value = "RefinableString";
    function fillNums() {
      const typeId = typeSel.value;
      let type = REFINABLE_TYPES[0];
      for (let ri = 0; ri < REFINABLE_TYPES.length; ri++) {
        if (REFINABLE_TYPES[ri].id === typeId) {
          type = REFINABLE_TYPES[ri];
          break;
        }
      }
      numSel.innerHTML = "";
      for (let i = 0; i < type.count; i++) {
        const opt = document.createElement("option");
        const num = String(i).padStart(2, "0");
        opt.value = num;
        opt.textContent = num;
        numSel.appendChild(opt);
      }
    }
    fillNums();
    typeSel.addEventListener("change", fillNums);
  }

  async function loadRefinableMappings(panel, invoke, siteUrl) {
    const out = panel.querySelector(".sp-toolkit-compass-refinable-out");
    const typeSel = panel.querySelector(".sp-toolkit-compass-refinable-type");
    const numSel = panel.querySelector(".sp-toolkit-compass-refinable-num");
    if (!out || !typeSel || !numSel || !siteUrl) return;
    const typeId = typeSel.value || "RefinableString";
    const num = numSel.value || "00";
    const propertyName = typeId + num;
    if (
      out.dataset.refinableLoadedKey === propertyName &&
      (out.querySelector(".refinable-mappings-list") ||
        out.querySelector(".refinable-mappings-header") ||
        (out.textContent && out.textContent.indexOf("No crawled properties mapped") >= 0))
    ) {
      return;
    }
    if (panel.dataset.compassRefinableLoading === "1") return;
    panel.dataset.compassRefinableLoading = "1";
    try {
    out.innerHTML = "<div class=\"sp-toolkit-panel-loading\">Loading…</div>";
    out.className = "sp-toolkit-compass-refinable-out";
    const res = await invoke({ action: "getRefinableMappings", siteUrl: siteUrl, propertyName: propertyName });
    if (!res) {
      out.textContent = "No response.";
      out.className = "sp-toolkit-compass-refinable-out sp-toolkit-compass-err";
      return;
    }
    if (!res.ok) {
      out.textContent = res.error || "Failed to load mappings.";
      out.className = "sp-toolkit-compass-refinable-out sp-toolkit-compass-err";
      delete out.dataset.refinableLoadedKey;
      return;
    }
    out.dataset.refinableLoadedKey = propertyName;
    if (!res.mappings || res.mappings.length === 0) {
      out.innerHTML =
        "<span class=\"refinable-mappings-header\">Crawled properties</span>" +
        (res.alias ? "<p class=\"refinable-alias\">Alias: " + escapeHtml(String(res.alias)) + "</p>" : "") +
        "<p>No crawled properties mapped for " + escapeHtml(propertyName) + ".</p>";
      out.className = "sp-toolkit-compass-refinable-out";
      noteCompassPanelContentChanged(panel);
      return;
    }
    const aliasPart = res.alias ? "<p class=\"refinable-alias\">Alias: " + escapeHtml(String(res.alias)) + "</p>" : "";
    const listItems = res.mappings
      .map(function (m) {
        return "<li>" + escapeHtml(m.name) + (m.type ? " <span class=\"sp-toolkit-ref-type\">" + escapeHtml(m.type) + "</span>" : "") + "</li>";
      })
      .join("");
    out.innerHTML = aliasPart + "<span class=\"refinable-mappings-header\">Crawled properties (" + res.mappings.length + ")</span><ul class=\"refinable-mappings-list\">" + listItems + "</ul>";
    out.className = "sp-toolkit-compass-refinable-out";
    applyFlowListAnimation(out, ".refinable-mappings-list li");
    noteCompassPanelContentChanged(panel);
    } finally {
      delete panel.dataset.compassRefinableLoading;
    }
  }

  function openRefinableAdmin(panel, siteUrl) {
    const typeSel = panel.querySelector(".sp-toolkit-compass-refinable-type");
    const numSel = panel.querySelector(".sp-toolkit-compass-refinable-num");
    if (!typeSel || !numSel || !siteUrl) return;
    const propertyName = (typeSel.value || "RefinableString") + (numSel.value || "00");
    const enc = encodeURIComponent(propertyName);
    window.open(siteUrl + "/_layouts/15/managedproperty.aspx?property=" + enc + "&level=sitecol", "_blank", "noopener,noreferrer");
  }

  function wireViewsPane(panel, onListOrLibraryView, siteUrl, currentListId) {
    const host = panel.querySelector(".sp-toolkit-compass-views-host");
    if (!host || host.dataset.wired === "1") return;
    host.dataset.wired = "1";
    host.innerHTML = "";
    const p = document.createElement("p");
    p.className = "sp-toolkit-compass-muted";
    p.textContent =
      "View Manager opens in a full tab. View formatter opens the JSON editor with list preview for the current page.";
    host.appendChild(p);
    const row = document.createElement("div");
    row.className = "sp-toolkit-compass-btnrow";
    const b1 = document.createElement("button");
    b1.type = "button";
    b1.className = "sp-toolkit-compass-primary";
    b1.textContent = "Open View Manager";
    const listIdClean = String(currentListId || "").replace(/[{}]/g, "").trim();
    const canOpenViewManager = !!(siteUrl && listIdClean);
    b1.disabled = !canOpenViewManager;
    b1.title = canOpenViewManager ? "Open View Manager" : "Open a list or library view first";
    b1.addEventListener("click", function () {
      if (!canOpenViewManager) return;
      try {
        chrome.runtime.sendMessage(
          { type: "SPOToolkitOpenViewManager", listId: listIdClean, webUrl: siteUrl },
          function (res) {
            if (chrome.runtime.lastError || (res && res.ok === false)) {
              console.warn("SPOToolkit: View Manager open failed", chrome.runtime.lastError || res);
            }
          }
        );
      } catch (e) {
        console.warn("SPOToolkit: View Manager open failed", e);
      }
    });
    row.appendChild(b1);
    const b2 = document.createElement("button");
    b2.type = "button";
    b2.className = "sp-toolkit-compass-secondary";
    b2.textContent = "View formatter";
    b2.disabled = !onListOrLibraryView;
    b2.title = onListOrLibraryView ? "Open View formatter" : "Open a list or library view first";
    b2.addEventListener("click", function () {
      if (!onListOrLibraryView) return;
      try {
        chrome.runtime.sendMessage({ type: "SPOToolkitOpenViewFormatter", previewUrl: window.location.href });
      } catch (e) {
        console.warn(e);
      }
    });
    row.appendChild(b2);
    host.appendChild(row);
    noteCompassPanelContentChanged(panel);
  }

  function getPageSizePromise() {
    return new Promise(function (resolve) {
      chrome.storage.local.get("pageSize", function (st) {
        const v = parseInt(st.pageSize, 10);
        resolve(v === 500 || v === 1000 || v === 5000 ? v : DEFAULT_PAGE_SIZE);
      });
    });
  }

  function getDefaultExportFormatPromise() {
    return new Promise(function (resolve) {
      chrome.storage.local.get("defaultExportFormat", function (st) {
        const f = st.defaultExportFormat;
        resolve(f === "csv" || f === "xlsx" ? f : DEFAULT_EXPORT_FORMAT);
      });
    });
  }

  function wireToolkitSettingsPane(panel) {
    const host = panel.querySelector(".sp-toolkit-compass-settings-host");
    if (!host || host.dataset.wired === "1") return;
    host.dataset.wired = "1";
    host.innerHTML =
      "<div class=\"sp-toolkit-compass-settings-shell\">" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitSettingLauncher\">Show launcher on SharePoint pages</label><input type=\"checkbox\" id=\"spToolkitSettingLauncher\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitSettingPopupSearch\">Open universal search when dropdown opens</label><input type=\"checkbox\" id=\"spToolkitSettingPopupSearch\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitSettingPopoutSearch\">Open universal search when popout opens</label><input type=\"checkbox\" id=\"spToolkitSettingPopoutSearch\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitSettingRememberTab\">Remember last tab when opening popout</label><input type=\"checkbox\" id=\"spToolkitSettingRememberTab\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitSettingShortcutsEnabled\">Enable keyboard shortcut macros</label><input type=\"checkbox\" id=\"spToolkitSettingShortcutsEnabled\"/></div>" +
      "<div id=\"spToolkitShortcutRows\" class=\"sp-toolkit-compass-settings-shortcuts\">" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutOpenPanel\">Open popout</label><input id=\"spToolkitShortcutOpenPanel\" type=\"text\" placeholder=\"alt+shift+s\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutOpenSearch\">Open popout + search</label><input id=\"spToolkitShortcutOpenSearch\" type=\"text\" placeholder=\"alt+shift+f\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutQuickLinks\">Open Quick Links tab</label><input id=\"spToolkitShortcutQuickLinks\" type=\"text\" placeholder=\"alt+shift+1\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutPageProps\">Open Page Props tab</label><input id=\"spToolkitShortcutPageProps\" type=\"text\" placeholder=\"alt+shift+2\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutColumns\">Open Columns tab</label><input id=\"spToolkitShortcutColumns\" type=\"text\" placeholder=\"alt+shift+3\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutReports\">Open Reports tab</label><input id=\"spToolkitShortcutReports\" type=\"text\" placeholder=\"alt+shift+4\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutRefinables\">Open Refinables tab</label><input id=\"spToolkitShortcutRefinables\" type=\"text\" placeholder=\"alt+shift+5\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutViews\">Open Views tab</label><input id=\"spToolkitShortcutViews\" type=\"text\" placeholder=\"alt+shift+6\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutSiteContents\">Open Site Contents tab</label><input id=\"spToolkitShortcutSiteContents\" type=\"text\" placeholder=\"alt+shift+7\"/></div>" +
      "<div class=\"sp-toolkit-compass-settings-row\"><label for=\"spToolkitShortcutSettings\">Open Settings tab</label><input id=\"spToolkitShortcutSettings\" type=\"text\" placeholder=\"alt+shift+8\"/></div>" +
      "</div>" +
      "<div class=\"sp-toolkit-compass-field\"><label class=\"sp-toolkit-compass-label\" for=\"spToolkitSettingCompassSize\">Launcher icon size</label>" +
      "<div class=\"sp-toolkit-compass-settings-range-wrap\"><input id=\"spToolkitSettingCompassSize\" class=\"sp-toolkit-compass-settings-range\" type=\"range\" min=\"75\" max=\"160\" step=\"1\"/><span id=\"spToolkitCompassSizeValue\" class=\"sp-toolkit-compass-settings-range-value\">100% (Medium)</span></div></div>" +
      "<div class=\"sp-toolkit-compass-field\"><label class=\"sp-toolkit-compass-label\" for=\"spToolkitSettingAnimationSpeed\">Animation speed</label>" +
      "<div class=\"sp-toolkit-compass-settings-range-wrap\"><input id=\"spToolkitSettingAnimationSpeed\" class=\"sp-toolkit-compass-settings-range\" type=\"range\" min=\"0\" max=\"100\" step=\"1\"/><span id=\"spToolkitAnimationValue\" class=\"sp-toolkit-compass-settings-range-value\">50%</span></div></div>" +
      "<div class=\"sp-toolkit-compass-btnrow\"><button type=\"button\" class=\"sp-toolkit-compass-secondary\" id=\"spToolkitSettingsBackBtn\">Back</button></div>" +
      "</div>";

    const launcherChk = host.querySelector("#spToolkitSettingLauncher");
    const popupSearchChk = host.querySelector("#spToolkitSettingPopupSearch");
    const popoutSearchChk = host.querySelector("#spToolkitSettingPopoutSearch");
    const rememberTabChk = host.querySelector("#spToolkitSettingRememberTab");
    const shortcutsEnabledChk = host.querySelector("#spToolkitSettingShortcutsEnabled");
    const shortcutRows = host.querySelector("#spToolkitShortcutRows");
    const shortcutInputs = {
      openPanel: host.querySelector("#spToolkitShortcutOpenPanel"),
      openSearch: host.querySelector("#spToolkitShortcutOpenSearch"),
      quicklinks: host.querySelector("#spToolkitShortcutQuickLinks"),
      context: host.querySelector("#spToolkitShortcutPageProps"),
      columns: host.querySelector("#spToolkitShortcutColumns"),
      reports: host.querySelector("#spToolkitShortcutReports"),
      refinableProps: host.querySelector("#spToolkitShortcutRefinables"),
      viewManager: host.querySelector("#spToolkitShortcutViews"),
      siteContents: host.querySelector("#spToolkitShortcutSiteContents"),
      toolkitSettings: host.querySelector("#spToolkitShortcutSettings")
    };
    const speedRange = host.querySelector("#spToolkitSettingAnimationSpeed");
    const speedValue = host.querySelector("#spToolkitAnimationValue");
    const sizeRange = host.querySelector("#spToolkitSettingCompassSize");
    const sizeValue = host.querySelector("#spToolkitCompassSizeValue");
    const backBtn = host.querySelector("#spToolkitSettingsBackBtn");

    function normalizeSeconds(v) {
      const n = Number(v);
      if (!Number.isFinite(n)) return DEFAULT_ANIMATION_SECONDS;
      return Math.min(MAX_ANIMATION_SECONDS, Math.max(MIN_ANIMATION_SECONDS, n));
    }

    function sliderToSeconds(v) {
      const pct = Math.min(100, Math.max(0, Number(v) || 0)) / 100;
      // Invert mapping so "up/right" means faster (smaller duration).
      return MAX_ANIMATION_SECONDS - pct * (MAX_ANIMATION_SECONDS - MIN_ANIMATION_SECONDS);
    }

    function secondsToSlider(seconds) {
      const s = normalizeSeconds(seconds);
      const pct = (MAX_ANIMATION_SECONDS - s) / (MAX_ANIMATION_SECONDS - MIN_ANIMATION_SECONDS);
      return Math.round(pct * 100);
    }

    function updateSpeedLabel(sliderValue, seconds) {
      if (!speedValue) return;
      const pct = Math.min(100, Math.max(0, Math.round(Number(sliderValue) || 0)));
      const sec = normalizeSeconds(seconds);
      speedValue.textContent = pct + "% (" + sec.toFixed(2) + "s)";
    }

    function updateSizeLabel(percent) {
      if (!sizeValue) return;
      const p = Math.min(160, Math.max(75, Math.round(Number(percent) || 100)));
      const tone = p < 95 ? "Small" : p > 115 ? "Large" : "Medium";
      sizeValue.textContent = p + "% (" + tone + ")";
    }

    function normalizeShortcutText(v) {
      return String(v || "")
        .toLowerCase()
        .replace(/\s+/g, "")
        .replace(/control/g, "ctrl")
        .replace(/command/g, "meta")
        .replace(/option/g, "alt");
    }

    function applyShortcutsEnabledUi(enabled) {
      if (shortcutRows) shortcutRows.style.opacity = enabled ? "1" : ".5";
      Object.keys(shortcutInputs).forEach(function (k) {
        if (shortcutInputs[k]) shortcutInputs[k].disabled = !enabled;
      });
    }

    chrome.storage.local.get(
      ["listsLauncherEnabled", "universalSearchOpenOnLoad", "compassUniversalSearchOpenOnLoad", COMPASS_REMEMBER_TAB_KEY, COMPASS_SHORTCUTS_ENABLED_KEY, COMPASS_SHORTCUTS_KEY, ANIMATION_SECONDS_KEY, LAUNCHER_ICON_SCALE_KEY],
      function (st) {
        if (launcherChk) launcherChk.checked = st.listsLauncherEnabled !== false;
        if (popupSearchChk) popupSearchChk.checked = !!st.universalSearchOpenOnLoad;
        if (popoutSearchChk) popoutSearchChk.checked = !!st.compassUniversalSearchOpenOnLoad;
        if (rememberTabChk) rememberTabChk.checked = st[COMPASS_REMEMBER_TAB_KEY] !== false;
        const shortcutsEnabled = !!st[COMPASS_SHORTCUTS_ENABLED_KEY];
        if (shortcutsEnabledChk) shortcutsEnabledChk.checked = shortcutsEnabled;
        applyShortcutsEnabledUi(shortcutsEnabled);
        const map = Object.assign({}, DEFAULT_COMPASS_SHORTCUTS, st[COMPASS_SHORTCUTS_KEY] || {});
        Object.keys(shortcutInputs).forEach(function (key) {
          if (shortcutInputs[key]) shortcutInputs[key].value = map[key] || "";
        });
        const speed = Number.isFinite(Number(st[ANIMATION_SECONDS_KEY])) ? Number(st[ANIMATION_SECONDS_KEY]) : DEFAULT_ANIMATION_SECONDS;
        const sliderPos = secondsToSlider(speed);
        if (speedRange) speedRange.value = String(sliderPos);
        updateSpeedLabel(sliderPos, speed);
        const sizeScaleRaw = Number(st[LAUNCHER_ICON_SCALE_KEY]);
        const sizeScale = Number.isFinite(sizeScaleRaw) ? sizeScaleRaw : DEFAULT_LAUNCHER_ICON_SCALE;
        const sizePercent = Math.round(sizeScale * 100);
        if (sizeRange) sizeRange.value = String(Math.min(160, Math.max(75, sizePercent)));
        updateSizeLabel(sizePercent);
      }
    );

    if (launcherChk) {
      launcherChk.addEventListener("change", function () {
        chrome.storage.local.set({ listsLauncherEnabled: launcherChk.checked });
      });
    }
    if (popupSearchChk) {
      popupSearchChk.addEventListener("change", function () {
        chrome.storage.local.set({ universalSearchOpenOnLoad: popupSearchChk.checked });
      });
    }
    if (popoutSearchChk) {
      popoutSearchChk.addEventListener("change", function () {
        chrome.storage.local.set({ compassUniversalSearchOpenOnLoad: popoutSearchChk.checked });
      });
    }
    if (rememberTabChk) {
      rememberTabChk.addEventListener("change", function () {
        chrome.storage.local.set({ [COMPASS_REMEMBER_TAB_KEY]: rememberTabChk.checked });
      });
    }
    if (shortcutsEnabledChk) {
      shortcutsEnabledChk.addEventListener("change", function () {
        const enabled = !!shortcutsEnabledChk.checked;
        applyShortcutsEnabledUi(enabled);
        chrome.storage.local.set({ [COMPASS_SHORTCUTS_ENABLED_KEY]: enabled });
      });
    }
    Object.keys(shortcutInputs).forEach(function (key) {
      const input = shortcutInputs[key];
      if (!input) return;
      input.addEventListener("change", function () {
        const normalized = normalizeShortcutText(input.value);
        input.value = normalized;
        chrome.storage.local.get(COMPASS_SHORTCUTS_KEY, function (st) {
          const next = Object.assign({}, DEFAULT_COMPASS_SHORTCUTS, st[COMPASS_SHORTCUTS_KEY] || {});
          next[key] = normalized;
          chrome.storage.local.set({ [COMPASS_SHORTCUTS_KEY]: next });
        });
      });
    });
    if (speedRange) {
      speedRange.addEventListener("input", function () {
        const sliderPos = Number(speedRange.value || 0);
        const nextSeconds = sliderToSeconds(sliderPos);
        updateSpeedLabel(sliderPos, nextSeconds);
        chrome.storage.local.set({ [ANIMATION_SECONDS_KEY]: nextSeconds });
      });
    }
    if (sizeRange) {
      sizeRange.addEventListener("input", function () {
        const sizePercent = Number(sizeRange.value || 100);
        const scale = Math.min(MAX_LAUNCHER_ICON_SCALE, Math.max(MIN_LAUNCHER_ICON_SCALE, sizePercent / 100));
        updateSizeLabel(sizePercent);
        chrome.storage.local.set({ [LAUNCHER_ICON_SCALE_KEY]: scale });
      });
    }
    if (backBtn) {
      backBtn.addEventListener("click", function () {
        activatePane(panel, panel.dataset.compassSettingsReturnPane || "siteContents");
      });
    }
  }

  function normalizeMatrixSiteKey(siteUrl) {
    if (!siteUrl) return "";
    try {
      var u = new URL(String(siteUrl).replace(/\/$/, ""));
      return (u.origin + u.pathname.replace(/\/$/, "")).toLowerCase();
    } catch (_) {
      return String(siteUrl).replace(/\/$/, "").toLowerCase();
    }
  }

  async function resolveReportContext(invoke, fallbackSiteUrl, fallbackListId) {
    let siteUrl = "";
    let listId = "";
    let viewId = "";
    try {
      const pc = await invoke({ action: "getPageContext" });
      if (pc && pc.ok) {
        siteUrl = String(pc.webAbsoluteUrl || pc.siteAbsoluteUrl || "").replace(/\/$/, "");
        listId = String(pc.pageListId || "").replace(/[{}]/g, "");
      }
    } catch (_) {}
    if (!siteUrl || !listId) {
      const ctx = parseContextFromUrl(typeof location !== "undefined" ? location.href : "");
      if (!siteUrl) siteUrl = String(ctx.webAbsoluteUrl || "").replace(/\/$/, "");
      if (!listId) listId = String(ctx.pageListId || "").replace(/[{}]/g, "");
      viewId = String(ctx.viewId || "").replace(/[{}]/g, "");
    }
    if (!siteUrl && fallbackSiteUrl) siteUrl = String(fallbackSiteUrl).replace(/\/$/, "");
    if (!listId && fallbackListId) listId = String(fallbackListId).replace(/[{}]/g, "");
    return { siteUrl: siteUrl, listId: listId, viewId: viewId };
  }

  async function resolveReportSelectionContext(invoke, fallbackSiteUrl, fallbackListId) {
    var ctx = await resolveReportContext(invoke, fallbackSiteUrl, fallbackListId);
    var onListPage = false;
    try {
      var chk = await invoke({ action: "checkListPage" });
      onListPage = !!(chk && chk.isListPage);
    } catch (_) {}
    return { siteUrl: ctx.siteUrl, listId: ctx.listId, viewId: ctx.viewId, onListPage: onListPage };
  }

  function buildReportsPane(panel, invoke, siteUrlFromCtx, currentListIdFromCtx) {
    const host = panel.querySelector(".sp-toolkit-compass-reports-host");
    if (!host) return;
    if (host.dataset.built === "18") {
      return;
    }
    if (host.dataset.building === "1") return;
    host.dataset.building = "1";
    try {
    host.innerHTML =
      "<div class=\"sp-toolkit-compass-reports-main\">" +
      "<div class=\"sp-toolkit-export-progress-console\">" +
      "<div class=\"sp-toolkit-export-progress-headwrap\">" +
      "<div class=\"sp-toolkit-export-progress-bar-wrap\"><div class=\"sp-toolkit-export-progress-bar\"></div></div>" +
      "<div class=\"sp-toolkit-export-progress-head\">" +
      "<div class=\"sp-toolkit-export-progress-status\">" +
      "<span class=\"sp-toolkit-export-progress-spinner\" aria-hidden=\"true\"></span>" +
      "<span class=\"sp-toolkit-export-progress-msg\">Export console</span></div>" +
      "<div class=\"sp-toolkit-export-progress-meta\">" +
      "<span class=\"sp-toolkit-export-progress-elapsed\"></span>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-export-cancel\" title=\"Stop export\">Cancel</button>" +
      "<span class=\"sp-toolkit-export-progress-pct\"></span></div></div></div>" +
      "<div class=\"sp-toolkit-export-progress-log\"></div></div>" +
      "<div class=\"sp-toolkit-compass-reports-config sp-toolkit-compass-form-stack\">" +
      "<div class=\"sp-toolkit-compass-reports-status\" role=\"status\"></div>" +
      "<div class=\"sp-toolkit-compass-field\">" +
      "<label class=\"sp-toolkit-compass-label\">Report type</label>" +
      "<select class=\"sp-toolkit-compass-report-select\">" +
      "<option value=\"exportCSV\">Export List / Library to CSV / Excel</option>" +
      "<option value=\"folderCount\">Folder &amp; Item Counts</option>" +
      "<option value=\"pathLengths\">Path Length Report</option>" +
      "<option value=\"permissionsMatrix\">Permissions Matrix</option>" +
      "</select></div>" +
      "<div class=\"sp-toolkit-compass-report-opts sp-toolkit-compass-form-stack\">" +
      "<div class=\"sp-toolkit-report-lists-row sp-toolkit-compass-matrix-card\" style=\"display:none\">" +
      "<div class=\"sp-toolkit-compass-field sp-toolkit-report-lists-block sp-toolkit-matrix-sites-block\">" +
      "<label class=\"sp-toolkit-compass-label\">Lists / libraries</label>" +
      "<div class=\"sp-toolkit-report-lists-toolbar sp-toolkit-matrix-sites-toolbar\">" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-report-lists-load\">Load lists</button>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-report-lists-expand\">Expand all</button>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-report-lists-collapse\">Collapse all</button>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-report-lists-all\">All</button>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-report-lists-none\">None</button>" +
      "<span class=\"sp-toolkit-report-lists-count\" aria-live=\"polite\"></span></div>" +
      "<input type=\"search\" class=\"sp-toolkit-report-lists-filter sp-toolkit-compass-context-filter\" placeholder=\"Filter by subsite path or list name…\" autocomplete=\"off\" />" +
      "<div class=\"sp-toolkit-report-lists-list\">Click Load lists to choose site lists and libraries.</div></div></div>" +
      "<div class=\"sp-toolkit-matrix-row sp-toolkit-compass-matrix-card\" style=\"display:none\">" +
      "<div class=\"sp-toolkit-compass-field sp-toolkit-matrix-sites-block\">" +
      "<label class=\"sp-toolkit-compass-label\">Sites to scan</label>" +
      "<div class=\"sp-toolkit-matrix-sites-toolbar\">" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-matrix-load\">Load sites</button>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-matrix-all\">All</button>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-matrix-none\">None</button></div>" +
      "<div class=\"sp-toolkit-matrix-sites-list\">Click Load sites to choose root site and subsites.</div></div></div>" +
      "</div>" +
      "<div class=\"sp-toolkit-compass-reports-footer\">" +
      "<div class=\"sp-toolkit-compass-field sp-toolkit-export-row\"><label class=\"sp-toolkit-compass-label\">Format</label>" +
      "<select class=\"sp-toolkit-compass-format-select\"><option value=\"xlsx\">Excel (.xlsx)</option><option value=\"csv\">CSV (.csv)</option></select></div>" +
      "<p class=\"sp-toolkit-compass-muted sp-toolkit-report-bundle-hint\">Multi-list exports download as one .zip folder bundle.</p>" +
      "<div class=\"sp-toolkit-compass-btnrow\">" +
      "<button type=\"button\" class=\"sp-toolkit-compass-primary sp-toolkit-btn-run-export\">Run</button>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-cols\">Choose columns</button>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-rpt-settings\">Settings</button></div></div>" +
      "</div></div>" +
      "<div class=\"sp-toolkit-compass-picker\">" +
      "<div class=\"sp-toolkit-compass-picker-hd\"><button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-pick-back\">← Back</button><span>Choose columns</span></div>" +
      "<div class=\"sp-toolkit-compass-picker-status\"></div><div class=\"sp-toolkit-compass-column-list\"></div>" +
      "<label class=\"sp-toolkit-compass-inline\"><input type=\"checkbox\" class=\"sp-toolkit-chk-versions\"/> Include version history</label>" +
      "<button type=\"button\" class=\"sp-toolkit-compass-primary sp-toolkit-btn-export-selected\">Export selected columns</button></div>" +
      "<div class=\"sp-toolkit-compass-settings\">" +
      "<div class=\"sp-toolkit-compass-picker-hd\"><button type=\"button\" class=\"sp-toolkit-compass-secondary sp-toolkit-btn-set-back\">← Back</button><span>Settings</span></div>" +
      "<div class=\"sp-toolkit-compass-field\"><label class=\"sp-toolkit-compass-label\">Page size</label><select class=\"sp-toolkit-page-size\"><option value=\"500\">500</option><option value=\"1000\">1000</option><option value=\"5000\">5000</option></select></div>" +
      "<div class=\"sp-toolkit-compass-field\"><label class=\"sp-toolkit-compass-label\">Default export format</label><select class=\"sp-toolkit-default-format\"><option value=\"xlsx\">Excel</option><option value=\"csv\">CSV</option></select></div>" +
      "<div class=\"sp-toolkit-matrix-settings-section\">" +
      "<label class=\"sp-toolkit-compass-label\">Permissions matrix</label>" +
      "<label class=\"sp-toolkit-chk-row\"><input type=\"checkbox\" class=\"sp-toolkit-chk-matrix\" checked/><span>Include subsites when loading site list</span></label>" +
      "<label class=\"sp-toolkit-chk-row\"><input type=\"checkbox\" class=\"sp-toolkit-chk-matrix-expand\"/><span>Expand groups in matrix</span></label>" +
      "<label class=\"sp-toolkit-chk-row\"><input type=\"checkbox\" class=\"sp-toolkit-chk-matrix-all-inherited\"/><span>Include ALL inherited rows (slow)</span></label>" +
      "<label class=\"sp-toolkit-chk-row\"><input type=\"checkbox\" class=\"sp-toolkit-chk-matrix-folder-links\" checked/><span>Include folder sharing links</span></label>" +
      "<label class=\"sp-toolkit-chk-row\"><input type=\"checkbox\" class=\"sp-toolkit-chk-matrix-sharing-fetch-all\" checked/><span>Fetch sharing link URLs for all unique items</span></label>" +
      "<div class=\"sp-toolkit-compass-matrix-grid\">" +
      "<div class=\"sp-toolkit-matrix-field\"><label class=\"sp-toolkit-matrix-field-label\" for=\"spToolkitMatrixMaxItems\">Max list items</label>" +
      "<input type=\"number\" id=\"spToolkitMatrixMaxItems\" class=\"sp-toolkit-matrix-max-items\" min=\"0\" value=\"2000\"/></div>" +
      "<div class=\"sp-toolkit-matrix-field\"><label class=\"sp-toolkit-matrix-field-label\" for=\"spToolkitMatrixPageSize\">Page size</label>" +
      "<select id=\"spToolkitMatrixPageSize\" class=\"sp-toolkit-matrix-page-size\"><option value=\"500\">500</option><option value=\"1000\">1000</option><option value=\"5000\" selected>5000</option></select></div></div></div>" +
      "</div>";

    const statusEl = host.querySelector(".sp-toolkit-compass-reports-status");
    const reportSelect = host.querySelector(".sp-toolkit-compass-report-select");
    const formatSelect = host.querySelector(".sp-toolkit-compass-format-select");
    const matrixRow = host.querySelector(".sp-toolkit-matrix-row");
    const reportListsRow = host.querySelector(".sp-toolkit-report-lists-row");
    const reportListsList = host.querySelector(".sp-toolkit-report-lists-list");
    const reportListsFilter = host.querySelector(".sp-toolkit-report-lists-filter");
    const reportListsCount = host.querySelector(".sp-toolkit-report-lists-count");
    const exportRow = host.querySelector(".sp-toolkit-export-row");
    const chkMatrix = host.querySelector(".sp-toolkit-chk-matrix");
    const chkMatrixExpand = host.querySelector(".sp-toolkit-chk-matrix-expand");
    const chkMatrixAllInherited = host.querySelector(".sp-toolkit-chk-matrix-all-inherited");
    const chkMatrixFolderLinks = host.querySelector(".sp-toolkit-chk-matrix-folder-links");
    const chkMatrixSharingFetchAll = host.querySelector(".sp-toolkit-chk-matrix-sharing-fetch-all");
    const matrixMaxItems = host.querySelector(".sp-toolkit-matrix-max-items");
    const matrixPageSize = host.querySelector(".sp-toolkit-matrix-page-size");
    const matrixSitesList = matrixRow.querySelector(".sp-toolkit-matrix-sites-list");
    const progressConsole = host.querySelector(".sp-toolkit-export-progress-console");
    const progressBarWrap = host.querySelector(".sp-toolkit-export-progress-bar-wrap");
    const progressBar = host.querySelector(".sp-toolkit-export-progress-bar");
    const progressElapsed = host.querySelector(".sp-toolkit-export-progress-elapsed");
    const progressMsg = host.querySelector(".sp-toolkit-export-progress-msg");
    const progressLog = host.querySelector(".sp-toolkit-export-progress-log");
    let matrixSitePlan = [];
    let reportListPlan = [];
    let reportListsExpandedPaths = new Set();
    let reportListsTreeIndex = {};
    const REPORT_EXPORT_PREFS_KEY = "reportExportPrefs";
    const reportsMain = host.querySelector(".sp-toolkit-compass-reports-main");
    const picker = host.querySelector(".sp-toolkit-compass-picker");
    const settings = host.querySelector(".sp-toolkit-compass-settings");
    const columnList = host.querySelector(".sp-toolkit-compass-column-list");
    const pickStatus = host.querySelector(".sp-toolkit-compass-picker-status");
    let pickerData = null;

    function setStatus(msg, kind) {
      if (!statusEl) return;
      statusEl.textContent = msg || "";
      statusEl.className = "sp-toolkit-compass-reports-status" + (kind ? " " + kind : "");
    }

    function normalizeListGuid(g) {
      return String(g || "").replace(/[{}]/g, "").toUpperCase();
    }

    function reportListEntryKey(entry) {
      return (String(entry.siteUrl || "").replace(/\/$/, "") + "|" + normalizeListGuid(entry.listId)).toLowerCase();
    }

    function getSelectedReportLists() {
      if (!reportListsList) return [];
      var selectedKeys = {};
      Array.from(reportListsList.querySelectorAll("input[type=checkbox]:checked")).forEach(function (cb) {
        var key = cb.getAttribute("data-key");
        if (key) selectedKeys[key] = true;
      });
      return reportListPlan.filter(function (entry) {
        return selectedKeys[reportListEntryKey(entry)];
      }).map(function (entry) {
        return {
          siteUrl: entry.siteUrl,
          listId: entry.listId,
          listTitle: entry.listTitle,
          siteTitle: entry.siteTitle,
          sitePath: entry.sitePath || "",
          baseTemplate: entry.baseTemplate,
          viewUrl: entry.viewUrl || ""
        };
      });
    }

    function defaultReportListSelectionKeys(entries, currentListId, onListPage, currentSiteUrl) {
      var normCurrent = normalizeListGuid(currentListId);
      if (onListPage && normCurrent) {
        var siteNorm = String(currentSiteUrl || "").replace(/\/$/, "").toLowerCase();
        var matches = entries.filter(function (e) {
          if (normalizeListGuid(e.listId) !== normCurrent) return false;
          if (!siteNorm) return true;
          return String(e.siteUrl || "").replace(/\/$/, "").toLowerCase() === siteNorm;
        });
        if (!matches.length) {
          matches = entries.filter(function (e) { return normalizeListGuid(e.listId) === normCurrent; });
        }
        if (matches.length) return matches.map(reportListEntryKey);
        return [];
      }
      return entries.map(reportListEntryKey);
    }

    function reportSiteGroupKey(entry) {
      return String(entry.siteUrl || "").replace(/\/$/, "").toLowerCase();
    }

    function normalizeTreePathKey(path) {
      return String(path || "").replace(/\/$/, "").toLowerCase() || "/";
    }

    function findReportSiteParentPath(pathKey, allPathKeys) {
      var parentKey = "";
      var bestLen = 0;
      allPathKeys.forEach(function (pk) {
        if (pk === pathKey) return;
        if (pathKey.indexOf(pk + "/") === 0 && pk.length > bestLen) {
          bestLen = pk.length;
          parentKey = pk;
        }
      });
      return parentKey;
    }

    function buildReportSiteTree(entries, libOnly) {
      var sitesByPath = {};
      (entries || []).forEach(function (entry) {
        if (libOnly && entry.baseTemplate !== 101) return;
        var path = formatReportSitePath(entry);
        var pathKey = normalizeTreePathKey(path);
        if (!sitesByPath[pathKey]) {
          sitesByPath[pathKey] = {
            path: path,
            pathKey: pathKey,
            siteUrl: entry.siteUrl || "",
            siteTitle: entry.siteTitle || path || "Site",
            lists: [],
            children: []
          };
        }
        sitesByPath[pathKey].lists.push(entry);
      });
      var pathKeys = Object.keys(sitesByPath).sort();
      var roots = [];
      var nodeByPath = {};
      pathKeys.forEach(function (pathKey) {
        var node = sitesByPath[pathKey];
        node.lists.sort(function (a, b) {
          return String(a.listTitle || "").localeCompare(String(b.listTitle || ""), undefined, { sensitivity: "base" });
        });
        nodeByPath[pathKey] = node;
        var parentKey = findReportSiteParentPath(pathKey, pathKeys);
        if (parentKey && nodeByPath[parentKey]) nodeByPath[parentKey].children.push(node);
        else roots.push(node);
      });
      function sortNodes(nodes) {
        nodes.sort(function (a, b) {
          return String(a.siteTitle || a.path).localeCompare(String(b.siteTitle || b.path), undefined, { sensitivity: "base" });
        });
        nodes.forEach(function (n) {
          if (n.children.length) sortNodes(n.children);
        });
      }
      sortNodes(roots);
      return { roots: roots, nodeByPath: nodeByPath };
    }

    function ensureReportTreePathExpanded(pathKey) {
      if (!pathKey) return;
      reportListsExpandedPaths.add(pathKey);
      var parentKey = findReportSiteParentPath(pathKey, Object.keys(reportListsTreeIndex));
      if (parentKey) ensureReportTreePathExpanded(parentKey);
    }

    function seedReportListsExpandedPaths(roots, nodeByPath, sel) {
      reportListsTreeIndex = nodeByPath || {};
      var validKeys = {};
      Object.keys(reportListsTreeIndex).forEach(function (k) {
        validKeys[k] = true;
      });
      Array.from(reportListsExpandedPaths).forEach(function (k) {
        if (!validKeys[k]) reportListsExpandedPaths.delete(k);
      });
      if (!reportListsExpandedPaths.size) {
        roots.forEach(function (r) {
          reportListsExpandedPaths.add(r.pathKey);
        });
      }
      function walk(node) {
        var hasSelected = node.lists.some(function (e) {
          return sel[reportListEntryKey(e).toLowerCase()];
        });
        if (hasSelected) ensureReportTreePathExpanded(node.pathKey);
        node.children.forEach(walk);
      }
      roots.forEach(walk);
    }

    function setReportTreeExpanded(expanded) {
      if (!reportListsList) return;
      reportListsList.querySelectorAll(".sp-toolkit-report-tree-node").forEach(function (node) {
        var pathKey = node.getAttribute("data-tree-path") || "";
        var toggle = node.querySelector(".sp-toolkit-report-tree-toggle");
        if (expanded) {
          reportListsExpandedPaths.add(pathKey);
          node.classList.remove("sp-toolkit-report-tree-collapsed");
          if (toggle) toggle.setAttribute("aria-expanded", "true");
        } else {
          reportListsExpandedPaths.delete(pathKey);
          node.classList.add("sp-toolkit-report-tree-collapsed");
          if (toggle) toggle.setAttribute("aria-expanded", "false");
        }
      });
    }

    function updateReportTreeSiteCheckboxes() {
      if (!reportListsList) return;
      reportListsList.querySelectorAll(".sp-toolkit-report-tree-node").forEach(function (node) {
        var siteCb = node.querySelector(":scope > .sp-toolkit-report-tree-site-heading .sp-toolkit-report-tree-site-cb");
        if (!siteCb) return;
        var listCbs = Array.from(node.querySelectorAll(".sp-toolkit-report-list-item input[type=checkbox]"));
        if (!listCbs.length) {
          siteCb.checked = false;
          siteCb.indeterminate = false;
          return;
        }
        var checkedCount = listCbs.filter(function (cb) { return cb.checked; }).length;
        siteCb.checked = checkedCount === listCbs.length;
        siteCb.indeterminate = checkedCount > 0 && checkedCount < listCbs.length;
      });
    }

    function renderReportListItem(entry, sel, depth) {
      var key = reportListEntryKey(entry);
      var siteKey = reportSiteGroupKey(entry);
      var id = "cmp-rpt-" + key.replace(/[^a-zA-Z0-9]/g, "_");
      var kindLabel = reportListKindLabel(entry);
      var kindClass = reportListKindClass(entry);
      var searchText = [
        entry.siteTitle,
        formatReportSitePath(entry),
        entry.listTitle,
        kindLabel,
        entry.listId
      ].join(" ").toLowerCase();
      var div = document.createElement("div");
      div.className = "sp-toolkit-matrix-site-item sp-toolkit-report-list-item";
      div.setAttribute("data-site-key", siteKey);
      div.setAttribute("data-search", searchText);
      div.style.setProperty("--tree-depth", String(depth));
      div.innerHTML =
        "<label for=\"" + id + "\"><input type=\"checkbox\" id=\"" + id + "\" data-key=\"" + escapeHtml(key) + "\" " + (sel[key.toLowerCase()] ? "checked" : "") + "/>" +
        "<span class=\"sp-toolkit-report-list-item-main\">" +
        "<span class=\"sp-toolkit-report-list-item-title\">" +
        "<span class=\"sp-toolkit-report-list-badge " + kindClass + "\">" + escapeHtml(kindLabel) + "</span>" +
        "<strong>" + escapeHtml(entry.listTitle || entry.listId) + "</strong></span>" +
        "<span class=\"sp-toolkit-report-list-meta\">" + (entry.itemCount || 0) + " items</span></span></label>";
      return div;
    }

    function renderReportSiteTreeNode(node, sel, depth) {
      var expanded = reportListsExpandedPaths.has(node.pathKey);
      var listCount = node.lists.length;
      var childCount = node.children.length;
      var metaParts = [];
      if (listCount) metaParts.push(listCount + (listCount === 1 ? " list" : " lists"));
      if (childCount) metaParts.push(childCount + (childCount === 1 ? " subsite" : " subsites"));
      var el = document.createElement("div");
      el.className = "sp-toolkit-report-tree-node" + (expanded ? "" : " sp-toolkit-report-tree-collapsed");
      el.setAttribute("data-tree-path", node.pathKey);
      el.setAttribute("data-site-key", reportSiteGroupKey({ siteUrl: node.siteUrl }));
      el.style.setProperty("--tree-depth", String(depth));
      el.innerHTML =
        "<div class=\"sp-toolkit-report-tree-site-heading\">" +
        "<button type=\"button\" class=\"sp-toolkit-report-tree-toggle\" aria-expanded=\"" + (expanded ? "true" : "false") + "\" aria-label=\"Toggle " + escapeHtml(node.siteTitle) + "\">" +
        "<span class=\"sp-toolkit-report-tree-chevron\" aria-hidden=\"true\"></span></button>" +
        "<div class=\"sp-toolkit-report-tree-site-label\">" +
        "<input type=\"checkbox\" class=\"sp-toolkit-report-tree-site-cb\" data-tree-path=\"" + escapeHtml(node.pathKey) + "\" title=\"Select all lists in this site\" />" +
        "<span class=\"sp-toolkit-report-tree-site-text\">" +
        "<span class=\"sp-toolkit-report-lists-site-title\">" + escapeHtml(node.siteTitle) + "</span>" +
        "<span class=\"sp-toolkit-report-lists-site-path\">" + escapeHtml(node.path) + "</span>" +
        (metaParts.length ? "<span class=\"sp-toolkit-report-tree-site-meta\">" + escapeHtml(metaParts.join(" · ")) + "</span>" : "") +
        "</span></div></div>" +
        "<div class=\"sp-toolkit-report-tree-body\"></div>";
      var body = el.querySelector(".sp-toolkit-report-tree-body");
      if (node.lists.length) {
        var listsWrap = document.createElement("div");
        listsWrap.className = "sp-toolkit-report-tree-lists";
        node.lists.forEach(function (entry) {
          listsWrap.appendChild(renderReportListItem(entry, sel, depth + 1));
        });
        body.appendChild(listsWrap);
      }
      node.children.forEach(function (child) {
        body.appendChild(renderReportSiteTreeNode(child, sel, depth + 1));
      });
      return el;
    }

    function formatReportSitePath(entry) {
      var p = String(entry.sitePath || "").trim();
      if (p) return p;
      try {
        return new URL(String(entry.siteUrl || "")).pathname.replace(/\/$/, "") || "/";
      } catch (_) {
        return String(entry.siteUrl || "");
      }
    }

    function reportListKindLabel(entry) {
      return entry.baseTemplate === 101 || entry.isLibrary ? "Library" : "List";
    }

    function reportListKindClass(entry) {
      return entry.baseTemplate === 101 || entry.isLibrary ? "library" : "list";
    }

    function updateReportListsSelectionCount() {
      if (!reportListsCount || !reportListsList) return;
      var visibleItems = Array.from(reportListsList.querySelectorAll(".sp-toolkit-report-list-item:not(.sp-toolkit-report-list-hidden)"));
      var checkedVisible = visibleItems.filter(function (item) {
        var cb = item.querySelector("input[type=checkbox]");
        return cb && cb.checked;
      }).length;
      var totalVisible = visibleItems.length;
      var totalAll = reportListsList.querySelectorAll(".sp-toolkit-report-list-item").length;
      if (!totalAll) {
        reportListsCount.textContent = "";
        return;
      }
      if (totalVisible === totalAll) {
        reportListsCount.textContent = checkedVisible + " of " + totalAll + " selected";
      } else {
        reportListsCount.textContent = checkedVisible + " of " + totalVisible + " shown selected";
      }
    }

    function applyReportListsFilter() {
      if (!reportListsList) return;
      var q = String(reportListsFilter && reportListsFilter.value || "").trim().toLowerCase();
      var items = reportListsList.querySelectorAll(".sp-toolkit-report-list-item");
      items.forEach(function (item) {
        var hay = String(item.getAttribute("data-search") || "").toLowerCase();
        item.classList.toggle("sp-toolkit-report-list-hidden", !!(q && hay.indexOf(q) < 0));
      });
      var treeNodes = reportListsList.querySelectorAll(".sp-toolkit-report-tree-node");
      treeNodes.forEach(function (node) {
        var pathKey = node.getAttribute("data-tree-path") || "";
        var toggle = node.querySelector(".sp-toolkit-report-tree-toggle");
        var hasVisibleList = node.querySelectorAll(".sp-toolkit-report-list-item:not(.sp-toolkit-report-list-hidden)").length > 0;
        var hasVisibleChildSite = false;
        node.querySelectorAll(":scope > .sp-toolkit-report-tree-body > .sp-toolkit-report-tree-node").forEach(function (child) {
          if (!child.classList.contains("sp-toolkit-report-list-hidden")) hasVisibleChildSite = true;
        });
        var show = !q || hasVisibleList || hasVisibleChildSite;
        node.classList.toggle("sp-toolkit-report-list-hidden", !show);
        if (q && show) {
          reportListsExpandedPaths.add(pathKey);
          node.classList.remove("sp-toolkit-report-tree-collapsed");
          if (toggle) toggle.setAttribute("aria-expanded", "true");
        } else if (!q) {
          var expanded = reportListsExpandedPaths.has(pathKey);
          node.classList.toggle("sp-toolkit-report-tree-collapsed", !expanded);
          if (toggle) toggle.setAttribute("aria-expanded", expanded ? "true" : "false");
        }
      });
      updateReportListsSelectionCount();
      updateReportTreeSiteCheckboxes();
    }

    function syncReportListsLayoutMode() {
      applyReportsPaneLayout(panel);
    }

    function renderReportLists(entries, selectedKeys, selectionCtx) {
      if (!reportListsList) return;
      reportListPlan = entries || [];
      var sel = {};
      var keys = selectedKeys;
      if (!keys) {
        var ctx = selectionCtx || {};
        keys = defaultReportListSelectionKeys(
          reportListPlan,
          ctx.listId || host.dataset.reportCurrentListId || currentListIdFromCtx,
          ctx.onListPage != null ? ctx.onListPage : host.dataset.reportOnListPage === "1",
          ctx.siteUrl || host.dataset.reportSiteUrl || ""
        );
      }
      (keys || []).forEach(function (k) { sel[String(k).toLowerCase()] = true; });
      if (!reportListPlan.length) {
        reportListsList.innerHTML = "<span>Click Load lists to choose lists and libraries.</span>";
        if (reportListsFilter) reportListsFilter.value = "";
        updateReportListsSelectionCount();
        syncReportListsLayoutMode();
        return;
      }
      var reportType = reportSelect ? reportSelect.value : "exportCSV";
      var libOnly = reportType === "folderCount" || reportType === "pathLengths";
      var tree = buildReportSiteTree(reportListPlan, libOnly);
      seedReportListsExpandedPaths(tree.roots, tree.nodeByPath, sel);
      reportListsList.classList.add("sp-toolkit-report-lists-tree");
      reportListsList.innerHTML = "";
      tree.roots.forEach(function (node) {
        reportListsList.appendChild(renderReportSiteTreeNode(node, sel, 0));
      });
      if (!reportListsList.children.length) {
        reportListsList.classList.remove("sp-toolkit-report-lists-tree");
        reportListsList.innerHTML = "<span>No document libraries found. Path/folder reports require libraries.</span>";
      }
      applyReportListsFilter();
      syncReportListsLayoutMode();
    }

    function buildReportExportPrefsPayload(extra) {
      return Object.assign({
        siteKey: host.dataset.reportSiteKey || "",
        includeSubsites: !!(chkMatrix && chkMatrix.checked),
        listPlan: reportListPlan,
        selectedKeys: Array.from(reportListsList ? reportListsList.querySelectorAll("input[type=checkbox]:checked") : []).map(function (cb) {
          return cb.getAttribute("data-key");
        }).filter(Boolean)
      }, extra || {});
    }

    function getSelectedMatrixPaths() {
      if (!matrixSitesList) return [];
      return Array.from(matrixSitesList.querySelectorAll("input[type=checkbox]:checked")).map(function (cb) {
        return cb.getAttribute("data-path");
      }).filter(Boolean);
    }

    function renderMatrixSites(plan, selectedPaths) {
      if (!matrixSitesList) return;
      matrixSitePlan = plan || [];
      var sel = {};
      (selectedPaths || plan.map(function (e) { return e.path; })).forEach(function (p) {
        sel[String(p).toLowerCase()] = true;
      });
      if (!plan.length) {
        matrixSitesList.innerHTML = "<span>Click Load sites to choose subsites.</span>";
        return;
      }
      matrixSitesList.innerHTML = "";
      plan.forEach(function (entry) {
        var path = entry.path || "";
        var id = "cmp-mx-" + path.replace(/[^a-zA-Z0-9]/g, "_");
        var div = document.createElement("div");
        div.className = "sp-toolkit-matrix-site-item";
        div.innerHTML =
          '<label><input type="checkbox" id="' + id + '" data-path="' + escapeHtml(path) + '" ' + (sel[path.toLowerCase()] ? "checked" : "") + "/>" +
          "<span><strong>" + escapeHtml(entry.title || path) + "</strong><br/><span style=\"opacity:.75;font-size:11px\">" +
          escapeHtml(path) + (entry.listCount || entry.itemCount ? " · " + (entry.listCount || 0) + " lists · " + (entry.itemCount || 0) + " items" : "") + "</span></span></label>";
        matrixSitesList.appendChild(div);
      });
    }

    function formatCompassElapsed(ms) {
      var sec = Math.max(0, Math.floor(ms / 1000));
      if (sec < 3600) {
        var m = Math.floor(sec / 60);
        var s = sec % 60;
        return m > 0 ? m + "m " + (s < 10 ? "0" : "") + s + "s" : sec + "s";
      }
      return Math.floor(sec / 3600) + "h " + Math.floor((sec % 3600) / 60) + "m";
    }

    function scrollCompassExportLogToEnd(logEl) {
      if (!logEl) return;
      requestAnimationFrame(function () {
        requestAnimationFrame(function () {
          logEl.scrollTop = logEl.scrollHeight;
        });
      });
    }

    let progressPollTimer = null;
    let progressDismissTimer = null;
    const EXPORT_COMPLETION_DISMISS_MS = 5000;
    const EXPORT_STALE_MS = 25000;

    function hideExportProgressConsole() {
      if (!progressConsole) return;
      progressConsole.classList.remove("sp-toolkit-export-console-visible");
      progressConsole.classList.remove("sp-toolkit-export-console-active");
      progressConsole.classList.remove("sp-toolkit-export-console-idle");
      if (progressBarWrap) progressBarWrap.classList.remove("sp-toolkit-export-bar-active");
      if (reportsMain) reportsMain.classList.remove("sp-toolkit-reports-export-active", "sp-toolkit-reports-export-tracking");
      var reportsScrollOff = host.closest(".sp-toolkit-compass-scroll");
      if (reportsScrollOff) reportsScrollOff.classList.remove("sp-toolkit-reports-scroll-locked");
      if (statusEl) statusEl.classList.remove("sp-toolkit-reports-status-hidden");
    }

    function clearExportProgressDismissTimer() {
      if (!progressDismissTimer) return;
      clearTimeout(progressDismissTimer);
      progressDismissTimer = null;
    }

    function scheduleExportProgressDismiss() {
      clearExportProgressDismissTimer();
      progressDismissTimer = setTimeout(function () {
        progressDismissTimer = null;
        try {
          chrome.runtime.sendMessage({ type: "SPCSVExportProgressClear" }, function () {
            hideExportProgressConsole();
          });
        } catch (_) {
          hideExportProgressConsole();
        }
      }, EXPORT_COMPLETION_DISMISS_MS);
    }

    function isStaleExportCompletion(data) {
      if (!data || data.active) return false;
      if (!data.finishedAt) return true;
      return Date.now() - data.finishedAt > EXPORT_COMPLETION_DISMISS_MS;
    }

    function refreshProgressConsole(state) {
      if (!progressConsole) return;
      function apply(data) {
        if (isStaleExportCompletion(data)) {
          clearExportProgressDismissTimer();
          try {
            chrome.runtime.sendMessage({ type: "SPCSVExportProgressClear" });
          } catch (_) {}
          hideExportProgressConsole();
          return;
        }
        if (!data || (!data.active && !(data.log && data.log.length))) {
          clearExportProgressDismissTimer();
          hideExportProgressConsole();
          return;
        }
        progressConsole.classList.add("sp-toolkit-export-console-visible");
        progressConsole.classList.toggle("sp-toolkit-export-console-active", !!data.active);
        progressConsole.classList.toggle("sp-toolkit-export-console-idle", !data.active);
        if (progressBarWrap) progressBarWrap.classList.toggle("sp-toolkit-export-bar-active", !!data.active);
        if (reportsMain) {
          reportsMain.classList.toggle("sp-toolkit-reports-export-active", !!data.active);
          reportsMain.classList.toggle("sp-toolkit-reports-export-tracking", !!(data.active || (data.log && data.log.length)));
        }
        if (data.active) {
          clearExportProgressDismissTimer();
          showReportsMainView();
          startProgressPolling();
          try { chrome.runtime.sendMessage({ type: "SPCSVExportEnsureWorker" }); } catch (_) {}
        } else {
          scheduleExportProgressDismiss();
          stopProgressPolling();
        }
        var reportsScroll = host.closest(".sp-toolkit-compass-scroll");
        if (reportsScroll) reportsScroll.classList.toggle("sp-toolkit-reports-scroll-locked", !!data.active);
        if (statusEl) statusEl.classList.add("sp-toolkit-reports-status-hidden");
        if (progressBar) progressBar.style.width = (data.percent != null ? Math.max(0, Math.min(100, data.percent)) : (data.active ? 8 : 100)) + "%";
        const pctEl = host.querySelector(".sp-toolkit-export-progress-pct");
        if (pctEl) pctEl.textContent = data.percent != null ? data.percent + "%" : (data.active ? "…" : "100%");
        var headline = String(data.message || (data.active ? "Export running…" : (data.success ? "Export complete" : "Export finished"))).replace(/\s*\(still working…\)+/gi, "").trim();
        var lastTouch = Math.max(data.updatedAt || 0, data.pulseAt || 0);
        if (data.active && lastTouch && Date.now() - lastTouch > EXPORT_STALE_MS) {
          headline += " (no recent updates…)";
        }
        if (progressMsg) progressMsg.textContent = headline;
        if (progressElapsed && data.startedAt) {
          progressElapsed.textContent = formatCompassElapsed((data.active ? Date.now() : (data.finishedAt || Date.now())) - data.startedAt);
          progressElapsed.style.display = "";
        } else if (progressElapsed) {
          progressElapsed.style.display = "none";
        }
        if (progressLog && Array.isArray(data.log)) {
          progressLog.textContent = data.log.map(function (line) {
            var d = line.t ? new Date(line.t) : null;
            var ts = d ? d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" }) : "";
            var msg = String(line.msg || "").replace(/\s*\(still working…\)+/gi, "").trim();
            return (ts ? "[" + ts + "] " : "") + msg;
          }).join("\n");
          scrollCompassExportLogToEnd(progressLog);
        }
        noteCompassPanelContentChanged(panel);
      }
      if (state) {
        apply(state);
        return;
      }
      try {
        chrome.runtime.sendMessage({ type: "SPCSVExportProgressGet" }, function (data) {
          if (chrome.runtime.lastError) return;
          apply(data);
        });
      } catch (_) {}
    }

    window.SPOT_refreshExportProgress = refreshProgressConsole;

    function startProgressPolling() {
      if (progressPollTimer) return;
      progressPollTimer = setInterval(function () {
        refreshProgressConsole();
      }, 1000);
    }
    function stopProgressPolling() {
      if (!progressPollTimer) return;
      clearInterval(progressPollTimer);
      progressPollTimer = null;
    }

    if (chrome.storage && chrome.storage.onChanged) {
      chrome.storage.onChanged.addListener(function (changes, area) {
        if (area === "session" && changes.spcsvExportProgress) {
          refreshProgressConsole(changes.spcsvExportProgress.newValue);
          if (changes.spcsvExportProgress.newValue && changes.spcsvExportProgress.newValue.active) {
            startProgressPolling();
          } else {
            stopProgressPolling();
          }
        }
      });
    }

    function buildMatrixPrefsPayload(extra) {
      return Object.assign({
        siteKey: host.dataset.matrixSiteKey || "",
        includeSubsites: !!(chkMatrix && chkMatrix.checked),
        expandGroups: !!(chkMatrixExpand && chkMatrixExpand.checked),
        includeAllInherited: !!(chkMatrixAllInherited && chkMatrixAllInherited.checked),
        includeFolderSharingLinks: !!(chkMatrixFolderLinks && chkMatrixFolderLinks.checked),
        sharingLinkFetchAll: !!(chkMatrixSharingFetchAll && chkMatrixSharingFetchAll.checked),
        maxListItems: matrixMaxItems ? parseInt(matrixMaxItems.value, 10) || 2000 : 2000,
        listItemPageSize: matrixPageSize ? parseInt(matrixPageSize.value, 10) || 5000 : 5000,
        sitePlan: matrixSitePlan || [],
        selectedPaths: getSelectedMatrixPaths()
      }, extra || {});
    }

    function applyMatrixControlPrefs(p) {
      p = p || {};
      if (chkMatrix && p.includeSubsites != null) chkMatrix.checked = !!p.includeSubsites;
      if (chkMatrixExpand && p.expandGroups != null) chkMatrixExpand.checked = !!p.expandGroups;
      if (chkMatrixAllInherited && p.includeAllInherited != null) chkMatrixAllInherited.checked = !!p.includeAllInherited;
      if (chkMatrixFolderLinks && p.includeFolderSharingLinks != null) chkMatrixFolderLinks.checked = !!p.includeFolderSharingLinks;
      if (chkMatrixSharingFetchAll && p.sharingLinkFetchAll != null) chkMatrixSharingFetchAll.checked = !!p.sharingLinkFetchAll;
      if (matrixMaxItems && p.maxListItems != null) matrixMaxItems.value = String(p.maxListItems);
      if (matrixPageSize && p.listItemPageSize != null) matrixPageSize.value = String(p.listItemPageSize);
    }

    async function refreshReportsPaneContext() {
      var selCtx = await resolveReportSelectionContext(invoke, siteUrlFromCtx, currentListIdFromCtx);
      var siteKey = normalizeMatrixSiteKey(selCtx.siteUrl);
      var prevKey = host.dataset.matrixSiteKey || "";
      var prevReportKey = host.dataset.reportSiteKey || "";
      host.dataset.matrixSiteKey = siteKey;
      host.dataset.reportSiteKey = siteKey;
      host.dataset.reportSiteUrl = selCtx.siteUrl || "";
      host.dataset.reportCurrentListId = selCtx.listId || "";
      host.dataset.reportOnListPage = selCtx.onListPage ? "1" : "0";
      var defaultKeys = null;
      chrome.storage.local.get(["matrixExportPrefs", REPORT_EXPORT_PREFS_KEY], function (r) {
        var p = r.matrixExportPrefs || {};
        applyMatrixControlPrefs(p);
        if (siteKey && p.siteKey === siteKey && Array.isArray(p.sitePlan) && p.sitePlan.length) {
          renderMatrixSites(p.sitePlan, p.selectedPaths);
        } else if (siteKey !== prevKey) {
          matrixSitePlan = [];
          if (matrixSitesList) {
            matrixSitesList.innerHTML = "<span>Click Load sites to choose root site and subsites.</span>";
          }
        }
        var rp = r[REPORT_EXPORT_PREFS_KEY] || {};
        if (siteKey && rp.siteKey === siteKey && Array.isArray(rp.listPlan) && rp.listPlan.length) {
          defaultKeys = defaultReportListSelectionKeys(rp.listPlan, selCtx.listId, selCtx.onListPage, selCtx.siteUrl);
          renderReportLists(rp.listPlan, defaultKeys, selCtx);
          chrome.storage.local.set({
            reportExportPrefs: buildReportExportPrefsPayload({
              siteKey: siteKey,
              listPlan: rp.listPlan,
              selectedKeys: defaultKeys
            })
          });
        } else if (siteKey !== prevReportKey) {
          reportListPlan = [];
          if (reportListsList) {
            reportListsList.innerHTML = "<span>Click Load lists to choose site lists and libraries.</span>";
          }
          if (siteKey) void loadReportListPlan();
        }
      });
    }

    async function loadReportListPlan() {
      if (reportListsList) reportListsList.innerHTML = "<span class=\"sp-toolkit-panel-loading\">Loading lists…</span>";
      var selCtx = await resolveReportSelectionContext(invoke, siteUrlFromCtx, currentListIdFromCtx);
      var siteUrl = selCtx.siteUrl || "";
      if (!siteUrl) {
        if (reportListsList) reportListsList.innerHTML = "<span>Open a SharePoint site page first.</span>";
        return;
      }
      var siteKey = normalizeMatrixSiteKey(siteUrl);
      host.dataset.reportSiteKey = siteKey;
      host.dataset.reportSiteUrl = siteUrl;
      host.dataset.reportCurrentListId = selCtx.listId || "";
      host.dataset.reportOnListPage = selCtx.onListPage ? "1" : "0";
      var reportType = reportSelect ? reportSelect.value : "exportCSV";
      var libOnly = reportType === "folderCount" || reportType === "pathLengths";
      return invoke({
        action: "getReportListPlan",
        siteUrl: siteUrl,
        includeSubsites: !!(chkMatrix && chkMatrix.checked),
        librariesOnly: false
      }).then(function (res) {
        if (!res || !res.ok) {
          if (reportListsList) reportListsList.innerHTML = "<span>" + escapeHtml((res && res.error) || "Failed to load lists.") + "</span>";
          return;
        }
        var entries = res.entries || [];
        var selectedKeys = defaultReportListSelectionKeys(entries, selCtx.listId, selCtx.onListPage, selCtx.siteUrl);
        renderReportLists(entries, selectedKeys, selCtx);
        chrome.storage.local.set({
          reportExportPrefs: buildReportExportPrefsPayload({
            siteKey: siteKey,
            listPlan: entries,
            selectedKeys: selectedKeys
          })
        });
        if (libOnly && !reportListsList.querySelector("input[type=checkbox]")) {
          setStatus("No document libraries in selection scope. Folder/path reports need libraries.", "error");
        }
      });
    }

    host._spotRefreshReportsContext = refreshReportsPaneContext;
    void refreshReportsPaneContext();

    host.querySelector(".sp-toolkit-btn-matrix-load")?.addEventListener("click", function () {
      if (matrixSitesList) matrixSitesList.innerHTML = "<span class=\"sp-toolkit-panel-loading\">Loading site inventory…</span>";
      setStatus("Loading sites… large sites can take up to a minute.", "info");
      resolveReportContext(invoke, siteUrlFromCtx, currentListIdFromCtx).then(function (ctx) {
        var siteUrl = ctx.siteUrl || "";
        if (!siteUrl) {
          if (matrixSitesList) matrixSitesList.innerHTML = "<span>Open a SharePoint site page first.</span>";
          setStatus("Open a SharePoint site page first.", "error");
          return;
        }
        var siteKey = normalizeMatrixSiteKey(siteUrl);
        host.dataset.matrixSiteKey = siteKey;
        return invoke({ action: "getMatrixScanPlan", siteUrl: siteUrl, includeSubsites: !!(chkMatrix && chkMatrix.checked) }).then(function (res) {
          if (!res || !res.ok) {
            if (matrixSitesList) matrixSitesList.innerHTML = "<span>" + escapeHtml((res && res.error) || "Failed to load sites.") + "</span>";
            setStatus((res && res.error) || "Failed to load sites.", "error");
            return;
          }
          renderMatrixSites(res.plan || [], (res.plan || []).map(function (e) { return e.path; }));
          setStatus("Loaded " + (res.plan || []).length + " site(s). Select subsites to scan, then Run.", "ok");
          chrome.storage.local.set({
            matrixExportPrefs: buildMatrixPrefsPayload({
              siteKey: siteKey,
              sitePlan: res.plan || [],
              selectedPaths: (res.plan || []).map(function (e) { return e.path; })
            })
          });
        });
      }).catch(function (err) {
        if (matrixSitesList) matrixSitesList.innerHTML = "<span>" + escapeHtml((err && err.message) || "Failed to load sites.") + "</span>";
        setStatus((err && err.message) || "Failed to load sites.", "error");
      });
    });
    host.querySelector(".sp-toolkit-btn-matrix-all")?.addEventListener("click", function () {
      matrixSitesList.querySelectorAll("input[type=checkbox]").forEach(function (cb) { cb.checked = true; });
    });
    host.querySelector(".sp-toolkit-btn-matrix-none")?.addEventListener("click", function () {
      matrixSitesList.querySelectorAll("input[type=checkbox]").forEach(function (cb) { cb.checked = false; });
    });

    host.querySelector(".sp-toolkit-btn-report-lists-load")?.addEventListener("click", function () {
      void loadReportListPlan();
    });
    host.querySelector(".sp-toolkit-btn-report-lists-expand")?.addEventListener("click", function () {
      setReportTreeExpanded(true);
    });
    host.querySelector(".sp-toolkit-btn-report-lists-collapse")?.addEventListener("click", function () {
      setReportTreeExpanded(false);
    });
    host.querySelector(".sp-toolkit-btn-report-lists-all")?.addEventListener("click", function () {
      if (!reportListsList) return;
      reportListsList.querySelectorAll(".sp-toolkit-report-list-item:not(.sp-toolkit-report-list-hidden) input[type=checkbox]").forEach(function (cb) { cb.checked = true; });
      updateReportListsSelectionCount();
      updateReportTreeSiteCheckboxes();
      chrome.storage.local.set({ reportExportPrefs: buildReportExportPrefsPayload() });
    });
    host.querySelector(".sp-toolkit-btn-report-lists-none")?.addEventListener("click", function () {
      if (!reportListsList) return;
      reportListsList.querySelectorAll(".sp-toolkit-report-list-item:not(.sp-toolkit-report-list-hidden) input[type=checkbox]").forEach(function (cb) { cb.checked = false; });
      updateReportListsSelectionCount();
      updateReportTreeSiteCheckboxes();
      chrome.storage.local.set({ reportExportPrefs: buildReportExportPrefsPayload() });
    });
    reportListsFilter?.addEventListener("input", function () {
      applyReportListsFilter();
    });
    reportListsList?.addEventListener("click", function (ev) {
      var toggle = ev.target.closest(".sp-toolkit-report-tree-toggle");
      if (!toggle || !reportListsList.contains(toggle)) return;
      ev.preventDefault();
      var node = toggle.closest(".sp-toolkit-report-tree-node");
      if (!node) return;
      var pathKey = node.getAttribute("data-tree-path") || "";
      if (node.classList.contains("sp-toolkit-report-tree-collapsed")) {
        node.classList.remove("sp-toolkit-report-tree-collapsed");
        reportListsExpandedPaths.add(pathKey);
        toggle.setAttribute("aria-expanded", "true");
      } else {
        node.classList.add("sp-toolkit-report-tree-collapsed");
        reportListsExpandedPaths.delete(pathKey);
        toggle.setAttribute("aria-expanded", "false");
      }
    });
    reportListsList?.addEventListener("change", function (ev) {
      if (ev.target.classList.contains("sp-toolkit-report-tree-site-cb")) {
        var node = ev.target.closest(".sp-toolkit-report-tree-node");
        if (!node) return;
        var checked = ev.target.checked;
        node.querySelectorAll(".sp-toolkit-report-list-item input[type=checkbox]").forEach(function (cb) {
          cb.checked = checked;
        });
        node.querySelectorAll(".sp-toolkit-report-tree-site-cb").forEach(function (cb) {
          cb.checked = checked;
          cb.indeterminate = false;
        });
      }
      updateReportListsSelectionCount();
      updateReportTreeSiteCheckboxes();
      chrome.storage.local.set({ reportExportPrefs: buildReportExportPrefsPayload() });
    });

    refreshProgressConsole();

    host.querySelector(".sp-toolkit-export-cancel")?.addEventListener("click", function () {
      try {
        chrome.runtime.sendMessage({ type: "SPCSVExportCancel" }, function () {
          refreshProgressConsole();
        });
      } catch (_) {}
    });

    function toggleReportOptions() {
      const v = reportSelect.value;
      const isMatrix = v === "permissionsMatrix";
      const showReportLists = v === "exportCSV" || v === "folderCount" || v === "pathLengths";
      const showExportFormat = v === "exportCSV";
      const show = showExportFormat || isMatrix || showReportLists;
      const reportOpts = host.querySelector(".sp-toolkit-compass-report-opts");
      if (reportOpts) reportOpts.style.display = show ? "flex" : "none";
      if (reportListsRow) reportListsRow.style.display = showReportLists ? "flex" : "none";
      matrixRow.style.display = isMatrix ? "flex" : "none";
      exportRow.style.display = showExportFormat ? "flex" : "none";
      var bundleHint = host.querySelector(".sp-toolkit-report-bundle-hint");
      if (bundleHint) bundleHint.style.display = showReportLists ? "" : "none";
      host.querySelector(".sp-toolkit-btn-cols").style.display = showExportFormat ? "" : "none";
      host.classList.toggle("sp-toolkit-reports-matrix-mode", isMatrix);
      syncReportListsLayoutMode();
      if (showReportLists && reportListPlan.length) renderReportLists(reportListPlan);
    }
    reportSelect.addEventListener("change", toggleReportOptions);
    toggleReportOptions();

    getDefaultExportFormatPromise().then(function (f) {
      formatSelect.value = f;
    });

    function showReportsMainView() {
      host.classList.remove("sp-toolkit-reports-settings-active", "sp-toolkit-reports-picker-active");
      settings.classList.remove("sp-toolkit-compass-settings-open");
      picker.classList.remove("sp-toolkit-compass-picker-open");
    }

    function showReportsSettingsView() {
      host.classList.remove("sp-toolkit-reports-picker-active");
      picker.classList.remove("sp-toolkit-compass-picker-open");
      host.classList.add("sp-toolkit-reports-settings-active");
      settings.classList.add("sp-toolkit-compass-settings-open");
    }

    function showReportsPickerView() {
      host.classList.remove("sp-toolkit-reports-settings-active");
      settings.classList.remove("sp-toolkit-compass-settings-open");
      host.classList.add("sp-toolkit-reports-picker-active");
      picker.classList.add("sp-toolkit-compass-picker-open");
    }

    host.querySelector(".sp-toolkit-btn-rpt-settings")?.addEventListener("click", function () {
      showReportsSettingsView();
      getPageSizePromise().then(function (sz) {
        host.querySelector(".sp-toolkit-page-size").value = String(sz);
      });
      getDefaultExportFormatPromise().then(function (f) {
        host.querySelector(".sp-toolkit-default-format").value = f;
      });
      noteCompassPanelContentChanged(panel);
    });
    function saveMatrixPrefsFromControls() {
      chrome.storage.local.set({ matrixExportPrefs: buildMatrixPrefsPayload() });
    }

    host.querySelector(".sp-toolkit-btn-set-back").addEventListener("click", function () {
      const ps = parseInt(host.querySelector(".sp-toolkit-page-size").value, 10);
      chrome.storage.local.set({ pageSize: ps });
      chrome.storage.local.set({ defaultExportFormat: host.querySelector(".sp-toolkit-default-format").value });
      saveMatrixPrefsFromControls();
      showReportsMainView();
      noteCompassPanelContentChanged(panel);
    });

    host.querySelector(".sp-toolkit-btn-cols")?.addEventListener("click", function () {
      showReportsPickerView();
      pickStatus.textContent = "Loading columns…";
      columnList.innerHTML = "";
      invoke({ action: "getColumns" }).then(function (response) {
        if (!response || !response.ok) {
          pickStatus.textContent = (response && response.error) || "Failed to load columns.";
          return;
        }
        pickerData = response;
        pickStatus.textContent = "";
        const defaultSet = new Set(pickerData.defaultColumns || []);
        const requiredSet = new Set(REQUIRED_FIELDS);
        (pickerData.pickerFields || []).forEach(function (f) {
          const iname = f.InternalName || f.Title;
          const title = f.Title || iname;
          const isDefault = defaultSet.has(iname);
          const isRequired = requiredSet.has(iname);
          const div = document.createElement("div");
          div.className = "sp-toolkit-compass-pick-item";
          const id = "cp-col-" + iname.replace(/[^a-zA-Z0-9]/g, "_");
          const showInternal = title !== iname;
          const titleHtml = escapeHtml(title) + (isRequired ? "<span class=\"sp-toolkit-pick-req\" aria-hidden=\"true\"> *</span>" : "");
          const internalHtml = showInternal ? "<span class=\"sp-toolkit-pick-internal\">" + escapeHtml(iname) + "</span>" : "";
          div.innerHTML =
            "<input type=\"checkbox\" id=\"" +
            id +
            "\" class=\"sp-toolkit-compass-pick-cb\" data-internal=\"" +
            escapeHtml(iname) +
            "\" " +
            (isDefault ? "checked " : "") +
            "/>" +
            "<label for=\"" +
            id +
            "\">" +
            titleHtml +
            internalHtml +
            "</label>";
          div.addEventListener("click", function (e) {
            if (e.target.closest("input,label")) return;
            const cb = div.querySelector("input[type=\"checkbox\"]");
            if (cb && !cb.disabled) cb.checked = !cb.checked;
          });
          columnList.appendChild(div);
        });
        noteCompassPanelContentChanged(panel);
      });
    });

    host.querySelector(".sp-toolkit-btn-pick-back")?.addEventListener("click", function () {
      showReportsMainView();
      noteCompassPanelContentChanged(panel);
    });

    host.querySelector(".sp-toolkit-btn-export-selected").addEventListener("click", function () {
      const selected = [];
      columnList.querySelectorAll("input[type=\"checkbox\"]:checked").forEach(function (cb) {
        selected.push(cb.getAttribute("data-internal"));
      });
      if (selected.length === 0) {
        pickStatus.textContent = "Select at least one column.";
        return;
      }
      runExport(selected);
    });

    async function runExport(selectedColumns) {
      setStatus("Starting…", "info");
      const reportType = reportSelect.value;
      const reportCtx = await resolveReportContext(invoke, siteUrlFromCtx, currentListIdFromCtx);
      const siteUrl = reportCtx.siteUrl || "";
      const listId = reportCtx.listId || "";
      const viewId = reportCtx.viewId || "";
      if (!siteUrl) {
        setStatus("Open a SharePoint site or list page first, then try again.", "error");
        return;
      }
      const needsListPicker = reportType === "exportCSV" || reportType === "folderCount" || reportType === "pathLengths";
      let selectedLists = needsListPicker ? getSelectedReportLists() : [];
      if (needsListPicker && !selectedLists.length) {
        if (!reportListPlan.length) {
          await loadReportListPlan();
          selectedLists = getSelectedReportLists();
        }
      }
      if (needsListPicker && !selectedLists.length) {
        setStatus("Select at least one list or library (Load lists, then check items).", "error");
        return;
      }
      if (needsListPicker && (reportType === "folderCount" || reportType === "pathLengths")) {
        selectedLists = selectedLists.filter(function (e) { return e.baseTemplate === 101; });
        if (!selectedLists.length) {
          setStatus("Folder count and path length reports require at least one document library.", "error");
          return;
        }
      }
      const effectiveListId = selectedLists.length === 1 ? normalizeListGuid(selectedLists[0].listId) : listId;
      const effectiveSiteUrl = selectedLists.length === 1 ? String(selectedLists[0].siteUrl || siteUrl).replace(/\/$/, "") : siteUrl;
      const includeVersions = !!(host.querySelector(".sp-toolkit-chk-versions") && host.querySelector(".sp-toolkit-chk-versions").checked);
      const pageLimit = await getPageSizePromise();
      const exportFormat = reportType === "permissionsMatrix" ? "xlsx" : (formatSelect.value || "xlsx");
      const msg = {
        action: "runExportCSV",
        siteUrl: effectiveSiteUrl,
        listId: effectiveListId,
        viewId: viewId,
        exportFilename: "",
        pageLimit: pageLimit,
        includeVersions: includeVersions,
        selectedColumns: selectedColumns && selectedColumns.length ? selectedColumns : null,
        report: reportType,
        format: exportFormat,
        matrixIncludeSubsites: !!(chkMatrix && chkMatrix.checked),
        matrixExpandGroups: !!(chkMatrixExpand && chkMatrixExpand.checked),
        matrixIncludeAllInherited: !!(chkMatrixAllInherited && chkMatrixAllInherited.checked),
        matrixIncludeFolderSharingLinks: !!(chkMatrixFolderLinks && chkMatrixFolderLinks.checked),
        matrixSharingLinkFetchAll: !!(chkMatrixSharingFetchAll && chkMatrixSharingFetchAll.checked),
        matrixMaxListItems: matrixMaxItems ? parseInt(matrixMaxItems.value, 10) || 2000 : 2000,
        matrixListItemPageSize: matrixPageSize ? parseInt(matrixPageSize.value, 10) || 5000 : 5000,
        matrixSelectedPaths: getSelectedMatrixPaths(),
        reportSelectedLists: selectedLists.length ? selectedLists : null,
        rootSiteTitle: selectedLists.length > 1 ? (function () {
          try {
            var segs = new URL(siteUrl).pathname.replace(/\/$/, "").split("/").filter(Boolean);
            return segs.length ? segs[segs.length - 1] : "Site";
          } catch (_) { return "Site"; }
        })() : null
      };
      const response = await invoke(msg);
      const statusMessage =
        response && response.message != null && response.message !== ""
          ? response.message
          : response && response.ok
            ? "Export running in a background tab. Keep working here — open the background tab for Snake and live progress."
            : (response && response.error) || "Export failed.";
      setStatus(statusMessage, response && response.ok ? "ok" : "error");
      if (response && response.ok) {
        refreshProgressConsole();
        startProgressPolling();
        showReportsMainView();
        noteCompassPanelContentChanged(panel);
      }
    }

    host.querySelector(".sp-toolkit-btn-run-export").addEventListener("click", function () {
      if (reportSelect.value === "exportCSV" || reportSelect.value === "folderCount" || reportSelect.value === "pathLengths" || reportSelect.value === "permissionsMatrix") {
        runExport(null);
      }
    });
    host.dataset.built = "18";
    } finally {
      delete host.dataset.building;
    }
  }

  function buildNavTabs(panel, invoke, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId) {
    const nav = panel.querySelector(".sp-toolkit-compass-nav-tabs");
    if (!nav || !siteUrl) return;
    function activateRememberedTab() {
      chrome.storage.local.get([COMPASS_ACTIVE_TAB_KEY, COMPASS_REMEMBER_TAB_KEY], function (st) {
        const remember = st[COMPASS_REMEMBER_TAB_KEY] !== false;
        const paneId = resolveRememberedCompassPane(remember ? st[COMPASS_ACTIVE_TAB_KEY] : null, onListOrLibraryView);
        switchCompassPane(panel, paneId, invoke, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId, { skipPersist: true });
      });
    }
    if (nav.dataset.compassTabsBuilt === "1" && nav.querySelector(".sp-toolkit-compass-pane-tab")) {
      applyListOnlyCompassTabs(panel, onListOrLibraryView);
      refreshCompassTabRegistry(panel);
      activateRememberedTab();
      return;
    }
    delete panel._compassPaneRegistry;
    nav.innerHTML = "";
    const defs = [
      ["quicklinks", "Quick Links"],
      ["context", "Page Props"],
      ["columns", "Columns"],
      ["reports", "Reports"],
      ["refinableProps", "Refinables"],
      ["viewManager", "Views"],
      ["siteContents", "Site Contents"]
    ];
    defs.forEach(function (pair) {
      const id = pair[0];
      const label = pair[1];
      const b = document.createElement("button");
      b.type = "button";
      b.className = "sp-toolkit-header-tab sp-toolkit-compass-pane-tab tab";
      b.setAttribute("data-compass-pane", id);
      b.textContent = label;
      b.addEventListener("click", function (ev) {
        if (b.disabled) return;
        ev.preventDefault();
        ev.stopPropagation();
        switchCompassPane(panel, id, invoke, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId);
      });
      nav.appendChild(b);
    });
    applyListOnlyCompassTabs(panel, onListOrLibraryView);
    ensureCompassTabIndicator(panel);
    nav.dataset.compassTabsBuilt = "1";
    wireCompassPaneRegistry(panel);
    activateRememberedTab();
    scheduleTabIndicatorUpdate(panel, true);
  }

  window.SPOT_applyCompassTabAvailability = function (panel, onListOrLibraryView) {
    if (!panel) panel = document.getElementById("sp-toolkit-lists-panel");
    if (panel) {
      applyListOnlyCompassTabs(panel, onListOrLibraryView);
      scheduleTabIndicatorUpdate(panel, true);
    }
  };

  window.SPOT_reopenCompassPanel = function (panel, invokeToolkitAction, ctx) {
    if (!panel || !invokeToolkitAction) return;
    if (!panel._compassPaneRegistry) {
      wireCompassPaneRegistry(panel);
    } else {
      refreshCompassTabRegistry(panel);
    }
    const siteUrl = (ctx && ctx.siteUrl) || "";
    const onSiteContentsPage = !!(ctx && ctx.onSiteContentsPage);
    const onListOrLibraryView = !!(ctx && ctx.onListOrLibraryView);
    const currentListId = (ctx && ctx.currentListId) || "";
    applyListOnlyCompassTabs(panel, onListOrLibraryView);
    chrome.storage.local.get([COMPASS_ACTIVE_TAB_KEY, COMPASS_REMEMBER_TAB_KEY], function (st) {
      const remember = st[COMPASS_REMEMBER_TAB_KEY] !== false;
      const paneId = resolveRememberedCompassPane(remember ? st[COMPASS_ACTIVE_TAB_KEY] : null, onListOrLibraryView);
      const reg = panel._compassPaneRegistry;
      if (
        reg &&
        reg.activeId === paneId &&
        isCompassPaneContentReady(panel, paneId, onListOrLibraryView)
      ) {
        return;
      }
      switchCompassPane(panel, paneId, invokeToolkitAction, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId);
    });
  };

  window.SPOT_initCompassToolkitPanels = function (panel, invokeToolkitAction, ctx) {
    const siteUrl = (ctx && ctx.siteUrl) || "";
    const onSiteContentsPage = !!(ctx && ctx.onSiteContentsPage);
    const onListOrLibraryView = !!(ctx && ctx.onListOrLibraryView);
    const currentListId = (ctx && ctx.currentListId) || "";
    if (panel.dataset.compassToolListeners === "1") {
      wireCompassPaneRegistry(panel);
      refreshCompassTabRegistry(panel);
      return;
    }
    wireCompassPaneRegistry(panel);
    prefetchCompassPaneData(panel, invokeToolkitAction, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId);
    buildNavTabs(panel, invokeToolkitAction, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId);
    wireViewsPane(panel, onListOrLibraryView, siteUrl, currentListId);
    wireToolkitSettingsPane(panel);
    window.SPOT_openCompassToolkitSettings = function () {
      const active = panel.querySelector(".sp-toolkit-compass-pane.sp-toolkit-compass-pane-active");
      panel.dataset.compassSettingsReturnPane = active ? active.getAttribute("data-compass-pane") || "siteContents" : "siteContents";
      activatePane(panel, "toolkitSettings");
    };
    window.SPOT_noteCompassPanelContentChanged = noteCompassPanelContentChanged;
    window.SPOT_resetCompassPanelSize = resetCompassPanelSize;
    window.SPOT_updateCompassTabIndicator = updateTabIndicator;
    window.SPOT_updateCompassPanelTitle = updateCompassPanelTitle;
    window.SPOT_clearCompassPagePropsCache = clearPagePropsCacheForUrl;
    if (!window._SPOT_compassPanelResizeBound) {
      window._SPOT_compassPanelResizeBound = true;
      window.addEventListener("resize", function () {
        const p = document.getElementById("sp-toolkit-lists-panel");
        if (p && typeof window.SPOT_noteCompassPanelContentChanged === "function") {
          window.SPOT_noteCompassPanelContentChanged(p);
        }
        if (p && typeof window.SPOT_updateCompassTabIndicator === "function") {
          window.SPOT_updateCompassTabIndicator(p, true);
        }
      });
    }
    noteCompassPanelContentChanged(panel);
    scheduleTabIndicatorUpdate(panel, true);
    if (panel.dataset.compassToolListeners === "1") return;
    panel.dataset.compassToolListeners = "1";

    const ctxFilter = panel.querySelector(".sp-toolkit-compass-context-filter");
    if (ctxFilter) {
      ctxFilter.addEventListener("input", function () {
        renderPagePropsList(panel);
      });
    }
    const btnCtx = panel.querySelector(".sp-toolkit-btn-refresh-context");
    if (btnCtx) {
      btnCtx.addEventListener("click", function () {
        clearPagePropsCacheForUrl(pagePropsCacheKey());
        invalidatePanelPageProps(panel);
        loadPagePropsPane(panel, invokeToolkitAction);
      });
    }
    const colFilt = panel.querySelector(".sp-toolkit-compass-columns-filter");
    if (colFilt) {
      colFilt.addEventListener("input", function () {
        renderColumnsList(panel);
      });
    }
    const refinType = panel.querySelector(".sp-toolkit-compass-refinable-type");
    const refinNum = panel.querySelector(".sp-toolkit-compass-refinable-num");
    if (refinType) refinType.addEventListener("change", function () { loadRefinableMappings(panel, invokeToolkitAction, siteUrl); });
    if (refinNum) refinNum.addEventListener("change", function () { loadRefinableMappings(panel, invokeToolkitAction, siteUrl); });
    const btnRef = panel.querySelector(".sp-toolkit-btn-refinable-config");
    if (btnRef) {
      btnRef.addEventListener("click", function () {
        openRefinableAdmin(panel, siteUrl);
      });
    }

    window.SPOT_prefetchCompassUniversalSearchData = async function (prefPanel, prefInvoke, prefCtx) {
      const su = (prefCtx && prefCtx.siteUrl) || "";
      const osc = !!(prefCtx && prefCtx.onSiteContentsPage);
      await fillQuickLinksPane(prefPanel, prefInvoke, su, osc);
      await loadPagePropsPane(prefPanel, prefInvoke);
    };

    window.SPOT_compassNavigateToPane = async function (paneId) {
      switchCompassPane(panel, paneId, invokeToolkitAction, siteUrl, onSiteContentsPage, onListOrLibraryView, currentListId);
    };

    if (typeof window.SPOT_initCompassUniversalSearch === "function") {
      window.SPOT_initCompassUniversalSearch(panel, invokeToolkitAction, {
        siteUrl: siteUrl,
        onSiteContentsPage: onSiteContentsPage,
        onListOrLibraryView: onListOrLibraryView,
        currentListId: currentListId
      });
    }
    chrome.storage.local.get("compassUniversalSearchOpenOnLoad", function (st) {
      if (!st || !st.compassUniversalSearchOpenOnLoad) return;
      if (typeof window.SPOT_compassOpenUniversalSearch === "function") {
        scheduleAfterCompassTabPaint(function () {
          window.SPOT_compassOpenUniversalSearch();
        });
      }
    });
  };
})();
