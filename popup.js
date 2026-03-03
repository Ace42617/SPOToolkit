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

// --- Tabs: Quick links, Page Properties, Reports ---
const TAB_IDS = ["quicklinks", "context", "searchSchema", "reports", "refinableProps", "viewManager"];
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
// Restore last active tab on open
(async function restorePopupTab() {
  await (async function updateListTabsVisibility() {
    const tabColumns = document.getElementById("tabColumns");
    const tabViewManager = document.getElementById("tabViewManager");
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
        if (tabColumns) tabColumns.style.display = "none";
        if (tabViewManager) tabViewManager.style.display = "none";
        return;
      }
      const response = await chrome.tabs.sendMessage(tab.id, { action: "checkListPage" });
      const show = response?.isListPage ? "" : "none";
      if (tabColumns) tabColumns.style.display = show;
      if (tabViewManager) tabViewManager.style.display = show;
    } catch (_) {
      if (tabColumns) tabColumns.style.display = "none";
      if (tabViewManager) tabViewManager.style.display = "none";
    }
  })();
  const { popupActiveTab } = await chrome.storage.local.get("popupActiveTab");
  const id = popupActiveTab && TAB_IDS.includes(popupActiveTab) ? popupActiveTab : "quicklinks";
  const tabEl = document.querySelector(".tab[data-tab=\"" + id + "\"]");
  if (tabEl && tabEl.offsetParent !== null) switchToTab(id);
  else switchToTab("quicklinks");
})();

// Re-apply saved tab when popup is shown again (e.g. after tab refresh then reopen)
function reapplySavedTab() {
  chrome.storage.local.get("popupActiveTab", ({ popupActiveTab }) => {
    const id = popupActiveTab && TAB_IDS.includes(popupActiveTab) ? popupActiveTab : "quicklinks";
    const tabEl = document.querySelector(".tab[data-tab=\"" + id + "\"]");
    if (tabEl && tabEl.offsetParent !== null) switchToTab(id);
  });
}
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") reapplySavedTab(); });
window.addEventListener("pageshow", reapplySavedTab);

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

async function renderQuickLinks() {
  const hint = document.getElementById("quicklinksHint");
  const content = document.getElementById("quicklinksContent");
  const currentSite = document.getElementById("quicklinksCurrentSite");
  const currentUser = document.getElementById("quicklinksCurrentUser");
  const modes = document.getElementById("quicklinksModes");
  const tenant = document.getElementById("quicklinksTenant");
  if (!content || !currentSite) return;
  currentSite.innerHTML = "";
  currentUser.innerHTML = "";
  modes.innerHTML = "";
  tenant.innerHTML = "";
  let tab = null;
  try {
    const [t] = await chrome.tabs.query({ active: true, currentWindow: true });
    tab = t;
    if (!tab?.url?.includes("sharepoint.com")) {
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

  addQuickLink(currentSite, "Site settings", siteBase + "/_layouts/15/settings.aspx", "settings");
  addQuickLink(currentSite, "Site contents", siteBase + "/_layouts/15/viewlsts.aspx", "folder");
  addQuickLink(currentSite, "Recycle bin", siteBase + "/_layouts/15/RecycleBin.aspx", "trash");
  addQuickLink(currentSite, "All People", siteBase + "/_layouts/15/people.aspx?MembershipGroupId=0", "users");
  addQuickLink(currentSite, "Storage metrics", siteBase + "/_layouts/15/storman.aspx", "chart");
  addQuickLink(currentSite, "ACS Grant App", siteBase + "/_layouts/15/appinv.aspx", "key");

  addQuickLink(currentUser, "Edit user profile", siteBase + "/_layouts/15/me.aspx", "user");
  addQuickLink(currentUser, "Login as another user", siteBase + "/_layouts/15/closeConnection.aspx?loginasanotheruser=1", "logout");

  const pageUrl = tab.url.split("?")[0];
  const q = (param) => pageUrl + (pageUrl.indexOf("?") >= 0 ? "&" : "?") + param;
  addQuickLink(modes, "MaintenanceMode", q("MaintenanceMode=true"), "wrench");
  addQuickLink(modes, "WebView", q("env=WebView"), "monitor");
  addQuickLink(modes, "WebViewList", q("env=WebViewList"), "list");

  addQuickLink(tenant, "Admin center", `https://${adminHost}`, "shield");
  addQuickLink(tenant, "Tenant site settings", `https://${adminHost}/_layouts/15/online/tenantsettings.aspx`, "settings");
  addQuickLink(tenant, "User profiles", `https://${adminHost}/_layouts/15/TenantProfileAdmin/ManageUserProfileServiceApplication.aspx`, "users");
  addQuickLink(tenant, "Term store", `https://${host}/_layouts/15/termstoremanager.aspx`, "tag");
  addQuickLink(tenant, "Search administration", `https://${adminHost}/_layouts/15/searchadmin/TA_SearchAdministration.aspx`, "search");
  addQuickLink(tenant, "API access", `https://${adminHost}/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement`, "key");
  addQuickLink(tenant, "Teams admin", "https://admin.teams.microsoft.com/dashboard", "teams");
  addQuickLink(tenant, "App catalog", `https://${host}/sites/AppCatalog/_layouts/15/appStore.aspx`, "package");
  addQuickLink(tenant, "Classic app catalog", `https://${host}/sites/AppCatalog`, "archive");
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
    if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
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
      navigator.clipboard.writeText(valStr).then(() => {
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
document.getElementById("contextFilterInput").addEventListener("input", () => renderContextList());
document.getElementById("contextFilterByValue").addEventListener("change", () => renderContextList());

document.getElementById("btnOpenViewManager").addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
    alert("Open a SharePoint list or library page first, then open View Manager.");
    return;
  }
  chrome.tabs.create({ url: chrome.runtime.getURL("views.html?tabId=" + tab.id) });
});

// --- Search Schema tab: managed metadata columns from site collection ---
let searchSchemaColumns = [];

async function loadSearchSchema() {
  const countEl = document.getElementById("searchSchemaCount");
  const listEl = document.getElementById("searchSchemaList");
  const errEl = document.getElementById("searchSchemaError");
  if (!listEl || !errEl) return;
  errEl.style.display = "none";
  listEl.innerHTML = "";
  if (countEl) countEl.textContent = "Loading…";
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
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
    renderSearchSchemaList();
  } catch (e) {
    searchSchemaColumns = [];
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
  const filtered = q
    ? searchSchemaColumns.filter(
        (c) =>
          (c.internalName && c.internalName.toLowerCase().indexOf(q) >= 0) ||
          (c.title && c.title.toLowerCase().indexOf(q) >= 0) ||
          (c.type && c.type.toLowerCase().indexOf(q) >= 0) ||
          (c.group && c.group.toLowerCase().indexOf(q) >= 0) ||
          (c.crawledProperty && c.crawledProperty.toLowerCase().indexOf(q) >= 0)
      )
    : searchSchemaColumns;
  if (countEl) countEl.textContent = "Columns: " + filtered.length;
  listEl.innerHTML = "";
  filtered.forEach((c) => {
    const row = document.createElement("div");
    row.className = "search-schema-row";
    row.innerHTML =
      '<span class="col-name">' + escapeHtml(c.title || "") +
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
    if (!tab?.url?.includes("sharepoint.com")) {
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
    if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
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
      refinableMappingsOut.innerHTML = "<span class=\"refinable-mappings-header\">Crawled properties</span>" + (res.alias ? "<p class=\"refinable-alias\">Alias: " + escapeHtml(res.alias) + "</p>" : "") + "<p>No crawled properties mapped for " + escapeHtml(propertyName) + ".</p>";
      refinableMappingsOut.className = "refinable-mappings-out";
      return;
    }
    var html = "";
    if (res.alias) html += "<p class=\"refinable-alias\">Alias: " + escapeHtml(res.alias) + "</p>";
    html += "<span class=\"refinable-mappings-header\">Crawled properties (" + res.mappings.length + ")</span><ul class=\"refinable-mappings-list\">";
    for (var i = 0; i < res.mappings.length; i++) {
      var m = res.mappings[i];
      var typeSpan = m.type ? "<span class=\"refinable-mapping-type\">" + escapeHtml(m.type) + "</span>" : "";
      html += "<li>" + escapeHtml(m.name) + typeSpan + "</li>";
    }
    html += "</ul>";
    refinableMappingsOut.innerHTML = html;
    refinableMappingsOut.className = "refinable-mappings-out";
  } catch (e) {
    refinableMappingsOut.textContent = "Error: " + (e.message || String(e)) + ". Try refreshing the SharePoint tab.";
    refinableMappingsOut.className = "refinable-mappings-out error";
  }
}

// Run once so number dropdown is populated if user opens Refinable Props first
initRefinableProps();

function parseContextFromUrl(pageUrl) {
  if (!pageUrl || !pageUrl.includes("sharepoint.com")) return { webAbsoluteUrl: "", pageListId: "", viewId: "" };
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
    if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
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
    if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
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
    if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
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
    var res = null;
    if (results && results.length) {
      // Prefer main frame (index 0) so library vs list reflects the primary page, not an iframe.
      var main = results[0] && results[0].result;
      if (main && main.isListPage === true) {
        res = main;
      } else {
        for (var i = 0; i < results.length; i++) {
          var r = results[i] && results[i].result;
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
