const btnExport = document.getElementById("btnExport");
const btnChooseColumns = document.getElementById("btnChooseColumns");
const btnSettings = document.getElementById("btnSettings");
const btnSettingsBack = document.getElementById("btnSettingsBack");
const pageSizeSelect = document.getElementById("pageSizeSelect");
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

const DEFAULT_PAGE_SIZE = 5000;

// --- Tabs: Quick links, Page Properties, Reports ---
document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    const id = tab.dataset.tab;
    document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));
    tab.classList.add("active");
    const panel = document.getElementById(id + "Panel");
    if (panel) panel.classList.add("active");
    document.body.classList.toggle("columns-tab-active", id === "searchSchema");
    if (id === "quicklinks") renderQuickLinks();
    if (id === "context") loadContextInfo();
    if (id === "searchSchema") loadSearchSchema();
  });
});
// Default tab is Quick links; populate it on open
renderQuickLinks();
// Show Columns and View Manager tabs only when current page is a list or library
(async function updateListTabsVisibility() {
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

function escapeHtml(s) {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}

function addQuickLink(ul, label, href) {
  const li = document.createElement("li");
  li.innerHTML = `<a href="${escapeHtml(href)}" target="_blank" rel="noopener">${escapeHtml(label)}</a>`;
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

  addQuickLink(currentSite, "Site settings", siteBase + "/_layouts/15/settings.aspx");
  addQuickLink(currentSite, "Tenant site settings", siteBase + "/_layouts/15/tenantsettings.aspx");
  addQuickLink(currentSite, "Site contents", siteBase + "/_layouts/15/viewlsts.aspx");
  addQuickLink(currentSite, "Recycle bin", siteBase + "/_layouts/15/RecycleBin.aspx");
  addQuickLink(currentSite, "All People", siteBase + "/_layouts/15/people.aspx?MembershipGroupId=0");
  addQuickLink(currentSite, "Storage metrics", siteBase + "/_layouts/15/storman.aspx");
  addQuickLink(currentSite, "ACS Grant App", siteBase + "/_layouts/15/appinv.aspx");

  addQuickLink(currentUser, "Edit user profile", siteBase + "/_layouts/15/me.aspx");
  addQuickLink(currentUser, "Login as another user", siteBase + "/_layouts/15/closeConnection.aspx?loginasanotheruser=1");

  const pageUrl = tab.url.split("?")[0];
  const q = (param) => pageUrl + (pageUrl.indexOf("?") >= 0 ? "&" : "?") + param;
  addQuickLink(modes, "?MaintenanceMode=true", q("MaintenanceMode=true"));
  addQuickLink(modes, "?env=WebView", q("env=WebView"));
  addQuickLink(modes, "?env=WebViewList", q("env=WebViewList"));

  addQuickLink(tenant, "Admin center", `https://${adminHost}`);
  addQuickLink(tenant, "User profiles", `https://${adminHost}/_layouts/15/tenantprofileadmin/home.aspx`);
  addQuickLink(tenant, "Term store", `https://${host}/_layouts/15/termstoremanager.aspx`);
  addQuickLink(tenant, "Search administration", `https://${adminHost}/_layouts/15/searchadmin/default.aspx`);
  addQuickLink(tenant, "API access", `https://${adminHost}/_layouts/15/apiaccess.aspx`);
  addQuickLink(tenant, "Teams admin", "https://admin.teams.microsoft.com");
  addQuickLink(tenant, "App catalog", `https://${host}/sites/AppCatalog`);
  addQuickLink(tenant, "Classic app catalog", `https://${host}/sites/AppCatalog`);
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

async function loadSettings() {
  const size = await getPageSize();
  pageSizeSelect.value = String(size);
}

function savePageSize() {
  const v = parseInt(pageSizeSelect.value, 10);
  chrome.storage.local.set({ pageSize: v });
}

function setStatus(message, type) {
  statusEl.textContent = message;
  statusEl.className = type || "info";
  statusEl.style.display = "block";
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
  const inPicker = pickerPanel.classList.contains("visible");
  const el = document.getElementById(inPicker ? "chkIncludeVersionsPicker" : "chkIncludeVersions");
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
}
if (reportSelect) reportSelect.addEventListener("change", toggleReportOptions);

btnExport.addEventListener("click", () => {
  if (reportSelect && (reportSelect.value === "exportCSV" || reportSelect.value === "folderCount" || reportSelect.value === "pathLengths" || reportSelect.value === "permissionsMatrix")) runExport();
});

btnChooseColumns.addEventListener("click", openColumnPicker);

btnSettings.addEventListener("click", () => {
  mainPanel.classList.add("hidden");
  pickerPanel.classList.remove("visible");
  settingsPanel.classList.add("visible");
  loadSettings();
});

btnSettingsBack.addEventListener("click", () => {
  settingsPanel.classList.remove("visible");
  mainPanel.classList.remove("hidden");
  savePageSize();
});

pageSizeSelect.addEventListener("change", savePageSize);

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
