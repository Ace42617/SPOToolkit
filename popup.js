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
    document.body.classList.toggle("views-tab-active", id === "views");
    document.body.classList.toggle("columns-tab-active", id === "searchSchema");
    if (id === "quicklinks") renderQuickLinks();
    if (id === "context") loadContextInfo();
    if (id === "searchSchema") loadSearchSchema();
    if (id === "views") loadViewsTab();
  });
});
// Default tab is Quick links; populate it on open
renderQuickLinks();
// Show Columns and Views tabs only when current page is a list or library
(async function updateListTabsVisibility() {
  const tabColumns = document.getElementById("tabColumns");
  const tabViews = document.getElementById("tabViews");
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
      if (tabColumns) tabColumns.style.display = "none";
      if (tabViews) tabViews.style.display = "none";
      return;
    }
    const response = await chrome.tabs.sendMessage(tab.id, { action: "checkListPage" });
    const show = response?.isListPage ? "" : "none";
    if (tabColumns) tabColumns.style.display = show;
    if (tabViews) tabViews.style.display = show;
  } catch (_) {
    if (tabColumns) tabColumns.style.display = "none";
    if (tabViews) tabViews.style.display = "none";
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
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => (typeof window._spPageContextInfo !== "undefined" ? window._spPageContextInfo : null),
      world: "MAIN"
    });
    const ctx = results && results[0] && results[0].result;
    contextData = ctx && typeof ctx === "object" ? ctx : null;
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
          (c.group && c.group.toLowerCase().indexOf(q) >= 0)
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
      '</span><span class="col-type">' + escapeHtml(c.type || "") + '</span>';
    listEl.appendChild(row);
  });
}

document.getElementById("btnRefreshSearchSchema").addEventListener("click", () => loadSearchSchema());
document.getElementById("searchSchemaFilter").addEventListener("input", () => renderSearchSchemaList());

// --- Views tab ---
let viewsData = null; // { views, fields, listId, listTitle, viewDetails? }
let viewColumnOrder = []; // internal names in display order
let viewFilters = [{ column: "", op: "Eq", value: "" }]; // { column, op, value }[]

const FILTER_OPS = [
  { value: "Eq", label: "Equals" },
  { value: "Neq", label: "Not equal" },
  { value: "Gt", label: "Greater than" },
  { value: "Geq", label: "Greater or equal" },
  { value: "Lt", label: "Less than" },
  { value: "Leq", label: "Less or equal" },
  { value: "Contains", label: "Contains" },
  { value: "BeginsWith", label: "Begins with" }
];

function getFieldOpts() {
  return (viewsData && viewsData.fields) ? viewsData.fields.map((f) => ({ value: f.internalName, label: f.title || f.internalName })) : [];
}

function renderViewColumnsList() {
  const viewColumnsList = document.getElementById("viewColumnsList");
  if (!viewColumnsList || !viewsData) return;
  const fieldByInternal = {};
  (viewsData.fields || []).forEach((f) => { fieldByInternal[f.internalName] = f; });
  viewColumnsList.innerHTML = "";
  (viewColumnOrder.length ? viewColumnOrder : (viewsData.fields || []).map((f) => f.internalName)).forEach((iname, idx) => {
    const f = fieldByInternal[iname] || { internalName: iname, title: iname };
    const div = document.createElement("div");
    div.className = "view-column-item";
    const id = "viewcol-" + (iname || "").replace(/[^a-zA-Z0-9]/g, "_");
    div.innerHTML =
      "<input type=\"checkbox\" id=\"" + id + "\" data-internal=\"" + escapeHtml(iname) + "\">" +
      "<label for=\"" + id + "\">" + escapeHtml(f.title || f.internalName || iname) + "</label>" +
      "<span class=\"col-move\"><button type=\"button\" class=\"col-move-up\" data-internal=\"" + escapeHtml(iname) + "\" title=\"Move up\">↑</button><button type=\"button\" class=\"col-move-down\" data-internal=\"" + escapeHtml(iname) + "\" title=\"Move down\">↓</button></span>";
    viewColumnsList.appendChild(div);
  });
  viewColumnsList.querySelectorAll(".col-move-up").forEach((btn) => {
    btn.addEventListener("click", () => {
      const iname = btn.dataset.internal;
      const i = viewColumnOrder.indexOf(iname);
      if (i > 0) {
        viewColumnOrder.splice(i, 1);
        viewColumnOrder.splice(i - 1, 0, iname);
        renderViewColumnsList();
      }
    });
  });
  viewColumnsList.querySelectorAll(".col-move-down").forEach((btn) => {
    btn.addEventListener("click", () => {
      const iname = btn.dataset.internal;
      const i = viewColumnOrder.indexOf(iname);
      if (i >= 0 && i < viewColumnOrder.length - 1) {
        viewColumnOrder.splice(i, 1);
        viewColumnOrder.splice(i + 1, 0, iname);
        renderViewColumnsList();
      }
    });
  });
}

function renderViewFilters() {
  const container = document.getElementById("viewFiltersContainer");
  const fieldOpts = getFieldOpts();
  if (!container) return;
  container.innerHTML = "";
  viewFilters.forEach((f, idx) => {
    const block = document.createElement("div");
    block.className = "view-filter-block";
    const colId = "viewFilterCol_" + idx;
    const opId = "viewFilterOp_" + idx;
    const valId = "viewFilterVal_" + idx;
    let colOpts = "<option value=\"\">— None —</option>";
    fieldOpts.forEach((o) => { colOpts += "<option value=\"" + escapeHtml(o.value) + "\">" + escapeHtml(o.label) + "</option>"; });
    let opOpts = "";
    FILTER_OPS.forEach((o) => { opOpts += "<option value=\"" + o.value + "\"" + (f.op === o.value ? " selected" : "") + ">" + o.label + "</option>"; });
    block.innerHTML =
      "<div class=\"filter-row\"><label>Column:</label><select id=\"" + colId + "\" class=\"view-filter-col\">" + colOpts + "</select></div>" +
      "<div class=\"filter-row\"><label>Operator:</label><select id=\"" + opId + "\" class=\"view-filter-op\">" + opOpts + "</select></div>" +
      "<div class=\"filter-row\"><label>Value:</label><input type=\"text\" id=\"" + valId + "\" class=\"view-filter-val\" placeholder=\"Filter value\"></div>" +
      (viewFilters.length > 1 ? "<a href=\"#\" class=\"view-filter-remove\" data-idx=\"" + idx + "\">Remove filter</a>" : "");
    container.appendChild(block);
    const colEl = document.getElementById(colId);
    const opEl = document.getElementById(opId);
    const valEl = document.getElementById(valId);
    if (colEl) colEl.value = f.column || "";
    if (opEl) opEl.value = f.op || "Eq";
    if (valEl) valEl.value = f.value || "";
    block.querySelector(".view-filter-remove")?.addEventListener("click", (e) => {
      e.preventDefault();
      viewFilters.splice(parseInt(block.querySelector(".view-filter-remove").dataset.idx, 10), 1);
      if (viewFilters.length === 0) viewFilters = [{ column: "", op: "Eq", value: "" }];
      renderViewFilters();
    });
  });
  container.querySelectorAll(".view-filter-col").forEach((el, i) => { el.addEventListener("change", () => { viewFilters[i].column = el.value; }); });
  container.querySelectorAll(".view-filter-op").forEach((el, i) => { el.addEventListener("change", () => { viewFilters[i].op = el.value; }); });
  container.querySelectorAll(".view-filter-val").forEach((el, i) => { el.addEventListener("input", () => { viewFilters[i].value = el.value; }); });
}

function collectViewFiltersFromDom() {
  const container = document.getElementById("viewFiltersContainer");
  if (!container) return viewFilters;
  const blocks = container.querySelectorAll(".view-filter-block");
  const out = [];
  blocks.forEach((block, i) => {
    const col = block.querySelector(".view-filter-col");
    const val = block.querySelector(".view-filter-val");
    const op = block.querySelector(".view-filter-op");
    if (col && col.value) out.push({ column: col.value, op: (op && op.value) || "Eq", valueType: "Text", value: (val && val.value) || "" });
  });
  return out;
}

async function loadViewsTab(selectViewId = null) {
  const viewSelect = document.getElementById("viewSelect");
  const viewNameInput = document.getElementById("viewNameInput");
  const viewColumnsList = document.getElementById("viewColumnsList");
  const viewSortColumn = document.getElementById("viewSortColumn");
  const viewGroupBy = document.getElementById("viewGroupBy");
  const viewsStatus = document.getElementById("viewsStatus");
  if (!viewSelect || !viewColumnsList) return;

  viewSelect.innerHTML = "<option value=\"\">Loading…</option>";
  viewColumnsList.innerHTML = "";
  document.getElementById("viewFiltersContainer").innerHTML = "";
  viewsStatus.style.display = "none";
  viewsData = null;

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !tab.url?.includes("sharepoint.com")) {
      viewSelect.innerHTML = "<option value=\"\">Open a list page first</option>";
      return;
    }
    const response = await chrome.tabs.sendMessage(tab.id, { action: "getViewsData" });
    if (!response?.ok) {
      viewSelect.innerHTML = "<option value=\"\">" + (response?.error || "Failed to load") + "</option>";
      return;
    }
    viewsData = { views: response.views || [], fields: response.fields || [], listId: response.listId || "", listTitle: response.listTitle || "" };
    viewColumnOrder = (viewsData.fields || []).map((f) => f.internalName);
    viewFilters = [{ column: "", op: "Eq", value: "" }];
    // Populate view dropdown
    viewSelect.innerHTML = "";
    const optNew = document.createElement("option");
    optNew.value = "__new__";
    optNew.textContent = "— Create new view —";
    viewSelect.appendChild(optNew);
    (viewsData.views || []).forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v.id || "";
      opt.textContent = (v.title || "View") + (v.defaultView ? " (default)" : "");
      viewSelect.appendChild(opt);
    });
    const fieldOpts = getFieldOpts();
    [viewSortColumn, viewGroupBy].forEach((sel) => {
      if (!sel) return;
      const current = sel.value;
      sel.innerHTML = "<option value=\"\">— None —</option>";
      fieldOpts.forEach((f) => {
        const o = document.createElement("option");
        o.value = f.value;
        o.textContent = f.label;
        sel.appendChild(o);
      });
      sel.value = current || "";
    });
    renderViewColumnsList();
    renderViewFilters();
    viewNameInput.value = "";
    viewSortColumn.value = "";
    document.getElementById("viewSortDirection").value = "asc";
    viewGroupBy.value = "";
    viewSelect.value = selectViewId && viewsData.views.some((v) => v.id === selectViewId) ? selectViewId : "__new__";
    viewSelect.dispatchEvent(new Event("change"));
  } catch (e) {
    viewSelect.innerHTML = "<option value=\"\">Error: " + escapeHtml(e.message || "Could not reach page") + "</option>";
  }
}

function onViewsViewSelectChange() {
  const viewSelect = document.getElementById("viewSelect");
  const viewNameRow = document.getElementById("viewNameRow");
  const viewNameInput = document.getElementById("viewNameInput");
  const val = viewSelect && viewSelect.value;
  const isNew = val === "__new__" || !val;
  if (viewNameRow) viewNameRow.style.display = isNew ? "" : "none";
  if (viewNameInput) viewNameInput.placeholder = isNew ? "View name (for new view)" : "View name";
  if (!viewsData || !val || val === "__new__") return;
  (async () => {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab?.id) return;
      const response = await chrome.tabs.sendMessage(tab.id, { action: "getViewsData", viewId: val });
      if (!response?.ok || !response.viewDetails) return;
      const d = response.viewDetails;
      viewNameInput.value = d.viewTitle || "";
      if (d.orderBy) {
        document.getElementById("viewSortColumn").value = d.orderBy.field || "";
        document.getElementById("viewSortDirection").value = d.orderBy.ascending !== false ? "asc" : "desc";
      } else {
        document.getElementById("viewSortColumn").value = "";
        document.getElementById("viewSortDirection").value = "asc";
      }
      document.getElementById("viewGroupBy").value = d.groupBy || "";
      viewFilters = (d.filters && d.filters.length) ? d.filters.map((f) => ({ column: f.field || "", op: f.op || "Eq", value: f.value || "" })) : (d.filter ? [{ column: d.filter.field || "", op: d.filter.op || "Eq", value: d.filter.value || "" }] : [{ column: "", op: "Eq", value: "" }]);
      const vf = (d.viewFields || []).map(String);
      const allInternal = (viewsData.fields || []).map((f) => f.internalName);
      viewColumnOrder = vf.length ? [...vf] : [...allInternal];
      allInternal.forEach((iname) => { if (viewColumnOrder.indexOf(iname) < 0) viewColumnOrder.push(iname); });
      renderViewFilters();
      renderViewColumnsList();
      const set = new Set(vf);
      document.querySelectorAll("#viewColumnsList input[type=checkbox]").forEach((cb) => {
        cb.checked = set.has(cb.dataset.internal || "");
      });
    } catch (_) {}
  })();
}

document.getElementById("viewSelect").addEventListener("change", onViewsViewSelectChange);

document.getElementById("viewAddFilterLink").addEventListener("click", (e) => {
  e.preventDefault();
  viewFilters.push({ column: "", op: "Eq", value: "" });
  renderViewFilters();
});

document.getElementById("btnSaveView").addEventListener("click", async () => {
  const viewSelect = document.getElementById("viewSelect");
  const viewNameInput = document.getElementById("viewNameInput");
  const viewsStatus = document.getElementById("viewsStatus");
  if (!viewsData || !viewsData.listId) {
    viewsStatus.textContent = "Load the Views tab first.";
    viewsStatus.className = "error";
    viewsStatus.style.display = "block";
    return;
  }
  const val = viewSelect && viewSelect.value;
  const isNew = val === "__new__" || !val;
  const viewTitle = (viewNameInput && viewNameInput.value) ? String(viewNameInput.value).trim() : "";
  if (isNew && !viewTitle) {
    viewsStatus.textContent = "Enter a name for the new view.";
    viewsStatus.className = "error";
    viewsStatus.style.display = "block";
    return;
  }
  const checkedSet = new Set(Array.from(document.querySelectorAll("#viewColumnsList input[type=checkbox]:checked")).map((c) => c.dataset.internal).filter(Boolean));
  const viewFields = viewColumnOrder.filter((iname) => checkedSet.has(iname));
  if (viewFields.length === 0) {
    viewsStatus.textContent = "Select at least one column.";
    viewsStatus.className = "error";
    viewsStatus.style.display = "block";
    return;
  }
  const sortCol = document.getElementById("viewSortColumn").value;
  const sortDir = document.getElementById("viewSortDirection").value;
  const filters = collectViewFiltersFromDom();
  const groupBy = document.getElementById("viewGroupBy").value;

  const payload = {
    listId: viewsData.listId,
    viewId: isNew ? "" : val,
    viewTitle: isNew ? viewTitle : (viewNameInput && viewNameInput.value) ? String(viewNameInput.value).trim() : viewTitle,
    viewFields,
    orderBy: sortCol ? { field: sortCol, ascending: sortDir !== "desc" } : null,
    filters,
    groupBy: groupBy || null
  };
  if (!payload.viewTitle) payload.viewTitle = "View";

  viewsStatus.textContent = "Saving…";
  viewsStatus.className = "info";
  viewsStatus.style.display = "block";

  try {
    let tab = (await chrome.tabs.query({ active: true, currentWindow: true }))[0];
    if (!tab?.id) throw new Error("No active tab");
    const response = await chrome.tabs.sendMessage(tab.id, { action: "saveView", payload });
    if (response?.ok) {
      const savedViewId = (response.detail && response.detail.viewId) || (isNew ? null : val);
      const savedViewTitle = (response.detail && response.detail.viewTitle) || viewTitle || "View";
      viewsStatus.textContent = "Save complete.";
      viewsStatus.className = "success";
      if (savedViewId && tab.url && tab.url.includes("sharepoint.com")) {
        try {
          const u = new URL(tab.url);
          u.searchParams.set("View", savedViewId);
          await chrome.tabs.update(tab.id, { url: u.toString() });
        } catch (_) {}
      }
      if (savedViewId && viewSelect) {
        let opt = viewSelect.querySelector("option[value=\"" + savedViewId.replace(/"/g, "\\\"") + "\"]");
        if (!opt) {
          opt = document.createElement("option");
          opt.value = savedViewId;
          opt.textContent = savedViewTitle;
          viewSelect.appendChild(opt);
        }
        viewSelect.value = savedViewId;
      }
    } else {
      viewsStatus.textContent = response?.message || "Save failed.";
      viewsStatus.className = "error";
    }
  } catch (e) {
    viewsStatus.textContent = "Error: " + (e.message || "Could not save.");
    viewsStatus.className = "error";
  }
});

async function getPageSize() {
  const st = await chrome.storage.local.get("pageSize");
  const v = parseInt(st.pageSize, 10);
  return (v === 500 || v === 1000 || v === 5000 || v === 10000 || v === 20000) ? v : DEFAULT_PAGE_SIZE;
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
      const exportFormat = (formatSelect && formatSelect.value) ? formatSelect.value : "csv";
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
  const showPanel = value === "exportCSV" || value === "folderCount" || value === "pathLengths" || value === "permissions" || value === "permissionsMatrix";
  reportOptions.style.display = showPanel ? "block" : "none";
  const matrixWholeSiteRow = document.getElementById("matrixWholeSiteRow");
  if (matrixWholeSiteRow) matrixWholeSiteRow.style.display = value === "permissionsMatrix" ? "" : "none";
  const exportOptionsRow = document.getElementById("exportOptionsRow");
  if (exportOptionsRow) exportOptionsRow.style.display = value === "exportCSV" ? "" : "none";
}
if (reportSelect) reportSelect.addEventListener("change", toggleReportOptions);

btnExport.addEventListener("click", () => {
  if (reportSelect && (reportSelect.value === "exportCSV" || reportSelect.value === "folderCount" || reportSelect.value === "pathLengths" || reportSelect.value === "permissions" || reportSelect.value === "permissionsMatrix")) runExport();
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
