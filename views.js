(function () {
  const params = new URLSearchParams(window.location.search);
  const tabId = params.get("tabId") ? parseInt(params.get("tabId"), 10) : null;

  // Sync dark/light mode from extension storage (same as popup)
  if (typeof chrome !== "undefined" && chrome.storage && chrome.storage.local) {
    chrome.storage.local.get("darkMode", function (st) {
      document.documentElement.classList.toggle("dark-mode", !!st.darkMode);
    });
    document.getElementById("darkModeToggle")?.addEventListener("click", function () {
      const isDark = document.documentElement.classList.toggle("dark-mode");
      chrome.storage.local.set({ darkMode: isDark });
    });
  }

  let sitePath = "";
  let listId = "";
  let listTitle = "";
  let views = [];
  let fields = [];
  let selectedViewId = null;
  let viewDetails = null;
  let columnOrder = [];
  let sortLevels = [];
  let filters = [];
  let groupByColumn = "";
  let groupExpand = true;

  const FILTER_OPS = [
    { value: "Eq", label: "equals" },
    { value: "Neq", label: "not equals" },
    { value: "Gt", label: "greater than" },
    { value: "Geq", label: "greater or equal" },
    { value: "Lt", label: "less than" },
    { value: "Leq", label: "less or equal" },
    { value: "Contains", label: "contains" },
    { value: "BeginsWith", label: "begins with" }
  ];

  function escapeHtml(s) {
    const div = document.createElement("div");
    div.textContent = s == null ? "" : s;
    return div.innerHTML;
  }
  function escapeAttr(s) {
    return escapeHtml(s == null ? "" : s).replace(/"/g, "&quot;");
  }

  function sendToTab(message) {
    return new Promise((resolve, reject) => {
      if (!tabId || !chrome || !chrome.tabs) {
        reject(new Error("No SharePoint tab. Open a list/library page and open View Manager from the popup."));
        return;
      }
      chrome.tabs.sendMessage(tabId, message, (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message || "Tab not ready. Try refreshing the SharePoint page."));
          return;
        }
        resolve(response);
      });
    });
  }

  function rest(method, path, body) {
    return sendToTab({ action: "rest", method: method || "GET", path: path, body: body });
  }

  function showContextStatus(msg, isError) {
    const el = document.getElementById("contextStatus");
    const lines = document.getElementById("contextLines");
    if (msg) {
      el.textContent = msg;
      el.className = "status " + (isError ? "error" : "loading");
      el.style.display = "block";
      if (lines) lines.classList.remove("visible");
    } else {
      el.style.display = "none";
      if (lines) lines.classList.add("visible");
    }
  }

  function showContext(webAbsoluteUrl, listUrl, listIdVal, listTitleVal) {
    listId = listIdVal || "";
    listTitle = listTitleVal || "";
    if (webAbsoluteUrl) {
      try {
        const u = new URL(webAbsoluteUrl);
        sitePath = u.pathname.replace(/\/$/, "") || "/";
      } catch (_) {
        sitePath = "/";
      }
    }
    const nameEl = document.getElementById("contextListName");
    const idEl = document.getElementById("contextListIdDisplay");
    const linesEl = document.getElementById("contextLines");
    if (nameEl) nameEl.textContent = listTitle || "—";
    if (idEl) idEl.textContent = listId || "—";
    if (linesEl) linesEl.classList.add("visible");
    showContextStatus(null);
    updateUrlWithList(listTitle);
  }

  function updateUrlWithList(listTitleVal) {
    if (typeof history === "undefined" || !history.replaceState) return;
    const slug = (listTitleVal || "")
      .trim()
      .replace(/\s+/g, "-")
      .replace(/[^a-zA-Z0-9_\-]/g, "")
      .replace(/-+/g, "-")
      .replace(/^-|-$/g, "")
      .slice(0, 80) || "list";
    const params = new URLSearchParams(window.location.search);
    if (tabId != null) params.set("tabId", String(tabId));
    params.set("list", slug);
    const url = window.location.pathname + "?" + params.toString();
    history.replaceState(null, "", url);
  }

  function copyListIdToClipboard() {
    if (!listId) return;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(listId).then(function () {
        const btn = document.getElementById("btnCopyListId");
        if (btn) { btn.textContent = "Copied!"; setTimeout(function () { btn.textContent = "Copy"; }, 1500); }
      }).catch(function () { fallbackCopyListId(); });
    } else {
      fallbackCopyListId();
    }
  }

  function fallbackCopyListId() {
    const el = document.getElementById("contextListIdDisplay");
    if (!el || !listId) return;
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(el);
    if (sel) { sel.removeAllRanges(); sel.addRange(range); }
    try {
      document.execCommand("copy");
      const btn = document.getElementById("btnCopyListId");
      if (btn) { btn.textContent = "Copied!"; setTimeout(function () { btn.textContent = "Copy"; }, 1500); }
    } catch (_) {}
    if (sel) sel.removeAllRanges();
  }

  async function loadContext() {
    if (!tabId) {
      document.getElementById("noListCard").classList.remove("hide");
      return false;
    }
    showContextStatus("Loading context…", false);
    try {
      const res = await sendToTab({ action: "getPageContext" });
      if (!res || !res.ok) {
        showContextStatus(res && res.error ? res.error : "Could not get page context.", true);
        document.getElementById("noListCard").classList.remove("hide");
        return false;
      }
      const webUrl = (res.webAbsoluteUrl || res.siteAbsoluteUrl || "").replace(/\/$/, "");
      const listUrl = (res.listUrl || "").trim();
      if (!webUrl) {
        showContextStatus("Not on a SharePoint site.", true);
        document.getElementById("noListCard").classList.remove("hide");
        return false;
      }
      try {
        const u = new URL(res.webAbsoluteUrl || res.siteAbsoluteUrl || "https://x/");
        sitePath = u.pathname.replace(/\/$/, "") || "/";
      } catch (_) {
        sitePath = "/";
      }
      if (!listUrl) {
        showContext(webUrl, "", "", "");
        showContextStatus("Not on a list or library. Open a list view and try again.", true);
        document.getElementById("noListCard").classList.remove("hide");
        return false;
      }
      const apiPath = sitePath + "/_api/web/GetList(@a)/Id?@a='" + encodeURIComponent(listUrl.replace(/'/g, "''")) + "'";
      const idRes = await rest("GET", apiPath);
      if (!idRes || !idRes.ok) {
        showContext(webUrl, listUrl, "", "");
        showContextStatus(idRes && idRes.error ? String(idRes.error) : "Could not resolve list ID.", true);
        return false;
      }
      let lid = (idRes.data && (idRes.data.value !== undefined ? idRes.data.value : (idRes.data.d && idRes.data.d.Id) || idRes.data.Id)) || "";
      if (lid && typeof lid === "string") lid = lid.replace(/[{}]/g, "").trim();
      const titlePath = sitePath + "/_api/web/lists(guid'" + lid + "')?$select=Title";
      const titleRes = await rest("GET", titlePath);
      let title = "";
      if (titleRes && titleRes.ok && titleRes.data) {
        const t = titleRes.data.Title || titleRes.data.title;
        if (t != null) title = String(t);
      }
      showContext(webUrl, listUrl, lid, title);
      document.getElementById("noListCard").classList.add("hide");
      return true;
    } catch (e) {
      showContextStatus(e && e.message ? e.message : "Failed to load context.", true);
      document.getElementById("noListCard").classList.remove("hide");
      return false;
    }
  }

  async function loadViewsAndFields() {
    if (!listId) return;
    try {
      const res = await sendToTab({ action: "getViewsData" });
      if (!res || res.error) {
        showSaveStatus(res && res.error ? res.error : "Failed to load views and columns.", true);
        return;
      }
      views = res.views || [];
      fields = res.fields || [];
      renderViewSelect();
      renderColumnList();
      renderSortLevels();
      renderFilterConditions();
      renderGroupBy();
      document.getElementById("editorCard").classList.remove("hide");
      selectDefaultView();
    } catch (e) {
      showSaveStatus(e && e.message ? e.message : "Failed to load data.", true);
    }
  }

  async function loadViewDetails(viewId) {
    if (!viewId) return;
    try {
      const res = await sendToTab({ action: "getViewsData", viewId: viewId });
      if (!res || res.error || !res.viewDetails) {
        showSaveStatus(res && res.error ? res.error : "Failed to load view details.", true);
        return;
      }
      viewDetails = res.viewDetails;
      columnOrder = (viewDetails.viewFields || []).slice();
      sortLevels = viewDetails.orderBy ? [{ field: viewDetails.orderBy.field, ascending: viewDetails.orderBy.ascending }] : [];
      filters = (viewDetails.filters || []).map(function (f) {
        return { field: f.field, op: f.op, value: f.value || "" };
      });
      groupByColumn = viewDetails.groupBy || "";
      groupExpand = true;
      document.getElementById("viewName").value = viewDetails.viewTitle || "";
      const viewMetaPath = sitePath + "/_api/web/lists(guid'" + listId.replace(/'/g, "''") + "')/views(guid'" + viewId.replace(/'/g, "''") + "')?$select=RowLimit,Scope";
      const metaRes = await rest("GET", viewMetaPath);
      if (metaRes && metaRes.ok && metaRes.data) {
        const rl = (metaRes.data.RowLimit != null ? metaRes.data.RowLimit : metaRes.data.rowLimit);
        if (rl != null) {
          const rlStr = String(rl);
          const rlEl = document.getElementById("viewRowLimit");
          if (rlEl.querySelector('option[value="' + rlStr + '"]')) rlEl.value = rlStr;
        }
        const sc = metaRes.data.Scope != null ? metaRes.data.Scope : metaRes.data.scope;
        const scopeEl = document.getElementById("viewScope");
        if (scopeEl) {
          if (sc === 1 || sc === "1") scopeEl.value = "1";
          else if (sc === 2 || sc === "2") scopeEl.value = "2";
          else scopeEl.value = "2";
        }
      }
      const rlEl = document.getElementById("viewRowLimit");
      if (!rlEl.value) rlEl.value = "100";
      renderColumnList();
      renderSortLevels();
      renderFilterConditions();
      renderGroupBy();
      updateSetDefaultVisibility();
    } catch (e) {
      showSaveStatus(e && e.message ? e.message : "Failed to load view.", true);
    }
  }

  function setNewViewDefaults() {
    viewDetails = null;
    columnOrder = [];
    sortLevels = [];
    filters = [];
    groupByColumn = "";
    groupExpand = true;
    document.getElementById("viewName").value = "";
    document.getElementById("viewPersonal").checked = false;
    document.getElementById("viewRowLimit").value = "100";
    document.getElementById("viewScope").value = "2";
    document.getElementById("btnDelete").classList.add("hide");
    document.getElementById("btnSetDefault").classList.add("hide");
    document.getElementById("btnDuplicate").classList.add("hide");
    renderColumnList();
    renderSortLevels();
    renderFilterConditions();
    renderGroupBy();
  }

  function renderViewSelect() {
    const sel = document.getElementById("viewSelect");
    const prevVal = sel.value;
    sel.innerHTML = "<option value=\"\">— Create new view —</option>";
    views.forEach(function (v) {
      const opt = document.createElement("option");
      opt.value = v.id;
      opt.textContent = v.title + (v.defaultView ? " (default)" : "");
      sel.appendChild(opt);
    });
    sel.value = selectedViewId || prevVal || "";
  }

  function selectDefaultView() {
    const defaultView = views.find(function (v) { return v.defaultView; }) || views[0];
    if (defaultView) {
      selectedViewId = defaultView.id;
      document.getElementById("viewSelect").value = defaultView.id;
      document.getElementById("btnDelete").classList.remove("hide");
      document.getElementById("btnDuplicate").classList.remove("hide");
      loadViewDetails(defaultView.id);
    } else {
      setNewViewDefaults();
    }
  }

  function onViewSelectChange() {
    const val = document.getElementById("viewSelect").value;
    selectedViewId = val || null;
    if (!val) {
      setNewViewDefaults();
      updateSetDefaultVisibility();
      return;
    }
    document.getElementById("btnDelete").classList.remove("hide");
    document.getElementById("btnDuplicate").classList.remove("hide");
    updateSetDefaultVisibility();
    loadViewDetails(val);
  }

  function updateSetDefaultVisibility() {
    const btn = document.getElementById("btnSetDefault");
    if (!btn) return;
    if (!selectedViewId) {
      btn.classList.add("hide");
      return;
    }
    const v = views.find(function (x) { return x.id === selectedViewId; });
    if (v && v.defaultView) {
      btn.classList.add("hide");
    } else {
      btn.classList.remove("hide");
    }
  }

  function renderColumnList() {
    const list = document.getElementById("colList");
    const search = (document.getElementById("colSearch").value || "").toLowerCase().trim();
    const ordered = columnOrder.slice();
    const rest = fields.filter(function (f) {
      return ordered.indexOf(f.internalName) < 0 &&
        (!search || (f.title || "").toLowerCase().indexOf(search) >= 0 || (f.internalName || "").toLowerCase().indexOf(search) >= 0);
    });
    rest.sort(function (a, b) { return (a.title || "").localeCompare(b.title || ""); });
    const toShow = ordered.filter(function (name) {
      const f = fields.find(function (x) { return x.internalName === name; });
      return f && (!search || (f.title || "").toLowerCase().indexOf(search) >= 0 || (f.internalName || "").toLowerCase().indexOf(search) >= 0);
    }).map(function (name) {
      return fields.find(function (f) { return f.internalName === name; });
    }).filter(Boolean);
    const restFiltered = search ? rest : rest;

    list.innerHTML = "";
    let draggedName = null;
    function addRow(field, isInView) {
      const div = document.createElement("div");
      div.className = "col-item";
      div.dataset.internalName = field.internalName;
      div.draggable = true;
      div.innerHTML =
        "<span class=\"drag-handle\" aria-hidden=\"true\">⋮⋮</span>" +
        "<span class=\"col-check\"><input type=\"checkbox\" class=\"col-checkbox\" " + (isInView ? "checked" : "") + "></span>" +
        "<span class=\"col-label\">" + escapeHtml(field.title || field.internalName) + "</span>" +
        "<span class=\"col-internal\">" + escapeHtml(field.internalName) + "</span>";
      const cb = div.querySelector(".col-checkbox");
      cb.addEventListener("change", function () {
        if (cb.checked) {
          if (columnOrder.indexOf(field.internalName) < 0) columnOrder.push(field.internalName);
        } else {
          columnOrder = columnOrder.filter(function (n) { return n !== field.internalName; });
        }
        renderColumnList();
      });
      div.addEventListener("dragstart", function (e) {
        draggedName = field.internalName;
        div.classList.add("dragging");
        e.dataTransfer.setData("text/plain", field.internalName);
        e.dataTransfer.effectAllowed = "move";
      });
      div.addEventListener("dragend", function () {
        div.classList.remove("dragging");
        draggedName = null;
      });
      div.addEventListener("dragover", function (e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
      });
      div.addEventListener("drop", function (e) {
        e.preventDefault();
        const name = e.dataTransfer.getData("text/plain");
        if (!name || name === field.internalName || name !== draggedName) return;
        const targetInOrder = columnOrder.indexOf(field.internalName) >= 0;
        let arr = columnOrder.filter(function (n) { return n !== name; });
        if (targetInOrder) {
          const insertIdx = arr.indexOf(field.internalName);
          arr.splice(insertIdx >= 0 ? insertIdx : arr.length, 0, name);
        } else {
          arr.push(name);
        }
        columnOrder = arr;
        renderColumnList();
      });
      list.appendChild(div);
    }
    toShow.forEach(function (f) { addRow(f, true); });
    restFiltered.forEach(function (f) { addRow(f, false); });
  }

  document.getElementById("colSearch").addEventListener("input", function () { renderColumnList(); });
  document.getElementById("btnColAddAll").addEventListener("click", function () {
    columnOrder = fields.map(function (f) { return f.internalName; });
    renderColumnList();
  });
  document.getElementById("btnColRemoveAll").addEventListener("click", function () {
    columnOrder = [];
    renderColumnList();
  });

  function renderSortLevels() {
    const container = document.getElementById("sortContainer");
    container.innerHTML = "";
    sortLevels.forEach(function (s, i) {
      const row = document.createElement("div");
      row.className = "sort-row";
      const fieldSel = document.createElement("select");
      fieldSel.className = "sort-field";
      fields.forEach(function (f) {
        const opt = document.createElement("option");
        opt.value = f.internalName;
        opt.textContent = f.title || f.internalName;
        if (f.internalName === s.field) opt.selected = true;
        fieldSel.appendChild(opt);
      });
      const dirSel = document.createElement("select");
      dirSel.className = "sort-dir";
      dirSel.innerHTML = "<option value=\"true\">Ascending</option><option value=\"false\">Descending</option>";
      dirSel.value = s.ascending ? "true" : "false";
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "link";
      btn.textContent = "Remove";
      btn.addEventListener("click", function () {
        sortLevels.splice(i, 1);
        renderSortLevels();
      });
      row.appendChild(fieldSel);
      row.appendChild(dirSel);
      row.appendChild(btn);
      container.appendChild(row);
      fieldSel.addEventListener("change", function () { sortLevels[i].field = fieldSel.value; });
      dirSel.addEventListener("change", function () { sortLevels[i].ascending = dirSel.value === "true"; });
    });
  }

  document.getElementById("btnAddSort").addEventListener("click", function () {
    sortLevels.push({ field: fields[0] ? fields[0].internalName : "", ascending: true });
    renderSortLevels();
  });

  function renderFilterConditions() {
    const container = document.getElementById("filterContainer");
    container.innerHTML = "";
    filters.forEach(function (f, i) {
      const row = document.createElement("div");
      row.className = "filter-row";
      const fieldSel = document.createElement("select");
      fieldSel.className = "filter-field";
      fields.forEach(function (fd) {
        const opt = document.createElement("option");
        opt.value = fd.internalName;
        opt.textContent = fd.title || fd.internalName;
        if (fd.internalName === f.field) opt.selected = true;
        fieldSel.appendChild(opt);
      });
      const opSel = document.createElement("select");
      opSel.className = "filter-op";
      FILTER_OPS.forEach(function (o) {
        const opt = document.createElement("option");
        opt.value = o.value;
        opt.textContent = o.label;
        if (o.value === f.op) opt.selected = true;
        opSel.appendChild(opt);
      });
      const valInput = document.createElement("input");
      valInput.type = "text";
      valInput.placeholder = "Value";
      valInput.value = f.value || "";
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "link";
      btn.textContent = "Remove";
      btn.addEventListener("click", function () {
        filters.splice(i, 1);
        renderFilterConditions();
      });
      row.appendChild(fieldSel);
      row.appendChild(opSel);
      row.appendChild(valInput);
      row.appendChild(btn);
      container.appendChild(row);
      fieldSel.addEventListener("change", function () { filters[i].field = fieldSel.value; });
      opSel.addEventListener("change", function () { filters[i].op = opSel.value; });
      valInput.addEventListener("input", function () { filters[i].value = valInput.value; });
    });
  }

  document.getElementById("btnAddFilter").addEventListener("click", function () {
    filters.push({
      field: fields[0] ? fields[0].internalName : "",
      op: "Eq",
      value: ""
    });
    renderFilterConditions();
  });

  function renderGroupBy() {
    const sel = document.getElementById("groupByColumn");
    sel.innerHTML = "<option value=\"\">— None —</option>";
    fields.forEach(function (f) {
      const opt = document.createElement("option");
      opt.value = f.internalName;
      opt.textContent = f.title || f.internalName;
      if (f.internalName === groupByColumn) opt.selected = true;
      sel.appendChild(opt);
    });
    document.getElementById("groupExpand").checked = groupExpand;
  }

  document.getElementById("groupByColumn").addEventListener("change", function () {
    groupByColumn = document.getElementById("groupByColumn").value || "";
  });
  document.getElementById("groupExpand").addEventListener("change", function () {
    groupExpand = document.getElementById("groupExpand").checked;
  });

  function getValueTypeForField(fieldInternalName) {
    const fd = fields.find(function (f) { return f.internalName === fieldInternalName; });
    const t = (fd && fd.typeAsString) ? String(fd.typeAsString) : "";
    if (/Integer|Counter|Boolean/i.test(t)) return "Integer";
    if (/Number|Currency|Decimal/i.test(t)) return "Number";
    if (/DateTime|Date/i.test(t)) return "DateTime";
    return "Text";
  }

  function buildViewQuery() {
    const parts = [];
    const hasFilter = filters.some(function (f) { return (f.value || "").trim() !== ""; });
    if (hasFilter) {
      const conds = filters.filter(function (f) { return (f.value || "").trim() !== ""; }).map(function (f) {
        const v = String(f.value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        const type = getValueTypeForField(f.field).replace(/"/g, "");
        return "<" + f.op + "><FieldRef Name=\"" + escapeAttr(f.field) + "\"/><Value Type=\"" + type + "\">" + v + "</Value></" + f.op + ">";
      });
      if (conds.length === 1) parts.push("<Where>" + conds[0] + "</Where>");
      else if (conds.length > 1) parts.push("<Where><And>" + conds.join("") + "</And></Where>");
    }
    if (sortLevels.length > 0) {
      const refs = sortLevels.filter(function (s) { return s.field; }).map(function (s) {
        return "<FieldRef Name=\"" + escapeAttr(s.field) + "\" Ascending=\"" + (s.ascending ? "True" : "False") + "\"/>";
      });
      if (refs.length > 0) parts.push("<OrderBy>" + refs.join("") + "</OrderBy>");
    }
    if (groupByColumn) {
      const collapse = groupExpand ? "FALSE" : "TRUE";
      parts.push("<GroupBy Collapse=\"" + collapse + "\"><FieldRef Name=\"" + escapeAttr(groupByColumn) + "\"/></GroupBy>");
    }
    return parts.join("");
  }

  function showSaveStatus(msg, isError) {
    const el = document.getElementById("saveStatus");
    el.textContent = msg || "";
    el.className = "status " + (isError ? "error" : "success");
    el.style.display = msg ? "block" : "none";
  }

  async function applyViewFields(viewId) {
    const vfBase = sitePath + "/_api/web/lists(guid'" + listId.replace(/'/g, "''") + "')/views(guid'" + viewId.replace(/'/g, "''") + "')/ViewFields";
    const rm = await rest("POST", vfBase + "/RemoveAllViewFields");
    if (!rm.ok && rm.status !== 404 && rm.status !== 501) throw new Error(rm.error || "RemoveAllViewFields failed");
    for (let i = 0; i < columnOrder.length; i++) {
      const name = String(columnOrder[i]).replace(/'/g, "''");
      const r = await rest("POST", vfBase + "/addviewfield('" + name + "')");
      if (!r.ok) throw new Error(r.error || "Add field failed");
    }
  }

  async function saveView() {
    const name = (document.getElementById("viewName").value || "").trim();
    if (!name) {
      showSaveStatus("Enter a view name.", true);
      return;
    }
    if (columnOrder.length === 0) {
      showSaveStatus("Select at least one column.", true);
      return;
    }
    if (!listId) {
      showSaveStatus("No list context. Refresh context.", true);
      return;
    }

    const personal = document.getElementById("viewPersonal").checked;
    const rowLimit = parseInt(document.getElementById("viewRowLimit").value, 10) || 100;
    const scope = parseInt(document.getElementById("viewScope").value, 10) || 2;
    const viewQuery = buildViewQuery();

    document.getElementById("btnSave").disabled = true;
    showSaveStatus("Saving…", false);

    try {
      const lb = sitePath + "/_api/web/lists(guid'" + listId.replace(/'/g, "''") + "')";
      let viewId = selectedViewId;

      if (!viewId) {
        const createRes = await rest("POST", lb + "/views", {
          Title: name,
          PersonalView: personal,
          RowLimit: rowLimit,
          Scope: scope,
          ViewQuery: viewQuery
        });
        if (!createRes || !createRes.ok) throw new Error(createRes && createRes.error ? (typeof createRes.error === "string" ? createRes.error : JSON.stringify(createRes.error)) : "Create failed");
        const created = createRes.data;
        viewId = (created.Id || created.id || "").replace(/[{}]/g, "").trim();
      } else {
        const patchRes = await rest("PATCH", lb + "/views(guid'" + viewId.replace(/'/g, "''") + "')", {
          Title: name,
          PersonalView: personal,
          RowLimit: rowLimit,
          Scope: scope,
          ViewQuery: viewQuery
        });
        if (!patchRes || !patchRes.ok) throw new Error(patchRes && patchRes.error ? String(patchRes.error) : "Update failed");
      }

      await applyViewFields(viewId);
      showSaveStatus("View saved successfully.", false);
      document.getElementById("btnSave").disabled = false;
      selectedViewId = viewId;
      document.getElementById("viewSelect").value = viewId;
      document.getElementById("btnDelete").classList.remove("hide");
      loadContext().then(function () { loadViewsAndFields(); });
    } catch (e) {
      showSaveStatus(e && e.message ? e.message : "Save failed.", true);
      document.getElementById("btnSave").disabled = false;
    }
  }

  async function deleteView() {
    if (!selectedViewId) return;
    if (!confirm("Delete this view? This cannot be undone.")) return;
    try {
      const path = sitePath + "/_api/web/lists(guid'" + listId.replace(/'/g, "''") + "')/views(guid'" + selectedViewId.replace(/'/g, "''") + "')";
      const res = await rest("DELETE", path);
      if (!res || !res.ok) throw new Error(res && res.error ? String(res.error) : "Delete failed");
      showSaveStatus("View deleted.", false);
      selectedViewId = null;
      document.getElementById("viewSelect").value = "";
      setNewViewDefaults();
      loadContext().then(function () { loadViewsAndFields(); });
    } catch (e) {
      showSaveStatus(e && e.message ? e.message : "Delete failed.", true);
    }
  }

  function duplicateView() {
    if (!selectedViewId && !viewDetails) return;
    selectedViewId = null;
    document.getElementById("viewSelect").value = "";
    document.getElementById("viewName").value = (document.getElementById("viewName").value || "").trim() + " (copy)";
    document.getElementById("btnDelete").classList.add("hide");
    document.getElementById("btnSetDefault").classList.add("hide");
    document.getElementById("btnDuplicate").classList.add("hide");
  }

  async function setDefaultView() {
    if (!selectedViewId || !listId) return;
    const btn = document.getElementById("btnSetDefault");
    if (btn) btn.disabled = true;
    showSaveStatus("Setting default view…", false);
    try {
      const path = sitePath + "/_api/web/lists(guid'" + listId.replace(/'/g, "''") + "')/views(guid'" + selectedViewId.replace(/'/g, "''") + "')";
      const res = await rest("PATCH", path, { DefaultView: true });
      if (!res || !res.ok) throw new Error(res && res.error ? String(res.error) : "Set default failed");
      showSaveStatus("Default view updated.", false);
      await loadViewsAndFields();
    } catch (e) {
      showSaveStatus(e && e.message ? e.message : "Set default failed.", true);
    }
    if (btn) btn.disabled = false;
  }

  document.getElementById("backToTab").addEventListener("click", function (e) {
    e.preventDefault();
    if (tabId && chrome.tabs) {
      chrome.tabs.update(tabId, { active: true });
      chrome.tabs.get(tabId, function (tab) {
        if (tab && tab.windowId) chrome.windows.update(tab.windowId, { focused: true });
      });
    }
  });

  document.getElementById("btnCopyListId").addEventListener("click", copyListIdToClipboard);

  document.getElementById("btnSave").addEventListener("click", saveView);
  document.getElementById("btnSetDefault").addEventListener("click", setDefaultView);
  document.getElementById("btnDelete").addEventListener("click", deleteView);
  document.getElementById("btnDuplicate").addEventListener("click", duplicateView);
  document.getElementById("viewSelect").addEventListener("change", onViewSelectChange);

  (async function init() {
    const ok = await loadContext();
    if (ok) await loadViewsAndFields();
  })();
})();
