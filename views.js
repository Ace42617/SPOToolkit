(function () {
  const params = new URLSearchParams(window.location.search);
  const tabId = params.get("tabId") ? parseInt(params.get("tabId"), 10) : null;
  const forcedViewManagerListId = (params.get("listId") || "").replace(/[{}]/g, "").trim();
  const forcedViewManagerWebUrl = (params.get("webUrl") || "").replace(/\/$/, "");
  const forceCreateNewView = params.get("create") === "1";

  // Sync dark/light mode and optional branding from extension storage (same as popup)
  if (typeof chrome !== "undefined" && chrome.storage && chrome.storage.local) {
    chrome.storage.local.get(["darkMode", "toolkitTitle", "toolkitSubtitle"], function (st) {
      document.documentElement.classList.toggle("dark-mode", !!st.darkMode);
      const titleEl = document.getElementById("headerTitle");
      const subEl = document.getElementById("headerSub");
      if (titleEl && st.toolkitTitle) titleEl.textContent = st.toolkitTitle;
      if (subEl && st.toolkitSubtitle) subEl.textContent = st.toolkitSubtitle;
    });
    document.getElementById("darkModeToggle")?.addEventListener("click", function () {
      const isDark = document.documentElement.classList.toggle("dark-mode");
      chrome.storage.local.set({ darkMode: isDark });
    });
  }

  const headerIconBtn = document.getElementById("headerIconBtn");
  const headerCompassImg = document.getElementById("headerCompassIcon");
  if (headerCompassImg && typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.getURL) {
    headerCompassImg.src = chrome.runtime.getURL("icon.png");
  }
  function showFactToastInPage(factText, sourceUrl) {
    const TOAST_ID = "vm-fact-toast";
    const DURATION_MS = 45000;
    let existing = document.getElementById(TOAST_ID);
    if (existing) existing.remove();
    const wrap = document.createElement("div");
    wrap.id = TOAST_ID;
    wrap.innerHTML =
      "<div class=\"vm-fact-wrap\">" +
      "<button type=\"button\" class=\"vm-fact-close\" aria-label=\"Close\">×</button>" +
      "<p class=\"vm-fact-text\"></p>" +
      (sourceUrl ? "<a class=\"vm-fact-source\" href=\"#\" target=\"_blank\" rel=\"noopener\">Learn More</a>" : "") +
      "</div>";
    wrap.querySelector(".vm-fact-text").textContent = factText || "";
    let closeTimeout;
    const close = function () {
      if (closeTimeout) clearTimeout(closeTimeout);
      const el = document.getElementById(TOAST_ID);
      if (el && el.parentNode) el.parentNode.removeChild(el);
    };
    wrap.querySelector(".vm-fact-close").addEventListener("click", close);
    if (sourceUrl) {
      const link = wrap.querySelector(".vm-fact-source");
      link.href = sourceUrl;
      link.addEventListener("click", function (e) { e.preventDefault(); window.open(sourceUrl, "_blank", "noopener"); });
    }
    document.body.appendChild(wrap);
    closeTimeout = setTimeout(close, DURATION_MS);
  }

  if (headerIconBtn) {
    headerIconBtn.addEventListener("click", function () {
      headerIconBtn.classList.remove("spin");
      headerIconBtn.offsetHeight;
      headerIconBtn.classList.add("spin");
      setTimeout(function () { headerIconBtn.classList.remove("spin"); }, 600);
      const facts = typeof window.SHAREPOINT_FACTS !== "undefined" ? window.SHAREPOINT_FACTS : [];
      if (!facts.length) return;
      const fact = facts[Math.floor(Math.random() * facts.length)];
      const factText = typeof fact === "string" ? fact : (fact && fact.text ? fact.text : "");
      const sourceUrl = typeof fact === "object" && fact && fact.sourceUrl && String(fact.sourceUrl).trim() ? fact.sourceUrl : "";
      if (typeof fact === "object" && fact && typeof fact.openUrl === "string") {
        if (chrome && chrome.tabs) chrome.tabs.create({ url: fact.openUrl });
      }
      showFactToastInPage(factText, sourceUrl);
      if (tabId && typeof chrome !== "undefined" && chrome.tabs) {
        chrome.tabs.sendMessage(tabId, { action: "showFact", fact: factText, sourceUrl: sourceUrl }, function () {
          if (chrome.runtime.lastError) { /* tab may be closed or content script not ready */ }
        });
      }
    });
  }

  let sitePath = "";
  /** Full web URL for REST (passed into getViewsData when list is opened off-list). */
  let contextWebAbsoluteUrl = "";
  let listId = "";
  let listTitle = "";
  let contentTypes = [];
  let views = [];
  let fields = [];
  let selectedViewId = null;
  let viewDetails = null;
  let columnOrder = [];
  let inViewSet = new Set();
  let sortLevels = [];
  let filters = [];
  let groupByColumn = "";
  let groupExpand = true;
  let groupByAscending = true;
  let groupByExtraLevels = [];
  let groupLimit = null;
  let loadedGroupByPrimary = "";

  const LIST_TYPE_LABELS = {
    100: "List",
    101: "Document library",
    106: "Events list",
    107: "Links list",
    109: "Picture library",
    110: "Survey",
    119: "Page library",
    171: "Issue tracking"
  };
  function listBaseTemplateLabel(bt) {
    if (bt == null || bt === "") return "";
    const n = typeof bt === "number" ? bt : parseInt(String(bt), 10);
    if (isNaN(n)) return "";
    return LIST_TYPE_LABELS[n] || "List";
  }

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

  function filterTypeaheadApi() {
    return typeof window !== "undefined" && window.SPFilterTypeahead ? window.SPFilterTypeahead : null;
  }

  function filterTypeaheadKindsForField(internalName) {
    const ta = filterTypeaheadApi();
    if (!ta) return [];
    const fd = fields.find(function (x) { return x.internalName === internalName; });
    return ta.kindsFromTypeAsString(fd && fd.typeAsString);
  }

  /** Text before caret: open `[…` token or last whitespace-delimited word. */
  function filterTypeaheadQuery(input) {
    const v = input.value;
    const start = input.selectionStart != null ? input.selectionStart : v.length;
    const before = v.slice(0, start);
    const bracketStart = before.lastIndexOf("[");
    if (bracketStart >= 0 && before.indexOf("]", bracketStart) < 0) {
      return before.slice(bracketStart);
    }
    const lastSeg = before.match(/[^\s]*$/);
    return lastSeg ? lastSeg[0] : "";
  }

  function filterTypeaheadInsertAtCaret(input, token, rowIndex) {
    const v = input.value;
    const start = input.selectionStart != null ? input.selectionStart : v.length;
    const before = v.slice(0, start);
    const after = v.slice(start);
    const bStart = before.lastIndexOf("[");
    const partialBracket = bStart >= 0 && before.indexOf("]", bStart) < 0;
    let nv;
    let caret;
    if (partialBracket) {
      nv = v.slice(0, bStart) + token + after;
      caret = bStart + token.length;
    } else if (!v.trim()) {
      nv = token;
      caret = token.length;
    } else {
      const seg = before.match(/[^\s]*$/);
      const segLen = seg ? seg[0].length : 0;
      const replaceStart = before.length - segLen;
      const prefix = v.slice(0, replaceStart);
      const needsSpace = prefix.length > 0 && !/\s$/.test(prefix);
      nv = prefix + (needsSpace ? " " : "") + token + after;
      caret = prefix.length + (needsSpace ? 1 : 0) + token.length;
    }
    input.value = nv;
    input.setSelectionRange(caret, caret);
    filters[rowIndex].value = nv;
  }

  function attachFilterValueTypeahead(input, rowIndex) {
    const wrap = input.parentElement;
    const ta = filterTypeaheadApi();
    if (!wrap || !ta) return;

    const listEl = document.createElement("ul");
    listEl.className = "filter-typeahead-list hide";
    listEl.setAttribute("role", "listbox");
    wrap.appendChild(listEl);

    let activeIdx = -1;

    function suggestionsForCurrentField() {
      const kinds = filterTypeaheadKindsForField(filters[rowIndex].field);
      const q = filterTypeaheadQuery(input);
      const search = q === "" || q === "[" ? "" : q;
      return ta.suggest(kinds, search);
    }

    function renderList() {
      const items = suggestionsForCurrentField();
      listEl.innerHTML = "";
      if (items.length === 0) {
        listEl.classList.add("hide");
        return;
      }
      items.forEach(function (s) {
        const li = document.createElement("li");
        li.className = "filter-typeahead-item";
        li.setAttribute("role", "option");
        li.dataset.token = s.token;
        const tokSpan = document.createElement("span");
        tokSpan.className = "filter-typeahead-token";
        tokSpan.textContent = s.token;
        li.appendChild(tokSpan);
        if (s.hint) {
          const hintSpan = document.createElement("span");
          hintSpan.className = "filter-typeahead-hint";
          hintSpan.textContent = s.hint;
          li.appendChild(hintSpan);
        }
        li.addEventListener("mousedown", function (e) {
          e.preventDefault();
          filterTypeaheadInsertAtCaret(input, s.token, rowIndex);
          listEl.classList.add("hide");
          input.focus();
        });
        listEl.appendChild(li);
      });
      listEl.classList.remove("hide");
      activeIdx = -1;
    }

    function updateActive(lis) {
      lis.forEach(function (li, j) { li.classList.toggle("active", j === activeIdx); });
      if (activeIdx >= 0 && lis[activeIdx]) lis[activeIdx].scrollIntoView({ block: "nearest" });
    }

    input.addEventListener("input", function () {
      filters[rowIndex].value = input.value;
      renderList();
    });
    input.addEventListener("focus", function () { renderList(); });
    input.addEventListener("blur", function () {
      setTimeout(function () {
        listEl.classList.add("hide");
        activeIdx = -1;
      }, 120);
    });
    input.addEventListener("keydown", function (e) {
      if (listEl.classList.contains("hide")) return;
      const lis = listEl.querySelectorAll(".filter-typeahead-item");
      if (!lis.length) return;
      if (e.key === "ArrowDown") {
        e.preventDefault();
        activeIdx = Math.min(activeIdx + 1, lis.length - 1);
        updateActive(lis);
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        activeIdx = Math.max(activeIdx - 1, 0);
        updateActive(lis);
      } else if (e.key === "Enter") {
        e.preventDefault();
        const idx = activeIdx >= 0 ? activeIdx : 0;
        filterTypeaheadInsertAtCaret(input, lis[idx].dataset.token, rowIndex);
        listEl.classList.add("hide");
        input.focus();
      } else if (e.key === "Escape") {
        e.preventDefault();
        listEl.classList.add("hide");
        activeIdx = -1;
      }
    });
  }

  function filterValuePlaceholderForField(internalName) {
    const kinds = filterTypeaheadKindsForField(internalName);
    if (kinds.indexOf("date") >= 0 && kinds.indexOf("person") >= 0) {
      return "Date tokens or [Me]…";
    }
    if (kinds.indexOf("date") >= 0) {
      return "Date or type [ for [Today]…";
    }
    if (kinds.indexOf("person") >= 0) {
      return "User id or [Me]…";
    }
    return "Value";
  }

  /** Normalize: first row has no join; others default to And. */
  function normalizeFilterJoins(arr) {
    arr.forEach(function (f, idx) {
      if (idx === 0) {
        delete f.join;
      } else {
        const j = String(f.join || "").toLowerCase();
        f.join = j === "or" ? "Or" : "And";
      }
    });
  }

  function matchBalancedOuter(s, tag) {
    const openRe = new RegExp("^<" + tag + "\\s*>", "i");
    const om = s.match(openRe);
    if (!om) return null;
    let i = om[0].length;
    let depth = 1;
    while (i < s.length && depth > 0) {
      const rest = s.slice(i);
      const no = rest.match(new RegExp("^<" + tag + "\\s*>", "i"));
      const nc = rest.match(new RegExp("^</" + tag + "\\s*>", "i"));
      if (nc && (!no || nc.index < no.index)) {
        depth--;
        i += nc[0].length;
        if (depth === 0) return s.slice(0, i);
        continue;
      }
      if (no) {
        depth++;
        i += no[0].length;
        continue;
      }
      i++;
    }
    return null;
  }

  function takeFirstFilterSegment(s) {
    s = s.trim();
    const leafTags = ["Eq", "Neq", "Gt", "Geq", "Lt", "Leq", "Contains", "BeginsWith"];
    for (let t = 0; t < leafTags.length; t++) {
      const tag = leafTags[t];
      const re = new RegExp("^<" + tag + "\\s*>[\\s\\S]*?</" + tag + "\\s*>", "i");
      const m = s.match(re);
      if (m) return [m[0], s.slice(m[0].length).trim()];
    }
    const tries = ["And", "Or"];
    for (let t = 0; t < tries.length; t++) {
      const block = matchBalancedOuter(s, tries[t]);
      if (block) return [block, s.slice(block.length).trim()];
    }
    return null;
  }

  function takeAllFilterSegments(body) {
    const segs = [];
    let rest = body.trim();
    while (rest.length) {
      const seg = takeFirstFilterSegment(rest);
      if (!seg) return null;
      segs.push(seg[0].trim());
      rest = seg[1].trim();
    }
    return segs.length ? segs : null;
  }

  function parseLeafCondString(s) {
    const norm = s.replace(/\s+/g, " ").trim();
    const re = /^<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)\s*>\s*<FieldRef\s+Name="([^"]+)"\s*\/>\s*<Value\s+Type="([^"]*)">([^<]*)<\/Value>\s*<\/\1\s*>$/i;
    const m = norm.match(re);
    if (!m) return null;
    const v = (m[4] || "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">");
    return { type: "leaf", field: m[2], op: m[1], value: v };
  }

  function parseWhereExpr(inner) {
    inner = inner.trim();
    const leaf = parseLeafCondString(inner);
    if (leaf) return leaf;
    const tries = ["And", "Or"];
    for (let ti = 0; ti < tries.length; ti++) {
      const tag = tries[ti];
      const full = matchBalancedOuter(inner, tag);
      if (!full || full.length !== inner.length) continue;
      const openM = inner.match(new RegExp("^<" + tag + "\\s*>", "i"));
      const body = inner.slice(openM[0].length, inner.length - ("</" + tag + ">").length).trim();
      const segs = takeAllFilterSegments(body);
      if (!segs || segs.length === 0) return null;
      const nodes = segs.map(function (seg) { return parseWhereExpr(seg); });
      if (nodes.some(function (n) { return !n; })) return null;
      let acc = nodes[0];
      for (let k = 1; k < nodes.length; k++) {
        acc = { type: "bin", op: tag, left: acc, right: nodes[k] };
      }
      return acc;
    }
    return null;
  }

  function binTreeToFilterRows(node) {
    if (node.type === "leaf") {
      return [{ field: node.field, op: node.op, value: node.value }];
    }
    const L = binTreeToFilterRows(node.left);
    const R = binTreeToFilterRows(node.right);
    if (R.length === 0) return L;
    R[0].join = node.op;
    return L.concat(R);
  }

  function parseViewQueryToFilterRows(viewQuery) {
    if (!viewQuery || typeof viewQuery !== "string") return null;
    const wm = viewQuery.replace(/\s+/g, " ").match(/<Where>\s*([\s\S]*?)\s*<\/Where>/i);
    if (!wm) return null;
    const tree = parseWhereExpr(wm[1].trim());
    if (!tree) return null;
    return binTreeToFilterRows(tree);
  }

  function editorQueryStateFromViewDetails(viewDetails) {
    const d = viewDetails || {};
    let sortLevels;
    if (d.orderByLevels && d.orderByLevels.length) {
      sortLevels = d.orderByLevels.map(function (s) {
        return { field: s.field, ascending: s.ascending !== false };
      });
    } else {
      sortLevels = d.orderBy ? [{ field: d.orderBy.field, ascending: d.orderBy.ascending !== false }] : [];
    }
    const groupByLevels = d.groupByLevels && d.groupByLevels.length ? d.groupByLevels : [];
    const groupByColumn = d.groupBy || (groupByLevels[0] && groupByLevels[0].field) || "";
    return {
      sortLevels: sortLevels,
      groupByColumn: groupByColumn,
      groupExpand: groupByColumn ? d.groupExpand !== false : true,
      groupByAscending: groupByLevels[0] ? groupByLevels[0].ascending !== false : true,
      groupByExtraLevels: groupByLevels.slice(1).map(function (s) {
        return { field: s.field, ascending: s.ascending !== false };
      }),
      groupLimit: d.groupLimit != null && d.groupLimit !== "" ? d.groupLimit : null,
    };
  }

  function applyEditorQueryState(state) {
    sortLevels = state.sortLevels || [];
    groupByColumn = state.groupByColumn || "";
    groupExpand = state.groupExpand !== false;
    groupByAscending = state.groupByAscending !== false;
    groupByExtraLevels = state.groupByExtraLevels || [];
    groupLimit = state.groupLimit != null ? state.groupLimit : null;
    loadedGroupByPrimary = groupByColumn;
  }

  function escapeHtml(s) {
    const div = document.createElement("div");
    div.textContent = s == null ? "" : s;
    return div.innerHTML;
  }
  function escapeAttr(s) {
    return escapeHtml(s == null ? "" : s).replace(/"/g, "&quot;");
  }

  function escapeCsvValue(val) {
    const s = String(val == null ? "" : val);
    if (/[",\r\n]/.test(s)) return "\"" + s.replace(/"/g, "\"\"") + "\"";
    return s;
  }

  function exportViewColumnsToCsv() {
    let headers = [];
    if (viewDetails && viewDetails.viewFields && viewDetails.viewFields.length) {
      headers = viewDetails.viewFields;
    } else {
      headers = columnOrder.filter(function (c) { return inViewSet.has(c); });
    }
    if (!headers.length) {
      showSaveStatus("No columns in view. Add columns to the view first.", true);
      return;
    }
    const csvLine = headers.map(escapeCsvValue).join(",");
    const csv = csvLine + "\r\n";
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const siteName = (sitePath.split("/").filter(Boolean).pop() || "site").replace(/[/\\:*?"<>|]/g, "-").trim() || "site";
    const libName = (listTitle || "list").replace(/[/\\:*?"<>|]/g, "-").replace(/\s+/g, " ").trim() || "list";
    a.download = siteName + "-" + libName + "-columns.csv";
    a.click();
    URL.revokeObjectURL(url);
    showSaveStatus("Column names exported to view-columns.csv", false);
    setTimeout(function () { showSaveStatus("", false); }, 2500);
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

  function showContext(webAbsoluteUrl, listUrl, listIdVal, listTitleVal, listKindLabel) {
    listId = listIdVal || "";
    listTitle = listTitleVal || "";
    if (webAbsoluteUrl) {
      contextWebAbsoluteUrl = String(webAbsoluteUrl).replace(/\/$/, "");
      try {
        const u = new URL(webAbsoluteUrl);
        sitePath = u.pathname.replace(/\/$/, "") || "/";
      } catch (_) {
        sitePath = "/";
      }
    }
    const nameEl = document.getElementById("contextListName");
    const kindEl = document.getElementById("contextListKind");
    const idEl = document.getElementById("contextListIdDisplay");
    const linesEl = document.getElementById("contextLines");
    const label = listKindLabel != null ? String(listKindLabel).trim() : "";
    if (kindEl) {
      kindEl.textContent = label;
      kindEl.classList.toggle("hide", !label);
      kindEl.setAttribute("aria-hidden", label ? "false" : "true");
    }
    if (nameEl) nameEl.textContent = listTitle || "—";
    if (idEl) idEl.textContent = listId || "—";
    if (linesEl) linesEl.classList.add("visible");
    showContextStatus(null);
    updateUrlWithList(listTitle);
    updateViewIdDisplay();
  }

  function updateViewIdDisplay() {
    const wrap = document.getElementById("contextViewIdWrap");
    const idEl = document.getElementById("contextViewIdDisplay");
    const btn = document.getElementById("btnCopyViewId");
    if (wrap) wrap.classList.toggle("hide", !selectedViewId);
    if (idEl) idEl.textContent = selectedViewId || "—";
    if (btn) btn.textContent = "Copy";
  }

  function updatePersonalViewVisibility() {
    const wrap = document.getElementById("viewPersonalWrap");
    if (!wrap) return;
    const show = !selectedViewId;
    wrap.classList.toggle("hide", !show);
    wrap.style.display = show ? "" : "none";
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

  function copyViewIdToClipboard() {
    if (!selectedViewId) return;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(selectedViewId).then(function () {
        const btn = document.getElementById("btnCopyViewId");
        if (btn) { btn.textContent = "Copied!"; setTimeout(function () { btn.textContent = "Copy"; }, 1500); }
      }).catch(function () { fallbackCopyViewId(); });
    } else {
      fallbackCopyViewId();
    }
  }

  function fallbackCopyViewId() {
    const el = document.getElementById("contextViewIdDisplay");
    if (!el || !selectedViewId) return;
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(el);
    if (sel) { sel.removeAllRanges(); sel.addRange(range); }
    try {
      document.execCommand("copy");
      const btn = document.getElementById("btnCopyViewId");
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
      if (forcedViewManagerListId) {
        let webUrl = forcedViewManagerWebUrl;
        if (!webUrl) {
          const pc = await sendToTab({ action: "getPageContext" });
          webUrl = (pc && (pc.webAbsoluteUrl || pc.siteAbsoluteUrl) || "").replace(/\/$/, "");
        }
        if (!webUrl) {
          showContextStatus("Could not determine site URL.", true);
          document.getElementById("noListCard").classList.remove("hide");
          return false;
        }
        try {
          const u = new URL(webUrl);
          sitePath = u.pathname.replace(/\/$/, "") || "/";
        } catch (_) {
          sitePath = "/";
        }
        const titlePath =
          sitePath + "/_api/web/lists(guid'" + forcedViewManagerListId.replace(/'/g, "''") + "')?$select=Title,BaseTemplate";
        const titleRes = await rest("GET", titlePath);
        let title = "";
        let baseTemplate = null;
        if (titleRes && titleRes.ok && titleRes.data) {
          const t = titleRes.data.Title || titleRes.data.title;
          if (t != null) title = String(t);
          if (titleRes.data.BaseTemplate != null) baseTemplate = titleRes.data.BaseTemplate;
        }
        showContext(webUrl, "", forcedViewManagerListId, title, listBaseTemplateLabel(baseTemplate));
        document.getElementById("noListCard").classList.add("hide");
        showContextStatus(null);
        return true;
      }
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
        showContext(webUrl, "", "", "", "");
        showContextStatus("Not on a list or library. Open a list view and try again.", true);
        document.getElementById("noListCard").classList.remove("hide");
        return false;
      }
      const apiPath = sitePath + "/_api/web/GetList(@a)/Id?@a='" + encodeURIComponent(listUrl.replace(/'/g, "''")) + "'";
      const idRes = await rest("GET", apiPath);
      if (!idRes || !idRes.ok) {
        showContext(webUrl, listUrl, "", "", "");
        showContextStatus(idRes && idRes.error ? String(idRes.error) : "Could not resolve list ID.", true);
        return false;
      }
      let lid = (idRes.data && (idRes.data.value !== undefined ? idRes.data.value : (idRes.data.d && idRes.data.d.Id) || idRes.data.Id)) || "";
      if (lid && typeof lid === "string") lid = lid.replace(/[{}]/g, "").trim();
      const titlePath = sitePath + "/_api/web/lists(guid'" + lid + "')?$select=Title,BaseTemplate";
      const titleRes = await rest("GET", titlePath);
      let title = "";
      let baseTemplate = null;
      if (titleRes && titleRes.ok && titleRes.data) {
        const t = titleRes.data.Title || titleRes.data.title;
        if (t != null) title = String(t);
        if (titleRes.data.BaseTemplate != null) baseTemplate = titleRes.data.BaseTemplate;
      }
      showContext(webUrl, listUrl, lid, title, listBaseTemplateLabel(baseTemplate));
      document.getElementById("noListCard").classList.add("hide");
      return true;
    } catch (e) {
      showContextStatus(e && e.message ? e.message : "Failed to load context.", true);
      document.getElementById("noListCard").classList.remove("hide");
      return false;
    }
  }

  /**
   * Content type IDs are hierarchical (parent ID + suffix). Longer prefixes first so
   * 0x0120D5 (Document Set) wins over 0x0120 (Folder). See MS Learn “Content type IDs”.
   */
  const CT_PARENT_PREFIXES = [
    ["0x0120D5", "Document Set"],
    ["0x0120", "Folder"],
    ["0x0110", "Form"],
    ["0x0107", "Task"],
    ["0x0106", "Link"],
    ["0x0105", "Message"],
    ["0x0104", "Discussion"],
    ["0x0102", "Event"],
    ["0x0101", "Document"],
    ["0x01", "Item"]
  ];

  /** Gallery parent for a built-in when the list CT has the same display name as that type. */
  const CT_BUILTIN_GALLERY_PARENT = {
    item: "",
    document: "Item",
    folder: "Item",
    "document set": "Folder",
    event: "Item",
    task: "Item",
    link: "Item",
    message: "Item",
    discussion: "Item",
    form: "Item"
  };

  function normalizeContentTypeId(id) {
    return String(id || "")
      .replace(/\s/g, "")
      .replace(/[{}]/g, "")
      .toUpperCase();
  }

  function inferImmediateParentFromContentTypeId(stringId) {
    const s = String(stringId || "").replace(/\s/g, "");
    if (!s) return "";
    const u = s.toUpperCase();
    if (u === "0X0101") return "Item";
    if (u === "0X0120") return "Item";
    if (u === "0X0120D5") return "Folder";
    for (let i = 0; i < CT_PARENT_PREFIXES.length; i++) {
      const p = CT_PARENT_PREFIXES[i][0].toUpperCase();
      if (u.length > p.length && u.indexOf(p) === 0) {
        return CT_PARENT_PREFIXES[i][1];
      }
    }
    return "";
  }

  /** Parent label for UI: ID inference + fix list CTs named like their parent type. */
  function resolveParentContentTypeLabel(stringId, contentTypeName) {
    const immediate = inferImmediateParentFromContentTypeId(stringId);
    if (!immediate) return "";
    const n = (contentTypeName || "").trim().toLowerCase();
    const key = immediate.trim().toLowerCase();
    if (n && n === key) {
      const gallery = CT_BUILTIN_GALLERY_PARENT[n];
      return gallery !== undefined && gallery !== "" ? gallery : immediate;
    }
    return immediate;
  }

  async function loadContentTypes() {
    const section = document.getElementById("contentTypesSection");
    const tbody = document.getElementById("contentTypesBody");
    const statusEl = document.getElementById("contentTypesStatus");
    if (!section || !tbody) return;
    if (!listId) {
      section.classList.add("hide");
      tbody.innerHTML = "";
      contentTypes = [];
      return;
    }
    section.classList.remove("hide");
    tbody.innerHTML = "";
    contentTypes = [];
    if (statusEl) {
      statusEl.textContent = "Loading content types…";
      statusEl.className = "content-types-hint loading";
    }
    try {
      const listCtBase =
        sitePath + "/_api/web/lists(guid'" + listId.replace(/'/g, "''") + "')/ContentTypes";
      const withParent =
        listCtBase +
        "?$select=Name,StringId,Id,Hidden,Group,Parent/Name,Parent/StringId&$expand=Parent&$orderby=Name";
      let res = await rest("GET", withParent);
      if (!res || !res.ok) {
        res = await rest(
          "GET",
          listCtBase + "?$select=Name,StringId,Id,Hidden,Group&$orderby=Name"
        );
      }
      if (!res || !res.ok) {
        if (statusEl) {
          statusEl.textContent = res && res.error ? String(res.error) : "Could not load content types.";
          statusEl.className = "content-types-hint error";
        }
        return;
      }
      const raw = res.data && (res.data.value != null ? res.data.value : res.data.d && res.data.d.results) || [];
      const arr = Array.isArray(raw) ? raw : [];
      const seen = new Set();
      contentTypes = [];
      for (let i = 0; i < arr.length; i++) {
        const c = arr[i];
        const name = (c.Name != null ? String(c.Name) : c.name != null ? String(c.name) : "").trim();
        let sid = "";
        if (c.StringId != null) sid = String(c.StringId).trim();
        else if (c.stringId != null) sid = String(c.stringId).trim();
        else if (c.Id != null) sid = String(c.Id).trim();
        const key = (sid || name).toLowerCase();
        if (!key || seen.has(key)) continue;
        seen.add(key);
        const hidden = !!(c.Hidden || c.hidden);
        let parentName = "";
        let parentSid = "";
        const par = c.Parent;
        if (par && typeof par === "object" && !par.__deferred) {
          parentName = (par.Name != null ? String(par.Name) : par.name != null ? String(par.name) : "").trim();
          if (par.StringId != null) parentSid = String(par.StringId).trim();
          else if (par.stringId != null) parentSid = String(par.stringId).trim();
          else if (par.Id != null) parentSid = String(par.Id).trim();
        }
        const sidNorm = sid ? normalizeContentTypeId(sid) : "";
        const parSidNorm = parentSid ? normalizeContentTypeId(parentSid) : "";
        const apiParentSameAsSelf =
          (parentName && name && parentName.toLowerCase() === name.toLowerCase()) ||
          (sidNorm && parSidNorm && parSidNorm === sidNorm);
        if (apiParentSameAsSelf) {
          parentName = "";
          parentSid = "";
        }
        if (!parentName && sid && sid !== "—") {
          parentName = resolveParentContentTypeLabel(sid, name);
        }
        contentTypes.push({
          name: name || "—",
          stringId: sid || "—",
          hidden: hidden,
          parentName: parentName,
          parentStringId: parentSid
        });
      }
      tbody.innerHTML = "";
      contentTypes.forEach(function (ct) {
        const tr = document.createElement("tr");
        const tdName = document.createElement("td");
        tdName.textContent = ct.hidden ? ct.name + " (hidden)" : ct.name;
        const tdParent = document.createElement("td");
        tdParent.className = "ct-parent-cell";
        if (ct.parentName) {
          tdParent.textContent = ct.parentName;
          if (ct.parentStringId) {
            tdParent.title = "Parent ID: " + ct.parentStringId;
          }
        } else {
          tdParent.textContent = "—";
          tdParent.classList.add("ct-parent-unknown");
        }
        const tdId = document.createElement("td");
        const wrap = document.createElement("div");
        wrap.className = "ct-id-wrap";
        const code = document.createElement("code");
        code.className = "ct-id-cell";
        code.textContent = ct.stringId;
        wrap.appendChild(code);
        const copyBtn = document.createElement("button");
        copyBtn.type = "button";
        copyBtn.className = "secondary ct-copy-btn";
        copyBtn.textContent = "Copy";
        copyBtn.title = "Copy content type ID";
        copyBtn.setAttribute("aria-label", "Copy content type ID for " + (ct.name || "type"));
        const idToCopy = ct.stringId && ct.stringId !== "—" ? ct.stringId : "";
        copyBtn.disabled = !idToCopy;
        copyBtn.addEventListener("click", function () {
          if (!idToCopy) return;
          if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(idToCopy).then(function () {
              copyBtn.textContent = "Copied!";
              setTimeout(function () { copyBtn.textContent = "Copy"; }, 1200);
            });
          }
        });
        wrap.appendChild(copyBtn);
        tdId.appendChild(wrap);
        tr.appendChild(tdName);
        tr.appendChild(tdParent);
        tr.appendChild(tdId);
        tbody.appendChild(tr);
      });
      if (statusEl) {
        statusEl.textContent =
          contentTypes.length + " content type" + (contentTypes.length !== 1 ? "s" : "") + " on this list/library.";
        statusEl.className = "content-types-hint";
      }
    } catch (e) {
      if (statusEl) {
        statusEl.textContent = e && e.message ? e.message : "Failed to load content types.";
        statusEl.className = "content-types-hint error";
      }
    }
  }

  async function loadListEditorData() {
    await Promise.all([loadViewsAndFields(), loadContentTypes()]);
  }

  async function loadViewsAndFields() {
    if (!listId) return;
    try {
      const res = await sendToTab({
        action: "getViewsData",
        listId: forcedViewManagerListId || undefined,
        webAbsoluteUrl: contextWebAbsoluteUrl || undefined
      });
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
      if (forceCreateNewView) {
        selectedViewId = null;
        const sel = document.getElementById("viewSelect");
        if (sel) sel.value = "";
        onViewSelectChange();
      } else {
        selectDefaultView();
      }
    } catch (e) {
      showSaveStatus(e && e.message ? e.message : "Failed to load data.", true);
    }
  }

  async function loadViewDetails(viewId) {
    if (!viewId) return;
    try {
      const res = await sendToTab({
        action: "getViewsData",
        viewId: viewId,
        listId: forcedViewManagerListId || undefined,
        webAbsoluteUrl: contextWebAbsoluteUrl || undefined
      });
      if (!res || res.error || !res.viewDetails) {
        showSaveStatus(res && res.error ? res.error : "Failed to load view details.", true);
        return;
      }
      viewDetails = res.viewDetails;
      const viewFields = viewDetails.viewFields || [];
      inViewSet = new Set(viewFields);
      const otherFields = (fields || []).filter(function (f) { return viewFields.indexOf(f.internalName) < 0; });
      otherFields.sort(function (a, b) { return (a.title || "").localeCompare(b.title || ""); });
      columnOrder = viewFields.slice();
      otherFields.forEach(function (f) { columnOrder.push(f.internalName); });
      applyEditorQueryState(editorQueryStateFromViewDetails(viewDetails));
      const parsedFq = parseViewQueryToFilterRows(viewDetails.viewQuery || "");
      if (parsedFq && parsedFq.length) {
        filters = parsedFq.map(function (row) {
          let val = row.value || "";
          if (isBooleanFieldName(row.field)) val = booleanFilterRawToDisplay(val);
          const o = { field: row.field, op: row.op, value: val };
          if (row.join) o.join = row.join;
          return o;
        });
      } else {
        filters = (viewDetails.filters || []).map(function (f, idx) {
          let val = f.value || "";
          if (isBooleanFieldName(f.field)) val = booleanFilterRawToDisplay(val);
          const o = { field: f.field, op: f.op, value: val };
          if (idx > 0) o.join = "And";
          return o;
        });
      }
      normalizeFilterJoins(filters);
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
    inViewSet = new Set();
    const sorted = (fields || []).slice().sort(function (a, b) { return (a.title || "").localeCompare(b.title || ""); });
    columnOrder = sorted.map(function (f) { return f.internalName; });
    sortLevels = [];
    filters = [];
    groupByColumn = "";
    groupExpand = true;
    groupByAscending = true;
    groupByExtraLevels = [];
    groupLimit = null;
    loadedGroupByPrimary = "";
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
      document.getElementById("btnDuplicate").classList.remove("hide");
      updateViewIdDisplay();
      updateDeleteButtonVisibility();
      loadViewDetails(defaultView.id);
    } else {
      setNewViewDefaults();
      updateViewIdDisplay();
    }
  }

  function onViewSelectChange() {
    const val = document.getElementById("viewSelect").value;
    selectedViewId = val || null;
    updateViewIdDisplay();
    updatePersonalViewVisibility();
    if (!val) {
      setNewViewDefaults();
      updateSetDefaultVisibility();
      return;
    }
    document.getElementById("btnDuplicate").classList.remove("hide");
    updateDeleteButtonVisibility();
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

  function updateDeleteButtonVisibility() {
    const btn = document.getElementById("btnDelete");
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
    const toShow = ordered.filter(function (name) {
      const f = fields.find(function (x) { return x.internalName === name; });
      return f && (!search || (f.title || "").toLowerCase().indexOf(search) >= 0 || (f.internalName || "").toLowerCase().indexOf(search) >= 0);
    }).map(function (name) {
      return fields.find(function (f) { return f.internalName === name; });
    }).filter(Boolean);

    list.innerHTML = "";
    let draggedName = null;
    function addRow(field) {
      const isInView = inViewSet.has(field.internalName);
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
          inViewSet.add(field.internalName);
        } else {
          inViewSet.delete(field.internalName);
        }
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
        const arr = columnOrder.filter(function (n) { return n !== name; });
        const insertIdx = arr.indexOf(field.internalName);
        arr.splice(insertIdx >= 0 ? insertIdx : arr.length, 0, name);
        columnOrder = arr;
        renderColumnList();
      });
      list.appendChild(div);
    }
    toShow.forEach(function (f) { addRow(f); });
  }

  document.getElementById("colSearch").addEventListener("input", function () { renderColumnList(); });
  document.getElementById("btnColAddAll").addEventListener("click", function () {
    columnOrder.forEach(function (name) { inViewSet.add(name); });
    renderColumnList();
  });
  document.getElementById("btnColRemoveAll").addEventListener("click", function () {
    inViewSet.clear();
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
    normalizeFilterJoins(filters);
    filters.forEach(function (f, i) {
      const block = document.createElement("div");
      block.className = "filter-block";
      if (i > 0) {
        const joinRow = document.createElement("div");
        joinRow.className = "filter-join-row";
        const legend = document.createElement("span");
        legend.textContent = "Match previous row with:";
        joinRow.appendChild(legend);
        ["And", "Or"].forEach(function (opName) {
          const id = "filterJoin_" + i + "_" + opName;
          const label = document.createElement("label");
          const radio = document.createElement("input");
          radio.type = "radio";
          radio.name = "filterJoinGroup_" + i;
          radio.value = opName;
          radio.id = id;
          radio.checked = (f.join || "And") === opName;
          radio.addEventListener("change", function () {
            if (radio.checked) filters[i].join = opName;
          });
          label.appendChild(radio);
          label.appendChild(document.createTextNode(" " + opName));
          joinRow.appendChild(label);
        });
        block.appendChild(joinRow);
      }
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
      const valueWrap = document.createElement("span");
      valueWrap.className = "filter-value-wrap";
      if (isBooleanFieldName(f.field)) {
        const valSel = document.createElement("select");
        valSel.className = "filter-value filter-value-bool";
        [["", "—"], ["Yes", "Yes"], ["No", "No"]].forEach(function (pair) {
          const opt = document.createElement("option");
          opt.value = pair[0];
          opt.textContent = pair[1];
          valSel.appendChild(opt);
        });
        const cur = (f.value || "").trim();
        valSel.value = cur === "Yes" || cur === "No" ? cur : "";
        valSel.addEventListener("change", function () { filters[i].value = valSel.value; });
        valueWrap.appendChild(valSel);
      } else {
        const valInput = document.createElement("input");
        valInput.type = "text";
        valInput.className = "filter-value";
        valInput.autocomplete = "off";
        valInput.spellcheck = false;
        valInput.value = f.value || "";
        valInput.placeholder = filterValuePlaceholderForField(f.field);
        valueWrap.appendChild(valInput);
        const kinds = filterTypeaheadKindsForField(f.field);
        if (kinds.length > 0 && filterTypeaheadApi()) {
          valueWrap.classList.add("filter-value-typeahead-wrap");
          attachFilterValueTypeahead(valInput, i);
        } else {
          valInput.addEventListener("input", function () { filters[i].value = valInput.value; });
        }
      }
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "link";
      btn.textContent = "Remove";
      btn.addEventListener("click", function () {
        filters.splice(i, 1);
        normalizeFilterJoins(filters);
        renderFilterConditions();
      });
      row.appendChild(fieldSel);
      row.appendChild(opSel);
      row.appendChild(valueWrap);
      row.appendChild(btn);
      block.appendChild(row);
      container.appendChild(block);
      fieldSel.addEventListener("change", function () {
        filters[i].field = fieldSel.value;
        if (isBooleanFieldName(fieldSel.value)) {
          filters[i].value = coerceBooleanFilterStoredValue(filters[i].value || "");
        }
        renderFilterConditions();
      });
      opSel.addEventListener("change", function () { filters[i].op = opSel.value; });
    });
  }

  document.getElementById("btnAddFilter").addEventListener("click", function () {
    const row = {
      field: fields[0] ? fields[0].internalName : "",
      op: "Eq",
      value: ""
    };
    if (filters.length > 0) row.join = "And";
    filters.push(row);
    normalizeFilterJoins(filters);
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
    if (groupByColumn !== loadedGroupByPrimary) {
      groupByExtraLevels = [];
      groupByAscending = true;
    }
  });
  document.getElementById("groupExpand").addEventListener("change", function () {
    groupExpand = document.getElementById("groupExpand").checked;
  });

  function isBooleanFieldName(fieldInternalName) {
    const fd = fields.find(function (f) { return f.internalName === fieldInternalName; });
    const t = (fd && fd.typeAsString) ? String(fd.typeAsString) : "";
    return /Boolean|YesNo/i.test(t);
  }

  /** Map CAML / API values to display Yes/No for Yes/No columns */
  function booleanFilterRawToDisplay(raw) {
    const s = String(raw == null ? "" : raw).trim();
    if (s === "") return "";
    const low = s.toLowerCase();
    if (low === "0" || low === "false" || low === "no") return "No";
    if (low === "1" || low === "true" || low === "yes") return "Yes";
    return s;
  }

  /** When switching a row to a Yes/No field, keep only Yes/No/blank */
  function coerceBooleanFilterStoredValue(v) {
    const d = booleanFilterRawToDisplay(v);
    if (d === "Yes" || d === "No" || d === "") return d;
    return "";
  }

  /** CAML Yes/No uses Type=Integer with 0 or 1 */
  function booleanFilterDisplayToCamlInteger(s) {
    const t = String(s || "").trim().toLowerCase();
    if (t === "yes" || t === "true" || t === "1") return "1";
    if (t === "no" || t === "false" || t === "0") return "0";
    return null;
  }

  function getValueTypeForField(fieldInternalName) {
    const fd = fields.find(function (f) { return f.internalName === fieldInternalName; });
    const t = (fd && fd.typeAsString) ? String(fd.typeAsString) : "";
    if (/Integer|Counter|Boolean|YesNo/i.test(t)) return "Integer";
    if (/Number|Currency|Decimal/i.test(t)) return "Number";
    if (/DateTime|Date/i.test(t)) return "DateTime";
    return "Text";
  }

  function buildOneFilterCondXml(f) {
    let type = getValueTypeForField(f.field).replace(/"/g, "");
    let inner;
    if (isBooleanFieldName(f.field)) {
      const camlInt = booleanFilterDisplayToCamlInteger(f.value);
      if (camlInt == null) return null;
      inner = camlInt;
      type = "Integer";
    } else {
      inner = String(f.value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }
    return "<" + f.op + "><FieldRef Name=\"" + escapeAttr(f.field) + "\"/><Value Type=\"" + type + "\">" + inner + "</Value></" + f.op + ">";
  }

  /** Left-associative And/Or chain matching rows with non-blank values (skips blank rows). */
  function buildFilterWhereXml() {
    const chunks = [];
    for (let i = 0; i < filters.length; i++) {
      const f = filters[i];
      if (!(f.value || "").trim()) continue;
      const xml = buildOneFilterCondXml(f);
      if (!xml) continue;
      chunks.push({ xml: xml, join: chunks.length === 0 ? null : (f.join || "And") });
    }
    if (chunks.length === 0) return "";
    let acc = chunks[0].xml;
    for (let k = 1; k < chunks.length; k++) {
      const op = chunks[k].join || "And";
      acc = "<" + op + ">" + acc + chunks[k].xml + "</" + op + ">";
    }
    return "<Where>" + acc + "</Where>";
  }

  function buildViewQuery() {
    const parts = [];
    const whereXml = buildFilterWhereXml();
    if (whereXml) parts.push(whereXml);
    if (sortLevels.length > 0) {
      const refs = sortLevels.filter(function (s) { return s.field; }).map(function (s) {
        return "<FieldRef Name=\"" + escapeAttr(s.field) + "\" Ascending=\"" + (s.ascending ? "True" : "False") + "\"/>";
      });
      if (refs.length > 0) parts.push("<OrderBy>" + refs.join("") + "</OrderBy>");
    }
    if (groupByColumn) {
      const collapse = groupExpand ? "FALSE" : "TRUE";
      const limitAttr = groupLimit != null && groupLimit !== "" ? " GroupLimit=\"" + String(groupLimit) + "\"" : "";
      const refs = [];
      refs.push("<FieldRef Name=\"" + escapeAttr(groupByColumn) + "\" Ascending=\"" + (groupByAscending ? "True" : "False") + "\"/>");
      if (groupByColumn === loadedGroupByPrimary) {
        groupByExtraLevels.forEach(function (s) {
          if (!s || !s.field || s.field === groupByColumn) return;
          refs.push("<FieldRef Name=\"" + escapeAttr(s.field) + "\" Ascending=\"" + (s.ascending ? "True" : "False") + "\"/>");
        });
      }
      parts.push("<GroupBy Collapse=\"" + collapse + "\"" + limitAttr + ">" + refs.join("") + "</GroupBy>");
    }
    return parts.join("");
  }

  function showSaveStatus(msg, isError) {
    const el = document.getElementById("saveStatus");
    el.textContent = msg || "";
    el.className = "status " + (isError ? "error" : "success");
    el.style.display = msg ? "block" : "none";
  }

  function getOrderedViewColumns() {
    return columnOrder.filter(function (n) { return inViewSet.has(n); });
  }

  async function applyViewFields(viewId) {
    const ordered = getOrderedViewColumns();
    const vfBase = sitePath + "/_api/web/lists(guid'" + listId.replace(/'/g, "''") + "')/views(guid'" + viewId.replace(/'/g, "''") + "')/ViewFields";
    const rm = await rest("POST", vfBase + "/RemoveAllViewFields");
    if (!rm.ok && rm.status !== 404 && rm.status !== 501) throw new Error(rm.error || "RemoveAllViewFields failed");
    for (let i = 0; i < ordered.length; i++) {
      const name = String(ordered[i]).replace(/'/g, "''");
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
    if (getOrderedViewColumns().length === 0) {
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
      updateViewIdDisplay();
      updatePersonalViewVisibility();
      updateDeleteButtonVisibility();
      loadContext().then(function () { loadListEditorData(); });
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
      loadContext().then(function () { loadListEditorData(); });
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
    updateViewIdDisplay();
    updatePersonalViewVisibility();
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
  document.getElementById("btnCopyViewId").addEventListener("click", copyViewIdToClipboard);

  document.getElementById("btnSave").addEventListener("click", saveView);
  document.getElementById("btnSetDefault").addEventListener("click", setDefaultView);
  document.getElementById("btnDelete").addEventListener("click", deleteView);
  document.getElementById("btnDuplicate").addEventListener("click", duplicateView);
  document.getElementById("btnExportViewColumns").addEventListener("click", exportViewColumnsToCsv);
  document.getElementById("viewSelect").addEventListener("change", onViewSelectChange);

  (async function init() {
    const ok = await loadContext();
    if (ok) await loadListEditorData();
  })();
})();
