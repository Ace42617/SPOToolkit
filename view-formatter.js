import { isSharePointOnlineUrl } from "./lib/sharePointUrl.mjs";

(function () {
  function escHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  /** Lex JSON-ish text for display (tolerates invalid/truncated input). */
  function highlightJsonHtml(text) {
    if (!text) return "";
    const n = text.length;
    let i = 0;
    const parts = [];
    while (i < n) {
      const ch = text[i];
      if (/\s/.test(ch)) {
        let j = i + 1;
        while (j < n && /\s/.test(text[j])) j++;
        parts.push(escHtml(text.slice(i, j)));
        i = j;
        continue;
      }
      if (ch === '"') {
        let j = i + 1;
        let segment = '"';
        while (j < n) {
          if (text[j] === "\\" && j + 1 < n) {
            segment += text[j] + text[j + 1];
            j += 2;
            continue;
          }
          if (text[j] === '"') {
            segment += '"';
            j++;
            break;
          }
          segment += text[j];
          j++;
        }
        let k = j;
        while (k < n && /\s/.test(text[k])) k++;
        const isKey = text[k] === ":";
        parts.push('<span class="vf-' + (isKey ? "jkey" : "jstr") + '">' + escHtml(segment) + "</span>");
        i = j;
        continue;
      }
      if (text.startsWith("true", i)) {
        parts.push('<span class="vf-jkw">' + escHtml("true") + "</span>");
        i += 4;
        continue;
      }
      if (text.startsWith("false", i)) {
        parts.push('<span class="vf-jkw">' + escHtml("false") + "</span>");
        i += 5;
        continue;
      }
      if (text.startsWith("null", i)) {
        parts.push('<span class="vf-jkw">' + escHtml("null") + "</span>");
        i += 4;
        continue;
      }
      if (ch === "-" && i + 1 < n && /\d/.test(text[i + 1]) || /\d/.test(ch)) {
        let j = i;
        if (text[j] === "-") j++;
        while (j < n && /[\d.eE+\-]/.test(text[j])) j++;
        parts.push('<span class="vf-jnum">' + escHtml(text.slice(i, j)) + "</span>");
        i = j;
        continue;
      }
      if ("{}[],:".indexOf(ch) >= 0) {
        parts.push('<span class="vf-jpunc">' + escHtml(ch) + "</span>");
        i++;
        continue;
      }
      parts.push(escHtml(ch));
      i++;
    }
    return parts.join("");
  }

  /** 1-based line that contains the character at index `pos` (0-based). */
  function offsetToLineOneBased(text, pos) {
    if (pos <= 0) return 1;
    let line = 1;
    const end = Math.min(pos, text.length);
    for (let i = 0; i < end; i++) {
      if (text.charCodeAt(i) === 10) line++;
    }
    return line;
  }

  /** Map 1-based line/column (as in engine messages) to 0-based offset. */
  function lineColOneBasedToOffset(text, line1, col1) {
    if (line1 < 1) return 0;
    let line = 1;
    let i = 0;
    const n = text.length;
    while (i < n && line < line1) {
      if (text.charCodeAt(i) === 10) line++;
      i++;
    }
    if (line < line1) return n;
    const colOffset = Math.max(0, (col1 || 1) - 1);
    return Math.min(i + colOffset, n);
  }

  /** Last non-whitespace character index before `pos`, or -1. */
  function charIndexBeforeWs(text, pos) {
    let i = Math.min(pos, text.length) - 1;
    while (i >= 0 && /[\s\r\n]/.test(text[i])) i--;
    return i;
  }

  /**
   * Engine `position` often points at the *next* token (e.g. next key) when a comma is missing.
   * Show the callout after the line where the fix belongs (usually the previous value’s closing token).
   */
  function displayLineForJsonError(text, pos, cleanedMessage) {
    const msg = (cleanedMessage || "").toLowerCase();
    const bounded = Math.min(Math.max(0, pos), text.length);
    const i = charIndexBeforeWs(text, bounded);

    const commaOrSep =
      /expected\s*','\s*or/i.test(msg) ||
      /expected\s*','\s*after/i.test(msg) ||
      (msg.includes("expected") && msg.includes("','"));

    if (commaOrSep && i >= 0) {
      let k = bounded;
      while (k < text.length && /[ \t\r\n]/.test(text[k])) k++;
      if (/after\s*array\s*element/i.test(msg) && k < text.length && text[k] === "}") {
        return offsetToLineOneBased(text, k);
      }
      const c = text[i];
      if (c === "}" || c === "]" || c === '"' || /[0-9]/.test(c)) {
        return offsetToLineOneBased(text, i);
      }
    }

    if (bounded <= 0) return 1;
    return offsetToLineOneBased(text, bounded);
  }

  /** Narrow vague "Expected ',' or ']'…" style text using nearby source. */
  function refineJsonErrorMessage(text, pos, cleaned) {
    const m = cleaned.trim();
    const mi = m.toLowerCase();
    const bounded = Math.min(Math.max(0, pos), text.length);
    const i = charIndexBeforeWs(text, bounded);
    const prev = i >= 0 ? text[i] : "";

    if (/expected\s*','\s*or\s*'\]'\s*after\s*property\s*value/i.test(m)) {
      if (prev === "}") return "Missing a comma (,) after `}` — add `,` before the next property.";
      if (prev === '"') return "Missing a comma (,) after this property value — add `,` before the next key.";
      return "Missing a comma (,) between properties (or the list ended too early).";
    }
    if (/expected\s*','\s*or\s*'}'\s*after\s*property\s*value/i.test(m)) {
      if (prev === "]") return "Missing a comma (,) after `]` here, or close the object with `}` if the array ended.";
      if (prev === "}") return "Missing a comma (,) after `}` or add `}` to close the object.";
      if (prev === '"') return "Missing a comma (,) between properties.";
      return m;
    }
    if (/expected\s*','\s*or\s*'\]'\s*after\s*array\s*element/i.test(m)) {
      return "Missing a comma (,) between array elements, or close the array with `]`.";
    }
    if (/expected\s*','\s*after\s*array\s*element/i.test(m)) {
      return "Missing a comma (,) between array elements.";
    }
    if (/expected\s*':'/i.test(m) && (mi.includes("property") || mi.includes("name"))) {
      return "Missing a colon (:) after the property name — use `\"key\": value`.";
    }
    if (/expected\s*'}'/i.test(m) && mi.includes("after")) {
      return "Missing `}` to close an object.";
    }
    if (/expected\s*'\]'/i.test(m) && mi.includes("after")) {
      return "Missing `]` to close an array.";
    }
    return m;
  }

  /** Insert index for floating chip: immediately after the last character tied to this error. */
  function computeAnchorOffset(text, pos, cleanedEngineMessage) {
    const msg = (cleanedEngineMessage || "").toLowerCase();
    const bounded = Math.min(Math.max(0, pos), text.length);
    const i = charIndexBeforeWs(text, bounded);
    const commaOrSep =
      /expected\s*','\s*or/i.test(msg) ||
      /expected\s*','\s*after/i.test(msg) ||
      (msg.includes("expected") && msg.includes("','"));
    if (commaOrSep && i >= 0) {
      let k = bounded;
      while (k < text.length && /[ \t\r\n]/.test(text[k])) k++;
      if (/after\s*array\s*element/i.test(msg) && k < text.length && text[k] === "}") {
        return k;
      }
      const c = text[i];
      if (c === "}" || c === "]" || c === '"' || /[0-9]/.test(c)) {
        return i + 1;
      }
    }
    return bounded;
  }

  /** String-aware bracket scan — same class of checks as JSON validators (e.g. jsonchecker.com / json.org rules). */
  function bracketBalanceHint(text) {
    let inStr = false;
    let esc = false;
    const stack = [];
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inStr) {
        if (esc) esc = false;
        else if (c === "\\") esc = true;
        else if (c === '"') inStr = false;
        continue;
      }
      if (c === '"') {
        inStr = true;
        continue;
      }
      if (c === "{") stack.push("}");
      else if (c === "[") stack.push("]");
      else if (c === "}" || c === "]") {
        const want = stack.pop();
        if (want !== c) {
          return (
            "[Brackets] Expected `" +
            (want == null ? "a matching opener" : want) +
            "` but found `" +
            c +
            "`."
          );
        }
      }
    }
    if (stack.length) {
      return (
        "[Brackets] " +
        stack.length +
        " unclosed opener(s) — still need " +
        stack
          .map(function (x) {
            return "`" + x + "`";
          })
          .join(", ") +
        "."
      );
    }
    return "";
  }

  /** Extra hints: comments, quotes, trailing commas (standard JSON / JSON Checker style guidance). */
  function enrichJsonDiagnostics(text, pos, message) {
    let m = message;
    const p = pos != null && !isNaN(pos) ? Math.min(Math.max(0, pos), text.length) : 0;
    const around = text.slice(Math.max(0, p - 48), Math.min(text.length, p + 48));
    const mi = m.toLowerCase();

    if (
      /\/\/|\/\*/.test(around) &&
      mi.indexOf("comment") < 0 &&
      (mi.indexOf("unexpected") >= 0 || mi.indexOf("invalid") >= 0)
    ) {
      m =
        m +
        " — JSON does not allow comments (`//` or `/* */`). Remove them (same rules as json.org and tools like JSON Checker).";
    }

    if (
      /'(?:[^'\\]|\\.)*'\s*:/.test(text.slice(0, Math.min(text.length, p + 120))) &&
      mi.indexOf("double quote") < 0 &&
      mi.indexOf("single") < 0
    ) {
      m = "JSON requires double-quoted keys and strings (not single quotes). " + m;
    }

    if (/,\s*[\]\}]/.test(around) && mi.indexOf("trailing") < 0 && mi.indexOf("comma") < 0) {
      m = "Trailing comma before `]` or `}` — standard JSON does not allow it; remove that comma.";
    }

    const bh = bracketBalanceHint(text);
    if (bh && m.indexOf("[Brackets]") < 0) m = m + " " + bh;

    return m;
  }

  /** Strip position / line-column from engine message; return { line, message, position, anchorOffset }. */
  function parseJsonError(text, err) {
    const raw = err && err.message ? String(err.message) : "Invalid JSON";
    const posMatch = raw.match(/at\s+position\s+(\d+)/i);
    const lineColMatch = raw.match(/\(\s*line\s+(\d+)\s+column\s+(\d+)\s*\)/i);
    let pos = null;
    if (posMatch) pos = parseInt(posMatch[1], 10);
    else if (lineColMatch) {
      pos = lineColOneBasedToOffset(text, parseInt(lineColMatch[1], 10), parseInt(lineColMatch[2], 10));
    }

    let cleaned = raw
      .replace(/\s*at\s+position\s+\d+\s*/gi, " ")
      .replace(/\s*\(\s*line\s+\d+\s+column\s+\d+\s*\)\s*/gi, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!cleaned) cleaned = "Invalid JSON";

    let line = null;
    let anchorOffset = 0;
    if (pos != null && !isNaN(pos)) {
      pos = Math.min(Math.max(0, pos), text.length);
      line = displayLineForJsonError(text, pos, cleaned);
      anchorOffset = computeAnchorOffset(text, pos, cleaned);
    }

    let message = refineJsonErrorMessage(text, pos != null ? pos : 0, cleaned);
    message = enrichJsonDiagnostics(text, pos != null ? pos : 0, message);
    if (line != null && !isNaN(line) && line >= 1) {
      message = "Line " + line + ": " + message;
    }

    return { line, message, position: pos, anchorOffset: anchorOffset };
  }

  function getJsonValidation(text) {
    const t = text.trim();
    if (!t) {
      return {
        valid: false,
        error: {
          line: 1,
          message: "Line 1: Add JSON before you can save or refresh the preview.",
          anchorOffset: 0,
        },
      };
    }
    try {
      JSON.parse(t);
      return { valid: true, error: null };
    } catch (e) {
      return { valid: false, error: parseJsonError(text, e) };
    }
  }

  function buildEditorHighlight(text) {
    const lines = text.split(/\r?\n/);
    const parts = [];
    for (let li = 0; li < lines.length; li++) {
      parts.push('<span class="vf-line">');
      parts.push(highlightJsonHtml(lines[li]));
      parts.push("</span>");
    }
    return parts.join("");
  }

  const params = new URLSearchParams(window.location.search);
  const rawSrc = params.get("src") || "";
  let previewUrl = "";
  try {
    previewUrl = decodeURIComponent(rawSrc);
  } catch (_) {
    previewUrl = rawSrc;
  }
  if (previewUrl && !isSharePointOnlineUrl(previewUrl, { allowTenantAdmin: false })) {
    previewUrl = "";
  }

  const tabIdRaw = params.get("tabId") || "";
  const spTabId = /^\d+$/.test(tabIdRaw) ? parseInt(tabIdRaw, 10) : null;

  const frame = document.getElementById("vfFrame");
  const previewEl = document.getElementById("vfPreview");
  const splitEl = document.getElementById("vfSplit");
  const resizer = document.getElementById("vfResizer");
  const jsonEl = document.getElementById("vfJson");
  const hlEl = document.getElementById("vfJsonHl");
  const statusEl = document.getElementById("vfJsonStatus");
  const labelEl = document.getElementById("vfSrcLabel");
  const bannerEl = document.getElementById("vfBanner");
  const previewMask = document.getElementById("vfPreviewMask");
  const btnRefresh = document.getElementById("vfRefreshFrame");
  const btnSaveToView = document.getElementById("vfSaveToView");
  const viewSelectEl = document.getElementById("vfViewSelect");
  const btnOpenViewManager = document.getElementById("vfOpenViewManager");
  const codeSearchInput = document.getElementById("vfCodeSearch");
  const codeSearchPrevBtn = document.getElementById("vfCodeSearchPrev");
  const codeSearchNextBtn = document.getElementById("vfCodeSearchNext");
  const codeSearchCountEl = document.getElementById("vfCodeSearchCount");

  let contextListId = "";
  let contextSitePath = "/";
  let contextViewId = "";
  let selectedViewId = "";
  let contextWebAbsoluteUrl = "";
  let codeSearchMatches = [];
  let codeSearchIndex = -1;

  function normGuid(s) {
    return String(s || "")
      .replace(/[{}]/g, "")
      .replace(/%7B|%7D/gi, "")
      .trim();
  }

  function restErrorMessage(payload) {
    if (payload == null) return "";
    if (typeof payload === "string") return payload;
    const oe = payload["odata.error"] || payload.error;
    if (oe && oe.message && typeof oe.message.value === "string") return oe.message.value;
    if (oe && typeof oe.message === "string") return oe.message;
    try {
      return JSON.stringify(payload);
    } catch (_) {
      return "Request failed";
    }
  }

  function sendToSpTab(message) {
    return new Promise(function (resolve, reject) {
      if (spTabId == null) {
        reject(new Error("NO_TAB"));
        return;
      }
      try {
        chrome.tabs.sendMessage(spTabId, message, function (response) {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message || "Could not reach the SharePoint tab."));
            return;
          }
          resolve(response);
        });
      } catch (e) {
        reject(e);
      }
    });
  }

  const STORAGE_SPLIT = "spotk-vf-split-pct";
  /** Draft cache is per list URL + view id so switching views does not reuse the wrong JSON. */
  function storageKeyForView(viewId) {
    let h = 0;
    const s = (previewUrl || "default") + "|" + normGuid(viewId || "");
    for (let i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
    return "spotk-vf-draft-" + (h >>> 0).toString(16);
  }
  const storageKeyDraft = function () {
    return storageKeyForView(selectedViewId || contextViewId);
  };

  function setPreviewFrameSrc(url) {
    if (!frame || !url) return;
    try {
      const u = new URL(url);
      u.searchParams.set("_spotkVf", "1");
      u.searchParams.set("_toolkitVfT", String(Date.now()));
      frame.src = u.toString();
    } catch (_) {
      const sep = url.indexOf("?") >= 0 ? "&" : "?";
      frame.src = url + sep + "_spotkVf=1&_toolkitVfT=" + String(Date.now());
    }
  }

  let livePreviewTimer = null;
  let livePreviewStatusTimer = null;
  function scheduleLivePreview(immediate) {
    if (!frame || !jsonEl || !previewUrl) return;
    const v = getJsonValidation(jsonEl.value);
    if (!v.valid) return;
    const send = function () {
      let origin;
      try {
        origin = new URL(previewUrlForSelectedView() || previewUrl).origin;
      } catch (_) {
        origin = "*";
      }
      try {
        frame.contentWindow.postMessage({ __spotkVfPreview: true, json: jsonEl.value }, origin);
      } catch (_) {}
    };
    if (immediate) {
      send();
    } else {
      if (livePreviewTimer) clearTimeout(livePreviewTimer);
      livePreviewTimer = setTimeout(send, 320);
    }
  }

  function previewUrlForSelectedView() {
    return buildPreviewUrlForView(selectedViewId || contextViewId) || previewUrl || "";
  }

  if (frame && previewUrl) {
    frame.addEventListener("load", function () {
      scheduleLivePreview(true);
    });
  }

  window.addEventListener("message", function (ev) {
    if (!frame || ev.source !== frame.contentWindow) return;
    if (!ev.data || ev.data.__spotkVfPreviewResult !== true) return;
    const r = ev.data.result || {};
    if (!r.ok || !statusEl) return;
    const v = getJsonValidation(jsonEl.value);
    if (!v.valid) return;
    statusEl.textContent = "Valid JSON — live preview updated";
    statusEl.className = "vf-json-status ok";
    if (livePreviewStatusTimer) clearTimeout(livePreviewStatusTimer);
    livePreviewStatusTimer = setTimeout(function () {
      if (getJsonValidation(jsonEl.value).valid) applyValidationUi();
    }, 2000);
  });

  const DEFAULT_TEMPLATE = [
    "{",
    '  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/view-formatting.schema.json",',
    '  "hideSelection": true,',
    '  "rowFormatter": {',
    '    "elmType": "div",',
    '    "style": { "padding-left": "12px" },',
    '    "children": [',
    "      {",
    '        "elmType": "span",',
    '        "style": { "font-weight": "600" },',
    '        "txtContent": "[$Title]"',
    "      }",
    "    ]",
    "  }",
    "}",
  ].join("\n");

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

  function updateLineGutter() {
    const lineNumsEl = document.getElementById("vfLineNums");
    if (!lineNumsEl || !jsonEl) return;
    const lines = (jsonEl.value || "").split(/\r?\n/);
    const n = Math.max(lines.length, 1);
    lineNumsEl.textContent = Array.from({ length: n }, (_, j) => String(j + 1)).join("\n");
    const gutter = lineNumsEl.closest(".vf-line-gutter");
    if (gutter) {
      const w = String(n).length;
      gutter.style.minWidth = Math.max(2.25, w + 1.4) + "ch";
    }
  }

  function syncEditorScrollTargets() {
    const hlPre = hlEl && hlEl.parentElement;
    const gScroll = document.getElementById("vfLineGutterScroll");
    if (hlPre) {
      hlPre.scrollTop = jsonEl.scrollTop;
      hlPre.scrollLeft = jsonEl.scrollLeft;
    }
    if (gScroll) gScroll.scrollTop = jsonEl.scrollTop;
  }

  function syncJsonHighlight() {
    if (!hlEl || !jsonEl) return;
    updateLineGutter();
    hlEl.innerHTML = buildEditorHighlight(jsonEl.value);
    syncEditorScrollTargets();
  }

  function persistDraftIfValid() {
    const v = getJsonValidation(jsonEl.value);
    if (!v.valid) return;
    try {
      localStorage.setItem(storageKeyDraft(), jsonEl.value);
    } catch (_) {}
  }

  function applyValidationUi() {
    const v = getJsonValidation(jsonEl.value);
    statusEl.textContent = "";
    statusEl.className = "vf-json-status";
    if (v.valid && jsonEl.value.trim()) {
      statusEl.textContent = "Valid JSON";
      statusEl.classList.add("ok");
    } else if (!v.valid) {
      statusEl.textContent = (v.error && v.error.message) || "Invalid JSON";
      statusEl.classList.add("err");
    }

    if (btnRefresh) btnRefresh.disabled = !v.valid;
    if (btnSaveToView) {
      const canSave = v.valid && spTabId != null;
      btnSaveToView.disabled = !canSave;
      if (spTabId == null) {
        btnSaveToView.title =
          "Open View formatter from the extension (popup or { } launcher) while the SharePoint list tab stays open.";
      } else {
        btnSaveToView.title =
          "Write JSON to the current list view in SharePoint (valid JSON only). Uses the list tab you opened this from.";
      }
    }

    if (previewMask) previewMask.hidden = v.valid;

    syncJsonHighlight();
  }

  function showBanner(text) {
    if (!bannerEl) return;
    bannerEl.textContent = text;
    bannerEl.hidden = false;
  }

  function hideBanner() {
    if (!bannerEl) return;
    bannerEl.hidden = true;
    bannerEl.textContent = "";
  }

  function prettyJsonOrEmptyObject(raw) {
    const text = String(raw || "").trim();
    if (!text) return "{}";
    try {
      return JSON.stringify(JSON.parse(text), null, 2);
    } catch (_) {
      return text;
    }
  }

  function sitePathFromPreviewUrl() {
    if (!previewUrl) return "/";
    try {
      return new URL(previewUrl).pathname.replace(/\/$/, "") || "/";
    } catch (_) {
      return "/";
    }
  }

  function webAbsoluteUrlFromPreviewUrl() {
    if (!previewUrl) return "";
    try {
      const u = new URL(previewUrl);
      const segs = u.pathname.split("/").filter(Boolean);
      let siteEnd = -1;
      for (let i = 0; i < segs.length; i++) {
        if ((segs[i] === "sites" || segs[i] === "teams" || segs[i] === "site") && segs[i + 1]) {
          siteEnd = i + 1;
          break;
        }
      }
      if (siteEnd >= 0) return u.origin + "/" + segs.slice(0, siteEnd + 1).join("/");
      return u.origin;
    } catch (_) {
      return "";
    }
  }

  function buildPreviewUrlForView(viewId) {
    if (!previewUrl) return "";
    try {
      const u = new URL(previewUrl);
      if (viewId) u.searchParams.set("View", "{" + normGuid(viewId).toUpperCase() + "}");
      else u.searchParams.delete("View");
      return u.toString();
    } catch (_) {
      return previewUrl;
    }
  }

  async function loadViewFormatterFor(viewId, refreshFrame) {
    if (!viewId || !contextListId || spTabId == null) return;
    const vid = normGuid(viewId);
    const path =
      contextSitePath +
      "/_api/web/lists(guid'" +
      contextListId.replace(/'/g, "''") +
      "')/views(guid'" +
      vid.replace(/'/g, "''") +
      "')?$select=Id,Title,CustomFormatter,DefaultView";
    const res = await sendToSpTab({ action: "rest", method: "GET", path: path });
    if (!res || !res.ok) {
      throw new Error(restErrorMessage(res && res.error) || "Failed to load view formatter JSON.");
    }
    const d = res.data || {};
    selectedViewId = normGuid(d.Id || d.id || vid);
    const rawCf = d.CustomFormatter || d.customFormatter || "";
    let text = prettyJsonOrEmptyObject(rawCf);
    try {
      const draft = localStorage.getItem(storageKeyForView(selectedViewId));
      if (draft && getJsonValidation(draft).valid) {
        let serverKey;
        let draftKey;
        try {
          serverKey = JSON.stringify(JSON.parse(text.trim() || "{}"));
          draftKey = JSON.stringify(JSON.parse(draft.trim()));
        } catch (_) {
          serverKey = text.trim();
          draftKey = draft.trim();
        }
        if (draftKey !== serverKey) text = draft;
      }
    } catch (_) {}
    jsonEl.value = text;
    if (viewSelectEl && selectedViewId) viewSelectEl.value = selectedViewId;
    if (labelEl) {
      const title = String(d.Title || "").trim();
      const src = (refreshFrame ? buildPreviewUrlForView(selectedViewId) : previewUrl) || previewUrl;
      labelEl.textContent = title ? "View: " + title : (src ? src.replace(/^https:\/\//, "") : "");
      labelEl.title = src || "";
    }
    applyValidationUi();
    if (refreshFrame) {
      const nextSrc = previewUrlForSelectedView();
      if (nextSrc) setPreviewFrameSrc(nextSrc);
    }
  }

  function refreshCodeSearchState(selectCurrent) {
    const q = String(codeSearchInput?.value || "");
    const text = String(jsonEl.value || "");
    codeSearchMatches = [];
    codeSearchIndex = -1;
    if (q) {
      const textL = text.toLowerCase();
      const qL = q.toLowerCase();
      let at = 0;
      while (at >= 0) {
        at = textL.indexOf(qL, at);
        if (at < 0) break;
        codeSearchMatches.push({ start: at, end: at + q.length });
        at += q.length || 1;
      }
      if (codeSearchMatches.length) codeSearchIndex = 0;
    }
    if (codeSearchCountEl) {
      codeSearchCountEl.textContent = q
        ? (codeSearchMatches.length ? String(codeSearchIndex + 1) + "/" + String(codeSearchMatches.length) : "0")
        : "";
    }
    if (selectCurrent && codeSearchMatches.length) {
      const m = codeSearchMatches[codeSearchIndex];
      jsonEl.focus();
      jsonEl.setSelectionRange(m.start, m.end);
      syncEditorScrollTargets();
    }
  }

  function moveCodeSearch(dir) {
    if (!codeSearchMatches.length) return;
    codeSearchIndex = (codeSearchIndex + dir + codeSearchMatches.length) % codeSearchMatches.length;
    const m = codeSearchMatches[codeSearchIndex];
    if (codeSearchCountEl) codeSearchCountEl.textContent = String(codeSearchIndex + 1) + "/" + String(codeSearchMatches.length);
    jsonEl.focus();
    jsonEl.setSelectionRange(m.start, m.end);
    syncEditorScrollTargets();
  }

  async function initViewSelectorAndLoadInitialJson() {
    if (!viewSelectEl || spTabId == null) return;
    try {
      const ctx = await sendToSpTab({ action: "getViewFormatContext" });
      if (!ctx || !ctx.ok) throw new Error((ctx && ctx.error) || "Could not read current list/view context.");
      contextListId = normGuid(ctx.listId);
      contextViewId = normGuid(ctx.viewId);
      contextSitePath = String(ctx.sitePath || "/");
      contextWebAbsoluteUrl = webAbsoluteUrlFromPreviewUrl();
      const vRes = await sendToSpTab({
        action: "getViewsData",
        listId: contextListId,
        webAbsoluteUrl: contextWebAbsoluteUrl || undefined,
      });
      if (!vRes || !vRes.ok) throw new Error((vRes && vRes.error) || "Could not load list views.");
      const views = Array.isArray(vRes.views) ? vRes.views : [];
      viewSelectEl.innerHTML = "";
      views.forEach((v) => {
        const id = normGuid(v.id);
        if (!id) return;
        const opt = document.createElement("option");
        opt.value = id;
        opt.textContent = String(v.title || "Untitled") + (v.defaultView ? " (Default)" : "");
        viewSelectEl.appendChild(opt);
      });
      if (!viewSelectEl.options.length) {
        const empty = document.createElement("option");
        empty.value = "";
        empty.textContent = "No views available";
        viewSelectEl.appendChild(empty);
        viewSelectEl.disabled = true;
        showBanner("No list views found to load formatter JSON.");
        return;
      }
      viewSelectEl.disabled = false;
      const initialViewId =
        contextViewId ||
        (views.find((v) => !!v.defaultView && normGuid(v.id)) || {}).id ||
        viewSelectEl.options[0].value;
      await loadViewFormatterFor(initialViewId, true);
      hideBanner();
      scheduleLivePreview(true);
    } catch (e) {
      showBanner((e && e.message) || "Could not initialize view selector.");
    }
  }

  if (labelEl) {
    labelEl.textContent = previewUrl ? previewUrl.replace(/^https:\/\//, "") : "No preview URL — paste a list view URL or open this page from SharePoint.";
    labelEl.title = previewUrl || "";
  }

  if (previewUrl && frame) {
    try {
      setPreviewFrameSrc(previewUrl);
    } catch (_) {
      showBanner("Could not set preview URL.");
    }
  } else {
    showBanner("Open View formatter from the extension (SharePoint list/library page) to load a preview URL.");
  }

  const applySplit = function (pct) {
    const p = Math.min(80, Math.max(22, pct));
    if (previewEl) previewEl.style.flex = "0 0 " + p + "%";
    try {
      localStorage.setItem(STORAGE_SPLIT, String(p));
    } catch (_) {}
  };

  let startPct = 50;
  try {
    const saved = localStorage.getItem(STORAGE_SPLIT);
    if (saved) startPct = parseFloat(saved, 10);
  } catch (_) {}
  if (!isNaN(startPct)) applySplit(startPct);

  if (resizer && splitEl && previewEl) {
    resizer.addEventListener("mousedown", function (e) {
      e.preventDefault();
      resizer.classList.add("active");
      const rect = splitEl.getBoundingClientRect();
      const total = rect.width || 1;
      const previewRect = previewEl.getBoundingClientRect();
      const startX = e.clientX;
      const startW = previewRect.width;

      const overlay = document.createElement("div");
      overlay.className = "vf-resize-overlay";
      overlay.setAttribute("aria-hidden", "true");
      document.body.appendChild(overlay);

      let dragDone = false;
      function move(ev) {
        const dx = ev.clientX - startX;
        const nextW = startW + dx;
        applySplit((nextW / total) * 100);
      }
      function up() {
        if (dragDone) return;
        dragDone = true;
        resizer.classList.remove("active");
        overlay.remove();
        document.removeEventListener("mousemove", move, true);
        document.removeEventListener("mouseup", up, true);
        window.removeEventListener("blur", up);
      }
      overlay.addEventListener("mousemove", move);
      overlay.addEventListener("mouseup", function (ev) {
        ev.preventDefault();
        up();
      });
      document.addEventListener("mousemove", move, true);
      document.addEventListener("mouseup", up, true);
      window.addEventListener("blur", up);
    });
  }

  if (spTabId == null) {
    try {
      const draft = localStorage.getItem(storageKeyDraft());
      jsonEl.value = draft || "{}";
    } catch (_) {
      jsonEl.value = "{}";
    }
  } else {
    jsonEl.value = "{}";
    if (statusEl) {
      statusEl.textContent = "Loading view JSON from SharePoint…";
      statusEl.className = "vf-json-status";
    }
  }

  applyValidationUi();

  jsonEl.addEventListener("scroll", function () {
    syncEditorScrollTargets();
  });

  let saveTimer = null;
  jsonEl.addEventListener("input", function () {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      if (getJsonValidation(jsonEl.value).valid) persistDraftIfValid();
    }, 400);
    applyValidationUi();
    if (codeSearchInput && codeSearchInput.value) refreshCodeSearchState(false);
    if (getJsonValidation(jsonEl.value).valid) scheduleLivePreview(false);
  });

  document.getElementById("darkModeToggle")?.addEventListener("click", function () {
    const isDark = document.documentElement.classList.toggle("dark-mode");
    updateDarkToggleLabel(isDark);
    try {
      chrome.storage.local.set({ darkMode: isDark });
    } catch (_) {}
  });

  document.getElementById("vfRefreshFrame")?.addEventListener("click", function () {
    const url = previewUrlForSelectedView();
    if (!url || !frame || btnRefresh.disabled) return;
    setPreviewFrameSrc(url);
  });

  document.getElementById("vfOpenTab")?.addEventListener("click", function () {
    if (!previewUrl) return;
    window.open(buildPreviewUrlForView(selectedViewId || contextViewId) || previewUrl, "_blank", "noopener,noreferrer");
  });

  document.getElementById("vfFormatJson")?.addEventListener("click", function () {
    const t = jsonEl.value.trim();
    if (!t) {
      applyValidationUi();
      return;
    }
    try {
      jsonEl.value = JSON.stringify(JSON.parse(t), null, 2);
      applyValidationUi();
      persistDraftIfValid();
      scheduleLivePreview(false);
    } catch (_e) {
      applyValidationUi();
    }
  });

  document.getElementById("vfCopyJson")?.addEventListener("click", function () {
    navigator.clipboard.writeText(jsonEl.value).then(function () {
      statusEl.textContent = "Copied";
      statusEl.className = "vf-json-status ok";
      setTimeout(applyValidationUi, 1200);
    });
  });

  document.getElementById("vfSaveToView")?.addEventListener("click", function () {
    if (!btnSaveToView || btnSaveToView.disabled) return;
    (async function () {
      const v = getJsonValidation(jsonEl.value);
      if (!v.valid) return;
      let canonical;
      try {
        canonical = JSON.stringify(JSON.parse(jsonEl.value.trim()));
      } catch (_) {
        return;
      }
      btnSaveToView.disabled = true;
      statusEl.textContent = "Saving to view…";
      statusEl.className = "vf-json-status";
      let doneMsg = "";
      let doneClass = "";
      try {
        const ctx = await sendToSpTab({ action: "getViewFormatContext" });
        if (!ctx || !ctx.ok) throw new Error((ctx && ctx.error) || "Could not read list from the SharePoint tab.");
        const listId = normGuid(ctx.listId || contextListId);
        let viewId = normGuid(selectedViewId || ctx.viewId || contextViewId);
        const sitePath = String(ctx.sitePath || contextSitePath || "/");
        if (!viewId) {
          const defPath =
            sitePath +
            "/_api/web/lists(guid'" +
            listId.replace(/'/g, "''") +
            "')/DefaultView?$select=Id";
          const defRes = await sendToSpTab({ action: "rest", method: "GET", path: defPath });
          if (!defRes || !defRes.ok) {
            throw new Error(
              restErrorMessage(defRes && defRes.error) ||
                "Could not resolve view id. Open the list view you want to format, or use a URL that includes View=."
            );
          }
          const d = defRes.data || {};
          viewId = normGuid(d.Id || d.id);
        }
        if (!viewId) throw new Error("Could not determine view id.");
        const patchPath =
          sitePath +
          "/_api/web/lists(guid'" +
          listId.replace(/'/g, "''") +
          "')/views(guid'" +
          viewId.replace(/'/g, "''") +
          "')";
        const patchRes = await sendToSpTab({
          action: "rest",
          method: "PATCH",
          path: patchPath,
          body: { CustomFormatter: canonical },
        });
        if (!patchRes || !patchRes.ok) {
          throw new Error(restErrorMessage(patchRes && patchRes.error) || "SharePoint rejected the update.");
        }
        selectedViewId = viewId;
        persistDraftIfValid();
        const afterSave = buildPreviewUrlForView(viewId) || previewUrl;
        if (afterSave) setPreviewFrameSrc(afterSave);
        doneMsg = "Saved to SharePoint view";
        doneClass = "ok";
      } catch (e) {
        doneMsg =
          e && e.message === "NO_TAB"
            ? "Open View formatter from the toolkit popup or { } launcher with the list page open."
            : e && e.message
              ? e.message
              : String(e);
        doneClass = "err";
      }
      applyValidationUi();
      if (doneMsg) {
        statusEl.textContent = doneMsg;
        statusEl.className = "vf-json-status " + doneClass;
      }
    })();
  });

  document.getElementById("vfStarterTemplate")?.addEventListener("click", function () {
    jsonEl.value = DEFAULT_TEMPLATE;
    applyValidationUi();
    persistDraftIfValid();
    scheduleLivePreview(true);
  });

  chrome.storage.local.get("darkMode", function (r) {
    const isDark = !!(r && r.darkMode);
    document.documentElement.classList.toggle("dark-mode", isDark);
    updateDarkToggleLabel(isDark);
    syncJsonHighlight();
  });

  if (frame && previewUrl) {
    frame.addEventListener("error", function () {
      showBanner("Preview failed to load. Use Open in tab, or reload the extension and this page if framing rules did not apply.");
    });
  }

  viewSelectEl?.addEventListener("change", function () {
    const nextId = normGuid(viewSelectEl.value);
    if (!nextId) return;
    statusEl.textContent = "Loading selected view…";
    statusEl.className = "vf-json-status";
    loadViewFormatterFor(nextId, true).catch((e) => {
      statusEl.textContent = (e && e.message) || "Could not load selected view.";
      statusEl.className = "vf-json-status err";
    });
  });

  btnOpenViewManager?.addEventListener("click", function () {
    if (spTabId == null) {
      showBanner("Open this formatter from a list/library tab to create and manage views.");
      return;
    }
    const url = new URL(chrome.runtime.getURL("views.html"));
    url.searchParams.set("tabId", String(spTabId));
    if (contextListId) url.searchParams.set("listId", contextListId);
    if (contextWebAbsoluteUrl) url.searchParams.set("webUrl", contextWebAbsoluteUrl);
    url.searchParams.set("create", "1");
    chrome.tabs.create({ url: url.toString() });
  });

  codeSearchInput?.addEventListener("input", () => refreshCodeSearchState(false));
  codeSearchInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      if (!codeSearchMatches.length) refreshCodeSearchState(true);
      else moveCodeSearch(e.shiftKey ? -1 : 1);
    }
  });
  codeSearchPrevBtn?.addEventListener("click", () => {
    if (!codeSearchMatches.length) refreshCodeSearchState(true);
    else moveCodeSearch(-1);
  });
  codeSearchNextBtn?.addEventListener("click", () => {
    if (!codeSearchMatches.length) refreshCodeSearchState(true);
    else moveCodeSearch(1);
  });

  void initViewSelectorAndLoadInitialJson();

})();
