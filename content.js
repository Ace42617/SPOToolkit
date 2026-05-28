// Runs on SharePoint pages. Injects scripts into page context so fetch uses session.

const PROGRESS_BOX_ID = "sp-csv-export-progress";
const SCRIPT_TIMEOUT = 15000;
/** JSON script node read by getViewsData.js (page context). */
const SP_VIEWS_PARAMS_SCRIPT_ID = "sp-views-params";

function getSpViewsParamsScriptId(requestId) {
  return requestId ? SP_VIEWS_PARAMS_SCRIPT_ID + "-" + requestId : SP_VIEWS_PARAMS_SCRIPT_ID;
}

function createSpViewsRequestId() {
  return "views-" + Date.now() + "-" + Math.random().toString(36).slice(2);
}

/** Injects/replaces #sp-views-params for View Manager REST (listId, optional viewId, webAbsoluteUrl). */
function attachSpViewsParamsScript(message, requestId) {
  const payload = {
    requestId: requestId || null,
    viewId: message.viewId || null,
    listId: message.listId || null,
    webAbsoluteUrl: message.webAbsoluteUrl || null,
  };
  let el = document.getElementById(getSpViewsParamsScriptId(requestId));
  if (el) el.remove();
  el = document.createElement("script");
  el.id = getSpViewsParamsScriptId(requestId);
  el.type = "application/json";
  el.textContent = JSON.stringify(payload);
  (document.head || document.documentElement).appendChild(el);
}

function showProgress(message) {
  let el = document.getElementById(PROGRESS_BOX_ID);
  if (!el) {
    el = document.createElement("div");
    el.id = PROGRESS_BOX_ID;
    el.style.cssText =
      "position:fixed;bottom:20px;right:20px;z-index:2147483647;font-family:'Segoe UI',sans-serif;min-width:280px;max-width:90vw;";
    const wrap = document.createElement("div");
    wrap.className = "sp-csv-progress-wrap";
    wrap.innerHTML =
      '<div class="sp-csv-progress-rotate"></div>' +
      '<div class="sp-csv-progress-mask"></div>' +
      '<div class="sp-csv-progress-content">' +
      '<div class="sp-csv-progress-spinner"></div>' +
      '<span class="sp-csv-msg">' + (message || "Starting…") + "</span>" +
      "</div>";
    el.appendChild(wrap);
    const style = document.createElement("style");
    style.textContent =
      ".sp-csv-progress-wrap{position:relative;border-radius:11px;overflow:hidden;}" +
      ".sp-csv-progress-rotate{position:absolute;left:-50%;top:-50%;width:200%;height:200%;background:conic-gradient(transparent,#37ae1c 12%,#3D4FD6 28%,transparent 42%);animation:spcsvrotate 4s linear infinite;z-index:0;}" +
      ".sp-csv-progress-wrap.dark-mode .sp-csv-progress-rotate{background:conic-gradient(transparent,#86efac 12%,#7B89F5 28%,transparent 42%);}" +
      ".sp-csv-progress-mask{position:absolute;left:6px;top:6px;width:calc(100% - 12px);height:calc(100% - 12px);background:#323130;border-radius:5px;z-index:1;}" +
      ".sp-csv-progress-content{position:relative;z-index:2;display:flex;align-items:center;gap:12px;padding:14px 20px;color:#fff;font-size:14px;}" +
      ".sp-csv-progress-spinner{width:20px;height:20px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;border-radius:50%;animation:spcsvspin 0.8s linear infinite;flex-shrink:0;}" +
      "@keyframes spcsvrotate{to{transform:rotate(1turn);}}" +
      "@keyframes spcsvspin{to{transform:rotate(360deg);}}";
    document.head.appendChild(style);
    document.body.appendChild(el);
    chrome.storage.local.get("darkMode", (r) => {
      if (r.darkMode && wrap) wrap.classList.add("dark-mode");
    });
  }
  const msgEl = el.querySelector(".sp-csv-msg");
  if (msgEl) msgEl.textContent = message || "…";
}

function finishProgress(success, message, stopReason) {
  const el = document.getElementById(PROGRESS_BOX_ID);
  if (!el) return;
  const rotate = el.querySelector(".sp-csv-progress-rotate");
  const mask = el.querySelector(".sp-csv-progress-mask");
  const msgEl = el.querySelector(".sp-csv-msg");
  const spinner = el.querySelector(".sp-csv-progress-spinner");
  if (rotate) {
    rotate.style.animation = "none";
    rotate.style.background = success ? "#107c10" : "#a4262c";
  }
  if (mask) mask.style.background = success ? "#107c10" : "#a4262c";
  if (spinner) spinner.style.animation = "none";
  if (msgEl) msgEl.textContent = message || (success ? "Done!" : "Error.");
  const hasEarlyStop = stopReason && /no View ID|returned 0 rows/.test(stopReason);
  const displayMs = (success && stopReason) ? 15000 : (success ? 3500 : 10000);
  if (hasEarlyStop) alert("SharePoint CSV Export: " + stopReason);
  setTimeout(() => {
    const e = document.getElementById(PROGRESS_BOX_ID);
    if (e && e.parentNode) e.parentNode.removeChild(e);
  }, displayMs);
}

const FACT_TOAST_ID = "sp-fact-toast";
const FACT_TOAST_STYLE_ID = "sp-fact-toast-styles";
const FACT_TOAST_DURATION_MS = 45000;
let factToastTimeoutId = null;

const FACT_TOAST_CSS =
  ".sp-fact-wrap{position:relative;background:#323130;color:#fff;border-radius:11px;padding:14px 36px 14px 16px;box-shadow:0 4px 24px rgba(0,0,0,.5);font-size:13px;line-height:1.5;}" +
  ".sp-fact-text{margin:0 0 8px;}" +
  ".sp-fact-source{display:inline-block;padding:6px 12px;background:rgba(255,255,255,.2);color:#fff;font-size:12px;font-weight:600;text-decoration:none;border-radius:6px;}" +
  ".sp-fact-source:hover{background:rgba(255,255,255,.3);}" +
  ".sp-fact-close{position:absolute;top:8px;right:8px;width:24px;height:24px;border:none;background:rgba(255,255,255,.15);color:#fff;border-radius:4px;font-size:18px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;}" +
  ".sp-fact-close:hover{background:rgba(255,255,255,.25);}";

function showFactToast(factText, sourceUrl) {
  if (factToastTimeoutId) {
    clearTimeout(factToastTimeoutId);
    factToastTimeoutId = null;
  }
  let el = document.getElementById(FACT_TOAST_ID);
  if (el) el.remove();
  if (!document.getElementById(FACT_TOAST_STYLE_ID)) {
    const style = document.createElement("style");
    style.id = FACT_TOAST_STYLE_ID;
    style.textContent = FACT_TOAST_CSS;
    document.head.appendChild(style);
  }
  el = document.createElement("div");
  el.id = FACT_TOAST_ID;
  el.style.cssText = "position:fixed;bottom:20px;right:20px;z-index:2147483647;font-family:'Segoe UI',sans-serif;max-width:320px;min-width:200px;";
  el.innerHTML =
    '<div class="sp-fact-wrap">' +
    '<button type="button" class="sp-fact-close" aria-label="Close">×</button>' +
    '<p class="sp-fact-text"></p>' +
    (sourceUrl ? '<a class="sp-fact-source" href="#" target="_blank" rel="noopener">Learn More</a>' : '') +
    '</div>';
  el.querySelector(".sp-fact-text").textContent = factText || "";
  if (sourceUrl) {
    const link = el.querySelector(".sp-fact-source");
    link.href = sourceUrl;
    link.addEventListener("click", (e) => { e.preventDefault(); window.open(sourceUrl, "_blank", "noopener"); });
  }
  document.body.appendChild(el);
  const close = () => {
    factToastTimeoutId = null;
    const node = document.getElementById(FACT_TOAST_ID);
    if (node && node.parentNode) node.parentNode.removeChild(node);
  };
  el.querySelector(".sp-fact-close").addEventListener("click", close);
  factToastTimeoutId = setTimeout(close, FACT_TOAST_DURATION_MS);
}

// --- Lists/Libraries launcher (floating button, bottom-right) ---
const LISTS_LAUNCHER_ID = "sp-toolkit-lists-launcher";
const LISTS_LAUNCHER_PANEL_ID = "sp-toolkit-lists-panel";
const LISTS_LAUNCHER_STYLE_ID = "sp-toolkit-lists-launcher-style";

const LISTS_LAUNCHER_CSS =
  "#" + LISTS_LAUNCHER_ID + "{position:fixed;bottom:20px;right:20px;z-index:2147483646;font-family:'Segoe UI',sans-serif;transition:opacity .15s ease-out;display:block;width:58px;flex-shrink:0;box-sizing:border-box;}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-rolled-up{opacity:0;pointer-events:none;}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn{width:58px!important;height:58px!important;min-width:58px!important;min-height:58px!important;max-width:58px!important;max-height:58px!important;flex:0 0 58px!important;border-radius:50%;border:2px solid rgba(61,79,214,.45);cursor:pointer;display:inline-flex!important;align-items:center;justify-content:center;padding:0!important;box-sizing:border-box!important;overflow:hidden;" +
  "background:linear-gradient(135deg,#f7f9ff 0%,#e7ecff 100%);color:#1a1d24;box-shadow:0 6px 20px rgba(0,0,0,.2);transition:transform .15s,box-shadow .15s,border-color .15s;}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn:hover{transform:scale(1.05);box-shadow:0 6px 20px rgba(0,0,0,.25);}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn:hover{border-color:rgba(61,79,214,.75);}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-dark .sp-toolkit-lists-btn{background:linear-gradient(135deg,#3d4048 0%,#2d3038 100%);border:2px solid rgba(134,239,172,.55);box-shadow:0 4px 16px rgba(0,0,0,.35);}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-dark .sp-toolkit-lists-btn:hover{box-shadow:0 6px 22px rgba(0,0,0,.5);}" +
  "#" + LISTS_LAUNCHER_ID + ".sp-toolkit-launcher-dark .sp-toolkit-lists-btn .sp-toolkit-launcher-compass{filter:brightness(0) invert(1);}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn .sp-toolkit-launcher-compass,#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn img.sp-toolkit-launcher-compass{width:30px!important;height:30px!important;min-width:30px!important;min-height:30px!important;max-width:30px!important;max-height:30px!important;display:block;flex-shrink:0;object-fit:contain;transition:transform .25s ease;}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn:hover .sp-toolkit-launcher-compass{animation:sp-toolkit-launcher-spin .7s ease-in-out 1;}" +
  "#" + LISTS_LAUNCHER_ID + " .sp-toolkit-lists-btn.sp-toolkit-launcher-click-spin .sp-toolkit-launcher-compass{animation:sp-toolkit-launcher-spin-click .55s ease-out 1 forwards;}" +
  "@keyframes sp-toolkit-launcher-spin{to{transform:rotate(360deg);}}" +
  "@keyframes sp-toolkit-launcher-spin-click{to{transform:rotate(1080deg);}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + "{position:absolute;bottom:68px;right:0;width:max-content;min-width:360px;max-width:min(92vw, 1000px);max-height:70vh;overflow:hidden;display:flex;flex-direction:column;border-radius:11px;box-shadow:0 8px 32px rgba(0,0,0,.25);transition:background .2s,color .2s,border-color .2s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark{background:#181b23!important;color:#fff!important;border:1px solid rgba(255,255,255,.08);color-scheme:dark;isolation:isolate;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light{background:#fff;color:#1C1F4A;border:1px solid #E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-header{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.12);flex-shrink:0;transition:border-color .2s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-lists-panel-header{border-bottom-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-left{display:flex;flex-direction:column;align-items:flex-start;gap:2px;min-width:0;flex:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tenant-line{font-size:11px;opacity:.7;color:inherit;margin-top:0;line-height:1.3;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:left;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tabs{display:flex;align-items:center;gap:2px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab{padding:4px 10px;font-size:12px;font-weight:500;border-radius:6px;background:transparent;border:none;color:rgba(255,255,255,.6);cursor:pointer;font-family:inherit;transition:background .15s,color .15s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.sp-toolkit-header-tab-label{color:rgba(255,255,255,.9);cursor:default;pointer-events:none;font-size:14px;padding-left:0;text-align:left;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.sp-toolkit-header-tab-label .sp-toolkit-label-site-name{font-weight:700;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.sp-toolkit-header-tab-label.active{background:transparent;color:rgba(255,255,255,.9);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-header-tab{color:#5A5F8A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-header-tab.sp-toolkit-header-tab-label{color:#1C1F4A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-header-tab.sp-toolkit-header-tab-label.active{background:transparent;color:#1C1F4A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab:hover{color:inherit;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.sp-toolkit-header-tab-label:hover{color:inherit;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-header-tab.active{background:rgba(55,174,28,.2);color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-header-tab.active{background:#EEF0FE;color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-subsites-view{display:none;overflow-y:auto;flex:1;min-height:0;padding:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-subsites-view.visible{display:block;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-icons{display:flex;align-items:center;gap:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-icons a{display:flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:6px;color:rgba(255,255,255,.85);transition:background .15s,color .15s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-lists-panel-icons a:hover{background:rgba(55,174,28,.25);color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-lists-panel-icons a{color:#5A5F8A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-lists-panel-icons a:hover{background:#F5F6FF;color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-icons a svg{width:18px;height:18px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-panel-body{display:flex;flex-direction:column;flex:1;min-height:0;overflow:hidden;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-wrap{overflow-y:auto;flex:1;min-height:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-subsites-block{margin-top:12px;padding-top:12px;border-top:1px solid rgba(255,255,255,.12);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-subsites-block{border-top-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-subsites-heading{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:rgba(255,255,255,.6);margin-bottom:8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-subsites-heading{color:#5A5F8A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-subsites-list{display:flex;flex-wrap:wrap;gap:6px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-subsites-list a{display:inline-block;font-size:12px;padding:4px 10px;border-radius:6px;background:rgba(55,174,28,.15);color:#86efac;text-decoration:none;transition:background .15s,color .15s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-subsites-list a:hover{background:rgba(55,174,28,.3);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-subsites-list a{background:#EEF0FE;color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-subsites-list a:hover{background:#E0E4FF;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table{width:max-content;border-collapse:collapse;font-size:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table th{color:rgba(255,255,255,.92)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table td{color:rgba(255,255,255,.92)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-name a{color:rgba(255,255,255,.95)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-icon{color:rgba(255,255,255,.88)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-settings a{color:rgba(255,255,255,.85)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-view-manager{color:rgba(255,255,255,.85)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-header-tab-label," +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-header-tab-label .sp-toolkit-label-site-name{color:rgba(255,255,255,.92)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-header-tenant-line{color:rgba(255,255,255,.75)!important;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th{text-align:left;padding:8px 10px;font-weight:600;border-bottom:1px solid rgba(255,255,255,.12);color:inherit;white-space:nowrap;cursor:pointer;user-select:none;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th:first-child," + "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td:first-child{padding-left:12px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th:hover{opacity:.9;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th .sp-toolkit-sort-icon{opacity:.6;margin-left:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table th{border-bottom-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row th{padding:4px 8px;vertical-align:middle;cursor:default;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row th:hover{opacity:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row input," +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row select{width:100%;max-width:140px;padding:6px 8px;font-size:12px;border-radius:6px;border:1px solid rgba(255,255,255,.15);background:#252a30;color:rgba(255,255,255,.95);font-family:inherit;cursor:pointer;appearance:auto;-webkit-appearance:menulist;-moz-appearance:menulist;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row input{background:#252a30;color:rgba(255,255,255,.95);cursor:text;appearance:none;-webkit-appearance:none;-moz-appearance:none;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row select:hover{border-color:rgba(255,255,255,.25);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-filter-row select:focus{outline:none;border-color:rgba(55,174,28,.5);box-shadow:0 0 0 2px rgba(55,174,28,.15);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row input," +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row select{background:#f5f6ff;border-color:#E4E6F5;color:#1C1F4A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row input{background:#fff;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row select:hover{border-color:#CDD0EE;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-filter-row select:focus{border-color:#37ae1c;box-shadow:0 0 0 2px rgba(55,174,28,.2);}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td{padding:8px 10px;border-bottom:1px solid rgba(255,255,255,.06);vertical-align:middle;white-space:nowrap;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table td{border-bottom-color:#E4E6F5;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-name{display:flex;align-items:center;gap:8px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-icon{flex-shrink:0;width:20px;height:20px;color:inherit;opacity:.85;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-name a{color:inherit;text-decoration:none;white-space:nowrap;min-width:0;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-contents-table .sp-toolkit-row-name a:hover{color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-contents-table .sp-toolkit-row-name a:hover{color:#37ae1c;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table tr.sp-toolkit-current .sp-toolkit-row-name a{font-weight:600;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table tr.sp-toolkit-current td:first-child{border-left:3px solid #37ae1c;padding-left:9px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td.sp-toolkit-type-cell{display:flex;align-items:center;color:inherit;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-type-pill{display:inline-flex;align-items:center;justify-content:center;line-height:1;font-size:11px;font-weight:700;padding:3px 8px;border-radius:999px;white-space:nowrap;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-dark .sp-toolkit-type-pill{background:rgba(55,174,28,.25);color:#86efac;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-type-pill{background:#EEF0FE;color:#5A5F8A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td:nth-child(3),#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table td:nth-child(4){text-align:center;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-settings{display:flex;align-items:center;justify-content:center;gap:2px;min-width:56px;text-align:center;vertical-align:middle;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-settings a{display:inline-flex;align-items:center;justify-content:center;color:inherit;opacity:.7;padding:4px;border-radius:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-settings a:hover{opacity:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-settings svg{width:18px;height:18px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-view-manager{display:inline-flex;align-items:center;justify-content:center;color:inherit;opacity:.7;padding:4px;margin:0;border:none;background:transparent;cursor:pointer;border-radius:4px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-view-manager:hover{opacity:1;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-view-manager svg{width:18px;height:18px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table th[data-column=\"settings\"]{cursor:default;min-width:56px;width:64px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table .sp-toolkit-row-name a:hover{text-decoration:underline;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading,.sp-toolkit-lists-error{padding:14px;font-size:13px;color:#9ca3bf;transition:color .2s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-light .sp-toolkit-lists-loading,.sp-toolkit-lists-panel-light .sp-toolkit-lists-error{color:#5A5F8A;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-lists-panel-open{animation:sp-toolkit-panel-fade-in .2s ease-out forwards;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + ".sp-toolkit-panel-rolling-up{animation:sp-toolkit-panel-fade-out .12s ease-out forwards;}" +
  "@keyframes sp-toolkit-panel-fade-in{0%{opacity:0;}100%{opacity:1;}}" +
  "@keyframes sp-toolkit-panel-fade-out{0%{opacity:1;}100%{opacity:0;}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading{display:flex;align-items:center;gap:10px;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading-dot{width:6px;height:6px;border-radius:50%;background:rgba(55,174,28,.8);animation:sp-toolkit-load-bounce 1s ease-in-out infinite;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading-dot:nth-child(2){animation-delay:.12s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-lists-loading-dot:nth-child(3){animation-delay:.24s;}" +
  "@keyframes sp-toolkit-load-bounce{0%,80%,100%{transform:scale(0.6);opacity:.6;}40%{transform:scale(1);opacity:1;}}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr{animation:sp-toolkit-item-in .3s ease-out backwards;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(1){animation-delay:0.02s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(2){animation-delay:0.04s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(3){animation-delay:0.06s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(4){animation-delay:0.08s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(5){animation-delay:0.1s;}" +
  "#" + LISTS_LAUNCHER_PANEL_ID + " .sp-toolkit-contents-table.sp-toolkit-lists-loaded tbody tr:nth-child(n+6){animation-delay:0.12s;}" +
  "@keyframes sp-toolkit-item-in{0%{opacity:0;transform:translateX(-6px);}100%{opacity:1;transform:translateX(0);}}";

let listsLauncherExpanded = false;
let listsLauncherListsListener = null;
let listsLauncherSubsitesListener = null;

function toggleListsLauncher() {
  const container = document.getElementById(LISTS_LAUNCHER_ID);
  const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
  if (!container || !panel) return;
  listsLauncherExpanded = !listsLauncherExpanded;
  if (listsLauncherExpanded) {
    panel.style.display = "flex";
    panel.classList.remove("sp-toolkit-panel-rolling-up");
    panel.classList.add("sp-toolkit-lists-panel-open");
    try { chrome.runtime.sendMessage({ type: "SPOToolkitPanelOpened" }); } catch (_) {}
  } else {
    panel.classList.remove("sp-toolkit-lists-panel-open");
    panel.classList.add("sp-toolkit-panel-rolling-up");
    setTimeout(function () {
      panel.style.display = "none";
      panel.classList.remove("sp-toolkit-panel-rolling-up");
    }, 130);
  }
  if (listsLauncherExpanded) {
    const loadingEl = panel.querySelector(".sp-toolkit-lists-loading");
    const errorEl = panel.querySelector(".sp-toolkit-lists-error");
    const tableEl = panel.querySelector(".sp-toolkit-contents-table");
    const contentsWrap = panel.querySelector(".sp-toolkit-contents-wrap");
    if (loadingEl) loadingEl.style.display = "flex";
    if (errorEl) errorEl.style.display = "none";
    if (contentsWrap) contentsWrap.style.display = "block";
    if (tableEl) {
      tableEl.classList.remove("sp-toolkit-lists-loaded");
      tableEl.style.display = "none";
      const tbody = tableEl.querySelector("tbody");
      if (tbody) tbody.innerHTML = "";
    }
    const prevSubsites = contentsWrap ? contentsWrap.querySelector(".sp-toolkit-subsites-block") : null;
    if (prevSubsites) prevSubsites.remove();
    const headerTabs = panel.querySelector(".sp-toolkit-header-tabs");
    if (headerTabs) headerTabs.innerHTML = "";
    const subsitesView = panel.querySelector(".sp-toolkit-subsites-view");
    if (subsitesView) { subsitesView.classList.remove("visible"); subsitesView.innerHTML = ""; }
    if (listsLauncherListsListener) {
      window.removeEventListener("message", listsLauncherListsListener);
      listsLauncherListsListener = null;
    }
    if (listsLauncherSubsitesListener) {
      window.removeEventListener("message", listsLauncherSubsitesListener);
      listsLauncherSubsitesListener = null;
    }
    listsLauncherListsListener = (ev) => {
      if (!ev.data || ev.data.__spcsv !== true || ev.data.type !== "SPCSVSiteListsResult") return;
      window.removeEventListener("message", listsLauncherListsListener);
      listsLauncherListsListener = null;
      const tableEl = panel.querySelector(".sp-toolkit-contents-table");
      const tbody = tableEl ? tableEl.querySelector("tbody") : null;
      const loadingEl = panel.querySelector(".sp-toolkit-lists-loading");
      const errorEl = panel.querySelector(".sp-toolkit-lists-error");
      const contentsWrap = panel.querySelector(".sp-toolkit-contents-wrap");
      if (loadingEl) loadingEl.style.display = "none";
      if (contentsWrap) contentsWrap.style.display = "block";
      if (ev.data.error) {
        if (errorEl) {
          errorEl.textContent = ev.data.error;
          errorEl.style.display = "block";
        }
        return;
      }
      if (errorEl) errorEl.style.display = "none";
      if (tableEl && !ev.data.error) {
        tableEl.classList.add("sp-toolkit-lists-loaded");
        tableEl.style.display = "table";
      }
      const siteUrl = (ev.data.siteUrl || "").replace(/\/$/, "");
      /* Fluent Key 24 regular — Fabric MDL2 "Permissions" / Unicode E8D7 (flicon.io) */
      const fluentPermissionsKeySvgInner = "<path fill=\"currentColor\" d=\"M18.2493 6.99982C18.2493 7.69017 17.6896 8.24982 16.9993 8.24982C16.3089 8.24982 15.7493 7.69017 15.7493 6.99982C15.7493 6.30946 16.3089 5.74982 16.9993 5.74982C17.6896 5.74982 18.2493 6.30946 18.2493 6.99982ZM15.4992 2.0498C11.885 2.0498 8.94922 4.98559 8.94922 8.5998C8.94922 8.98701 8.99939 9.36017 9.05968 9.70382C9.07749 9.80529 9.04493 9.89344 8.99046 9.94791L2.75467 16.1837C2.23895 16.6994 1.94922 17.3989 1.94922 18.1282V20.2998C1.94922 21.2663 2.73272 22.0498 3.69922 22.0498H6.19922C7.16572 22.0498 7.94922 21.2663 7.94922 20.2998V19.0498H9.69922C10.3896 19.0498 10.9492 18.4902 10.9492 17.7998V16.0498H12.6992C13.3741 16.0498 13.9241 15.515 13.9484 14.846C14.4451 14.9738 14.9689 15.0498 15.4992 15.0498C19.1134 15.0498 22.0492 12.114 22.0492 8.4998C22.0492 4.86866 19.0963 2.0498 15.4992 2.0498ZM10.4492 8.5998C10.4492 5.81402 12.7134 3.5498 15.4992 3.5498C18.3021 3.5498 20.5492 5.73095 20.5492 8.4998C20.5492 11.2856 18.285 13.5498 15.4992 13.5498C14.8199 13.5498 14.1206 13.3787 13.4947 13.1104C13.2629 13.0111 12.9968 13.0349 12.7864 13.1737C12.5759 13.3125 12.4492 13.5477 12.4492 13.7998V14.5498H10.6992C10.0089 14.5498 9.44922 15.1094 9.44922 15.7998V17.5498H7.69922C7.00886 17.5498 6.44922 18.1094 6.44922 18.7998V20.2998C6.44922 20.4379 6.33729 20.5498 6.19922 20.5498H3.69922C3.56115 20.5498 3.44922 20.4379 3.44922 20.2998V18.1282C3.44922 17.7967 3.58091 17.4788 3.81534 17.2443L10.0511 11.0086C10.4695 10.5902 10.6349 10.0018 10.5371 9.44461C10.4834 9.13865 10.4492 8.8622 10.4492 8.5998Z\"/>";
      const lists = Array.isArray(ev.data.lists) ? ev.data.lists : [];
      const currentPath = (location.pathname || "").split("?")[0].replace(/\/$/, "") || "/";
      const onSiteContentsPage = /viewlsts\.aspx/i.test(location.pathname || "");
      let currentListTitle = null;
      for (let i = 0; i < lists.length; i++) {
        const viewPath = lists[i].viewUrl ? (function () { try { return new URL(lists[i].viewUrl).pathname.replace(/\/$/, "") || "/"; } catch (_) { return ""; } }()) : "";
        if (viewPath && (currentPath === viewPath || currentPath.startsWith(viewPath + "/"))) {
          currentListTitle = lists[i].title || "Untitled";
          break;
        }
      }
      const iconsEl = panel.querySelector(".sp-toolkit-lists-panel-icons");
      if (iconsEl && siteUrl) {
        iconsEl.innerHTML = "";
        const addIcon = (href, title, svgPath, filled) => {
          const a = document.createElement("a");
          a.href = href;
          a.target = "_blank";
          a.rel = "noopener noreferrer";
          a.title = title;
          a.setAttribute("aria-label", title);
          if (filled) {
            a.innerHTML = "<svg viewBox=\"0 0 24 24\" fill=\"none\" xmlns=\"http://www.w3.org/2000/svg\">" + svgPath + "</svg>";
          } else {
            a.innerHTML = "<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\">" + svgPath + "</svg>";
          }
          iconsEl.appendChild(a);
        };
        if (!onSiteContentsPage) addIcon(siteUrl + "/_layouts/15/viewlsts.aspx", "Site contents", "<rect x=\"3\" y=\"3\" width=\"7\" height=\"7\" rx=\"1\"/><rect x=\"14\" y=\"3\" width=\"7\" height=\"7\" rx=\"1\"/><rect x=\"3\" y=\"14\" width=\"7\" height=\"7\" rx=\"1\"/><rect x=\"14\" y=\"14\" width=\"7\" height=\"7\" rx=\"1\"/>");
        addIcon(siteUrl + "/_layouts/15/settings.aspx", "Site settings", "<circle cx=\"12\" cy=\"12\" r=\"3\"/><path d=\"M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z\"/>");
        addIcon(siteUrl + "/_layouts/15/user.aspx", "Site permissions", fluentPermissionsKeySvgInner, true);
        const currentListId = ev.data.currentListId && String(ev.data.currentListId).replace(/^\{|\}$/g, "").trim();
        const onListOrLibraryView = !!(currentListId || currentListTitle) && !onSiteContentsPage;
        if (currentListId) {
          const listGuid = "{" + currentListId.toUpperCase() + "}";
          addIcon(siteUrl + "/_layouts/15/listedit.aspx?List=" + encodeURIComponent(listGuid), "Current list/library settings", "<line x1=\"4\" y1=\"21\" x2=\"4\" y2=\"14\"/><line x1=\"4\" y1=\"10\" x2=\"4\" y2=\"3\"/><line x1=\"12\" y1=\"21\" x2=\"12\" y2=\"12\"/><line x1=\"12\" y1=\"8\" x2=\"12\" y2=\"3\"/><line x1=\"20\" y1=\"21\" x2=\"20\" y2=\"16\"/><line x1=\"20\" y1=\"12\" x2=\"20\" y2=\"3\"/><line x1=\"1\" y1=\"14\" x2=\"7\" y2=\"14\"/><line x1=\"9\" y1=\"8\" x2=\"15\" y2=\"8\"/><line x1=\"17\" y1=\"16\" x2=\"23\" y2=\"16\"/>");
        }
        addIcon(siteUrl + "/_layouts/15/RecycleBin.aspx", "Recycle bin", "<path fill=\"currentColor\" d=\"M7.03454 3.5C6.13437 3.5 5.4063 4.17543 5.2982 5.02738C5.36232 5.00954 5.4299 5 5.49971 5H18.4997C18.5695 5 18.6371 5.00954 18.7012 5.02738C18.5931 4.17543 17.8651 3.5 16.9649 3.5H7.03454ZM6.85063 19.8306C6.8918 20.2114 7.21327 20.5 7.59629 20.5H16.4031C16.7861 20.5 17.1076 20.2114 17.1488 19.8306L18.5905 6.49456C18.5608 6.49815 18.5304 6.5 18.4997 6.5H5.49971C5.46897 6.5 5.43866 6.49815 5.4089 6.49456L6.85063 19.8306ZM3.80337 5.59932C3.59559 3.67734 5.10136 2 7.03454 2H16.9649C18.8981 2 20.4038 3.67733 20.196 5.59932L18.6401 19.9918C18.5166 21.1342 17.5522 22 16.4031 22H7.59629C6.44725 22 5.48282 21.1342 5.35932 19.9918L3.80337 5.59932ZM11.7919 10.4094C11.8909 10.2614 12.1085 10.2614 12.2075 10.4094L12.8763 11.4092C13.1067 11.7535 13.5725 11.8459 13.9167 11.6156C14.261 11.3852 14.3534 10.9194 14.1231 10.5752L13.4542 9.57538C12.7612 8.53947 11.2382 8.53946 10.5452 9.57538L9.87634 10.5752C9.64602 10.9194 9.7384 11.3852 10.0827 11.6156C10.427 11.8459 10.8928 11.7535 11.1231 11.4092L11.7919 10.4094ZM9.63558 12.5991C9.99077 12.8122 10.1059 13.2729 9.89283 13.6281L9.5287 14.2349C9.32874 14.5682 9.5688 14.9922 9.95745 14.9922H10.7497C11.1639 14.9922 11.4997 15.328 11.4997 15.7422C11.4997 16.1564 11.1639 16.4922 10.7497 16.4922H9.95745C8.40286 16.4922 7.44263 14.7962 8.24246 13.4632L8.60659 12.8563C8.8197 12.5011 9.2804 12.386 9.63558 12.5991ZM14.1087 13.6276C13.8958 13.2723 14.0113 12.8117 14.3666 12.5988C14.722 12.3859 15.1826 12.5014 15.3954 12.8568L15.7594 13.4644C16.558 14.7974 15.5977 16.4922 14.0437 16.4922H13.2497C12.8355 16.4922 12.4997 16.1564 12.4997 15.7422C12.4997 15.328 12.8355 14.9922 13.2497 14.9922H14.0437C14.4322 14.9922 14.6723 14.5685 14.4727 14.2352L14.1087 13.6276Z\"/>", true);
        if (onListOrLibraryView) {
          const vfA = document.createElement("a");
          vfA.href = "#";
          vfA.title = "View formatter — JSON editor and list preview";
          vfA.setAttribute("aria-label", "View formatter");
          vfA.innerHTML = "<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M8 3H7a2 2 0 0 0-2 2v5a2 2 0 0 1-2 2 2 2 0 0 1 2 2v5c0 1.1.9 2 2 2h1\"/><path d=\"M16 21h1a2 2 0 0 0 2-2v-5c0-1.1.9-2 2-2a2 2 0 0 1-2-2V5a2 2 0 0 0-2-2h-1\"/></svg>";
          vfA.addEventListener("click", function (ev) {
            ev.preventDefault();
            ev.stopPropagation();
            try {
              chrome.runtime.sendMessage({ type: "SPOToolkitOpenViewFormatter", previewUrl: window.location.href });
            } catch (err) {
              console.warn("SPOToolkit: View formatter", err);
            }
          });
          iconsEl.appendChild(vfA);
        }
      }
      function formatModified(s) {
        if (!s) return "";
        try {
          const d = new Date(s);
          if (isNaN(d.getTime())) return String(s);
          const m = d.getMonth() + 1, day = d.getDate(), y = d.getFullYear();
          const h = d.getHours(), min = d.getMinutes(), am = h < 12;
          const h12 = (h % 12) || 12;
          return m + "/" + day + "/" + y + " " + h12 + ":" + (min < 10 ? "0" : "") + min + " " + (am ? "AM" : "PM");
        } catch (_) { return String(s); }
      }
      const iconLibrary = "<svg class=\"sp-toolkit-row-icon\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z\"/><line x1=\"12\" y1=\"11\" x2=\"12\" y2=\"17\"/><line x1=\"9\" y1=\"14\" x2=\"15\" y2=\"14\"/></svg>";
      const iconList = "<svg class=\"sp-toolkit-row-icon\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><rect x=\"3\" y=\"4\" width=\"18\" height=\"18\" rx=\"2\" ry=\"2\"/><line x1=\"16\" y1=\"2\" x2=\"16\" y2=\"6\"/><line x1=\"8\" y1=\"2\" x2=\"8\" y2=\"6\"/><line x1=\"3\" y1=\"10\" x2=\"21\" y2=\"10\"/></svg>";
      const thead = tableEl ? tableEl.querySelector("thead") : null;
      if (thead && tableEl) {
        thead.innerHTML = "";
        const cols = [{ key: "name", label: "Name" }, { key: "type", label: "Type" }, { key: "items", label: "Items" }, { key: "modified", label: "Modified" }, { key: "settings", label: "" }];
        const filterRow = document.createElement("tr");
        filterRow.className = "sp-toolkit-filter-row";
        const filterNameTh = document.createElement("th");
        filterNameTh.setAttribute("data-column", "name");
        const nameInput = document.createElement("input");
        nameInput.type = "text";
        nameInput.placeholder = "Filter…";
        nameInput.setAttribute("aria-label", "Filter by name");
        filterNameTh.appendChild(nameInput);
        filterRow.appendChild(filterNameTh);
        const uniqueTypes = [];
        const typeSet = {};
        lists.forEach(function (item) {
          const tl = item.typeLabel != null ? item.typeLabel : (item.isLibrary ? "Document library" : "List");
          if (!typeSet[tl]) { typeSet[tl] = true; uniqueTypes.push(tl); }
        });
        uniqueTypes.sort();
        const filterTypeTh = document.createElement("th");
        filterTypeTh.setAttribute("data-column", "type");
        const typeSelect = document.createElement("select");
        typeSelect.setAttribute("aria-label", "Filter by type");
        const typeAll = document.createElement("option");
        typeAll.value = "";
        typeAll.textContent = "All types";
        typeSelect.appendChild(typeAll);
        uniqueTypes.forEach(function (t) {
          const opt = document.createElement("option");
          opt.value = t;
          opt.textContent = t;
          typeSelect.appendChild(opt);
        });
        filterTypeTh.appendChild(typeSelect);
        filterRow.appendChild(filterTypeTh);
        filterRow.appendChild(document.createElement("th"));
        filterRow.appendChild(document.createElement("th"));
        filterRow.appendChild(document.createElement("th"));
        thead.appendChild(filterRow);
        const headerRow = document.createElement("tr");
        cols.forEach(function (c) {
          const th = document.createElement("th");
          th.setAttribute("data-column", c.key);
          th.innerHTML = c.key === "settings" ? c.label : c.label + " <span class=\"sp-toolkit-sort-icon\" aria-hidden=\"true\">↕</span>";
          headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        panel._siteContentsLists = lists;
        panel._siteContentsCurrentPath = currentPath;
        panel._siteContentsSort = { col: "type", dir: 1 };
        panel._siteContentsFilter = { name: "", type: "" };
        function renderRows() {
          if (!tbody) return;
          const listData = panel._siteContentsLists || [];
          const sort = panel._siteContentsSort || { col: null, dir: 1 };
          const filter = panel._siteContentsFilter || { name: "", type: "" };
          let filtered = listData.filter(function (item) {
            const title = (item.title || "Untitled").toLowerCase();
            const typeLabel = item.typeLabel != null ? item.typeLabel : (item.isLibrary ? "Document library" : "List");
            if (filter.name && title.indexOf(filter.name.toLowerCase()) === -1) return false;
            if (filter.type && typeLabel !== filter.type) return false;
            return true;
          });
          const col = sort.col;
          const dir = sort.dir;
          if (col === "name") filtered.sort(function (a, b) { const x = (a.title || "").toLowerCase(), y = (b.title || "").toLowerCase(); return dir * (x < y ? -1 : x > y ? 1 : 0); });
          else if (col === "type") filtered.sort(function (a, b) { const x = (a.typeLabel != null ? a.typeLabel : (a.isLibrary ? "Document library" : "List")); const y = (b.typeLabel != null ? b.typeLabel : (b.isLibrary ? "Document library" : "List")); return dir * (x < y ? -1 : x > y ? 1 : 0); });
          else if (col === "items") filtered.sort(function (a, b) { const x = (a.itemCount != null ? a.itemCount : 0) | 0; const y = (b.itemCount != null ? b.itemCount : 0) | 0; return dir * (x - y); });
          else if (col === "modified") filtered.sort(function (a, b) { const x = new Date(a.modified || 0).getTime(); const y = new Date(b.modified || 0).getTime(); return dir * (x - y); });
          tbody.innerHTML = "";
          const path = panel._siteContentsCurrentPath || "";
          filtered.forEach(function (item) {
            const viewPath = item.viewUrl ? (function () { try { return new URL(item.viewUrl).pathname.replace(/\/$/, "") || "/"; } catch (_) { return ""; } }()) : "";
            const isCurrent = viewPath && (path === viewPath || path.startsWith(viewPath + "/"));
            const tr = document.createElement("tr");
            if (isCurrent) tr.classList.add("sp-toolkit-current");
            const nameCell = document.createElement("td");
            const nameDiv = document.createElement("div");
            nameDiv.className = "sp-toolkit-row-name";
            nameDiv.innerHTML = (item.isLibrary ? iconLibrary : iconList) + " ";
            const a = document.createElement("a");
            a.href = item.viewUrl || "#";
            a.textContent = item.title || "Untitled";
            a.addEventListener("click", function (e) { e.preventDefault(); e.stopPropagation(); closeListsLauncherPanel(); if (item.viewUrl) requestAnimationFrame(function () { window.location.href = item.viewUrl; }); });
            nameDiv.appendChild(a);
            nameCell.appendChild(nameDiv);
            tr.appendChild(nameCell);
            const typeLabel = item.typeLabel != null ? item.typeLabel : (item.isLibrary ? "Document library" : "List");
            const typeCell = document.createElement("td");
            typeCell.className = "sp-toolkit-type-cell";
            const pill = document.createElement("span");
            pill.className = "sp-toolkit-type-pill";
            pill.textContent = typeLabel;
            typeCell.appendChild(pill);
            tr.appendChild(typeCell);
            const itemsCell = document.createElement("td");
            itemsCell.textContent = (item.itemCount != null ? item.itemCount : 0).toString();
            tr.appendChild(itemsCell);
            const modCell = document.createElement("td");
            modCell.textContent = formatModified(item.modified);
            tr.appendChild(modCell);
            const settingsCell = document.createElement("td");
            settingsCell.className = "sp-toolkit-row-settings";
            const listGuid = item.id != null ? ("{" + String(item.id).replace(/[{}]/g, "").trim().toUpperCase() + "}") : "";
            if (listGuid && siteUrl) {
              const guidNoBraces = String(item.id).replace(/[{}]/g, "").trim().toUpperCase();
              const vmBtn = document.createElement("button");
              vmBtn.type = "button";
              vmBtn.className = "sp-toolkit-row-view-manager";
              vmBtn.title = "View Manager for this list";
              vmBtn.setAttribute("aria-label", "View Manager for " + (item.title || "list"));
              vmBtn.innerHTML = "<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z\"/><circle cx=\"12\" cy=\"12\" r=\"3\"/></svg>";
              vmBtn.addEventListener("click", function (e) {
                e.preventDefault();
                e.stopPropagation();
                try {
                  chrome.runtime.sendMessage({ type: "SPOToolkitOpenViewManager", listId: guidNoBraces, webUrl: siteUrl }, function (res) {
                    if (chrome.runtime.lastError || (res && res.ok === false)) {
                      console.warn("SPOToolkit: View Manager open failed", chrome.runtime.lastError || res);
                    }
                  });
                } catch (err) {
                  console.warn("SPOToolkit: View Manager open failed", err);
                }
              });
              settingsCell.appendChild(vmBtn);
              const settingsLink = document.createElement("a");
              settingsLink.href = siteUrl + "/_layouts/15/listedit.aspx?List=" + encodeURIComponent(listGuid);
              settingsLink.target = "_blank";
              settingsLink.rel = "noopener";
              settingsLink.title = "List settings";
              settingsLink.setAttribute("aria-label", "Settings for " + (item.title || "list"));
              settingsLink.innerHTML = "<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><line x1=\"4\" y1=\"21\" x2=\"4\" y2=\"14\"/><line x1=\"4\" y1=\"10\" x2=\"4\" y2=\"3\"/><line x1=\"12\" y1=\"21\" x2=\"12\" y2=\"12\"/><line x1=\"12\" y1=\"8\" x2=\"12\" y2=\"3\"/><line x1=\"20\" y1=\"21\" x2=\"20\" y2=\"16\"/><line x1=\"20\" y1=\"12\" x2=\"20\" y2=\"3\"/><line x1=\"1\" y1=\"14\" x2=\"7\" y2=\"14\"/><line x1=\"9\" y1=\"8\" x2=\"15\" y2=\"8\"/><line x1=\"17\" y1=\"16\" x2=\"23\" y2=\"16\"/></svg>";
              settingsCell.appendChild(settingsLink);
              const permsLink = document.createElement("a");
              permsLink.href = siteUrl + "/_layouts/15/user.aspx?obj=" + encodeURIComponent(listGuid + ",list");
              permsLink.target = "_blank";
              permsLink.rel = "noopener";
              permsLink.title = "List/library permissions";
              permsLink.setAttribute("aria-label", "Permissions for " + (item.title || "list"));
              permsLink.innerHTML = "<svg viewBox=\"0 0 24 24\" fill=\"none\" xmlns=\"http://www.w3.org/2000/svg\">" + fluentPermissionsKeySvgInner + "</svg>";
              settingsCell.appendChild(permsLink);
            }
            tr.appendChild(settingsCell);
            tbody.appendChild(tr);
          });
        }
        function updateSortIcons() {
          const sort = panel._siteContentsSort || { col: null, dir: 1 };
          headerRow.querySelectorAll("th[data-column]").forEach(function (th) {
            const key = th.getAttribute("data-column");
            const icon = th.querySelector(".sp-toolkit-sort-icon");
            if (!icon || key === "settings") return;
            if (sort.col !== key) { icon.textContent = "↕"; return; }
            icon.textContent = sort.dir === 1 ? "↑" : "↓";
          });
        }
        headerRow.querySelectorAll("th[data-column]").forEach(function (th) {
          th.addEventListener("click", function () {
            const key = th.getAttribute("data-column");
            if (key === "settings") return;
            const sort = panel._siteContentsSort || { col: null, dir: 1 };
            if (sort.col === key) panel._siteContentsSort = { col: key, dir: -sort.dir };
            else panel._siteContentsSort = { col: key, dir: 1 };
            updateSortIcons();
            renderRows();
          });
        });
        nameInput.addEventListener("input", function () { panel._siteContentsFilter = panel._siteContentsFilter || {}; panel._siteContentsFilter.name = nameInput.value.trim(); renderRows(); });
        typeSelect.addEventListener("change", function () { panel._siteContentsFilter = panel._siteContentsFilter || {}; panel._siteContentsFilter.type = typeSelect.value; renderRows(); });
        renderRows();
        updateSortIcons();
      }
      const headerTabsEl = panel.querySelector(".sp-toolkit-header-tabs");
      const contentsWrapEl = panel.querySelector(".sp-toolkit-contents-wrap");
      const subsitesViewEl = panel.querySelector(".sp-toolkit-subsites-view");
      function escapeHtml(str) {
        const div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
      }
      function siteContentsLabelHtml(tenantName, siteName, listTitle) {
        const t = (tenantName || "").trim();
        const s = (siteName || "").trim();
        const list = (listTitle || "").trim();
        if (s && list) return "<span class=\"sp-toolkit-label-site-name\">" + escapeHtml(s) + "</span>" + " | " + escapeHtml(list) + " | Contents";
        if (s) return "<span class=\"sp-toolkit-label-site-name\">" + escapeHtml(s) + "</span>" + " | Contents";
        if (t) return escapeHtml(t) + " | Site Contents";
        return "Site Contents";
      }
      const labelHtml = siteContentsLabelHtml(ev.data.tenantName, ev.data.siteName, currentListTitle);
      const tenantName = (ev.data.tenantName || "").trim();
      const siteUrlForLabel = (siteUrl || ev.data.siteUrl || "").replace(/\/$/, "");
      const tenantLineText = tenantName && siteUrlForLabel ? tenantName + " - " + siteUrlForLabel : tenantName || siteUrlForLabel || "";
      const headerLeftEl = panel.querySelector(".sp-toolkit-header-left");
      function applyTenantLine(container, text) {
        if (!container) return;
        const lines = container.querySelectorAll(".sp-toolkit-header-tenant-line");
        const first = lines[0];
        if (first) {
          first.textContent = text;
          first.style.display = text ? "" : "none";
          for (let i = 1; i < lines.length; i++) lines[i].remove();
        } else if (text) {
          const el = document.createElement("div");
          el.className = "sp-toolkit-header-tenant-line";
          el.textContent = text;
          el.setAttribute("aria-hidden", "true");
          container.appendChild(el);
        }
      }
      if (headerTabsEl && !headerTabsEl.querySelector(".sp-toolkit-header-tab")) {
        const tabContents = document.createElement("button");
        tabContents.type = "button";
        tabContents.className = "sp-toolkit-header-tab sp-toolkit-header-tab-label active";
        tabContents.innerHTML = labelHtml;
        tabContents.addEventListener("click", function () {
          headerTabsEl.querySelectorAll(".sp-toolkit-header-tab").forEach(function (t) { t.classList.remove("active"); });
          tabContents.classList.add("active");
          if (contentsWrapEl) contentsWrapEl.style.display = "block";
          if (subsitesViewEl) subsitesViewEl.classList.remove("visible");
        });
        headerTabsEl.appendChild(tabContents);
        applyTenantLine(headerLeftEl, tenantLineText);
      } else {
        const existingLabel = headerTabsEl && headerTabsEl.querySelector(".sp-toolkit-header-tab-label");
        if (existingLabel) existingLabel.innerHTML = labelHtml;
        applyTenantLine(headerLeftEl, tenantLineText);
      }
    };
    listsLauncherSubsitesListener = (ev) => {
      if (!ev.data || ev.data.__spcsv !== true || ev.data.type !== "SPCSVSubsitesResult") return;
      window.removeEventListener("message", listsLauncherSubsitesListener);
      listsLauncherSubsitesListener = null;
      const subsites = Array.isArray(ev.data.subsites) ? ev.data.subsites : [];
      if (subsites.length === 0) return;
      const contentsWrap = panel.querySelector(".sp-toolkit-contents-wrap");
      const subsitesView = panel.querySelector(".sp-toolkit-subsites-view");
      const headerTabs = panel.querySelector(".sp-toolkit-header-tabs");
      if (!headerTabs || !subsitesView) return;
      if (!headerTabs.querySelector(".sp-toolkit-header-tab")) {
        const tabContents = document.createElement("button");
        tabContents.type = "button";
        tabContents.className = "sp-toolkit-header-tab sp-toolkit-header-tab-label active";
        tabContents.textContent = "Site Contents";
        tabContents.addEventListener("click", function () {
          headerTabs.querySelectorAll(".sp-toolkit-header-tab").forEach(function (t) { t.classList.remove("active"); });
          tabContents.classList.add("active");
          if (contentsWrap) contentsWrap.style.display = "block";
          subsitesView.classList.remove("visible");
        });
        headerTabs.appendChild(tabContents);
      }
      if (!headerTabs.querySelector(".sp-toolkit-header-tab-subsites")) {
        const tabSubsites = document.createElement("button");
        tabSubsites.type = "button";
        tabSubsites.className = "sp-toolkit-header-tab sp-toolkit-header-tab-subsites";
        tabSubsites.textContent = "Subsites";
        tabSubsites.addEventListener("click", function () {
          headerTabs.querySelectorAll(".sp-toolkit-header-tab").forEach(function (t) { t.classList.remove("active"); });
          tabSubsites.classList.add("active");
          if (contentsWrap) contentsWrap.style.display = "none";
          subsitesView.classList.add("visible");
        });
        headerTabs.appendChild(tabSubsites);
      }
      subsitesView.innerHTML = "";
      const list = document.createElement("div");
      list.className = "sp-toolkit-subsites-list";
      subsites.forEach(function (s) {
        const a = document.createElement("a");
        a.href = s.url || "#";
        a.target = "_blank";
        a.rel = "noopener";
        a.textContent = s.title || "Untitled";
        list.appendChild(a);
      });
      subsitesView.appendChild(list);
    };
    window.addEventListener("message", listsLauncherListsListener);
    window.addEventListener("message", listsLauncherSubsitesListener);
    const script = document.createElement("script");
    script.src = chrome.runtime.getURL("getSiteLists.js");
    script.onload = () => script.remove();
    script.onerror = () => {
      if (listsLauncherListsListener) {
        window.removeEventListener("message", listsLauncherListsListener);
        listsLauncherListsListener = null;
      }
      if (listsLauncherSubsitesListener) {
        window.removeEventListener("message", listsLauncherSubsitesListener);
        listsLauncherSubsitesListener = null;
      }
      const loadingEl = panel.querySelector(".sp-toolkit-lists-loading");
      const errorEl = panel.querySelector(".sp-toolkit-lists-error");
      if (loadingEl) loadingEl.style.display = "none";
      if (errorEl) {
        errorEl.textContent = "Failed to load lists.";
        errorEl.style.display = "block";
      }
    };
    (document.head || document.documentElement).appendChild(script);
  }
}

function closeListsLauncherPanel(e) {
  const container = document.getElementById(LISTS_LAUNCHER_ID);
  const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
  if (!container || !panel) return;
  if (e && e.target && container.contains(e.target)) return;
  if (!listsLauncherExpanded) return;
  listsLauncherExpanded = false;
  panel.classList.remove("sp-toolkit-lists-panel-open");
  panel.classList.add("sp-toolkit-panel-rolling-up");
  setTimeout(function () {
    panel.style.display = "none";
    panel.classList.remove("sp-toolkit-panel-rolling-up");
  }, 100);
}

function destroyListsLauncher() {
  const el = document.getElementById(LISTS_LAUNCHER_ID);
  if (el && el.parentNode) el.parentNode.removeChild(el);
}

function initListsLauncher() {
  if (window !== window.top) return;
  destroyListsLauncher();
  chrome.storage.local.get("listsLauncherEnabled", (r) => {
    if (r.listsLauncherEnabled === false) return;
    if (document.getElementById(LISTS_LAUNCHER_ID)) return;
    initListsLauncherUI();
  });
}

function initListsLauncherUI() {
  if (window !== window.top) return;
  if (document.getElementById(LISTS_LAUNCHER_ID)) return;
  if (!document.getElementById(LISTS_LAUNCHER_STYLE_ID)) {
    const style = document.createElement("style");
    style.id = LISTS_LAUNCHER_STYLE_ID;
    style.textContent = LISTS_LAUNCHER_CSS;
    document.head.appendChild(style);
  }
  const container = document.createElement("div");
  container.id = LISTS_LAUNCHER_ID;
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "sp-toolkit-lists-btn";
  btn.setAttribute("aria-label", "Site Contents");
  const compassImg = document.createElement("img");
  compassImg.className = "sp-toolkit-launcher-compass";
  compassImg.src = chrome.runtime.getURL("icon.png");
  compassImg.alt = "";
  compassImg.width = 24;
  compassImg.height = 24;
  btn.appendChild(compassImg);
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    btn.classList.remove("sp-toolkit-launcher-click-spin");
    void btn.offsetWidth;
    btn.classList.add("sp-toolkit-launcher-click-spin");
    setTimeout(function () { btn.classList.remove("sp-toolkit-launcher-click-spin"); }, 600);
    toggleListsLauncher();
  });
  const panel = document.createElement("div");
  panel.id = LISTS_LAUNCHER_PANEL_ID;
  panel.className = "sp-toolkit-lists-panel";
  panel.style.display = "none";
  panel.innerHTML =
    "<div class=\"sp-toolkit-lists-panel-header\"><div class=\"sp-toolkit-header-left\"><div class=\"sp-toolkit-header-tabs\"></div></div><div class=\"sp-toolkit-lists-panel-icons\"></div></div>" +
    "<div class=\"sp-toolkit-lists-panel-body\">" +
    "<div class=\"sp-toolkit-contents-wrap\">" +
    "<div class=\"sp-toolkit-lists-loading\"><span class=\"sp-toolkit-lists-loading-dot\"></span><span class=\"sp-toolkit-lists-loading-dot\"></span><span class=\"sp-toolkit-lists-loading-dot\"></span><span>Loading…</span></div>" +
    "<div class=\"sp-toolkit-lists-error\" style=\"display:none\"></div>" +
    "<table class=\"sp-toolkit-contents-table\"><thead></thead><tbody></tbody></table>" +
    "</div>" +
    "<div class=\"sp-toolkit-subsites-view\"></div>" +
    "</div>";
  panel.addEventListener("click", (e) => e.stopPropagation());
  container.appendChild(panel);
  container.appendChild(btn);
  document.body.appendChild(container);
  document.addEventListener("click", closeListsLauncherPanel);
  applyListsPanelTheme(panel);
}

function applyListsPanelTheme(panel) {
  if (!panel) return;
  const container = panel.parentNode && panel.parentNode.id === LISTS_LAUNCHER_ID ? panel.parentNode : document.getElementById(LISTS_LAUNCHER_ID);
  chrome.storage.local.get("darkMode", (r) => {
    panel.classList.remove("sp-toolkit-lists-panel-dark", "sp-toolkit-lists-panel-light");
    panel.classList.add(r.darkMode ? "sp-toolkit-lists-panel-dark" : "sp-toolkit-lists-panel-light");
    if (container) {
      container.classList.remove("sp-toolkit-launcher-dark", "sp-toolkit-launcher-light");
      container.classList.add(r.darkMode ? "sp-toolkit-launcher-dark" : "sp-toolkit-launcher-light");
    }
  });
}

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (window !== window.top || areaName !== "local") return;
  if (changes.listsLauncherEnabled) {
    const enabled = changes.listsLauncherEnabled.newValue !== false;
    if (enabled) requestAnimationFrame(function () { requestAnimationFrame(function () { initListsLauncherUI(); }); });
    else destroyListsLauncher();
    return;
  }
  if (changes.darkMode) {
    const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
    if (panel) applyListsPanelTheme(panel);
  }
});

function setupListeners() {
  window.removeEventListener("message", onPageMessage);
  window.addEventListener("message", onPageMessage);
}

function onPageMessage(e) {
  if (!e.data || e.data.__spcsv !== true) return;
  const d = e.data.detail || {};
  if (e.data.type === "SPCSVExportProgress") {
    let msg = d.message || "Export in progress…";
    if (d.totalCount != null && d.totalCount > 0 && d.currentCount != null) {
      msg = Math.min(100, Math.round((d.currentCount / d.totalCount) * 100)) + "% complete";
    }
    showProgress(msg);
  } else if (e.data.type === "SPCSVExportDone") {
    const success = !!(d && d.success);
    const message = (d && d.message) || (success ? "Done!" : "Export failed.");
    finishProgress(success, message, d && d.stopReason || "");
    try {
      chrome.runtime.sendMessage({ type: "SPCSVExportDone", success, message });
    } catch (err) {}
  }
}

// Inject a script and wait for one postMessage response. Handles timeout and script load error.
function injectAndWait(scriptName, messageType, parseData, sendResponse, options = {}) {
  const timeoutMs = options.timeoutMs ?? SCRIPT_TIMEOUT;
  const errorPayload = options.errorPayload ?? { ok: false, error: "Timeout or load failed" };
  let done = false;
  function finish(payload) {
    if (done) return;
    done = true;
    clearTimeout(tid);
    window.removeEventListener("message", listener);
    try {
      if (options.afterFinish) options.afterFinish();
    } catch (_) {}
    sendResponse(payload);
  }
  const listener = (ev) => {
    if (!ev.data || ev.data.__spcsv !== true || ev.data.type !== messageType) return;
    if (options.matchesResponse && !options.matchesResponse(ev.data)) return;
    finish(parseData(ev.data));
  };
  if (options.beforeInject) options.beforeInject();
  const script = document.createElement("script");
  script.src = chrome.runtime.getURL(scriptName) + (options.urlSuffix || "");
  script.onload = () => script.remove();
  script.onerror = () => finish(errorPayload);
  window.addEventListener("message", listener);
  const tid = setTimeout(() => finish(errorPayload), timeoutMs);
  (document.head || document.documentElement).appendChild(script);
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message.type === "SPOToolkitPopupOpened") {
    const launcher = document.getElementById(LISTS_LAUNCHER_ID);
    const panel = document.getElementById(LISTS_LAUNCHER_PANEL_ID);
    if (launcher) launcher.classList.remove("sp-toolkit-launcher-rolled-up");
    if (panel && listsLauncherExpanded) {
      listsLauncherExpanded = false;
      panel.classList.remove("sp-toolkit-lists-panel-open");
      panel.classList.add("sp-toolkit-panel-rolling-up");
      setTimeout(function () {
        panel.style.display = "none";
        panel.classList.remove("sp-toolkit-panel-rolling-up");
      }, 130);
    }
    sendResponse({});
    return true;
  }
  if (message.type === "SPOToolkitPopupClosed") {
    const launcher = document.getElementById(LISTS_LAUNCHER_ID);
    if (launcher) launcher.classList.add("sp-toolkit-launcher-rolled-up");
    sendResponse({});
    return true;
  }
  if (message.action === "showFact") {
    showFactToast(message.fact || "", message.sourceUrl || "");
    sendResponse({});
    return true;
  }
  if (message.action === "getPageContext") {
    injectAndWait(
      "getPageContext.js",
      "SPCSVPageContext",
      (data) => {
        const web = (data.webAbsoluteUrl || "").replace(/\/$/, "");
        const site = (data.siteAbsoluteUrl || "").replace(/\/$/, "");
        const primaryWeb = web || site;
        if (!primaryWeb) return { ok: false, error: "No context from page" };
        return {
          ok: true,
          webAbsoluteUrl: primaryWeb,
          siteAbsoluteUrl: site,
          pageListId: data.pageListId || "",
          listUrl: data.listUrl || "",
          siteId: data.siteId || "",
          webId: data.webId || ""
        };
      },
      sendResponse,
      { timeoutMs: 5000, errorPayload: { ok: false, error: "No context from page" } }
    );
    return true;
  }

  if (message.action === "getViewFormatContext") {
    injectAndWait(
      "getViewFormatContext.js",
      "SPCSVViewFormatContext",
      (data) => {
        const normGuid = (s) =>
          String(s || "")
            .replace(/[{}]/g, "")
            .replace(/%7B|%7D/gi, "")
            .trim();
        const listId = normGuid(data.listId);
        const viewId = normGuid(data.viewId);
        const webAbsoluteUrl = String(data.webAbsoluteUrl || "").replace(/\/$/, "");
        if (!listId || !webAbsoluteUrl) {
          return {
            ok: false,
            error: "Could not read list/site. Open a list or library view on SharePoint.",
          };
        }
        let sitePath = "/";
        try {
          sitePath = new URL(webAbsoluteUrl).pathname.replace(/\/$/, "") || "/";
        } catch (_) {}
        return { ok: true, listId, viewId, sitePath };
      },
      sendResponse,
      { timeoutMs: 6000, errorPayload: { ok: false, error: "Timeout reading page context." } }
    );
    return true;
  }

  if (message.action === "getPageContextJson") {
    injectAndWait(
      "getPageContextJson.js",
      "SPCSVPageContextJson",
      (data) => (data.error ? { ok: false, error: data.error } : { ok: true, data: data.data || {} }),
      sendResponse,
      { timeoutMs: 12000, errorPayload: { ok: false, error: "Timeout or page did not return context" } }
    );
    return true;
  }

  if (message.action === "rest") {
    const method = (message.method || "GET").toUpperCase();
    const path = message.path || "";
    const body = message.body;
    (async () => {
      try {
        let digest = null;
        if (method === "POST" || method === "PATCH" || method === "DELETE") {
          const apiIdx = path.indexOf("/_api/");
          const sitePath = apiIdx >= 0 ? path.substring(0, apiIdx) || "/" : "/";
          const contextUrl = location.origin + sitePath + "/_api/contextinfo";
          const cr = await fetch(contextUrl, {
            method: "POST",
            credentials: "include",
            headers: { Accept: "application/json;odata=nometadata" }
          });
          if (cr.ok) {
            const cj = await cr.json();
            digest = cj.FormDigestValue || null;
          }
        }
        const opts = {
          method,
          credentials: "include",
          headers: { Accept: "application/json;odata=nometadata" }
        };
        if (digest) opts.headers["X-RequestDigest"] = digest;
        if (body != null && method !== "GET") {
          opts.headers["Content-Type"] = "application/json;odata=nometadata";
          opts.body = typeof body === "string" ? body : JSON.stringify(body);
        }
        const r = await fetch(location.origin + path, opts);
        const text = await r.text();
        let data = text;
        try { data = JSON.parse(text); } catch (_) {}
        sendResponse(r.ok ? { ok: true, data } : { ok: false, status: r.status, error: data });
      } catch (e) {
        sendResponse({ ok: false, error: e && e.message ? e.message : String(e) });
      }
    })();
    return true;
  }

  if (message.action === "checkListPage") {
    injectAndWait(
      "checkListPage.js",
      "SPCSVListPageCheck",
      (data) => ({ isListPage: !!data.isListPage }),
      sendResponse,
      { timeoutMs: 3000, errorPayload: { isListPage: false } }
    );
    return true;
  }

  if (message.action === "getListType") {
    injectAndWait(
      "getListType.js",
      "SPCSVListTypeResult",
      (data) => ({
        isListPage: !!data.isListPage,
        isLibrary: !!data.isLibrary
      }),
      sendResponse,
      { timeoutMs: 6000, errorPayload: { isListPage: false, isLibrary: false } }
    );
    return true;
  }

  if (message.action === "getRefinableMappings") {
    injectAndWait(
      "getRefinableMappings.js",
      "SPCSVRefinableMappingsResult",
      (data) => ({
        ok: !!data.ok,
        error: data.error || null,
        mappings: Array.isArray(data.mappings) ? data.mappings : [],
        alias: data.alias != null ? data.alias : null
      }),
      sendResponse,
      {
        beforeInject() {
          let el = document.getElementById("sp-refinable-params");
          if (el) el.remove();
          el = document.createElement("script");
          el.id = "sp-refinable-params";
          el.type = "application/json";
          el.textContent = JSON.stringify({
            siteUrl: message.siteUrl || "",
            propertyName: message.propertyName || ""
          });
          (document.head || document.documentElement).appendChild(el);
        },
        timeoutMs: 15000,
        errorPayload: { ok: false, error: "Timeout or load failed", mappings: [], alias: null }
      }
    );
    return true;
  }

  if (message.action === "getSearchSchema") {
    injectAndWait(
      "getSearchSchema.js",
      "SPCSVSearchSchemaResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, columns: data.columns || [], siteUrl: data.siteUrl || "", listId: data.listId || "" },
      sendResponse,
      { errorPayload: { ok: false, error: "Timeout loading columns" } }
    );
    return true;
  }

  if (message.action === "getColumns") {
    injectAndWait(
      "getFields.js",
      "SPCSVFieldsResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, ...data },
      sendResponse,
      { errorPayload: { ok: false, error: "Timeout loading columns" } }
    );
    return true;
  }

  if (message.action === "getViewsData") {
    const requestId = createSpViewsRequestId();
    injectAndWait(
      "getViewsData.js",
      "SPCSVViewsDataResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, views: data.views || [], fields: data.fields || [], listId: data.listId, listTitle: data.listTitle, viewDetails: data.viewDetails },
      sendResponse,
      {
        beforeInject() {
          attachSpViewsParamsScript(message, requestId);
        },
        urlSuffix: "?spcsvRequestId=" + encodeURIComponent(requestId),
        matchesResponse(data) {
          return data.requestId === requestId;
        },
        afterFinish() {
          const el = document.getElementById(getSpViewsParamsScriptId(requestId));
          if (el) el.remove();
        },
        errorPayload: { ok: false, error: "Timeout loading views data" }
      }
    );
    return true;
  }

  if (message.action !== "runExportCSV") return;

  (async () => {
    showProgress("Starting export");
    setupListeners();
    const { siteUrl, listId, viewId, exportFilename, pageLimit, includeVersions, selectedColumns, report, format, permissionsMatrixWholeSite } = message;
    const params = {
      u: (siteUrl || "").replace(/\/$/, ""),
      lid: listId || "",
      vid: viewId || "",
      f: exportFilename || "",
      pl: pageLimit ?? 1000,
      incVer: includeVersions !== false
    };
    if (report) params.report = report;
    if (format) params.format = format;
    if (selectedColumns && Array.isArray(selectedColumns) && selectedColumns.length > 0) params.cols = selectedColumns;
    if (permissionsMatrixWholeSite === true) params.matrixWholeSite = true;
    let el = document.getElementById("spcsv-params-json");
    if (el) el.remove();
    el = document.createElement("script");
    el.id = "spcsv-params-json";
    el.type = "application/json";
    el.textContent = JSON.stringify(params);
    (document.head || document.documentElement).appendChild(el);

    function injectExportScript() {
      const script = document.createElement("script");
      script.src = chrome.runtime.getURL("exportCSV.js");
      script.onload = () => { script.remove(); sendResponse({ ok: true, message: "Export started. Watch the progress bar for status." }); };
      script.onerror = () => {
        finishProgress(false, "Failed to load export script.");
        sendResponse({ ok: false, error: "Failed to load exportCSV.js" });
      };
      (document.head || document.documentElement).appendChild(script);
    }

    if (format === "xlsx") {
      const exportUrl = chrome.runtime.getURL("exportCSV.js");
      const preloadScript = document.createElement("script");
      preloadScript.src = chrome.runtime.getURL("jszip-preload.js");
      preloadScript.onload = () => {
        const jszipScript = document.createElement("script");
        jszipScript.src = chrome.runtime.getURL("jszip.min.js");
        jszipScript.onload = () => {
          const restoreScript = document.createElement("script");
          restoreScript.src = chrome.runtime.getURL("jszip-restore.js") + "?url=" + encodeURIComponent(exportUrl);
          restoreScript.onload = () => { restoreScript.remove(); sendResponse({ ok: true, message: "Export started. Watch the progress bar for status." }); };
          restoreScript.onerror = () => {
            finishProgress(false, "Failed to load export script.");
            sendResponse({ ok: false, error: "Failed to load jszip-restore.js" });
          };
          (document.head || document.documentElement).appendChild(restoreScript);
        };
        jszipScript.onerror = () => {
          finishProgress(false, "Failed to load JSZip for XLSX.");
          sendResponse({ ok: false, error: "Failed to load jszip.min.js" });
        };
        (document.head || document.documentElement).appendChild(jszipScript);
      };
      preloadScript.onerror = () => {
        finishProgress(false, "Failed to load XLSX preload.");
        sendResponse({ ok: false, error: "Failed to load jszip-preload.js" });
      };
      (document.head || document.documentElement).appendChild(preloadScript);
    } else {
      injectExportScript();
    }
  })();
  return true;
});

if (document.body) initListsLauncher();
else document.addEventListener("DOMContentLoaded", initListsLauncher);
