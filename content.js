// Runs on SharePoint pages. Injects scripts into page context so fetch uses session.

const PROGRESS_BOX_ID = "sp-csv-export-progress";
const SCRIPT_TIMEOUT = 15000;

function showProgress(message) {
  let el = document.getElementById(PROGRESS_BOX_ID);
  if (!el) {
    el = document.createElement("div");
    el.id = PROGRESS_BOX_ID;
    el.style.cssText =
      "position:fixed;top:20px;left:50%;transform:translateX(-50%);z-index:2147483647;font-family:'Segoe UI',sans-serif;min-width:400px;max-width:90vw;";
    el.innerHTML =
      '<div style="display:flex;align-items:center;gap:12px;padding:16px 24px;background:#323130;color:#fff;border-radius:8px;box-shadow:0 4px 24px rgba(0,0,0,0.5);font-size:14px;border:3px solid #0078d4;">' +
      '<div style="width:20px;height:20px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;border-radius:50%;animation:spcsvspin 0.8s linear infinite;flex-shrink:0;"></div>' +
      '<span class="sp-csv-msg" style="flex:1;">' + (message || "Starting…") + "</span></div>";
    const style = document.createElement("style");
    style.textContent = "@keyframes spcsvspin { to { transform: rotate(360deg); } }";
    document.head.appendChild(style);
    document.body.appendChild(el);
  }
  const msgEl = el.querySelector(".sp-csv-msg");
  if (msgEl) msgEl.textContent = message || "…";
}

function finishProgress(success, message, stopReason) {
  const el = document.getElementById(PROGRESS_BOX_ID);
  if (!el) return;
  const box = el.querySelector("div");
  const msgEl = el.querySelector(".sp-csv-msg");
  const spinner = box && box.querySelector("div");
  if (box) box.style.background = success ? "#107c10" : "#a4262c";
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
    finishProgress(!!(d && d.success), (d && d.message) || (d && d.success ? "Done!" : "Export failed."), d && d.stopReason || "");
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
    sendResponse(payload);
  }
  const listener = (ev) => {
    if (!ev.data || ev.data.__spcsv !== true || ev.data.type !== messageType) return;
    finish(parseData(ev.data));
  };
  if (options.beforeInject) options.beforeInject();
  const script = document.createElement("script");
  script.src = chrome.runtime.getURL(scriptName);
  script.onload = () => script.remove();
  script.onerror = () => finish(errorPayload);
  window.addEventListener("message", listener);
  const tid = setTimeout(() => finish(errorPayload), timeoutMs);
  (document.head || document.documentElement).appendChild(script);
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message.action === "getPageContext") {
    injectAndWait(
      "getPageContext.js",
      "SPCSVPageContext",
      (data) => {
        const url = data.webAbsoluteUrl || data.siteAbsoluteUrl || "";
        return url ? { ok: true, webAbsoluteUrl: url, pageListId: data.pageListId || "", listUrl: data.listUrl || "" } : { ok: false, error: "No context from page" };
      },
      sendResponse,
      { timeoutMs: 5000, errorPayload: { ok: false, error: "No context from page" } }
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

  if (message.action === "getSearchSchema") {
    injectAndWait(
      "getSearchSchema.js",
      "SPCSVSearchSchemaResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, columns: data.columns || [], siteUrl: data.siteUrl },
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
    injectAndWait(
      "getViewsData.js",
      "SPCSVViewsDataResult",
      (data) => data.error ? { ok: false, error: data.error } : { ok: true, views: data.views || [], fields: data.fields || [], listId: data.listId, listTitle: data.listTitle, viewDetails: data.viewDetails },
      sendResponse,
      {
        beforeInject() {
          let el = document.getElementById("sp-views-params");
          if (el) el.remove();
          el = document.createElement("script");
          el.id = "sp-views-params";
          el.type = "application/json";
          el.textContent = JSON.stringify({ viewId: message.viewId || null });
          (document.head || document.documentElement).appendChild(el);
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
