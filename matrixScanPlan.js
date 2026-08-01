// Injected page script: permissions matrix scan plan (site/subsite inventory).
(function () {
  function readParams() {
    try {
      var el = document.getElementById("sp-matrix-scan-params-json");
      if (el && el.textContent) return JSON.parse(el.textContent);
    } catch (_) {}
    return {};
  }

  var params = readParams();
  var includeSubsites = params.includeSubsites !== false;

  function post(type, payload) {
    window.postMessage(Object.assign({ __spcsv: true, type: type }, payload || {}), "*");
  }

  function normalizePath(urlOrPath) {
    if (!urlOrPath) return "";
    var s = String(urlOrPath).trim();
    if (s.charAt(0) === "/") return s.replace(/\/$/, "") || "/";
    try { return new URL(s).pathname.replace(/\/$/, "") || "/"; } catch (_) { return s.replace(/\/$/, ""); }
  }

  var apiOrigin = String(location.origin || "").replace(/\/$/, "");
  var spi = typeof window._spPageContextInfo === "object" ? window._spPageContextInfo : null;
  if (spi && spi.webAbsoluteUrl) {
    try { apiOrigin = new URL(spi.webAbsoluteUrl).origin.replace(/\/$/, ""); } catch (_) {}
  } else if (params.siteUrl) {
    try { apiOrigin = new URL(String(params.siteUrl)).origin.replace(/\/$/, ""); } catch (_) {}
  }

  function webUrlFromPath(path) {
    var p = normalizePath(path);
    if (!p || p === "/") return apiOrigin;
    return apiOrigin + p;
  }

  async function fetchJson(url, accept) {
    var r = await fetch(url, { credentials: "include", headers: { Accept: accept || "application/json;odata=nometadata" } });
    if (!r.ok) throw new Error("HTTP " + r.status + " for " + url);
    return r.json();
  }

  async function buildPlan() {
    var accept = "application/json;odata=nometadata";
    var plan = [];
    var seen = {};

    async function addWeb(path) {
      var p = normalizePath(path);
      if (!p) return;
      var key = p.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      var webUrl = webUrlFromPath(p);
      var j = await fetchJson(webUrl + "/_api/web?$select=Title,ServerRelativeUrl,Url,HasUniqueRoleAssignments", accept);
      var sr = normalizePath(j.ServerRelativeUrl || j.serverRelativeUrl || p);
      plan.push({
        title: (j.Title || j.title || sr || "Site").trim(),
        path: sr,
        url: (j.Url || j.url || webUrl).trim(),
        hasUnique: j.HasUniqueRoleAssignments === true,
        listCount: 0,
        itemCount: 0
      });
      if (!includeSubsites) return;
      var subs = await fetchJson(webUrl + "/_api/web/webs?$select=Title,ServerRelativeUrl,Url&$top=500", accept);
      var items = subs.value || subs.results || [];
      for (var i = 0; i < items.length; i++) {
        var subPath = items[i].ServerRelativeUrl || items[i].serverRelativeUrl;
        if (!subPath && (items[i].Url || items[i].url)) subPath = normalizePath(items[i].Url || items[i].url);
        if (subPath) await addWeb(subPath);
      }
    }

    var startPath = normalizePath(params.siteUrl || (spi && spi.webServerRelativeUrl) || "/");
    if (startPath && startPath !== "/" && startPath.indexOf("/") === 0) {
      await addWeb(startPath);
    } else {
      var root = await fetchJson(apiOrigin + "/_api/web?$select=ServerRelativeUrl", accept);
      await addWeb(root.ServerRelativeUrl || root.serverRelativeUrl || "/");
    }
    return plan;
  }

  buildPlan().then(function (plan) {
    post("SPMatrixScanPlanResult", { ok: true, plan: plan });
  }).catch(function (e) {
    post("SPMatrixScanPlanResult", { ok: false, error: (e && e.message) || String(e) });
  });
})();
