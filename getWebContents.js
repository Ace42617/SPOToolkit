// Injected page script: fetch lists + child webs for one SharePoint web URL.
(function () {
  var REQUEST_QUERY = "spcsvRequestId";
  var activeRequestId = getCurrentRequestId();

  function getCurrentRequestId() {
    try {
      var script = document.currentScript;
      var src = script && script.src ? String(script.src) : "";
      if (!src) return "";
      return new URL(src, window.location.href).searchParams.get(REQUEST_QUERY) || "";
    } catch (e) {
      return "";
    }
  }

  function findParamsElement(requestId) {
    var nodes = document.querySelectorAll('script[data-sp-site-contents-web-params="1"]');
    if (!nodes || !nodes.length) return null;
    if (requestId) {
      for (var i = 0; i < nodes.length; i++) {
        if (nodes[i].getAttribute("data-request-id") === requestId) return nodes[i];
      }
      return null;
    }
    return nodes[nodes.length - 1] || null;
  }

  function readParams() {
    try {
      var el = findParamsElement(activeRequestId);
      if (el && el.textContent) {
        var parsed = JSON.parse(el.textContent);
        if (!activeRequestId && parsed && parsed.requestId) {
          activeRequestId = String(parsed.requestId);
        }
        return parsed;
      }
    } catch (_) {}
    return {};
  }

  var params = readParams();
  var requestId = activeRequestId || params.requestId || "";
  var siteUrl = String(params.webUrl || "").replace(/\/$/, "");
  var accept = "application/json;odata=nometadata";
  var typeLabels = {
    100: "List",
    101: "Document library",
    106: "Events list",
    107: "Links list",
    109: "Picture library",
    110: "Survey",
    119: "Page library",
    171: "Issue tracking"
  };

  function post(payload) {
    window.postMessage(Object.assign({
      __spcsv: true,
      type: "SPCSVWebContentsResult",
      requestId: requestId,
      webUrl: siteUrl
    }, payload || {}), "*");
  }

  if (!siteUrl) {
    post({ ok: false, error: "No web URL", lists: [], subsites: [] });
    return;
  }

  function mapListsJson(j) {
    var results = (j.value != null ? j.value : (j.d && j.d.results) != null ? j.d.results : (j.d && j.d) != null ? j.d : []);
    if (!Array.isArray(results)) results = [];
    var base = location.origin;
    return results.map(function (item) {
      var title = item.Title || item.title || "";
      var viewUrl = item.DefaultViewUrl || item.defaultViewUrl || "";
      if (viewUrl && viewUrl.indexOf("http") !== 0) viewUrl = base + (viewUrl.indexOf("/") === 0 ? viewUrl : "/" + viewUrl);
      var baseTemplate = item.BaseTemplate != null ? item.BaseTemplate : item.baseTemplate;
      var isLibrary = (baseTemplate === 101 || baseTemplate === 109 || baseTemplate === 119);
      var typeLabel = typeLabels[baseTemplate] || "List";
      var itemCount = item.ItemCount != null ? item.ItemCount : (item.itemCount != null ? item.itemCount : 0);
      var modified = item.LastItemModifiedDate || item.lastItemModifiedDate || "";
      return {
        id: item.Id,
        title: title,
        viewUrl: viewUrl || "",
        isLibrary: isLibrary,
        typeLabel: typeLabel,
        itemCount: itemCount,
        modified: modified,
        webUrl: siteUrl
      };
    }).filter(function (item) { return item.viewUrl; });
  }

  function mapSubsitesJson(j) {
    var results = (j.value != null ? j.value : (j.d && j.d.results) != null ? j.d.results : (j.d && j.d) != null ? j.d : []);
    if (!Array.isArray(results)) results = [];
    var base = location.origin;
    return results.map(function (item) {
      var title = item.Title || item.title || "";
      var url = (item.Url || item.url || "").replace(/\/$/, "");
      if (url && url.indexOf("http") !== 0) url = base + (url.indexOf("/") === 0 ? url : "/" + url);
      return { title: title, url: url };
    }).filter(function (item) { return item.url; });
  }

  var listsUrl = siteUrl + "/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,DefaultViewUrl,BaseTemplate,ItemCount,LastItemModifiedDate&$orderby=Title asc&$top=500";
  var subsitesUrl = siteUrl + "/_api/web/webs?$select=Title,Url&$top=200";

  Promise.all([
    fetch(listsUrl, { credentials: "include", headers: { Accept: accept } }).then(function (r) {
      if (!r.ok) throw new Error("HTTP " + r.status + " loading lists");
      return r.json();
    }),
    fetch(subsitesUrl, { credentials: "include", headers: { Accept: accept } }).then(function (r) {
      if (!r.ok) return { value: [] };
      return r.json();
    }).catch(function () { return { value: [] }; })
  ])
    .then(function (pair) {
      post({
        ok: true,
        lists: mapListsJson(pair[0] || {}),
        subsites: mapSubsitesJson(pair[1] || {})
      });
    })
    .catch(function (err) {
      post({
        ok: false,
        error: (err && err.message) || "Failed to load web contents",
        lists: [],
        subsites: []
      });
    });
})();
