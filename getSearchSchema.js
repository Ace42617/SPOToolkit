// Injected: fetches list columns (InternalName, Title, Type, Group) for current list via REST. Uses list GUID like getFields/getViewsData.
(function () {
  var spi = window._spPageContextInfo || {};
  var siteUrl = (spi.webAbsoluteUrl || spi.siteAbsoluteUrl || "").replace(/\/$/, "");
  var listUrl = spi.listUrl || spi.listServerRelativeUrl;
  if (!siteUrl || !listUrl) {
    window.postMessage({ __spcsv: true, type: "SPCSVSearchSchemaResult", error: "Open a list or library page.", columns: [] }, "*");
    return;
  }
  var accept = "application/json;odata=nometadata";
  var enc = encodeURIComponent("'" + listUrl.replace(/'/g, "''") + "'");
  fetch(siteUrl + "/_api/web/GetList(@u)/Id?@u=" + enc, { credentials: "include", headers: { Accept: accept } })
    .then(function (r) { return r.json(); })
    .then(function (j) {
      var listId = (j.value || "").replace(/[{}]/g, "");
      if (!listId) throw new Error("Could not get list ID");
      var lb = siteUrl + "/_api/web/lists(guid'" + listId + "')";
      return fetch(lb + "/fields?$select=InternalName,Title,TypeAsString,Group,EntityPropertyName&$orderby=Title", { credentials: "include", headers: { Accept: accept } });
    })
    .then(function (r) {
      if (!r.ok) throw new Error(r.status + " " + r.statusText);
      return r.json();
    })
    .then(function (data) {
      var raw = data.value || data.d?.results || [];
      var columns = [];
      for (var i = 0; i < raw.length; i++) {
        var f = raw[i];
        var internalName = f.InternalName || f.Title || "";
        if (!internalName || /^_|^vti_|^ows_|^tp_/.test(internalName)) continue;
        columns.push({
          internalName: internalName,
          title: f.Title || f.InternalName || "",
          type: f.TypeAsString || "",
          group: f.Group || "",
          alias: f.EntityPropertyName || null
        });
      }
      window.postMessage({ __spcsv: true, type: "SPCSVSearchSchemaResult", columns: columns, error: null }, "*");
    })
    .catch(function (err) {
      window.postMessage({ __spcsv: true, type: "SPCSVSearchSchemaResult", error: (err && err.message) || String(err), columns: [] }, "*");
    });
})();
