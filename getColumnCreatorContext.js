// Injected: site + list field catalogs for column creator internal-name prediction.
(function () {
  var spi = window._spPageContextInfo || {};
  var siteUrl = (spi.webAbsoluteUrl || spi.siteAbsoluteUrl || "").replace(/\/$/, "");
  var listUrl = spi.listUrl || spi.listServerRelativeUrl;
  var accept = "application/json;odata=nometadata";

  function post(payload) {
    window.postMessage(Object.assign({ __spcsv: true, type: "SPCSVColumnCreatorContextResult" }, payload), "*");
  }

  if (!siteUrl) {
    post({ ok: false, error: "Open a SharePoint site.", siteFields: [], listFields: [], siteUrl: "", listId: "" });
    return;
  }

  function mapField(f) {
    return {
      internalName: f.InternalName || f.Title || "",
      title: f.Title || f.InternalName || "",
      type: f.TypeAsString || "",
      group: f.Group || "",
    };
  }

  function isHiddenSystemField(internalName) {
    if (!internalName) return true;
    if (/^_|^vti_|^ows_|^tp_/.test(internalName)) return true;
    return false;
  }

  function fetchJson(url) {
    return fetch(url, { credentials: "include", headers: { Accept: accept } }).then(function (r) {
      if (!r.ok) throw new Error(r.status + " " + r.statusText);
      return r.json();
    });
  }

  var siteFieldsPromise = fetchJson(
    siteUrl + "/_api/web/fields?$select=InternalName,Title,TypeAsString,Group&$orderby=Title"
  ).then(function (data) {
    var raw = data.value || data.d?.results || [];
    var fields = [];
    for (var i = 0; i < raw.length; i++) {
      var f = raw[i];
      var internalName = f.InternalName || f.Title || "";
      if (isHiddenSystemField(internalName)) continue;
      fields.push(mapField(f));
    }
    return fields;
  });

  if (!listUrl) {
    siteFieldsPromise
      .then(function (siteFields) {
        post({ ok: true, siteUrl: siteUrl, listId: "", siteFields: siteFields, listFields: [] });
      })
      .catch(function (err) {
        post({ ok: false, error: (err && err.message) || String(err), siteUrl: siteUrl, listId: "", siteFields: [], listFields: [] });
      });
    return;
  }

  var enc = encodeURIComponent("'" + listUrl.replace(/'/g, "''") + "'");
  fetchJson(siteUrl + "/_api/web/GetList(@u)/Id?@u=" + enc)
    .then(function (j) {
      var listId = (j.value || "").replace(/[{}]/g, "");
      if (!listId) throw new Error("Could not get list ID");
      var lb = siteUrl + "/_api/web/lists(guid'" + listId + "')";
      return Promise.all([
        siteFieldsPromise,
        fetchJson(lb + "/fields?$select=InternalName,Title,TypeAsString,Group&$orderby=Title"),
      ]).then(function (results) {
        var listRaw = results[1].value || results[1].d?.results || [];
        var listFields = [];
        for (var i = 0; i < listRaw.length; i++) {
          var lf = listRaw[i];
          var iname = lf.InternalName || lf.Title || "";
          if (isHiddenSystemField(iname)) continue;
          listFields.push(mapField(lf));
        }
        post({
          ok: true,
          siteUrl: siteUrl,
          listId: listId,
          siteFields: results[0],
          listFields: listFields,
        });
      });
    })
    .catch(function (err) {
      post({
        ok: false,
        error: (err && err.message) || String(err),
        siteUrl: siteUrl,
        listId: "",
        siteFields: [],
        listFields: [],
      });
    });
})();
