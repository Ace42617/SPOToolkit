// Injected: fetch list views, list fields, and optionally one view's ViewFields + ViewQuery.
(function () {
  var spi = window._spPageContextInfo || {};
  var siteUrl = (spi.webAbsoluteUrl || spi.siteAbsoluteUrl || "").replace(/\/$/, "");
  var listUrl = spi.listUrl || spi.listServerRelativeUrl;
  var params = {};
  try {
    var el = document.getElementById("sp-views-params");
    if (el && el.textContent) params = JSON.parse(el.textContent);
  } catch (_) {}
  function send(data) {
    window.postMessage({ __spcsv: true, type: "SPCSVViewsDataResult", ...data }, "*");
  }
  if (!siteUrl || !listUrl) {
    send({ error: "Open a list or library page.", views: [], fields: [] });
    return;
  }
  var accept = "application/json;odata=nometadata";
  var enc = encodeURIComponent("'" + listUrl.replace(/'/g, "''") + "'");
  fetch(siteUrl + "/_api/web/GetList(@u)/Id?@u=" + enc, { credentials: "include", headers: { Accept: accept } })
    .then(function (r) { return r.json(); })
    .then(function (j) {
      var listId = (j.value || "").replace(/[{}]/g, "");
      if (!listId) throw new Error("Could not get list ID");
      return fetch(siteUrl + "/_api/web/lists(guid'" + listId + "')?$select=Title", { credentials: "include", headers: { Accept: accept } })
        .then(function (r) { return r.json(); })
        .then(function (listData) {
          var listTitle = (listData.Title || listData.title || "").replace(/'/g, "''");
          return { listId: listId, listTitle: listTitle };
        });
    })
    .then(function (ctx) {
      var lb = siteUrl + "/_api/web/lists(guid'" + ctx.listId + "')";
      return fetch(lb + "/views?$select=Id,Title,DefaultView", { credentials: "include", headers: { Accept: accept } })
        .then(function (r) { return r.json(); })
        .then(function (j) {
          var views = (j.value || j.d?.results || []).map(function (v) {
            return { id: (v.Id || "").replace(/[{}]/g, ""), title: v.Title || "", defaultView: !!v.DefaultView };
          });
          return fetch(lb + "/fields?$select=InternalName,Title,Hidden,TypeAsString&$orderby=Title&$filter=Hidden eq false", { credentials: "include", headers: { Accept: accept } })
            .then(function (r) { return r.json(); })
            .then(function (j2) {
              var raw = j2.value || j2.d?.results || [];
              var fields = [];
              for (var i = 0; i < raw.length; i++) {
                var f = raw[i];
                var iname = f.InternalName || f.Title || "";
                if (iname && !/^_|^vti_|^ows_|^tp_/.test(iname)) fields.push({ internalName: iname, title: f.Title || iname, typeAsString: f.TypeAsString || "Text" });
              }
              return { ctx: ctx, views: views, fields: fields };
            });
        });
    })
    .then(function (data) {
      var viewId = params.viewId && String(params.viewId).replace(/[{}]/g, "");
      if (!viewId) {
        send({ views: data.views, fields: data.fields, listId: data.ctx.listId, listTitle: data.ctx.listTitle.replace(/''/g, "'") });
        return;
      }
      var lb = siteUrl + "/_api/web/lists(guid'" + data.ctx.listId + "')";
      var vBase = lb + "/views(guid'" + viewId + "')";
      Promise.all([
        fetch(vBase + "/ViewFields", { credentials: "include", headers: { Accept: accept } }).then(function (r) { return r.json(); }),
        fetch(vBase + "?$select=ViewQuery,Title", { credentials: "include", headers: { Accept: accept } }).then(function (r) { return r.json(); })
      ]).then(function (results) {
        var vfItems = results[0].Items || results[0].value || results[0].d?.results || [];
        var viewFields = [].concat(vfItems).filter(Boolean).map(function (x) { return String(typeof x === "object" ? (x.Name || x) : x); });
        var viewQuery = (results[1].ViewQuery || "").replace(/\s+/g, " ");
        var viewTitle = results[1].Title || "";
        var orderBy = null, filters = [], groupBy = null;
        var orderMatch = /<OrderBy>\s*<FieldRef\s+Name="([^"]+)"\s+Ascending="(True|False)"/i.exec(viewQuery);
        if (orderMatch) orderBy = { field: orderMatch[1], ascending: orderMatch[2].toLowerCase() === "true" };
        var groupMatch = /<GroupBy>\s*<FieldRef\s+Name="([^"]+)"/i.exec(viewQuery);
        if (groupMatch) groupBy = groupMatch[1];
        var condRegex = /<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)>\s*<FieldRef\s+Name="([^"]+)"\s*\/>\s*<Value\s+Type="([^"]*)">([^<]*)<\/Value>/gi;
        var m;
        while ((m = condRegex.exec(viewQuery)) !== null) {
          filters.push({ op: m[1], field: m[2], valueType: m[3] || "Text", value: (m[4] || "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">") });
        }
        send({
          views: data.views,
          fields: data.fields,
          listId: data.ctx.listId,
          listTitle: data.ctx.listTitle.replace(/''/g, "'"),
          viewDetails: { viewFields: viewFields, viewQuery: viewQuery, viewTitle: viewTitle, orderBy: orderBy, filters: filters, groupBy: groupBy }
        });
      }).catch(function (err) {
        send({ views: data.views, fields: data.fields, listId: data.ctx.listId, listTitle: data.ctx.listTitle.replace(/''/g, "'"), error: (err && err.message) || "Failed to load view details" });
      });
    })
    .catch(function (err) {
      send({ error: (err && err.message) || String(err), views: [], fields: [] });
    });
})();
