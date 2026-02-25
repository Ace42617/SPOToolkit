// Injected: create or update a list view (fields + ViewQuery). Reads payload from script tag sp-save-view-payload.
// normKey/canonicalName map SharePoint display aliases to canonical API names (same for all lists).
(function () {
  function send(success, message, detail) {
    window.postMessage({ __spcsv: true, type: "SPCSVSaveViewResult", success: !!success, message: message || "", detail: detail || {} }, "*");
  }
  var payload = null;
  try {
    var el = document.getElementById("sp-save-view-payload");
    if (el && el.textContent) payload = JSON.parse(el.textContent);
  } catch (_) {}
  if (!payload || !payload.listId) {
    send(false, "Missing payload or listId.");
    return;
  }
  var spi = window._spPageContextInfo || {};
  var siteUrl = (spi.webAbsoluteUrl || spi.siteAbsoluteUrl || "").replace(/\/$/, "");
  var listId = String(payload.listId).replace(/[{}]/g, "").trim();
  var viewId = payload.viewId ? String(payload.viewId).replace(/[{}]/g, "").trim() : "";
  var viewTitle = (payload.viewTitle || "").trim() || "New View";
  var viewFields = Array.isArray(payload.viewFields) ? payload.viewFields : [];
  var orderBy = payload.orderBy && payload.orderBy.field ? payload.orderBy : null;
  var filters = Array.isArray(payload.filters) ? payload.filters : (payload.filter && payload.filter.field ? [payload.filter] : []);
  var groupBy = payload.groupBy && String(payload.groupBy).trim() ? String(payload.groupBy).trim() : null;

  function normalizeGuid(g) {
    if (g == null) return "";
    var s = String(g).trim();
    s = s.replace(/[{}]/g, "");
    return s;
  }

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function getDigestFromDomOrContext() {
    var el = document.getElementById("__REQUESTDIGEST");
    if (el && el.value) return el.value;
    if (window._spPageContextInfo && window._spPageContextInfo.formDigestValue) return window._spPageContextInfo.formDigestValue;
    return null;
  }

  function getDigest() {
    var d = getDigestFromDomOrContext();
    if (d) return Promise.resolve(d);
    return fetch(siteUrl + "/_api/contextinfo", {
      method: "POST",
      credentials: "include",
      headers: { "Accept": "application/json;odata=nometadata" }
    }).then(function (r) { return r.ok ? r.json() : Promise.reject(new Error("contextinfo failed")); })
      .then(function (j) { return j.FormDigestValue || null; });
  }

  function listBaseUrl() { return siteUrl + "/_api/web/lists(guid'" + listId + "')"; }

  function escapeCamlValue(val) {
    var s = String(val == null ? "" : val);
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function buildViewQuery() {
    var parts = [];
    if (orderBy && orderBy.field) {
      var asc = orderBy.ascending !== false ? "True" : "False";
      parts.push("<OrderBy><FieldRef Name=\"" + orderBy.field.replace(/"/g, "&quot;") + "\" Ascending=\"" + asc + "\"/></OrderBy>");
    }
    if (filters.length > 0) {
      var conds = [];
      for (var fi = 0; fi < filters.length; fi++) {
        var f = filters[fi];
        if (!f || !f.field) continue;
        var op = (f.op || "Eq").replace(/[^a-zA-Z]/g, "");
        if (!op) op = "Eq";
        var vType = (f.valueType || "Text").replace(/[^a-zA-Z#]/g, "") || "Text";
        var v = escapeCamlValue(f.value || "");
        conds.push("<" + op + "><FieldRef Name=\"" + f.field.replace(/"/g, "&quot;") + "\"/><Value Type=\"" + vType + "\">" + v + "</Value></" + op + ">");
      }
      if (conds.length === 1) parts.push("<Where>" + conds[0] + "</Where>");
      else if (conds.length > 1) parts.push("<Where><And>" + conds.join("") + "</And></Where>");
    }
    if (groupBy) {
      parts.push("<GroupBy><FieldRef Name=\"" + groupBy.replace(/"/g, "&quot;") + "\"/></GroupBy>");
    }
    return parts.length ? parts.join("") : "";
  }

  function removeAllViewFields(viewId) {
    return getDigest().then(function (digest) {
      if (!digest) return Promise.reject(new Error("Could not get request digest"));
      var vfBase = listBaseUrl() + "/views(guid'" + normalizeGuid(viewId) + "')/ViewFields";
      return fetch(vfBase + "/RemoveAllViewFields", {
        method: "POST",
        credentials: "include",
        headers: { "Accept": "application/json;odata=nometadata", "X-RequestDigest": digest }
      }).then(function (ra) {
        if (ra.ok) return { ok: true };
        return fetch(vfBase, { credentials: "include", headers: { "Accept": "application/json;odata=nometadata" } })
          .then(function (r) { return r.json(); })
          .then(function (j) {
            var items = j.Items || j.value || [];
            return Promise.all(items.map(function (name) {
              var n = String(name).replace(/'/g, "''");
              return fetch(vfBase + "/removeviewfield('" + n + "')", {
                method: "POST",
                credentials: "include",
                headers: { "Accept": "application/json;odata=nometadata", "X-RequestDigest": digest }
              });
            })).then(function () { return { ok: true }; });
          });
      });
    });
  }

  function addViewField(viewId, internalName) {
    return getDigest().then(function (digest) {
      if (!digest) return Promise.reject(new Error("Could not get request digest"));
      var vfBase = listBaseUrl() + "/views(guid'" + normalizeGuid(viewId) + "')/ViewFields";
      var name = String(internalName).replace(/'/g, "''");
      return fetch(vfBase + "/addviewfield('" + name + "')", {
        method: "POST",
        credentials: "include",
        headers: { "Accept": "application/json;odata=nometadata", "X-RequestDigest": digest }
      }).then(function (r) { return { ok: r.ok, error: r.ok ? null : (r.status + "") }; });
    });
  }

  function setViewFieldsBatch(viewId, fieldNames) {
    function normKey(n) {
      if (n === "Id") return "ID";
      if (n === "LinkFilename" || n === "LinkFilenameNoMenu" || n === "FileLeafRef" || n === "Name") return "FileLeafRef";
      if (n === "LinkTitle" || n === "LinkTitleNoMenu" || n === "Title") return "Title";
      return n;
    }
    function canonicalName(n) {
      if (n === "Id") return "ID";
      if (n === "LinkFilename" || n === "LinkFilenameNoMenu" || n === "Name") return "FileLeafRef";
      if (n === "LinkTitle" || n === "LinkTitleNoMenu") return "Title";
      return n;
    }
    var seen = {};
    var names = [];
    for (var i = 0; i < fieldNames.length; i++) {
      var n = String(fieldNames[i]).trim();
      if (!n) continue;
      var k = normKey(n);
      if (seen[k]) continue;
      seen[k] = true;
      names.push(canonicalName(n));
    }
    if (!seen["ID"]) { seen["ID"] = true; names.unshift("ID"); }
    return removeAllViewFields(viewId).then(function () { return sleep(300); }).then(function () {
      return names.reduce(function (p, name) {
        return p.then(function () { return addViewField(viewId, name); });
      }, Promise.resolve());
    }).then(function () { return { ok: true }; });
  }

  function setViewQueryById(viewId, viewQuery) {
    return getDigest().then(function (digest) {
      if (!digest) return Promise.reject(new Error("Could not get request digest"));
      var apiUrl = listBaseUrl() + "/views(guid'" + normalizeGuid(viewId) + "')";
      return fetch(apiUrl, {
        method: "POST",
        credentials: "include",
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": digest,
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE"
        },
        body: JSON.stringify({ ViewQuery: viewQuery })
      }).then(function (r) { return { ok: r.ok, error: r.ok ? null : (r.status + "") }; });
    });
  }

  var accept = { "Accept": "application/json;odata=nometadata" };
  var lb = listBaseUrl();

  (function run() {
    var isNew = !viewId;
    var p = isNew
      ? getDigest().then(function (digest) {
        if (!digest) return Promise.reject(new Error("Could not get request digest"));
        return fetch(lb + "/views", {
          method: "POST",
          credentials: "include",
          headers: {
            "Accept": "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": digest
          },
          body: JSON.stringify({
            Title: viewTitle,
            PersonalView: false,
            RowLimit: 5000,
            Scope: 2,
            ViewQuery: buildViewQuery()
          })
        }).then(function (r) {
          if (!r.ok) return r.text().then(function (t) { return Promise.reject(new Error("Create view failed: " + r.status + " " + t)); });
          return r.json();
        }).then(function (created) {
          viewId = normalizeGuid(created.Id);
          return viewId;
        });
      })
      : Promise.resolve(viewId);

    p.then(function () {
      if (viewFields.length > 0) return setViewFieldsBatch(viewId, viewFields);
      return { ok: true };
    }).then(function (sf) {
      if (!sf.ok) return Promise.reject(new Error("Set view fields failed"));
      var vq = buildViewQuery();
      if (vq) return setViewQueryById(viewId, vq);
      return { ok: true };
    }).then(function (sq) {
      if (!sq.ok) return Promise.reject(new Error("Set view query failed"));
      send(true, "Save complete.", { viewId: viewId, viewTitle: viewTitle });
    }).catch(function (err) {
      send(false, (err && err.message) || String(err), {});
    });
  })();
})();
