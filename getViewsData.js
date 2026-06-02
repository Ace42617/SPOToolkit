// Injected: View Manager REST loader. Behavior must match lib/viewsDataCore.mjs — run `npm test` after edits.
// Lite copy: SP-Developer-Toolkit-Lite/getViewsData.js
(function () {
  var ACCEPT_NOMETADATA = "application/json;odata=nometadata";
  var RESULT_TYPE = "SPCSVViewsDataResult";

  function normalizeWebUrl(url) {
    return String(url || "").replace(/\/$/, "");
  }

  function parseInjectParams(text) {
    try {
      if (text == null || String(text).trim() === "") return {};
      return JSON.parse(String(text));
    } catch (e) {
      return {};
    }
  }

  function resolveSiteUrl(paramWeb, pageContext) {
    var pc = pageContext || {};
    var p =
      paramWeb != null && String(paramWeb).trim() !== ""
        ? normalizeWebUrl(paramWeb)
        : "";
    return normalizeWebUrl(p || pc.webAbsoluteUrl || pc.siteAbsoluteUrl || "");
  }

  function normalizeGuidParam(raw) {
    if (raw == null || raw === "") return "";
    return String(raw).replace(/[{}]/g, "").trim();
  }

  function sqlEscapeListTitle(title) {
    return String(title || "").replace(/'/g, "''");
  }

  function unescapeListTitleForDisplay(sqlTitle) {
    return String(sqlTitle || "").replace(/''/g, "'");
  }

  function isBoardScaffoldField(f) {
    var title = (f.Title || "").trim();
    var iname = (f.InternalName || "").trim();
    var norm = iname.replace(/_x0020_/gi, " ").replace(/_x005f_/gi, "_");
    var lowT = title.toLowerCase();
    if (lowT.indexOf("board view choice") >= 0 || norm.toLowerCase().indexOf("board view choice") >= 0) return true;
    var ts = f.TypeAsString || "";
    if (/\schoice\d+$/i.test(title) && /^Choice$/i.test(ts) && (f.ReadOnlyField === true || f.CanBeDeleted === false))
      return true;
    return false;
  }

  function mergeStubFieldsForView(fieldsArr, viewFields, filters, orderBy, groupBy) {
    var known = Object.create(null);
    for (var i = 0; i < fieldsArr.length; i++) known[fieldsArr[i].internalName] = true;
    function addStub(name) {
      if (!name || known[name]) return;
      known[name] = true;
      fieldsArr.push({ internalName: name, title: name, typeAsString: "Text" });
    }
    var j;
    if (viewFields) for (j = 0; j < viewFields.length; j++) addStub(viewFields[j]);
    if (filters) for (j = 0; j < filters.length; j++) if (filters[j].field) addStub(filters[j].field);
    if (orderBy && orderBy.field) addStub(orderBy.field);
    if (groupBy) addStub(groupBy);
  }

  function mapODataViewRow(v) {
    return {
      id: String(v.Id || "").replace(/[{}]/g, ""),
      title: v.Title || "",
      defaultView: !!v.DefaultView,
    };
  }

  var SKIP_INTERNAL = /^_|^vti_|^ows_|^tp_/;

  function mapFieldRowToModel(f) {
    var iname = f.InternalName || f.Title || "";
    if (!iname || SKIP_INTERNAL.test(iname)) return null;
    if (isBoardScaffoldField(f)) return null;
    return { internalName: iname, title: f.Title || iname, typeAsString: f.TypeAsString || "Text" };
  }

  function rawFieldsToModels(rows) {
    var fields = [];
    for (var i = 0; i < rows.length; i++) {
      var m = mapFieldRowToModel(rows[i]);
      if (m) fields.push(m);
    }
    return fields;
  }

  function responseJsonOrThrow(response, url) {
    if (response && response.ok === false) {
      var status = response.status ? " " + response.status : "";
      var fail = function (detail) {
        var trimmed = detail ? ": " + String(detail).slice(0, 300) : "";
        throw new Error("SharePoint request failed" + status + " for " + url + trimmed);
      };
      if (typeof response.text === "function") {
        return response.text().then(fail, function () { fail(""); });
      }
      fail("");
    }
    return response.json();
  }

  function fetchODataAllPages(fetchImpl, startUrl, accept) {
    var acc = accept != null ? accept : ACCEPT_NOMETADATA;
    var accum = [];
    function step(url) {
      if (!url) return Promise.resolve(accum);
      return fetchImpl(url, { credentials: "include", headers: { Accept: acc } })
        .then(function (r) {
          return responseJsonOrThrow(r, url);
        })
        .then(function (j) {
          var batch = j.value || (j.d && j.d.results) || [];
          for (var i = 0; i < batch.length; i++) accum.push(batch[i]);
          var next = j["@odata.nextLink"] || j["odata.nextLink"];
          return step(next || null);
        });
    }
    return step(startUrl);
  }

  function fetchJson(fetchImpl, url, accept) {
    var acc = accept != null ? accept : ACCEPT_NOMETADATA;
    return fetchImpl(url, { credentials: "include", headers: { Accept: acc } }).then(function (r) {
      return responseJsonOrThrow(r, url);
    });
  }

  function resolveListContext(siteUrl, listUrl, forcedListId, fetchImpl, accept) {
    var acc = accept != null ? accept : ACCEPT_NOMETADATA;
    if (forcedListId && siteUrl) {
      return fetchJson(fetchImpl, siteUrl + "/_api/web/lists(guid'" + forcedListId + "')?$select=Title", acc).then(
        function (listData) {
          var listTitle = sqlEscapeListTitle(listData.Title || listData.title || "");
          return { listId: forcedListId, listTitle: listTitle };
        }
      );
    }
    if (!siteUrl || !listUrl) {
      return Promise.reject(new Error("MISSING_LIST_CONTEXT"));
    }
    var enc = encodeURIComponent("'" + String(listUrl).replace(/'/g, "''") + "'");
    return fetchJson(fetchImpl, siteUrl + "/_api/web/GetList(@u)/Id?@u=" + enc, acc).then(function (j) {
      var listId = String(j.value || "").replace(/[{}]/g, "");
      if (!listId) throw new Error("Could not get list ID");
      return fetchJson(fetchImpl, siteUrl + "/_api/web/lists(guid'" + listId + "')?$select=Title", acc).then(function (listData) {
        var listTitle = sqlEscapeListTitle(listData.Title || listData.title || "");
        return { listId: listId, listTitle: listTitle };
      });
    });
  }

  var FIELDS_SELECT =
    "InternalName,Title,Hidden,TypeAsString,ReadOnlyField,CanBeDeleted&$orderby=Title&$filter=Hidden eq false";

  function loadViewsAndFieldModels(siteUrl, ctx, fetchImpl, accept) {
    var acc = accept != null ? accept : ACCEPT_NOMETADATA;
    var lb = siteUrl + "/_api/web/lists(guid'" + ctx.listId + "')";
    return Promise.all([
      fetchODataAllPages(fetchImpl, lb + "/views?$select=Id,Title,DefaultView", acc),
      fetchODataAllPages(fetchImpl, lb + "/fields?$select=" + FIELDS_SELECT, acc),
    ]).then(function (results) {
      var viewRows = results[0] || [];
      var rawFields = results[1] || [];
      return {
        ctx: ctx,
        views: viewRows.map(mapODataViewRow),
        fields: rawFieldsToModels(rawFields),
      };
    });
  }

  function viewFieldsResponseToNames(vfJson) {
    var d = vfJson && vfJson.d;
    var vfItems = (vfJson && (vfJson.Items || vfJson.value || (d && d.results))) || [];
    return []
      .concat(vfItems)
      .filter(Boolean)
      .map(function (x) {
        return String(typeof x === "object" ? x.Name || x : x);
      });
  }

  function parseViewQueryParts(viewQuery) {
    var q = String(viewQuery || "").replace(/\s+/g, " ");
    var orderBy = null;
    var groupBy = null;
    var filters = [];
    var orderMatch = /<OrderBy>\s*<FieldRef\s+Name="([^"]+)"\s+Ascending="(True|False)"/i.exec(q);
    if (orderMatch) orderBy = { field: orderMatch[1], ascending: orderMatch[2].toLowerCase() === "true" };
    var groupMatch = /<GroupBy>\s*<FieldRef\s+Name="([^"]+)"/i.exec(q);
    if (groupMatch) groupBy = groupMatch[1];
    var condRegex =
      /<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)>\s*<FieldRef\s+Name="([^"]+)"\s*\/>\s*<Value\s+Type="([^"]*)">([^<]*)<\/Value>/gi;
    var m;
    while ((m = condRegex.exec(q)) !== null) {
      filters.push({
        op: m[1],
        field: m[2],
        valueType: m[3] || "Text",
        value: (m[4] || "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">"),
      });
    }
    return { viewQuery: q, orderBy: orderBy, filters: filters, groupBy: groupBy };
  }

  function loadSingleViewDetails(siteUrl, ctx, viewId, fields, fetchImpl, accept) {
    var acc = accept != null ? accept : ACCEPT_NOMETADATA;
    var lb = siteUrl + "/_api/web/lists(guid'" + ctx.listId + "')";
    var vBase = lb + "/views(guid'" + viewId + "')";
    return Promise.all([
      fetchJson(fetchImpl, vBase + "/ViewFields", acc),
      fetchJson(fetchImpl, vBase + "?$select=ViewQuery,Title", acc),
    ]).then(function (results) {
      var vfJson = results[0];
      var metaJson = results[1];
      var viewFields = viewFieldsResponseToNames(vfJson);
      var viewQuery = (metaJson.ViewQuery || "").replace(/\s+/g, " ");
      var viewTitle = metaJson.Title || "";
      var parsed = parseViewQueryParts(viewQuery);
      var orderBy = parsed.orderBy;
      var filters = parsed.filters;
      var groupBy = parsed.groupBy;
      var fieldsOut = fields.slice();
      mergeStubFieldsForView(fieldsOut, viewFields, filters, orderBy, groupBy);
      return {
        viewFields: viewFields,
        viewQuery: viewQuery,
        viewTitle: viewTitle,
        orderBy: orderBy,
        filters: filters,
        groupBy: groupBy,
        fieldsOut: fieldsOut,
      };
    });
  }

  function sendResult(data) {
    window.postMessage(Object.assign({ __spcsv: true, type: RESULT_TYPE }, data), "*");
  }

  async function run() {
    var el = document.getElementById("sp-views-params");
    var params = parseInjectParams(el && el.textContent);
    var pageContext = window._spPageContextInfo || {};
    var accept = ACCEPT_NOMETADATA;
    var siteUrl = resolveSiteUrl(params.webAbsoluteUrl, pageContext);
    var listUrl = pageContext.listUrl || pageContext.listServerRelativeUrl;
    var forcedListId = normalizeGuidParam(params.listId);

    var ctx;
    try {
      ctx = await resolveListContext(siteUrl, listUrl, forcedListId || null, fetch, accept);
    } catch (e) {
      if (e && e.message === "MISSING_LIST_CONTEXT") {
        sendResult({ error: "Open a list or library page.", views: [], fields: [] });
        return;
      }
      sendResult({ error: (e && e.message) || String(e), views: [], fields: [] });
      return;
    }

    var data;
    try {
      data = await loadViewsAndFieldModels(siteUrl, ctx, fetch, accept);
    } catch (e) {
      sendResult({ error: (e && e.message) || String(e), views: [], fields: [] });
      return;
    }

    var viewId = params.viewId ? normalizeGuidParam(params.viewId) : "";
    if (!viewId) {
      sendResult({
        views: data.views,
        fields: data.fields,
        listId: data.ctx.listId,
        listTitle: unescapeListTitleForDisplay(data.ctx.listTitle),
      });
      return;
    }

    try {
      var details = await loadSingleViewDetails(siteUrl, data.ctx, viewId, data.fields, fetch, accept);
      sendResult({
        views: data.views,
        fields: details.fieldsOut,
        listId: data.ctx.listId,
        listTitle: unescapeListTitleForDisplay(data.ctx.listTitle),
        viewDetails: {
          viewFields: details.viewFields,
          viewQuery: details.viewQuery,
          viewTitle: details.viewTitle,
          orderBy: details.orderBy,
          filters: details.filters,
          groupBy: details.groupBy,
        },
      });
    } catch (err) {
      sendResult({
        views: data.views,
        fields: data.fields,
        listId: data.ctx.listId,
        listTitle: unescapeListTitleForDisplay(data.ctx.listTitle),
        error: (err && err.message) || "Failed to load view details",
      });
    }
  }

  run().catch(function (e) {
    window.postMessage(
      {
        __spcsv: true,
        type: RESULT_TYPE,
        error: (e && e.message) || String(e),
        views: [],
        fields: [],
      },
      "*"
    );
  });
})();
