// Injected: View Manager REST loader. Behavior must match lib/viewsDataCore.mjs — run `npm test` after edits.
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

  function mergeStubFieldsForView(fieldsArr, viewFields, filters, orderBy, groupBy, orderByLevels, groupByLevels) {
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
    if (orderByLevels && orderByLevels.length) {
      for (j = 0; j < orderByLevels.length; j++) if (orderByLevels[j] && orderByLevels[j].field) addStub(orderByLevels[j].field);
    } else if (orderBy && orderBy.field) addStub(orderBy.field);
    if (groupByLevels && groupByLevels.length) {
      for (j = 0; j < groupByLevels.length; j++) {
        var g = groupByLevels[j];
        addStub(g && g.field ? g.field : g);
      }
    } else if (groupBy) addStub(groupBy);
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

  function fetchODataAllPages(fetchImpl, startUrl, accept) {
    var acc = accept != null ? accept : ACCEPT_NOMETADATA;
    var accum = [];
    function step(url) {
      if (!url) return Promise.resolve(accum);
      return fetchImpl(url, { credentials: "include", headers: { Accept: acc } })
        .then(function (r) {
          return r.json();
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
      return r.json();
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

  function parseCamlFieldRefs(xml) {
    var refs = [];
    var re = /<FieldRef\b([^>]*?)\/?\s*>/gi;
    var m;
    while ((m = re.exec(String(xml || ""))) !== null) {
      var attrs = m[1] || "";
      var nameM = /\bName\s*=\s*["']([^"']+)["']/i.exec(attrs);
      if (!nameM) continue;
      var ascM = /\bAscending\s*=\s*["'](True|False)["']/i.exec(attrs);
      refs.push({
        field: nameM[1],
        ascending: ascM ? ascM[1].toLowerCase() === "true" : true,
      });
    }
    return refs;
  }

  function camlAttrTrueFalse(attrs, name) {
    var re = new RegExp("\\b" + name + "\\s*=\\s*[\"'](True|False)[\"']", "i");
    var m = re.exec(attrs || "");
    if (!m) return null;
    return m[1].toLowerCase() === "true";
  }

  function parseViewQueryParts(viewQuery) {
    var q = String(viewQuery || "").replace(/\s+/g, " ");
    var filters = [];
    var orderByLevels = [];
    var orderBlock = /<OrderBy\b[^>]*>([\s\S]*?)<\/OrderBy>/i.exec(q);
    if (orderBlock) orderByLevels = parseCamlFieldRefs(orderBlock[1]);
    var orderBy = orderByLevels.length ? orderByLevels[0] : null;

    var groupBy = null;
    var groupByLevels = [];
    var groupExpand = true;
    var groupLimit = null;
    var groupBlock = /<GroupBy\b([^>]*)>([\s\S]*?)<\/GroupBy>/i.exec(q);
    if (groupBlock) {
      var gAttrs = groupBlock[1] || "";
      var collapse = camlAttrTrueFalse(gAttrs, "Collapse");
      groupExpand = collapse === false;
      var lim = /\bGroupLimit\s*=\s*["'](\d+)["']/i.exec(gAttrs);
      if (lim) groupLimit = parseInt(lim[1], 10);
      groupByLevels = parseCamlFieldRefs(groupBlock[2]);
      if (groupByLevels.length) groupBy = groupByLevels[0].field;
    }

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
    return {
      viewQuery: q,
      orderBy: orderBy,
      orderByLevels: orderByLevels,
      filters: filters,
      groupBy: groupBy,
      groupByLevels: groupByLevels,
      groupExpand: groupExpand,
      groupLimit: groupLimit,
    };
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
      var orderByLevels = parsed.orderByLevels;
      var filters = parsed.filters;
      var groupBy = parsed.groupBy;
      var groupByLevels = parsed.groupByLevels;
      var groupExpand = parsed.groupExpand;
      var groupLimit = parsed.groupLimit;
      var fieldsOut = fields.slice();
      mergeStubFieldsForView(fieldsOut, viewFields, filters, orderBy, groupBy, orderByLevels, groupByLevels);
      return {
        viewFields: viewFields,
        viewQuery: viewQuery,
        viewTitle: viewTitle,
        orderBy: orderBy,
        orderByLevels: orderByLevels,
        filters: filters,
        groupBy: groupBy,
        groupByLevels: groupByLevels,
        groupExpand: groupExpand,
        groupLimit: groupLimit,
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
          orderByLevels: details.orderByLevels,
          filters: details.filters,
          groupBy: details.groupBy,
          groupByLevels: details.groupByLevels,
          groupExpand: details.groupExpand,
          groupLimit: details.groupLimit,
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
