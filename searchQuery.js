// Injected: runs SharePoint search postquery, returns rows with all requested managed properties. Reads params from #sp-search-query-params.
(function () {
  var el = document.getElementById("sp-search-query-params");
  if (!el || !el.textContent) {
    window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: false, error: "No params", rows: [] }, "*");
    return;
  }
  var params;
  try {
    params = JSON.parse(el.textContent);
  } catch (e) {
    window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: false, error: "Invalid params", rows: [] }, "*");
    return;
  }
  var siteUrl = (params.siteUrl || "").replace(/\/$/, "");
  var queryText = (params.queryText || "*").trim() || "*";
  var rowLimit = Math.min(Math.max(parseInt(params.rowLimit, 10) || 25, 1), 500);

  function buildSelectProperties() {
    var props = [
      "Title", "Path", "DefaultEncodingURL", "DocId", "Author", "Size", "Write", "Created", "LastModifiedTime",
      "HitHighlightedSummary", "Rank", "ListItemID", "UniqueId", "SiteTitle", "WebId", "SiteId", "ContentTypeId", "ContentType",
      "Filename", "Editor", "ModifiedBy", "CreatedBy", "ParentLink", "ParentTitle", "ListID", "ListTitle", "IsDocument", "OriginalPath",
      "SPSiteUrl", "WebUrl", "SiteName", "CollapsingStatus", "HitHighlightedProperties", "ViewContentTypeId"
    ];
    var i, s;
    for (i = 0; i < 100; i++) {
      s = i < 10 ? "0" + i : "" + i;
      props.push("RefinableString" + s);
    }
    for (i = 100; i < 200; i++) props.push("RefinableString" + i);
    for (i = 0; i < 20; i++) {
      s = i < 10 ? "0" + i : "" + i;
      props.push("RefinableDate" + s);
    }
    for (i = 0; i < 50; i++) {
      s = i < 10 ? "0" + i : "" + i;
      props.push("RefinableInt" + s);
    }
    for (i = 0; i < 10; i++) {
      s = i < 10 ? "0" + i : "" + i;
      props.push("RefinableDecimal" + s);
    }
    for (i = 0; i < 10; i++) {
      s = i < 10 ? "0" + i : "" + i;
      props.push("RefinableDouble" + s);
    }
    for (i = 0; i < 5; i++) {
      s = i < 10 ? "0" + i : "" + i;
      props.push("RefinableYesNo" + s);
    }
    return props;
  }

  var url = siteUrl + "/_api/search/postquery";
  var requestBody = {
    request: {
      Querytext: queryText,
      RowLimit: rowLimit,
      SelectProperties: { results: buildSelectProperties() }
    }
  };

  function parseResultRows(data) {
    var rows = [];
    var d = data && data.d;
    if (!d) return { rows: null, error: "No 'd' in response. Top-level keys: " + (data ? Object.keys(data).join(", ") : "null") };
    var payload = d.postquery || d.query || d;
    if (!payload) return { rows: null, error: "No postquery/query in d. Keys: " + Object.keys(d).join(", ") };
    var primary = payload.PrimaryQueryResult;
    if (!primary) {
      var msg = (payload.ErrorCode ? "ErrorCode " + payload.ErrorCode + " " : "") || "";
      if (payload.Messages && payload.Messages.results && payload.Messages.results[0]) msg += (payload.Messages.results[0].Value || "");
      return { rows: null, error: msg || "No PrimaryQueryResult. Payload keys: " + Object.keys(payload).join(", ") };
    }
    var table = primary.RelevantResults && primary.RelevantResults.Table;
    if (!table || !table.Rows) return { rows: [], totalRows: (primary.RelevantResults && primary.RelevantResults.TotalRows) || 0 };
    var rowResults = table.Rows.results != null ? table.Rows.results : table.Rows;
    if (!Array.isArray(rowResults)) rowResults = rowResults ? [rowResults] : [];
    for (var r = 0; r < rowResults.length; r++) {
      var row = rowResults[r];
      var cells = row.Cells && (row.Cells.results != null ? row.Cells.results : row.Cells);
      if (!Array.isArray(cells)) cells = cells ? [cells] : [];
      var obj = {};
      for (var c = 0; c < cells.length; c++) {
        var key = cells[c].Key;
        var val = cells[c].Value;
        if (key != null && val != null && String(val).trim() !== "") obj[key] = val;
      }
      rows.push(obj);
    }
    var total = (primary.RelevantResults && primary.RelevantResults.TotalRows) || rows.length;
    return { rows: rows, totalRows: total };
  }

  function parseXmlResult(text) {
    var rows = [];
    function byLocalName(parent, localName) {
      if (!parent) return null;
      var list = parent.getElementsByTagName("*");
      for (var i = 0; i < list.length; i++) {
        var el = list[i];
        if ((el.localName || el.tagName || "").replace(/^[^:]+:/, "") === localName) return el;
      }
      return null;
    }
    function directChildrenByLocalName(parent, localName) {
      var out = [];
      if (!parent || !parent.childNodes) return out;
      for (var i = 0; i < parent.childNodes.length; i++) {
        var el = parent.childNodes[i];
        if (el.nodeType !== 1) continue;
        var name = (el.localName || el.tagName || "").replace(/^[^:]+:/, "");
        if (name === localName) out.push(el);
      }
      return out;
    }
    try {
      var parser = new DOMParser();
      var doc = parser.parseFromString(text, "text/xml");
      if (doc.querySelector("parsererror")) return { rows: null, error: "XML parse error" };
      var query = doc.documentElement;
      var primary = byLocalName(query, "PrimaryQueryResult");
      if (!primary) return { rows: null, error: "No PrimaryQueryResult in XML" };
      var relevant = byLocalName(primary, "RelevantResults");
      if (!relevant) return { rows: [], totalRows: 0 };
      var table = byLocalName(relevant, "Table");
      var totalEl = byLocalName(relevant, "RowCount");
      var totalRows = totalEl ? parseInt(totalEl.textContent, 10) || 0 : 0;
      if (!table) return { rows: [], totalRows: totalRows };
      var rowsContainer = byLocalName(table, "Rows");
      var rowEls = rowsContainer ? directChildrenByLocalName(rowsContainer, "element") : [];
      for (var i = 0; i < rowEls.length; i++) {
        var rowEl = rowEls[i];
        var cellsEl = byLocalName(rowEl, "Cells");
        if (!cellsEl) continue;
        var cellEls = directChildrenByLocalName(cellsEl, "element");
        var obj = {};
        for (var j = 0; j < cellEls.length; j++) {
          var keyEl = byLocalName(cellEls[j], "Key");
          var valEl = byLocalName(cellEls[j], "Value");
          var key = keyEl ? keyEl.textContent : "";
          var val = valEl ? valEl.textContent : "";
          if (key && val !== undefined && String(val).trim() !== "") obj[key.trim()] = String(val).trim();
        }
        if (Object.keys(obj).length > 0) rows.push(obj);
      }
      return { rows: rows, totalRows: totalRows || rows.length };
    } catch (e) {
      return { rows: null, error: "XML parse: " + (e.message || String(e)) };
    }
  }

  function runGetSearch() {
    var getUrl = siteUrl + "/_api/search/query?querytext='" + encodeURIComponent(queryText) + "'&rowlimit=" + rowLimit + "&selectproperties='Title,Path,Author,Size,Write'";
    fetch(getUrl, { method: "GET", credentials: "include", headers: { "Accept": "application/json;odata=verbose" } })
      .then(function (r) { return r.text().then(function (text) { return { ok: r.ok, status: r.status, text: text }; }); })
      .then(function (result) {
        var debug = { method: "GET", url: getUrl, status: result.status, preview: result.text ? result.text.substring(0, 500) : "" };
        if (result.text && result.text.length > 500) debug.preview += "\n... [truncated " + (result.text.length - 500) + " chars]";
        try {
          var j = JSON.parse(result.text);
          if (j && j.d) debug.dKeys = Object.keys(j.d).join(", ");
        } catch (_) {}
        if (!result.ok) {
          window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: false, error: "GET search failed: HTTP " + result.status, rows: [], debug: debug }, "*");
          return;
        }
        var data;
        var parsed;
        try {
          data = JSON.parse(result.text);
          parsed = parseResultRows(data);
        } catch (e) {
          if (/^\s*<\?xml|^\s*</.test(result.text)) {
            parsed = parseXmlResult(result.text);
            if (parsed.rows !== null) {
              window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: true, rows: parsed.rows || [], totalRows: parsed.totalRows || 0, error: null, debug: debug }, "*");
              return;
            }
          }
          window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: false, error: "GET: Invalid JSON", rows: [], debug: debug }, "*");
          return;
        }
        if (parsed.error && parsed.rows === null) {
          window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: false, error: "GET " + parsed.error, rows: [], debug: debug }, "*");
          return;
        }
        window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: true, rows: parsed.rows || [], totalRows: parsed.totalRows || 0, error: null, debug: debug }, "*");
      })
      .catch(function (err) {
        window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: false, error: "GET " + ((err && err.message) || String(err)), rows: [], debug: { method: "GET", url: getUrl, error: String(err) } }, "*");
      });
  }

  function makeDebug(result, data) {
    var debug = { method: "POST", url: url, status: result.status };
    if (result.text) {
      debug.preview = result.text.substring(0, 500);
      if (result.text.length > 500) debug.preview += "\n... [truncated " + (result.text.length - 500) + " chars]";
    }
    if (data && data.d) debug.dKeys = Object.keys(data.d).join(", ");
    return debug;
  }

  fetch(url, {
    method: "POST",
    credentials: "include",
    headers: {
      "Content-Type": "application/json",
      "Accept": "application/json;odata=verbose"
    },
    body: JSON.stringify(requestBody)
  })
    .then(function (r) {
      return r.text().then(function (text) {
        return { ok: r.ok, status: r.status, text: text };
      });
    })
    .then(function (result) {
      if (!result.ok) {
        runGetSearch();
        return;
      }
      var data;
      var parsed;
      try {
        data = JSON.parse(result.text);
        parsed = parseResultRows(data);
      } catch (e) {
        if (/^\s*<\?xml|^\s*</.test(result.text)) {
          parsed = parseXmlResult(result.text);
          if (parsed.rows !== null) {
            window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: true, rows: parsed.rows, totalRows: parsed.totalRows || 0, error: null, debug: makeDebug(result, null) }, "*");
            return;
          }
        }
        window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: false, error: "Invalid JSON response", rows: [], debug: makeDebug(result, null) }, "*");
        return;
      }
      if (parsed.rows === null) {
        if (/^\s*<\?xml|^\s*</.test(result.text)) {
          var xmlParsed = parseXmlResult(result.text);
          if (xmlParsed.rows !== null) {
            window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: true, rows: xmlParsed.rows, totalRows: xmlParsed.totalRows || 0, error: null, debug: makeDebug(result, null) }, "*");
            return;
          }
        }
        runGetSearch();
        return;
      }
      window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: true, rows: parsed.rows, totalRows: parsed.totalRows || 0, error: null, debug: makeDebug(result, data) }, "*");
    })
    .catch(function (err) {
      window.postMessage({ __spcsv: true, type: "SPCSVSearchQueryResult", ok: false, error: (err && err.message) || String(err), rows: [], debug: { method: "POST", url: url, error: String(err) } }, "*");
    });
})();
