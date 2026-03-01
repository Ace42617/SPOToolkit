// Injected: fetches managedproperty.aspx and parses mapped/crawled properties. Reads params from #sp-refinable-params.
(function () {
  var el = document.getElementById("sp-refinable-params");
  if (!el || !el.textContent) {
    window.postMessage({ __spcsv: true, type: "SPCSVRefinableMappingsResult", ok: false, error: "No params", mappings: [], alias: null }, "*");
    return;
  }
  var params;
  try {
    params = JSON.parse(el.textContent);
  } catch (e) {
    window.postMessage({ __spcsv: true, type: "SPCSVRefinableMappingsResult", ok: false, error: "Invalid params", mappings: [], alias: null }, "*");
    return;
  }
  var siteUrl = (params.siteUrl || "").replace(/\/$/, "");
  var propertyName = params.propertyName || "";
  if (!siteUrl || !propertyName) {
    window.postMessage({ __spcsv: true, type: "SPCSVRefinableMappingsResult", ok: false, error: "Missing siteUrl or propertyName", mappings: [], alias: null }, "*");
    return;
  }
  var navTerms = ["managed properties", "crawled properties", "categories"];
  function isNavTerm(text) {
    var t = (text || "").toLowerCase().trim();
    return navTerms.indexOf(t) >= 0;
  }
  function looksLikeCrawledProperty(name) {
    if (!name || name.length > 120) return false;
    var t = name.trim();
    if (t.indexOf(" ") >= 0) return false;
    return /^ows_[a-zA-Z0-9_]+$/i.test(t);
  }
  function extractCrawledNamesFromText(text) {
    var names = [];
    if (!text || typeof text !== "string") return names;
    var m;
    var re = /ows_[a-zA-Z0-9_]+/gi;
    while ((m = re.exec(text)) !== null) names.push(m[0]);
    return names;
  }
  function filterCrawledOnly(items) {
    var out = [];
    for (var i = 0; i < items.length; i++) if (looksLikeCrawledProperty(items[i].name)) out.push(items[i]);
    return out;
  }
  function collectFromTable(table, requireLink) {
    var found = [];
    var rows = table.querySelectorAll("tr");
    var startRow = 0;
    if (rows.length > 1) {
      var firstText = (rows[0].textContent || "").toLowerCase();
      if (firstText.indexOf("crawled") >= 0 || firstText.indexOf("mapping") >= 0 || firstText.indexOf("name") >= 0)
        startRow = 1;
    }
    for (var r = startRow; r < rows.length; r++) {
      var cells = rows[r].querySelectorAll("td");
      if (cells.length === 0) continue;
      var name = "";
      var nameCell = cells[0];
      if (nameCell) {
        var link = nameCell.querySelector("a");
        if (requireLink && !link) continue;
        name = (link ? link.textContent : nameCell.textContent || "").trim();
      }
      if (!name || isNavTerm(name)) continue;
      var type = cells.length > 1 ? (cells[1].textContent || "").trim() : "";
      found.push({ name: name, type: type });
    }
    return found;
  }
  function findMappings(doc) {
    var mappings = [];
    var body = doc.body || doc.documentElement;
    if (!body) return mappings;
    var bodyText = (body.textContent || "").toLowerCase();
    if (bodyText.indexOf("access denied") >= 0 || bodyText.indexOf("sign in") >= 0) return mappings;

    function nextTableAfter(node) {
      var n = node.nextSibling;
      for (var i = 0; i < 25; i++) {
        if (!n) {
          n = node.parentNode;
          node = n;
          if (!n || n === body) break;
          n = n.nextSibling;
          continue;
        }
        if (n.nodeType === 1) {
          if (n.tagName === "TABLE") return n;
          var tbl = n.querySelector && n.querySelector("table");
          if (tbl) return tbl;
        }
        n = n.nextSibling;
      }
      return null;
    }
    function allTablesAfter(node, maxTables) {
      var tables = [];
      var n = node.nextSibling;
      var count = 0;
      while (n && count < (maxTables || 5)) {
        if (n.nodeType === 1) {
          if (n.tagName === "TABLE") {
            tables.push(n);
            count++;
          } else {
            var tbls = n.querySelectorAll && n.querySelectorAll("table");
            if (tbls) for (var i = 0; i < tbls.length && count < (maxTables || 5); i++) {
              tables.push(tbls[i]);
              count++;
            }
          }
        }
        n = n.nextSibling;
        if (!n && node.parentNode && node.parentNode !== body) {
          n = node.parentNode.nextSibling;
          node = node.parentNode;
        }
      }
      return tables;
    }
    function sectionTextAfter(node, maxChars) {
      var out = [];
      var len = 0;
      var limit = maxChars || 8000;
      var n = node.nextSibling;
      while (n && len < limit) {
        if (n.nodeType === 3) {
          out.push(n.textContent);
          len += (n.textContent || "").length;
        } else if (n.nodeType === 1) {
          var t = n.textContent || "";
          out.push(t);
          len += t.length;
        }
        n = n.nextSibling;
        if (!n && node.parentNode && node.parentNode !== body) {
          n = node.parentNode.nextSibling;
          node = node.parentNode;
        }
      }
      return out.join("");
    }
    var walker = doc.createTreeWalker(body, 1, null, false);
    var mappingsLabel = null;
    while (walker.nextNode()) {
      var node = walker.currentNode;
      var text = (node.textContent || "").trim();
      if (text.toLowerCase().indexOf("mappings to crawled") >= 0 || (text === "Crawled Properties" && node.querySelector && !node.querySelector("table"))) {
        if (node.childNodes.length <= 15) {
          mappingsLabel = node;
          break;
        }
        if (!mappingsLabel) mappingsLabel = node;
      }
    }
    if (mappingsLabel) {
      var seen = {};
      var tbls = allTablesAfter(mappingsLabel, 5);
      for (var ti = 0; ti < tbls.length; ti++) {
        var raw = collectFromTable(tbls[ti], false);
        for (var ri = 0; ri < raw.length; ri++) {
          var nam = raw[ri].name;
          if (looksLikeCrawledProperty(nam) && !seen[nam]) {
            seen[nam] = true;
            mappings.push(raw[ri]);
          }
          var extracted = extractCrawledNamesFromText(nam);
          for (var ei = 0; ei < extracted.length; ei++) {
            var ex = extracted[ei];
            if (!seen[ex]) {
              seen[ex] = true;
              mappings.push({ name: ex, type: raw[ri].type || "" });
            }
          }
        }
        var cells = tbls[ti].querySelectorAll("td");
        for (var ci = 0; ci < cells.length; ci++) {
          var cellText = (cells[ci].textContent || "").trim();
          var names = extractCrawledNamesFromText(cellText);
          for (var ni = 0; ni < names.length; ni++) {
            if (!seen[names[ni]]) {
              seen[names[ni]] = true;
              mappings.push({ name: names[ni], type: "" });
            }
          }
        }
      }
      var sectionTxt = sectionTextAfter(mappingsLabel, 6000);
      var namesInSection = extractCrawledNamesFromText(sectionTxt);
      for (var si = 0; si < namesInSection.length; si++) {
        var sn = namesInSection[si];
        if (!seen[sn]) {
          seen[sn] = true;
          mappings.push({ name: sn, type: "" });
        }
      }
      if (mappings.length > 0) return mappings;
      var tbl = nextTableAfter(mappingsLabel);
      if (!tbl && mappingsLabel.parentNode) {
        var parent = mappingsLabel.parentNode;
        if (parent.querySelector) {
          var inParent = parent.querySelector("table");
          if (inParent && inParent !== mappingsLabel && mappingsLabel.compareDocumentPosition(inParent) === 4) tbl = inParent;
        }
      }
      if (!tbl) tbl = mappingsLabel.querySelector && mappingsLabel.querySelector("table");
      if (tbl) {
        var fromTbl = filterCrawledOnly(collectFromTable(tbl, false));
        if (fromTbl.length > 0) return fromTbl;
      }
    }
    var bodyFullText = (body.textContent || "").trim();
    var fallbackNames = extractCrawledNamesFromText(bodyFullText);
    if (fallbackNames.length === 0) {
      var bodyHtml = (body.innerHTML || "").trim();
      fallbackNames = extractCrawledNamesFromText(bodyHtml);
    }
    if (fallbackNames.length === 0) {
      var links = body.querySelectorAll('a[href*="crawledproperty"], a[href*="CrawledProperty"]');
      for (var li = 0; li < links.length; li++) {
        var linkText = (links[li].textContent || "").trim();
        if (looksLikeCrawledProperty(linkText)) fallbackNames.push(linkText);
      }
    }
    if (fallbackNames.length === 0) {
      var textareas = body.querySelectorAll("textarea, [role='listbox'], select");
      for (var ta = 0; ta < textareas.length; ta++) {
        var t = (textareas[ta].value || textareas[ta].textContent || "").trim();
        var fromTa = extractCrawledNamesFromText(t);
        for (var ft = 0; ft < fromTa.length; ft++) fallbackNames.push(fromTa[ft]);
      }
    }
    if (fallbackNames.length > 0) {
      var seen = {};
      for (var fi = 0; fi < fallbackNames.length; fi++) {
        if (!seen[fallbackNames[fi]]) {
          seen[fallbackNames[fi]] = true;
          mappings.push({ name: fallbackNames[fi], type: "" });
        }
      }
      return mappings;
    }
    var tables = doc.querySelectorAll("table");
    for (var t = 0; t < tables.length; t++) {
      var table = tables[t];
      var rows = table.querySelectorAll("tr");
      if (rows.length < 1) continue;
      var firstRowCells = rows[0].querySelectorAll("td, th");
      var firstRowText = (rows[0].textContent || "").toLowerCase();
      if (firstRowCells.length <= 3 && firstRowText.indexOf("managed properties") >= 0 && firstRowText.indexOf("crawled properties") >= 0 && firstRowText.indexOf("categories") >= 0)
        continue;
      var withLinks = filterCrawledOnly(collectFromTable(table, true));
      if (withLinks.length > 0) return withLinks;
      var withAny = filterCrawledOnly(collectFromTable(table, false));
      if (withAny.length > 0) return withAny;
    }
    var lists = doc.querySelectorAll("ul, ol");
    for (var L = 0; L < lists.length; L++) {
      var list = lists[L];
      var items = list.querySelectorAll("li");
      for (var i = 0; i < items.length; i++) {
        var name = (items[i].textContent || "").trim();
        if (name && looksLikeCrawledProperty(name)) mappings.push({ name: name, type: "" });
      }
      if (mappings.length > 0) return mappings;
    }
    return mappings;
  }
  function getAlias(doc) {
    var body = doc.body || doc.documentElement;
    if (!body) return null;
    var inputs = body.querySelectorAll && body.querySelectorAll("input[type='text'], input:not([type])");
    if (inputs) {
      for (var i = 0; i < inputs.length; i++) {
        var inp = inputs[i];
        var id = (inp.id || "").toLowerCase();
        var name = (inp.name || "").toLowerCase();
        if ((id.indexOf("alias") >= 0 || name.indexOf("alias") >= 0) && inp.value) return inp.value.trim();
      }
    }
    var walker = doc.createTreeWalker(body, 1, null, false);
    while (walker.nextNode()) {
      var node = walker.currentNode;
      var text = (node.textContent || "").trim();
      if (text === "Alias:" || (text === "Alias" || (text.toLowerCase().indexOf("alias:") === 0 && text.length < 50))) {
        var next = node.nextSibling;
        for (var i = 0; i < 5 && next; i++) {
          if (next.nodeType === 1) {
            var input = next.querySelector && next.querySelector("input[type='text'], input:not([type])");
            if (input && input.value) return input.value.trim();
            var val = (next.textContent || "").trim();
            if (val && val !== text && val.length < 200 && val.indexOf("Define an alias") < 0) return val;
          }
          next = next.nextSibling;
        }
        var parent = node.parentNode;
        if (parent && parent.querySelector) {
          var inps = parent.querySelectorAll("input[type='text'], input:not([type])");
          for (var j = 0; j < inps.length; j++) if (inps[j].value) return inps[j].value.trim();
        }
        var tr = node;
        while (tr && tr.tagName !== "TR") tr = tr.parentNode;
        if (tr) {
          var cells = tr.querySelectorAll("td");
          for (var c = 0; c < cells.length - 1; c++) {
            if ((cells[c].textContent || "").trim().toLowerCase().indexOf("alias") >= 0) {
              var nextCell = cells[c + 1];
              var inp = nextCell.querySelector && nextCell.querySelector("input[type='text'], input:not([type])");
              if (inp && inp.value) return inp.value.trim();
              var v = (nextCell.textContent || "").trim();
              if (v && v.length < 200 && v.toLowerCase().indexOf("define an alias") < 0) return v;
              break;
            }
          }
        }
        break;
      }
    }
    return null;
  }
  var url = siteUrl + "/_layouts/15/managedproperty.aspx?property=" + encodeURIComponent(propertyName) + "&level=sitecol";
  fetch(url, { credentials: "include", headers: { "Accept": "text/html" } })
    .then(function (r) { return r.text(); })
    .then(function (html) {
      var mappings = [];
      var alias = null;
      try {
        var parser = new DOMParser();
        var doc = parser.parseFromString(html, "text/html");
        mappings = findMappings(doc);
        alias = getAlias(doc);
        if (mappings.length === 0) {
          var bodyText = (doc.body && doc.body.textContent || "").toLowerCase();
          if (bodyText.indexOf("access denied") >= 0 || bodyText.indexOf("sign in") >= 0) {
            window.postMessage({ __spcsv: true, type: "SPCSVRefinableMappingsResult", ok: false, error: "Access denied or sign-in required", mappings: [], alias: null }, "*");
            return;
          }
        }
      } catch (parseErr) {
        window.postMessage({ __spcsv: true, type: "SPCSVRefinableMappingsResult", ok: false, error: "Parse error: " + (parseErr.message || ""), mappings: [], alias: null }, "*");
        return;
      }
      window.postMessage({ __spcsv: true, type: "SPCSVRefinableMappingsResult", ok: true, mappings: mappings, alias: alias, error: null }, "*");
    })
    .catch(function (err) {
      window.postMessage({ __spcsv: true, type: "SPCSVRefinableMappingsResult", ok: false, error: (err && err.message) || String(err), mappings: [], alias: null }, "*");
    });
})();
