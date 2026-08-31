// Injected page script: site permissions matrix (aligned with Export-SitePermissionsMatrix.ps1 defaults).
// Also powers Everything Bagel (same scan + Folder Counts / Path Lengths sheets).
(function () {
  function readParams() {
    try {
      var el = document.getElementById("spcsv-params-json");
      if (el && el.textContent) return JSON.parse(el.textContent);
    } catch (_) {}
    return null;
  }

  var params = readParams();
  if (!params) {
    console.error("Permissions matrix export – missing params.");
    return;
  }

  var IS_BAGEL = params.report === "everythingBagel";
  var REPORT_KEY = IS_BAGEL ? "everythingBagel" : "permissionsMatrix";
  var REPORT_LABEL = IS_BAGEL ? "Everything Bagel" : "Permissions matrix";

  try {
    window.postMessage({
      __spcsv: true,
      type: "SPCSVExportStarted",
      detail: { message: REPORT_LABEL + " export started…", report: REPORT_KEY }
    }, "*");
  } catch (_) {}

  var exportCancelled = false;
  window.__SPOToolkitExportCancel = false;
  window.addEventListener("message", function (ev) {
    if (ev.data && ev.data.__spcsv === true && ev.data.type === "SPCSVExportCancel") exportCancelled = true;
  });
  try {
    window.addEventListener("spotoolkit-export-cancel", function () { exportCancelled = true; });
  } catch (_) {}

  function checkCancelled() {
    if (exportCancelled || window.__SPOToolkitExportCancel) throw new Error("Export cancelled.");
  }

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  var overallPct = 0;
  var progressPlan = { totalItems: 0, doneItems: 0, currentListLoaded: 0, totalLists: 0, websTotal: 0 };

  function fmtProgressCount(n) {
    n = Math.max(0, Math.round(n || 0));
    try { return n.toLocaleString(); } catch (_) { return String(n); }
  }

  function reportProgress(message, opts) {
    opts = opts || {};
    var isPulse = !!opts.pulse;
    if (!isPulse) {
      lastProgressAt = Date.now();
      lastProgressMessage = message;
    }
    var detail = { message: message, logLine: message, report: REPORT_KEY };
    if (opts.logLine) detail.logLine = opts.logLine;
    if (isPulse) {
      detail.pulse = true;
      detail.message = lastProgressMessage || message;
      delete detail.logLine;
    }
    if (opts.percent != null) {
      overallPct = Math.max(overallPct, Math.min(99, Math.round(opts.percent)));
      detail.percent = overallPct;
    } else {
      detail.percent = overallPct;
    }
    window.postMessage({ __spcsv: true, type: "SPCSVExportProgress", detail: detail }, "*");
  }

  function progressHeadline(pct) {
    var p = pct != null ? Math.round(pct) : overallPct;
    if (progressPlan.totalItems > 0) {
      var done = Math.min(progressPlan.doneItems + progressPlan.currentListLoaded, progressPlan.totalItems);
      return p + "% — " + fmtProgressCount(done) + " of " + fmtProgressCount(progressPlan.totalItems) + " items";
    }
    return p + "% complete";
  }

  function itemProgressPct(loadedInCurrentList) {
    var partial = progressPlan.doneItems + Math.max(0, loadedInCurrentList || 0);
    var frac = partial / Math.max(1, progressPlan.totalItems);
    return 5 + frac * 90;
  }

  function setProgress(logLine, pct) {
    var nextPct = Math.max(overallPct, Math.min(99, Math.round(pct)));
    reportProgress(progressHeadline(nextPct), { percent: nextPct, logLine: logLine });
  }

  function logProgress(logLine) {
    reportProgress(progressHeadline(overallPct), { logLine: logLine });
  }

  var lastProgressAt = Date.now();
  var lastProgressMessage = "";
  var progressPulseTimer = null;

  function startProgressPulse() {
    if (progressPulseTimer) return;
    progressPulseTimer = setInterval(function () {
      if (Date.now() - lastProgressAt < 4000) return;
      if (!lastProgressMessage) return;
      reportProgress(lastProgressMessage || progressHeadline(overallPct), { percent: overallPct, pulse: true });
    }, 4000);
  }

  function stopProgressPulse() {
    if (progressPulseTimer) {
      clearInterval(progressPulseTimer);
      progressPulseTimer = null;
    }
  }

  function reportDone(success, message) {
    if (!success) {
      console.error(REPORT_LABEL + ":", message);
    }
    window.postMessage({ __spcsv: true, type: "SPCSVExportDone", detail: { success: success, message: message, stopReason: "" } }, "*");
  }

  function progressMsg(suffix) {
    return REPORT_LABEL + ": " + suffix;
  }

  function normalizeGuid(g) {
    if (g == null) return "";
    return String(g).replace(/[{}]/g, "").trim().toLowerCase();
  }

  function fieldUserDisplayName(fieldValue) {
    if (fieldValue == null) return "";
    if (typeof fieldValue === "string") {
      var m = fieldValue.match(/;#(.+)$/);
      return m ? m[1] : fieldValue;
    }
    if (typeof fieldValue === "object") {
      var lv = fieldValue.LookupValue || fieldValue.lookupValue;
      if (lv) return String(lv);
      var t = fieldValue.Title || fieldValue.title;
      if (t) return String(t);
    }
    return "";
  }

  function escapeCSV(val) {
    var s = val == null ? "" : String(val);
    if (s.indexOf('"') >= 0 || s.indexOf(",") >= 0 || s.indexOf("\n") >= 0 || s.indexOf("\r") >= 0) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function escXml(s) {
    if (s == null) return "";
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function colToLetter(c) {
    var s = "";
    for (;;) {
      s = String.fromCharCode((c % 26) + 65) + s;
      c = Math.floor(c / 26) - 1;
      if (c < 0) break;
    }
    return s;
  }

  function sanitizeXmlText(s) {
    if (s == null) return "";
    var out = "";
    var str = String(s);
    for (var i = 0; i < str.length; i++) {
      var code = str.charCodeAt(i);
      if (code === 9 || code === 10 || code === 13) { out += str.charAt(i); continue; }
      if (code < 32 || code === 65535 || (code >= 0xD800 && code <= 0xDFFF) || code === 0xFFFE || code === 0xFFFF) continue;
      out += str.charAt(i);
    }
    return out;
  }

  function coerceCellValue(val) {
    if (val == null) return "";
    if (typeof val === "number" && !isNaN(val)) return val;
    var s = sanitizeXmlText(val);
    if (s.trim() === "") return "";
    if (/^-?\d+$/.test(s.trim())) return parseInt(s.trim(), 10);
    if (/^-?\d+\.\d*$/.test(s.trim()) || /^-?\d*\.\d+$/.test(s.trim())) {
      var n = parseFloat(s.trim());
      if (!isNaN(n)) return n;
    }
    return s;
  }

  function isEmptyCellValue(v) {
    return v === "" || v == null || (typeof v === "number" && isNaN(v));
  }

  function mergeRef(m) {
    return colToLetter(m.c1) + (m.r1 + 1) + ":" + colToLetter(m.c2) + (m.r2 + 1);
  }

  function xmlNumber(v) {
    if (typeof v !== "number" || isNaN(v) || !isFinite(v)) return "0";
    if (Number.isInteger(v)) return String(v);
    var s = String(v);
    if (/e/i.test(s)) return String(Number(v.toPrecision(15)));
    return s;
  }

  function worksheetOpenXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
  }

  function sheetPrXml(tabColor) {
    return tabColor ? '<sheetPr><tabColor rgb="' + tabColor + '"/></sheetPr>' : "";
  }

  function sheetViewsFreezeXml() {
    return '<sheetViews><sheetView workbookViewId="0" showGridLines="0" zoomScale="100">' +
      '<pane ySplit="1" topLeftCell="A2" state="frozen" activePane="bottomLeft"/>' +
      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' +
      "</sheetView></sheetViews>";
  }

  function sheetViewsDashboardXml() {
    return '<sheetViews><sheetView workbookViewId="0" showGridLines="0" zoomScale="100" showRowColHeaders="1">' +
      '<pane xSplit="2" ySplit="10" topLeftCell="C11" state="frozen" activePane="bottomRight"/>' +
      '<selection pane="topRight" activeCell="C1" sqref="C1"/>' +
      '<selection pane="bottomLeft" activeCell="A11" sqref="A11"/>' +
      '<selection pane="bottomRight" activeCell="C11" sqref="C11"/>' +
      "</sheetView></sheetViews>";
  }

  var STYLE_ROW_LIMIT = 50000;
  var COL_STYLE_DEFAULT = 20;

  var DATA_COL_WIDTHS = {
    matrix: [21.95, 36, 72, 13.3, 15.3, 35.6, 38, 18.8, 52, 17.3, 17.3, 16.3, 12, 12, 14.3, 12, 13.3, 20, 12, 20, 19.3, 20, 18.3, 19.3, 20, 20, 20],
    groups: [13.3, 53.45, 28, 30.35, 30.35, 12.3, 36, 52, 44, 16.7, 17.3],
    sharing: [18.8, 80, 17.3, 18.3, 13.3, 72, 17.3, 20, 18.3, 20.3, 39.8, 72, 29.3, 60, 14.3, 40, 68, 10.3, 60],
    allItems: [15.65, 38, 80, 13.3, 11.3, 22, 21.95, 27.2, 21.95, 32],
    folderCounts: [18, 28, 40, 22, 12, 14, 14, 16, 12, 16, 56],
    pathLengths: [18, 28, 56, 56, 16, 16]
  };

  function buildWorkbookStylesXml() {
    if (typeof GOLD_WORKBOOK_STYLES_XML === "string" && GOLD_WORKBOOK_STYLES_XML) {
      return GOLD_WORKBOOK_STYLES_XML;
    }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/></styleSheet>';
  }

  function colWidthsForSheet(tableKind, colCount) {
    var base = DATA_COL_WIDTHS[tableKind] || [];
    var out = [];
    for (var i = 0; i < colCount; i++) out.push(i < base.length ? base[i] : 14);
    return out;
  }

  function buildColsXml(colWidths, opts) {
    opts = opts || {};
    if (!colWidths || !colWidths.length) return "";
    var styleId = opts.styleId != null ? opts.styleId : COL_STYLE_DEFAULT;
    var hiddenFrom = opts.hiddenFrom;
    var parts = [];
    for (var ci = 0; ci < colWidths.length; ci++) {
      if (colWidths[ci] <= 0) continue;
      var hiddenAttr = hiddenFrom != null && ci >= hiddenFrom ? ' hidden="1"' : "";
      parts.push('<col min="' + (ci + 1) + '" max="' + (ci + 1) + '" width="' + colWidths[ci] + '" customWidth="1" style="' + styleId + '"' + hiddenAttr + "/>");
    }
    return parts.length ? "<cols>" + parts.join("") + "</cols>" : "";
  }

  function cellStyleForTable(r, c, val, tableKind, roleStartCol) {
    if (r === 0) return tableKind === "allItems" ? 40 : 28;
    if (r > STYLE_ROW_LIMIT) return 0;
    var sVal = val == null ? "" : String(val);
    var alt = (r % 2 === 0) ? 29 : 30;
    if (tableKind === "matrix") {
      if (c === 4) {
        if (sVal === "Custom") return 31;
        if (sVal === "Inherited") return 32;
        if (sVal === "Top level") return 33;
      }
      if (c === 9 && sVal === "Yes") return 34;
      if (roleStartCol >= 0 && c >= roleStartCol && sVal === "X") return 38;
      return alt;
    }
    if (tableKind === "sharing") {
      if (c === 0 && sVal === "SharingLink") return 37;
      if (c === 0 && sVal === "DirectPermission") return 38;
      if (c === 6 && sVal === "Yes") return 34;
      return alt;
    }
    if (tableKind === "groups") {
      if (c === 10 && sVal === "Yes") return 34;
      return alt;
    }
    if (tableKind === "allItems") {
      if (c === 5 && sVal === "Yes") return 34;
      return 41;
    }
    if (tableKind === "folderCounts" || tableKind === "pathLengths") return alt;
    return alt;
  }

  function stripBagelPathPrefix(s) {
    if (!s || typeof s !== "string") return s || "";
    var idx = s.indexOf("#");
    return idx >= 0 ? s.slice(idx + 1).trim() : s;
  }

  function bagelItemPathAndType(itm, knownFolderPaths) {
    var path = stripBagelPathPrefix(itm.FileRef || "");
    if (!path) {
      var dirRef0 = stripBagelPathPrefix(itm.FileDirRef || "");
      var leaf0 = stripBagelPathPrefix(itm.FileLeafRef || "");
      path = (dirRef0 + "/" + leaf0).replace(/\/+/g, "/").replace(/^\/+/, "/");
    }
    if (!path || path === "/") return null;
    path = path.replace(/\/+/g, "/");
    if (path.length > 1 && path.charAt(path.length - 1) === "/") path = path.slice(0, -1);
    if (!path) path = "/";
    var fsVal = itm.FSObjType != null ? itm.FSObjType : itm.FileSystemObjectType;
    var fsoNum = parseInt(fsVal, 10);
    var fsFolder = fsVal === 1 || fsVal === "1" || String(fsVal).toLowerCase() === "folder" || fsoNum === 1;
    var isFolder = !!fsFolder || (knownFolderPaths && knownFolderPaths[path]);
    var parentPath;
    if (isFolder) {
      parentPath = path.replace(/\/[^/]+$/, "") || "/";
    } else {
      var dirRef = stripBagelPathPrefix(itm.FileDirRef || "");
      dirRef = dirRef.replace(/\/+/g, "/");
      if (dirRef.length > 1 && dirRef.charAt(dirRef.length - 1) === "/") dirRef = dirRef.slice(0, -1);
      parentPath = dirRef || path.replace(/\/[^/]+$/, "") || "/";
    }
    return { path: path, isFolder: !!isFolder, parentPath: parentPath || "/" };
  }

  function appendBagelFolderCountRows(outRows, items, siteName, listTitle) {
    if (!items || !items.length) return;
    var knownFolderPaths = {};
    for (var i = 0; i < items.length; i++) {
      var dirRef = stripBagelPathPrefix(items[i].FileDirRef || "");
      if (dirRef) {
        dirRef = dirRef.replace(/\/+/g, "/");
        if (dirRef.length > 1 && dirRef.charAt(dirRef.length - 1) === "/") dirRef = dirRef.slice(0, -1);
        if (dirRef) knownFolderPaths[dirRef] = true;
      }
    }
    var parsed = [];
    for (var j = 0; j < items.length; j++) {
      var x = bagelItemPathAndType(items[j], knownFolderPaths);
      if (x) parsed.push(x);
    }
    if (!parsed.length) return;
    var pathToFolder = {};
    var root = null;
    for (var k = 0; k < parsed.length; k++) {
      var it0 = parsed[k];
      var candidate = it0.isFolder ? it0.path : it0.path.replace(/\/[^/]+$/, "").replace(/\/$/, "") || "/";
      if (!root || candidate.length < root.length) root = candidate;
    }
    root = (root || "/").replace(/\/$/, "") || "/";
    for (var m = 0; m < parsed.length; m++) {
      var it = parsed[m];
      var parent = it.parentPath || "/";
      if (parent && !pathToFolder[parent]) pathToFolder[parent] = { directFolders: 0, directFiles: 0, children: [] };
      if (it.isFolder) {
        if (!pathToFolder[it.path]) pathToFolder[it.path] = { directFolders: 0, directFiles: 0, children: [] };
        if (pathToFolder[parent]) {
          pathToFolder[parent].directFolders++;
          if (pathToFolder[parent].children.indexOf(it.path) < 0) pathToFolder[parent].children.push(it.path);
        }
      } else if (pathToFolder[parent]) {
        pathToFolder[parent].directFiles++;
      }
    }
    var allFolders = Object.keys(pathToFolder).filter(function (p) { return pathToFolder[p]; });
    allFolders.sort(function (a, b) { return b.split("/").length - a.split("/").length; });
    var totalRecursive = {};
    var totalNestedFolders = {};
    var levelsDeep = {};
    for (var t = 0; t < allFolders.length; t++) {
      var fp = allFolders[t];
      var info = pathToFolder[fp];
      var sum = info.directFiles;
      var nestedCount = 0;
      var maxChildDepth = -1;
      for (var c = 0; c < info.children.length; c++) {
        var ch = info.children[c];
        sum += 1 + (totalRecursive[ch] != null ? totalRecursive[ch] : 0);
        nestedCount += 1 + (totalNestedFolders[ch] != null ? totalNestedFolders[ch] : 0);
        var cd = levelsDeep[ch] != null ? levelsDeep[ch] : 0;
        if (cd > maxChildDepth) maxChildDepth = cd;
      }
      totalRecursive[fp] = sum;
      totalNestedFolders[fp] = nestedCount;
      levelsDeep[fp] = info.children.length === 0 ? 0 : 1 + maxChildDepth;
    }
    var folderPaths = Object.keys(pathToFolder).sort();
    var displayRoot = root;
    if (folderPaths.length > 0) {
      var shortest = folderPaths[0];
      for (var si = 1; si < folderPaths.length; si++) {
        if (folderPaths[si].length < shortest.length) shortest = folderPaths[si];
      }
      displayRoot = shortest;
    }
    var normRoot = (displayRoot || "/").replace(/\/+$/, "") || "/";
    function folderLevelFromPath(p) {
      if (!p || p === "/") return 0;
      var norm = p.replace(/\/+$/, "").replace(/^\/+/, "/");
      if (norm === normRoot || norm.length <= normRoot.length) return 0;
      if (normRoot !== "/" && norm.indexOf(normRoot + "/") !== 0) return 0;
      var rel = normRoot === "/" ? norm.replace(/^\/+/, "") : norm.slice(normRoot.length).replace(/^\/+/, "");
      return rel ? rel.split("/").filter(function (s) { return s.length > 0; }).length : 0;
    }
    for (var f = 0; f < folderPaths.length; f++) {
      var path = folderPaths[f];
      var finfo = pathToFolder[path];
      var folderPathDisplay = (path === displayRoot || path === displayRoot + "/") ? "/" : (path.indexOf(displayRoot) === 0 ? "/" + path.slice(displayRoot.length).replace(/^\//, "") : path);
      var folderName = path.replace(/.*\//, "") || path || "/";
      outRows.push([
        siteName,
        listTitle,
        folderPathDisplay,
        folderName,
        folderLevelFromPath(path),
        finfo.directFolders,
        finfo.directFiles,
        totalNestedFolders[path] != null ? totalNestedFolders[path] : 0,
        levelsDeep[path] != null ? levelsDeep[path] : 0,
        totalRecursive[path] != null ? totalRecursive[path] : 0,
        path
      ]);
    }
  }

  function appendBagelPathLengthRows(outRows, items, siteName, listTitle) {
    if (!items || !items.length) return;
    var knownFolderPaths = {};
    for (var i = 0; i < items.length; i++) {
      var dirRef = stripBagelPathPrefix(items[i].FileDirRef || "");
      if (dirRef) {
        dirRef = dirRef.replace(/\/+$/, "").replace(/^\/+/, "/");
        if (dirRef.length > 1 && dirRef.charAt(dirRef.length - 1) === "/") dirRef = dirRef.slice(0, -1);
        if (dirRef) knownFolderPaths[dirRef] = true;
      }
    }
    var shortestPath = null;
    for (var j = 0; j < items.length; j++) {
      var x = bagelItemPathAndType(items[j], knownFolderPaths);
      if (!x || x.isFolder) continue;
      if (!shortestPath || x.path.length < shortestPath.length) shortestPath = x.path;
    }
    var normRoot = (shortestPath ? shortestPath.replace(/\/[^/]+$/, "").replace(/\/+$/, "") : "") || "/";
    var fileEntries = [];
    for (var k = 0; k < items.length; k++) {
      var y = bagelItemPathAndType(items[k], knownFolderPaths);
      if (!y || y.isFolder) continue;
      var path = y.path;
      var pathAfterLibrary = (normRoot === "/" || path.indexOf(normRoot + "/") !== 0) ? path : path.slice(normRoot.length).replace(/^\/+/, "") || "";
      if (normRoot !== "/" && path === normRoot) pathAfterLibrary = "";
      var encodedPath = encodeURIComponent(pathAfterLibrary);
      fileEntries.push({
        pathAfterLibrary: pathAfterLibrary,
        encodedPath: encodedPath,
        friendlyLen: pathAfterLibrary.length,
        encodedLen: encodedPath.length
      });
    }
    fileEntries.sort(function (a, b) {
      if (b.encodedLen !== a.encodedLen) return b.encodedLen - a.encodedLen;
      return b.friendlyLen - a.friendlyLen;
    });
    for (var n = 0; n < fileEntries.length; n++) {
      var e = fileEntries[n];
      outRows.push([siteName, listTitle, e.pathAfterLibrary, e.encodedPath, e.friendlyLen, e.encodedLen]);
    }
  }

  function buildDashboardStylesXml() { return buildWorkbookStylesXml(); }

  function buildDashboardSheetXml(dashboard, getStrIndex, tabColor) {
    var cells = dashboard.cells || [];
    var merges = dashboard.merges || [];
    var rowHeights = dashboard.rowHeights || {};
    var colWidths = dashboard.colWidths || [];
    var rowsMap = {};
    var maxR = 0;
    var maxC = 0;
    for (var i = 0; i < cells.length; i++) {
      var cell = cells[i];
      if (cell.r > maxR) maxR = cell.r;
      if (cell.c > maxC) maxC = cell.c;
      if (!rowsMap[cell.r]) rowsMap[cell.r] = {};
      var ref = colToLetter(cell.c) + (cell.r + 1);
      var sAttr = cell.s != null ? ' s="' + cell.s + '"' : "";
      var cellXml;
      if (cell.inlineXml) {
        cellXml = '<c r="' + ref + '"' + sAttr + ' t="inlineStr">' + cell.inlineXml + "</c>";
      } else {
        var v = coerceCellValue(cell.v);
        var isNum = typeof v === "number" && !isNaN(v);
        cellXml = isNum
          ? '<c r="' + ref + '"' + sAttr + ' t="n"><v>' + xmlNumber(v) + "</v></c>"
          : '<c r="' + ref + '"' + sAttr + ' t="s"><v>' + getStrIndex(v) + "</v></c>";
      }
      rowsMap[cell.r][cell.c] = cellXml;
    }
    for (var mi = 0; mi < merges.length; mi++) {
      var merge = merges[mi];
      if (merge.r1 > maxR) maxR = merge.r1;
      if (merge.r2 > maxR) maxR = merge.r2;
      if (merge.c1 > maxC) maxC = merge.c1;
      if (merge.c2 > maxC) maxC = merge.c2;
      for (var mr = merge.r1; mr <= merge.r2; mr++) {
        if (!rowsMap[mr]) rowsMap[mr] = {};
      }
    }
    for (var fillR = 0; fillR <= maxR; fillR++) {
      if (!rowsMap[fillR]) rowsMap[fillR] = {};
    }
    var rowKeys = Object.keys(rowsMap).map(function (k) { return parseInt(k, 10); }).sort(function (a, b) { return a - b; });
    var sheetRows = [];
    for (var ri = 0; ri < rowKeys.length; ri++) {
      var r = rowKeys[ri];
      var ht = rowHeights[String(r)];
      var htAttr = ht != null ? ' ht="' + ht + '" customHeight="1"' : "";
      var colKeys = Object.keys(rowsMap[r]).map(function (k) { return parseInt(k, 10); }).sort(function (a, b) { return a - b; });
      var rowCells = [];
      for (var ck = 0; ck < colKeys.length; ck++) rowCells.push(rowsMap[r][colKeys[ck]]);
      sheetRows.push('<row r="' + (r + 1) + '"' + htAttr + ">" + rowCells.join("") + "</row>");
    }
    var colsXml = "";
    if (colWidths.length) {
      colsXml = buildColsXml(colWidths, { styleId: COL_STYLE_DEFAULT, hiddenFrom: dashboard.colHiddenFrom });
    }
    var mergeXml = merges.length
      ? '<mergeCells count="' + merges.length + '">' + merges.map(function (m) {
        return '<mergeCell ref="' + mergeRef(m) + '"/>';
      }).join("") + "</mergeCells>"
      : "";
    var dim = maxC >= 0 && maxR >= 0 ? "A1:" + colToLetter(maxC) + (maxR + 1) : "A1";
    return worksheetOpenXml() + sheetPrXml(tabColor) +
      '<dimension ref="' + dim + '"/>' + sheetViewsDashboardXml() +
      '<sheetFormatPr defaultRowHeight="15"/>' + colsXml +
      "<sheetData>" + sheetRows.join("") + "</sheetData>" + mergeXml + "</worksheet>";
  }

  function aoaToXlsxBlobMulti(sheets) {
    if (!window.JSZip) return Promise.reject(new Error("JSZip not loaded"));
    if (!sheets || !sheets.length) return Promise.reject(new Error("No sheets"));
    var sharedStrings = [];
    var sharedIndex = {};
    var totalStringRefs = 0;
    function getStrIndex(val) {
      var s = sanitizeXmlText(val == null ? "" : val);
      totalStringRefs++;
      if (sharedIndex[s] === undefined) {
        sharedIndex[s] = sharedStrings.length;
        sharedStrings.push(s);
      }
      return sharedIndex[s];
    }
    function buildSheetRows(aoa, opts) {
      opts = opts || {};
      var sparseBody = !!opts.sparseBody;
      var tableKind = opts.tableKind || "";
      var roleStartCol = opts.roleStartCol != null ? opts.roleStartCol : -1;
      var nr = aoa.length;
      var styleBody = !!tableKind && (nr - 1) <= STYLE_ROW_LIMIT;
      var nc = nr > 0 ? (aoa[0] || []).length : 0;
      var sheetRows = [];
      for (var r = 0; r < nr; r++) {
        var row = aoa[r];
        var cells = [];
        var styleRow = styleBody || r === 0;
        for (var c = 0; c < nc; c++) {
          var raw = row[c];
          var v = coerceCellValue(raw);
          if (sparseBody && r > 0 && isEmptyCellValue(v)) continue;
          var ref = colToLetter(c) + (r + 1);
          var styleIdx = styleRow ? cellStyleForTable(r, c, v, tableKind, roleStartCol) : 0;
          var sAttr = styleIdx ? ' s="' + styleIdx + '"' : "";
          var isNum = typeof v === "number" && !isNaN(v);
          if (isNum) cells.push('<c r="' + ref + '"' + sAttr + ' t="n"><v>' + xmlNumber(v) + "</v></c>");
          else cells.push('<c r="' + ref + '"' + sAttr + ' t="s"><v>' + getStrIndex(v) + "</v></c>");
        }
        if (cells.length) {
          var rowHt = "";
          if (r === 0) rowHt = ' ht="28" customHeight="1"';
          else if (styleBody && tableKind) rowHt = ' ht="17" customHeight="1"';
          sheetRows.push('<row r="' + (r + 1) + '"' + rowHt + ">" + cells.join("") + "</row>");
        }
      }
      return { rows: sheetRows, nr: nr, nc: nc };
    }

    function buildDataWorksheetXml(built, opts) {
      opts = opts || {};
      var filterRange = built.nc > 0 && built.nr > 0 ? "A1:" + colToLetter(built.nc - 1) + built.nr : "A1";
      // Always enable AutoFilter (All Items is often >>10k rows; the old cap left that sheet unfilterable).
      var filterXml = built.nr > 0 && built.nc > 0 ? ('<autoFilter ref="' + filterRange + '"/>') : "";
      var freezeXml = built.nr > 1 ? sheetViewsFreezeXml() : "";
      var formatXml = built.nr > 1 ? '<sheetFormatPr defaultRowHeight="15"/>' : "";
      var colsXml = opts.colWidths ? buildColsXml(opts.colWidths, { styleId: COL_STYLE_DEFAULT }) : "";
      return worksheetOpenXml() + sheetPrXml(opts.tabColor) +
        '<dimension ref="' + filterRange + '"/>' + freezeXml + formatXml + colsXml +
        "<sheetData>" + built.rows.join("") + "</sheetData>" + filterXml + "</worksheet>";
    }
    var sheetXmls = [];
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].dashboard) {
        sheetXmls.push(buildDashboardSheetXml(sheets[s].dashboard, getStrIndex, sheets[s].tabColor));
        continue;
      }
      var aoa = sheets[s].aoa || [];
      var sparse = aoa.length > 5000 && sheets[s].tableKind !== "matrix";
      var built = buildSheetRows(aoa, {
        sparseBody: sparse,
        tableKind: sheets[s].tableKind || "",
        roleStartCol: sheets[s].roleStartCol != null ? sheets[s].roleStartCol : -1
      });
      sheetXmls.push(buildDataWorksheetXml(built, {
        tabColor: sheets[s].tabColor,
        colWidths: colWidthsForSheet(sheets[s].tableKind || "", built.nc)
      }));
    }
    var sstItems = sharedStrings.map(function (str) {
      var t = escXml(sanitizeXmlText(str));
      if (/[\s\u00a0]/.test(str) || str.length !== str.trim().length) return '<si><t xml:space="preserve">' + t + "</t></si>";
      return "<si><t>" + t + "</t></si>";
    });
    var sstXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' +
      totalStringRefs + '" uniqueCount="' + sharedStrings.length + '">' + sstItems.join("") + "</sst>";
    var contentTypesParts = [
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    ];
    for (var i = 0; i < sheets.length; i++) {
      contentTypesParts.push('<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
    }
    contentTypesParts.push(
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    );
    var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + contentTypesParts.join("") + "</Types>";
    var workbookRelsParts = [];
    for (var j = 0; j < sheets.length; j++) {
      workbookRelsParts.push('<Relationship Id="rId' + (j + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (j + 1) + '.xml"/>');
    }
    workbookRelsParts.push(
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
      '<Relationship Id="rId' + (sheets.length + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    );
    var workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + workbookRelsParts.join("") + "</Relationships>";
    var sheetEls = [];
    for (var k = 0; k < sheets.length; k++) {
      var sName = escXml(String(sheets[k].name || "Sheet" + (k + 1)).slice(0, 31));
      var sId = k + 1;
      sheetEls.push('<sheet name="' + sName + '" sheetId="' + sId + '" r:id="rId' + sId + '"/>');
    }
    var workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>' + sheetEls.join("") + '</sheets><calcPr fullCalcOnLoad="1"/></workbook>';
    var stylesXml = buildWorkbookStylesXml();
    var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var zip = new window.JSZip();
    zip.file("[Content_Types].xml", contentTypes);
    zip.file("_rels/.rels", rels);
    zip.file("xl/workbook.xml", workbookXml);
    zip.file("xl/_rels/workbook.xml.rels", workbookRels);
    for (var z = 0; z < sheetXmls.length; z++) zip.file("xl/worksheets/sheet" + (z + 1) + ".xml", sheetXmls[z]);
    zip.file("xl/sharedStrings.xml", sstXml);
    zip.file("xl/styles.xml", stylesXml);
    return zip.generateAsync({
      type: "arraybuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 }
    }).then(function (ab) {
      return new Blob([ab], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    });
  }

  function downloadBlobFile(blob, filename) {
    filename = filename || "export.xlsx";
    var link = document.createElement("a");
    var url = URL.createObjectURL(blob);
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    setTimeout(function () {
      URL.revokeObjectURL(url);
      if (link.parentNode) link.parentNode.removeChild(link);
    }, 60000);
    return Promise.resolve();
  }

  function downloadWorkbook(sheets, filenameBase) {
    var fn = (filenameBase || "PermissionsMatrix").replace(/\.(csv|xls|xlsx|xml)$/i, "") + ".xlsx";
    if (params.format === "csv") {
      var matrix = sheets.filter(function (s) { return s.name === "Permissions Matrix"; })[0] || sheets[0];
      var aoa = matrix && matrix.aoa ? matrix.aoa : [[]];
      var lines = aoa.map(function (row) { return row.map(escapeCSV).join(","); });
      var blob = new Blob(["\ufeff" + lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
      return downloadBlobFile(blob, fn.replace(/\.xlsx$/i, ".csv"));
    }
    return aoaToXlsxBlobMulti(sheets).then(function (blob) {
      return downloadBlobFile(blob, fn);
    });
  }

  var spi = typeof window._spPageContextInfo === "object" ? window._spPageContextInfo : null;
  var apiOrigin = (typeof location !== "undefined" && location.origin) ? String(location.origin).replace(/\/$/, "") : "";
  if (!apiOrigin && params.u) {
    try { apiOrigin = new URL(String(params.u)).origin.replace(/\/$/, ""); } catch (_) {}
  }

  function serverRelativeFromAny(urlOrPath) {
    if (!urlOrPath) return "";
    var s = String(urlOrPath).trim();
    if (s.charAt(0) === "/") return s.replace(/\/$/, "") || "/";
    try { return new URL(s).pathname.replace(/\/$/, "") || "/"; } catch (_) { return s.replace(/\/$/, ""); }
  }

  function webUrlFromPath(serverRelativePath) {
    var path = serverRelativeFromAny(serverRelativePath);
    if (!path || path === "/") return apiOrigin;
    return apiOrigin + path;
  }

  function normalizeApiUrl(url) {
    if (!url || !apiOrigin) return url;
    try {
      var u = new URL(String(url), apiOrigin);
      if (u.origin !== apiOrigin) return apiOrigin + u.pathname + u.search;
      return u.href;
    } catch (_) {
      return url;
    }
  }

  var siteUrl = apiOrigin;
  if (spi && spi.webServerRelativeUrl) {
    siteUrl = webUrlFromPath(spi.webServerRelativeUrl);
  } else if (params.u) {
    siteUrl = normalizeApiUrl(String(params.u).replace(/\/$/, ""));
  } else if (spi && spi.webAbsoluteUrl) {
    siteUrl = normalizeApiUrl(String(spi.webAbsoluteUrl).replace(/\/$/, ""));
  }

  var INCLUDE_SUBSITES = params.matrixIncludeSubsites !== false;
  var INCLUDE_ALL_INHERITED = params.matrixIncludeAllInherited === true;
  var EXPAND_GROUPS = params.matrixExpandGroups === true;
  var MAX_LIST_ITEMS = params.matrixMaxListItems > 0 ? parseInt(params.matrixMaxListItems, 10) : 2000;
  var PAGE_SIZE = params.matrixListItemPageSize > 0 ? Math.min(5000, Math.max(100, parseInt(params.matrixListItemPageSize, 10))) : 5000;
  var HAS_EXPLICIT_SITE_SELECTION = Array.isArray(params.matrixSelectedPaths);
  var SELECTED_SITE_PATHS = HAS_EXPLICIT_SITE_SELECTION
    ? params.matrixSelectedPaths.map(function (p) { return serverRelativeFromAny(p).toLowerCase(); }).filter(Boolean)
    : null;
  var SELECTED_LIST_KEYS = null;
  if (Array.isArray(params.reportSelectedLists) && params.reportSelectedLists.length) {
    SELECTED_LIST_KEYS = {};
    for (var sli = 0; sli < params.reportSelectedLists.length; sli++) {
      var sle = params.reportSelectedLists[sli] || {};
      var slGuid = normalizeGuid(sle.listId || sle.ListId || sle.id);
      if (!slGuid) continue;
      var slSite = String(sle.siteUrl || "").replace(/\/$/, "").toLowerCase();
      SELECTED_LIST_KEYS[slGuid] = true;
      if (slSite) SELECTED_LIST_KEYS[slSite + "|" + slGuid] = true;
    }
  }
  function listIsSelectedForMatrix(webUrl, listId) {
    if (!SELECTED_LIST_KEYS) return true;
    var guid = normalizeGuid(listId);
    if (!guid) return false;
    if (SELECTED_LIST_KEYS[guid]) return true;
    var siteKey = String(webUrl || "").replace(/\/$/, "").toLowerCase();
    return !!(siteKey && SELECTED_LIST_KEYS[siteKey + "|" + guid]);
  }
  var INCLUDE_FOLDER_SHARING_LINKS = params.matrixIncludeFolderSharingLinks !== false;
  var SHARING_LINK_FETCH_ALL = params.matrixSharingLinkFetchAll !== false;
  var REQUEST_PACE_MS = 60;
  var SHARING_PACE_MS = 750;
  var groupUsersCache = {};
  var sharingQueue = [];
  var sharingLinkUrlSet = {};
  var requestDigestCache = {};
  var sharingApiSkipCache = {};
  var SHARING_API_BODY = JSON.stringify({ request: { populateInheritedLinks: true } });
  var SHARING_HEADERS_VERBOSE = {
    Accept: "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose"
  };

  var NON_INHERITABLE_ROLES = {
    "Limited Access": 1, "Restricted View": 1, "Restricted Read": 1, "Web-Only Limited Access": 1
  };

  var EXCLUDED_LISTS = {
    "Access Requests": 1, "App Packages": 1, appdata: 1, appfiles: 1, "Apps in Testing": 1,
    "Cache Profiles": 1, "Composed Looks": 1, "Content and Structure Reports": 1,
    "Content type publishing error log": 1, "Converted Forms": 1, "Device Channels": 1,
    "Form Templates": 1, fpdatasources: 1, "Get started with Apps for Office and SharePoint": 1,
    "List Template Gallery": 1, "Long Running Operation Status": 1, "Maintenance Log Library": 1,
    "Master Docs": 1, "Master Page Gallery": 1, MicroFeed: 1, NintexFormXml: 1,
    "Quick Deploy Items": 1, "Relationships List": 1, "Reusable Content": 1,
    "Reporting Metadata": 1, "Reporting Templates": 1, "Search Config List": 1,
    "Preservation Hold Library": 1, "Solution Gallery": 1,
    "Suggested Content Browser Locations": 1, "Theme Gallery": 1, TaxonomyHiddenList: 1,
    "User Information List": 1, "Web Part Gallery": 1, wfpub: 1, wfsvc: 1,
    "Workflow History": 1, "Workflow Tasks": 1, "Style Library": 1
  };

  var PREFERRED_ROLES = [
    "Full Control", "Design", "Edit", "Contribute", "Read", "View Only",
    "Create new subsites", "Approve", "Manage Hierarchy", "Restricted Read",
    "Restricted Interfaces for Translation"
  ];

  var SHARING_LINK_COLUMNS = [
    "RecordType", "ShareLinkUrl", "ShareLinkType", "ShareLinkScope", "LinkRoles", "LinkUsers",
    "External user", "Expiration", "BlocksDownload", "RequiresPassword", "ShareId", "SiteUrl", "ListTitle",
    "ListUrl", "ObjectType", "ItemName", "RelativeUrl", "ItemId", "Access"
  ];

  function principalTypeLabel(pt) {
    var n = parseInt(pt, 10);
    if (n === 1) return "User";
    if (n === 2) return "Distribution list";
    if (n === 4) return "Security group";
    if (n === 8) return "SharePoint group";
    return "User";
  }

  function sharingPrincipalTypeName(pt) {
    var n = parseInt(pt, 10);
    if (n === 1) return "User";
    if (n === 8) return "SharePointGroup";
    if (n === 4) return "SecurityGroup";
    if (n === 2) return "DistributionList";
    return "Type" + n;
  }

  function isExternalUser(loginName, email) {
    if (email && /#EXT#@/i.test(String(email))) return true;
    if (!loginName) return false;
    return /#ext#/i.test(String(loginName));
  }

  function testSharingLink(login, title) {
    var l = String(login || "");
    var t = String(title || "");
    return l.indexOf("SharingLinks.") === 0 || t.indexOf("SharingLinks.") === 0;
  }

  function getSharingLinkLabel(login, title) {
    var src = String(login || "").indexOf("SharingLinks.") === 0 ? login : (String(title || "").indexOf("SharingLinks.") === 0 ? title : "");
    if (!src) return "Sharing link";
    var m = String(src).match(/SharingLinks\.[^.]+\.([^.]+)\./i);
    if (!m) return "Sharing link";
    var k = m[1];
    var m2;
    if ((m2 = k.match(/^Organization(.+)$/i))) return "Organization (" + m2[1].toLowerCase() + ") sharing link";
    if ((m2 = k.match(/^Anonymous(.+)$/i))) return "Anonymous (" + m2[1].toLowerCase() + ") sharing link";
    if ((m2 = k.match(/^Users(.+)$/i))) return "Specific people (" + m2[1].toLowerCase() + ") sharing link";
    return k + " sharing link";
  }

  function testListScopedPrincipal(login, title) {
    var b = (String(title || "") + "|" + String(login || "")).toLowerCase();
    return /limited access system group|sharinglinks\.|sharing link/.test(b);
  }

  function roleNamesFromSet(roleSet) {
    var out = [];
    for (var k in roleSet) if (roleSet[k]) out.push(k);
    return out;
  }

  function formatSharingAccessEntry(principalType, title, login, email, roles) {
    var typeLabel = principalType === "User" ? "User" : (principalType === "SharePointGroup" ? "Group" : principalType);
    var roleText = roles.length ? roles.join(", ") : "(no roles)";
    if (principalType === "User") {
      var identity = email ? (title && title !== email ? title + " <" + email + ">" : email) : (title || login);
      return "User: " + identity + " [" + roleText + "]";
    }
    return typeLabel + ": " + (title || login) + " [" + roleText + "]";
  }

  function shouldEmitMatrixRow(row, roleNames) {
    var inh = row[4];
    if (inh !== "Inherited") return true;
    for (var i = 0; i < roleNames.length; i++) {
      if (row[11 + i] === "X" && !NON_INHERITABLE_ROLES[roleNames[i]]) return true;
    }
    return false;
  }

  function normalizePath(p) {
    if (!p) return "";
    return String(p).replace(/\\/g, "/").replace(/\/+/g, "/").replace(/\/$/, "").replace(/^\//, "").trim();
  }

  function matrixItemName(itemPath, itemType, siteName) {
    if (itemType === "Site" || itemType === "Subsite") return siteName || itemPath || "";
    if (!itemPath) return "";
    var path = String(itemPath).replace(/\/$/, "");
    var leaf = path;
    var idx = path.lastIndexOf("/");
    if (idx >= 0) leaf = path.slice(idx + 1);
    try { return decodeURIComponent(leaf.replace(/\+/g, " ")); } catch (_) { return leaf; }
  }

  function webItemType(serverRelativeUrl, rootPath) {
    var p = normalizePath(serverRelativeUrl);
    var root = normalizePath(rootPath);
    if (!p || !root || p === root) return "Site";
    return "Subsite";
  }

  async function fetchWithRetry(url, opts, tries) {
    tries = tries || 8;
    var normalized = normalizeApiUrl(url);
    checkCancelled();
    if (/throttle/i.test(normalized)) {
      await sleep(5000);
      throw new Error("SharePoint is throttling requests. Wait a minute and try again.");
    }
    for (var attempt = 1; attempt <= tries; attempt++) {
      checkCancelled();
      var r = await fetch(normalized, opts);
      var finalUrl = (r.url || normalized || "").toString();
      if (/throttle/i.test(finalUrl)) {
        await sleep(Math.min(5000 * attempt, 30000));
        continue;
      }
      if (r.ok) {
        if (REQUEST_PACE_MS > 0) await sleep(REQUEST_PACE_MS);
        return r;
      }
      if ([406, 429, 503, 504].indexOf(r.status) >= 0) {
        var ra = r.headers.get("Retry-After");
        await sleep(ra ? parseInt(ra, 10) * 1000 : Math.min(2000 * Math.pow(2, attempt - 1), 45000));
        continue;
      }
      if (r.status === 404) return r;
      throw new Error("HTTP " + r.status);
    }
    throw new Error("SharePoint throttled too many requests. Wait a minute and try again.");
  }

  async function fetchJson(url, accept) {
    var r = await fetchWithRetry(normalizeApiUrl(url), { credentials: "include", headers: { Accept: accept } }, 8);
    if (!r.ok) throw new Error("HTTP " + r.status);
    return r.json();
  }

  async function fetchJsonOptional(url, accept) {
    try {
      var r = await fetchWithRetry(normalizeApiUrl(url), { credentials: "include", headers: { Accept: accept } }, 4);
      if (!r.ok) return null;
      return r.json();
    } catch (_) {
      return null;
    }
  }

  async function fetchAllWebs(startUrl, includeSubsites) {
    var accept = "application/json;odata=nometadata";
    var webs = [];
    var seen = {};

    async function addWebByPath(serverRelativePath, walkSubs) {
      var path = serverRelativeFromAny(serverRelativePath);
      if (!path) return;
      var pathKey = path.toLowerCase();
      if (seen[pathKey]) return;
      seen[pathKey] = true;

      var webApiUrl = webUrlFromPath(path);
      var j = await fetchJson(webApiUrl + "/_api/web?$select=Title,ServerRelativeUrl,Url,HasUniqueRoleAssignments", accept);
      var srUrl = serverRelativeFromAny(j.ServerRelativeUrl || j.serverRelativeUrl || path);
      webs.push({
        url: webUrlFromPath(srUrl),
        title: (j.Title || j.title || "Site").trim(),
        serverRelativeUrl: srUrl,
        hasUnique: j.HasUniqueRoleAssignments === true,
        absoluteUrl: (j.Url || j.url || webApiUrl).trim()
      });
      if (walkSubs === false || !includeSubsites) return;
      var subs = await fetchJson(webUrlFromPath(srUrl) + "/_api/web/webs?$select=Title,ServerRelativeUrl,Url&$top=500", accept);
      var items = subs.value || subs.results || [];
      for (var i = 0; i < items.length; i++) {
        var subPath = items[i].ServerRelativeUrl || items[i].serverRelativeUrl;
        if (!subPath && (items[i].Url || items[i].url)) {
          subPath = serverRelativeFromAny(items[i].Url || items[i].url);
        }
        if (subPath) await addWebByPath(subPath);
      }
    }

    if (HAS_EXPLICIT_SITE_SELECTION) {
      if (!SELECTED_SITE_PATHS || !SELECTED_SITE_PATHS.length) {
        throw new Error("No lists selected for the permissions matrix. Load lists and select at least one list or library.");
      }
      for (var spi = 0; spi < SELECTED_SITE_PATHS.length; spi++) {
        await addWebByPath(SELECTED_SITE_PATHS[spi], false);
      }
      return webs;
    }

    var startPath = serverRelativeFromAny(startUrl);
    if (startPath && startPath !== "/" && startPath.indexOf("/") === 0) {
      await addWebByPath(startPath);
    } else {
      var root = await fetchJson(apiOrigin + "/_api/web?$select=ServerRelativeUrl", accept);
      await addWebByPath(root.ServerRelativeUrl || root.serverRelativeUrl || "/");
    }
    return webs;
  }

  async function fetchRoleNames(webUrl) {
    var accept = "application/json;odata=nometadata";
    var roleNames = [];
    try {
      var j = await fetchJson(webUrl + "/_api/web/roledefinitions?$select=Name,Order&$filter=Hidden eq false&$orderby=Order asc", accept);
      var defs = j.value || j.results || [];
      for (var i = 0; i < defs.length; i++) {
        var name = (defs[i].Name || "").trim();
        if (name && roleNames.indexOf(name) < 0) roleNames.push(name);
      }
    } catch (_) {}
    if (roleNames.length === 0) roleNames = PREFERRED_ROLES.slice();
    else {
      var ordered = [];
      PREFERRED_ROLES.forEach(function (r) { if (roleNames.indexOf(r) >= 0) ordered.push(r); });
      roleNames.forEach(function (r) { if (ordered.indexOf(r) < 0) ordered.push(r); });
      roleNames = ordered;
    }
    return roleNames;
  }

  async function fetchRoleAssignments(raUrl, accept) {
    var all = [];
    var base = normalizeApiUrl(raUrl);
    var next = base + (base.indexOf("?") >= 0 ? "&" : "?") + "$expand=Member,RoleDefinitionBindings&$top=5000";
    while (next) {
      var j = await fetchJson(next, accept);
      all = all.concat(j.value || j.results || []);
      next = j["@odata.nextLink"] ? normalizeApiUrl(j["@odata.nextLink"]) : null;
    }
    return all;
  }

  async function fetchGroupUsers(webUrl, groupId, accept) {
    if (!groupId) return [];
    var cacheKey = webUrl + "|" + groupId;
    if (Object.prototype.hasOwnProperty.call(groupUsersCache, cacheKey)) {
      return groupUsersCache[cacheKey];
    }
    var j = await fetchJsonOptional(webUrl + "/_api/web/sitegroups/GetById(" + groupId + ")/users?$select=Id,Title,LoginName,Email,PrincipalType", accept);
    var users = (j && (j.value || j.results)) || [];
    groupUsersCache[cacheKey] = users;
    return users;
  }

  function recordGroupMember(groupCatalog, siteName, siteUrlAbs, sitePath, groupTitle, groupLogin, groupId, user) {
    var key = sitePath + "|" + groupId + "|" + (user.loginName || user.title);
    if (groupCatalog[key]) return;
    groupCatalog[key] = true;
    var ext = isExternalUser(user.loginName, user.email) ? "Yes" : "";
    if (ext === "Yes" && user.loginName) {
      externalUserLogins[user.loginName.toLowerCase()] = true;
      if (!externalUsersBySite[siteName]) externalUsersBySite[siteName] = {};
      externalUsersBySite[siteName][user.loginName.toLowerCase()] = true;
    }
    groupMemberRows.push([
      siteName, siteUrlAbs, sitePath, groupTitle, groupLogin, groupId,
      user.title || "", user.loginName || "", user.email || "",
      user.principalType || "User", ext
    ]);
  }

  async function registerGroupMembers(webUrl, groupId, groupTitle, groupLogin, accept, groupCatalog, siteCtx) {
    var gkey = (siteCtx.sitePath || "") + "|" + groupId;
    if (registeredGroups[gkey]) return;
    registeredGroups[gkey] = true;
    var users = await fetchGroupUsers(webUrl, groupId, accept);
    if (!users.length) {
      recordGroupMember(groupCatalog, siteCtx.siteName, siteCtx.webUrl, siteCtx.sitePath, groupTitle, groupLogin, groupId, {
        title: "(no members or unable to read)",
        loginName: "",
        email: "",
        principalType: ""
      });
      return;
    }
    for (var i = 0; i < users.length; i++) {
      var u = users[i];
      recordGroupMember(groupCatalog, siteCtx.siteName, siteCtx.webUrl, siteCtx.sitePath, groupTitle, groupLogin, groupId, {
        title: (u.Title || u.LoginName || "").trim(),
        loginName: (u.LoginName || "").trim(),
        email: (u.Email || u.email || "").trim(),
        principalType: principalTypeLabel(u.PrincipalType)
      });
    }
  }

  async function principalsFromAssignments(webUrl, assignments, accept, groupCatalog, siteCtx) {
    var principals = [];
    for (var i = 0; i < assignments.length; i++) {
      var ra = assignments[i];
      var m = ra.Member || ra.member;
      if (!m) continue;
      var bindings = ra.RoleDefinitionBindings || ra.roleDefinitionBindings || [];
      var roleSet = {};
      for (var b = 0; b < bindings.length; b++) {
        var rn = (bindings[b].Name || bindings[b].name || "").trim();
        if (rn) roleSet[rn] = true;
      }
      var ptNum = parseInt(m.PrincipalType, 10);
      var title = (m.Title || m.LoginName || "").trim();
      var login = (m.LoginName || "").trim();
      var email = (m.Email || m.email || "").trim();
      var groupId = ptNum === 8 && m.Id != null ? String(m.Id) : "";

      if (testSharingLink(login, title)) {
        var label = getSharingLinkLabel(login, title);
        if (groupId && siteCtx) await registerGroupMembers(webUrl, groupId, label, login, accept, groupCatalog, siteCtx);
        if (EXPAND_GROUPS && groupId) {
          var linkMembers = await fetchGroupUsers(webUrl, groupId, accept);
          for (var lm = 0; lm < linkMembers.length; lm++) {
            var lu = linkMembers[lm];
            if (parseInt(lu.PrincipalType, 10) !== 1) continue;
            principals.push({
              title: (lu.Title || lu.LoginName || "").trim(),
              loginName: (lu.LoginName || "").trim(),
              principalType: "User",
              roleSet: roleSet,
              givenThrough: "Sharing Link",
              details: label
            });
          }
        } else {
          principals.push({
            title: label,
            loginName: label,
            principalType: "SharePoint group",
            roleSet: roleSet,
            givenThrough: "Sharing Link",
            details: label
          });
        }
        continue;
      }

      if (testListScopedPrincipal(login, title)) continue;

      principals.push({
        title: title, loginName: login, email: email,
        principalType: principalTypeLabel(m.PrincipalType),
        roleSet: roleSet, givenThrough: "Explicit", details: ""
      });

      if (ptNum === 8 && groupId && siteCtx) {
        await registerGroupMembers(webUrl, groupId, title, login, accept, groupCatalog, siteCtx);
        if (EXPAND_GROUPS) {
          var members = await fetchGroupUsers(webUrl, groupId, accept);
          for (var mi = 0; mi < members.length; mi++) {
            var mu = members[mi];
            if (parseInt(mu.PrincipalType, 10) !== 1) continue;
            var mul = (mu.LoginName || "").trim();
            var mut = (mu.Title || mul || "").trim();
            if (testListScopedPrincipal(mul, mut)) continue;
            principals.push({
              title: mut, loginName: mul, email: (mu.Email || mu.email || "").trim(),
              principalType: "User", roleSet: roleSet,
              givenThrough: title + " Members", details: "",
              inheritanceOverride: "Inherited"
            });
          }
        }
      }
    }
    return principals;
  }

  function buildAccessText(assignments) {
    var entries = [];
    for (var i = 0; i < assignments.length; i++) {
      var ra = assignments[i];
      var m = ra.Member || ra.member;
      if (!m) continue;
      var bindings = ra.RoleDefinitionBindings || ra.roleDefinitionBindings || [];
      var roles = [];
      for (var b = 0; b < bindings.length; b++) {
        var rn = (bindings[b].Name || bindings[b].name || "").trim();
        if (rn) roles.push(rn);
      }
      if (testSharingLink(m.LoginName || m.loginName, m.Title || m.title)) {
        roles = roles.filter(function (r) { return !NON_INHERITABLE_ROLES[r]; });
        if (!roles.length) continue;
      }
      entries.push(formatSharingAccessEntry(
        sharingPrincipalTypeName(m.PrincipalType),
        (m.Title || "").trim(),
        (m.LoginName || "").trim(),
        (m.Email || m.email || "").trim(),
        roles
      ));
    }
    return entries.join("; ");
  }

  function assignmentsHaveSharingLink(assignments) {
    for (var i = 0; i < assignments.length; i++) {
      var m = (assignments[i].Member || assignments[i].member);
      if (m && testSharingLink(m.LoginName || m.loginName, m.Title || m.title)) return true;
    }
    return false;
  }

  function queueSharingItem(entry) {
    if (entry.itemType === "Folder" && !INCLUDE_FOLDER_SHARING_LINKS) return;
    sharingQueue.push(entry);
  }

  async function getRequestDigest(webUrl) {
    var key = String(webUrl || "").toLowerCase();
    if (requestDigestCache[key]) return requestDigestCache[key];
    var r = await fetch(webUrl + "/_api/contextinfo", {
      method: "POST",
      credentials: "include",
      headers: { Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose" }
    });
    if (!r.ok) throw new Error("ContextInfo HTTP " + r.status);
    var j = await r.json();
    var digest = j.d && j.d.GetContextWebInformation && j.d.GetContextWebInformation.FormDigestValue;
    if (!digest) throw new Error("No FormDigestValue");
    requestDigestCache[key] = digest;
    return digest;
  }

  function parseSharingLinksFromResponse(j) {
    var info = j.d || j;
    var pi = info.permissionsInformation || info.PermissionsInformation || {};
    var links = pi.links || pi.Links || [];
    if (links && links.results) links = links.results;
    if (!Array.isArray(links)) links = [];
    return links;
  }

  function linkUsersFromSharingLink(link) {
    var emails = [];
    var ids = link.grantedToIdentitiesV2 || link.GrantedToIdentitiesV2 || link.grantedToIdentities || link.GrantedToIdentities || [];
    for (var i = 0; i < ids.length; i++) {
      var u = ids[i].user || ids[i].User;
      if (u && (u.email || u.Email)) emails.push(u.email || u.Email);
    }
    return emails.join("|");
  }

  function sharingLinkHasExternal(link) {
    var ids = link.grantedToIdentitiesV2 || link.GrantedToIdentitiesV2 || link.grantedToIdentities || link.GrantedToIdentities || [];
    for (var i = 0; i < ids.length; i++) {
      var u = ids[i].user || ids[i].User;
      if (u && isExternalUser(u.loginName || u.LoginName, u.email || u.Email)) return true;
    }
    return !!(link.hasExternalGuestInvitees || link.HasExternalGuestInvitees);
  }

  async function fetchSharingPostWithRetry(url, digest, tries) {
    tries = tries || 12;
    var normalized = normalizeApiUrl(url);
    for (var attempt = 1; attempt <= tries; attempt++) {
      var r = await fetch(normalized, {
        method: "POST",
        credentials: "include",
        headers: Object.assign({ "X-RequestDigest": digest }, SHARING_HEADERS_VERBOSE),
        body: SHARING_API_BODY
      });
      var finalUrl = (r.url || normalized || "").toString();
      if (/throttle/i.test(finalUrl)) {
        await sleep(Math.min(8000 * attempt, 45000));
        continue;
      }
      if ([429, 503, 504, 406].indexOf(r.status) >= 0) {
        var ra = r.headers.get("Retry-After");
        await sleep(ra ? parseInt(ra, 10) * 1000 : Math.min(4000 * Math.pow(2, attempt - 1), 60000));
        continue;
      }
      if (SHARING_PACE_MS > 0) await sleep(SHARING_PACE_MS);
      return r;
    }
    return null;
  }

  async function fetchSharingLinksForItem(webUrl, listId, itemId) {
    var cacheKey = listId + "|" + itemId;
    if (sharingApiSkipCache[cacheKey] === "skip") return [];
    try {
      var digest = await getRequestDigest(webUrl);
      var expand = "$expand=permissionsInformation,pickerSettings";
      var urls = [
        webUrl + "/_api/web/lists(guid'" + listId + "')/items(" + itemId + ")/GetSharingInformation?" + expand,
        webUrl + "/_api/web/lists(guid'" + listId + "')/GetItemById(" + itemId + ")/GetSharingInformation?" + expand
      ];
      for (var ui = 0; ui < urls.length; ui++) {
        var r = await fetchSharingPostWithRetry(urls[ui], digest, 12);
        if (!r) continue;
        if (r.status === 400 || r.status === 403 || r.status === 404) continue;
        if (!r.ok) continue;
        var links = parseSharingLinksFromResponse(await r.json());
        return links;
      }
      sharingApiSkipCache[cacheKey] = "skip";
      return [];
    } catch (_) {
      sharingApiSkipCache[cacheKey] = "skip";
      return [];
    }
  }

  function pushSharingLinkRow(link, linkUrl, entry) {
    var details = link.linkDetails || link.LinkDetails || {};
    var roles = link.roles || link.Roles;
    var rolesText = Array.isArray(roles) ? roles.join("|") : (roles == null ? "" : String(roles));
    sharingRows.push([
      "SharingLink",
      linkUrl,
      details.type || details.Type || link.linkKind || link.LinkKind || "",
      details.scope || details.Scope || "",
      rolesText,
      linkUsersFromSharingLink(link),
      sharingLinkHasExternal(link) ? "Yes" : "No",
      link.expirationDateTime || link.ExpirationDateTime || details.expiration || details.Expiration || "",
      details.blocksDownload != null ? details.blocksDownload : (details.BlocksDownload != null ? details.BlocksDownload : ""),
      link.hasPassword != null ? link.hasPassword : (link.HasPassword != null ? link.HasPassword : ""),
      link.id || link.Id || link.shareId || link.ShareId || "",
      entry.siteUrl,
      entry.listTitle,
      entry.listUrl,
      entry.objectType,
      entry.itemName,
      entry.itemPath,
      entry.itemId,
      ""
    ]);
  }

  function pushDirectPermissionRow(entry) {
    sharingRows.push([
      "DirectPermission", "", "", "", "", "",
      entry.hasExt ? "Yes" : "", "", "", "", "",
      entry.siteUrl, entry.listTitle, entry.listUrl,
      entry.objectType, entry.itemName, entry.itemPath, entry.itemId, entry.accessText
    ]);
  }

  async function flushSharingQueue(listTitle) {
    if (!sharingQueue.length) return;
    var queue = sharingQueue.slice();
    sharingQueue = [];
    logProgress(progressMsg("Fetching sharing links for " + queue.length + " item(s)" + (listTitle ? " in " + listTitle : "") + "…"));
    for (var i = 0; i < queue.length; i++) {
      var entry = queue[i];
      logProgress(progressMsg("Sharing links " + (i + 1) + "/" + queue.length + " — " + (entry.itemName || entry.itemPath)));
      var links = [];
      if (SHARING_LINK_FETCH_ALL || entry.needsLinkApi) {
        links = await fetchSharingLinksForItem(entry.webUrl, entry.listId, entry.itemId);
      }
      var added = false;
      for (var li = 0; li < links.length; li++) {
        var link = links[li];
        var details = link.linkDetails || link.LinkDetails || link;
        var linkUrl = String(details.url || details.Url || details.webUrl || details.WebUrl || link.shareUrl || link.ShareUrl || "").trim();
        if (!linkUrl || sharingLinkUrlSet[linkUrl.toLowerCase()]) continue;
        sharingLinkUrlSet[linkUrl.toLowerCase()] = true;
        pushSharingLinkRow(link, linkUrl, entry);
        added = true;
      }
      if (!added) pushDirectPermissionRow(entry);
    }
  }

  function pushMatrixRows(target, siteName, itemPath, itemType, inheritance, details, principals, roleNames) {
    for (var i = 0; i < principals.length; i++) {
      var p = principals[i];
      var rowInh = p.inheritanceOverride || inheritance;
      var rowDetails = p.details || details || "";
      var row = [
        siteName,
        matrixItemName(itemPath, itemType, siteName),
        itemPath,
        itemType,
        rowInh,
        rowDetails,
        p.title,
        p.principalType,
        p.loginName,
        isExternalUser(p.loginName, p.email) ? "Yes" : "",
        p.givenThrough || "Explicit"
      ];
      for (var ri = 0; ri < roleNames.length; ri++) {
        row.push(p.roleSet && p.roleSet[roleNames[ri]] ? "X" : "");
      }
      if (shouldEmitMatrixRow(row, roleNames)) target.push(row);
    }
  }

  function itemTypeFromList(baseTemplate, fsObjType) {
    if (parseInt(baseTemplate, 10) === 101) {
      return parseInt(fsObjType, 10) === 1 ? "Folder" : "File";
    }
    return "List item";
  }

  function sharingObjectType(baseTemplate, fsObjType) {
    if (parseInt(baseTemplate, 10) === 101) {
      return parseInt(fsObjType, 10) === 1 ? "Folder" : "File";
    }
    return "Item";
  }

  function fmtNum(n) {
    if (typeof n !== "number" || isNaN(n)) return n == null ? "" : String(n);
    return n.toLocaleString("en-US");
  }

  function dashCell(cells, r, c, v, style) {
    cells.push({ r: r, c: c, v: v, s: style == null ? 0 : style });
  }

  function buildHeroInlineXml(title, subtitle) {
    function run(text, opts) {
      var pr = "<rPr>";
      if (opts.bold) pr += "<b/>";
      pr += '<sz val="' + opts.sz + '"/>';
      pr += '<color rgb="' + opts.color + '"/>';
      pr += '<rFont val="Segoe UI"/>';
      pr += "</rPr>";
      return "<r>" + pr + '<t xml:space="preserve">' + sanitizeXmlText(text) + "</t></r>";
    }
    return "<is>" +
      run(title + "\n", { bold: true, sz: 22, color: "FFFFFFFF" }) +
      run(subtitle, { bold: true, sz: 10, color: "FFBACCE0" }) +
      "</is>";
  }

  function dashMargins(cells, r) {
    dashCell(cells, r, 0, "", 21);
    dashCell(cells, r, 1, "", 21);
    dashCell(cells, r, 12, "", 21);
    dashCell(cells, r, 13, "", 21);
  }

  function dashFillBand(cells, r, c1, c2, style) {
    for (var c = c1; c <= c2; c++) dashCell(cells, r, c, "", style);
  }

  function overviewValueStyle(val, sectionLead, firstSection) {
    if (sectionLead) return firstSection ? 10 : 8;
    var cv = coerceCellValue(val);
    if (typeof cv === "number" && !isNaN(cv)) return 15;
    var s = String(val == null ? "" : val).trim();
    if (/^-?\d+$/.test(s) || /^\d{1,3}(,\d{3})+$/.test(s)) return 15;
    if (/^https?:\/\//i.test(s)) return 18;
    return 15;
  }

  function buildSummaryDashboard(sourceUrl, elapsedSec, stats, siteDetails) {
    var PAD = 2;
    function SR(n) { return n + PAD - 1; }
    function SC(n) { return n + PAD - 1; }

    var generated = new Date();
    var pad = function (n) { return n < 10 ? "0" + n : String(n); };
    var genStr = generated.getFullYear() + "-" + pad(generated.getMonth() + 1) + "-" + pad(generated.getDate()) + " " +
      pad(generated.getHours()) + ":" + pad(generated.getMinutes()) + ":" + pad(generated.getSeconds());
    var elapsed = elapsedSec >= 3600 ? Math.floor(elapsedSec / 3600) + "h " + Math.floor((elapsedSec % 3600) / 60) + "m " + (elapsedSec % 60) + "s" :
      (elapsedSec >= 60 ? Math.floor(elapsedSec / 60) + "m " + (elapsedSec % 60) + "s" : elapsedSec + "s");

    var overviewRows = [];
    function addLine(section, metric, value) {
      overviewRows.push([section, metric, value == null ? "" : String(value)]);
    }
    addLine("Report", "Generated at", genStr);
    addLine("Report", "Source site URL", sourceUrl);
    addLine("Report", "Elapsed time", elapsed);
    addLine("Report", "Include subsites", INCLUDE_SUBSITES ? "Yes" : "No");
    addLine("Report", "Expand groups", EXPAND_GROUPS ? "Yes" : "No");
    addLine("Report", "Include all inherited matrix rows", INCLUDE_ALL_INHERITED ? "Yes" : "No");
    addLine("Report", "Max list items for full matrix", MAX_LIST_ITEMS);
    addLine("Report", "List item page size", PAGE_SIZE);
    addLine("Report", "Include folder sharing links", INCLUDE_FOLDER_SHARING_LINKS ? "Yes" : "No");
    addLine("Report", "Fetch sharing link URLs (REST)", SHARING_LINK_FETCH_ALL ? "All unique items" : "Sharing-link principals only");
    addLine("Report", "Export engine", "SharePoint Toolkit (browser extension)");
    addLine("Scope", "Sites and subsites scanned", stats.sites);
    addLine("Scope", "Root sites", stats.rootSites);
    addLine("Scope", "Subsites", stats.subsites);
    addLine("Scope", "Lists", stats.lists);
    addLine("Scope", "Libraries", stats.libraries);
    addLine("Content", "Total SharePoint items scanned", stats.itemsScanned);
    addLine("Content", "List items", stats.listItems);
    addLine("Content", "Files", stats.files);
    addLine("Content", "Folders", stats.folders);
    addLine("Permissions", "Items/folders/files with unique permissions", stats.uniqueItems);
    addLine("Permissions", "Sites with unique permissions", stats.sitesWithUnique);
    addLine("Permissions", "Lists/libraries with unique permissions", stats.listsWithUnique);
    addLine("Principals", "Sharing link permission rows", stats.sharingLinkRows);
    addLine("Principals", "External users (distinct)", stats.externalUsers || 0);
    addLine("Coverage", "Distinct item paths in matrix", stats.matrixPaths);
    addLine("Coverage", "Matrix rows", stats.matrixRows);
    addLine("Coverage", "All Items sheet rows", stats.itemsScanned);
    addLine("Sharing links", "Sharing link sheet rows", stats.sharingSheetRows);
    addLine("Groups", "Group member rows", stats.groupMemberRows);

    var cells = [];
    var merges = [];
    var heroSubtitle = "Source: " + sourceUrl + "  |  Generated: " + genStr + "  |  Elapsed: " + elapsed;
    cells.push({
      r: SR(1), c: SC(1), s: 1,
      inlineXml: buildHeroInlineXml("Permissions scan summary", heroSubtitle)
    });
    dashFillBand(cells, SR(1), SC(1) + 1, SC(10), 1);
    dashFillBand(cells, SR(2), SC(1), SC(10), 1);
    merges.push({ r1: SR(1), c1: SC(1), r2: SR(2), c2: SC(10) });
    dashFillBand(cells, SR(3), SC(1), SC(10), 22);
    merges.push({ r1: SR(3), c1: SC(1), r2: SR(3), c2: SC(10) });

    var totalSites = (stats.rootSites || 0) + (stats.subsites || 0);
    dashFillBand(cells, SR(5), SC(1), SC(3), 21);
    dashFillBand(cells, SR(6), SC(1), SC(3), 21);
    dashCell(cells, SR(5), SC(4), "", 21);
    dashCell(cells, SR(6), SC(4), "", 21);
    dashCell(cells, SR(5), SC(5), "Sites\n" + fmtNum(totalSites), 2);
    merges.push({ r1: SR(5), c1: SC(5), r2: SR(5), c2: SC(7) });
    dashCell(cells, SR(5), SC(8), "Items scanned\n" + fmtNum(stats.itemsScanned), 3);
    merges.push({ r1: SR(5), c1: SC(8), r2: SR(5), c2: SC(10) });
    dashCell(cells, SR(6), SC(5), "External users\n" + fmtNum(stats.externalUsers || 0), 4);
    merges.push({ r1: SR(6), c1: SC(5), r2: SR(6), c2: SC(7) });
    dashCell(cells, SR(6), SC(8), "Unique permissions\n" + fmtNum(stats.uniqueItems), 5);
    merges.push({ r1: SR(6), c1: SC(8), r2: SR(6), c2: SC(10) });
    dashFillBand(cells, SR(7), SC(1), SC(10), 21);
    merges.push({ r1: SR(7), c1: SC(1), r2: SR(7), c2: SC(10) });

    dashCell(cells, SR(8), SC(1), "SCAN DETAILS", 23);
    merges.push({ r1: SR(8), c1: SC(1), r2: SR(8), c2: SC(3) });
    dashFillBand(cells, SR(8), SC(4), SC(10), 21);

    dashCell(cells, SR(9), SC(1), "Section", 24);
    dashCell(cells, SR(9), SC(2), "Metric", 24);
    dashCell(cells, SR(9), SC(3), "Value", 24);
    dashFillBand(cells, SR(9), SC(4), SC(10), 21);

    var rowHeights = {};
    rowHeights["0"] = 18;
    rowHeights["1"] = 18;
    rowHeights[String(SR(1))] = 44;
    rowHeights[String(SR(2))] = 6;
    rowHeights[String(SR(3))] = 5;
    rowHeights[String(SR(5))] = 52;
    rowHeights[String(SR(6))] = 52;
    rowHeights[String(SR(7))] = 10;
    rowHeights[String(SR(8))] = 20;
    rowHeights[String(SR(9))] = 28;

    var overviewStart = SR(10);
    var dataRow = overviewStart;
    var prevSec = "";
    var firstSection = true;
    for (var oi = 0; oi < overviewRows.length; oi++) {
      var sec = overviewRows[oi][0];
      var metric = overviewRows[oi][1];
      var val = overviewRows[oi][2];
      var isNewSec = sec !== prevSec;
      if (isNewSec) {
        if (firstSection) {
          dashCell(cells, dataRow, SC(1), sec, 11);
          dashCell(cells, dataRow, SC(2), metric, 9);
          dashCell(cells, dataRow, SC(3), val, overviewValueStyle(val, true, true));
          firstSection = false;
        } else {
          dashCell(cells, dataRow, SC(1), sec, 6);
          dashCell(cells, dataRow, SC(2), metric, 7);
          dashCell(cells, dataRow, SC(3), val, overviewValueStyle(val, true, false));
        }
        prevSec = sec;
      } else {
        dashCell(cells, dataRow, SC(1), "", 12);
        dashCell(cells, dataRow, SC(2), metric, 13);
        dashCell(cells, dataRow, SC(3), val, overviewValueStyle(val, false, false));
      }
      dashFillBand(cells, dataRow, SC(4), SC(10), 21);
      rowHeights[String(dataRow)] = 20;
      dataRow++;
    }

    var spacerRow = dataRow;
    dashFillBand(cells, spacerRow, SC(1), SC(10), 21);
    rowHeights[String(spacerRow)] = 10;

    var siteLabelRow = spacerRow + 1;
    dashCell(cells, siteLabelRow, SC(1), "SITE BREAKDOWN", 25);
    merges.push({ r1: siteLabelRow, c1: SC(1), r2: siteLabelRow, c2: SC(10) });
    rowHeights[String(siteLabelRow)] = 18;

    var headerRow = siteLabelRow + 1;
    var siteHeaders = ["Site Name", "Site Path", "Type", "Lists/Libraries", "Items", "Files", "Folders", "List items", "Unique items", "External users", "Site unique perms"];
    for (var hi = 0; hi < siteHeaders.length; hi++) {
      dashCell(cells, headerRow, SC(1 + hi), siteHeaders[hi], 24);
    }
    rowHeights[String(headerRow)] = 28;

    var siteRow = headerRow + 1;
    for (var si = 0; si < siteDetails.length; si++) {
      var sd = siteDetails[si];
      var siteValues = [
        sd.siteName, sd.sitePath, sd.type, sd.listsLibs, sd.items, sd.files, sd.folders,
        sd.listItems, sd.uniqueItems, sd.externalUsers != null ? sd.externalUsers : 0, sd.siteUniquePerms
      ];
      for (var hci = 0; hci < siteValues.length; hci++) {
        dashCell(cells, siteRow, SC(1 + hci), siteValues[hci], (hci % 2 === 0) ? 6 : 26);
      }
      rowHeights[String(siteRow)] = 18;
      siteRow++;
    }

    var lastRow = siteRow - 1;
    for (var pr = 0; pr <= lastRow; pr++) dashMargins(cells, pr);
    dashFillBand(cells, 0, 0, 13, 21);
    dashFillBand(cells, 1, 0, 13, 21);

    return {
      cells: cells,
      merges: merges,
      colHiddenFrom: 12,
      rowHeights: rowHeights,
      colWidths: [10, 10, 0.6, 40, 26, 30, 16, 16, 16, 16, 16, 16, 12, 10]
    };
  }

  var matrixRows = [];
  var allItemRows = [];
  var groupMemberRows = [];
  var sharingRows = [];
  var groupCatalog = {};
  var registeredGroups = {};
  var matrixPathSet = {};
  var externalUserLogins = {};
  var externalUsersBySite = {};
  var stats = {
    sites: 0, rootSites: 0, subsites: 0, lists: 0, libraries: 0, uniqueItems: 0,
    itemsScanned: 0, listItems: 0, files: 0, folders: 0, sitesWithUnique: 0, listsWithUnique: 0,
    sharingLinkRows: 0, matrixRows: 0, matrixPaths: 0, groupMemberRows: 0, sharingSheetRows: 0
  };
  var siteDetails = [];

  (async function run() {
    var runStart = Date.now();
    try {
      if (!siteUrl) {
        reportDone(false, "Could not detect site. Open a SharePoint site, list, or library page.");
        return;
      }

      setProgress(progressMsg("Discovering site structure…"), 2);
      startProgressPulse();
      checkCancelled();
      var webs = await fetchAllWebs(siteUrl, INCLUDE_SUBSITES);
      if (SELECTED_SITE_PATHS && SELECTED_SITE_PATHS.length) {
        webs = webs.filter(function (w) {
          return SELECTED_SITE_PATHS.indexOf((w.serverRelativeUrl || "").toLowerCase()) >= 0;
        });
      }
      if (!webs.length) {
        reportDone(false, "No webs found to scan.");
        return;
      }
      setProgress(progressMsg("Found " + webs.length + " site(s) to scan…"), 5);

      var roleNames = await fetchRoleNames(siteUrl);
      var matrixHeader = ["Site Name", "Name", "Item path", "Item Type", "Inheritance", "Details", "User/group", "Principal type", "Account name", "External user", "Given through"].concat(roleNames);
      var allItemsHeader = ["Site Name", "List/Library", "Item path", "Item Type", "Item Id", "Broken permissions", "Created", "Created By", "Modified", "Modified By"];
      var groupHeader = ["Site Name", "Site URL", "Site Path", "Group Name", "Group Login", "Group Id", "Member Name", "Member Login", "Member Email", "Member Type", "External user"];
      var bagelFolderHeader = ["Site", "Library", "FolderPath", "FolderName", "FolderLevel", "DirectFoldersCount", "DirectFilesCount", "TotalNestedFolders", "LevelsDeep", "TotalItemsRecursive", "ServerRelativeUrl"];
      var bagelPathHeader = ["Site", "Library", "Path", "EncodedPath", "FriendlyCharCount", "EncodedCharCount"];
      var bagelFolderRows = [];
      var bagelPathRows = [];

      var rootPath = webs[0].serverRelativeUrl;
      var accept = "application/json;odata=nometadata";
      var acceptVerbose = { Accept: "application/json;odata=verbose" };

      setProgress(progressMsg("Inventorying lists…"), 4);
      var webScanPlans = [];
      progressPlan.websTotal = webs.length;
      for (var wpi = 0; wpi < webs.length; wpi++) {
        var invWeb = webs[wpi];
        var invListsResp = await fetchJson(invWeb.url + "/_api/web/lists?$select=Id,Title,BaseTemplate,HasUniqueRoleAssignments,ItemCount,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden eq false&$top=5000", accept);
        var invLists = invListsResp.value || invListsResp.results || [];
        var invScannable = [];
        for (var ilsi = 0; ilsi < invLists.length; ilsi++) {
          var ilt = (invLists[ilsi].Title || invLists[ilsi].title || "").trim();
          if (!ilt || EXCLUDED_LISTS[ilt]) continue;
          var invListId = normalizeGuid(invLists[ilsi].Id || invLists[ilsi].id);
          if (!listIsSelectedForMatrix(invWeb.url, invListId)) continue;
          invScannable.push(invLists[ilsi]);
          progressPlan.totalItems += Math.max(0, parseInt(invLists[ilsi].ItemCount, 10) || 0);
        }
        if (!invScannable.length && SELECTED_LIST_KEYS) continue;
        webScanPlans.push({ web: invWeb, lists: invScannable });
        progressPlan.totalLists += invScannable.length;
        setProgress(progressMsg(invWeb.title + " — " + invScannable.length + " list(s)"), 4 + ((wpi + 1) / webs.length) * 1);
      }
      if (!webScanPlans.length) {
        reportDone(false, "No selected lists/libraries found to scan.");
        return;
      }
      setProgress(progressMsg("Ready — " + webScanPlans.length + " site(s), " + progressPlan.totalLists + " list(s), " + fmtProgressCount(progressPlan.totalItems) + " item(s)"), 5);

      for (var wi = 0; wi < webScanPlans.length; wi++) {
        checkCancelled();
        var scanPlan = webScanPlans[wi];
        var web = scanPlan.web;
        var scannableLists = scanPlan.lists;
        var webUrl = web.url;
        var siteName = web.title;
        var webPath = web.serverRelativeUrl;
        var webType = webItemType(webPath, rootPath);
        var siteCtx = { siteName: siteName, webUrl: web.absoluteUrl || webUrl, sitePath: webPath };
        stats.sites++;
        if (webType === "Subsite") stats.subsites++; else stats.rootSites++;
        if (web.hasUnique) stats.sitesWithUnique++;

        var siteDetail = {
          siteName: siteName, sitePath: webPath, type: webType,
          listsLibs: 0, items: 0, files: 0, folders: 0, listItems: 0, uniqueItems: 0,
          siteUniquePerms: web.hasUnique ? "Yes" : "No"
        };

        logProgress(progressMsg(siteName + " (" + (wi + 1) + "/" + webs.length + ")…"));

        if (!(webType === "Subsite" && !web.hasUnique)) {
          var webAssignments = await fetchRoleAssignments(webUrl + "/_api/web/roleassignments", accept);
          var webPrincipals = await principalsFromAssignments(webUrl, webAssignments, accept, groupCatalog, siteCtx);
          var before = matrixRows.length;
          pushMatrixRows(matrixRows, siteName, webPath, webType, web.hasUnique ? "Custom" : "Top level", "", webPrincipals, roleNames);
          if (matrixRows.length > before) matrixPathSet[webPath] = true;
        }

        for (var li = 0; li < scannableLists.length; li++) {
          checkCancelled();
          var lst = scannableLists[li];
          var listTitle = (lst.Title || lst.title || "").trim();
          if (!listTitle || EXCLUDED_LISTS[listTitle]) continue;
          var listId = normalizeGuid(lst.Id || lst.id);
          if (!listId) continue;
          var listBase = webUrl + "/_api/web/lists(guid'" + listId + "')";
          var baseTemplate = lst.BaseTemplate != null ? lst.BaseTemplate : lst.baseTemplate;
          var isLibrary = parseInt(baseTemplate, 10) === 101;
          siteDetail.listsLibs++;
          if (isLibrary) stats.libraries++; else stats.lists++;
          var listHasUnique = lst.HasUniqueRoleAssignments === true;
          if (listHasUnique) stats.listsWithUnique++;
          var itemCount = Math.max(0, parseInt(lst.ItemCount, 10) || 0);
          var listRootUrl = (lst.RootFolder && (lst.RootFolder.ServerRelativeUrl || lst.RootFolder.serverRelativeUrl)) || "";
          var listRootType = isLibrary ? "Library" : "List";
          var skipNonUniqueMatrix = itemCount > MAX_LIST_ITEMS && !INCLUDE_ALL_INHERITED;

          try {
            if (listHasUnique) {
              var listAssignments = await fetchRoleAssignments(listBase + "/roleassignments", accept);
              var listPrincipals = await principalsFromAssignments(webUrl, listAssignments, accept, groupCatalog, siteCtx);
              var lb = matrixRows.length;
              pushMatrixRows(matrixRows, siteName, listRootUrl, listRootType, "Custom", "", listPrincipals, roleNames);
              if (matrixRows.length > lb) matrixPathSet[listRootUrl] = true;
            }

            var allItems = [];
            var nextUrl = listBase + "/items?$select=Id,FileRef,FileDirRef,FileLeafRef,FSObjType,HasUniqueRoleAssignments,Created,Modified,Author/Title,Editor/Title&$expand=Author,Editor&$top=" + PAGE_SIZE;
            var pageNum = 0;
            progressPlan.currentListLoaded = 0;
            setProgress(progressMsg("Scanning " + listTitle + (itemCount ? " (~" + fmtProgressCount(itemCount) + " items)" : "") + "…"), itemProgressPct(0));
            while (nextUrl) {
              checkCancelled();
              pageNum++;
              var itemResp = await fetchWithRetry(normalizeApiUrl(nextUrl), { credentials: "include", headers: acceptVerbose }, 8);
              if (!itemResp.ok) throw new Error("HTTP " + itemResp.status);
              var d = await itemResp.json();
              var batch = (d.d && d.d.results) || d.value || [];
              allItems = allItems.concat(batch);
              progressPlan.currentListLoaded = allItems.length;
              if (pageNum === 1 || pageNum % 3 === 0) {
                setProgress(progressMsg(listTitle + " — loaded " + fmtProgressCount(allItems.length) + " item(s)…"), itemProgressPct(allItems.length));
              }
              nextUrl = (d.d && d.d.__next) ? normalizeApiUrl(d.d.__next) : (d["@odata.nextLink"] ? normalizeApiUrl(d["@odata.nextLink"]) : null);
              if (!nextUrl || batch.length === 0) break;
              await sleep(80);
            }
            if (itemCount > 0 && allItems.length < itemCount && allItems.length < itemCount * 0.9) {
              logProgress(progressMsg(listTitle + " returned " + allItems.length + " of ~" + itemCount + " items"));
            }
            if (IS_BAGEL && isLibrary && allItems.length) {
              appendBagelFolderCountRows(bagelFolderRows, allItems, siteName, listTitle);
              appendBagelPathLengthRows(bagelPathRows, allItems, siteName, listTitle);
            }

            var uniqueInList = 0;
            for (var ii = 0; ii < allItems.length; ii++) {
              if (ii > 0 && ii % 25 === 0) checkCancelled();
              var itm = allItems[ii];
              var itemPath = itm.FileRef || (itm.FileDirRef && itm.FileLeafRef ? (itm.FileDirRef + "/" + itm.FileLeafRef).replace(/\/+/g, "/") : null) || itm.FileLeafRef || "";
              if (itemPath && itemPath.indexOf("#") >= 0) itemPath = itemPath.slice(itemPath.indexOf("#") + 1).trim();
              var itemType = itemTypeFromList(baseTemplate, itm.FSObjType);
              var hasUnique = itm.HasUniqueRoleAssignments === true;
              stats.itemsScanned++;
              siteDetail.items++;
              if (itemType === "File") { stats.files++; siteDetail.files++; }
              else if (itemType === "Folder") { stats.folders++; siteDetail.folders++; }
              else { stats.listItems++; siteDetail.listItems++; }
              if (hasUnique) { stats.uniqueItems++; siteDetail.uniqueItems++; }

              allItemRows.push([
                siteName, listTitle, itemPath, itemType, itm.Id || "",
                hasUnique ? "Yes" : "No",
                itm.Created || "", fieldUserDisplayName(itm.Author),
                itm.Modified || "", fieldUserDisplayName(itm.Editor)
              ]);

              if (hasUnique) {
                uniqueInList++;
                if (uniqueInList === 1 || uniqueInList % 10 === 0) {
                  logProgress(progressMsg(listTitle + " — unique permissions " + uniqueInList + "…"));
                }
                var itemAssignments = await fetchRoleAssignments(listBase + "/items(" + itm.Id + ")/RoleAssignments", accept);
                var itemPrincipals = await principalsFromAssignments(webUrl, itemAssignments, accept, groupCatalog, siteCtx);
                var ib = matrixRows.length;
                pushMatrixRows(matrixRows, siteName, itemPath, itemType, "Custom", "", itemPrincipals, roleNames);
                if (matrixRows.length > ib) matrixPathSet[itemPath] = true;
                for (var mr = ib; mr < matrixRows.length; mr++) {
                  if (matrixRows[mr][10] === "Sharing Link") stats.sharingLinkRows++;
                }
                var accessText = buildAccessText(itemAssignments);
                var hasExt = itemAssignments.some(function (a) {
                  var mem = a.Member || a.member;
                  return mem && parseInt(mem.PrincipalType, 10) === 1 && isExternalUser(mem.LoginName || mem.loginName, mem.Email || mem.email);
                });
                queueSharingItem({
                  webUrl: webUrl,
                  listId: listId,
                  listTitle: listTitle,
                  listUrl: listRootUrl,
                  siteUrl: siteCtx.webUrl,
                  objectType: sharingObjectType(baseTemplate, itm.FSObjType),
                  itemName: matrixItemName(itemPath, itemType, siteName),
                  itemPath: itemPath,
                  itemId: itm.Id || "",
                  itemType: itemType,
                  accessText: accessText,
                  hasExt: hasExt,
                  needsLinkApi: assignmentsHaveSharingLink(itemAssignments)
                });
                if (uniqueInList % 5 === 0) await sleep(120);
              } else if (!skipNonUniqueMatrix && INCLUDE_ALL_INHERITED) {
                var inhUrl = listHasUnique ? listBase + "/roleassignments" : webUrl + "/_api/web/roleassignments";
                var inhAssignments = await fetchRoleAssignments(inhUrl, accept);
                var inhPrincipals = await principalsFromAssignments(webUrl, inhAssignments, accept, groupCatalog, siteCtx);
                for (var ip = 0; ip < inhPrincipals.length; ip++) {
                  inhPrincipals[ip].givenThrough = "Inherited (From ItemPath: " + (listHasUnique ? listRootUrl : webPath) + ")";
                }
                var iib = matrixRows.length;
                pushMatrixRows(matrixRows, siteName, itemPath, itemType, "Inherited", "", inhPrincipals, roleNames);
                if (matrixRows.length > iib) matrixPathSet[itemPath] = true;
              }
            }
            await flushSharingQueue(listTitle);
            progressPlan.doneItems += allItems.length;
            progressPlan.currentListLoaded = 0;
            setProgress(progressMsg("Finished " + listTitle), itemProgressPct(0));
          } catch (listErr) {
            progressPlan.doneItems += itemCount;
            progressPlan.currentListLoaded = 0;
            setProgress(progressMsg("Skipped " + listTitle + " (" + ((listErr && listErr.message) || listErr) + ")"), itemProgressPct(0));
          }
        }
        siteDetails.push(siteDetail);
      }

      stats.externalUsers = Object.keys(externalUserLogins).length;
      for (var sdi = 0; sdi < siteDetails.length; sdi++) {
        var sn = siteDetails[sdi].siteName;
        siteDetails[sdi].externalUsers = externalUsersBySite[sn] ? Object.keys(externalUsersBySite[sn]).length : 0;
      }

      stats.matrixRows = matrixRows.length;
      stats.matrixPaths = Object.keys(matrixPathSet).length;
      stats.groupMemberRows = groupMemberRows.length;
      stats.sharingSheetRows = sharingRows.length;

      var elapsedSec = Math.max(0, Math.round((Date.now() - runStart) / 1000));
      var summaryDashboard = buildSummaryDashboard(siteUrl, elapsedSec, stats, siteDetails);

      checkCancelled();
      setProgress(progressMsg("Building workbook…"), 98);
      var defaultLabel = IS_BAGEL ? "EverythingBagel" : "PermissionsMatrix";
      var siteLabel = (webs[0] && webs[0].title) ? webs[0].title.replace(/[^\w\s-]/g, "").trim() : defaultLabel;
      var ts = new Date();
      var tsPad = function (n) { return n < 10 ? "0" + n : String(n); };
      var stamp = ts.getFullYear() + tsPad(ts.getMonth() + 1) + tsPad(ts.getDate()) + "-" + tsPad(ts.getHours()) + tsPad(ts.getMinutes()) + tsPad(ts.getSeconds());
      var fnBase = params.f ? params.f.replace(/\.(csv|xls|xlsx|xml)$/i, "") : (siteLabel + "-" + defaultLabel + "-" + stamp);

      var sheets = [
        { name: "Summary", dashboard: summaryDashboard, tabColor: "FF0D9488" },
        { name: "Permissions Matrix", aoa: [matrixHeader].concat(matrixRows), tableKind: "matrix", roleStartCol: 11, tabColor: "FF0F2744" },
        { name: "Group Members", aoa: [groupHeader].concat(groupMemberRows), tableKind: "groups", tabColor: "FF6366F1" },
        { name: "Sharing Links", aoa: [SHARING_LINK_COLUMNS].concat(sharingRows), tableKind: "sharing", tabColor: "FFEA580C" },
        { name: "All Items", aoa: [allItemsHeader].concat(allItemRows), tableKind: "allItems", tabColor: "FF2563EB" }
      ];
      if (IS_BAGEL) {
        sheets.push(
          { name: "Folder Counts", aoa: [bagelFolderHeader].concat(bagelFolderRows), tableKind: "folderCounts", tabColor: "FF059669" },
          { name: "Path Lengths", aoa: [bagelPathHeader].concat(bagelPathRows), tableKind: "pathLengths", tabColor: "FF7C3AED" }
        );
      }

      await downloadWorkbook(sheets, fnBase);
      stopProgressPulse();
      setProgress(progressMsg("Download complete."), 100);
      reportDone(
        true,
        IS_BAGEL
          ? ("Done! Everything Bagel downloaded (" + stats.matrixRows + " matrix rows, " + bagelFolderRows.length + " folder rows, " + bagelPathRows.length + " path rows, " + stats.sites + " site(s)).")
          : ("Done! Permissions matrix downloaded (" + stats.matrixRows + " matrix rows, " + stats.sites + " site(s)).")
      );
    } catch (e) {
      stopProgressPulse();
      if (exportCancelled || (e && e.message === "Export cancelled.")) {
        reportDone(false, "Export cancelled.");
        return;
      }
      reportDone(false, (e && e.message) || String(e));
    }
  })();
})();
