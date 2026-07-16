// Runs in page context on SharePoint. Creates temp RPC view, pages by ID range via owssvr.dll, exports CSV.
// Uses SharePoint system field names (ID, Title, FileLeafRef, etc.) — same across all lists/libraries.
(function () {
  // -----------------------------
  // Param bootstrap (extension-style)
  // -----------------------------
  var params = null;
  function readParams() {
    try {
      var el = document.getElementById("spcsv-params-json");
      if (el && el.textContent) return JSON.parse(el.textContent);
    } catch (_) {}
    try {
      var s = document.currentScript && document.currentScript.getAttribute("src");
      if (s && s.indexOf("?p=") >= 0) return JSON.parse(decodeURIComponent(s.slice(s.indexOf("?p=") + 3)));
    } catch (e) {}
    return null;
  }
  params = readParams();
  if (!params) {
    console.error("SharePoint CSV Export – Missing params.");
    return;
  }

  var exportReport = params.report || "exportCSV";
  var exportOverallPct = 0;

  try {
    window.postMessage({
      __spcsv: true,
      type: "SPCSVExportStarted",
      detail: {
        message: "Export started…",
        report: exportReport
      }
    }, "*");
  } catch (_) {}

  // -----------------------------
  // Helpers
  // -----------------------------
  function normalizeGuid(g) {
    if (g == null) return "";
    var s = String(g).trim();
    try { s = decodeURIComponent(s); } catch (_) {}
    s = s.replace(/%7B|%7D/ig, "");
    s = s.replace(/[{}]/g, "");
    return s.trim();
  }

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function fmtExportCount(n) {
    n = Math.max(0, Math.round(n || 0));
    try { return n.toLocaleString(); } catch (_) { return String(n); }
  }

  function reportProgress(message, opts) {
    opts = opts || {};
    var detail = { message: message || "", logLine: message || "", report: exportReport };
    if (opts.logLine) detail.logLine = opts.logLine;
    var pct = opts.percent;
    if (pct == null && opts.currentCount != null && opts.totalCount != null && opts.totalCount > 0) {
      pct = Math.min(99, Math.round((opts.currentCount / opts.totalCount) * 100));
    }
    if (pct != null) {
      exportOverallPct = Math.max(exportOverallPct, Math.min(99, Math.round(pct)));
      detail.percent = exportOverallPct;
    } else if (exportOverallPct) {
      detail.percent = exportOverallPct;
    }
    if (opts.currentCount != null && opts.totalCount != null && opts.totalCount > 0) {
      var headline = exportOverallPct + "% — " + fmtExportCount(opts.currentCount) + " of " + fmtExportCount(opts.totalCount);
      if (message) headline += " — " + message;
      detail.message = headline;
      if (!opts.logLine && message) detail.logLine = message;
    } else if (message) {
      detail.message = message;
    }
    window.postMessage({ __spcsv: true, type: "SPCSVExportProgress", detail: detail }, "*");
  }

  function reportDone(success, message, stopReason) {
    if (!success) {
      console.error("SharePoint CSV Export:", message);
      try { alert("SharePoint CSV Export failed:\n\n" + message); } catch (_) {}
    }
    window.postMessage(
      { __spcsv: true, type: "SPCSVExportDone", detail: { success: success, message: message, stopReason: stopReason || "" } },
      "*"
    );
  }

  function cleanValue(val) {
    if (val == null || val === undefined) return "";
    var s = String(val);
    if (s.indexOf(";#") >= 0) {
      var parts = s.split(";#");
      return parts.filter(function (_, i) { return i % 2 === 1; }).join(", ") || s;
    }
    return s;
  }

  function escapeCSV(val) {
    var s = cleanValue(val);
    if (s.indexOf('"') >= 0 || s.indexOf(",") >= 0 || s.indexOf("\n") >= 0 || s.indexOf("\r") >= 0) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function escXml(s) {
    if (s == null || s === undefined) return "";
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function parseCSVToAoa(csvStr) {
    var s = (csvStr || "").replace(/\ufeff/g, "");
    var lines = [];
    var inQuotes = false;
    var field = "";
    var row = [];
    for (var i = 0; i < s.length; i++) {
      var c = s.charAt(i);
      if (inQuotes) {
        if (c === '"') {
          if (s.charAt(i + 1) === '"') { field += '"'; i++; }
          else inQuotes = false;
        } else field += c;
      } else {
        if (c === '"') { inQuotes = true; }
        else if (c === ",") { row.push(field); field = ""; }
        else if (c === "\n" || c === "\r") {
          if (c === "\r" && s.charAt(i + 1) === "\n") i++;
          row.push(field);
          lines.push(row);
          row = [];
          field = "";
        } else field += c;
      }
    }
    row.push(field);
    if (row.length > 1 || field !== "") lines.push(row);
    return lines;
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

  function aoaToExcelXml(aoa, sheetName) {
    sheetName = sheetName || "Sheet1";
    var headers = aoa[0] || [];
    var headerRow = headers
      .map(function (h) { return "<Cell ss:StyleID=\"Header\"><Data ss:Type=\"String\">" + escXml(h) + "</Data></Cell>"; })
      .join("");
    var dataRows = [];
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var cells = headers
        .map(function (_, c) {
          var v = row[c];
          var isNum = typeof v === "number" && !isNaN(v);
          var val = v == null ? "" : (isNum ? v : escXml(v));
          return "<Cell><Data ss:Type=\"" + (isNum ? "Number" : "String") + "\">" + val + "</Data></Cell>";
        })
        .join("");
      dataRows.push("<Row>" + cells + "</Row>");
    }
    var tableRows = "<Row>" + headerRow + "</Row>" + dataRows.join("");
    var cols = headers.map(function () { return "<Column ss:Width=\"120\"/>"; }).join("");
    var nr = aoa.length;
    var nc = headers.length;
    var tableAttrs = " ss:ExpandedColumnCount=\"" + nc + "\" ss:ExpandedRowCount=\"" + nr + "\"";
    var worksheetOptions = "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><FilterOn/></WorksheetOptions>";
    return "<?xml version=\"1.0\"?>\n<?mso-application progid=\"Excel.Sheet\"?>\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\n<Styles>\n<Style ss:ID=\"Default\" ss:Name=\"Normal\"><Font ss:Size=\"11\"/></Style>\n<Style ss:ID=\"Header\"><Font ss:Bold=\"1\" ss:Size=\"11\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#4472C4\" ss:Pattern=\"Solid\"/></Style>\n</Styles>\n<Worksheet ss:Name=\"" + escXml(sheetName) + "\">\n<Table" + tableAttrs + ">\n" + cols + "\n" + tableRows + "\n</Table>\n" + worksheetOptions + "\n</Worksheet>\n</Workbook>";
  }

  function aoaToXlsxBlob(aoa, sheetName) {
    return aoaToXlsxBlobMulti([{ name: sheetName || "Sheet1", aoa: aoa }]);
  }

  function aoaToXlsxBlobMulti(sheets) {
    if (!window.JSZip) return Promise.reject(new Error("JSZip not loaded"));
    if (!sheets || sheets.length === 0) return Promise.reject(new Error("No sheets"));
    var sharedStrings = [];
    var sharedIndex = {};
    function getStrIndex(val) {
      var s = val == null ? "" : String(val);
      if (sharedIndex[s] === undefined) {
        sharedIndex[s] = sharedStrings.length;
        sharedStrings.push(s);
      }
      return sharedIndex[s];
    }
    function buildSheetRows(aoa) {
      var nr = aoa.length;
      var nc = nr > 0 ? (aoa[0] || []).length : 0;
      var sheetRows = [];
      for (var r = 0; r < nr; r++) {
        var row = aoa[r];
        var cells = [];
        for (var c = 0; c < nc; c++) {
          var v = row[c];
          if (typeof v === "string" && v.trim() !== "") {
            var s = v.trim();
            if (/^-?\d+$/.test(s)) v = parseInt(s, 10);
            else if (/^-?\d+\.\d*$/.test(s) || /^-?\d*\.\d+$/.test(s)) { var n = parseFloat(s); if (!isNaN(n)) v = n; }
          }
          var ref = colToLetter(c) + (r + 1);
          var isNum = typeof v === "number" && !isNaN(v);
          if (isNum) cells.push("<c r=\"" + ref + "\" t=\"n\"><v>" + v + "</v></c>");
          else cells.push("<c r=\"" + ref + "\" t=\"s\"><v>" + getStrIndex(v) + "</v></c>");
        }
        sheetRows.push("<row r=\"" + (r + 1) + "\">" + cells.join("") + "</row>");
      }
      return { rows: sheetRows, nr: nr, nc: nc };
    }
    var sheetXmls = [];
    for (var s = 0; s < sheets.length; s++) {
      var sh = sheets[s];
      var aoa = sh.aoa || [];
      var built = buildSheetRows(aoa);
      var filterRange = built.nc > 0 && built.nr > 0 ? "A1:" + colToLetter(built.nc - 1) + built.nr : "A1";
      sheetXmls.push("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"" + filterRange + "\"/><sheetData>" + built.rows.join("") + "</sheetData><autoFilter ref=\"" + filterRange + "\"/></worksheet>");
    }
    var sstItems = sharedStrings.map(function (s) {
      var t = escXml(s);
      if (/[\s\u00a0]/.test(s) || s.length !== s.trim().length) return "<si><t xml:space=\"preserve\">" + t + "</t></si>";
      return "<si><t>" + t + "</t></si>";
    });
    var sstXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"" + sharedStrings.length + "\" uniqueCount=\"" + sharedStrings.length + "\">" + sstItems.join("") + "</sst>";
    var contentTypesParts = ["<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>", "<Default Extension=\"xml\" ContentType=\"application/xml\"/>", "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"];
    for (var i = 0; i < sheets.length; i++) contentTypesParts.push("<Override PartName=\"/xl/worksheets/sheet" + (i + 1) + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
    contentTypesParts.push("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>", "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
    var contentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" + contentTypesParts.join("") + "</Types>";
    var workbookRelsParts = [];
    for (var j = 0; j < sheets.length; j++) workbookRelsParts.push("<Relationship Id=\"rId" + (j + 1) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet" + (j + 1) + ".xml\"/>");
    workbookRelsParts.push("<Relationship Id=\"rId" + (sheets.length + 1) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>", "<Relationship Id=\"rId" + (sheets.length + 2) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
    var workbookRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + workbookRelsParts.join("") + "</Relationships>";
    var sheetEls = [];
    for (var k = 0; k < sheets.length; k++) sheetEls.push("<sheet name=\"" + escXml((sheets[k].name || "Sheet" + (k + 1)).toString().slice(0, 31)) + "\" sheetId=\"" + (k + 1) + "\" r:id=\"rId" + (k + 1) + "\"/>");
    var workbookXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>" + sheetEls.join("") + "</sheets></workbook>";
    var stylesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs></styleSheet>";
    var rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>";
    var zip = new window.JSZip();
    zip.file("[Content_Types].xml", contentTypes);
    zip.file("_rels/.rels", rels);
    zip.file("xl/workbook.xml", workbookXml);
    zip.file("xl/_rels/workbook.xml.rels", workbookRels);
    for (var z = 0; z < sheetXmls.length; z++) zip.file("xl/worksheets/sheet" + (z + 1) + ".xml", sheetXmls[z]);
    zip.file("xl/sharedStrings.xml", sstXml);
    zip.file("xl/styles.xml", stylesXml);
    return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  function downloadExport(data, filenameBase, opts) {
    opts = opts || {};
    var isXls = params.format === "xls";
    var isXlsx = params.format === "xlsx";
    var ext = isXlsx ? ".xlsx" : (isXls ? ".xls" : ".csv");
    var fn = (filenameBase || "export").replace(/\.(csv|xls|xlsx|xml)$/i, "") + ext;
    if (isXls || isXlsx) {
      var aoa;
      if (data.sheets && data.sheets.length > 0 && isXlsx) {
        aoa = null;
      } else if (data.csv != null) {
        aoa = parseCSVToAoa(data.csv);
      } else if (data.rows && data.columns) {
        var cols = data.columns;
        aoa = [cols.slice()];
        for (var i = 0; i < data.rows.length; i++) {
          aoa.push(cols.map(function (c) { return getCellValue(data.rows[i], c); }));
        }
      } else {
        reportDone(false, "Excel: no data");
        return;
      }
      if (isXlsx) {
        var xlsxSheets = data.sheets && data.sheets.length > 0 ? data.sheets : [{ name: opts.sheetName || "Sheet1", aoa: aoa }];
        aoaToXlsxBlobMulti(xlsxSheets).then(function (blob) {
          var link = document.createElement("a");
          link.href = URL.createObjectURL(blob);
          link.download = fn;
          link.click();
          URL.revokeObjectURL(link.href);
        }).catch(function (err) {
          reportDone(false, err && err.message ? err.message : "XLSX export failed.");
        });
        return;
      }
      var xml = aoaToExcelXml(aoa, opts.sheetName || "Sheet1");
      var blob = new Blob([xml], { type: "application/vnd.ms-excel" });
      var link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = fn;
      link.click();
      URL.revokeObjectURL(link.href);
    } else {
      var csv = data.csv != null ? data.csv : (data.rows && data.columns ? toCSVFromRows(data.rows, data.columns) : "");
      if (data.rows && data.columns && csv.indexOf("\ufeff") !== 0) csv = "\ufeff" + csv;
      var blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
      var link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = fn;
      link.click();
      URL.revokeObjectURL(link.href);
    }
  }

  var SCHEMA_ATTRS = { "id": 1, "name": 1, "dt": 1, "rs": 1, "s": 1, "z": 1, "commandcontent": 1 };

  function preprocessOwssvrXml(xmlText) {
    var str = xmlText;
    function stripDup(str, attrName) {
      return str.replace(/<([a-z:]*row)([^>]*)>/gi, function (match, tagName, attrs) {
        var attrRegex = new RegExp("(\\s)" + attrName + '="[^"]*"', "g");
        var first = true;
        var newAttrs = attrs.replace(attrRegex, function (m) {
          if (first) { first = false; return m; }
          return "";
        });
        return "<" + tagName + newAttrs + ">";
      });
    }
    str = stripDup(str, "ows_FileLeafRef");
    str = stripDup(str, "ows_Title");
    str = stripDup(str, "ows_ID");
    return str;
  }

  function parseOwssvrXml(xmlText) {
    xmlText = preprocessOwssvrXml(xmlText);
    var parser = new DOMParser();
    var doc = parser.parseFromString(xmlText, "text/xml");
    var err = doc.querySelector("parsererror");
    if (err) throw new Error("XML parse error: " + (err.textContent || "invalid or duplicate attributes"));
    var rows = [];
    var columns = [];

    var all = doc.getElementsByTagName("*");
    for (var i = 0; i < all.length; i++) {
      var el = all[i];
      var local = (el.localName || el.nodeName || "").toLowerCase();
      if (local !== "row") continue;

      var row = {};
      for (var j = 0; j < el.attributes.length; j++) {
        var attr = el.attributes[j];
        var attrName = (attr.name || attr.nodeName || attr.localName || "");
        if (SCHEMA_ATTRS[attrName.toLowerCase()]) continue;
        var colName = attrName.replace(/^ows_/, "").replace(/^ows:/, "");
        if (colName) {
          row[colName] = attr.value;
          if (colName === "Id" && !row.ID) row.ID = attr.value;
          if (colName === "ID" && !row.Id) row.Id = attr.value;
          if (columns.indexOf(colName) < 0) columns.push(colName);
        }
      }
      if (Object.keys(row).length > 0) rows.push(row);
    }

    columns.sort(function (a, b) {
      if (a === "ID") return -1;
      if (b === "ID") return 1;
      return a.localeCompare(b);
    });

    return { rows: rows, columns: columns };
  }

  var SIZE_COLUMNS = { "FileSizeDisplay": 1, "File_x0020_Size": 1, "SMTotalSize": 1, "FileSize": 1 };

  function sanitizeColumnDisplayName(colName) {
    var map = {
      "ID": "ID",
      "Id": "ID",
      "FileSizeDisplay": "File Size (MB)",
      "File_x0020_Size": "File Size (MB)",
      "SMTotalSize": "File Size (MB)",
      "FileSize": "File Size (MB)",
      "FileLeafRef": "Name",
      "FSObjType": "Object Type",
      "ItemChildCount": "Item Child Count",
      "DocIcon": "Doc Icon",
      "_UIVersionString": "Version"
    };
    if (map[colName]) return map[colName];
    var s = colName.replace(/_x0020_/g, " ").replace(/_x002e_/g, ".").replace(/^_/, "");
    return s.replace(/([a-z])([A-Z])/g, "$1 $2").replace(/\s+/g, " ").trim();
  }

  function formatSizeAsMB(val) {
    if (val == null || val === "") return "";
    var s = String(val).trim();
    var bytes = parseInt(s, 10);
    if (!isNaN(bytes) && bytes >= 0) return (bytes / (1024 * 1024)).toFixed(2);
    var n = parseFloat(s.replace(/[^\d.]/g, ""));
    if (!isNaN(n) && n >= 0) return n.toFixed(2);
    return s;
  }

  function getCellValue(row, colName) {
    var v = row[colName];
    if (v != null && String(v).trim() !== "") {
      if (SIZE_COLUMNS[colName]) return formatSizeAsMB(v);
      return cleanValue(v);
    }
    if (colName === "ID" || colName === "Id") return cleanValue(row.ID || row.Id || "");
    if (colName === "Title") return cleanValue(row.Title || row.title || "");
    if (colName === "Name" || colName === "FileLeafRef") return cleanValue(row.Name || row.FileLeafRef || row.LinkFilename || "");
    if (colName === "FSObjType") return row.FSObjType != null ? cleanValue(row.FSObjType) : "";
    return v != null ? cleanValue(v) : "";
  }

  function toCSV(rowMapById, columns) {
    var header = columns.map(function (c) { return escapeCSV(sanitizeColumnDisplayName(c)); }).join(",");
    var ids = Object.keys(rowMapById).sort(function (a, b) { return parseInt(a, 10) - parseInt(b, 10); });
    var body = ids.map(function (id) {
      var r = rowMapById[id] || {};
      if (!r.ID && !r.Id) r.ID = id;
      return columns.map(function (c) { return escapeCSV(getCellValue(r, c)); }).join(",");
    });
    return header + "\r\n" + body.join("\r\n");
  }

  function toCSVFromRows(rows, columns) {
    var header = columns.map(function (c) { return escapeCSV(sanitizeColumnDisplayName(c)); }).join(",");
    var sorted = rows.slice().sort(function (a, b) {
      var idA = parseInt(a.ID || a.Id || "0", 10);
      var idB = parseInt(b.ID || b.Id || "0", 10);
      if (idA !== idB) return idA - idB;
      var vA = a._UIVersionString || a.Version || "";
      var vB = b._UIVersionString || b.Version || "";
      return String(vA).localeCompare(String(vB));
    });
    var body = sorted.map(function (r) {
      return columns.map(function (c) { return escapeCSV(getCellValue(r, c)); }).join(",");
    });
    return header + "\r\n" + body.join("\r\n");
  }

  function getDigestFromDomOrContext() {
    var el = document.getElementById("__REQUESTDIGEST");
    if (el && el.value) return el.value;
    if (window._spPageContextInfo && window._spPageContextInfo.formDigestValue) return window._spPageContextInfo.formDigestValue;
    return null;
  }

  // -----------------------------
  // Inputs / Context
  // -----------------------------
  var siteUrl = String(params.u || "").replace(/\/$/, "");
  var listId = normalizeGuid(params.lid || "");

  if (typeof window._spPageContextInfo === "object" && window._spPageContextInfo) {
    var spi = window._spPageContextInfo;
    if (spi.webAbsoluteUrl) siteUrl = String(spi.webAbsoluteUrl).replace(/\/$/, "");
    if (spi.pageListId) listId = normalizeGuid(spi.pageListId);
  }

  var PAGE_LIMIT = (params.pl != null && params.pl > 0) ? parseInt(params.pl, 10) : 1000;

  // List type: 101 = document library; set before export so owssvr URL can vary for lists
  var listBaseTemplate = null;

  // Working view name (used for owssvr download - all columns + show all items without folders)
  var RPC_VIEW_TITLE = String(params.vt || "RPC");

  // Delete working view at end? (default true unless keep=1)
  var DELETE_VIEW_AT_END = !(params.keep === 1 || params.keep === "1" || params.keep === true || params.keep === "true");

  function getSiteNameFromUrl(url) {
    try {
      var u = new URL(url);
      var segs = u.pathname.replace(/\/$/, "").split("/").filter(Boolean);
      return segs.length > 0 ? segs[segs.length - 1] : "Site";
    } catch (_) { return "Site"; }
  }
  function sanitizeFilenamePart(s) {
    return String(s).replace(/[\s\\/:*?"<>|]/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "") || "Export";
  }
  function buildDefaultExportFilename() {
    var siteName = sanitizeFilenamePart(getSiteNameFromUrl(siteUrl));
    var listTitle = getListTitleFromContext();
    var listPart = (listTitle && String(listTitle).trim()) ? sanitizeFilenamePart(String(listTitle).trim()) : "";
    var d = new Date();
    var pad = function (n) { n = parseInt(n, 10); return (n >= 0 && n <= 9) ? "0" + n : String(n); };
    var dt = d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) + "_" + pad(d.getHours()) + "-" + pad(d.getMinutes());
    return listPart ? siteName + "_" + listPart + "_" + dt : siteName + "_" + dt;
  }
  var exportFilename = (params.f && String(params.f).trim())
    ? String(params.f).trim().replace(/[\s\\/:*?"<>|]/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "") || buildDefaultExportFilename()
    : buildDefaultExportFilename();

  // -----------------------------
  // REST helpers
  // -----------------------------
  async function getDigest(forceRefresh) {
    if (!forceRefresh) {
      var d = getDigestFromDomOrContext();
      if (d) return d;
    }
    var r = await fetch(siteUrl + "/_api/contextinfo", {
      method: "POST",
      credentials: "include",
      headers: { "Accept": "application/json;odata=nometadata" }
    });
    if (!r.ok) return null;
    var j = await r.json();
    return j.FormDigestValue || null;
  }

  function listBaseUrl() {
    return siteUrl + "/_api/web/lists(guid'" + listId + "')";
  }

  async function getItemCount() {
    try {
      var r = await fetch(listBaseUrl() + "?$select=ItemCount", {
        credentials: "include",
        headers: { "Accept": "application/json;odata=nometadata" }
      });
      if (!r.ok) return null;
      var j = await r.json();
      var n = parseInt(j.ItemCount, 10);
      return isNaN(n) || n < 0 ? null : n;
    } catch (_) { return null; }
  }

  async function getListBaseTemplate() {
    try {
      var r = await fetch(listBaseUrl() + "?$select=BaseTemplate", {
        credentials: "include",
        headers: { "Accept": "application/json;odata=nometadata" }
      });
      if (!r.ok) return null;
      var j = await r.json();
      var t = j.BaseTemplate;
      return t != null ? parseInt(t, 10) : null;
    } catch (_) { return null; }
  }

  var BATCH_SIZE_VERSIONS = 100;
  var VERSIONS_PAGE_SIZE = 5000;

  /** Fetches all versions for one item, following @odata.nextLink. Returns array of version objects or null on failure. */
  async function fetchAllVersionsForItem(itemId, selectQuery) {
    var lb = listBaseUrl();
    var url = lb + "/items(" + itemId + ")/versions?$select=" + encodeURIComponent(selectQuery) + "&$top=" + VERSIONS_PAGE_SIZE;
    var all = [];
    var accept = { "Accept": "application/json;odata=nometadata" };
    while (url) {
      var r = await fetch(url, { credentials: "include", headers: accept });
      if (!r.ok) return null;
      var j = await r.json();
      var page = j.value || j.d?.results || [];
      if (!Array.isArray(page)) return null;
      for (var i = 0; i < page.length; i++) all.push(page[i]);
      url = j["@odata.nextLink"] || null;
      if (page.length < VERSIONS_PAGE_SIZE) break;
    }
    return all;
  }

  function buildBatchBody(itemIds, selectQuery) {
    var boundary = "batch_" + Math.random().toString(36).slice(2) + "_" + Date.now();
    var lb = listBaseUrl();
    var lines = [];
    for (var i = 0; i < itemIds.length; i++) {
      var itemId = itemIds[i];
      var url = lb + "/items(" + itemId + ")/versions?$select=" + encodeURIComponent(selectQuery);
      lines.push("--" + boundary);
      lines.push("Content-Type: application/http");
      lines.push("Content-Transfer-Encoding: binary");
      lines.push("");
      lines.push("GET " + url + " HTTP/1.1");
      lines.push("Accept: application/json;odata=nometadata");
      lines.push("");
    }
    lines.push("--" + boundary + "--");
    return { body: lines.join("\r\n"), boundary: boundary };
  }

  function parseBatchResponse(responseText, requestBoundary, responseContentType) {
    var boundary = requestBoundary;
    if (responseContentType) {
      var m = /boundary\s*=\s*["']?([^"'\s;]+)/i.exec(responseContentType);
      if (m) boundary = m[1].trim();
    }
    var sep = "\r\n--" + boundary;
    var parts = responseText.split(sep);
    var results = [];
    for (var p = 0; p < parts.length; p++) {
      var part = parts[p].trim();
      if (!part || part.indexOf("HTTP/1.1") < 0) continue;
      var dbl = part.indexOf("\r\n\r\n");
      if (dbl < 0) dbl = part.indexOf("\n\n");
      if (dbl < 0) continue;
      var rest = part.slice(dbl + (part.indexOf("\r\n\r\n") >= 0 ? 4 : 2));
      var secondDbl = rest.indexOf("\r\n\r\n");
      if (secondDbl < 0) secondDbl = rest.indexOf("\n\n");
      var body = (secondDbl >= 0 ? rest.slice(secondDbl + (rest.indexOf("\r\n\r\n") >= 0 ? 4 : 2)) : rest).trim();
      body = body.replace(/\r\n--.*$/s, "").trim();
      if (!body) continue;
      try {
        var j = JSON.parse(body);
        results.push(j);
      } catch (_) {}
    }
    return results;
  }

  async function getMinListItemId() {
    try {
      var r = await fetch(listBaseUrl() + "/items?$select=Id&$orderby=Id asc&$top=1", {
        credentials: "include",
        headers: { "Accept": "application/json;odata=nometadata" }
      });
      if (!r.ok) return null;
      var j = await r.json();
      var items = j.value || j.d?.results || [];
      if (items.length === 0) return null;
      var id = items[0].ID || items[0].Id;
      var n = parseInt(id, 10);
      return isNaN(n) || n < 1 ? null : n;
    } catch (_) { return null; }
  }

  async function getMaxListItemId() {
    try {
      var r = await fetch(listBaseUrl() + "/items?$select=Id&$orderby=Id desc&$top=1", {
        credentials: "include",
        headers: { "Accept": "application/json;odata=nometadata" }
      });
      if (!r.ok) return null;
      var j = await r.json();
      var items = j.value || j.d?.results || [];
      if (items.length === 0) return null;
      var id = items[0].ID || items[0].Id;
      var n = parseInt(id, 10);
      return isNaN(n) || n < 1 ? null : n;
    } catch (_) { return null; }
  }

  // Wait for page/context to be ready (avoids 403 when script runs before SharePoint is ready).
  async function waitForPageReady(maxMs) {
    maxMs = maxMs || 4000;
    var deadline = Date.now() + maxMs;
    while (Date.now() < deadline) {
      if (document.readyState === "complete" && (typeof window._spPageContextInfo === "object" && window._spPageContextInfo)) {
        await sleep(400);
        return;
      }
      await sleep(300);
    }
  }

  // Creates a temporary RPC view with ALL columns and "Show all items without folders" (Scope 1).
  // Always creates fresh: deletes existing "RPC" view if present, then creates new one.
  async function getOrCreateRpcViewForDownload() {
    var accept = { "Accept": "application/json;odata=nometadata" };
    var lb = listBaseUrl();

    // Fetch all views and find by Title (more reliable than $filter which may not work on views)
    var viewsResp = await fetch(lb + "/views?$select=Id,Title", { credentials: "include", headers: accept });
    if (!viewsResp.ok) {
      var t0 = "";
      try { t0 = await viewsResp.text(); } catch (_) {}
      return { ok: false, error: "List views failed: " + viewsResp.status, detail: t0 };
    }

    var viewsJson = await viewsResp.json();
    var views = viewsJson.value || viewsJson.d?.results || [];
    var rpcView = null;
    for (var i = 0; i < views.length; i++) {
      if ((views[i].Title || "").trim() === RPC_VIEW_TITLE) {
        rpcView = views[i];
        break;
      }
    }

    // Delete existing RPC view so we always create fresh with correct Scope/fields
    if (rpcView) {
      await deleteViewById(normalizeGuid(rpcView.Id));
      await sleep(200);
    }

    // Create new view with Scope 2 (RecursiveAll). Retry on 403/503 (often timing/session not ready on first load).
    var createResp = null;
    var lastStatus = 0;
    var lastDetail = "";
    var maxAttempts = 5;
    var retryDelayMs = 1500;
    for (var attempt = 1; attempt <= maxAttempts; attempt++) {
      var digest = await getDigest(attempt > 1);
      if (!digest) return { ok: false, error: "Could not get request digest" };

      createResp = await fetch(lb + "/views", {
        method: "POST",
        credentials: "include",
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": digest
        },
        body: JSON.stringify({
          Title: RPC_VIEW_TITLE,
          PersonalView: false,
          RowLimit: Math.min(PAGE_LIMIT, 5000000),
          Scope: 2,  // RecursiveAll = Show all items (files + folders)
          ViewQuery: ""
        })
      });

      lastStatus = createResp.status;
      if (createResp.ok) break;
      try { lastDetail = await createResp.text(); } catch (_) {}
      if (createResp.status !== 403 && createResp.status !== 503) break;
      await sleep(attempt * retryDelayMs);
    }

    if (!createResp || !createResp.ok) {
      return { ok: false, error: "Create RPC view failed: " + lastStatus, detail: lastDetail };
    }

    var created = await createResp.json();
    var viewId = normalizeGuid(created.Id);

    // Use custom columns from picker, or default view fields
    var fieldNames;
    if (params.cols && Array.isArray(params.cols) && params.cols.length > 0) {
      fieldNames = mergeRequiredFields(params.cols);
      var valid = await getListFieldInternalNames();
      if (valid.ok && valid.names) {
        var filtered = [];
        for (var fi = 0; fi < fieldNames.length; fi++) {
          if (valid.names[fieldNames[fi]]) filtered.push(fieldNames[fi]);
        }
        if (filtered.length > 0) fieldNames = filtered;
      }
    } else {
      var vf = await getViewFieldsFromDefaultView();
      if (!vf.ok) return vf;
      fieldNames = vf.fields;
    }
    var reportNeedsPathCols =
      params.report === "folderCount" ||
      params.report === "pathLengths" ||
      params.report === "exportCSV";
    if (reportNeedsPathCols) {
      fieldNames = fieldNames.slice();
      var vrPath = await getListFieldInternalNames();
      var vnPath = (vrPath.ok && vrPath.names) ? vrPath.names : {};
      if (fieldNames.indexOf("FileRef") < 0 && vnPath["FileRef"]) fieldNames.push("FileRef");
      if ((params.report === "folderCount" || params.report === "pathLengths") && fieldNames.indexOf("ContentTypeId") < 0 && vnPath["ContentTypeId"]) {
        fieldNames.push("ContentTypeId");
      }
      if (fieldNames.indexOf("FileDirRef") < 0 && vnPath["FileDirRef"]) fieldNames.push("FileDirRef");
    }
    if (!fieldNames || fieldNames.length === 0) {
      return { ok: false, error: "No columns selected" };
    }

    var sf = await setViewFieldsBatch(viewId, fieldNames);
    if (!sf.ok) return { ok: false, error: "Failed setting ViewFields: " + (sf.error || "unknown") };
    await sleep(300);

    return { ok: true, viewId: viewId, fieldNames: fieldNames };
  }

  function mergeRequiredFields(fieldNames) {
    var required = ["ID", "Title", "FileLeafRef", "FSObjType", "Created", "Modified", "Author", "Editor", "_UIVersionString"];
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
    var out = [];
    for (var i = 0; i < required.length; i++) {
      var k = normKey(required[i]);
      if (!seen[k]) { seen[k] = true; out.push(required[i]); }
    }
    for (var j = 0; j < fieldNames.length; j++) {
      var n = String(fieldNames[j] || "").trim();
      if (n) {
        var k = normKey(n);
        if (!seen[k]) { seen[k] = true; out.push(canonicalName(n)); }
      }
    }
    return out;
  }

  // Get ViewFields from the list's default view (e.g. "All Items"). These produce valid owssvr XML.
  async function getViewFieldsFromDefaultView() {
    var accept = { "Accept": "application/json;odata=nometadata" };
    var lb = listBaseUrl();

    var r = await fetch(lb + "/DefaultView/ViewFields", { credentials: "include", headers: accept });
    if (!r.ok) {
      var t = "";
      try { t = await r.text(); } catch (_) {}
      return { ok: false, error: "Get default ViewFields failed: " + r.status, detail: t };
    }

    var j = await r.json();
    var raw = j.Items || j.value || j.d?.results || [];
    var seen = {};
    var fields = [];
    var exclude = {
      "VirusStatus": 1, "VirusVendorID": 1, "VirusInfo": 1, "VirusInfoEx": 1,
      "SMTotalSize": 1, "SMLastModifiedDate": 1, "InstanceID": 1, "SyncClientId": 1,
      "ProgId": 1, "ScopeId": 1, "PermMask": 1, "UniqueId": 1, "WorkflowVersion": 1,
      "WorkflowInstanceID": 1, "FormData": 1, "RestrictType": 1, "NoCrawl": 1,
      "IsCurrentVersion": 1
    };
    function excludeByName(n) {
      if (exclude[n]) return true;
      if (n.indexOf("_") === 0 && n !== "_UIVersionString") return true;
      if (n.indexOf("vti_") === 0 || n.indexOf("ows_") === 0 || n.indexOf("tp_") === 0) return true;
      return false;
    }
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
    for (var i = 0; i < raw.length; i++) {
      var item = raw[i];
      var name = (typeof item === "string") ? item : (item && (item.Name || item.InternalName || item.Title || item));
      name = String(name || "").trim();
      if (name && !excludeByName(name)) {
        var k = normKey(name);
        if (!seen[k]) {
          seen[k] = true;
          fields.push(canonicalName(name));
        }
      }
    }
    if (!seen["ID"]) { seen["ID"] = true; fields.unshift("ID"); }

    // Always include: Title, FileLeafRef (name), Version, Created, Modified, Created By, Modified By
    var required = ["ID", "Title", "FileLeafRef", "FSObjType", "Created", "Modified", "Author", "Editor", "_UIVersionString"];
    for (var ri = 0; ri < required.length; ri++) {
      if (!seen[normKey(required[ri])]) {
        seen[normKey(required[ri])] = true;
        fields.push(required[ri]);
      }
    }

    return { ok: true, fields: fields };
  }

  async function deleteViewById(viewId) {
    var digest = await getDigest();
    if (!digest) return { ok: false, error: "Could not get request digest" };

    var url = listBaseUrl() + "/views(guid'" + normalizeGuid(viewId) + "')";
    var resp = await fetch(url, {
      method: "POST",
      credentials: "include",
      headers: {
        "Accept": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE"
      }
    });

    var err = "";
    if (!resp.ok) { try { err = await resp.text(); } catch (_) {} }
    return { ok: resp.ok, status: resp.status, error: err };
  }

  async function setViewQueryById(viewId, viewQuery) {
    var vId = normalizeGuid(viewId);
    var apiUrl = listBaseUrl() + "/views(guid'" + vId + "')";
    var lastStatus = 0;
    var lastError = "";
    var maxAttempts = 3;
    var retryDelayMs = 1200;

    for (var attempt = 1; attempt <= maxAttempts; attempt++) {
      var digest = await getDigest(true);
      if (!digest) return { ok: false, error: "Could not get request digest" };

      var resp = await fetch(apiUrl, {
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
      });

      lastStatus = resp.status;
      lastError = "";
      if (!resp.ok) { try { lastError = await resp.text(); } catch (_) {} }

      if (resp.ok) return { ok: true, status: resp.status, error: "" };
      var isSecurityValidation = lastStatus === 403 ||
        (lastError && (lastError.indexOf("security validation") !== -1 || lastError.indexOf("2130575252") !== -1 || lastError.indexOf("invalid and might be corrupted") !== -1));
      if (!isSecurityValidation || attempt >= maxAttempts) return { ok: false, status: lastStatus, error: lastError };
      await sleep(attempt * retryDelayMs);
    }
    return { ok: false, status: lastStatus, error: lastError };
  }

  async function removeAllViewFields(viewId) {
    var accept = { "Accept": "application/json;odata=nometadata" };
    var vfBase = listBaseUrl() + "/views(guid'" + normalizeGuid(viewId) + "')/ViewFields";
    var maxAttempts = 3;
    var retryDelayMs = 1200;

    for (var attempt = 1; attempt <= maxAttempts; attempt++) {
      var digest = await getDigest(true);
      if (!digest) return { ok: false, error: "Could not get request digest" };

      var ra = await fetch(vfBase + "/RemoveAllViewFields", {
        method: "POST",
        credentials: "include",
        headers: { ...accept, "X-RequestDigest": digest }
      });
      if (ra.ok) return { ok: true };
      var errText = "";
      try { errText = await ra.text(); } catch (_) {}
      var isSecurityValidation = ra.status === 403 ||
        (errText && (errText.indexOf("security validation") !== -1 || errText.indexOf("2130575252") !== -1 || errText.indexOf("invalid and might be corrupted") !== -1));
      if (!isSecurityValidation || attempt >= maxAttempts) {
        if (ra.status === 404 || ra.status === 501) break;
        return { ok: false, error: "RemoveAllViewFields failed: " + ra.status, detail: errText };
      }
      await sleep(attempt * retryDelayMs);
    }

    var cur = await fetch(vfBase, { credentials: "include", headers: accept });
    if (!cur.ok) {
      var t = "";
      try { t = await cur.text(); } catch (_) {}
      return { ok: false, error: "Failed reading ViewFields: " + cur.status, detail: t };
    }
    var j = await cur.json();
    var items = j.Items || j.value || [];
    for (var i = 0; i < items.length; i++) {
      var f = String(items[i]).replace(/'/g, "''");
      for (var attempt = 1; attempt <= 2; attempt++) {
        var digest = await getDigest(true);
        if (!digest) return { ok: false, error: "Could not get request digest" };
        var r = await fetch(vfBase + "/removeviewfield('" + f + "')", {
          method: "POST",
          credentials: "include",
          headers: { ...accept, "X-RequestDigest": digest }
        });
        if (r.ok) break;
        if (r.status !== 403 && r.status !== 503) break;
        await sleep(attempt * 800);
      }
    }
    return { ok: true };
  }

  async function addViewField(viewId, internalName) {
    var accept = { "Accept": "application/json;odata=nometadata" };
    var vfBase = listBaseUrl() + "/views(guid'" + normalizeGuid(viewId) + "')/ViewFields";
    var name = String(internalName).replace(/'/g, "''");
    var maxAttempts = 3;
    var retryDelayMs = 1200;
    var lastStatus = 0;
    var lastError = "";

    for (var attempt = 1; attempt <= maxAttempts; attempt++) {
      var digest = await getDigest(true);
      if (!digest) return { ok: false, status: 0, error: "Could not get request digest" };

      var r = await fetch(vfBase + "/addviewfield('" + name + "')", {
        method: "POST",
        credentials: "include",
        headers: { ...accept, "X-RequestDigest": digest }
      });
      lastStatus = r.status;
      lastError = "";
      if (!r.ok) { try { lastError = await r.text(); } catch (_) {} }
      if (r.ok) return { ok: true, status: r.status, error: "" };
      var isSecurityValidation = lastStatus === 403 ||
        (lastError && (lastError.indexOf("security validation") !== -1 || lastError.indexOf("2130575252") !== -1 || lastError.indexOf("invalid and might be corrupted") !== -1));
      if (!isSecurityValidation || attempt >= maxAttempts) return { ok: false, status: lastStatus, error: lastError };
      await sleep(attempt * retryDelayMs);
    }
    return { ok: false, status: lastStatus, error: lastError };
  }

  async function setViewFieldsBatch(viewId, fieldNames) {
    var rm = await removeAllViewFields(viewId);
    if (!rm.ok) return rm;

    // Allow removeAllViewFields to propagate (new views may have default fields)
    await sleep(300);

    // Deduplicate: owssvr XML fails with "Attribute ows_X redefined" if same field appears twice.
    // Id/ID, LinkFilename/FileLeafRef/Name, LinkTitle/Title produce same ows_ attributes - only add one.
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

    var failed = 0;
    for (var i = 0; i < names.length; i++) {
      var r = await addViewField(viewId, names[i]);
      if (!r.ok) failed++;
    }
    if (failed > names.length / 2) {
      return { ok: false, error: "View field setup failed (403 or permission denied). Try refreshing the page and run the export again." };
    }
    return { ok: true };
  }

  async function getListFieldInternalNames() {
    var accept = { "Accept": "application/json;odata=nometadata" };
    var r = await fetch(
      listBaseUrl() + "/fields?$select=InternalName,TypeAsString&$top=5000",
      { credentials: "include", headers: accept }
    );
    if (!r.ok) return { ok: false, error: "Get fields failed: " + r.status };
    var j = await r.json();
    var fields = j.value || j.d?.results || [];
    var names = {};
    var lookupSet = {};
    for (var i = 0; i < fields.length; i++) {
      var f = fields[i];
      var n = f && f.InternalName;
      if (n) {
        names[n] = true;
        var t = (f.TypeAsString || f.Type || "").toLowerCase();
        if (t.indexOf("lookup") >= 0 || t === "user" || t.indexOf("user") >= 0) lookupSet[n] = true;
      }
    }
    return { ok: true, names: names, lookupSet: lookupSet };
  }

  function filterLookupFieldsForRest(fieldNames, lookupSet) {
    if (!lookupSet) return fieldNames;
    return fieldNames.filter(function (n) { return !lookupSet[n]; });
  }

  function flattenRestValue(val) {
    if (val == null || val === undefined) return "";
    if (typeof val === "object") {
      if (val.LookupValue != null) return String(val.LookupValue);
      if (Array.isArray(val)) {
        return val.map(function (v) {
          return (v && v.LookupValue != null) ? String(v.LookupValue) : String(v);
        }).join(";#");
      }
      return JSON.stringify(val);
    }
    return String(val);
  }

  function restItemToRow(item) {
    var row = {};
    for (var k in item) {
      if (!Object.prototype.hasOwnProperty.call(item, k)) continue;
      if (k === "odata.type" || k === "odata.id" || k === "odata.etag" || k.indexOf("odata.") === 0) continue;
      if (k === "ID" || k === "Id") {
        row.ID = String(item[k]);
      } else {
        row[k] = flattenRestValue(item[k]);
      }
    }
    return row;
  }

  async function getEligibleFieldNames() {
    var accept = { "Accept": "application/json;odata=nometadata" };
    var r = await fetch(
      listBaseUrl() + "/fields?$select=InternalName,Hidden,ReadOnlyField,Sealed,ShowInListDefault,TypeAsString",
      { credentials: "include", headers: accept }
    );
    if (!r.ok) return { ok: false, error: "Get fields failed: " + r.status };
    var j = await r.json();
    var fields = j.value || j.d?.results || [];
    var skip = {
      "Attachments": 1, "ContentTypeId": 1, "ComplianceAssetId": 1, "MetaInfo": 1, "Edit": 1,
      "AppAuthor": 1, "AppEditor": 1, "MediaServiceImageTags": 1, "MediaServiceAutoTags": 1,
      "USPSGATS": 1, "VirusStatus": 1, "VirusVendorID": 1, "VirusInfo": 1, "VirusInfoEx": 1,
      "SMTotalSize": 1, "SMLastModifiedDate": 1, "InstanceID": 1, "SyncClientId": 1,
      "ProgId": 1, "ScopeId": 1, "PermMask": 1, "UniqueId": 1, "WorkflowVersion": 1,
      "WorkflowInstanceID": 1, "FormData": 1, "RestrictType": 1, "NoCrawl": 1, "IsCurrentVersion": 1
    };
    var skipPrefix = ["LinkTitle", "LinkTitleNoMenu", "LinkFilename", "LinkFilenameNoMenu", "vti_", "ows_", "tp_"];
    function isSkipName(n) {
      if (skip[n]) return true;
      if (n.indexOf("_") === 0 && n !== "_UIVersionString") return true;
      if (n.indexOf("c000_") === 0) return true;
      for (var i = 0; i < skipPrefix.length; i++) { if (n.indexOf(skipPrefix[i]) === 0) return true; }
      return false;
    }
    function isLookup(f) {
      var t = (f.TypeAsString || f.Type || "").toLowerCase();
      return t.indexOf("lookup") >= 0 || t === "user" || t.indexOf("user") >= 0;
    }
    var seen = {};
    var eligible = [];
    for (var i = 0; i < fields.length; i++) {
      var f = fields[i];
      var name = f.InternalName;
      if (!name || seen[name]) continue;
      if (f.Hidden || f.ReadOnlyField || f.Sealed) continue;
      if (f.ShowInListDefault === false) continue;
      if (isSkipName(name)) continue;
      if (isLookup(f)) continue;
      seen[name] = true;
      eligible.push(name);
    }
    eligible = eligible.filter(function (n) { return n !== "ID"; });
    return { ok: true, fields: eligible };
  }

  function buildOwssvrUrl(viewId) {
    var guid = function (g) { return "{" + normalizeGuid(g).toUpperCase() + "}"; };
    var incVer = params.incVer !== false;
    var isDocLib = listBaseTemplate === 101;
    // For lists (not doc lib), version expansion often needs RootFolder=* and RowLimit=0
    var rowLimit = (incVer && !isDocLib) ? 0 : Math.max(PAGE_LIMIT * 20, 50000);
    var url = siteUrl + "/_vti_bin/owssvr.dll?Cmd=Display&XMLDATA=1&List=" + encodeURIComponent(guid(listId)) +
      "&View=" + encodeURIComponent(guid(viewId)) + "&IncludeVersions=" + (incVer ? "TRUE" : "FALSE") + "&RowLimit=" + rowLimit;
    if (incVer && !isDocLib) url += "&RootFolder=*";
    if (isDocLib) {
      // Use list root path so RecursiveAll view returns all items in all folders; RootFolder=* can limit to one folder with ID-range paging
      var listRoot = (window._spPageContextInfo && (window._spPageContextInfo.listUrl || window._spPageContextInfo.listServerRelativeUrl)) || "";
      if (listRoot) url += "&RootFolder=" + encodeURIComponent(listRoot);
      else url += "&RootFolder=*";
    }
    return url;
  }

  function colNormKey(c) {
    if (c === "Id") return "ID";
    if (c === "Name") return "FileLeafRef";
    return c;
  }

  function getRowVal(row, keys) {
    var lowerKeys = keys.map(function (k) { return k.toLowerCase(); });
    for (var p in row) {
      if (!Object.prototype.hasOwnProperty.call(row, p)) continue;
      if (lowerKeys.indexOf(p.toLowerCase()) >= 0) {
        var v = row[p];
        if (v != null && String(v).trim() !== "") return String(v).trim();
      }
    }
    return "";
  }

  function stripIdHashPrefix(s) {
    if (!s || typeof s !== "string") return s || "";
    var idx = s.indexOf("#");
    return idx >= 0 ? s.slice(idx + 1).trim() : s;
  }

  function isDocumentSet(row) {
    var ct = getRowVal(row, ["ContentTypeId", "ContentTypeID", "ContentType"]);
    ct = stripIdHashPrefix(ct);
    return ct.toLowerCase().indexOf("0x0120d520") === 0;
  }

  function getItemPathAndType(row, knownFolderPaths) {
    var path = getRowVal(row, ["FileRef", "ServerRelativeUrl", "FileRefUrl"]);
    path = stripIdHashPrefix(path);
    if (!path) {
      var dirRef = stripIdHashPrefix(getRowVal(row, ["FileDirRef"]));
      var leaf = getRowVal(row, ["FileLeafRef", "Name", "LinkFilename", "BaseName"]);
      path = (dirRef + "/" + leaf).replace(/\/+/g, "/").replace(/^\/+/, "/");
    }
    if (!path || path === "/") return null;
    path = path.replace(/\/+/g, "/");
    if (path.length > 1 && path.charAt(path.length - 1) === "/") path = path.slice(0, -1);
    if (!path) path = "/";
    var fsVal = getRowVal(row, ["FSObjType", "FileSystemObjectType"]);
    var fsoNum = parseInt(fsVal, 10);
    var fsFolder = fsVal === "1" || String(fsVal).toLowerCase() === "folder" || fsoNum === 1;
    var isFolder = !!fsFolder || isDocumentSet(row) || (knownFolderPaths && knownFolderPaths[path]);
    var parentPath;
    if (isFolder) {
      parentPath = path.replace(/\/[^/]+$/, "") || "/";
    } else {
      var dirRef = stripIdHashPrefix(getRowVal(row, ["FileDirRef"]));
      dirRef = dirRef.replace(/\/+/g, "/");
      if (dirRef.length > 1 && dirRef.charAt(dirRef.length - 1) === "/") dirRef = dirRef.slice(0, -1);
      parentPath = dirRef || path.replace(/\/[^/]+$/, "") || "/";
    }
    return { path: path, isFolder: !!isFolder, parentPath: parentPath || "/" };
  }

  function buildFolderCountReportFromRows(rows) {
    var knownFolderPaths = {};
    for (var i = 0; i < rows.length; i++) {
      var dirRef = stripIdHashPrefix(getRowVal(rows[i], ["FileDirRef"]));
      if (dirRef) {
        dirRef = dirRef.replace(/\/+/g, "/");
        if (dirRef.length > 1 && dirRef.charAt(dirRef.length - 1) === "/") dirRef = dirRef.slice(0, -1);
        if (dirRef) knownFolderPaths[dirRef] = true;
      }
    }
    var items = [];
    for (var i = 0; i < rows.length; i++) {
      var x = getItemPathAndType(rows[i], knownFolderPaths);
      if (x) items.push(x);
    }
    if (items.length === 0) return "FolderPath,FolderName,FolderLevel,DirectFoldersCount,DirectFilesCount,TotalNestedFolders,LevelsDeep,TotalItemsRecursive,ServerRelativeUrl\r\n";
    var pathToFolder = {};
    var root = null;
    for (var j = 0; j < items.length; j++) {
      var it0 = items[j];
      var candidate = it0.isFolder ? it0.path : it0.path.replace(/\/[^/]+$/, "").replace(/\/$/, "") || "/";
      if (!root || candidate.length < root.length) root = candidate;
    }
    root = (root || "/").replace(/\/$/, "") || "/";
    for (var k = 0; k < items.length; k++) {
      var it = items[k];
      var parent = it.parentPath || "/";
      if (parent && !pathToFolder[parent]) pathToFolder[parent] = { directFolders: 0, directFiles: 0, children: [] };
      if (it.isFolder) {
        if (!pathToFolder[it.path]) pathToFolder[it.path] = { directFolders: 0, directFiles: 0, children: [] };
        if (pathToFolder[parent]) {
          pathToFolder[parent].directFolders++;
          if (pathToFolder[parent].children.indexOf(it.path) < 0) pathToFolder[parent].children.push(it.path);
        }
      } else {
        if (pathToFolder[parent]) pathToFolder[parent].directFiles++;
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
    // Library root (e.g. sites/site/LibraryName) = level 0; don't count path above library in folder level or depth
    var normRoot = (displayRoot || "/").replace(/\/+$/, "") || "/";
    function folderLevelFromPath(p) {
      if (!p || p === "/") return 0;
      var norm = p.replace(/\/+$/, "").replace(/^\/+/, "/");
      if (norm === normRoot || norm.length <= normRoot.length) return 0;
      if (normRoot !== "/" && norm.indexOf(normRoot + "/") !== 0) return 0;
      var rel = normRoot === "/" ? norm.replace(/^\/+/, "") : norm.slice(normRoot.length).replace(/^\/+/, "");
      return rel ? rel.split("/").filter(function (s) { return s.length > 0; }).length : 0;
    }
    var lines = ["FolderPath,FolderName,FolderLevel,DirectFoldersCount,DirectFilesCount,TotalNestedFolders,LevelsDeep,TotalItemsRecursive,ServerRelativeUrl"];
    for (var f = 0; f < folderPaths.length; f++) {
      var path = folderPaths[f];
      var info = pathToFolder[path];
      var folderPathDisplay = (path === displayRoot || path === displayRoot + "/") ? "/" : (path.indexOf(displayRoot) === 0 ? "/" + path.slice(displayRoot.length).replace(/^\//, "") : path);
      var folderName = path.replace(/.*\//, "") || path || "/";
      var row = [
        escapeCSV(folderPathDisplay),
        escapeCSV(folderName),
        String(folderLevelFromPath(path)),
        String(info.directFolders),
        String(info.directFiles),
        String(totalNestedFolders[path] != null ? totalNestedFolders[path] : 0),
        String(levelsDeep[path] != null ? levelsDeep[path] : 0),
        String(totalRecursive[path] != null ? totalRecursive[path] : 0),
        escapeCSV(path)
      ];
      lines.push(row.join(","));
    }
    return "\ufeff" + lines.join("\r\n");
  }

  function buildPathLengthsReportFromRows(rows) {
    var knownFolderPaths = {};
    for (var i = 0; i < rows.length; i++) {
      var dirRef = stripIdHashPrefix(getRowVal(rows[i], ["FileDirRef"]));
      if (dirRef) {
        dirRef = dirRef.replace(/\/+$/, "").replace(/^\/+/, "/");
        if (dirRef.length > 1 && dirRef.charAt(dirRef.length - 1) === "/") dirRef = dirRef.slice(0, -1);
        if (dirRef) knownFolderPaths[dirRef] = true;
      }
    }
    var shortestPath = null;
    for (var i = 0; i < rows.length; i++) {
      var x = getItemPathAndType(rows[i], knownFolderPaths);
      if (!x || x.isFolder) continue;
      if (!shortestPath || x.path.length < shortestPath.length) shortestPath = x.path;
    }
    var normRoot = (shortestPath ? shortestPath.replace(/\/[^/]+$/, "").replace(/\/+$/, "") : "") || "/";
    var fileEntries = [];
    for (var j = 0; j < rows.length; j++) {
      var y = getItemPathAndType(rows[j], knownFolderPaths);
      if (!y || y.isFolder) continue;
      var path = y.path;
      var pathAfterLibrary = (normRoot === "/" || path.indexOf(normRoot + "/") !== 0) ? path : path.slice(normRoot.length).replace(/^\/+/, "") || "";
      if (normRoot !== "/" && path === normRoot) pathAfterLibrary = "";
      var encodedPath = encodeURIComponent(pathAfterLibrary);
      var friendlyLen = pathAfterLibrary.length;
      var encodedLen = encodedPath.length;
      fileEntries.push({ pathAfterLibrary: pathAfterLibrary, encodedPath: encodedPath, friendlyLen: friendlyLen, encodedLen: encodedLen });
    }
    fileEntries.sort(function (a, b) {
      if (b.encodedLen !== a.encodedLen) return b.encodedLen - a.encodedLen;
      return b.friendlyLen - a.friendlyLen;
    });
    var lines = ["Path,EncodedPath,FriendlyCharCount,EncodedCharCount"];
    for (var k = 0; k < fileEntries.length; k++) {
      var e = fileEntries[k];
      lines.push([escapeCSV(e.pathAfterLibrary), escapeCSV(e.encodedPath), String(e.friendlyLen), String(e.encodedLen)].join(","));
    }
    return "\ufeff" + lines.join("\r\n");
  }

  function finishOwssvrExport(parsed) {
    if (params.report === "folderCount") {
      var csv = buildFolderCountReportFromRows(parsed.rows);
      var fnBase = (exportFilename || "FolderCounts").replace(/\.(csv|xls|xlsx|xml)$/i, "") + "_FolderCounts";
      downloadExport({ csv: csv }, fnBase, { sheetName: "FolderCounts" });
      reportDone(true, "Done! Folder count report downloaded.");
      return;
    }
    if (params.report === "pathLengths") {
      var pathCsv = buildPathLengthsReportFromRows(parsed.rows);
      var pathFnBase = (exportFilename || "PathLengths").replace(/\.(csv|xls|xlsx|xml)$/i, "") + "_PathLengths";
      downloadExport({ csv: pathCsv }, pathFnBase, { sheetName: "PathLengths" });
      reportDone(true, "Done! Path lengths report downloaded.");
      return;
    }
    var rows = parsed.rows;
    if (rows.length === 0) {
      reportDone(false, "No rows to export. View setup may have failed (check permissions). Try refreshing the page and run again.");
      return;
    }
    var allColumns = { ID: true, Title: true, FileLeafRef: true, FSObjType: true };
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      for (var k in r) {
        if (Object.prototype.hasOwnProperty.call(r, k)) allColumns[k] = true;
      }
    }
    var seenCol = {};
    var columns = [];
    var colOrder = { "ID": 0, "Title": 1, "FileLeafRef": 2, "FSObjType": 3 };
    for (var c in allColumns) {
      var k = colNormKey(c);
      if (seenCol[k]) continue;
      seenCol[k] = true;
      columns.push(k);
    }
    columns.sort(function (a, b) {
      var oa = colOrder[colNormKey(a)] != null ? colOrder[colNormKey(a)] : 99;
      var ob = colOrder[colNormKey(b)] != null ? colOrder[colNormKey(b)] : 99;
      if (oa !== ob) return oa - ob;
      return a.localeCompare(b);
    });
    reportProgress("Exporting CSV");
    var fnBase = exportFilename.replace(/\.(csv|xls|xlsx|xml)$/i, "");
    downloadExport({ rows: rows, columns: columns }, fnBase, { sheetName: "Export" });
    reportDone(true, "Done! " + rows.length.toLocaleString() + " rows exported.");
  }

  function buildViewQueryForIdRange(startId, endId) {
    // User's format: Gt (exclusive lower) + Lt (exclusive upper) -> IDs (startId..endId-1)
    return "<Where><And>" +
      "<Gt><FieldRef Name='ID'/><Value Type='Counter'>" + (startId - 1) + "</Value></Gt>" +
      "<Lt><FieldRef Name='ID'/><Value Type='Counter'>" + endId + "</Value></Lt>" +
      "</And></Where>";
  }

  function shouldStopAfterEmptyIdRange(startId, maxIdResolved, emptyStreak, maxEmptyStreak) {
    // A resolved maximum is authoritative: sparse lists can have arbitrarily large
    // gaps after bulk deletions, so only stop after scanning past that item ID.
    if (maxIdResolved != null) return startId > maxIdResolved;
    return emptyStreak >= maxEmptyStreak;
  }

  async function exportViaOwssvrWithView(viewId, itemCount, retryCount) {
    retryCount = retryCount || 0;
    var allRows = [];
    var allColumns = { ID: true, Title: true, FileLeafRef: true, FSObjType: true };
    var pageNum = 0;
    var totalCount = (itemCount != null && itemCount > 0) ? itemCount : null;

    var minId = null;
    var maxIdResolved = null;
    await Promise.all([
      getMinListItemId().then(function (x) { minId = x; }),
      getMaxListItemId().then(function (x) { maxIdResolved = x; })
    ]);

    var baseStartId = 1;
    if (minId != null) baseStartId = Math.floor(minId / 100) * 100;

    var MAX_PAGES = 500;
    if (maxIdResolved != null && maxIdResolved >= baseStartId) {
      var spanIds = maxIdResolved - baseStartId;
      MAX_PAGES = Math.min(500000, Math.ceil(spanIds / PAGE_LIMIT) + 10);
    }

    var emptyStreak = 0;
    var maxEmptyStreak = 2000;

    while (pageNum < MAX_PAGES) {
      pageNum++;
      var startId = baseStartId + (pageNum - 1) * PAGE_LIMIT;
      var endId = baseStartId + pageNum * PAGE_LIMIT;

      if (maxIdResolved != null && startId > maxIdResolved) {
        break;
      }

      var viewQuery = buildViewQueryForIdRange(startId, endId);

      var uq = await setViewQueryById(viewId, viewQuery);
      if (!uq.ok) {
        if (pageNum === 1) {
          reportDone(false, "View filter failed: " + (uq.error || "unknown"));
          return;
        }
        break;
      }
      await sleep(500);

      var url = buildOwssvrUrl(viewId);
      var resp = await fetch(url, { credentials: "include", redirect: "follow" });
      if (!resp.ok && resp.status === 404) {
        if (pageNum === 1) {
          reportDone(false, "owssvr returned 404 (view may not be supported on this site)");
          return;
        }
        break;
      }
      if (!resp.ok) {
        if (resp.status === 503 && allRows.length > 0) break;
        reportDone(false, "owssvr.dll HTTP " + resp.status + ": " + resp.statusText);
        return;
      }
      var xmlText = await resp.text();
      var parsed;
      try {
        parsed = parseOwssvrXml(xmlText);
      } catch (e) {
        if (pageNum === 1) {
          reportDone(false, "owssvr XML parse error: " + (e.message || String(e)));
          return;
        }
        break;
      }
      if (parsed.rows.length === 0) {
        if (pageNum === 1 && retryCount < 1) {
          await sleep(2000);
          return exportViaOwssvrWithView(viewId, itemCount, retryCount + 1);
        }
        emptyStreak++;
        if (shouldStopAfterEmptyIdRange(startId, maxIdResolved, emptyStreak, maxEmptyStreak)) break;
        continue;
      }
      emptyStreak = 0;
      for (var i = 0; i < parsed.rows.length; i++) {
        var r = parsed.rows[i];
        allRows.push(r);
        for (var k in r) {
          if (Object.prototype.hasOwnProperty.call(r, k)) allColumns[k] = true;
        }
      }
      reportProgress("", { currentCount: allRows.length, totalCount: totalCount });
    }

    if (allRows.length === 0) {
      reportDone(false, "No rows found.");
      return;
    }

    var rowsToExport = allRows;
    if (params.incVer === false) {
      var byId = {};
      function versionNum(v) {
        var s = String(v || "0").split(".")[0];
        return parseInt(s, 10) || 0;
      }
      for (var i = 0; i < allRows.length; i++) {
        var r = allRows[i];
        var id = r.ID || r.Id || "";
        if (!id) continue;
        var cur = byId[id];
        var vCur = cur ? versionNum(cur._UIVersionString || cur.Version) : 0;
        var vNew = versionNum(r._UIVersionString || r.Version);
        if (!cur || vNew >= vCur) byId[id] = r;
      }
      rowsToExport = Object.keys(byId).sort(function (a, b) { return parseInt(a, 10) - parseInt(b, 10); }).map(function (id) { return byId[id]; });
    }

    if (rowsToExport.length === 0) {
      reportDone(false, "No data to export; view setup may have failed (check permissions). Try refreshing the page and run again.");
      return;
    }
    finishOwssvrExport({ rows: rowsToExport, columns: [] });
  }

  async function exportViaRest(itemCountParam) {
    var itemCount = itemCountParam;
    if (itemCount == null) itemCount = await getItemCount();
    var fieldNames;
    if (params.cols && Array.isArray(params.cols) && params.cols.length > 0) {
      fieldNames = mergeRequiredFields(params.cols);
      var valid = await getListFieldInternalNames();
      if (valid.ok && valid.names) {
        var filtered = [];
        for (var fi = 0; fi < fieldNames.length; fi++) {
          if (valid.names[fieldNames[fi]]) filtered.push(fieldNames[fi]);
        }
        if (filtered.length > 0) fieldNames = filtered;
      }
      fieldNames = filterLookupFieldsForRest(fieldNames, valid && valid.lookupSet);
    } else {
      var fr = await getEligibleFieldNames();
      if (!fr.ok) {
        reportDone(false, fr.error || "Failed to get list fields");
        return;
      }
      var valid = await getListFieldInternalNames();
      var validSet = (valid.ok && valid.names) ? valid.names : null;
      var base = ["ID"];
      if (!validSet || validSet["FSObjType"]) base.push("FSObjType");
      var rest = fr.fields.filter(function (n) { return n !== "ID" && n !== "FSObjType" && (!validSet || validSet[n]); });
      fieldNames = base.concat(rest);
      if (fieldNames.length <= 1) fieldNames = ["ID", "Title"].concat(fr.fields.filter(function (n) { return n !== "ID" && n !== "Title"; }).slice(0, 50));
    }
    var vrest = (typeof valid !== "undefined" && valid && valid.names) ? valid.names : {};
    var restNeedsPathCols =
      params.report === "folderCount" ||
      params.report === "pathLengths" ||
      params.report === "exportCSV";
    if (restNeedsPathCols) {
      if (fieldNames.indexOf("FileRef") < 0 && vrest["FileRef"]) fieldNames = fieldNames.concat(["FileRef"]);
      if ((params.report === "folderCount" || params.report === "pathLengths") && fieldNames.indexOf("ContentTypeId") < 0 && vrest["ContentTypeId"]) {
        fieldNames = fieldNames.concat(["ContentTypeId"]);
      }
      if (fieldNames.indexOf("FileDirRef") < 0 && vrest["FileDirRef"]) fieldNames = fieldNames.concat(["FileDirRef"]);
    }
    var select = fieldNames.join(",");
    var selectWithVersion = (select.indexOf("VersionLabel") >= 0 || select.indexOf("_UIVersionString") >= 0)
      ? select
      : select + ",VersionLabel";
    var accept = { "Accept": "application/json;odata=nometadata" };
    var rowMap = {};
    var allColumns = { ID: true, Title: true, FileLeafRef: true, FSObjType: true };
    var maxItemIdResolved = await getMaxListItemId();
    var lastId = 0;
    var pageNum = 0;
    var MAX_PAGES = 200;
    if (maxItemIdResolved != null && maxItemIdResolved >= 1) {
      MAX_PAGES = Math.min(500000, Math.ceil(maxItemIdResolved / PAGE_LIMIT) + 10);
    }

    if (params.incVer !== false) {
      reportProgress("List export with version history (paginated per item). Starting…");
      await sleep(500);
      var itemIds = [];
      lastId = 0;
      pageNum = 0;
      while (pageNum < MAX_PAGES) {
        pageNum++;
        var idUrl = listBaseUrl() + "/items?$select=ID&$orderby=ID&$top=" + PAGE_LIMIT;
        if (lastId > 0) idUrl += "&$filter=ID gt " + lastId;
        var idResp = await fetch(idUrl, { credentials: "include", headers: accept });
        if (!idResp.ok) break;
        var idJson = await idResp.json();
        var idItems = idJson.value || idJson.d?.results || [];
        if (idItems.length === 0) break;
        for (var vi = 0; vi < idItems.length; vi++) {
          var o = idItems[vi];
          var iid = o.ID != null ? o.ID : o.Id;
          if (iid != null) itemIds.push(String(iid));
        }
        lastId = parseInt(idItems[idItems.length - 1].ID || idItems[idItems.length - 1].Id, 10) || lastId;
        if (idItems.length < PAGE_LIMIT) break;
      }
      var rows = [];
      for (var idx = 0; idx < itemIds.length; idx++) {
        var itemId = itemIds[idx];
        var versions = await fetchAllVersionsForItem(itemId, selectWithVersion);
        if (versions && versions.length > 0) {
          for (var vi = 0; vi < versions.length; vi++) {
            var row = restItemToRow(versions[vi]);
            row.ID = itemId;
            if (versions[vi].VersionLabel != null) row._UIVersionString = String(versions[vi].VersionLabel);
            rows.push(row);
            for (var k in row) { if (Object.prototype.hasOwnProperty.call(row, k)) allColumns[k] = true; }
          }
        } else {
          var currentUrl = listBaseUrl() + "/items(" + itemId + ")?$select=" + encodeURIComponent(selectWithVersion);
          var currentResp = await fetch(currentUrl, { credentials: "include", headers: accept });
          if (currentResp.ok) {
            var currentJson = await currentResp.json();
            var row = restItemToRow(currentJson);
            if (!row.ID) row.ID = itemId;
            if (currentJson.VersionLabel != null) row._UIVersionString = String(currentJson.VersionLabel);
            rows.push(row);
            for (var k in row) { if (Object.prototype.hasOwnProperty.call(row, k)) allColumns[k] = true; }
          }
        }
        if ((idx + 1) % 50 === 0 || idx === itemIds.length - 1) {
          reportProgress("", { currentCount: idx + 1, totalCount: itemIds.length });
        }
      }
      if (rows.length === 0) {
        reportDone(false, "No rows found.");
        return;
      }
      if (params.report === "folderCount") {
        var folderCsvVer = buildFolderCountReportFromRows(rows);
        var folderFnBaseVer = (exportFilename || "FolderCounts").replace(/\.(csv|xls|xlsx|xml)$/i, "") + "_FolderCounts";
        downloadExport({ csv: folderCsvVer }, folderFnBaseVer, { sheetName: "FolderCounts" });
        reportDone(true, "Done! Folder count report downloaded.");
        return;
      }
      if (params.report === "pathLengths") {
        var pathCsvVer = buildPathLengthsReportFromRows(rows);
        var pathFnBaseVer = (exportFilename || "PathLengths").replace(/\.(csv|xls|xlsx|xml)$/i, "") + "_PathLengths";
        downloadExport({ csv: pathCsvVer }, pathFnBaseVer, { sheetName: "PathLengths" });
        reportDone(true, "Done! Path lengths report downloaded.");
        return;
      }
      reportProgress("Building CSV…");
      try {
        var seenCol = {};
        var columns = [];
        var colOrder = { "ID": 0, "Title": 1, "FileLeafRef": 2, "FSObjType": 3 };
        for (var c in allColumns) {
          var k = colNormKey(c);
          if (seenCol[k]) continue;
          seenCol[k] = true;
          columns.push(k);
        }
        columns.sort(function (a, b) {
          var oa = colOrder[colNormKey(a)] != null ? colOrder[colNormKey(a)] : 99;
          var ob = colOrder[colNormKey(b)] != null ? colOrder[colNormKey(b)] : 99;
          if (oa !== ob) return oa - ob;
          return a.localeCompare(b);
        });
        var fnBaseVer = exportFilename.replace(/\.(csv|xls|xlsx|xml)$/i, "");
        downloadExport({ rows: rows, columns: columns }, fnBaseVer, { sheetName: "Export" });
        reportDone(true, "Done! " + rows.length.toLocaleString() + " rows exported (with versions) via REST API.");
      } catch (err) {
        reportDone(false, "Failed to build or download CSV: " + (err.message || String(err)));
      }
      return;
    }

    while (pageNum < MAX_PAGES) {
      pageNum++;
      var url = listBaseUrl() + "/items?$select=" + encodeURIComponent(select) +
        "&$orderby=ID&$top=" + PAGE_LIMIT;
      if (lastId > 0) url += "&$filter=ID gt " + lastId;

      var resp = await fetch(url, { credentials: "include", headers: accept });
      if (!resp.ok) {
        var errBody = "";
        try { errBody = await resp.text(); } catch (_) {}
        reportDone(false, "REST API HTTP " + resp.status + ": " + resp.statusText + (errBody && errBody.length < 300 ? "\n" + errBody : ""));
        return;
      }
      var j = await resp.json();
      var items = j.value || j.d?.results || [];
      if (items.length === 0) break;

      for (var i = 0; i < items.length; i++) {
        var row = restItemToRow(items[i]);
        var id = row.ID;
        if (!id) continue;
        rowMap[id] = row;
        for (var k in row) {
          if (Object.prototype.hasOwnProperty.call(row, k)) allColumns[k] = true;
        }
      }
      lastId = parseInt(items[items.length - 1].ID || items[items.length - 1].Id, 10) || lastId;
      reportProgress("", { currentCount: Object.keys(rowMap).length, totalCount: (itemCount != null && itemCount > 0) ? itemCount : null });
      if (maxItemIdResolved != null && lastId >= maxItemIdResolved) break;
      if (items.length < PAGE_LIMIT) break;
    }

    var totalRows = Object.keys(rowMap).length;
    if (totalRows === 0) {
      reportDone(false, "No rows found.");
      return;
    }
    if (params.report === "folderCount") {
      var restRows = Object.keys(rowMap).sort(function (a, b) { return parseInt(a, 10) - parseInt(b, 10); }).map(function (id) { return rowMap[id]; });
      var folderCsv = buildFolderCountReportFromRows(restRows);
      var folderFnBase = (exportFilename || "FolderCounts").replace(/\.(csv|xls|xlsx|xml)$/i, "") + "_FolderCounts";
      downloadExport({ csv: folderCsv }, folderFnBase, { sheetName: "FolderCounts" });
      reportDone(true, "Done! Folder count report downloaded.");
      return;
    }
    if (params.report === "pathLengths") {
      var pathRestRows = Object.keys(rowMap).sort(function (a, b) { return parseInt(a, 10) - parseInt(b, 10); }).map(function (id) { return rowMap[id]; });
      var pathCsvRest = buildPathLengthsReportFromRows(pathRestRows);
      var pathFnBaseRest = (exportFilename || "PathLengths").replace(/\.(csv|xls|xlsx|xml)$/i, "") + "_PathLengths";
      downloadExport({ csv: pathCsvRest }, pathFnBaseRest, { sheetName: "PathLengths" });
      reportDone(true, "Done! Path lengths report downloaded.");
      return;
    }
    var seenCol = {};
    var columns = [];
    var colOrder = { "ID": 0, "Title": 1, "FileLeafRef": 2, "FSObjType": 3 };
    for (var c in allColumns) {
      var k = colNormKey(c);
      if (seenCol[k]) continue;
      seenCol[k] = true;
      columns.push(k);
    }
    columns.sort(function (a, b) {
      var oa = colOrder[colNormKey(a)] != null ? colOrder[colNormKey(a)] : 99;
      var ob = colOrder[colNormKey(b)] != null ? colOrder[colNormKey(b)] : 99;
      if (oa !== ob) return oa - ob;
      return a.localeCompare(b);
    });
    reportProgress("Exporting CSV");
    var restRowsArr = Object.keys(rowMap).sort(function (a, b) { return parseInt(a, 10) - parseInt(b, 10); }).map(function (id) { return rowMap[id]; });
    var fnBaseRest = exportFilename.replace(/\.(csv|xls|xlsx|xml)$/i, "");
    downloadExport({ rows: restRowsArr, columns: columns }, fnBaseRest, { sheetName: "Export" });
    reportDone(true, "Done! " + totalRows.toLocaleString() + " rows exported via REST API.");
  }

  async function resolveListIdFromListUrl() {
    var spi = window._spPageContextInfo;
    if (!spi || !siteUrl) return null;
    var listUrl = spi.listUrl || spi.listServerRelativeUrl;
    if (!listUrl) return null;
    try {
      var enc = encodeURIComponent("'" + listUrl + "'");
      var r = await fetch(siteUrl + "/_api/web/GetList(@listUrl)/Id?@listUrl=" + enc, {
        credentials: "include",
        headers: { "Accept": "application/json;odata=nometadata" }
      });
      if (!r.ok) return null;
      var j = await r.json();
      var id = (j.value || "").replace(/[{}]/g, "").trim();
      return id ? normalizeGuid(id) : null;
    } catch (_) { return null; }
  }

  function isExternalUser(loginName) {
    if (!loginName || typeof loginName !== "string") return false;
    var ln = loginName.toLowerCase();
    return ln.indexOf("#ext#") >= 0 || ln.indexOf("c:0t.c|") >= 0;
  }

  function getListTitleFromContext() {
    var spi = window._spPageContextInfo;
    if (!spi) return null;
    if (spi.listTitle && String(spi.listTitle).trim()) return String(spi.listTitle).trim();
    if (spi.listUrl) {
      var segments = String(spi.listUrl).split("/").filter(Boolean);
      if (segments.length > 0) return segments[segments.length - 1];
    }
    return null;
  }

  function listByTitleUrl(title) {
    if (!title) return null;
    var safe = String(title).replace(/'/g, "''");
    return siteUrl + "/_api/web/lists/GetByTitle('" + safe + "')";
  }

  async function exportPermissionsReport() {
    var accept = { "Accept": "application/json;odata=nometadata" };
    var listTitle = getListTitleFromContext();
    var listBase = listTitle ? listByTitleUrl(listTitle) : listBaseUrl();
    reportProgress("Permissions: Loading list…");
    var listMeta = await fetch(listBase + "?$select=HasUniqueRoleAssignments,BaseTemplate,Title,ItemCount", { credentials: "include", headers: accept });
    if (!listMeta.ok) {
      reportDone(false, "Could not load list: " + listMeta.status);
      return;
    }
    var listData = await listMeta.json();
    var listHasUnique = listData.HasUniqueRoleAssignments === true;
    var isDocLib = parseInt(listData.BaseTemplate, 10) === 101;
    var totalItemCount = Math.max(0, parseInt(listData.ItemCount, 10) || 0);
    var principalScopes = {};
    var principalCache = {};
    var uniquePermResults = [];
    var PAGE_SIZE = 5000;
    var CONCURRENCY = 6;
    async function fetchWithRetry(url, opts, tries) {
      tries = tries || 8;
      for (var attempt = 1; attempt <= tries; attempt++) {
        var r = await fetch(url, opts);
        if (r.ok) return r;
        if ([429, 503, 504].indexOf(r.status) >= 0) {
          var ra = r.headers.get("Retry-After");
          var wait = ra ? (parseInt(ra, 10) * 1000) : Math.min(500 * Math.pow(2, attempt - 1), 20000);
          await sleep(wait);
          continue;
        }
        throw new Error("HTTP " + r.status + ": " + (await r.text()));
      }
      throw new Error("Throttled too many times.");
    }
    function addPrincipal(member, scope, scopePath) {
      if (!member) return;
      var id = member.Id != null ? String(member.Id) : (member.id != null ? String(member.id) : "");
      if (!id) return;
      var title = member.Title || member.LoginName || "";
      var loginName = member.LoginName || "";
      var email = member.Email || "";
      var principalType = member.PrincipalType != null ? parseInt(member.PrincipalType, 10) : (member.PrincipalType != null ? parseInt(member.PrincipalType, 10) : 1);
      if (!principalScopes[id]) {
        principalScopes[id] = { id: id, title: title, loginName: loginName, email: email, principalType: principalType, scopes: [] };
      }
      var scopeKey = scope === "List" ? "List" : (scopePath || scope);
      if (principalScopes[id].scopes.indexOf(scopeKey) < 0) principalScopes[id].scopes.push(scopeKey);
    }
    async function resolvePrincipalById(principalId) {
      if (principalCache[principalId]) return principalCache[principalId];
      var pid = parseInt(principalId, 10);
      if (isNaN(pid)) return null;
      try {
        var userResp = await fetch(siteUrl + "/_api/web/GetUserById(" + pid + ")", { credentials: "include", headers: accept });
        if (userResp.ok) {
          var userData = await userResp.json();
          var m = {
            Id: userData.Id != null ? userData.Id : userData.id,
            Title: userData.Title,
            LoginName: userData.LoginName,
            Email: userData.Email || "",
            PrincipalType: userData.PrincipalType != null ? userData.PrincipalType : 1
          };
          principalCache[principalId] = m;
          return m;
        }
      } catch (_) {}
      try {
        var groupResp = await fetch(siteUrl + "/_api/web/sitegroups/GetById(" + pid + ")", { credentials: "include", headers: accept });
        if (groupResp.ok) {
          var groupData = await groupResp.json();
          var gm = {
            Id: groupData.Id != null ? groupData.Id : groupData.id,
            Title: groupData.Title || groupData.LoginName,
            LoginName: groupData.LoginName || groupData.Title || "",
            Email: groupData.Email || "",
            PrincipalType: groupData.PrincipalType != null ? groupData.PrincipalType : 8
          };
          principalCache[principalId] = gm;
          return gm;
        }
      } catch (_) {}
      principalCache[principalId] = null;
      return null;
    }
    function raResults(data) {
      return data.value || data.results || (data.d && data.d.results) || (data.d && data.d.value) || [];
    }
    if (listHasUnique) {
      reportProgress("Permissions: Loading list role assignments…");
      var listRa = await fetch(listBase + "/roleassignments?$expand=Member,RoleDefinitionBindings", { credentials: "include", headers: accept });
      if (listRa.ok) {
        var listRaData = await listRa.json();
        var results = raResults(listRaData);
        for (var i = 0; i < results.length; i++) {
          var ra = results[i];
          var m = ra.Member || ra.member;
          if (!m && (ra.PrincipalId != null || ra.principalId != null)) {
            var pid = ra.PrincipalId != null ? ra.PrincipalId : ra.principalId;
            m = await resolvePrincipalById(String(pid));
          }
          if (m) addPrincipal(m, "List", null);
        }
      }
    } else {
      reportProgress("Permissions: Loading site role assignments (list inherits)…");
      var webRa = await fetch(siteUrl + "/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings", { credentials: "include", headers: accept });
      if (webRa.ok) {
        var webRaData = await webRa.json();
        var webResults = raResults(webRaData);
        for (var w = 0; w < webResults.length; w++) {
          var wra = webResults[w];
          var wm = wra.Member || wra.member;
          if (!wm && (wra.PrincipalId != null || wra.principalId != null)) {
            var wpid = wra.PrincipalId != null ? wra.PrincipalId : wra.principalId;
            wm = await resolvePrincipalById(String(wpid));
          }
          if (wm) addPrincipal(wm, "List", null);
        }
      }
    }
    var acceptVerbose = { "Accept": "application/json;odata=verbose" };
    var allItems = [];
    var nextUrl = listBase + "/items?$select=Id,FileRef,FileDirRef,FileLeafRef,FSObjType,HasUniqueRoleAssignments&$top=" + PAGE_SIZE;
    reportProgress("Permissions: Scanning items…", { currentCount: 0, totalCount: totalItemCount });
    while (nextUrl) {
      var r = await fetchWithRetry(nextUrl, { credentials: "include", headers: acceptVerbose }, 6);
      var d = await r.json();
      var batch = (d.d && d.d.results) || [];
      allItems = allItems.concat(batch);
      nextUrl = (d.d && d.d.__next) || null;
      reportProgress("Permissions: Scanning items…", { currentCount: allItems.length, totalCount: totalItemCount });
      if (batch.length < PAGE_SIZE) break;
      await sleep(100);
    }
    var exceptions = [];
    for (var ei = 0; ei < allItems.length; ei++) {
      var itm = allItems[ei];
      if (itm.HasUniqueRoleAssignments !== true) continue;
      var itemLabel = itm.FileRef || (itm.FileDirRef && itm.FileLeafRef ? (itm.FileDirRef + "/" + itm.FileLeafRef).replace(/\/+/g, "/") : null) || itm.FileLeafRef || ("(ID " + itm.Id + ")");
      if (itemLabel && itemLabel.indexOf("#") >= 0) itemLabel = itemLabel.slice(itemLabel.indexOf("#") + 1).trim();
      exceptions.push({ Id: itm.Id, Item: itemLabel });
    }
    async function getPrincipalsForItem(itemId) {
      var url = listBase + "/items(" + itemId + ")/RoleAssignments?$select=Member/Title,Member/LoginName,Member/PrincipalType&$expand=Member&$top=5000";
      var principals = [];
      var next = url;
      while (next) {
        var resp = await fetchWithRetry(next, { credentials: "include", headers: accept }, 8);
        var data = await resp.json();
        var page = data.value || data.results || [];
        for (var pi = 0; pi < page.length; pi++) {
          var m = page[pi].Member || page[pi].member;
          if (m) principals.push(m.Title || m.LoginName || "");
        }
        next = data["@odata.nextLink"] || null;
      }
      var seen = {};
      var uniq = [];
      for (var u = 0; u < principals.length; u++) {
        var name = (principals[u] || "").trim();
        var key = name.toLowerCase();
        if (!name || seen[key]) continue;
        seen[key] = true;
        uniq.push(name);
      }
      return uniq.join(", ");
    }
    var resultIdx = 0;
    reportProgress("Permissions: Loading principals for " + exceptions.length + " item(s)…", { currentCount: 0, totalCount: exceptions.length });
    function runWorker() {
      return new Promise(function (resolve) {
        function work() {
          var i = resultIdx++;
          if (i >= exceptions.length) { resolve(); return; }
          var ex = exceptions[i];
          getPrincipalsForItem(ex.Id).then(function (principals) {
            uniquePermResults.push({ Item: ex.Item, Principals: principals });
            reportProgress("Permissions: Loading principals…", { currentCount: uniquePermResults.length, totalCount: exceptions.length });
            work();
          }).catch(function (err) {
            uniquePermResults.push({ Item: ex.Item, Principals: "ERROR: " + (err && err.message ? err.message : String(err)) });
            reportProgress("Permissions: Loading principals…", { currentCount: uniquePermResults.length, totalCount: exceptions.length });
            work();
          });
        }
        work();
      });
    }
    var workers = [];
    for (var w = 0; w < CONCURRENCY; w++) workers.push(runWorker());
    await Promise.all(workers);
    uniquePermResults.sort(function (a, b) { return (a.Item || "").localeCompare(b.Item || ""); });
    var groupIdToName = {};
    for (var pid in principalScopes) {
      var p = principalScopes[pid];
      var pt = p.principalType;
      if (pt === 4 || pt === 8) {
        groupIdToName[pid] = p.title || p.loginName || ("Group " + pid);
      }
    }
    var numGroups = Object.keys(groupIdToName).length;
    if (numGroups > 0) {
      reportProgress("Permissions: Resolving " + numGroups + " group(s)…");
    }
    var userMap = {};
    function ensureUser(id, title, loginName, email) {
      if (!userMap[id]) userMap[id] = { title: title || "", loginName: loginName || "", email: email || "", groups: [], scopes: [] };
    }
    function addScopesToUser(id, scopes, groupName) {
      if (!userMap[id]) userMap[id] = { title: "", loginName: "", email: "", groups: [], scopes: [] };
      for (var si = 0; scopes && si < scopes.length; si++) {
        if (userMap[id].scopes.indexOf(scopes[si]) < 0) userMap[id].scopes.push(scopes[si]);
      }
      if (groupName && userMap[id].groups.indexOf(groupName) < 0) userMap[id].groups.push(groupName);
    }
    for (var principalId in principalScopes) {
      var pr = principalScopes[principalId];
      if (pr.principalType === 4 || pr.principalType === 8) continue;
      ensureUser(principalId, pr.title, pr.loginName, pr.email);
      for (var si = 0; pr.scopes && si < pr.scopes.length; si++) {
        if (userMap[principalId].scopes.indexOf(pr.scopes[si]) < 0) userMap[principalId].scopes.push(pr.scopes[si]);
      }
    }
    var gIndex = 0;
    for (var gid in groupIdToName) {
      gIndex++;
      if (numGroups > 3 && (gIndex % 5 === 0 || gIndex === numGroups)) {
        reportProgress("Permissions: Resolving group " + gIndex + " of " + numGroups + "…");
      }
      var gScopes = (principalScopes[gid] && principalScopes[gid].scopes) ? principalScopes[gid].scopes : [];
      var gName = groupIdToName[gid];
      try {
        var usersUrl = siteUrl + "/_api/web/sitegroups/GetById(" + gid + ")/users";
        var usersResp = await fetch(usersUrl, { credentials: "include", headers: accept });
        if (!usersResp.ok) continue;
        var usersData = await usersResp.json();
        var groupUsers = usersData.value || usersData.results || [];
        for (var u = 0; u < groupUsers.length; u++) {
          var gu = groupUsers[u];
          var uid = gu.Id != null ? String(gu.Id) : "";
          if (!uid) continue;
          ensureUser(uid, gu.Title, gu.LoginName, gu.Email);
          addScopesToUser(uid, gScopes, gName);
        }
      } catch (_) {}
      await sleep(50);
    }
    function formatScopeStr(scopes) {
      if (!scopes || !scopes.length) return "List";
      var list = [];
      var hasList = false;
      for (var s = 0; s < scopes.length; s++) {
        var v = (scopes[s] && String(scopes[s]).trim()) || "";
        if (v === "List") hasList = true;
        else if (v && list.indexOf(v) < 0) list.push(v);
      }
      if (hasList) list.unshift("List");
      return list.length ? list.join("; ") : "List";
    }
    reportProgress("Permissions: Building report…");
    var rows = [];
    for (var uid in userMap) {
      var u = userMap[uid];
      rows.push({
        User: u.title,
        LoginName: u.loginName,
        Email: u.email,
        Group: (u.groups && u.groups.length) ? u.groups.join("; ") : "",
        InternalOrExternal: isExternalUser(u.loginName) ? "External" : "Internal",
        UniquePermissionPaths: formatScopeStr(u.scopes)
      });
    }
    if (rows.length === 0 && Object.keys(principalScopes).length > 0) {
      for (var pid in principalScopes) {
        var pr = principalScopes[pid];
        rows.push({
          User: pr.title || pr.loginName || ("Principal " + pid),
          LoginName: pr.loginName || "",
          Email: pr.email || "",
          Group: (pr.principalType === 4 || pr.principalType === 8) ? "(group)" : "",
          InternalOrExternal: isExternalUser(pr.loginName) ? "External" : "Internal",
          UniquePermissionPaths: formatScopeStr(pr.scopes)
        });
      }
    }
    var lines = ["User,LoginName,Email,Group,InternalOrExternal,UniquePermissionPaths"];
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      lines.push([
        escapeCSV(row.User),
        escapeCSV(row.LoginName),
        escapeCSV(row.Email),
        escapeCSV(row.Group),
        escapeCSV(row.InternalOrExternal),
        escapeCSV(row.UniquePermissionPaths)
      ].join(","));
    }
    var csv = "\ufeff" + lines.join("\r\n");
    var fnBase = (exportFilename || "Permissions").replace(/\.(csv|xls|xlsx|xml)$/i, "") + "_Permissions";
    if (params.format === "xlsx") {
      var mainAoa = parseCSVToAoa(csv);
      var itemsAoa = [["Item", "Principals"]];
      for (var id = 0; id < uniquePermResults.length; id++) {
        var d = uniquePermResults[id];
        itemsAoa.push([d.Item || "", d.Principals || ""]);
      }
      downloadExport({ sheets: [{ name: "Permissions", aoa: mainAoa }, { name: "Items with unique permissions", aoa: itemsAoa }] }, fnBase, {});
    } else {
      downloadExport({ csv: csv }, fnBase, { sheetName: "Permissions" });
    }
    reportDone(true, "Done! Permissions report downloaded (" + rows.length + " rows).");
  }

  function principalTypeLabel(pt) {
    var n = parseInt(pt, 10);
    if (n === 1) return "User";
    if (n === 2) return "Distribution list";
    if (n === 4) return "Security group";
    if (n === 8) return "SharePoint group";
    return "User";
  }

  async function exportPermissionsMatrixReport() {
    var accept = { "Accept": "application/json;odata=nometadata" };
    var acceptVerbose = { "Accept": "application/json;odata=verbose" };
    var PAGE_SIZE = 5000;
    var CONCURRENCY = 6;
    async function fetchWithRetry(url, opts, tries) {
      tries = tries || 8;
      for (var attempt = 1; attempt <= tries; attempt++) {
        var r = await fetch(url, opts);
        if (r.ok) return r;
        if ([429, 503, 504].indexOf(r.status) >= 0) {
          var ra = r.headers.get("Retry-After");
          var wait = ra ? (parseInt(ra, 10) * 1000) : Math.min(500 * Math.pow(2, attempt - 1), 20000);
          await sleep(wait);
          continue;
        }
        throw new Error("HTTP " + r.status);
      }
      throw new Error("Throttled too many times.");
    }

    if (params.matrixWholeSite) {
      reportProgress("Permissions matrix: Loading role definitions…");
      var roleDefRespWs = await fetch(siteUrl + "/_api/web/roledefinitions?$select=Name,Order&$filter=Hidden eq false&$orderby=Order asc", { credentials: "include", headers: accept });
      var roleNamesWs = [];
      if (roleDefRespWs.ok) {
        var roleDefDataWs = await roleDefRespWs.json();
        var defsWs = roleDefDataWs.value || roleDefDataWs.results || [];
        for (var rdw = 0; rdw < defsWs.length; rdw++) {
          var nw = (defsWs[rdw].Name || "").trim();
          if (nw && roleNamesWs.indexOf(nw) < 0) roleNamesWs.push(nw);
        }
      }
      if (roleNamesWs.length === 0) roleNamesWs = ["Full Control", "Design", "Contribute", "Edit", "Read", "View Only", "Restricted Read", "Approve", "Manage Hierarchy", "Restricted View", "Restricted Interfaces for Translation"];
      reportProgress("Permissions matrix: Loading site lists…");
      var listsResp = await fetch(siteUrl + "/_api/web/lists?$select=Id,Title,BaseTemplate,HasUniqueRoleAssignments,ItemCount&$filter=Hidden eq false&$top=5000", { credentials: "include", headers: accept });
      if (!listsResp.ok) {
        reportDone(false, "Could not load site lists: " + listsResp.status);
        return;
      }
      var listsData = await listsResp.json();
      var allLists = listsData.value || listsData.results || [];
      var webResp = await fetch(siteUrl + "/_api/web?$select=Title,ServerRelativeUrl", { credentials: "include", headers: accept });
      var webTitle = "Site";
      var webPath = "";
      if (webResp.ok) {
        var webData = await webResp.json();
        webTitle = (webData.Title || webData.title || "Site").trim();
        webPath = (webData.ServerRelativeUrl || webData.serverRelativeUrl || "").trim();
      }
      var allMatrixRowsWs = [];
      var webRaResp = await fetch(siteUrl + "/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings&$top=5000", { credentials: "include", headers: accept });
      if (webRaResp.ok) {
        var webRaData = await webRaResp.json();
        var webRaPage = webRaData.value || webRaData.results || [];
        var webRaNext = webRaData["@odata.nextLink"];
        while (webRaNext) {
          var wnResp = await fetchWithRetry(webRaNext, { credentials: "include", headers: accept }, 6);
          var wnData = await wnResp.json();
          webRaPage = webRaPage.concat(wnData.value || wnData.results || []);
          webRaNext = wnData["@odata.nextLink"] || null;
        }
        var webPrincipals = [];
        for (var wpi = 0; wpi < webRaPage.length; wpi++) {
          var wra = webRaPage[wpi];
          var wm = wra.Member || wra.member;
          if (!wm) continue;
          var wbindings = wra.RoleDefinitionBindings || wra.roleDefinitionBindings || [];
          var wroleSet = {};
          for (var wbi = 0; wbi < wbindings.length; wbi++) {
            var wrn = (wbindings[wbi].Name || wbindings[wbi].name || "").trim();
            if (wrn) wroleSet[wrn] = true;
          }
          var wpt = principalTypeLabel(wm.PrincipalType);
          var wtit = (wm.Title || wm.LoginName || "").trim();
          var wlog = (wm.LoginName || "").trim();
          var wisGrp = (parseInt(wm.PrincipalType, 10) === 4 || parseInt(wm.PrincipalType, 10) === 8);
          webPrincipals.push({ title: wtit, loginName: wlog, principalType: wpt, givenThroughSuffix: "", roleSet: wroleSet, isGroup: wisGrp, groupId: wisGrp ? (wm.Id != null ? String(wm.Id) : "") : "" });
        }
        var webExpanded = [];
        for (var wei = 0; wei < webPrincipals.length; wei++) {
          var wip = webPrincipals[wei];
          if (!wip.isGroup) {
            webExpanded.push({ title: wip.title, loginName: wip.loginName, principalType: wip.principalType, givenThroughSuffix: wip.givenThroughSuffix, roleSet: wip.roleSet });
            continue;
          }
          webExpanded.push({ title: wip.title, loginName: wip.loginName, principalType: wip.principalType, givenThroughSuffix: wip.givenThroughSuffix, roleSet: wip.roleSet });
          if (!wip.groupId) continue;
          try {
            var wguResp = await fetch(siteUrl + "/_api/web/sitegroups/GetById(" + wip.groupId + ")/users", { credentials: "include", headers: accept });
            if (!wguResp.ok) continue;
            var wguData = await wguResp.json();
            var wgUsers = wguData.value || wguData.results || [];
            for (var wg = 0; wg < wgUsers.length; wg++) {
              var wu = wgUsers[wg];
              webExpanded.push({
                title: (wu.Title || wu.LoginName || "").trim(),
                loginName: (wu.LoginName || "").trim(),
                principalType: "User",
                givenThroughSuffix: " - " + (wip.title || "Group") + " Members",
                roleSet: wip.roleSet
              });
            }
          } catch (_) {}
        }
        var siteGivenBase = "Inherited (From ItemPath: " + (webPath || webTitle) + ")";
        for (var wxi = 0; wxi < webExpanded.length; wxi++) {
          var wp = webExpanded[wxi];
          allMatrixRowsWs.push({
            containerName: webTitle,
            containerType: "Page",
            itemPath: webPath || webTitle,
            inheritance: "Inherited",
            details: "",
            userOrGroup: wp.title,
            principalType: wp.principalType,
            accountName: wp.loginName,
            givenThrough: siteGivenBase + (wp.givenThroughSuffix || ""),
            roleSet: wp.roleSet
          });
        }
      }
      for (var listIdx = 0; listIdx < allLists.length; listIdx++) {
        var lst = allLists[listIdx];
        var listIdCur = normalizeGuid(lst.Id);
        var listBaseCur = siteUrl + "/_api/web/lists(guid'" + listIdCur + "')";
        var listTitleCur = (lst.Title || lst.title || "").trim() || ("List " + listIdCur);
        var listHasUniqueCur = lst.HasUniqueRoleAssignments === true;
        var totalItemCountCur = Math.max(0, parseInt(lst.ItemCount, 10) || 0);
        var containerTypeCur = parseInt(lst.BaseTemplate, 10) === 101 ? "Library" : "List";
        reportProgress("Permissions matrix: " + listTitleCur + "…", { currentCount: listIdx + 1, totalCount: allLists.length + 1 });
        var inheritedSourcePathCur = "";
        try {
          if (listHasUniqueCur) {
            var lrfCur = await fetch(listBaseCur + "/rootfolder?$select=ServerRelativeUrl", { credentials: "include", headers: accept });
            if (lrfCur.ok) {
              var lrfDataCur = await lrfCur.json();
              inheritedSourcePathCur = (lrfDataCur.ServerRelativeUrl || lrfDataCur.serverRelativeUrl || "").trim();
            }
          }
          if (!inheritedSourcePathCur) {
            inheritedSourcePathCur = (webPath || webTitle);
          }
          if (!inheritedSourcePathCur) inheritedSourcePathCur = listTitleCur;
        } catch (_) {}
        var inheritedRaUrlCur = listHasUniqueCur ? (listBaseCur + "/roleassignments") : (siteUrl + "/_api/web/roleassignments");
        var inheritedRaRespCur = await fetch(inheritedRaUrlCur + "?$expand=Member,RoleDefinitionBindings&$top=5000", { credentials: "include", headers: accept });
        var inheritedPrincipalsCur = [];
        if (inheritedRaRespCur.ok) {
          var inheritedRaDataCur = await inheritedRaRespCur.json();
          var iraPageCur = inheritedRaDataCur.value || inheritedRaDataCur.results || [];
          var iraNextCur = inheritedRaDataCur["@odata.nextLink"];
          while (iraNextCur) {
            var iraRCur = await fetchWithRetry(iraNextCur, { credentials: "include", headers: accept }, 6);
            var iraDCur = await iraRCur.json();
            iraPageCur = iraPageCur.concat(iraDCur.value || iraDCur.results || []);
            iraNextCur = iraDCur["@odata.nextLink"] || null;
          }
          for (var irac = 0; irac < iraPageCur.length; irac++) {
            var raCur = iraPageCur[irac];
            var mCur = raCur.Member || raCur.member;
            if (!mCur) continue;
            var bindCur = raCur.RoleDefinitionBindings || raCur.roleDefinitionBindings || [];
            var roleSetCur = {};
            for (var rbc = 0; rbc < bindCur.length; rbc++) {
              var rnc = (bindCur[rbc].Name || bindCur[rbc].name || "").trim();
              if (rnc) roleSetCur[rnc] = true;
            }
            var ptCur = principalTypeLabel(mCur.PrincipalType);
            var titCur = (mCur.Title || mCur.LoginName || "").trim();
            var logCur = (mCur.LoginName || "").trim();
            var isGrpCur = (parseInt(mCur.PrincipalType, 10) === 4 || parseInt(mCur.PrincipalType, 10) === 8);
            inheritedPrincipalsCur.push({ title: titCur, loginName: logCur, principalType: ptCur, givenThroughSuffix: "", roleSet: roleSetCur, isGroup: isGrpCur, groupId: isGrpCur ? (mCur.Id != null ? String(mCur.Id) : "") : "" });
          }
          var expandedCur = [];
          for (var eic = 0; eic < inheritedPrincipalsCur.length; eic++) {
            var ipCur = inheritedPrincipalsCur[eic];
            if (!ipCur.isGroup) {
              expandedCur.push({ title: ipCur.title, loginName: ipCur.loginName, principalType: ipCur.principalType, givenThroughSuffix: ipCur.givenThroughSuffix, roleSet: ipCur.roleSet });
              continue;
            }
            expandedCur.push({ title: ipCur.title, loginName: ipCur.loginName, principalType: ipCur.principalType, givenThroughSuffix: ipCur.givenThroughSuffix, roleSet: ipCur.roleSet });
            if (!ipCur.groupId) continue;
            try {
              var guRespCur = await fetch(siteUrl + "/_api/web/sitegroups/GetById(" + ipCur.groupId + ")/users", { credentials: "include", headers: accept });
              if (!guRespCur.ok) continue;
              var guDataCur = await guRespCur.json();
              var gUsersCur = guDataCur.value || guDataCur.results || [];
              for (var guc = 0; guc < gUsersCur.length; guc++) {
                var uCur = gUsersCur[guc];
                expandedCur.push({
                  title: (uCur.Title || uCur.LoginName || "").trim(),
                  loginName: (uCur.LoginName || "").trim(),
                  principalType: "User",
                  givenThroughSuffix: " - " + (ipCur.title || "Group") + " Members",
                  roleSet: ipCur.roleSet
                });
              }
            } catch (_) {}
          }
          inheritedPrincipalsCur = expandedCur;
        }
        var allItemsCur = [];
        var nextUrlCur = listBaseCur + "/items?$select=Id,FileRef,FileDirRef,FileLeafRef,FSObjType,HasUniqueRoleAssignments&$top=" + PAGE_SIZE;
        while (nextUrlCur) {
          var rCur = await fetchWithRetry(nextUrlCur, { credentials: "include", headers: acceptVerbose }, 6);
          var dCur = await rCur.json();
          var batchCur = (dCur.d && dCur.d.results) || [];
          allItemsCur = allItemsCur.concat(batchCur);
          nextUrlCur = (dCur.d && dCur.d.__next) || null;
          if (batchCur.length < PAGE_SIZE) break;
          await sleep(100);
        }
        var exceptionsCur = [];
        var inheritedItemsCur = [];
        for (var eic2 = 0; eic2 < allItemsCur.length; eic2++) {
          var itmCur = allItemsCur[eic2];
          var itemLabelCur = itmCur.FileRef || (itmCur.FileDirRef && itmCur.FileLeafRef ? (itmCur.FileDirRef + "/" + itmCur.FileLeafRef).replace(/\/+/g, "/") : null) || itmCur.FileLeafRef || ("(ID " + itmCur.Id + ")");
          if (itemLabelCur && itemLabelCur.indexOf("#") >= 0) itemLabelCur = itemLabelCur.slice(itemLabelCur.indexOf("#") + 1).trim();
          if (itmCur.HasUniqueRoleAssignments === true) {
            exceptionsCur.push({ Id: itmCur.Id, Item: itemLabelCur });
          } else {
            inheritedItemsCur.push({ Id: itmCur.Id, Item: itemLabelCur });
          }
        }
        function normalizePathCur(p) {
          if (!p || typeof p !== "string") return "";
          return p.replace(/\\/g, "/").replace(/\/+/g, "/").replace(/\/$/, "").replace(/^\//, "").trim();
        }
        function getPathAncestorsCur(path) {
          var n = normalizePathCur(path);
          if (!n) return [];
          var parts = n.split("/").filter(Boolean);
          if (parts.length <= 1) return [];
          var ancestors = [];
          for (var a = 1; a < parts.length; a++) ancestors.push(parts.slice(0, a).join("/"));
          return ancestors;
        }
        function getTopmostInheritanceSourceCur(itemPath, exceptionPathSetCur, defaultRootCur) {
          var ancestors = getPathAncestorsCur(itemPath);
          for (var ai = 0; ai < ancestors.length; ai++) {
            if (exceptionPathSetCur[ancestors[ai]]) return "/" + ancestors[ai];
          }
          return defaultRootCur || inheritedSourcePathCur;
        }
        var exceptionPathSetCur = {};
        for (var exi = 0; exi < exceptionsCur.length; exi++) exceptionPathSetCur[normalizePathCur(exceptionsCur[exi].Item)] = true;
        async function getMatrixRowsForItemCur(itemIdCur, itemPathCur) {
          var urlCur = listBaseCur + "/items(" + itemIdCur + ")/RoleAssignments?$expand=Member,RoleDefinitionBindings&$top=5000";
          var rowsCur = [];
          var nextCur = urlCur;
          while (nextCur) {
            var respCur = await fetchWithRetry(nextCur, { credentials: "include", headers: accept }, 8);
            var dataCur = await respCur.json();
            var pageCur = dataCur.value || dataCur.results || [];
            for (var pi = 0; pi < pageCur.length; pi++) {
              var ra = pageCur[pi];
              var m = ra.Member || ra.member;
              if (!m) continue;
              var bindings = ra.RoleDefinitionBindings || ra.roleDefinitionBindings || [];
              var roleSet = {};
              for (var bi = 0; bi < bindings.length; bi++) {
                var rn = (bindings[bi].Name || bindings[bi].name || "").trim();
                if (rn) roleSet[rn] = true;
              }
              var principalType = principalTypeLabel(m.PrincipalType);
              var title = (m.Title || m.LoginName || "").trim();
              var loginName = (m.LoginName || "").trim();
              var isGroup = (parseInt(m.PrincipalType, 10) === 4 || parseInt(m.PrincipalType, 10) === 8);
              rowsCur.push({ title: title, loginName: loginName, principalType: principalType, givenThrough: "Explicit", roleSet: roleSet, isGroup: isGroup, groupId: isGroup ? (m.Id != null ? String(m.Id) : "") : "" });
            }
            nextCur = dataCur["@odata.nextLink"] || null;
          }
          var expandedCur2 = [];
          for (var ri = 0; ri < rowsCur.length; ri++) {
            var row = rowsCur[ri];
            if (!row.isGroup) {
              expandedCur2.push({ title: row.title, loginName: row.loginName, principalType: row.principalType, givenThrough: row.givenThrough, roleSet: row.roleSet, viaGroup: false });
              continue;
            }
            expandedCur2.push({ title: row.title, loginName: row.loginName, principalType: row.principalType, givenThrough: row.givenThrough, roleSet: row.roleSet, viaGroup: false });
            if (!row.groupId) continue;
            try {
              var usersRespCur = await fetch(siteUrl + "/_api/web/sitegroups/GetById(" + row.groupId + ")/users", { credentials: "include", headers: accept });
              if (!usersRespCur.ok) continue;
              var usersDataCur = await usersRespCur.json();
              var groupUsersCur = usersDataCur.value || usersDataCur.results || [];
              for (var gu = 0; gu < groupUsersCur.length; gu++) {
                var u = groupUsersCur[gu];
                expandedCur2.push({
                  title: (u.Title || u.LoginName || "").trim(),
                  loginName: (u.LoginName || "").trim(),
                  principalType: "User",
                  givenThrough: (row.title || "Group") + " Members",
                  roleSet: row.roleSet,
                  viaGroup: true
                });
              }
            } catch (_) {}
          }
          return expandedCur2.map(function (e) {
            return {
              itemPath: itemPathCur,
              inheritance: e.viaGroup ? "Inherited" : "Custom",
              details: "",
              userOrGroup: e.title,
              principalType: e.principalType,
              accountName: e.loginName,
              givenThrough: e.givenThrough,
              roleSet: e.roleSet
            };
          });
        }
        for (var exIdx = 0; exIdx < exceptionsCur.length; exIdx++) {
          var exCur = exceptionsCur[exIdx];
          var itemRowsCur = await getMatrixRowsForItemCur(exCur.Id, exCur.Item);
          for (var irc = 0; irc < itemRowsCur.length; irc++) {
            var rowCur = itemRowsCur[irc];
            rowCur.containerName = listTitleCur;
            rowCur.containerType = containerTypeCur;
            allMatrixRowsWs.push(rowCur);
          }
        }
        var givenThroughBaseCur = "Inherited (From ItemPath: " + inheritedSourcePathCur + ")";
        for (var ii = 0; ii < inheritedItemsCur.length; ii++) {
          var pathCur = inheritedItemsCur[ii].Item;
          var sourcePathCur = getTopmostInheritanceSourceCur(pathCur, exceptionPathSetCur, inheritedSourcePathCur);
          var givenBaseCur = "Inherited (From ItemPath: " + sourcePathCur + ")";
          for (var ip = 0; ip < inheritedPrincipalsCur.length; ip++) {
            var pCur = inheritedPrincipalsCur[ip];
            allMatrixRowsWs.push({
              containerName: listTitleCur,
              containerType: containerTypeCur,
              itemPath: pathCur,
              inheritance: "Inherited",
              details: "",
              userOrGroup: pCur.title,
              principalType: pCur.principalType,
              accountName: pCur.loginName,
              givenThrough: givenBaseCur + (pCur.givenThroughSuffix || ""),
              roleSet: pCur.roleSet
            });
          }
        }
      }
      var headerWs = ["List/Library/Page name", "Type", "Item path", "Inheritance", "Details", "User/group", "Principal type", "Account name", "Given through"].concat(roleNamesWs);
      var aoaWs = [headerWs];
      for (var mr = 0; mr < allMatrixRowsWs.length; mr++) {
        var rowWs = allMatrixRowsWs[mr];
        var rWs = [rowWs.containerName, rowWs.containerType, rowWs.itemPath, rowWs.inheritance, rowWs.details, rowWs.userOrGroup, rowWs.principalType, rowWs.accountName, rowWs.givenThrough];
        for (var rn = 0; rn < roleNamesWs.length; rn++) rWs.push(rowWs.roleSet[roleNamesWs[rn]] ? "X" : "");
        aoaWs.push(rWs);
      }
      reportProgress("Permissions matrix: Building report…");
      var fnBaseWs = (exportFilename || "PermissionsMatrix").replace(/\.(csv|xls|xlsx|xml)$/i, "") + "_PermissionsMatrix_WholeSite";
      if (params.format === "xlsx") {
        downloadExport({ sheets: [{ name: "Permissions Matrix", aoa: aoaWs }] }, fnBaseWs, {});
      } else {
        var csvLinesWs = [headerWs.map(escapeCSV).join(",")];
        for (var c = 0; c < aoaWs.length - 1; c++) csvLinesWs.push(aoaWs[c + 1].map(escapeCSV).join(","));
        downloadExport({ csv: "\ufeff" + csvLinesWs.join("\r\n") }, fnBaseWs, { sheetName: "Permissions Matrix" });
      }
      reportDone(true, "Done! Permissions matrix (whole site) downloaded (" + allMatrixRowsWs.length + " rows).");
      return;
    }

    var listTitle = getListTitleFromContext();
    var listBase = listTitle ? listByTitleUrl(listTitle) : listBaseUrl();
    reportProgress("Permissions matrix: Loading list…");
    var listMeta = await fetch(listBase + "?$select=ItemCount,HasUniqueRoleAssignments", { credentials: "include", headers: accept });
    if (!listMeta.ok) {
      reportDone(false, "Could not load list: " + listMeta.status);
      return;
    }
    var listData = await listMeta.json();
    var totalItemCount = Math.max(0, parseInt(listData.ItemCount, 10) || 0);
    var listHasUnique = listData.HasUniqueRoleAssignments === true;
    reportProgress("Permissions matrix: Loading role definitions…");
    var roleDefResp = await fetch(siteUrl + "/_api/web/roledefinitions?$select=Name,Order&$filter=Hidden eq false&$orderby=Order asc", { credentials: "include", headers: accept });
    var roleNames = [];
    if (roleDefResp.ok) {
      var roleDefData = await roleDefResp.json();
      var defs = roleDefData.value || roleDefData.results || [];
      for (var rd = 0; rd < defs.length; rd++) {
        var name = (defs[rd].Name || "").trim();
        if (name && roleNames.indexOf(name) < 0) roleNames.push(name);
      }
    }
    if (roleNames.length === 0) roleNames = ["Full Control", "Design", "Contribute", "Edit", "Read", "View Only", "Restricted Read", "Approve", "Manage Hierarchy", "Restricted View", "Restricted Interfaces for Translation"];
    var inheritedSourcePath = "";
    try {
      if (listHasUnique) {
        var lrf = await fetch(listBase + "/rootfolder?$select=ServerRelativeUrl", { credentials: "include", headers: accept });
        if (lrf.ok) {
          var lrfData = await lrf.json();
          inheritedSourcePath = (lrfData.ServerRelativeUrl || lrfData.serverRelativeUrl || "").trim();
        }
      }
      if (!inheritedSourcePath) {
        var wu = await fetch(siteUrl + "/_api/web?$select=ServerRelativeUrl", { credentials: "include", headers: accept });
        if (wu.ok) {
          var wuData = await wu.json();
          inheritedSourcePath = (wuData.ServerRelativeUrl || wuData.serverRelativeUrl || "").trim();
        }
      }
      if (!inheritedSourcePath && typeof window._spPageContextInfo !== "undefined") {
        inheritedSourcePath = (window._spPageContextInfo.listUrl || window._spPageContextInfo.listServerRelativeUrl || window._spPageContextInfo.serverRequestPath || "").trim();
      }
      if (!inheritedSourcePath) inheritedSourcePath = listTitle || "list";
    } catch (_) {}
    reportProgress("Permissions matrix: Loading list/web role assignments (for inherited)…");
    var inheritedRaUrl = listHasUnique ? (listBase + "/roleassignments") : (siteUrl + "/_api/web/roleassignments");
    var inheritedRaResp = await fetch(inheritedRaUrl + "?$expand=Member,RoleDefinitionBindings&$top=5000", { credentials: "include", headers: accept });
    var inheritedPrincipals = [];
    if (inheritedRaResp.ok) {
      var inheritedRaData = await inheritedRaResp.json();
      var iraPage = inheritedRaData.value || inheritedRaData.results || [];
      var iraNext = inheritedRaData["@odata.nextLink"];
      var allRa = iraPage.slice();
      while (iraNext) {
        var iraR = await fetchWithRetry(iraNext, { credentials: "include", headers: accept }, 6);
        var iraD = await iraR.json();
        var iraBatch = iraD.value || iraD.results || [];
        allRa = allRa.concat(iraBatch);
        iraNext = iraD["@odata.nextLink"] || null;
      }
      for (var ira = 0; ira < allRa.length; ira++) {
        var ra = allRa[ira];
        var m = ra.Member || ra.member;
        if (!m) continue;
        var bindings = ra.RoleDefinitionBindings || ra.roleDefinitionBindings || [];
        var roleSet = {};
        for (var rb = 0; rb < bindings.length; rb++) {
          var rn = (bindings[rb].Name || bindings[rb].name || "").trim();
          if (rn) roleSet[rn] = true;
        }
        var pt = principalTypeLabel(m.PrincipalType);
        var tit = (m.Title || m.LoginName || "").trim();
        var log = (m.LoginName || "").trim();
        var isGrp = (parseInt(m.PrincipalType, 10) === 4 || parseInt(m.PrincipalType, 10) === 8);
        inheritedPrincipals.push({ title: tit, loginName: log, principalType: pt, givenThroughSuffix: "", roleSet: roleSet, isGroup: isGrp, groupId: isGrp ? (m.Id != null ? String(m.Id) : "") : "" });
      }
      var expandedInherited = [];
      for (var ei = 0; ei < inheritedPrincipals.length; ei++) {
        var ip = inheritedPrincipals[ei];
        if (!ip.isGroup) {
          expandedInherited.push({ title: ip.title, loginName: ip.loginName, principalType: ip.principalType, givenThroughSuffix: ip.givenThroughSuffix, roleSet: ip.roleSet });
          continue;
        }
        expandedInherited.push({ title: ip.title, loginName: ip.loginName, principalType: ip.principalType, givenThroughSuffix: ip.givenThroughSuffix, roleSet: ip.roleSet });
        if (!ip.groupId) continue;
        try {
          var guResp = await fetch(siteUrl + "/_api/web/sitegroups/GetById(" + ip.groupId + ")/users", { credentials: "include", headers: accept });
          if (!guResp.ok) continue;
          var guData = await guResp.json();
          var gUsers = guData.value || guData.results || [];
          for (var g = 0; g < gUsers.length; g++) {
            var u = gUsers[g];
            expandedInherited.push({
              title: (u.Title || u.LoginName || "").trim(),
              loginName: (u.LoginName || "").trim(),
              principalType: "User",
              givenThroughSuffix: " - " + (ip.title || "Group") + " Members",
              roleSet: ip.roleSet
            });
          }
        } catch (_) {}
      }
      inheritedPrincipals = expandedInherited;
    }
    var allItems = [];
    var nextUrl = listBase + "/items?$select=Id,FileRef,FileDirRef,FileLeafRef,FSObjType,HasUniqueRoleAssignments&$top=" + PAGE_SIZE;
    reportProgress("Permissions matrix: Scanning items…", { currentCount: 0, totalCount: totalItemCount });
    while (nextUrl) {
      var r = await fetchWithRetry(nextUrl, { credentials: "include", headers: acceptVerbose }, 6);
      var d = await r.json();
      var batch = (d.d && d.d.results) || [];
      allItems = allItems.concat(batch);
      nextUrl = (d.d && d.d.__next) || null;
      reportProgress("Permissions matrix: Scanning items…", { currentCount: allItems.length, totalCount: totalItemCount });
      if (batch.length < PAGE_SIZE) break;
      await sleep(100);
    }
    var exceptions = [];
    var inheritedItems = [];
    for (var ei = 0; ei < allItems.length; ei++) {
      var itm = allItems[ei];
      var itemLabel = itm.FileRef || (itm.FileDirRef && itm.FileLeafRef ? (itm.FileDirRef + "/" + itm.FileLeafRef).replace(/\/+/g, "/") : null) || itm.FileLeafRef || ("(ID " + itm.Id + ")");
      if (itemLabel && itemLabel.indexOf("#") >= 0) itemLabel = itemLabel.slice(itemLabel.indexOf("#") + 1).trim();
      if (itm.HasUniqueRoleAssignments === true) {
        exceptions.push({ Id: itm.Id, Item: itemLabel });
      } else {
        inheritedItems.push({ Id: itm.Id, Item: itemLabel });
      }
    }
    function normalizePath(p) {
      if (!p || typeof p !== "string") return "";
      return p.replace(/\\/g, "/").replace(/\/+/g, "/").replace(/\/$/, "").replace(/^\//, "").trim();
    }
    function getPathAncestors(path) {
      var n = normalizePath(path);
      if (!n) return [];
      var parts = n.split("/").filter(Boolean);
      if (parts.length <= 1) return [];
      var ancestors = [];
      for (var a = 1; a < parts.length; a++) {
        ancestors.push(parts.slice(0, a).join("/"));
      }
      return ancestors;
    }
    function getTopmostInheritanceSource(itemPath, exceptionPathSet, defaultRoot) {
      var ancestors = getPathAncestors(itemPath);
      for (var ai = 0; ai < ancestors.length; ai++) {
        var anc = ancestors[ai];
        if (exceptionPathSet[anc]) return "/" + anc;
      }
      return defaultRoot || inheritedSourcePath;
    }
    var exceptionPathSet = {};
    for (var exi = 0; exi < exceptions.length; exi++) {
      exceptionPathSet[normalizePath(exceptions[exi].Item)] = true;
    }
    async function getMatrixRowsForItem(itemId, itemPath) {
      var url = listBase + "/items(" + itemId + ")/RoleAssignments?$expand=Member,RoleDefinitionBindings&$top=5000";
      var rows = [];
      var next = url;
      while (next) {
        var resp = await fetchWithRetry(next, { credentials: "include", headers: accept }, 8);
        var data = await resp.json();
        var page = data.value || data.results || [];
        for (var pi = 0; pi < page.length; pi++) {
          var ra = page[pi];
          var m = ra.Member || ra.member;
          if (!m) continue;
          var bindings = ra.RoleDefinitionBindings || ra.roleDefinitionBindings || [];
          var roleSet = {};
          for (var bi = 0; bi < bindings.length; bi++) {
            var rn = (bindings[bi].Name || bindings[bi].name || "").trim();
            if (rn) roleSet[rn] = true;
          }
          var principalType = principalTypeLabel(m.PrincipalType);
          var title = (m.Title || m.LoginName || "").trim();
          var loginName = (m.LoginName || "").trim();
          var isGroup = (parseInt(m.PrincipalType, 10) === 4 || parseInt(m.PrincipalType, 10) === 8);
          rows.push({ title: title, loginName: loginName, principalType: principalType, givenThrough: "Explicit", roleSet: roleSet, isGroup: isGroup, groupId: isGroup ? (m.Id != null ? String(m.Id) : "") : "" });
        }
        next = data["@odata.nextLink"] || null;
      }
      var expanded = [];
      for (var ri = 0; ri < rows.length; ri++) {
        var row = rows[ri];
        if (!row.isGroup) {
          expanded.push({ title: row.title, loginName: row.loginName, principalType: row.principalType, givenThrough: row.givenThrough, roleSet: row.roleSet, viaGroup: false });
          continue;
        }
        expanded.push({ title: row.title, loginName: row.loginName, principalType: row.principalType, givenThrough: row.givenThrough, roleSet: row.roleSet, viaGroup: false });
        if (!row.groupId) continue;
        try {
          var usersResp = await fetch(siteUrl + "/_api/web/sitegroups/GetById(" + row.groupId + ")/users", { credentials: "include", headers: accept });
          if (!usersResp.ok) continue;
          var usersData = await usersResp.json();
          var groupUsers = usersData.value || usersData.results || [];
          for (var gu = 0; gu < groupUsers.length; gu++) {
            var u = groupUsers[gu];
            expanded.push({
              title: (u.Title || u.LoginName || "").trim(),
              loginName: (u.LoginName || "").trim(),
              principalType: "User",
              givenThrough: (row.title || "Group") + " Members",
              roleSet: row.roleSet,
              viaGroup: true
            });
          }
        } catch (_) {}
      }
      return expanded.map(function (e) {
        return {
          itemPath: itemPath,
          inheritance: e.viaGroup ? "Inherited" : "Custom",
          details: "",
          userOrGroup: e.title,
          principalType: e.principalType,
          accountName: e.loginName,
          givenThrough: e.givenThrough,
          roleSet: e.roleSet
        };
      });
    }
    var matrixRows = [];
    var resultIdx = 0;
    var completedCount = 0;
    reportProgress("Permissions matrix: Loading permissions…", { currentCount: 0, totalCount: exceptions.length });
    function runWorker() {
      return new Promise(function (resolve) {
        function work() {
          var i = resultIdx++;
          if (i >= exceptions.length) { resolve(); return; }
          var ex = exceptions[i];
          getMatrixRowsForItem(ex.Id, ex.Item).then(function (itemRows) {
            for (var ir = 0; ir < itemRows.length; ir++) matrixRows.push(itemRows[ir]);
            completedCount++;
            reportProgress("Permissions matrix: Loading permissions…", { currentCount: completedCount, totalCount: exceptions.length });
            work();
          }).catch(function () { completedCount++; work(); });
        }
        work();
      });
    }
    var workers = [];
    for (var w = 0; w < CONCURRENCY; w++) workers.push(runWorker());
    await Promise.all(workers);
    reportProgress("Permissions matrix: Adding inherited items…");
    for (var ii = 0; ii < inheritedItems.length; ii++) {
      var path = inheritedItems[ii].Item;
      var sourcePath = getTopmostInheritanceSource(path, exceptionPathSet, inheritedSourcePath);
      var givenThroughBase = "Inherited (From ItemPath: " + sourcePath + ")";
      for (var ip = 0; ip < inheritedPrincipals.length; ip++) {
        var p = inheritedPrincipals[ip];
        matrixRows.push({
          itemPath: path,
          inheritance: "Inherited",
          details: "",
          userOrGroup: p.title,
          principalType: p.principalType,
          accountName: p.loginName,
          givenThrough: givenThroughBase + (p.givenThroughSuffix || ""),
          roleSet: p.roleSet
        });
      }
    }
    var header = ["Item path", "Inheritance", "Details", "User/group", "Principal type", "Account name", "Given through"].concat(roleNames);
    var aoa = [header];
    for (var mr = 0; mr < matrixRows.length; mr++) {
      var row = matrixRows[mr];
      var r = [row.itemPath, row.inheritance, row.details, row.userOrGroup, row.principalType, row.accountName, row.givenThrough];
      for (var rn = 0; rn < roleNames.length; rn++) {
        r.push(row.roleSet[roleNames[rn]] ? "X" : "");
      }
      aoa.push(r);
    }
    reportProgress("Permissions matrix: Building report…");
    var fnBase = (exportFilename || "PermissionsMatrix").replace(/\.(csv|xls|xlsx|xml)$/i, "") + "_PermissionsMatrix";
    if (params.format === "xlsx") {
      downloadExport({ sheets: [{ name: "Permissions Matrix", aoa: aoa }] }, fnBase, {});
    } else {
      var csvLines = [header.map(escapeCSV).join(",")];
      for (var c = 0; c < aoa.length - 1; c++) csvLines.push(aoa[c + 1].map(escapeCSV).join(","));
      downloadExport({ csv: "\ufeff" + csvLines.join("\r\n") }, fnBase, { sheetName: "Permissions Matrix" });
    }
    reportDone(true, "Done! Permissions matrix downloaded (" + matrixRows.length + " rows).");
  }

  // -----------------------------
  // Main
  // -----------------------------
  (async function () {
    var rpcViewId = "";
    try {
      await waitForPageReady(6000);

      if (!siteUrl) {
        reportDone(false, "Could not detect site. Ensure you are on a SharePoint list/library page.");
        return;
      }
      if (!listId && window._spPageContextInfo) {
        var resolved = await resolveListIdFromListUrl();
        if (resolved) listId = resolved;
      }
      if (!listId && !params.matrixWholeSite) {
        reportDone(false, "Could not detect list. Ensure you are on a SharePoint list/library view page (URL should have List= or pageListId), or use Permissions matrix with Include whole site.");
        return;
      }

      listBaseTemplate = listId ? await getListBaseTemplate() : null;
      if (params.report === "permissions") {
        await exportPermissionsReport();
        return;
      }
      if (params.report === "permissionsMatrix") {
        reportDone(false, "Permissions matrix now runs via permissionsMatrixExport.js. Reload the extension and try again.");
        return;
      }
      if ((params.report === "folderCount" || params.report === "pathLengths") && listBaseTemplate !== 101) {
        reportDone(false, "Folder count and path lengths reports are only available for document libraries.");
        return;
      }
      var itemCount = await getItemCount();

      var isDocLib = listBaseTemplate === 101;
      var useRestForVersions = !isDocLib && params.incVer !== false;
      if (useRestForVersions) {
        await exportViaRest(itemCount);
      } else {
        var vr = await getOrCreateRpcViewForDownload();
        if (!vr.ok) {
          reportDone(false, "RPC view creation failed: " + (vr.error || "unknown"));
          return;
        }
        rpcViewId = vr.viewId;
        try {
          await sleep(500);
          await exportViaOwssvrWithView(rpcViewId, itemCount);
        } finally {
          if (DELETE_VIEW_AT_END && rpcViewId) {
            try { await deleteViewById(rpcViewId); } catch (_) {}
          }
        }
      }
    } catch (e) {
      console.error("SharePoint CSV Export – Error:", e);
      reportDone(false, "Error: " + (e.message || String(e)));
    }
  })();
})();
