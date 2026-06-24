// Injected: fetches list fields for column picker. Uses SharePoint system field names for defaults (not list-specific).
(function () {
  var spi = window._spPageContextInfo || {};
  var siteUrl = (spi.webAbsoluteUrl || "").replace(/\/$/, "");
  var listUrl = spi.listUrl || spi.listServerRelativeUrl;
  if (!listUrl) {
    window.postMessage({ __spcsv: true, type: "SPCSVFieldsResult", error: "Not on list page" }, "*");
    return;
  }
  var listId;
  var enc = encodeURIComponent("'" + listUrl.replace(/'/g, "''") + "'");
  fetch(siteUrl + "/_api/web/GetList(@u)/Id?@u=" + enc, {
    credentials: "include",
    headers: { "Accept": "application/json;odata=nometadata" }
  })
    .then(function (r) { return r.json(); })
    .then(function (j) {
      listId = (j.value || "").replace(/[{}]/g, "");
      if (!listId) throw new Error("Could not get list ID");
      var lb = siteUrl + "/_api/web/lists(guid'" + listId + "')";
      return Promise.all([
        fetch(lb + "/DefaultView/ViewFields", { credentials: "include", headers: { "Accept": "application/json;odata=nometadata" } }).then(function (x) { return x.json(); }),
        fetch(lb + "/fields?$select=InternalName,Title,Hidden&$orderby=Title", { credentials: "include", headers: { "Accept": "application/json;odata=nometadata" } }).then(function (x) { return x.json(); })
      ]);
    })
    .then(function (results) {
      var vfRaw = results[0].Items || results[0].value || results[0].d?.results || [];
      var allFields = results[1].value || results[1].d?.results || [];
      var sysExclude = {
        LinkFilename: 1, LinkFilenameNoMenu: 1, LinkTitle: 1, LinkTitleNoMenu: 1,
        VirusStatus: 1, VirusVendorID: 1, VirusInfo: 1, VirusInfoEx: 1,
        SMTotalSize: 1, SMLastModifiedDate: 1, InstanceID: 1, SyncClientId: 1,
        ProgId: 1, ScopeId: 1, PermMask: 1, UniqueId: 1, WorkflowVersion: 1,
        WorkflowInstanceID: 1, FormData: 1, RestrictType: 1, NoCrawl: 1,
        IsCurrentVersion: 1,
        Attachments: 1, ContentTypeId: 1, ComplianceAssetId: 1, MetaInfo: 1,
        Edit: 1, AppAuthor: 1, AppEditor: 1, MediaServiceImageTags: 1, MediaServiceAutoTags: 1
      };
      function isSystemField(n) {
        if (sysExclude[n]) return true;
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
      var defaultNames = [];
      var seen = {};
      for (var i = 0; i < vfRaw.length; i++) {
        var item = vfRaw[i];
        var name = (typeof item === "string") ? item : (item && (item.Name || item.InternalName || item));
        name = String(name || "").trim();
        if (name && !isSystemField(name)) {
          var k = normKey(name);
          if (!seen[k]) { seen[k] = true; defaultNames.push(canonicalName(name)); }
        }
      }
      // System field names (same on all lists/libraries)
      var required = ["ID", "Title", "FileLeafRef", "FSObjType", "Created", "Modified", "Author", "Editor", "_UIVersionString"];
      for (var r = 0; r < required.length; r++) {
        if (!seen[required[r]]) defaultNames.push(required[r]);
      }
      var fieldMap = {};
      var pickerFields = [];
      for (var f = 0; f < allFields.length; f++) {
        var fn = allFields[f];
        var iname = fn.InternalName || fn.Title;
        if (iname && !fn.Hidden && !isSystemField(iname)) {
          fieldMap[iname] = fn.Title || iname;
          pickerFields.push({ InternalName: iname, Title: fn.Title || iname });
        }
      }
      for (var p = 0; p < defaultNames.length; p++) {
        if (!fieldMap[defaultNames[p]]) {
          pickerFields.push({ InternalName: defaultNames[p], Title: defaultNames[p] });
        }
      }
      var pathExtras = ["FileRef", "FileDirRef"];
      for (var pe = 0; pe < pathExtras.length; pe++) {
        var px = pathExtras[pe];
        if (fieldMap[px]) continue;
        for (var ai = 0; ai < allFields.length; ai++) {
          var af = allFields[ai];
          var ain = af.InternalName || af.Title;
          if (ain === px) {
            fieldMap[px] = af.Title || px;
            pickerFields.push({ InternalName: px, Title: af.Title || px });
            break;
          }
        }
      }
      window.postMessage({
        __spcsv: true,
        type: "SPCSVFieldsResult",
        defaultColumns: defaultNames,
        pickerFields: pickerFields,
        fieldTitles: fieldMap
      }, "*");
    })
    .catch(function (err) {
      window.postMessage({ __spcsv: true, type: "SPCSVFieldsResult", error: String(err && err.message || err) }, "*");
    });
})();
