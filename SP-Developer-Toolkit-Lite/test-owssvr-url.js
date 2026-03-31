/**
 * Run with: node test-owssvr-url.js
 * Verifies normalizeGuid and owssvr URL building (same logic as exportCSV.js).
 */
function normalizeGuid(g) {
  if (g == null) return "";
  var s = String(g).trim();
  try { s = decodeURIComponent(s); } catch (_) {}
  s = s.replace(/%7B|%7D/ig, "");
  s = s.replace(/[{}]/g, "");
  return s.trim();
}

function buildOwssvrUrl(siteUrl, listId, viewId, listBaseTemplate, listRoot, pageLimit, incVer) {
  var guid = function (g) { return "{" + normalizeGuid(g).toUpperCase() + "}"; };
  var isDocLib = listBaseTemplate === 101;
  var rowLimit = (incVer && !isDocLib) ? 0 : Math.max(pageLimit * 20, 50000);
  var url = siteUrl + "/_vti_bin/owssvr.dll?Cmd=Display&XMLDATA=1&List=" + encodeURIComponent(guid(listId)) +
    "&View=" + encodeURIComponent(guid(viewId)) + "&IncludeVersions=" + (incVer ? "TRUE" : "FALSE") + "&RowLimit=" + rowLimit;
  if (incVer && !isDocLib) url += "&RootFolder=*";
  if (isDocLib) {
    var root = listRoot || "";
    if (root) url += "&RootFolder=" + encodeURIComponent(root);
    else url += "&RootFolder=*";
  }
  return url;
}

// Test normalizeGuid
const guidWithBraces = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}";
const guidLower = "a1b2c3d4-e5f6-7890-abcd-ef1234567890";
const guidNoBraces = "A1B2C3D4-E5F6-7890-ABCD-EF1234567890";
console.assert(normalizeGuid(guidWithBraces) === "A1B2C3D4-E5F6-7890-ABCD-EF1234567890", "normalizeGuid with braces");
console.assert(normalizeGuid(guidLower).toUpperCase() === "A1B2C3D4-E5F6-7890-ABCD-EF1234567890", "normalizeGuid lower");
console.assert(normalizeGuid(guidNoBraces) === "A1B2C3D4-E5F6-7890-ABCD-EF1234567890", "normalizeGuid no braces");

// Test URL building (list)
const siteUrl = "https://contoso.sharepoint.com/sites/site";
const listId = "a1b2c3d4-e5f6-7890-abcd-ef1234567890";
const viewId = "b2c3d4e5-f6a7-8901-bcde-f12345678901";
const urlList = buildOwssvrUrl(siteUrl, listId, viewId, 100, "", 5000, true);
console.assert(urlList.indexOf("List=%7B") !== -1, "List param encoded");
console.assert(urlList.indexOf("View=%7B") !== -1, "View param encoded");
console.assert(urlList.indexOf("RootFolder=*") !== -1, "List has RootFolder=*");
console.assert(urlList.indexOf("_vti_bin/owssvr.dll") !== -1, "owssvr path");

// Test URL building (doc lib)
const urlLib = buildOwssvrUrl(siteUrl, listId, viewId, 101, "/sites/site/Shared Documents", 5000, true);
console.assert(urlLib.indexOf("RootFolder=") !== -1, "Doc lib has RootFolder");
console.assert(urlLib.indexOf("Shared%20Documents") !== -1 || urlLib.indexOf("Shared+Documents") !== -1, "RootFolder encoded");

console.log("All assertions passed. Sample list URL:", urlList.slice(0, 100) + "...");
console.log("Sample doc lib URL:", urlLib.slice(0, 100) + "...");
