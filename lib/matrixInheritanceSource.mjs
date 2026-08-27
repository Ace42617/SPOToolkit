/**
 * Resolve the SharePoint securable object that an inheriting item actually
 * draws its ACL from. Keep in sync with the inlined copies in
 * permissionsMatrixExport.js (injected classic page scripts cannot import
 * from lib/) and with Get-InheritanceSource in Export-SitePermissionsMatrix.ps1.
 */

export function normalizeServerRelativePath(path) {
  if (path == null) return "";
  var s = String(path).replace(/\\/g, "/").trim();
  if (!s) return "";
  s = s.replace(/\/+/g, "/");
  if (s.length > 1) s = s.replace(/\/$/, "");
  return s;
}

export function parentServerRelativePath(path) {
  var n = normalizeServerRelativePath(path);
  if (!n || n === "/") return "";
  var i = n.lastIndexOf("/");
  if (i <= 0) return "";
  return n.slice(0, i);
}

function uniquePathLookup(uniquePermPaths) {
  var lookup = Object.create(null);
  var list;
  if (Array.isArray(uniquePermPaths)) {
    list = uniquePermPaths;
  } else if (uniquePermPaths && typeof uniquePermPaths === "object") {
    list = Object.keys(uniquePermPaths).filter(function (k) {
      return uniquePermPaths[k];
    });
  } else {
    list = [];
  }
  for (var i = 0; i < list.length; i++) {
    var u = normalizeServerRelativePath(list[i]);
    if (u) lookup[u.toLowerCase()] = u;
  }
  return lookup;
}

/**
 * Nearest uniquely permissioned ancestor path, else the list root.
 * Walks parents of itemPath (not the item itself), matching the PowerShell
 * Get-InheritanceSource helper.
 */
export function getInheritanceSourcePath(itemPath, uniquePermPaths, listRootUrl) {
  var root = normalizeServerRelativePath(listRootUrl);
  var lookup = uniquePathLookup(uniquePermPaths);
  var parent = parentServerRelativePath(itemPath);
  while (parent && (!root || parent.length >= root.length)) {
    var hit = lookup[parent.toLowerCase()];
    if (hit) return hit;
    if (root && parent.toLowerCase() === root.toLowerCase()) break;
    parent = parentServerRelativePath(parent);
  }
  return root;
}

/**
 * Which RoleAssignments collection to read for an inherited item.
 * Matches Get-ItemRows in Export-SitePermissionsMatrix.ps1:
 * web when the source is the list root of a list that still inherits,
 * item when a uniquely permissioned folder/file ancestor has an item id,
 * otherwise the list.
 */
export function inheritedRoleAssignmentKind(sourcePath, listRootUrl, listHasUnique, sourceItemId) {
  var src = normalizeServerRelativePath(sourcePath);
  var root = normalizeServerRelativePath(listRootUrl);
  var atRoot = !!src && !!root && src.toLowerCase() === root.toLowerCase();
  var srcId =
    sourceItemId != null && String(sourceItemId).trim() !== "" && String(sourceItemId) !== "0"
      ? sourceItemId
      : 0;
  if (atRoot) srcId = 0;
  if (atRoot && !listHasUnique && !srcId) return "web";
  if (!atRoot && srcId) return "item";
  return "list";
}

export function inheritedRoleAssignmentUrl(kind, listBase, webUrl, sourceItemId) {
  if (kind === "web") return String(webUrl || "").replace(/\/$/, "") + "/_api/web/roleassignments";
  if (kind === "item") return String(listBase || "") + "/items(" + sourceItemId + ")/RoleAssignments";
  return String(listBase || "") + "/roleassignments";
}
