/**
 * Path Lengths reports must strip the document library root, not the parent
 * of the shortest file path. Prefer RootFolder/ServerRelativeUrl; fall back to
 * the shortest observed parent/FileDirRef only when the root is unavailable.
 *
 * Keep the relative-path decision logic in sync with the inlined copies in
 * exportCSV.js and permissionsMatrixExport.js (injected classic page scripts
 * cannot import from lib/).
 */

/**
 * @param {unknown} p
 * @returns {string}
 */
export function normalizeServerRelativePath(p) {
  if (!p || typeof p !== "string") return "";
  var s = p.replace(/\/+/g, "/");
  if (s.length > 1 && s.charAt(s.length - 1) === "/") s = s.slice(0, -1);
  return s || "";
}

/**
 * @param {unknown} listRootUrl
 * @param {Iterable<unknown>} [parentPaths]
 * @returns {string}
 */
export function resolvePathLengthLibraryRoot(listRootUrl, parentPaths) {
  var explicit = normalizeServerRelativePath(listRootUrl);
  if (explicit) return explicit;
  var shortest = "";
  if (parentPaths) {
    for (var p of parentPaths) {
      var n = normalizeServerRelativePath(p);
      if (!n || n === "/") continue;
      if (!shortest || n.length < shortest.length) shortest = n;
    }
  }
  return shortest || "/";
}

/**
 * @param {unknown} fullPath
 * @param {unknown} libraryRoot
 * @returns {string}
 */
export function pathRelativeToLibraryRoot(fullPath, libraryRoot) {
  var path = normalizeServerRelativePath(fullPath);
  var normRoot = normalizeServerRelativePath(libraryRoot) || "/";
  if (!path) return "";
  if (normRoot === "/") return path.replace(/^\/+/, "") || "";
  if (path === normRoot) return "";
  if (path.indexOf(normRoot + "/") === 0) {
    return path.slice(normRoot.length).replace(/^\/+/, "") || "";
  }
  // Path outside the known library root — keep full server-relative path.
  return path;
}
