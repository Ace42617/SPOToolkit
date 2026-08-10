/**
 * Classic SharePoint lookup / FieldUserValue display strings use "ID;#Value".
 * Keep in sync with the inlined copies in exportCSV.js and permissionsMatrixExport.js
 * (injected classic page scripts cannot import from lib/).
 */

/**
 * Strip a leading classic lookup prefix ("123;#…") only.
 * Do not strip at an arbitrary "#" — SharePoint Online allows "#" in
 * file and folder names, so FileRef/FileLeafRef may contain "#" mid-path.
 *
 * @param {unknown} s
 * @returns {string}
 */
export function stripIdHashPrefix(s) {
  if (!s || typeof s !== "string") return s || "";
  var m = /^(\d+;#)/.exec(s);
  return m ? s.slice(m[1].length).trim() : s;
}
