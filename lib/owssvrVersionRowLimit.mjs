/**
 * OWSSVR RowLimit helpers for version-expanded document library exports.
 * Keep behavioral decisions in sync with exportCSV.js variants
 * (injected classic page scripts — cannot import from lib/).
 */

/**
 * Match exportCSV.js buildOwssvrUrl RowLimit selection.
 * Lists with versions use RowLimit=0; doc libs keep a finite cap.
 * @param {number|string} pageLimit
 * @param {boolean} includeVersions
 * @param {boolean} isDocLib
 * @returns {number}
 */
export function computeOwssvrRowLimit(pageLimit, includeVersions, isDocLib) {
  const pl = Math.max(1, parseInt(pageLimit, 10) || 1000);
  if (includeVersions && !isDocLib) return 0;
  return Math.max(pl * 20, 50000);
}

/**
 * Safer ID-range page size for versioned doc-lib OWSSVR pages.
 * Leaves ~100 versions/item of headroom under the RowLimit.
 * @param {number|string} pageLimit
 * @param {number} rowLimit
 * @returns {number}
 */
export function versionedDocLibIdPageLimit(pageLimit, rowLimit) {
  const pl = Math.max(1, parseInt(pageLimit, 10) || 1000);
  const rl = Number(rowLimit);
  if (!(rl > 0)) return pl;
  const capped = Math.max(50, Math.floor(rl / 100));
  return Math.min(pl, capped);
}

/**
 * True when a version-expanded doc-lib OWSSVR response is at/over RowLimit,
 * which SharePoint uses as a hard cap (older version rows are dropped).
 * @param {{ includeVersions: boolean, isDocLib: boolean, rowCount: number, rowLimit: number }} opts
 * @returns {boolean}
 */
export function isOwssvrVersionPageTruncated(opts) {
  const includeVersions = !!(opts && opts.includeVersions);
  const isDocLib = !!(opts && opts.isDocLib);
  if (!includeVersions || !isDocLib) return false;
  const limit = Number(opts && opts.rowLimit);
  const count = Number(opts && opts.rowCount);
  if (!(limit > 0) || !(count >= 0)) return false;
  return count >= limit;
}

/**
 * User-facing message when version expansion saturated RowLimit.
 * Callers must not download a file after this failure.
 * @param {number} rowLimit
 * @param {number} pageLimit
 * @returns {string}
 */
export function formatOwssvrVersionTruncationMessage(rowLimit, pageLimit) {
  return (
    "Version history export hit the OWSSVR RowLimit (" +
    String(rowLimit) +
    " rows) for an ID page of " +
    String(pageLimit) +
    " items. Older versions were likely omitted. No file downloaded. " +
    "Retry with a smaller page size, or export without versions."
  );
}
