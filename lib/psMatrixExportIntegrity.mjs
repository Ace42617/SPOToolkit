/**
 * Integrity helpers for PowerShell Permissions Matrix export
 * (scripts/Export-SitePermissionsMatrix.ps1).
 *
 * The .ps1 cannot import this module; keep message/decision wording in sync
 * with Scan-ListItems fail-closed throw + $scanCompletedOk export guard.
 */

/**
 * User-facing message when a list/library item scan fails mid-export.
 * Callers must not write a workbook after this failure.
 * @param {string} listTitle
 * @param {string} listUrl
 * @param {unknown} err
 * @returns {string}
 */
export function formatPsMatrixListScanFailureMessage(listTitle, listUrl, err) {
  const title = String(listTitle || "list").trim() || "list";
  const url = String(listUrl || "").trim();
  let detail = "";
  if (err && typeof err === "object" && "message" in err && err.message) {
    detail = String(err.message);
  } else {
    detail = String(err || "unknown error");
  }
  if (!detail.trim()) detail = "unknown error";
  return (
    "Permissions matrix failed while scanning list '" +
    title +
    "' (" +
    url +
    "): " +
    detail +
    ". No workbook downloaded. Resolve access or throttling, then retry."
  );
}

/**
 * Terminal workbook export is only allowed after a fully successful scan.
 * @param {{ scanCompletedOk: boolean, itemScanFailures?: unknown[] }} state
 * @returns {boolean}
 */
export function canExportPsMatrixWorkbook(state) {
  const s = state || {};
  if (!s.scanCompletedOk) return false;
  if (Array.isArray(s.itemScanFailures) && s.itemScanFailures.length > 0) {
    return false;
  }
  return true;
}
