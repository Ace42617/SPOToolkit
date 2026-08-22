/**
 * Integrity helpers for PowerShell sharing-links export
 * (scripts/Export-SiteSharingLinks.ps1).
 *
 * The .ps1 cannot import this module; keep message/decision wording in sync
 * with the REST enumeration fail-closed throw + $scanCompletedOk CSV guard.
 */

/**
 * User-facing message when a list/library item scan fails mid-export.
 * Callers must not write a CSV after this failure.
 * @param {string} listTitle
 * @param {unknown} err
 * @returns {string}
 */
export function formatPsSharingLinksListScanFailureMessage(listTitle, err) {
  const title = String(listTitle || "list").trim() || "list";
  let detail = "";
  if (err && typeof err === "object" && "message" in err && err.message) {
    detail = String(err.message);
  } else {
    detail = String(err || "unknown error");
  }
  if (!detail.trim()) detail = "unknown error";
  return (
    "Sharing-links export failed while scanning list '" +
    title +
    "': " +
    detail +
    ". No CSV downloaded. Resolve access or throttling, then retry."
  );
}

/**
 * Terminal CSV export is only allowed after a fully successful scan.
 * @param {{ scanCompletedOk: boolean, itemScanFailures?: unknown[] }} state
 * @returns {boolean}
 */
export function canExportPsSharingLinksCsv(state) {
  const s = state || {};
  if (!s.scanCompletedOk) return false;
  if (Array.isArray(s.itemScanFailures) && s.itemScanFailures.length > 0) {
    return false;
  }
  return true;
}
