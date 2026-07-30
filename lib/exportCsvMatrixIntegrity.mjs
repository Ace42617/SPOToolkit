/**
 * Integrity helpers for Permissions Matrix export in exportCSV.js variants
 * (Lite/Experimental primary path; root retains a parallel implementation).
 * Injected classic page scripts cannot import from lib/ — keep inline copies
 * in sync via source-sync asserts in the companion test.
 */

/**
 * User-facing message when /_api/web/roledefinitions cannot be loaded.
 * Callers must not download a workbook after this failure.
 * @param {string|number} detail
 * @returns {string}
 */
export function formatMatrixRoleDefsLoadFailureMessage(detail) {
  const d =
    detail != null && String(detail).trim() !== ""
      ? String(detail).trim()
      : "unknown error";
  return "Failed to load role definitions: " + d + ". No file downloaded.";
}

/**
 * User-facing message when roledefinitions returns no usable names.
 * Never substitute a hard-coded preferred role list (that silently drops
 * custom role grants from the matrix columns).
 * @returns {string}
 */
export function formatMatrixRoleDefsEmptyMessage() {
  return "No role definitions returned for this site. No file downloaded.";
}

/**
 * User-facing message when a uniquely permissioned item's RoleAssignments
 * fetch fails mid-export. Callers must not download a workbook afterward.
 * @param {string} itemLabel
 * @param {unknown} err
 * @returns {string}
 */
export function formatMatrixItemScanFailureMessage(itemLabel, err) {
  const label = String(itemLabel || "item").trim() || "item";
  const detail =
    err && typeof err === "object" && "message" in err && err.message
      ? String(err.message)
      : String(err || "unknown error");
  return (
    'Permissions matrix failed while loading permissions for "' +
    label +
    '": ' +
    detail +
    ". No file downloaded. Resolve access or throttling, then retry."
  );
}

/**
 * Terminal success is only allowed when role names loaded and no item
 * RoleAssignments scan errors occurred.
 * @param {{ roleNamesOk: boolean, itemScanError?: unknown }} opts
 * @returns {boolean}
 */
export function canReportExportCsvMatrixSuccess({ roleNamesOk, itemScanError }) {
  if (!roleNamesOk) return false;
  if (itemScanError) return false;
  return true;
}
