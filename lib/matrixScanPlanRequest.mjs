/**
 * Request-correlation helpers for Permissions Matrix "Load sites" scan plans.
 * Injected classic page scripts cannot import from lib/ — keep inline copies in
 * content.js / matrixScanPlan.js in sync via source-sync asserts in the companion test.
 */

export const MATRIX_SCAN_PARAMS_SCRIPT_ID = "sp-matrix-scan-params-json";
export const MATRIX_SCAN_REQUEST_QUERY = "spcsvRequestId";
export const MATRIX_SCAN_RESULT_TYPE = "SPMatrixScanPlanResult";

/**
 * DOM id for the JSON params script node for one Load-sites request.
 * Unique ids prevent overlapping Load-sites calls from overwriting each other's params.
 * @param {string|null|undefined} requestId
 * @returns {string}
 */
export function getMatrixScanParamsScriptId(requestId) {
  return requestId
    ? MATRIX_SCAN_PARAMS_SCRIPT_ID + "-" + requestId
    : MATRIX_SCAN_PARAMS_SCRIPT_ID;
}

/**
 * Whether a page postMessage belongs to the outstanding Load-sites wait.
 * Untagged or mismatched results must be ignored so a late/overlapping plan
 * cannot resolve the wrong caller (wrong matrixSelectedPaths scope).
 * @param {any} data
 * @param {string} requestId
 * @returns {boolean}
 */
export function matchesMatrixScanPlanResponse(data, requestId) {
  if (!data || data.__spcsv !== true || data.type !== MATRIX_SCAN_RESULT_TYPE) {
    return false;
  }
  if (!requestId) return false;
  return data.requestId === requestId;
}

/**
 * Parse spcsvRequestId from an injected script URL.
 * @param {string} src
 * @param {string} [baseHref]
 * @returns {string}
 */
export function extractMatrixScanRequestIdFromScriptSrc(src, baseHref) {
  try {
    const href = String(src || "");
    if (!href) return "";
    return (
      new URL(href, baseHref || "https://example.invalid/").searchParams.get(
        MATRIX_SCAN_REQUEST_QUERY
      ) || ""
    );
  } catch (_) {
    return "";
  }
}
