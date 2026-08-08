/**
 * Request-correlation helpers for Reports "Load lists" inventory plans.
 * Injected classic page scripts cannot import from lib/ — keep inline copies in
 * content.js / reportListPlan.js in sync via source-sync asserts in the companion test.
 */

export const REPORT_LIST_PLAN_PARAMS_SCRIPT_ID = "sp-report-list-plan-params-json";
export const REPORT_LIST_PLAN_REQUEST_QUERY = "spcsvRequestId";
export const REPORT_LIST_PLAN_RESULT_TYPE = "SPReportListPlanResult";

/**
 * DOM id for the JSON params script node for one Load-lists request.
 * Unique ids prevent overlapping Load-lists calls from overwriting each other's params.
 * @param {string|null|undefined} requestId
 * @returns {string}
 */
export function getReportListPlanParamsScriptId(requestId) {
  return requestId
    ? REPORT_LIST_PLAN_PARAMS_SCRIPT_ID + "-" + requestId
    : REPORT_LIST_PLAN_PARAMS_SCRIPT_ID;
}

/**
 * Whether a page postMessage belongs to the outstanding Load-lists wait.
 * Untagged or mismatched results must be ignored so a late/overlapping plan
 * cannot resolve the wrong caller (wrong reportSelectedLists / matrix scope).
 * @param {any} data
 * @param {string} requestId
 * @returns {boolean}
 */
export function matchesReportListPlanResponse(data, requestId) {
  if (!data || data.__spcsv !== true || data.type !== REPORT_LIST_PLAN_RESULT_TYPE) {
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
export function extractReportListPlanRequestIdFromScriptSrc(src, baseHref) {
  try {
    const href = String(src || "");
    if (!href) return "";
    return (
      new URL(href, baseHref || "https://example.invalid/").searchParams.get(
        REPORT_LIST_PLAN_REQUEST_QUERY
      ) || ""
    );
  } catch (_) {
    return "";
  }
}
