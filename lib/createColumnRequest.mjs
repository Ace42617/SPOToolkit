/**
 * Request-correlation helpers for Compass/popup column create.
 * Injected classic page scripts cannot import from lib/ — keep inline copies in
 * content.js / createColumn.js in sync via source-sync asserts in the companion test.
 */

export const COLUMN_CREATE_PARAMS_SCRIPT_ID = "sp-column-create-params";
export const COLUMN_CREATE_REQUEST_QUERY = "spcsvRequestId";
export const COLUMN_CREATE_RESULT_TYPE = "SPCSVCreateColumnResult";

/**
 * DOM id for the JSON params script node for one createColumn request.
 * Unique ids prevent overlapping Add/Create clicks from overwriting each other's SchemaXml.
 * @param {string|null|undefined} requestId
 * @returns {string}
 */
export function getColumnCreateParamsScriptId(requestId) {
  return requestId
    ? COLUMN_CREATE_PARAMS_SCRIPT_ID + "-" + requestId
    : COLUMN_CREATE_PARAMS_SCRIPT_ID;
}

/**
 * Whether a page postMessage belongs to the outstanding createColumn wait.
 * Untagged or mismatched results must be ignored so a late/overlapping create
 * cannot resolve the wrong caller (wrong field schema / success for a different column).
 * @param {any} data
 * @param {string} requestId
 * @returns {boolean}
 */
export function matchesColumnCreateResponse(data, requestId) {
  if (!data || data.__spcsv !== true || data.type !== COLUMN_CREATE_RESULT_TYPE) {
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
export function extractColumnCreateRequestIdFromScriptSrc(src, baseHref) {
  try {
    const href = String(src || "");
    if (!href) return "";
    return (
      new URL(href, baseHref || "https://example.invalid/").searchParams.get(
        COLUMN_CREATE_REQUEST_QUERY
      ) || ""
    );
  } catch (_) {
    return "";
  }
}
