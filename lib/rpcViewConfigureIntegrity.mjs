/**
 * Guards for temporary RPC view field setup during OWSSVR CSV/XLSX export.
 * Keep behavioral decisions in sync with exportCSV.js (injected classic page
 * script — cannot import from lib/).
 */

const SETUP_FAILED_FALLBACK =
  "View field setup failed. No file downloaded. Try refreshing the page and run the export again.";

/**
 * Map setViewFieldsBatch's result onto tryConfigureRpcView's return value.
 *
 * Field-setup failure must fail closed. Returning { ok: true, fieldsConfigured: false }
 * lets getOrCreateRpcViewForDownload ignore the result and still OWSSVR-export a
 * cleared or sparse view as a successful download.
 *
 * @param {{ ok?: boolean, error?: string } | null | undefined} setViewFieldsResult
 * @returns {{ ok: true, fieldsConfigured: true } | { ok: false, fieldsConfigured: false, error: string }}
 */
export function interpretTryConfigureRpcView(setViewFieldsResult) {
  if (setViewFieldsResult && setViewFieldsResult.ok === true) {
    return { ok: true, fieldsConfigured: true };
  }
  var error =
    setViewFieldsResult &&
    typeof setViewFieldsResult.error === "string" &&
    setViewFieldsResult.error
      ? setViewFieldsResult.error
      : SETUP_FAILED_FALLBACK;
  return { ok: false, fieldsConfigured: false, error: error };
}

/**
 * After tryConfigureRpcView on an owned RPC view, either fail closed or return
 * the success payload used for OWSSVR export.
 *
 * @param {{ ok?: boolean, error?: string } | null | undefined} configureResult
 * @param {object} successPayload
 * @returns {object}
 */
export function ownedRpcViewAfterConfigure(configureResult, successPayload) {
  if (!configureResult || configureResult.ok !== true) {
    return {
      ok: false,
      error:
        (configureResult && configureResult.error) || SETUP_FAILED_FALLBACK
    };
  }
  return successPayload;
}
