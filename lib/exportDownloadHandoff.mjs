/**
 * Page → content-script export download handoff.
 *
 * Page scripts post:
 *   { __spcsv: true, type: "SPCSVExportDownload*", detail: { requestId, ... } }
 *
 * content.js unwraps once: `const d = e.data.detail || {}`
 * Handlers must use `d` (not `d.detail`) or requestId is lost and downloads time out.
 */

/** Unwrap the page postMessage once, matching content.js onPageMessage. */
export function unwrapPageExportDetail(data) {
  if (!data || data.__spcsv !== true) return {};
  return data.detail && typeof data.detail === "object" ? data.detail : {};
}

/** True when a download handoff payload has a usable requestId. */
export function hasExportDownloadRequestId(detail) {
  return !!(detail && detail.requestId);
}
