/**
 * Guards for the OWSSVR export borrowed-view fallback.
 * Keep behavioral decisions in sync with exportCSV.js (injected classic page
 * script — cannot import from lib/).
 */

/**
 * Whether export may mutate ViewFields on the chosen OWSSVR view.
 *
 * Borrowed library views (All Documents / default) are marked read-only and must
 * never go through setViewFieldsBatch → RemoveAllViewFields. Doing so rewrites
 * a live production view while the UI claims the fallback is read-only.
 *
 * @param {{ ownView?: boolean, readOnlyView?: boolean } | null | undefined} viewInfo
 * @returns {boolean}
 */
export function shouldConfigureOwssvrExportView(viewInfo) {
  if (!viewInfo || typeof viewInfo !== "object") return false;
  if (viewInfo.readOnlyView === true) return false;
  if (viewInfo.ownView === false) return false;
  return true;
}
