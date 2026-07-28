/**
 * Integrity helpers for temporary RPC view field setup during CSV/XLSX export.
 * Keep behavioral decisions in sync with exportCSV.js /
 * SP-Developer-Toolkit-Experimental/exportCSV.js (injected classic page
 * scripts — cannot import from lib/).
 */

/**
 * Decide whether a setViewFieldsBatch run succeeded.
 * Any failed addViewField must fail closed: the previous half-failure
 * tolerance silently omitted up to half of the selected columns while the
 * export still reported success.
 *
 * @param {number} failed
 * @param {number} total
 * @returns {{ ok: true } | { ok: false, error: string }}
 */
export function evaluateViewFieldsBatchResult(failed, total) {
  const f = Number(failed);
  const t = Number(total);
  const failedCount = Number.isFinite(f) && f > 0 ? Math.floor(f) : 0;
  const totalCount = Number.isFinite(t) && t > 0 ? Math.floor(t) : 0;

  if (totalCount <= 0) {
    return {
      ok: false,
      error:
        "View field setup failed: no columns to attach. No file downloaded. Try refreshing the page and run the export again."
    };
  }
  if (failedCount > 0) {
    return {
      ok: false,
      error:
        "View field setup failed for " +
        failedCount +
        " of " +
        totalCount +
        " columns (403 or permission denied). No file downloaded. Try refreshing the page and run the export again."
    };
  }
  return { ok: true };
}
