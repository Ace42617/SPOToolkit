/**
 * SharePoint recycle bin layout pages:
 * - RecycleBin.aspx — first-stage (end-user) bin for the current web.
 * - AdminRecycleBin.aspx — site collection admin bin; second stage via OOB hash #view=13.
 * Quick links load RecycleBin.aspx first in the same tab, then this URL (sequence runs in background.js).
 */

export function normalizeTrailingSlash(url) {
  if (!url || typeof url !== "string") return "";
  return url.replace(/\/$/, "");
}

/** OOB “second-stage recycle bin” link (after first-stage recycle bin has loaded in-session). */
export const ADMIN_RECYCLE_BIN_SECOND_STAGE_SUFFIX = "/_layouts/15/AdminRecycleBin.aspx?view=5#view=13";

/**
 * @param {string} siteCollectionRootAbsoluteUrl — must be site collection root (e.g. _spPageContextInfo.siteAbsoluteUrl)
 */
export function secondStageRecycleBinUrl(siteCollectionRootAbsoluteUrl) {
  const base = normalizeTrailingSlash(siteCollectionRootAbsoluteUrl);
  if (!base) return "";
  return base + ADMIN_RECYCLE_BIN_SECOND_STAGE_SUFFIX;
}
