/**
 * OData next-page URL from a SharePoint REST JSON payload.
 *
 * Invoke-PnPSPRestMethod defaults to Accept odata=nometadata. That JSON uses
 * `@odata.nextLink`. PnP's parsed (non -Raw) path only copies JSON property
 * `odata.nextLink` (no @) onto the PSObject, so `$response.'odata.nextLink'`
 * is empty and paging stops after the first $top page.
 */

export const ODATA_NEXT_LINK_NAMES = Object.freeze([
  "@odata.nextLink",
  "odata.nextLink"
]);

/**
 * @param {unknown} payload parsed REST body (nometadata, minimalmetadata, or verbose)
 * @returns {string|null} absolute or relative next-page URL
 */
export function getODataNextLinkFromRestPayload(payload) {
  if (payload == null || typeof payload !== "object") return null;
  const obj = payload;
  for (const name of ODATA_NEXT_LINK_NAMES) {
    const v = obj[name];
    if (v != null && String(v).trim()) return String(v).trim();
  }
  const d = obj.d;
  if (d && typeof d === "object" && d.__next != null && String(d.__next).trim()) {
    return String(d.__next).trim();
  }
  return null;
}
