/**
 * Child-web discovery helpers for site/subsite inventory.
 * Keep in sync with permissionsMatrixExport.js, matrixScanPlan.js, and
 * reportListPlan.js (injected classic page scripts cannot import this module).
 */

/** Page size used for /_api/web/webs discovery. SharePoint may return a nextLink beyond this. */
export const CHILD_WEBS_PAGE_SIZE = 500;

/** Build the first-page child-webs REST URL for a web. */
export function childWebsApiUrl(webUrl) {
  const base = String(webUrl || "").replace(/\/$/, "");
  return (
    base +
    "/_api/web/webs?$select=Title,ServerRelativeUrl,Url&$top=" +
    CHILD_WEBS_PAGE_SIZE
  );
}

/** Extract OData collection items from a nometadata/verbose-ish payload. */
export function extractODataValue(json) {
  if (!json || typeof json !== "object") return [];
  if (Array.isArray(json.value)) return json.value;
  if (Array.isArray(json.results)) return json.results;
  return [];
}

/** Extract the next page URL, if any. */
export function extractODataNextLink(json) {
  if (!json || typeof json !== "object") return null;
  const link = json["@odata.nextLink"] || json["odata.nextLink"] || null;
  return link ? String(link) : null;
}

/**
 * Fetch every direct child web under webUrl, following @odata.nextLink.
 *
 * @param {(url: string, accept?: string) => Promise<object>} fetchJson
 * @param {string} webUrl Absolute web URL (no trailing slash required)
 * @param {string} [accept]
 * @param {(url: string) => string} [normalizeUrl] Optional URL normalizer (e.g. same-origin rewrite)
 * @returns {Promise<object[]>}
 */
export async function fetchAllChildWebs(fetchJson, webUrl, accept, normalizeUrl) {
  const all = [];
  let next = childWebsApiUrl(webUrl);
  while (next) {
    const j = await fetchJson(next, accept);
    const items = extractODataValue(j);
    for (let i = 0; i < items.length; i++) all.push(items[i]);
    const link = extractODataNextLink(j);
    next = link ? (normalizeUrl ? normalizeUrl(link) : link) : null;
  }
  return all;
}
