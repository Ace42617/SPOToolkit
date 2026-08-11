/**
 * Request-correlation helpers for Compass Site Contents web inventory.
 * Injected classic page scripts cannot import from lib/ — keep inline copies in
 * content.js / getWebContents.js in sync via source-sync asserts in the companion test.
 */

export const WEB_CONTENTS_PARAMS_ATTR = "data-sp-site-contents-web-params";
export const WEB_CONTENTS_REQUEST_ATTR = "data-request-id";
export const WEB_CONTENTS_REQUEST_QUERY = "spcsvRequestId";
export const WEB_CONTENTS_RESULT_TYPE = "SPCSVWebContentsResult";

/**
 * Whether a page postMessage belongs to the outstanding Site Contents wait.
 * Untagged or mismatched results must be ignored so overlapping injects cannot
 * cross-resolve (wrong web inventory / permanent Loading).
 * @param {any} data
 * @param {string} requestId
 * @returns {boolean}
 */
export function matchesWebContentsResponse(data, requestId) {
  if (!data || data.__spcsv !== true || data.type !== WEB_CONTENTS_RESULT_TYPE) {
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
export function extractWebContentsRequestIdFromScriptSrc(src, baseHref) {
  try {
    const href = String(src || "");
    if (!href) return "";
    return (
      new URL(href, baseHref || "https://example.invalid/").searchParams.get(
        WEB_CONTENTS_REQUEST_QUERY
      ) || ""
    );
  } catch (_) {
    return "";
  }
}

/**
 * Find the params script node for one Site Contents request.
 * Prefer request-scoped lookup so overlapping injects do not read the last node.
 * @param {{ querySelectorAll: (sel: string) => ArrayLike<{ getAttribute: (name: string) => string|null, textContent?: string|null }> }} documentLike
 * @param {string} requestId
 * @returns {{ getAttribute: (name: string) => string|null, textContent?: string|null }|null}
 */
export function findWebContentsParamsElement(documentLike, requestId) {
  if (!documentLike || typeof documentLike.querySelectorAll !== "function") return null;
  const nodes = documentLike.querySelectorAll(
    'script[' + WEB_CONTENTS_PARAMS_ATTR + '="1"]'
  );
  if (!nodes || !nodes.length) return null;
  if (requestId) {
    for (let i = 0; i < nodes.length; i++) {
      const el = nodes[i];
      if (el && el.getAttribute(WEB_CONTENTS_REQUEST_ATTR) === requestId) return el;
    }
    return null;
  }
  return nodes[nodes.length - 1] || null;
}

/**
 * Whether a Site Contents child-web cache entry should block a new fetch.
 * Successfully loaded inventories block; in-flight loads block; errors must not
 * (empty lists arrays are truthy and previously stranded nodes after failure).
 * @param {{ loading?: boolean, lists?: any[]|null, error?: string }|null|undefined} cache
 * @returns {boolean}
 */
export function shouldSkipSiteContentsWebFetch(cache) {
  if (!cache) return false;
  if (cache.loading) return true;
  if (cache.error) return false;
  return Array.isArray(cache.lists);
}
