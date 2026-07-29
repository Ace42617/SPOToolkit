/**
 * Integrity helpers for REST per-item version history fetches during export.
 * Keep behavioral decisions in sync with exportCSV.js variants
 * (injected classic page scripts — cannot import from lib/).
 */

/**
 * fetchAllVersionsForItem returns null on HTTP / payload failure.
 * Callers must not treat null like an empty version list and fall through
 * to current-item-only while still reporting success "with versions".
 * @param {unknown} versions
 * @returns {boolean}
 */
export function isRestVersionsFetchFailure(versions) {
  return versions == null;
}

/**
 * User-facing message when a required versions fetch fails.
 * Callers must not download a file after this failure.
 * @param {string|number} itemId
 * @returns {string}
 */
export function formatRestVersionsFetchFailureMessage(itemId) {
  const id = itemId != null && String(itemId).trim() !== "" ? String(itemId) : "?";
  return (
    "Version history export failed while fetching versions for item " +
    id +
    ". No file downloaded. Resolve access or throttling, then retry."
  );
}
