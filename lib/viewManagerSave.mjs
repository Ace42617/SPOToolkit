/**
 * View Manager save-roundtrip helpers for SP.View Scope and PersonalView.
 * views.js is a classic page script and cannot import from lib/ — keep the
 * inline copies there in sync via source-sync asserts in the companion test.
 *
 * SPViewScope:
 * 0 Default      — files and folders of the current folder
 * 1 Recursive    — all files of all folders (no folder items)
 * 2 RecursiveAll — all files and all subfolders of all folders
 * 3 FilesOnly    — files of the current folder only
 */

export const VIEW_SCOPE_DEFAULT = 0;
export const VIEW_SCOPE_RECURSIVE = 1;
export const VIEW_SCOPE_RECURSIVE_ALL = 2;
export const VIEW_SCOPE_FILES_ONLY = 3;

/**
 * Map a REST Scope value (number, numeric string, or enum name) to SPViewScope.
 * Unknown values fall back to Default (0), never RecursiveAll (2).
 * @param {unknown} raw
 * @returns {0|1|2|3}
 */
export function normalizeViewScopeValue(raw) {
  if (raw == null || raw === "") return VIEW_SCOPE_DEFAULT;
  if (typeof raw === "string") {
    const t = raw.trim().toLowerCase();
    if (t === "default") return VIEW_SCOPE_DEFAULT;
    if (t === "recursive") return VIEW_SCOPE_RECURSIVE;
    if (t === "recursiveall") return VIEW_SCOPE_RECURSIVE_ALL;
    if (t === "filesonly") return VIEW_SCOPE_FILES_ONLY;
  }
  const n = typeof raw === "number" ? raw : parseInt(String(raw), 10);
  if (
    n === VIEW_SCOPE_DEFAULT ||
    n === VIEW_SCOPE_RECURSIVE ||
    n === VIEW_SCOPE_RECURSIVE_ALL ||
    n === VIEW_SCOPE_FILES_ONLY
  ) {
    return n;
  }
  return VIEW_SCOPE_DEFAULT;
}

/**
 * Body for create (POST) vs update (PATCH) of a view.
 * PersonalView is only settable at create time; including it on PATCH can
 * convert a personal view to public (or fail the save) because the editor
 * checkbox is hidden for existing views and stays unchecked.
 * @param {{ title: string, rowLimit: number, scope: unknown, viewQuery: string, personalView?: boolean, isCreate: boolean }} opts
 */
export function buildViewSaveBody(opts) {
  const body = {
    Title: opts.title,
    RowLimit: opts.rowLimit,
    Scope: normalizeViewScopeValue(opts.scope),
    ViewQuery: opts.viewQuery
  };
  if (opts.isCreate) body.PersonalView = !!opts.personalView;
  return body;
}
