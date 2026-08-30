/**
 * View Manager rebuilds ViewQuery Where from editor rows (Eq/Neq/Contains/…).
 * SharePoint also emits <In>, <Membership>, <Includes>, IsNull, [Today], [Me], etc.
 * If those do not parse into editor rows, Save must not write an empty/partial Where.
 */

export function extractViewQueryWhereXml(viewQuery) {
  const q = String(viewQuery || "");
  const m = q.match(/<Where\b[^>]*>[\s\S]*?<\/Where>/i);
  return m ? m[0] : "";
}

/**
 * True when the live view has a <Where> that the editor could not represent.
 * Callers must not fall back to a regex that only picks Eq/Neq leaves out of a
 * larger tree (that silently drops <In> / <Membership> siblings).
 */
export function shouldPreserveUnparsedWhere(viewQuery, parsedFilterRows) {
  const whereXml = extractViewQueryWhereXml(viewQuery);
  if (!whereXml) return false;
  return !(Array.isArray(parsedFilterRows) && parsedFilterRows.length > 0);
}

export function unparsedWhereLoadMessage() {
  return (
    "This view has a filter View Manager cannot edit (for example \"is one of\", " +
    "group membership, or multi-value includes). Save will keep the existing filter as-is."
  );
}

export function unparsedWhereSaveBlockMessage() {
  return (
    "Cannot save: this view's existing filter cannot be edited here. " +
    "Remove the new filter rows to keep the original filter, or change the filter in SharePoint."
  );
}

/**
 * @param {{ editorFilters: Array, preservedWhereXml: string, builtWhereXml: string }} input
 * @returns {{ ok: true, whereXml: string } | { ok: false, error: string }}
 */
export function resolveSaveWhereXml(input) {
  const src = input || {};
  const preserved = String(src.preservedWhereXml || "");
  const editorFilters = Array.isArray(src.editorFilters) ? src.editorFilters : [];
  if (preserved) {
    if (editorFilters.length > 0) {
      return { ok: false, error: unparsedWhereSaveBlockMessage() };
    }
    return { ok: true, whereXml: preserved };
  }
  return { ok: true, whereXml: String(src.builtWhereXml || "") };
}
