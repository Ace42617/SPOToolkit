/**
 * Integrity helpers for Permissions Matrix browser export.
 * Keep behavioral decisions in sync with permissionsMatrixExport.js
 * (injected classic page script — cannot import from lib/).
 */

/**
 * Order fetched role definition names: preferred roles first, then the rest.
 * @param {string[]} fetchedNames
 * @param {string[]} preferredRoles
 * @returns {string[]}
 */
export function orderRoleNames(fetchedNames, preferredRoles) {
  const names = Array.isArray(fetchedNames) ? fetchedNames.filter(Boolean) : [];
  const preferred = Array.isArray(preferredRoles) ? preferredRoles : [];
  const ordered = [];
  preferred.forEach((r) => {
    if (names.indexOf(r) >= 0 && ordered.indexOf(r) < 0) ordered.push(r);
  });
  names.forEach((r) => {
    if (ordered.indexOf(r) < 0) ordered.push(r);
  });
  return ordered;
}

/**
 * Resolve role column names from a successful roledefinitions payload.
 * Must fail closed when SharePoint returns no usable names — never substitute
 * a hard-coded preferred list (that silently drops custom role grants).
 * @returns {{ ok: true, roleNames: string[] } | { ok: false, error: string, roleNames: string[] }}
 */
export function resolveRoleNamesFromDefs(defs, preferredRoles) {
  const roleNames = [];
  const list = Array.isArray(defs) ? defs : [];
  for (let i = 0; i < list.length; i++) {
    const name = String((list[i] && (list[i].Name || list[i].name)) || "").trim();
    if (name && roleNames.indexOf(name) < 0) roleNames.push(name);
  }
  if (roleNames.length === 0) {
    return {
      ok: false,
      error: "No role definitions returned for this site. No file downloaded.",
      roleNames: []
    };
  }
  return { ok: true, roleNames: orderRoleNames(roleNames, preferredRoles) };
}

/** Terminal success is only allowed when role names loaded and no list scan errors. */
export function canReportMatrixSuccess({ roleNamesOk, scanErrors }) {
  if (!roleNamesOk) return false;
  if (Array.isArray(scanErrors) && scanErrors.length > 0) return false;
  return true;
}

/**
 * User-facing message when a list/library scan fails mid-export.
 * Callers must not download a workbook after this failure.
 */
export function formatMatrixListScanFailureMessage(listTitle, err) {
  const title = String(listTitle || "list").trim() || "list";
  const detail = err && err.message ? err.message : String(err || "unknown error");
  return (
    'Permissions matrix failed while scanning "' +
    title +
    '": ' +
    detail +
    ". No file downloaded. Resolve access or throttling, then retry."
  );
}
