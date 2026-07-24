/**
 * Sharing-link attribution helpers for Permissions Matrix export.
 * Keep in sync with permissionsMatrixExport.js (injected classic page script).
 */

/** Request body for GetSharingInformation. Never ask for inherited parent links. */
export function buildSharingApiRequestBody() {
  return { request: { populateInheritedLinks: false } };
}

/**
 * True when SharePoint marks the link as inherited from an ancestor.
 * Direct links on the requested item should return false/undefined.
 */
export function isInheritedSharingLink(link) {
  if (!link || typeof link !== "object") return false;
  const details = link.linkDetails || link.LinkDetails || {};
  const candidates = [
    link.isInherited,
    link.IsInherited,
    details.isInherited,
    details.IsInherited,
    link.inheritedFrom,
    link.InheritedFrom,
    details.inheritedFrom,
    details.InheritedFrom
  ];
  for (let i = 0; i < candidates.length; i++) {
    const v = candidates[i];
    if (v === true) return true;
    if (typeof v === "string" && v.trim()) return true;
    if (v && typeof v === "object") return true;
  }
  return false;
}

/**
 * Decide whether to record a sharing link against the current queue entry.
 * Rejects blank URLs, already-seen URLs, and inherited/ancestor links so
 * remediation targets the owning item path (not a uniquely permissioned child).
 */
export function shouldAttributeSharingLink(link, linkUrl, seenUrlSet) {
  const url = String(linkUrl || "").trim();
  if (!url) return false;
  if (isInheritedSharingLink(link)) return false;
  const key = url.toLowerCase();
  if (seenUrlSet && seenUrlSet[key]) return false;
  return true;
}
