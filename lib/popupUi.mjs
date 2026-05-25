/**
 * Pure helpers shared by extension popups (full + Lite). Covered by test/popupUi.test.mjs.
 */

/** Classic list column settings: FldEdit.aspx */
export function listColumnSettingsUrl(siteUrl, listId, internalName) {
  const base = String(siteUrl || "").replace(/\/$/, "");
  const id = String(listId || "").replace(/[{}]/g, "");
  if (!base || !id || !internalName) return "";
  return (
    base +
    "/_layouts/15/FldEdit.aspx?List=" +
    encodeURIComponent("{" + id + "}") +
    "&Field=" +
    encodeURIComponent(internalName)
  );
}

export function searchSchemaListMetaFromResponse(response) {
  return {
    siteUrl: String(response?.siteUrl || "").replace(/\/$/, ""),
    listId: String(response?.listId || "").replace(/[{}]/g, ""),
  };
}

export function searchSchemaColumnMatchesFilter(column, qLower) {
  if (!qLower) return true;
  const parts = [column.internalName, column.title, column.type, column.group, column.crawledProperty];
  return parts.some((p) => p && String(p).toLowerCase().includes(qLower));
}

export function displayedUniversalSearchItems(items, perGroupLimit = 18) {
  const limit = Math.max(0, Number(perGroupLimit) || 0);
  const groups = new Map();
  for (const item of Array.isArray(items) ? items : []) {
    const group = item?.group || "";
    if (!groups.has(group)) groups.set(group, []);
    groups.get(group).push(item);
  }

  const displayed = [];
  groups.forEach((groupItems) => {
    displayed.push(...groupItems.slice(0, limit));
  });
  return displayed;
}

/**
 * @param {HTMLElement} contentRoot — #quicklinksContent
 * @param {string} queryRaw — filter text (case-insensitive substring)
 */
export function applyQuicklinksFilter(contentRoot, queryRaw) {
  const q = String(queryRaw || "").toLowerCase().trim();
  for (const section of contentRoot.querySelectorAll(".ql-section")) {
    if (!q) {
      section.style.display = "";
      for (const li of section.querySelectorAll("li")) li.style.display = "";
      continue;
    }
    const labelText = (section.querySelector(".ql-section-label")?.textContent || "").toLowerCase();
    const sectionHit = labelText.includes(q);
    let any = false;
    for (const li of section.querySelectorAll("li")) {
      const a = li.querySelector("a");
      const text = (a?.textContent || "").toLowerCase();
      const href = (a?.getAttribute("href") || "").toLowerCase();
      const hit = sectionHit || text.includes(q) || href.includes(q);
      li.style.display = hit ? "" : "none";
      if (hit) any = true;
    }
    section.style.display = any ? "" : "none";
  }
}
