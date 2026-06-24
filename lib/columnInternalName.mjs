/**
 * SharePoint field internal name prediction (XmlConvert.EncodeName-style + uniqueness).
 * Used by the column creator to preview the static name before creation.
 */

const MAX_INTERNAL_NAME_LENGTH = 32;

function isXmlNameStartChar(code) {
  return (
    (code >= 0x41 && code <= 0x5a) ||
    (code >= 0x61 && code <= 0x7a) ||
    code === 0x5f
  );
}

function isXmlNameChar(code) {
  return (
    isXmlNameStartChar(code) ||
    (code >= 0x30 && code <= 0x39) ||
    code === 0x2d ||
    code === 0x2e ||
    code === 0x3a
  );
}

function encodeCodePoint(code) {
  return "_x" + code.toString(16).padStart(4, "0") + "_";
}

/**
 * Encode a display title the way SharePoint derives a field static name.
 * @param {string} title
 * @returns {string}
 */
export function encodeInternalNameFromTitle(title) {
  const raw = String(title || "").trim();
  if (!raw) return "";

  let out = "";
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    const code = raw.charCodeAt(i);
    const atStart = out.length === 0;
    const allowed = atStart ? isXmlNameStartChar(code) : isXmlNameChar(code);
    out += allowed ? ch : encodeCodePoint(code);
  }

  if (!out) out = encodeCodePoint(raw.charCodeAt(0));
  if (out.length > MAX_INTERNAL_NAME_LENGTH) out = out.slice(0, MAX_INTERNAL_NAME_LENGTH);
  return out;
}

function normalizeKey(name) {
  return String(name || "").trim().toLowerCase();
}

function fieldIndex(fields) {
  const byInternal = new Map();
  const byTitle = new Map();
  for (const field of fields || []) {
    const internal = String(field.internalName || field.InternalName || "").trim();
    const title = String(field.title || field.Title || "").trim();
    if (internal) byInternal.set(normalizeKey(internal), field);
    if (title) byTitle.set(normalizeKey(title), field);
  }
  return { byInternal, byTitle };
}

function existingInternalNames(fields) {
  const names = new Set();
  for (const field of fields || []) {
    const internal = String(field.internalName || field.InternalName || "").trim();
    if (internal) names.add(internal);
  }
  return names;
}

function fitWithSuffix(base, suffix) {
  const maxBase = Math.max(1, MAX_INTERNAL_NAME_LENGTH - suffix.length);
  return base.slice(0, maxBase) + suffix;
}

/**
 * Pick a unique internal name against an existing set (SharePoint appends 0, 1, 2…).
 * @param {string} base
 * @param {Set<string>|string[]} taken
 * @returns {string}
 */
export function resolveUniqueInternalName(base, taken) {
  const encoded = encodeInternalNameFromTitle(base);
  if (!encoded) return "";
  const used = new Set();
  for (const name of taken || []) {
    const key = String(name || "").trim();
    if (key) used.add(normalizeKey(key));
  }
  if (!used.has(normalizeKey(encoded))) return encoded;

  for (let i = 0; i < 1000; i++) {
    const suffix = String(i);
    const candidate = fitWithSuffix(encoded, suffix);
    if (!used.has(normalizeKey(candidate))) return candidate;
  }
  return fitWithSuffix(encoded, "0");
}

/**
 * @typedef {'site'|'listNew'|'listFromSite'} ColumnPlacement
 */

/**
 * Predict the internal name for a column about to be created.
 * @param {object} input
 * @param {string} input.title Display name
 * @param {ColumnPlacement} input.placement
 * @param {Array<{internalName?:string,title?:string,InternalName?:string,Title?:string}>} [input.siteFields]
 * @param {Array<{internalName?:string,title?:string,InternalName?:string,Title?:string}>} [input.listFields]
 * @returns {{
 *   internalName: string,
 *   placement: ColumnPlacement,
 *   reusedSiteColumn: boolean,
 *   alreadyOnList: boolean,
 *   conflictResolved: boolean,
 *   notes: string[],
 *   crawledProperty: string,
 * }}
 */
export function predictColumnInternalName(input) {
  const title = String(input?.title || "").trim();
  const placement = input?.placement === "site" || input?.placement === "listFromSite" ? input.placement : "listNew";
  const siteFields = Array.isArray(input?.siteFields) ? input.siteFields : [];
  const listFields = Array.isArray(input?.listFields) ? input.listFields : [];
  const siteIndex = fieldIndex(siteFields);
  const listIndex = fieldIndex(listFields);
  const listNames = existingInternalNames(listFields);
  const siteNames = existingInternalNames(siteFields);
  const notes = [];

  if (!title) {
    return {
      internalName: "",
      placement,
      reusedSiteColumn: false,
      alreadyOnList: false,
      conflictResolved: false,
      notes: ["Enter a display name to preview the internal name."],
      crawledProperty: "",
    };
  }

  const siteMatch = siteIndex.byTitle.get(normalizeKey(title));
  const listMatch = listIndex.byTitle.get(normalizeKey(title));

  if (placement === "listFromSite") {
    if (!siteMatch) {
      const encoded = resolveUniqueInternalName(title, listNames);
      notes.push("No site column with this display name yet; a new site column would be created first.");
      return {
        internalName: encoded,
        placement,
        reusedSiteColumn: false,
        alreadyOnList: listIndex.byInternal.has(normalizeKey(encoded)),
        conflictResolved: encoded !== encodeInternalNameFromTitle(title),
        notes,
        crawledProperty: crawledForInternal(encoded),
      };
    }
    const internal = String(siteMatch.internalName || siteMatch.InternalName || "").trim();
    const onList = listIndex.byInternal.has(normalizeKey(internal));
    if (onList) notes.push("This site column is already on the current list.");
    else notes.push("Reuses the existing site column static name when added to the list.");
    return {
      internalName: internal,
      placement,
      reusedSiteColumn: true,
      alreadyOnList: onList,
      conflictResolved: false,
      notes,
      crawledProperty: crawledForInternal(internal),
    };
  }

  if (placement === "site") {
    if (siteMatch) {
      const internal = String(siteMatch.internalName || siteMatch.InternalName || "").trim();
      notes.push("A site column with this display name already exists.");
      return {
        internalName: internal,
        placement,
        reusedSiteColumn: true,
        alreadyOnList: listIndex.byInternal.has(normalizeKey(internal)),
        conflictResolved: false,
        notes,
        crawledProperty: crawledForInternal(internal),
      };
    }
    const encoded = resolveUniqueInternalName(title, siteNames);
    if (encoded !== encodeInternalNameFromTitle(title)) notes.push("Adjusted to avoid a duplicate static name on the site.");
    return {
      internalName: encoded,
      placement,
      reusedSiteColumn: false,
      alreadyOnList: listIndex.byInternal.has(normalizeKey(encoded)),
      conflictResolved: encoded !== encodeInternalNameFromTitle(title),
      notes,
      crawledProperty: crawledForInternal(encoded),
    };
  }

  // listNew — list-only column
  if (siteMatch) {
    const siteInternal = String(siteMatch.internalName || siteMatch.InternalName || "").trim();
    if (!listIndex.byInternal.has(normalizeKey(siteInternal))) {
      notes.push("A site column with this display name exists; add it from the site instead to reuse " + siteInternal + ".");
    }
  }
  if (listMatch) {
    const internal = String(listMatch.internalName || listMatch.InternalName || "").trim();
    notes.push("A list column with this display name already exists.");
    return {
      internalName: internal,
      placement,
      reusedSiteColumn: false,
      alreadyOnList: true,
      conflictResolved: false,
      notes,
      crawledProperty: crawledForInternal(internal),
    };
  }

  const encoded = resolveUniqueInternalName(title, listNames);
  if (encoded !== encodeInternalNameFromTitle(title)) notes.push("Adjusted to avoid a duplicate static name on this list.");
  return {
    internalName: encoded,
    placement,
    reusedSiteColumn: false,
    alreadyOnList: false,
    conflictResolved: encoded !== encodeInternalNameFromTitle(title),
    notes,
    crawledProperty: crawledForInternal(encoded),
  };
}

export function crawledForInternal(internalName) {
  const internal = String(internalName || "");
  if (!internal) return "";
  let crawled = "ows_" + internal;
  if (/_x003a__x0020_/i.test(internal)) {
    crawled = "ows_" + internal.replace(/_x003a__x0020_/gi, ":_x0020_");
  }
  return crawled;
}
