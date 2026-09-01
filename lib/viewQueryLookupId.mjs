/**
 * Person / lookup view filters in SharePoint CAML use
 * `<FieldRef Name="AssignedTo" LookupId="TRUE" /><Value Type="Integer">23</Value>`.
 * The View Manager leaf parser used to require FieldRef with only Name, so those
 * conditions never loaded and Save rewrote ViewQuery without them.
 */

const VALUE_OPS = "Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith";

export function decodeCamlText(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

export function encodeCamlText(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function encodeCamlAttr(s) {
  return encodeCamlText(s).replace(/"/g, "&quot;");
}

/** Parse attributes inside a FieldRef tag (without the FieldRef word). */
export function parseCamlFieldRefAttrs(attrStr) {
  const s = String(attrStr || "");
  const nameM = /\bName\s*=\s*"([^"]+)"/i.exec(s);
  if (!nameM) return null;
  const lookupM = /\bLookupId\s*=\s*"(TRUE|True|true|1)"/i.exec(s);
  return { name: nameM[1], lookupId: Boolean(lookupM) };
}

export function fieldTypeUsesLookupId(typeAsString) {
  return /^(User|UserMulti|Lookup|LookupMulti)$/i.test(String(typeAsString || "").trim());
}

export function isNumericLookupFilterValue(value) {
  return /^\d+$/.test(String(value || "").trim());
}

/**
 * Emit LookupId="TRUE" + integer value when the loaded CAML had LookupId,
 * or when the editor value is a numeric id on a Person/Lookup column.
 */
export function filterShouldWriteLookupId(filterRow, typeAsString) {
  if (!filterRow || !isNumericLookupFilterValue(filterRow.value)) return false;
  if (filterRow.lookupId) return true;
  return fieldTypeUsesLookupId(typeAsString);
}

const LEAF_RE = new RegExp(
  "^<(" +
    VALUE_OPS +
    ")\\s*>\\s*<FieldRef\\b([^>]*?)(?:\\s*\\/>|><\\/FieldRef\\s*>)\\s*<Value\\s+Type=\"([^\"]*)\">([^<]*)<\\/Value>\\s*<\\/\\1\\s*>$",
  "i"
);

export function parseLookupAwareLeafXml(xml) {
  const norm = String(xml || "").replace(/\s+/g, " ").trim();
  const m = LEAF_RE.exec(norm);
  if (!m) return null;
  const fr = parseCamlFieldRefAttrs(m[2]);
  if (!fr) return null;
  const row = {
    type: "leaf",
    field: fr.name,
    op: m[1],
    value: decodeCamlText(m[4]),
    valueType: m[3] || "Text",
  };
  if (fr.lookupId) row.lookupId = true;
  return row;
}

/** Global scan used by getViewsData / viewsDataCore fallback filter lists. */
export function scanLookupAwareFilterLeaves(viewQuery) {
  const q = String(viewQuery || "").replace(/\s+/g, " ");
  const re = new RegExp(
    "<(" +
      VALUE_OPS +
      ")\\s*>\\s*<FieldRef\\b([^>]*?)(?:\\s*\\/>|><\\/FieldRef\\s*>)\\s*<Value\\s+Type=\"([^\"]*)\">([^<]*)<\\/Value>",
    "gi"
  );
  const filters = [];
  let m;
  while ((m = re.exec(q)) !== null) {
    const fr = parseCamlFieldRefAttrs(m[2]);
    if (!fr) continue;
    const row = {
      op: m[1],
      field: fr.name,
      valueType: m[3] || "Text",
      value: decodeCamlText(m[4]),
    };
    if (fr.lookupId) row.lookupId = true;
    filters.push(row);
  }
  return filters;
}

export function buildLookupAwareLeafXml(filterRow, typeAsString) {
  if (!filterRow || !filterRow.op || !filterRow.field) return null;
  const op = String(filterRow.op);
  const name = encodeCamlAttr(filterRow.field);
  const inner = encodeCamlText(filterRow.value);
  if (filterShouldWriteLookupId(filterRow, typeAsString)) {
    return (
      "<" +
      op +
      "><FieldRef Name=\"" +
      name +
      "\" LookupId=\"TRUE\"/><Value Type=\"Integer\">" +
      inner +
      "</Value></" +
      op +
      ">"
    );
  }
  const type = String(filterRow.valueType || "Text").replace(/"/g, "") || "Text";
  return (
    "<" + op + "><FieldRef Name=\"" + name + "\"/><Value Type=\"" + type + "\">" + inner + "</Value></" + op + ">"
  );
}
