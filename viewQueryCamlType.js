/**
 * CAML <Value Type="..."> resolution for View Manager filter round-trip.
 * Loaded by views.html before views.js; also runnable under Node for tests.
 *
 * SharePoint stores comparison types on the Value element (ContentTypeId, Integer,
 * Number, User, Lookup, …). Inferring Type from REST TypeAsString (or a Text stub
 * for hidden fields) silently changes Eq/Geq semantics on Save.
 */
(function (root) {
  "use strict";

  function sanitizeCamlValueType(raw) {
    const s = String(raw == null ? "" : raw).trim();
    if (/^[A-Za-z][A-Za-z0-9]{0,63}$/.test(s)) return s;
    return "";
  }

  /**
   * Infer CAML Value Type from internal name + REST TypeAsString when the live
   * ViewQuery Type was not preserved (new editor rows).
   */
  function camlValueTypeFromField(internalName, typeAsString) {
    const name = String(internalName == null ? "" : internalName);
    const t = String(typeAsString == null ? "" : typeAsString);
    if (name === "ContentTypeId" || /^ContentTypeId$/i.test(t)) return "ContentTypeId";
    if (name === "FSObjType" || name === "EventType") return "Integer";
    if (/Integer|Counter|Boolean|YesNo/i.test(t)) return "Integer";
    if (/Number|Currency|Decimal/i.test(t)) return "Number";
    if (/DateTime|Date/i.test(t)) return "DateTime";
    return "Text";
  }

  /**
   * Prefer the Type parsed from live ViewQuery; otherwise infer from the field.
   */
  function resolveFilterCamlValueType(storedType, internalName, typeAsString) {
    const stored = sanitizeCamlValueType(storedType);
    if (stored) return stored;
    return camlValueTypeFromField(internalName, typeAsString);
  }

  function buildFilterValueTypeAttr(f, typeAsString, isBooleanField) {
    if (isBooleanField) return "Integer";
    return resolveFilterCamlValueType(f && f.valueType, f && f.field, typeAsString);
  }

  var api = {
    sanitizeCamlValueType: sanitizeCamlValueType,
    camlValueTypeFromField: camlValueTypeFromField,
    resolveFilterCamlValueType: resolveFilterCamlValueType,
    buildFilterValueTypeAttr: buildFilterValueTypeAttr
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  if (typeof window !== "undefined") {
    window.SPViewQueryCamlType = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : this);
