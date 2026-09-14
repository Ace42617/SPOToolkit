/**
 * Decode SharePoint delimited field values for CSV/XLSX export.
 *
 * OWSSVR / classic encodings:
 * - Lookup / Person: `id;#value` or `id;#value;#id;#value`
 * - MultiChoice: `;#value;#value;#` (values only, leading/trailing `;#`)
 *
 * REST nometadata MultiChoice arrives as a string array. Join those with ", "
 * — never `;#` — so lookup-pair decoding cannot drop every other choice.
 */

export function cleanSharePointDelimitedValue(val) {
  if (val == null || val === undefined) return "";
  const s = String(val);
  if (s.indexOf(";#") < 0) return s;
  const parts = s.split(";#");
  const looksLikeLookupPairs = parts.length >= 2 && parts.length % 2 === 0;
  if (looksLikeLookupPairs) {
    let allIdsNumeric = true;
    for (let i = 0; i < parts.length; i += 2) {
      if (!/^-?\d+$/.test(String(parts[i] || "").trim())) {
        allIdsNumeric = false;
        break;
      }
    }
    if (allIdsNumeric) {
      const values = [];
      for (let j = 1; j < parts.length; j += 2) {
        if (String(parts[j] || "") !== "") values.push(parts[j]);
      }
      if (values.length) return values.join(", ");
    }
  }
  const choices = [];
  for (let k = 0; k < parts.length; k++) {
    if (String(parts[k] || "") !== "") choices.push(parts[k]);
  }
  return choices.length ? choices.join(", ") : s;
}

export function flattenRestValue(val) {
  if (val == null || val === undefined) return "";
  if (typeof val === "object") {
    if (val.LookupValue != null) return String(val.LookupValue);
    if (Array.isArray(val)) {
      return val
        .map(function (v) {
          return v && v.LookupValue != null ? String(v.LookupValue) : String(v);
        })
        .filter(function (v) {
          return v !== "";
        })
        .join(", ");
    }
    return JSON.stringify(val);
  }
  return String(val);
}
