import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { describe, it } from "node:test";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

function read(rel) {
  return readFileSync(join(root, rel), "utf8");
}

const helperSrc = read("viewQueryCamlType.js");
const helperModule = { exports: {} };
Function("module", "window", helperSrc)(helperModule, {});
const camlType = helperModule.exports;

const viewsSrc = read("views.js");
const htmlSrc = read("views.html");
const getViewsDataSrc = read("getViewsData.js");
const coreSrc = read("lib/viewsDataCore.mjs");

const LEAF_RE =
  /^<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)\s*>\s*<FieldRef\s+Name="([^"]+)"\s*\/>\s*<Value\s+Type="([^"]*)">([^<]*)<\/Value>\s*<\/\1\s*>$/i;

function parseLeaf(xml) {
  const norm = xml.replace(/\s+/g, " ").trim();
  const m = LEAF_RE.exec(norm);
  if (!m) return null;
  const v = (m[4] || "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">");
  const valueType = (m[3] || "").trim();
  const leaf = { field: m[2], op: m[1], value: v };
  if (valueType) leaf.valueType = valueType;
  return leaf;
}

function buildLeaf(f, typeAsString) {
  const type = camlType.buildFilterValueTypeAttr(f, typeAsString, false);
  const inner = String(f.value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return "<" + f.op + "><FieldRef Name=\"" + f.field + "\"/><Value Type=\"" + type + "\">" + inner + "</Value></" + f.op + ">";
}

function roundTrip(xml, typeAsString) {
  const leaf = parseLeaf(xml);
  assert.ok(leaf, "leaf should parse: " + xml);
  return buildLeaf(leaf, typeAsString);
}

describe("View Manager CAML Value Type round-trip", () => {
  it("views.js leaf regex still matches the helper parse regex", () => {
    assert.ok(
      viewsSrc.includes(
        'const re = /^<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)\\s*>\\s*<FieldRef\\s+Name="([^"]+)"\\s*\\/>\\s*<Value\\s+Type="([^"]*)">([^<]*)<\\/Value>\\s*<\\/\\1\\s*>$/i;'
      )
    );
  });

  it("parse keeps Value Type and save uses it (source-sync)", () => {
    assert.match(viewsSrc, /valueType = \(m\[3\] \|\| ""\)\.trim\(\)/);
    assert.match(viewsSrc, /if \(valueType\) leaf\.valueType = valueType/);
    assert.match(viewsSrc, /if \(node\.valueType\) row\.valueType = node\.valueType/);
    assert.match(viewsSrc, /if \(row\.valueType && !isBooleanFieldName\(row\.field\)\) o\.valueType = row\.valueType/);
    assert.match(viewsSrc, /if \(f\.valueType && !isBooleanFieldName\(f\.field\)\) o\.valueType = f\.valueType/);
    assert.match(viewsSrc, /delete filters\[i\]\.valueType/);
    assert.match(viewsSrc, /resolveFilterValueType\(f\)/);
    assert.match(viewsSrc, /SPViewQueryCamlType/);
    assert.match(htmlSrc, /viewQueryCamlType\.js/);
  });

  it("getViewsData fallback still captures valueType", () => {
    assert.match(getViewsDataSrc, /valueType: m\[3\] \|\| "Text"/);
    assert.match(coreSrc, /valueType: m\[3\] \|\| "Text"/);
  });

  it("Eq ContentTypeId keeps Type=ContentTypeId when field is a Text stub", () => {
    const input =
      '<Eq><FieldRef Name="ContentTypeId" /><Value Type="ContentTypeId">0x010100</Value></Eq>';
    const out = roundTrip(input, "Text");
    assert.match(out, /<Value Type="ContentTypeId">0x010100<\/Value>/);
    assert.doesNotMatch(out, /Type="Text"/);
  });

  it("BeginsWith ContentTypeId keeps Type=ContentTypeId", () => {
    const input =
      '<BeginsWith><FieldRef Name="ContentTypeId"/><Value Type="ContentTypeId">0x0101</Value></BeginsWith>';
    const out = roundTrip(input, "Text");
    assert.match(out, /<Value Type="ContentTypeId">0x0101<\/Value>/);
  });

  it("FSObjType Integer 0 is not rewritten as Text", () => {
    const input = '<Eq><FieldRef Name="FSObjType" /><Value Type="Integer">0</Value></Eq>';
    const out = roundTrip(input, "Text");
    assert.match(out, /<Value Type="Integer">0<\/Value>/);
  });

  it("infers Integer for FSObjType even without stored Type", () => {
    assert.equal(camlType.camlValueTypeFromField("FSObjType", "Text"), "Integer");
    assert.equal(camlType.camlValueTypeFromField("FSObjType", "Lookup"), "Integer");
    assert.equal(camlType.camlValueTypeFromField("EventType", ""), "Integer");
    assert.equal(camlType.camlValueTypeFromField("ContentTypeId", "Text"), "ContentTypeId");
  });

  it("calculated Number Geq keeps Type=Number despite TypeAsString Calculated", () => {
    const input = '<Geq><FieldRef Name="Score" /><Value Type="Number">10</Value></Geq>';
    const out = roundTrip(input, "Calculated");
    assert.match(out, /<Value Type="Number">10<\/Value>/);
  });

  it("calculated DateTime Geq keeps Type=DateTime", () => {
    const input =
      '<Geq><FieldRef Name="DueCalc" /><Value Type="DateTime">2024-01-01T00:00:00Z</Value></Geq>';
    const out = roundTrip(input, "Calculated");
    assert.match(out, /<Value Type="DateTime">2024-01-01T00:00:00Z<\/Value>/);
  });

  it("visible DateTime Geq still emits DateTime", () => {
    const input =
      '<Geq><FieldRef Name="DueDate" /><Value Type="DateTime">2024-01-01T00:00:00Z</Value></Geq>';
    const out = roundTrip(input, "DateTime");
    assert.match(out, /<Value Type="DateTime">2024-01-01T00:00:00Z<\/Value>/);
  });

  it("User Type=User display-name filter is not rewritten as Text", () => {
    const input = '<Eq><FieldRef Name="Author" /><Value Type="User">John Smith</Value></Eq>';
    const out = roundTrip(input, "User");
    assert.match(out, /<Value Type="User">John Smith<\/Value>/);
  });

  it("FileDirRef Type=Lookup is preserved", () => {
    const input =
      '<Eq><FieldRef Name="FileDirRef" /><Value Type="Lookup">/sites/a/Shared Documents/folder</Value></Eq>';
    const out = roundTrip(input, "Text");
    assert.match(out, /<Value Type="Lookup">\/sites\/a\/Shared Documents\/folder<\/Value>/);
  });

  it("ordinary Text/Number fields still infer from TypeAsString when Type was not stored", () => {
    assert.equal(camlType.resolveFilterCamlValueType("", "Title", "Text"), "Text");
    assert.equal(camlType.resolveFilterCamlValueType("", "Qty", "Number"), "Number");
    assert.equal(camlType.resolveFilterCamlValueType("", "Amount", "Currency"), "Number");
    assert.equal(camlType.resolveFilterCamlValueType("", "DueDate", "DateTime"), "DateTime");
  });

  it("stored Type wins over a conflicting TypeAsString", () => {
    assert.equal(
      camlType.resolveFilterCamlValueType("ContentTypeId", "ContentTypeId", "Text"),
      "ContentTypeId"
    );
    assert.equal(camlType.resolveFilterCamlValueType("Number", "Score", "Calculated"), "Number");
  });

  it("rejects unsafe stored Type tokens", () => {
    assert.equal(camlType.sanitizeCamlValueType('Text"/><foo'), "");
    assert.equal(camlType.sanitizeCamlValueType("Text Type"), "");
    assert.equal(
      camlType.resolveFilterCamlValueType('Text"/><foo', "Title", "Text"),
      "Text"
    );
  });

  it("boolean Value Type stays Integer", () => {
    const f = { field: "YesNo", op: "Eq", value: "1", valueType: "Boolean" };
    assert.equal(camlType.buildFilterValueTypeAttr(f, "Boolean", true), "Integer");
  });

  it("helper is the UMD loaded by views.html", () => {
    assert.match(helperSrc, /window\.SPViewQueryCamlType = api/);
    assert.match(helperSrc, /camlValueTypeFromField/);
    assert.match(helperSrc, /resolveFilterCamlValueType/);
  });
});
