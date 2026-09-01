import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, it } from "node:test";
import {
  buildLookupAwareLeafXml,
  filterShouldWriteLookupId,
  parseCamlFieldRefAttrs,
  parseLookupAwareLeafXml,
  scanLookupAwareFilterLeaves,
} from "../lib/viewQueryLookupId.mjs";
import { parseViewQueryParts } from "../lib/viewsDataCore.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

function read(rel) {
  return readFileSync(join(root, rel), "utf8");
}

const LOOKUP_ID_LEAF =
  '<Eq><FieldRef Name="AssignedTo" LookupId="TRUE" /><Value Type="Integer">23</Value></Eq>';
const LOOKUP_ID_LEAF_ATTR_ORDER =
  '<Eq><FieldRef LookupId="TRUE" Name="Department" /><Value Type="Integer">5</Value></Eq>';
const STATUS_LEAF = '<Eq><FieldRef Name="Status" /><Value Type="Text">Active</Value></Eq>';
const MIXED_WHERE =
  "<Where><And>" + STATUS_LEAF + LOOKUP_ID_LEAF + "</And></Where>";

const OLD_LEAF_RE =
  /^<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)\s*>\s*<FieldRef\s+Name="([^"]+)"\s*\/>\s*<Value\s+Type="([^"]*)">([^<]*)<\/Value>\s*<\/\1\s*>$/i;
const OLD_SCAN_RE =
  /<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)>\s*<FieldRef\s+Name="([^"]+)"\s*\/>\s*<Value\s+Type="([^"]*)">([^<]*)<\/Value>/gi;

describe("LookupId person/lookup CAML", () => {
  it("does not match the old Name-only FieldRef leaf regex", () => {
    const norm = LOOKUP_ID_LEAF.replace(/\s+/g, " ").trim();
    assert.equal(OLD_LEAF_RE.exec(norm), null);
  });

  it("old fallback scan keeps Status and drops AssignedTo LookupId", () => {
    const found = [];
    let m;
    OLD_SCAN_RE.lastIndex = 0;
    while ((m = OLD_SCAN_RE.exec(MIXED_WHERE)) !== null) {
      found.push(m[2]);
    }
    assert.deepEqual(found, ["Status"]);
  });

  it("parses LookupId=TRUE on Assigned To (modern person filter)", () => {
    const leaf = parseLookupAwareLeafXml(LOOKUP_ID_LEAF);
    assert.ok(leaf);
    assert.equal(leaf.field, "AssignedTo");
    assert.equal(leaf.op, "Eq");
    assert.equal(leaf.value, "23");
    assert.equal(leaf.lookupId, true);
    assert.equal(leaf.valueType, "Integer");
  });

  it("parses LookupId before Name", () => {
    const leaf = parseLookupAwareLeafXml(LOOKUP_ID_LEAF_ATTR_ORDER);
    assert.ok(leaf);
    assert.equal(leaf.field, "Department");
    assert.equal(leaf.lookupId, true);
    assert.equal(leaf.value, "5");
  });

  it("still parses Name-only FieldRef filters", () => {
    const leaf = parseLookupAwareLeafXml(STATUS_LEAF);
    assert.ok(leaf);
    assert.equal(leaf.field, "Status");
    assert.equal(leaf.lookupId, undefined);
    assert.equal(leaf.value, "Active");
  });

  it("scan keeps both Status and AssignedTo LookupId siblings", () => {
    const rows = scanLookupAwareFilterLeaves(MIXED_WHERE);
    assert.equal(rows.length, 2);
    assert.equal(rows[0].field, "Status");
    assert.equal(rows[0].lookupId, undefined);
    assert.equal(rows[1].field, "AssignedTo");
    assert.equal(rows[1].lookupId, true);
    assert.equal(rows[1].value, "23");
  });

  it("parseViewQueryParts fallback list includes LookupId filters", () => {
    const parsed = parseViewQueryParts(MIXED_WHERE);
    assert.equal(parsed.filters.length, 2);
    assert.equal(parsed.filters[1].field, "AssignedTo");
    assert.equal(parsed.filters[1].lookupId, true);
  });

  it("round-trips LookupId + Integer on save", () => {
    const leaf = parseLookupAwareLeafXml(LOOKUP_ID_LEAF);
    const xml = buildLookupAwareLeafXml(leaf, "Text");
    assert.match(xml, /LookupId="TRUE"/);
    assert.match(xml, /Name="AssignedTo"/);
    assert.match(xml, /<Value Type="Integer">23<\/Value>/);
    const again = parseLookupAwareLeafXml(xml);
    assert.equal(again.lookupId, true);
    assert.equal(again.value, "23");
  });

  it("writes LookupId for a numeric id on a User column even if lookupId was not loaded", () => {
    assert.equal(
      filterShouldWriteLookupId({ field: "AssignedTo", op: "Eq", value: "23" }, "User"),
      true
    );
    const xml = buildLookupAwareLeafXml({ field: "AssignedTo", op: "Eq", value: "23" }, "User");
    assert.match(xml, /LookupId="TRUE"/);
  });

  it("does not write LookupId for a choice/text value", () => {
    assert.equal(
      filterShouldWriteLookupId({ field: "Status", op: "Eq", value: "Active" }, "Choice"),
      false
    );
    const xml = buildLookupAwareLeafXml(
      { field: "Status", op: "Eq", value: "Active", valueType: "Text" },
      "Choice"
    );
    assert.doesNotMatch(xml, /LookupId/);
  });

  it("parseCamlFieldRefAttrs reads Name and LookupId independently", () => {
    assert.deepEqual(parseCamlFieldRefAttrs(' Name="AssignedTo" LookupId="TRUE" '), {
      name: "AssignedTo",
      lookupId: true,
    });
    assert.deepEqual(parseCamlFieldRefAttrs(' Name="Status" '), {
      name: "Status",
      lookupId: false,
    });
  });
});

describe("source-sync: View Manager load/save must keep LookupId", () => {
  it("views.js leaf parser accepts extra FieldRef attributes and writes LookupId", () => {
    const src = read("views.js");
    const leaf = src.match(/function parseLeafCondString\([\s\S]*?\n  \}/);
    assert.ok(leaf, "parseLeafCondString not found");
    assert.match(leaf[0], /FieldRef\\b/);
    assert.match(leaf[0], /parseCamlFieldRefAttrs/);
    assert.match(leaf[0], /lookupId/);

    const build = src.match(/function buildOneFilterCondXml\([\s\S]*?\n  \}/);
    assert.ok(build, "buildOneFilterCondXml not found");
    assert.match(build[0], /filterShouldWriteLookupId/);
    assert.match(build[0], /LookupId=\\"TRUE\\"/);

    assert.match(src, /if \(row\.lookupId\) o\.lookupId = true;/);
    assert.match(src, /if \(f\.lookupId\) o\.lookupId = true;/);
  });

  it("getViewsData.js fallback scan keeps LookupId FieldRefs", () => {
    const src = read("getViewsData.js");
    const fn = src.match(/function parseViewQueryParts\([\s\S]*?\n  \}/);
    assert.ok(fn, "parseViewQueryParts not found");
    assert.match(fn[0], /FieldRef\\b/);
    assert.match(fn[0], /lookupId = true/);
    assert.doesNotMatch(
      fn[0],
      /FieldRef\\s\+Name="\(\[\^"\]\+\)"\\s\*\\\/>/
    );
  });
});
