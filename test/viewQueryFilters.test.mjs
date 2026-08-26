import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { parseViewQueryParts } from "../lib/viewsDataCore.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const helperSource = readFileSync(join(root, "viewQueryFilters.js"), "utf8");
const viewsSource = readFileSync(join(root, "views.js"), "utf8");
const viewsHtml = readFileSync(join(root, "views.html"), "utf8");
const getViewsDataSource = readFileSync(join(root, "getViewsData.js"), "utf8");

const module = { exports: {} };
Function("module", "window", helperSource)(module, {});
const api = module.exports;

const CHECKED_OUT = [
  "<Where>",
  "<And>",
  "<Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq>",
  "<IsNull><FieldRef Name=\"CheckoutUser\" /></IsNull>",
  "</And>",
  "</Where>",
].join("");

const TODAY_OFFSET = [
  "<Where>",
  "<Geq>",
  "<FieldRef Name=\"DueDate\" />",
  "<Value Type=\"DateTime\"><Today OffsetDays=\"-7\" /></Value>",
  "</Geq>",
  "</Where>",
].join("");

const ME_LOOKUP = [
  "<Where>",
  "<Eq>",
  "<FieldRef Name=\"Author\" LookupId=\"TRUE\" />",
  "<Value Type=\"Integer\"><UserID /></Value>",
  "</Eq>",
  "</Where>",
].join("");

const INCLUDE_TIME = [
  "<Where>",
  "<Geq>",
  "<FieldRef Name=\"Modified\" />",
  "<Value Type=\"DateTime\" IncludeTimeValue=\"TRUE\">2024-01-01T00:00:00Z</Value>",
  "</Geq>",
  "</Where>",
].join("");

describe("IsNull / IsNotNull round-trip (View Manager save)", () => {
  it("loads a SharePoint 'is empty' filter instead of returning null", () => {
    const rows = api.parseViewQueryToFilterRows(
      "<Where><IsNull><FieldRef Name=\"Status\" /></IsNull></Where>"
    );
    assert.deepEqual(rows, [{ field: "Status", op: "IsNull", value: "" }]);
  });

  it("loads IsNull when FieldRef is not self-closing", () => {
    const rows = api.parseViewQueryToFilterRows(
      "<Where><IsNull><FieldRef Name=\"Status\"></FieldRef></IsNull></Where>"
    );
    assert.deepEqual(rows, [{ field: "Status", op: "IsNull", value: "" }]);
  });

  it("loads IsNotNull", () => {
    const rows = api.parseViewQueryToFilterRows(
      "<Where><IsNotNull><FieldRef Name=\"AssignedTo\" /></IsNotNull></Where>"
    );
    assert.deepEqual(rows, [{ field: "AssignedTo", op: "IsNotNull", value: "" }]);
  });

  it("keeps IsNull when mixed with Eq (checked-out files view)", () => {
    const rows = api.parseViewQueryToFilterRows(CHECKED_OUT);
    assert.equal(rows.length, 2);
    assert.equal(rows[0].field, "FSObjType");
    assert.equal(rows[0].op, "Eq");
    assert.equal(rows[0].value, "0");
    assert.equal(rows[1].field, "CheckoutUser");
    assert.equal(rows[1].op, "IsNull");
    assert.equal(rows[1].join, "And");
  });

  it("save emits IsNull even when the editor value is blank", () => {
    const xml = api.buildFilterWhereXml([{ field: "Status", op: "IsNull", value: "" }]);
    assert.match(xml, /<Where><IsNull><FieldRef Name="Status"\/><\/IsNull><\/Where>/);
    assert.doesNotMatch(xml, /<Value/);
  });

  it("save keeps mixed Eq + IsNull instead of dropping Where", () => {
    const loaded = api.parseViewQueryToFilterRows(CHECKED_OUT);
    const xml = api.buildFilterWhereXml(loaded, { getValueType: () => "Integer" });
    assert.match(xml, /<IsNull><FieldRef Name="CheckoutUser"\/><\/IsNull>/);
    assert.match(xml, /<Eq><FieldRef Name="FSObjType"\/><Value Type="Integer">0<\/Value><\/Eq>/);
  });

  it("still skips blank Eq rows so incomplete editor rows are not written", () => {
    const xml = api.buildFilterWhereXml([{ field: "Title", op: "Eq", value: "   " }]);
    assert.equal(xml, "");
  });
});

describe("nested CAML tokens and extra Value attributes", () => {
  it("loads Today OffsetDays as [Today]-N and writes CAML back", () => {
    const rows = api.parseViewQueryToFilterRows(TODAY_OFFSET);
    assert.equal(rows[0].field, "DueDate");
    assert.equal(rows[0].op, "Geq");
    assert.equal(rows[0].value, "[Today]-7");
    const xml = api.buildFilterWhereXml(rows);
    assert.match(xml, /<Today OffsetDays="-7" \/>/);
    assert.doesNotMatch(xml, /\[Today\]/);
  });

  it("loads UserID as [Me] and writes LookupId + UserID", () => {
    const rows = api.parseViewQueryToFilterRows(ME_LOOKUP);
    assert.equal(rows[0].value, "[Me]");
    const xml = api.buildFilterWhereXml(rows);
    assert.match(xml, /LookupId="TRUE"/);
    assert.match(xml, /<UserID \/>/);
  });

  it("does not drop a date filter that has IncludeTimeValue", () => {
    const rows = api.parseViewQueryToFilterRows(INCLUDE_TIME);
    assert.equal(rows[0].field, "Modified");
    assert.equal(rows[0].value, "2024-01-01T00:00:00Z");
    const xml = api.buildFilterWhereXml(rows, { getValueType: () => "DateTime" });
    assert.match(xml, /2024-01-01T00:00:00Z/);
  });
});

describe("parseViewQueryParts fallback (getViewsData / viewsDataCore)", () => {
  it("extracts IsNull so a failed tree parse still has the empty-column filter", () => {
    const parsed = parseViewQueryParts("<Where><IsNull><FieldRef Name=\"Status\" /></IsNull></Where>");
    assert.equal(parsed.filters.length, 1);
    assert.equal(parsed.filters[0].op, "IsNull");
    assert.equal(parsed.filters[0].field, "Status");
    assert.equal(parsed.filters[0].value, "");
  });

  it("extracts IncludeTimeValue date values", () => {
    const parsed = parseViewQueryParts(INCLUDE_TIME);
    assert.equal(parsed.filters[0].field, "Modified");
    assert.equal(parsed.filters[0].value, "2024-01-01T00:00:00Z");
  });

  it("maps Today OffsetDays in the fallback decoder", () => {
    const parsed = parseViewQueryParts(TODAY_OFFSET);
    assert.equal(parsed.filters[0].value, "[Today]-7");
  });
});

describe("source sync", () => {
  it("views.html loads viewQueryFilters.js before views.js", () => {
    assert.match(viewsHtml, /<script src="viewQueryFilters\.js"><\/script>/);
    const helperAt = viewsHtml.indexOf("viewQueryFilters.js");
    const viewsAt = viewsHtml.indexOf('src="views.js"');
    assert.ok(helperAt >= 0 && viewsAt > helperAt);
  });

  it("views.js parses and saves through SPViewQueryFilters (not the old value-only skip)", () => {
    assert.match(viewsSource, /viewQueryFiltersApi\(\)/);
    assert.match(viewsSource, /parseViewQueryToFilterRows/);
    assert.match(viewsSource, /buildFilterWhereXml\(filters,/);
    assert.match(viewsSource, /isUnaryFilterOp/);
    assert.doesNotMatch(
      viewsSource,
      /if \(!\(f\.value \|\| ""\)\.trim\(\)\) continue;/
    );
    assert.doesNotMatch(
      viewsSource,
      /leafTags = \["Eq", "Neq", "Gt", "Geq", "Lt", "Leq", "Contains", "BeginsWith"\]/
    );
  });

  it("getViewsData.js fallback extracts IsNull/IsNotNull", () => {
    assert.match(getViewsDataSource, /IsNull\|IsNotNull/);
    assert.match(getViewsDataSource, /camlValueInnerToEditorFallback/);
  });

  it("helper FILTER_OPS includes is empty / is not empty", () => {
    const labels = api.FILTER_OPS.map((o) => o.value);
    assert.ok(labels.includes("IsNull"));
    assert.ok(labels.includes("IsNotNull"));
  });
});
