import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  parseViewQueryParts,
  parseCamlFieldRefs,
  editorQueryStateFromViewDetails,
  mergeStubFieldsForView,
} from "../lib/viewsDataCore.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const viewsSource = readFileSync(join(root, "views.js"), "utf8");
const getViewsDataSource = readFileSync(join(root, "getViewsData.js"), "utf8");

const MULTI_SORT_GROUPED = [
  "<Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where>",
  "<OrderBy>",
  "<FieldRef Name=\"Status\" Ascending=\"TRUE\" />",
  "<FieldRef Name=\"Modified\" Ascending=\"FALSE\" />",
  "</OrderBy>",
  "<GroupBy Collapse=\"TRUE\" GroupLimit=\"100\">",
  "<FieldRef Name=\"Dept\" Ascending=\"TRUE\" />",
  "<FieldRef Name=\"Category\" Ascending=\"FALSE\" />",
  "</GroupBy>",
].join("");

describe("parseViewQueryParts OrderBy", () => {
  it("keeps every OrderBy FieldRef, not only the first", () => {
    const parsed = parseViewQueryParts(MULTI_SORT_GROUPED);
    assert.equal(parsed.orderBy.field, "Status");
    assert.equal(parsed.orderBy.ascending, true);
    assert.deepEqual(parsed.orderByLevels, [
      { field: "Status", ascending: true },
      { field: "Modified", ascending: false },
    ]);
  });

  it("treats OrderBy FieldRef without Ascending as ascending (All Documents)", () => {
    const parsed = parseViewQueryParts('<OrderBy><FieldRef Name="FileLeafRef" /></OrderBy>');
    assert.deepEqual(parsed.orderByLevels, [{ field: "FileLeafRef", ascending: true }]);
    assert.equal(parsed.orderBy.field, "FileLeafRef");
  });

  it("reads Ascending before Name", () => {
    const parsed = parseViewQueryParts(
      '<OrderBy><FieldRef Ascending="FALSE" Name="Modified" /></OrderBy>'
    );
    assert.deepEqual(parsed.orderByLevels, [{ field: "Modified", ascending: false }]);
  });
});

describe("parseViewQueryParts GroupBy", () => {
  it("loads GroupBy even when Collapse and GroupLimit attributes are present", () => {
    const parsed = parseViewQueryParts(MULTI_SORT_GROUPED);
    assert.equal(parsed.groupBy, "Dept");
    assert.equal(parsed.groupExpand, false);
    assert.equal(parsed.groupLimit, 100);
    assert.deepEqual(parsed.groupByLevels, [
      { field: "Dept", ascending: true },
      { field: "Category", ascending: false },
    ]);
  });

  it("treats Collapse=FALSE as expand-by-default", () => {
    const parsed = parseViewQueryParts(
      '<GroupBy Collapse="FALSE"><FieldRef Name="Status" /></GroupBy>'
    );
    assert.equal(parsed.groupBy, "Status");
    assert.equal(parsed.groupExpand, true);
  });

  it("still parses a GroupBy tag with no attributes", () => {
    const parsed = parseViewQueryParts("<GroupBy><FieldRef Name=\"Owner\" /></GroupBy>");
    assert.equal(parsed.groupBy, "Owner");
    assert.equal(parsed.groupExpand, false);
  });
});

describe("editorQueryStateFromViewDetails", () => {
  it("loads secondary sorts and collapsed two-level grouping for save", () => {
    const parsed = parseViewQueryParts(MULTI_SORT_GROUPED);
    const state = editorQueryStateFromViewDetails(parsed);
    assert.equal(state.sortLevels.length, 2);
    assert.equal(state.sortLevels[1].field, "Modified");
    assert.equal(state.sortLevels[1].ascending, false);
    assert.equal(state.groupByColumn, "Dept");
    assert.equal(state.groupExpand, false);
    assert.equal(state.groupLimit, 100);
    assert.equal(state.groupByExtraLevels.length, 1);
    assert.equal(state.groupByExtraLevels[0].field, "Category");
  });

  it("does not drop FileLeafRef sort when Ascending is omitted", () => {
    const parsed = parseViewQueryParts('<OrderBy><FieldRef Name="FileLeafRef" /></OrderBy>');
    const state = editorQueryStateFromViewDetails(parsed);
    assert.deepEqual(state.sortLevels, [{ field: "FileLeafRef", ascending: true }]);
  });
});

describe("mergeStubFieldsForView", () => {
  it("stubs every sort and group field, not only the first OrderBy", () => {
    const parsed = parseViewQueryParts(MULTI_SORT_GROUPED);
    const fields = [];
    mergeStubFieldsForView(
      fields,
      ["LinkFilename"],
      parsed.filters,
      parsed.orderBy,
      parsed.groupBy,
      parsed.orderByLevels,
      parsed.groupByLevels
    );
    const names = fields.map((f) => f.internalName).sort();
    assert.deepEqual(names, ["Category", "Dept", "FSObjType", "LinkFilename", "Modified", "Status"]);
  });
});

describe("parseCamlFieldRefs", () => {
  it("accepts self-closing and attribute-reordered FieldRefs", () => {
    assert.deepEqual(parseCamlFieldRefs('<FieldRef Name="A"/><FieldRef Ascending="False" Name="B" />'), [
      { field: "A", ascending: true },
      { field: "B", ascending: false },
    ]);
  });
});

describe("ViewQuery parse source sync", () => {
  it("getViewsData.js parses OrderBy/GroupBy blocks instead of first-FieldRef-only regexes", () => {
    assert.match(getViewsDataSource, /function parseCamlFieldRefs\s*\(/);
    assert.match(getViewsDataSource, /<OrderBy\\b\[\^>\]\*>\(\[\\s\\S\]\*\?\)<\\\/OrderBy>/);
    assert.match(getViewsDataSource, /<GroupBy\\b\(\[\^>\]\*\)>\(\[\\s\\S\]\*\?\)<\\\/GroupBy>/);
    assert.match(getViewsDataSource, /orderByLevels/);
    assert.match(getViewsDataSource, /groupByLevels/);
    assert.match(getViewsDataSource, /groupExpand: details\.groupExpand/);
    assert.doesNotMatch(
      getViewsDataSource,
      /<OrderBy>\\s\*<FieldRef\\s\+Name="\(\[\^"\]\+\)"\\s\+Ascending=/
    );
    assert.doesNotMatch(getViewsDataSource, /<GroupBy>\\s\*<FieldRef\\s\+Name=/);
  });

  it("views.js loads all sort levels and GroupBy collapse from viewDetails", () => {
    assert.match(viewsSource, /function editorQueryStateFromViewDetails\s*\(/);
    assert.match(viewsSource, /applyEditorQueryState\(editorQueryStateFromViewDetails\(viewDetails\)\)/);
    assert.match(viewsSource, /groupByExtraLevels/);
    assert.match(viewsSource, /GroupLimit=/);
    assert.doesNotMatch(
      viewsSource,
      /sortLevels = viewDetails\.orderBy \? \[\{ field: viewDetails\.orderBy\.field/
    );
    assert.doesNotMatch(viewsSource, /groupByColumn = viewDetails\.groupBy \|\| "";\s*groupExpand = true;/);
  });
});
