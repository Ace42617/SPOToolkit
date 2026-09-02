import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { describe, it } from "node:test";
import { fileURLToPath } from "node:url";
import {
  extractViewQueryWhereXml,
  flattenAndRebuildWhereXml,
  groupedWhereLoadMessage,
  groupedWhereSaveBlockMessage,
  resolveGroupedWhereSave,
  shouldPreserveGroupedWhere,
} from "../lib/viewQueryAndOr.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const viewsJs = readFileSync(join(root, "views.js"), "utf8");

function leaf(op, field, value) {
  return "<" + op + "><FieldRef Name=\"" + field + "\"/><Value Type=\"Text\">" + value + "</Value></" + op + ">";
}

const ACTIVE = leaf("Eq", "Status", "Active");
const HIGH = leaf("Eq", "Priority", "High");
const EMEA = leaf("Eq", "Region", "EMEA");
const DRAFT = leaf("Eq", "Status", "Draft");

/** Status=Active OR (Priority=High AND Region=EMEA) */
const GROUPED_OR_AND =
  "<Where><Or>" + ACTIVE + "<And>" + HIGH + EMEA + "</And></Or></Where>";

/** (Status=Active OR Priority=High) AND Region=EMEA — left-associative, representable */
const LEFT_ASSOC_AND_OR =
  "<Where><And><Or>" + ACTIVE + HIGH + "</Or>" + EMEA + "</And></Where>";

/** All-And chain SharePoint often emits as n-ary And */
const ALL_AND =
  "<Where><And>" + ACTIVE + HIGH + EMEA + "</And></Where>";

const ALL_OR =
  "<Where><Or>" + ACTIVE + DRAFT + "</Or></Where>";

const SINGLE_EQ = "<Where>" + ACTIVE + "</Where>";

describe("flattenAndRebuildWhereXml corrupts grouped And/Or", () => {
  it("turns A OR (B AND C) into (A OR B) AND C", () => {
    const rebuilt = flattenAndRebuildWhereXml(GROUPED_OR_AND);
    assert.equal(rebuilt, LEFT_ASSOC_AND_OR);
    assert.notEqual(rebuilt, GROUPED_OR_AND);
  });

  it("round-trips a left-associative mixed chain", () => {
    assert.equal(flattenAndRebuildWhereXml(LEFT_ASSOC_AND_OR), LEFT_ASSOC_AND_OR);
  });

  it("round-trips uniform And / Or / single leaf", () => {
    assert.match(flattenAndRebuildWhereXml(ALL_AND), /<And>/);
    assert.equal(flattenAndRebuildWhereXml(ALL_OR), ALL_OR);
    assert.equal(flattenAndRebuildWhereXml(SINGLE_EQ), SINGLE_EQ);
  });
});

describe("shouldPreserveGroupedWhere", () => {
  it("preserves A OR (B AND C)", () => {
    assert.equal(shouldPreserveGroupedWhere(GROUPED_OR_AND), true);
  });

  it("preserves (A AND B) OR (C AND D)", () => {
    const q =
      "<Where><Or><And>" + ACTIVE + HIGH + "</And><And>" + DRAFT + EMEA + "</And></Or></Where>";
    assert.equal(shouldPreserveGroupedWhere(q), true);
  });

  it("does not preserve left-associative mixed And/Or", () => {
    assert.equal(shouldPreserveGroupedWhere(LEFT_ASSOC_AND_OR), false);
  });

  it("does not preserve uniform And, Or, or a single Eq", () => {
    assert.equal(shouldPreserveGroupedWhere(ALL_AND), false);
    assert.equal(shouldPreserveGroupedWhere(ALL_OR), false);
    assert.equal(shouldPreserveGroupedWhere(SINGLE_EQ), false);
  });

  it("does not preserve when there is no Where", () => {
    assert.equal(shouldPreserveGroupedWhere('<OrderBy><FieldRef Name="Title" Ascending="True"/></OrderBy>'), false);
    assert.equal(shouldPreserveGroupedWhere(""), false);
  });
});

describe("extractViewQueryWhereXml", () => {
  it("keeps the original grouped Where next to OrderBy", () => {
    const q = GROUPED_OR_AND + '<OrderBy><FieldRef Name="FileLeafRef" Ascending="True"/></OrderBy>';
    assert.equal(extractViewQueryWhereXml(q), GROUPED_OR_AND);
  });
});

describe("resolveGroupedWhereSave", () => {
  it("writes the original grouped Where when the editor is empty", () => {
    const r = resolveGroupedWhereSave({
      editorFilters: [],
      preservedWhereXml: GROUPED_OR_AND,
      builtWhereXml: LEFT_ASSOC_AND_OR,
    });
    assert.equal(r.ok, true);
    assert.equal(r.whereXml, GROUPED_OR_AND);
  });

  it("refuses save if the user added editor filters on a grouped Where", () => {
    const r = resolveGroupedWhereSave({
      editorFilters: [{ field: "Title", op: "Contains", value: "x" }],
      preservedWhereXml: GROUPED_OR_AND,
      builtWhereXml: "<Where><Contains>...</Contains></Where>",
    });
    assert.equal(r.ok, false);
    assert.match(r.error, /grouped And\/Or/i);
  });

  it("uses rebuilt Where for ordinary parsed filters", () => {
    const r = resolveGroupedWhereSave({
      editorFilters: [{ field: "Status", op: "Eq", value: "Active" }],
      preservedWhereXml: "",
      builtWhereXml: SINGLE_EQ,
    });
    assert.equal(r.ok, true);
    assert.equal(r.whereXml, SINGLE_EQ);
  });
});

describe("messages", () => {
  it("warns that Save keeps the existing filter", () => {
    assert.match(groupedWhereLoadMessage(), /keep the existing filter/i);
  });

  it("blocks replacing a grouped filter with editor rows", () => {
    assert.match(groupedWhereSaveBlockMessage(), /cannot edit/i);
  });
});

describe("views.js source sync", () => {
  it("tracks preserved grouped Where and refuses a lossy rebuild", () => {
    assert.match(viewsJs, /preservedGroupedWhereXml/);
    assert.match(viewsJs, /shouldPreserveGroupedWhere/);
    assert.match(viewsJs, /grouped And\/Or filters that cannot be edited here/);
  });

  it("keeps grouped Where on save and refuses editor rows that would replace it", () => {
    const saveIdx = viewsJs.indexOf("async function saveView()");
    const patchIdx = viewsJs.indexOf("ViewQuery: viewQuery", saveIdx);
    const preserveGuard = viewsJs.indexOf("preservedGroupedWhereXml && filters.length", saveIdx);
    const buildIdx = viewsJs.indexOf("function buildViewQuery()");
    const preserveInBuild = viewsJs.indexOf("preservedGroupedWhereXml && filters.length === 0", buildIdx);
    assert.ok(saveIdx >= 0 && patchIdx > saveIdx, "saveView PATCHes ViewQuery");
    assert.ok(preserveGuard > saveIdx && preserveGuard < patchIdx, "grouped-Where guard runs before ViewQuery PATCH");
    assert.ok(preserveInBuild > buildIdx && preserveInBuild < saveIdx, "buildViewQuery emits preserved grouped Where");
  });
});
