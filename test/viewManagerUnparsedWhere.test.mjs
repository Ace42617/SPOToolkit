import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { describe, it } from "node:test";
import { fileURLToPath } from "node:url";
import {
  extractViewQueryWhereXml,
  resolveSaveWhereXml,
  shouldPreserveUnparsedWhere,
  unparsedWhereLoadMessage,
  unparsedWhereSaveBlockMessage,
} from "../lib/viewManagerUnparsedWhere.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const viewsJs = readFileSync(join(root, "views.js"), "utf8");

const IN_WHERE =
  '<Where><In><FieldRef Name="Status"/><Values>' +
  '<Value Type="Text">Active</Value><Value Type="Text">Pending</Value>' +
  "</Values></In></Where>";

const MEMBERSHIP_WHERE =
  '<Where><Or><Eq><FieldRef Name="AssignedTo"/><Value Type="Integer">12</Value></Eq>' +
  '<Membership Type="CurrentUserGroups"><FieldRef Name="AssignedTo"/></Membership></Or></Where>';

const EQ_WHERE =
  '<Where><Eq><FieldRef Name="FSObjType"/><Value Type="Integer">0</Value></Eq></Where>';

describe("extractViewQueryWhereXml", () => {
  it("extracts a Where block from a full ViewQuery", () => {
    const q = IN_WHERE + '<OrderBy><FieldRef Name="FileLeafRef" Ascending="True"/></OrderBy>';
    assert.equal(extractViewQueryWhereXml(q), IN_WHERE);
  });

  it("returns empty when there is no Where", () => {
    assert.equal(extractViewQueryWhereXml('<OrderBy><FieldRef Name="Title" Ascending="True"/></OrderBy>'), "");
    assert.equal(extractViewQueryWhereXml(""), "");
  });
});

describe("shouldPreserveUnparsedWhere", () => {
  it("preserves In and Membership when the editor has no rows", () => {
    assert.equal(shouldPreserveUnparsedWhere(IN_WHERE, []), true);
    assert.equal(shouldPreserveUnparsedWhere(MEMBERSHIP_WHERE, null), true);
  });

  it("does not preserve when the editor parsed the Where", () => {
    assert.equal(shouldPreserveUnparsedWhere(EQ_WHERE, [{ field: "FSObjType", op: "Eq", value: "0" }]), false);
  });

  it("does not preserve when there is no Where", () => {
    assert.equal(shouldPreserveUnparsedWhere("<OrderBy/>", []), false);
  });
});

describe("resolveSaveWhereXml", () => {
  it("writes the original In filter when the editor is empty", () => {
    const r = resolveSaveWhereXml({
      editorFilters: [],
      preservedWhereXml: IN_WHERE,
      builtWhereXml: "",
    });
    assert.equal(r.ok, true);
    assert.equal(r.whereXml, IN_WHERE);
  });

  it("refuses save if the user added editor filters on top of an unparsed Where", () => {
    const r = resolveSaveWhereXml({
      editorFilters: [{ field: "Title", op: "Contains", value: "x" }],
      preservedWhereXml: IN_WHERE,
      builtWhereXml: "<Where><Contains>...</Contains></Where>",
    });
    assert.equal(r.ok, false);
    assert.match(r.error, /cannot be edited here/i);
  });

  it("uses rebuilt Where for ordinary parsed filters", () => {
    const built = EQ_WHERE;
    const r = resolveSaveWhereXml({
      editorFilters: [{ field: "FSObjType", op: "Eq", value: "0" }],
      preservedWhereXml: "",
      builtWhereXml: built,
    });
    assert.equal(r.ok, true);
    assert.equal(r.whereXml, built);
  });
});

describe("messages", () => {
  it("warns that Save keeps the existing filter", () => {
    assert.match(unparsedWhereLoadMessage(), /keep the existing filter/i);
  });

  it("blocks replacing an unparsed filter with editor rows", () => {
    assert.match(unparsedWhereSaveBlockMessage(), /cannot be edited here/i);
  });
});

describe("views.js source sync", () => {
  it("tracks preserved Where and does not use partial fallback filters when Where is unparsed", () => {
    assert.match(viewsJs, /preservedWhereXml/);
    assert.match(viewsJs, /extractViewQueryWhereXml/);
    assert.match(viewsJs, /shouldPreserveUnparsedWhere/);
    assert.match(viewsJs, /This view has a filter View Manager cannot edit/);
  });

  it("keeps unparsed Where on save and refuses editor rows that would replace it", () => {
    const saveIdx = viewsJs.indexOf("async function saveView()");
    const patchIdx = viewsJs.indexOf("ViewQuery: viewQuery", saveIdx);
    const preserveGuard = viewsJs.indexOf("preservedWhereXml && filters.length", saveIdx);
    const buildIdx = viewsJs.indexOf("function buildViewQuery()");
    const preserveInBuild = viewsJs.indexOf("preservedWhereXml && filters.length === 0", buildIdx);
    assert.ok(saveIdx >= 0 && patchIdx > saveIdx, "saveView PATCHes ViewQuery");
    assert.ok(preserveGuard > saveIdx && preserveGuard < patchIdx, "unparsed-Where guard runs before ViewQuery PATCH");
    assert.ok(preserveInBuild > buildIdx && preserveInBuild < saveIdx, "buildViewQuery emits preserved Where");
  });
});
