import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  VIEW_SCOPE_DEFAULT,
  VIEW_SCOPE_RECURSIVE,
  VIEW_SCOPE_RECURSIVE_ALL,
  VIEW_SCOPE_FILES_ONLY,
  normalizeViewScopeValue,
  buildViewSaveBody
} from "../lib/viewManagerSave.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const viewsSource = readFileSync(join(root, "views.js"), "utf8");
const viewsHtml = readFileSync(join(root, "views.html"), "utf8");

describe("normalizeViewScopeValue", () => {
  it("keeps SharePoint Default (0) instead of coercing to RecursiveAll", () => {
    assert.equal(normalizeViewScopeValue(0), VIEW_SCOPE_DEFAULT);
    assert.equal(normalizeViewScopeValue("0"), VIEW_SCOPE_DEFAULT);
    assert.equal(normalizeViewScopeValue("Default"), VIEW_SCOPE_DEFAULT);
  });

  it("preserves Recursive, RecursiveAll, and FilesOnly", () => {
    assert.equal(normalizeViewScopeValue(1), VIEW_SCOPE_RECURSIVE);
    assert.equal(normalizeViewScopeValue("1"), VIEW_SCOPE_RECURSIVE);
    assert.equal(normalizeViewScopeValue("Recursive"), VIEW_SCOPE_RECURSIVE);
    assert.equal(normalizeViewScopeValue(2), VIEW_SCOPE_RECURSIVE_ALL);
    assert.equal(normalizeViewScopeValue("recursiveall"), VIEW_SCOPE_RECURSIVE_ALL);
    assert.equal(normalizeViewScopeValue(3), VIEW_SCOPE_FILES_ONLY);
    assert.equal(normalizeViewScopeValue("FilesOnly"), VIEW_SCOPE_FILES_ONLY);
  });

  it("falls back to Default for unknown values, not RecursiveAll", () => {
    assert.equal(normalizeViewScopeValue(null), VIEW_SCOPE_DEFAULT);
    assert.equal(normalizeViewScopeValue(""), VIEW_SCOPE_DEFAULT);
    assert.equal(normalizeViewScopeValue(99), VIEW_SCOPE_DEFAULT);
    assert.equal(normalizeViewScopeValue("nope"), VIEW_SCOPE_DEFAULT);
  });
});

describe("buildViewSaveBody", () => {
  const base = {
    title: "All Documents",
    rowLimit: 30,
    scope: 0,
    viewQuery: "<OrderBy><FieldRef Name=\"FileLeafRef\" Ascending=\"True\"/></OrderBy>"
  };

  it("does not send PersonalView when updating an existing view", () => {
    const body = buildViewSaveBody({ ...base, personalView: false, isCreate: false });
    assert.equal(body.Scope, 0);
    assert.equal("PersonalView" in body, false);
  });

  it("sends PersonalView only when creating a view", () => {
    const personal = buildViewSaveBody({ ...base, personalView: true, isCreate: true });
    assert.equal(personal.PersonalView, true);
    const shared = buildViewSaveBody({ ...base, personalView: false, isCreate: true });
    assert.equal(shared.PersonalView, false);
  });

  it("does not treat Scope 0 as falsy RecursiveAll on save", () => {
    const body = buildViewSaveBody({ ...base, scope: "0", isCreate: false });
    assert.equal(body.Scope, 0);
    assert.notEqual(body.Scope, 2);
  });
});

describe("View Manager source sync", () => {
  it("views.js uses normalizeViewScopeValue and omits PersonalView on existing-view PATCH", () => {
    assert.match(viewsSource, /function normalizeViewScopeValue\s*\(/);
    assert.match(viewsSource, /function buildViewSaveBody\s*\(/);
    assert.match(viewsSource, /scopeEl\.value = String\(normalizeViewScopeValue\(sc\)\)/);
    assert.match(viewsSource, /isCreate:\s*!viewId/);
    assert.match(viewsSource, /if \(opts\.isCreate\) body\.PersonalView = !!opts\.personalView;/);
    assert.doesNotMatch(
      viewsSource,
      /parseInt\(document\.getElementById\("viewScope"\)\.value,\s*10\)\s*\|\|\s*2/
    );
    assert.doesNotMatch(viewsSource, /else scopeEl\.value = "2"/);
  });

  it("views.html exposes Default and FilesOnly scope options", () => {
    assert.match(viewsHtml, /<option value="0" selected>/);
    assert.match(viewsHtml, /<option value="3">/);
    assert.match(viewsHtml, /<option value="1">/);
    assert.match(viewsHtml, /<option value="2">/);
  });
});
