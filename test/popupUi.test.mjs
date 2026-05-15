/**
 * Tests for lib/popupUi.mjs (popup helpers).
 */
import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  groupUniversalSearchItems,
  listColumnSettingsUrl,
  searchSchemaListMetaFromResponse,
  searchSchemaColumnMatchesFilter,
} from "../lib/popupUi.mjs";

describe("listColumnSettingsUrl", () => {
  it("builds FldEdit URL with encoded List and Field", () => {
    const u = listColumnSettingsUrl("https://contoso.sharepoint.com/sites/a", "abc-def", "MyField");
    assert.ok(u.startsWith("https://contoso.sharepoint.com/sites/a/_layouts/15/FldEdit.aspx?"));
    assert.ok(u.includes("List=" + encodeURIComponent("{abc-def}")));
    assert.ok(u.includes("Field=" + encodeURIComponent("MyField")));
  });
  it("strips trailing slash from site", () => {
    const u = listColumnSettingsUrl("https://x/sites/w/", "g", "Title");
    assert.ok(u.startsWith("https://x/sites/w/_layouts"));
  });
  it("normalizes listId braces", () => {
    const u = listColumnSettingsUrl("https://x", "{guid}", "T");
    assert.ok(u.includes(encodeURIComponent("{guid}")));
  });
  it("returns empty when missing parts", () => {
    assert.equal(listColumnSettingsUrl("", "g", "T"), "");
    assert.equal(listColumnSettingsUrl("https://x", "", "T"), "");
    assert.equal(listColumnSettingsUrl("https://x", "g", ""), "");
  });
});

describe("searchSchemaListMetaFromResponse", () => {
  it("normalizes siteUrl and listId", () => {
    assert.deepEqual(searchSchemaListMetaFromResponse({ siteUrl: "https://a/", listId: "{x}" }), {
      siteUrl: "https://a",
      listId: "x",
    });
  });
  it("handles empty", () => {
    assert.deepEqual(searchSchemaListMetaFromResponse(null), { siteUrl: "", listId: "" });
  });
});

describe("searchSchemaColumnMatchesFilter", () => {
  const col = { internalName: "Foo", title: "Bar", type: "Text", group: "G", crawledProperty: "ows_Foo" };
  it("matches empty query", () => {
    assert.equal(searchSchemaColumnMatchesFilter(col, ""), true);
  });
  it("matches internal name", () => {
    assert.equal(searchSchemaColumnMatchesFilter(col, "foo"), true);
  });
  it("matches crawled property", () => {
    assert.equal(searchSchemaColumnMatchesFilter(col, "ows_"), true);
  });
  it("no match", () => {
    assert.equal(searchSchemaColumnMatchesFilter(col, "zzz"), false);
  });
});

describe("groupUniversalSearchItems", () => {
  it("returns rendered items in the same grouped order as the UI", () => {
    const items = [
      { group: "Tabs", title: "Overview" },
      { group: "Links", title: "List settings" },
      { group: "Tabs", title: "Reports" },
      { group: "Tools", title: "View Manager" },
      { group: "Links", title: "Site contents" },
    ];

    const result = groupUniversalSearchItems(items);

    assert.deepEqual(result.groups.map((g) => [g.group, g.items.map((it) => it.title)]), [
      ["Tabs", ["Overview", "Reports"]],
      ["Links", ["List settings", "Site contents"]],
      ["Tools", ["View Manager"]],
    ]);
    assert.deepEqual(result.renderedItems.map((it) => it.title), [
      "Overview",
      "Reports",
      "List settings",
      "Site contents",
      "View Manager",
    ]);
  });

  it("omits per-group overflow from the keyboard selection model", () => {
    const items = [
      { group: "Columns", title: "Column 1" },
      { group: "Columns", title: "Column 2" },
      { group: "Columns", title: "Column 3" },
      { group: "Tabs", title: "Settings" },
    ];

    const result = groupUniversalSearchItems(items, 2);

    assert.deepEqual(result.groups.map((g) => [g.group, g.items.map((it) => it.title)]), [
      ["Columns", ["Column 1", "Column 2"]],
      ["Tabs", ["Settings"]],
    ]);
    assert.deepEqual(result.renderedItems.map((it) => it.title), ["Column 1", "Column 2", "Settings"]);
  });
});
