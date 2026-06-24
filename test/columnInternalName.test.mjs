import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  encodeInternalNameFromTitle,
  resolveUniqueInternalName,
  predictColumnInternalName,
  crawledForInternal,
} from "../lib/columnInternalName.mjs";

describe("encodeInternalNameFromTitle", () => {
  it("keeps simple names", () => {
    assert.equal(encodeInternalNameFromTitle("Status"), "Status");
  });
  it("encodes spaces", () => {
    assert.equal(encodeInternalNameFromTitle("Hello World"), "Hello_x0020_World");
  });
  it("encodes punctuation", () => {
    assert.equal(encodeInternalNameFromTitle("A:B"), "A_x003a_B");
  });
  it("trims display name", () => {
    assert.equal(encodeInternalNameFromTitle("  Title  "), "Title");
  });
});

describe("resolveUniqueInternalName", () => {
  it("returns base when unused", () => {
    assert.equal(resolveUniqueInternalName("Hello World", ["Other"]), "Hello_x0020_World");
  });
  it("appends numeric suffix on conflict", () => {
    assert.equal(resolveUniqueInternalName("Hello World", ["Hello_x0020_World"]), "Hello_x0020_World0");
  });
});

describe("predictColumnInternalName", () => {
  const siteFields = [{ title: "Owner", internalName: "Owner" }];
  const listFields = [{ title: "Owner", internalName: "Owner" }];

  it("reuses site column when adding from site", () => {
    const p = predictColumnInternalName({
      title: "Owner",
      placement: "listFromSite",
      siteFields,
      listFields: [],
    });
    assert.equal(p.internalName, "Owner");
    assert.equal(p.reusedSiteColumn, true);
    assert.equal(p.alreadyOnList, false);
  });

  it("flags site column already on list", () => {
    const p = predictColumnInternalName({
      title: "Owner",
      placement: "listFromSite",
      siteFields,
      listFields,
    });
    assert.equal(p.alreadyOnList, true);
  });

  it("warns when list-only duplicates site title", () => {
    const p = predictColumnInternalName({
      title: "Owner",
      placement: "listNew",
      siteFields,
      listFields: [],
    });
    assert.ok(p.notes.some((n) => /site column/i.test(n)));
  });
});

describe("crawledForInternal", () => {
  it("prefixes ows_", () => {
    assert.equal(crawledForInternal("Title"), "ows_Title");
  });
  it("rewrites lookup colon encoding", () => {
    assert.equal(crawledForInternal("Advocate_x003a__x0020_Email"), "ows_Advocate:_x0020_Email");
  });
});
