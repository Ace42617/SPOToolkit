import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import {
  buildSharingApiRequestBody,
  isInheritedSharingLink,
  shouldAttributeSharingLink
} from "../lib/sharingLinkAttribution.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

describe("buildSharingApiRequestBody", () => {
  it("does not request inherited/ancestor sharing links", () => {
    assert.deepEqual(buildSharingApiRequestBody(), {
      request: { populateInheritedLinks: false }
    });
  });
});

describe("isInheritedSharingLink", () => {
  it("detects common inheritance markers", () => {
    assert.equal(isInheritedSharingLink({ isInherited: true }), true);
    assert.equal(isInheritedSharingLink({ LinkDetails: { IsInherited: true } }), true);
    assert.equal(isInheritedSharingLink({ linkDetails: { inheritedFrom: "/sites/a/Shared Documents" } }), true);
    assert.equal(isInheritedSharingLink({ shareUrl: "https://x" }), false);
    assert.equal(isInheritedSharingLink(null), false);
  });
});

describe("shouldAttributeSharingLink", () => {
  it("attributes a direct link once to the current item", () => {
    const seen = {};
    const link = { shareUrl: "https://contoso.sharepoint.com/:f:/s/x" };
    assert.equal(shouldAttributeSharingLink(link, link.shareUrl, seen), true);
    seen[link.shareUrl.toLowerCase()] = true;
    assert.equal(shouldAttributeSharingLink(link, link.shareUrl, seen), false);
  });

  it("never attributes an inherited link to a uniquely permissioned child", () => {
    const seen = {};
    const inherited = {
      isInherited: true,
      linkDetails: { url: "https://contoso.sharepoint.com/:f:/s/parent-link" }
    };
    assert.equal(
      shouldAttributeSharingLink(inherited, inherited.linkDetails.url, seen),
      false
    );
    assert.equal(seen["https://contoso.sharepoint.com/:f:/s/parent-link"], undefined);
  });
});

describe("permissionsMatrixExport.js wiring", () => {
  const source = readFileSync(join(root, "permissionsMatrixExport.js"), "utf8");

  it("requests populateInheritedLinks:false", () => {
    assert.match(
      source,
      /populateInheritedLinks:\s*false/
    );
    assert.doesNotMatch(
      source,
      /populateInheritedLinks:\s*true/
    );
  });

  it("skips inherited links before URL dedupe attribution", () => {
    assert.match(source, /function isInheritedSharingLink\s*\(/);
    assert.match(source, /if\s*\(\s*isInheritedSharingLink\s*\(\s*link\s*\)\s*\)\s*continue/);
  });
});
