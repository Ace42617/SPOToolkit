import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  normalizeTrailingSlash,
  secondStageRecycleBinUrl,
  ADMIN_RECYCLE_BIN_SECOND_STAGE_SUFFIX,
} from "../lib/recycleBinUrls.mjs";

describe("normalizeTrailingSlash", () => {
  it("strips one trailing slash", () => {
    assert.equal(normalizeTrailingSlash("https://a.sharepoint.com/sites/x/"), "https://a.sharepoint.com/sites/x");
  });
  it("empty", () => {
    assert.equal(normalizeTrailingSlash(""), "");
  });
});

describe("secondStageRecycleBinUrl", () => {
  it("uses AdminRecycleBin with view=5#view=13 (OOB second-stage pattern)", () => {
    const u = secondStageRecycleBinUrl("https://contoso.sharepoint.com/sites/hr");
    assert.equal(
      u,
      "https://contoso.sharepoint.com/sites/hr" + ADMIN_RECYCLE_BIN_SECOND_STAGE_SUFFIX
    );
    assert.ok(u.includes("AdminRecycleBin.aspx"));
    assert.ok(u.includes("view=5#view=13"));
    assert.ok(!u.includes("/_layouts/15/RecycleBin.aspx"));
  });
  it("returns empty for missing base", () => {
    assert.equal(secondStageRecycleBinUrl(""), "");
  });
});
