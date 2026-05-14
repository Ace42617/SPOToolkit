import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { normalizeSharePointPreviewUrl } from "../lib/viewFormatterSecurity.mjs";

describe("normalizeSharePointPreviewUrl", () => {
  it("allows normal SharePoint HTTPS URLs", () => {
    assert.equal(
      normalizeSharePointPreviewUrl("https://contoso.sharepoint.com/sites/hr/Lists/Tasks/AllItems.aspx"),
      "https://contoso.sharepoint.com/sites/hr/Lists/Tasks/AllItems.aspx"
    );
  });

  it("rejects userinfo host confusion", () => {
    assert.equal(normalizeSharePointPreviewUrl("https://contoso.sharepoint.com@evil.example/"), "");
  });

  it("rejects non-SharePoint and non-HTTPS URLs", () => {
    assert.equal(normalizeSharePointPreviewUrl("https://contoso.sharepoint.com.evil.example/"), "");
    assert.equal(normalizeSharePointPreviewUrl("http://contoso.sharepoint.com/sites/hr"), "");
  });

  it("rejects tenant admin pages", () => {
    assert.equal(normalizeSharePointPreviewUrl("https://contoso-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx"), "");
  });
});

describe("View Formatter iframe rules", () => {
  it("does not ship global static DNR header-stripping rules", async () => {
    const manifest = JSON.parse(await readFile(new URL("../manifest.json", import.meta.url), "utf8"));
    assert.equal(manifest.declarative_net_request, undefined);
    assert.ok(manifest.permissions.includes("declarativeNetRequest"));
  });
});
