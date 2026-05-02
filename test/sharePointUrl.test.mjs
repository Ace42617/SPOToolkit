import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { isSharePointOnlineUrl } from "../lib/sharePointUrl.mjs";

describe("isSharePointOnlineUrl", () => {
  it("accepts HTTPS SharePoint Online hosts", () => {
    assert.equal(isSharePointOnlineUrl("https://contoso.sharepoint.com/sites/hr"), true);
    assert.equal(isSharePointOnlineUrl("https://contoso-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx"), true);
  });

  it("rejects lookalike hosts that only contain sharepoint.com", () => {
    assert.equal(isSharePointOnlineUrl("https://contoso.sharepoint.com.evil.example/sites/hr"), false);
    assert.equal(isSharePointOnlineUrl("https://evil-sharepoint.com/sites/hr"), false);
  });

  it("requires HTTPS and can exclude tenant admin hosts", () => {
    assert.equal(isSharePointOnlineUrl("http://contoso.sharepoint.com/sites/hr"), false);
    assert.equal(
      isSharePointOnlineUrl("https://contoso-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx", {
        allowTenantAdmin: false,
      }),
      false
    );
  });
});
