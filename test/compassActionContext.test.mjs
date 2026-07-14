import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const source = await readFile(new URL("../compassActionContext.js", import.meta.url), "utf8");

function loadResolver() {
  const window = {};
  Function("window", source)(window);
  return window.SPOT_resolveCompassActionContext;
}

describe("resolveCompassActionContext", () => {
  it("uses and normalizes the context returned at action time", async () => {
    const resolveContext = loadResolver();
    const calls = [];
    const result = await resolveContext(
      async (message) => {
        calls.push(message);
        return {
          ok: true,
          webAbsoluteUrl: "https://contoso.sharepoint.com/sites/current/",
          pageListId: "{CURRENT-LIST}",
        };
      },
      "getPageContext",
      true
    );

    assert.deepEqual(calls, [{ action: "getPageContext" }]);
    assert.equal(result.siteUrl, "https://contoso.sharepoint.com/sites/current");
    assert.equal(result.listId, "CURRENT-LIST");
  });

  it("supports column-creator context responses", async () => {
    const resolveContext = loadResolver();
    const result = await resolveContext(
      async () => ({
        ok: true,
        siteUrl: "https://contoso.sharepoint.com/sites/current",
        listId: "{LIST-B}",
        siteFields: [{ internalName: "Owner" }],
      }),
      "getColumnCreatorContext",
      true
    );

    assert.equal(result.listId, "LIST-B");
    assert.equal(result.siteFields[0].internalName, "Owner");
  });

  it("fails closed when the current page is not a list", async () => {
    const resolveContext = loadResolver();

    await assert.rejects(
      resolveContext(
        async () => ({
          ok: true,
          webAbsoluteUrl: "https://contoso.sharepoint.com/sites/current",
          pageListId: "",
        }),
        "getPageContext",
        true
      ),
      /Open a list or library view/
    );
  });

  it("propagates context lookup failures instead of using a cached target", async () => {
    const resolveContext = loadResolver();

    await assert.rejects(
      resolveContext(async () => ({ ok: false, error: "No context from page" }), "getPageContext", true),
      /No context from page/
    );
  });
});
