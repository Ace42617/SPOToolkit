import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";

const sourceUrl = new URL("../getViewsData.js", import.meta.url);

function waitForPosts(posts, expectedCount) {
  return new Promise((resolve, reject) => {
    let attempts = 0;
    function check() {
      if (posts.length >= expectedCount) {
        resolve();
        return;
      }
      attempts++;
      if (attempts > 20) {
        reject(new Error("Timed out waiting for injected results"));
        return;
      }
      setTimeout(check, 0);
    }
    check();
  });
}

describe("getViewsData injected request correlation", () => {
  it("reads the params script that matches the current script request id", async () => {
    const source = await readFile(sourceUrl, "utf8");
    const posts = [];
    const fetchUrls = [];
    const elements = new Map();
    elements.set("sp-views-params", {
      textContent: JSON.stringify({
        requestId: "stale",
        listId: "wrong-list",
        viewId: "wrong-view",
        webAbsoluteUrl: "https://tenant/sites/wrong",
      }),
    });
    elements.set("sp-views-params-req-a", {
      textContent: JSON.stringify({
        requestId: "req-a",
        listId: "list-a",
        viewId: "view-a",
        webAbsoluteUrl: "https://tenant/sites/site",
      }),
    });
    elements.set("sp-views-params-req-b", {
      textContent: JSON.stringify({
        requestId: "req-b",
        listId: "list-b",
        viewId: "view-b",
        webAbsoluteUrl: "https://tenant/sites/site",
      }),
    });

    const document = {
      currentScript: null,
      getElementById(id) {
        return elements.get(id) || null;
      },
    };
    const window = {
      location: { href: "https://tenant/sites/site/Lists/List/AllItems.aspx" },
      _spPageContextInfo: {
        webAbsoluteUrl: "https://tenant/sites/site",
        listUrl: "/sites/site/Lists/List",
      },
      postMessage(data) {
        posts.push(data);
      },
    };
    const fetch = async (url) => {
      fetchUrls.push(String(url));
      if (String(url).includes("/lists(guid'list-a')?$select=Title")) {
        return { ok: true, json: async () => ({ Title: "List A" }) };
      }
      if (String(url).includes("/lists(guid'list-b')?$select=Title")) {
        return { ok: true, json: async () => ({ Title: "List B" }) };
      }
      if (String(url).includes("/views?$select")) {
        return { ok: true, json: async () => ({ value: [] }) };
      }
      if (String(url).includes("/fields?$select")) {
        return { ok: true, json: async () => ({ value: [{ InternalName: "Title", Title: "Title" }] }) };
      }
      if (String(url).includes("/views(guid'view-a')/ViewFields")) {
        return { ok: true, json: async () => ({ Items: ["Title"] }) };
      }
      if (String(url).includes("/views(guid'view-a')?$select=ViewQuery,Title")) {
        return { ok: true, json: async () => ({ Title: "View A", ViewQuery: "" }) };
      }
      if (String(url).includes("/views(guid'view-b')/ViewFields")) {
        return { ok: true, json: async () => ({ Items: ["Title"] }) };
      }
      if (String(url).includes("/views(guid'view-b')?$select=ViewQuery,Title")) {
        return { ok: true, json: async () => ({ Title: "View B", ViewQuery: "" }) };
      }
      throw new Error("Unexpected fetch URL: " + url);
    };

    const context = vm.createContext({ document, window, fetch, URL });
    document.currentScript = { src: "https://extension.test/getViewsData.js?spcsvRequestId=req-a" };
    vm.runInContext(source, context);
    document.currentScript = { src: "https://extension.test/getViewsData.js?spcsvRequestId=req-b" };
    vm.runInContext(source, context);

    await waitForPosts(posts, 2);
    const byRequest = Object.fromEntries(posts.map((post) => [post.requestId, post]));
    assert.equal(byRequest["req-a"].viewDetails.viewTitle, "View A");
    assert.equal(byRequest["req-b"].viewDetails.viewTitle, "View B");
    assert.ok(fetchUrls.some((url) => url.includes("/lists(guid'list-a')")));
    assert.ok(fetchUrls.some((url) => url.includes("/lists(guid'list-b')")));
    assert.equal(fetchUrls.some((url) => url.includes("wrong-list") || url.includes("wrong-view")), false);
  });
});
