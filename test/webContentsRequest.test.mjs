import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { promisify } from "node:util";
import vm from "node:vm";
import {
  WEB_CONTENTS_PARAMS_ATTR,
  WEB_CONTENTS_REQUEST_ATTR,
  WEB_CONTENTS_REQUEST_QUERY,
  WEB_CONTENTS_RESULT_TYPE,
  extractWebContentsRequestIdFromScriptSrc,
  findWebContentsParamsElement,
  matchesWebContentsResponse,
  shouldSkipSiteContentsWebFetch
} from "../lib/webContentsRequest.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const readFileAsync = promisify(readFile);

function waitForPosts(posts, expectedCount) {
  return new Promise((resolve, reject) => {
    let attempts = 0;
    function check() {
      if (posts.length >= expectedCount) {
        resolve();
        return;
      }
      attempts++;
      if (attempts > 40) {
        reject(new Error("Timed out waiting for injected web contents results"));
        return;
      }
      setTimeout(check, 0);
    }
    check();
  });
}

describe("webContentsRequest helpers", () => {
  it("matches only the outstanding request id", () => {
    const ok = {
      __spcsv: true,
      type: WEB_CONTENTS_RESULT_TYPE,
      requestId: "req-a",
      ok: true,
      lists: [{ title: "Docs" }],
      subsites: []
    };
    assert.equal(matchesWebContentsResponse(ok, "req-a"), true);
    assert.equal(matchesWebContentsResponse(ok, "req-b"), false);
    assert.equal(
      matchesWebContentsResponse(
        { __spcsv: true, type: WEB_CONTENTS_RESULT_TYPE, ok: true, lists: [] },
        "req-a"
      ),
      false
    );
    assert.equal(matchesWebContentsResponse(null, "req-a"), false);
  });

  it("extracts spcsvRequestId from injected script src", () => {
    assert.equal(
      extractWebContentsRequestIdFromScriptSrc(
        "chrome-extension://x/getWebContents.js?spcsvRequestId=req-42"
      ),
      "req-42"
    );
    assert.equal(extractWebContentsRequestIdFromScriptSrc(""), "");
  });

  it("finds the params node for the current request id, not the last node", () => {
    const nodes = [
      {
        getAttribute(name) {
          if (name === WEB_CONTENTS_REQUEST_ATTR) return "req-a";
          if (name === WEB_CONTENTS_PARAMS_ATTR) return "1";
          return null;
        },
        textContent: JSON.stringify({
          requestId: "req-a",
          webUrl: "https://tenant.sharepoint.com/sites/hr"
        })
      },
      {
        getAttribute(name) {
          if (name === WEB_CONTENTS_REQUEST_ATTR) return "req-b";
          if (name === WEB_CONTENTS_PARAMS_ATTR) return "1";
          return null;
        },
        textContent: JSON.stringify({
          requestId: "req-b",
          webUrl: "https://tenant.sharepoint.com/sites/finance"
        })
      }
    ];
    const documentLike = {
      querySelectorAll(sel) {
        assert.match(sel, /data-sp-site-contents-web-params/);
        return nodes;
      }
    };
    const el = findWebContentsParamsElement(documentLike, "req-a");
    assert.equal(el, nodes[0]);
    assert.equal(findWebContentsParamsElement(documentLike, "req-missing"), null);
  });

  it("does not skip fetch after failed loads with empty lists arrays", () => {
    assert.equal(shouldSkipSiteContentsWebFetch(null), false);
    assert.equal(shouldSkipSiteContentsWebFetch({ loading: true }), true);
    assert.equal(
      shouldSkipSiteContentsWebFetch({ loading: false, lists: [], error: "boom" }),
      false
    );
    assert.equal(
      shouldSkipSiteContentsWebFetch({ loading: false, lists: [], error: "" }),
      true
    );
    assert.equal(
      shouldSkipSiteContentsWebFetch({ loading: false, lists: null, error: "timeout" }),
      false
    );
  });
});

describe("getWebContents.js request correlation", () => {
  it("reads the params node matching the current script request id and tags results", async () => {
    const source = await readFileAsync(join(root, "getWebContents.js"), "utf8");
    const posts = [];
    const fetchUrls = [];
    const nodes = [
      {
        getAttribute(name) {
          if (name === "data-request-id") return "stale";
          return name === "data-sp-site-contents-web-params" ? "1" : null;
        },
        textContent: JSON.stringify({
          requestId: "stale",
          webUrl: "https://tenant.sharepoint.com/sites/wrong"
        })
      },
      {
        getAttribute(name) {
          if (name === "data-request-id") return "req-a";
          return name === "data-sp-site-contents-web-params" ? "1" : null;
        },
        textContent: JSON.stringify({
          requestId: "req-a",
          webUrl: "https://tenant.sharepoint.com/sites/hr"
        })
      },
      {
        getAttribute(name) {
          if (name === "data-request-id") return "req-b";
          return name === "data-sp-site-contents-web-params" ? "1" : null;
        },
        textContent: JSON.stringify({
          requestId: "req-b",
          webUrl: "https://tenant.sharepoint.com/sites/finance"
        })
      }
    ];

    const document = {
      currentScript: null,
      querySelectorAll(sel) {
        assert.match(sel, /data-sp-site-contents-web-params/);
        return nodes;
      }
    };
    const window = {
      location: {
        href: "https://tenant.sharepoint.com/sites/hr/SitePages/Home.aspx",
        origin: "https://tenant.sharepoint.com"
      },
      postMessage(data) {
        posts.push(data);
      }
    };
    const location = window.location;
    const fetch = async (url) => {
      const u = String(url);
      fetchUrls.push(u);
      if (u.includes("/sites/hr/_api/web/lists?")) {
        return {
          ok: true,
          json: async () => ({
            value: [
              {
                Id: "11111111-1111-1111-1111-111111111111",
                Title: "HR Docs",
                BaseTemplate: 101,
                ItemCount: 3,
                DefaultViewUrl: "/sites/hr/Shared Documents/Forms/AllItems.aspx",
                LastItemModifiedDate: "2026-08-01T00:00:00Z"
              }
            ]
          })
        };
      }
      if (u.includes("/sites/finance/_api/web/lists?")) {
        return {
          ok: true,
          json: async () => ({
            value: [
              {
                Id: "22222222-2222-2222-2222-222222222222",
                Title: "Finance Docs",
                BaseTemplate: 101,
                ItemCount: 5,
                DefaultViewUrl: "/sites/finance/Shared Documents/Forms/AllItems.aspx",
                LastItemModifiedDate: "2026-08-01T00:00:00Z"
              }
            ]
          })
        };
      }
      if (u.includes("/_api/web/webs?")) {
        return { ok: true, json: async () => ({ value: [] }) };
      }
      throw new Error("Unexpected fetch URL: " + u);
    };

    const context = vm.createContext({
      document,
      window,
      location,
      fetch,
      URL,
      Math
    });
    document.currentScript = {
      src:
        "https://extension.test/getWebContents.js?" +
        WEB_CONTENTS_REQUEST_QUERY +
        "=req-a"
    };
    vm.runInContext(source, context);
    document.currentScript = {
      src:
        "https://extension.test/getWebContents.js?" +
        WEB_CONTENTS_REQUEST_QUERY +
        "=req-b"
    };
    vm.runInContext(source, context);

    await waitForPosts(posts, 2);
    const byRequest = Object.fromEntries(posts.map((post) => [post.requestId, post]));
    assert.equal(byRequest["req-a"].ok, true);
    assert.equal(byRequest["req-b"].ok, true);
    assert.equal(byRequest["req-a"].lists[0].title, "HR Docs");
    assert.equal(byRequest["req-a"].webUrl, "https://tenant.sharepoint.com/sites/hr");
    assert.equal(byRequest["req-b"].lists[0].title, "Finance Docs");
    assert.equal(byRequest["req-b"].webUrl, "https://tenant.sharepoint.com/sites/finance");
    assert.ok(fetchUrls.some((url) => url.includes("/sites/hr/_api/web/lists?")));
    assert.ok(fetchUrls.some((url) => url.includes("/sites/finance/_api/web/lists?")));
    assert.equal(fetchUrls.some((url) => url.includes("/sites/wrong")), false);
  });
});

describe("content.js site contents correlation wiring", () => {
  const contentSource = readFileSync(join(root, "content.js"), "utf8");
  const webSource = readFileSync(join(root, "getWebContents.js"), "utf8");

  it("content.js scopes Site Contents injects and times out stranded waits", () => {
    assert.match(contentSource, /function shouldSkipSiteContentsWebFetch\s*\(/);
    assert.match(contentSource, /function fetchSiteContentsWeb\s*\(/);
    assert.match(
      contentSource,
      /spcsvRequestId="\s*\+\s*encodeURIComponent\(requestId\)/
    );
    assert.match(contentSource, /Timed out loading web contents/);
    assert.match(contentSource, /shouldSkipSiteContentsWebFetch\(cache\)/);
    assert.match(
      contentSource,
      /lists:\s*res\.ok\s*\n?\s*\?\s*\(res\.lists/
    );
    assert.match(
      contentSource,
      /if\s*\(!cache\s*\|\|\s*cache\.error\s*\|\|\s*\(!cache\.lists\s*&&\s*!cache\.loading\)\)/
    );
    assert.doesNotMatch(
      contentSource,
      /if\s*\(cache\s*&&\s*\(cache\.loading\s*\|\|\s*cache\.lists\)\)\s*return;/
    );
  });

  it("getWebContents.js echoes requestId and reads the request-scoped params node", () => {
    assert.match(webSource, /spcsvRequestId/);
    assert.match(webSource, /function findParamsElement\s*\(/);
    assert.match(webSource, /getAttribute\("data-request-id"\)\s*===\s*requestId/);
    assert.match(webSource, /document\.currentScript/);
    assert.doesNotMatch(
      webSource,
      /nodes\.length\s*\?\s*nodes\[nodes\.length\s*-\s*1\]\s*:\s*null;\s*\n\s*if\s*\(el\s*&&\s*el\.textContent\)\s*return\s*JSON\.parse/
    );
  });
});
