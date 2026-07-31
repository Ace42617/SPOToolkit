import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { promisify } from "node:util";
import vm from "node:vm";
import {
  MATRIX_SCAN_PARAMS_SCRIPT_ID,
  MATRIX_SCAN_REQUEST_QUERY,
  MATRIX_SCAN_RESULT_TYPE,
  extractMatrixScanRequestIdFromScriptSrc,
  getMatrixScanParamsScriptId,
  matchesMatrixScanPlanResponse
} from "../lib/matrixScanPlanRequest.mjs";

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
        reject(new Error("Timed out waiting for injected matrix scan plan results"));
        return;
      }
      setTimeout(check, 0);
    }
    check();
  });
}

describe("matrixScanPlanRequest helpers", () => {
  it("builds unique params script ids per request", () => {
    assert.equal(getMatrixScanParamsScriptId(null), MATRIX_SCAN_PARAMS_SCRIPT_ID);
    assert.equal(
      getMatrixScanParamsScriptId("req-a"),
      MATRIX_SCAN_PARAMS_SCRIPT_ID + "-req-a"
    );
  });

  it("matches only the outstanding request id", () => {
    const ok = {
      __spcsv: true,
      type: MATRIX_SCAN_RESULT_TYPE,
      requestId: "req-a",
      ok: true,
      plan: [{ path: "/sites/a" }]
    };
    assert.equal(matchesMatrixScanPlanResponse(ok, "req-a"), true);
    assert.equal(matchesMatrixScanPlanResponse(ok, "req-b"), false);
    assert.equal(
      matchesMatrixScanPlanResponse(
        { __spcsv: true, type: MATRIX_SCAN_RESULT_TYPE, ok: true, plan: [] },
        "req-a"
      ),
      false
    );
    assert.equal(matchesMatrixScanPlanResponse(null, "req-a"), false);
  });

  it("extracts spcsvRequestId from injected script src", () => {
    assert.equal(
      extractMatrixScanRequestIdFromScriptSrc(
        "chrome-extension://x/matrixScanPlan.js?spcsvRequestId=req-42"
      ),
      "req-42"
    );
    assert.equal(extractMatrixScanRequestIdFromScriptSrc(""), "");
  });
});

describe("matrixScanPlan.js request correlation", () => {
  it("reads the params node matching the current script request id and tags results", async () => {
    const source = await readFileAsync(join(root, "matrixScanPlan.js"), "utf8");
    const posts = [];
    const fetchUrls = [];
    const elements = new Map();
    elements.set(MATRIX_SCAN_PARAMS_SCRIPT_ID, {
      textContent: JSON.stringify({
        requestId: "stale",
        siteUrl: "https://tenant.sharepoint.com/sites/wrong",
        includeSubsites: true
      })
    });
    elements.set(getMatrixScanParamsScriptId("req-a"), {
      textContent: JSON.stringify({
        requestId: "req-a",
        siteUrl: "https://tenant.sharepoint.com/sites/hr",
        includeSubsites: false
      })
    });
    elements.set(getMatrixScanParamsScriptId("req-b"), {
      textContent: JSON.stringify({
        requestId: "req-b",
        siteUrl: "https://tenant.sharepoint.com/sites/finance",
        includeSubsites: false
      })
    });

    const document = {
      currentScript: null,
      getElementById(id) {
        return elements.get(id) || null;
      }
    };
    const window = {
      location: {
        href: "https://tenant.sharepoint.com/sites/hr/SitePages/Home.aspx",
        origin: "https://tenant.sharepoint.com"
      },
      _spPageContextInfo: {
        webAbsoluteUrl: "https://tenant.sharepoint.com/sites/hr",
        webServerRelativeUrl: "/sites/hr"
      },
      postMessage(data) {
        posts.push(data);
      }
    };
    const location = window.location;
    const fetch = async (url) => {
      const u = String(url);
      fetchUrls.push(u);
      if (u.includes("/sites/hr/_api/web?") && u.includes("HasUniqueRoleAssignments")) {
        return {
          ok: true,
          json: async () => ({
            Title: "HR",
            ServerRelativeUrl: "/sites/hr",
            Url: "https://tenant.sharepoint.com/sites/hr",
            HasUniqueRoleAssignments: true
          })
        };
      }
      if (u.includes("/sites/finance/_api/web?") && u.includes("HasUniqueRoleAssignments")) {
        return {
          ok: true,
          json: async () => ({
            Title: "Finance",
            ServerRelativeUrl: "/sites/finance",
            Url: "https://tenant.sharepoint.com/sites/finance",
            HasUniqueRoleAssignments: false
          })
        };
      }
      if (u.includes("/_api/web/lists?")) {
        return {
          ok: true,
          json: async () => ({ value: [{ Title: "Documents", ItemCount: 3, BaseTemplate: 101 }] })
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
        "https://extension.test/matrixScanPlan.js?" +
        MATRIX_SCAN_REQUEST_QUERY +
        "=req-a"
    };
    vm.runInContext(source, context);
    document.currentScript = {
      src:
        "https://extension.test/matrixScanPlan.js?" +
        MATRIX_SCAN_REQUEST_QUERY +
        "=req-b"
    };
    vm.runInContext(source, context);

    await waitForPosts(posts, 2);
    const byRequest = Object.fromEntries(posts.map((post) => [post.requestId, post]));
    assert.equal(byRequest["req-a"].ok, true);
    assert.equal(byRequest["req-b"].ok, true);
    assert.equal(byRequest["req-a"].plan[0].path, "/sites/hr");
    assert.equal(byRequest["req-b"].plan[0].path, "/sites/finance");
    assert.ok(fetchUrls.some((url) => url.includes("/sites/hr/_api/web?")));
    assert.ok(fetchUrls.some((url) => url.includes("/sites/finance/_api/web?")));
    assert.equal(
      fetchUrls.some((url) => url.includes("/sites/wrong")),
      false
    );
  });
});

describe("content.js matrix scan plan correlation wiring", () => {
  const contentSource = readFileSync(join(root, "content.js"), "utf8");
  const planSource = readFileSync(join(root, "matrixScanPlan.js"), "utf8");

  it("content.js scopes Load-sites params and ignores mismatched responses", () => {
    assert.match(contentSource, /function getMatrixScanParamsScriptId\s*\(/);
    assert.match(contentSource, /function createMatrixScanRequestId\s*\(/);
    assert.match(contentSource, /function attachMatrixScanParamsScript\s*\(/);
    assert.match(contentSource, /matchesResponse\s*\(\s*data\s*\)\s*\{/);
    assert.match(contentSource, /data\.requestId\s*===\s*requestId/);
    assert.match(
      contentSource,
      /urlSuffix:\s*"\?spcsvRequestId="\s*\+\s*encodeURIComponent\(requestId\)/
    );
    assert.match(contentSource, /options\.matchesResponse/);
    assert.match(contentSource, /options\.afterFinish/);
    assert.doesNotMatch(
      contentSource,
      /getMatrixScanPlan[\s\S]{0,400}el\.id\s*=\s*"sp-matrix-scan-params-json"/
    );
  });

  it("matrixScanPlan.js echoes requestId and reads the request-scoped params node", () => {
    assert.match(planSource, /spcsvRequestId/);
    assert.match(planSource, /function getParamsScriptId\s*\(/);
    assert.match(planSource, /PARAMS_SCRIPT_ID\s*\+\s*"-"\s*\+\s*requestId/);
    assert.match(planSource, /msg\.requestId\s*=\s*activeRequestId/);
    assert.match(planSource, /getElementById\(getParamsScriptId\(activeRequestId\)\)/);
  });
});
