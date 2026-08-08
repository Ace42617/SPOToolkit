import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { promisify } from "node:util";
import vm from "node:vm";
import {
  REPORT_LIST_PLAN_PARAMS_SCRIPT_ID,
  REPORT_LIST_PLAN_REQUEST_QUERY,
  REPORT_LIST_PLAN_RESULT_TYPE,
  extractReportListPlanRequestIdFromScriptSrc,
  getReportListPlanParamsScriptId,
  matchesReportListPlanResponse
} from "../lib/reportListPlanRequest.mjs";

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
        reject(new Error("Timed out waiting for injected report list plan results"));
        return;
      }
      setTimeout(check, 0);
    }
    check();
  });
}

describe("reportListPlanRequest helpers", () => {
  it("builds unique params script ids per request", () => {
    assert.equal(getReportListPlanParamsScriptId(null), REPORT_LIST_PLAN_PARAMS_SCRIPT_ID);
    assert.equal(
      getReportListPlanParamsScriptId("req-a"),
      REPORT_LIST_PLAN_PARAMS_SCRIPT_ID + "-req-a"
    );
  });

  it("matches only the outstanding request id", () => {
    const ok = {
      __spcsv: true,
      type: REPORT_LIST_PLAN_RESULT_TYPE,
      requestId: "req-a",
      ok: true,
      entries: [{ listId: "AAA", sitePath: "/sites/a" }]
    };
    assert.equal(matchesReportListPlanResponse(ok, "req-a"), true);
    assert.equal(matchesReportListPlanResponse(ok, "req-b"), false);
    assert.equal(
      matchesReportListPlanResponse(
        { __spcsv: true, type: REPORT_LIST_PLAN_RESULT_TYPE, ok: true, entries: [] },
        "req-a"
      ),
      false
    );
    assert.equal(matchesReportListPlanResponse(null, "req-a"), false);
  });

  it("extracts spcsvRequestId from injected script src", () => {
    assert.equal(
      extractReportListPlanRequestIdFromScriptSrc(
        "chrome-extension://x/reportListPlan.js?spcsvRequestId=req-42"
      ),
      "req-42"
    );
    assert.equal(extractReportListPlanRequestIdFromScriptSrc(""), "");
  });
});

describe("reportListPlan.js request correlation", () => {
  it("reads the params node matching the current script request id and tags results", async () => {
    const source = await readFileAsync(join(root, "reportListPlan.js"), "utf8");
    const posts = [];
    const fetchUrls = [];
    const elements = new Map();
    elements.set(REPORT_LIST_PLAN_PARAMS_SCRIPT_ID, {
      textContent: JSON.stringify({
        requestId: "stale",
        siteUrl: "https://tenant.sharepoint.com/sites/wrong",
        includeSubsites: true,
        librariesOnly: false
      })
    });
    elements.set(getReportListPlanParamsScriptId("req-a"), {
      textContent: JSON.stringify({
        requestId: "req-a",
        siteUrl: "https://tenant.sharepoint.com/sites/hr",
        includeSubsites: false,
        librariesOnly: false
      })
    });
    elements.set(getReportListPlanParamsScriptId("req-b"), {
      textContent: JSON.stringify({
        requestId: "req-b",
        siteUrl: "https://tenant.sharepoint.com/sites/finance",
        includeSubsites: false,
        librariesOnly: false
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
      if (u.includes("/sites/hr/_api/web?") && !u.includes("/lists") && !u.includes("/webs")) {
        return {
          ok: true,
          json: async () => ({
            Title: "HR",
            ServerRelativeUrl: "/sites/hr",
            Url: "https://tenant.sharepoint.com/sites/hr"
          })
        };
      }
      if (u.includes("/sites/finance/_api/web?") && !u.includes("/lists") && !u.includes("/webs")) {
        return {
          ok: true,
          json: async () => ({
            Title: "Finance",
            ServerRelativeUrl: "/sites/finance",
            Url: "https://tenant.sharepoint.com/sites/finance"
          })
        };
      }
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
                DefaultViewUrl: "/sites/hr/Shared Documents/Forms/AllItems.aspx"
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
                DefaultViewUrl: "/sites/finance/Shared Documents/Forms/AllItems.aspx"
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
        "https://extension.test/reportListPlan.js?" +
        REPORT_LIST_PLAN_REQUEST_QUERY +
        "=req-a"
    };
    vm.runInContext(source, context);
    document.currentScript = {
      src:
        "https://extension.test/reportListPlan.js?" +
        REPORT_LIST_PLAN_REQUEST_QUERY +
        "=req-b"
    };
    vm.runInContext(source, context);

    await waitForPosts(posts, 2);
    const byRequest = Object.fromEntries(posts.map((post) => [post.requestId, post]));
    assert.equal(byRequest["req-a"].ok, true);
    assert.equal(byRequest["req-b"].ok, true);
    assert.equal(byRequest["req-a"].entries[0].listTitle, "HR Docs");
    assert.equal(byRequest["req-a"].entries[0].sitePath, "/sites/hr");
    assert.equal(byRequest["req-b"].entries[0].listTitle, "Finance Docs");
    assert.equal(byRequest["req-b"].entries[0].sitePath, "/sites/finance");
    assert.ok(fetchUrls.some((url) => url.includes("/sites/hr/_api/web?")));
    assert.ok(fetchUrls.some((url) => url.includes("/sites/finance/_api/web?")));
    assert.equal(
      fetchUrls.some((url) => url.includes("/sites/wrong")),
      false
    );
  });
});

describe("content.js report list plan correlation wiring", () => {
  const contentSource = readFileSync(join(root, "content.js"), "utf8");
  const planSource = readFileSync(join(root, "reportListPlan.js"), "utf8");

  it("content.js scopes Load-lists params and ignores mismatched responses", () => {
    assert.match(contentSource, /function getReportListPlanParamsScriptId\s*\(/);
    assert.match(contentSource, /function createReportListPlanRequestId\s*\(/);
    assert.match(contentSource, /function attachReportListPlanParamsScript\s*\(/);
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
      /getReportListPlan[\s\S]{0,500}el\.id\s*=\s*"sp-report-list-plan-params-json"/
    );
  });

  it("reportListPlan.js echoes requestId and reads the request-scoped params node", () => {
    assert.match(planSource, /spcsvRequestId/);
    assert.match(planSource, /function getParamsScriptId\s*\(/);
    assert.match(planSource, /PARAMS_SCRIPT_ID\s*\+\s*"-"\s*\+\s*requestId/);
    assert.match(planSource, /msg\.requestId\s*=\s*activeRequestId/);
    assert.match(planSource, /getElementById\(getParamsScriptId\(activeRequestId\)\)/);
  });
});
