import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { promisify } from "node:util";
import vm from "node:vm";
import {
  COLUMN_CREATE_PARAMS_SCRIPT_ID,
  COLUMN_CREATE_REQUEST_QUERY,
  COLUMN_CREATE_RESULT_TYPE,
  extractColumnCreateRequestIdFromScriptSrc,
  getColumnCreateParamsScriptId,
  matchesColumnCreateResponse
} from "../lib/createColumnRequest.mjs";

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
        reject(new Error("Timed out waiting for injected create-column results"));
        return;
      }
      setTimeout(check, 0);
    }
    check();
  });
}

describe("createColumnRequest helpers", () => {
  it("builds unique params script ids per request", () => {
    assert.equal(getColumnCreateParamsScriptId(null), COLUMN_CREATE_PARAMS_SCRIPT_ID);
    assert.equal(
      getColumnCreateParamsScriptId("req-a"),
      COLUMN_CREATE_PARAMS_SCRIPT_ID + "-req-a"
    );
  });

  it("matches only the outstanding request id", () => {
    const ok = {
      __spcsv: true,
      type: COLUMN_CREATE_RESULT_TYPE,
      requestId: "req-a",
      ok: true,
      internalName: "Department",
      title: "Department"
    };
    assert.equal(matchesColumnCreateResponse(ok, "req-a"), true);
    assert.equal(matchesColumnCreateResponse(ok, "req-b"), false);
    assert.equal(
      matchesColumnCreateResponse(
        { __spcsv: true, type: COLUMN_CREATE_RESULT_TYPE, ok: true, internalName: "Department" },
        "req-a"
      ),
      false
    );
    assert.equal(matchesColumnCreateResponse(null, "req-a"), false);
  });

  it("extracts spcsvRequestId from injected script src", () => {
    assert.equal(
      extractColumnCreateRequestIdFromScriptSrc(
        "chrome-extension://x/createColumn.js?spcsvRequestId=req-42"
      ),
      "req-42"
    );
    assert.equal(extractColumnCreateRequestIdFromScriptSrc(""), "");
  });
});

describe("createColumn.js request correlation", () => {
  it("reads the params node matching the current script request id and tags results", async () => {
    const source = await readFileAsync(join(root, "createColumn.js"), "utf8");
    const posts = [];
    const createBodies = [];
    const elements = new Map();
    const staleSchema = '<Field Type="Boolean" Name="StaleFlag" DisplayName="Stale" />';
    const schemaA = '<Field Type="Choice" Name="Department" DisplayName="Department" />';
    const schemaB = '<Field Type="User" Name="Owner" DisplayName="Owner" />';
    elements.set(COLUMN_CREATE_PARAMS_SCRIPT_ID, {
      textContent: JSON.stringify({
        requestId: "stale",
        siteUrl: "https://tenant.sharepoint.com/sites/hr",
        listId: "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        target: "list",
        schemaXml: staleSchema,
        siteFieldInternal: "StaleFlag"
      })
    });
    elements.set(getColumnCreateParamsScriptId("req-a"), {
      textContent: JSON.stringify({
        requestId: "req-a",
        siteUrl: "https://tenant.sharepoint.com/sites/hr",
        listId: "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        target: "list",
        schemaXml: schemaA,
        siteFieldInternal: "Department"
      })
    });
    elements.set(getColumnCreateParamsScriptId("req-b"), {
      textContent: JSON.stringify({
        requestId: "req-b",
        siteUrl: "https://tenant.sharepoint.com/sites/hr",
        listId: "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        target: "list",
        schemaXml: schemaB,
        siteFieldInternal: "Owner"
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
        href: "https://tenant.sharepoint.com/sites/hr/Lists/Staff/AllItems.aspx",
        origin: "https://tenant.sharepoint.com"
      },
      postMessage(data) {
        posts.push(data);
      }
    };
    const fetch = async (url, opts) => {
      const u = String(url);
      if (u.endsWith("/_api/contextinfo")) {
        return {
          ok: true,
          json: async () => ({ FormDigestValue: "digest" })
        };
      }
      if (u.includes("/fields/createfieldasxml")) {
        const body = opts && opts.body ? JSON.parse(opts.body) : {};
        const xml = (body.parameters && body.parameters.SchemaXml) || "";
        createBodies.push(xml);
        const nameMatch = xml.match(/Name="([^"]+)"/i);
        const titleMatch = xml.match(/DisplayName="([^"]+)"/i);
        const typeMatch = xml.match(/Type="([^"]+)"/i);
        const internalName = nameMatch ? nameMatch[1] : "";
        return {
          ok: true,
          text: async () =>
            JSON.stringify({
              d: {
                InternalName: internalName,
                Title: titleMatch ? titleMatch[1] : internalName,
                TypeAsString: typeMatch ? typeMatch[1] : ""
              }
            })
        };
      }
      throw new Error("Unexpected fetch URL: " + u);
    };

    const context = vm.createContext({
      document,
      window,
      location: window.location,
      fetch,
      URL,
      Promise,
      setTimeout,
      clearTimeout,
      JSON,
      Object,
      String,
      Error
    });
    document.currentScript = {
      src:
        "https://extension.test/createColumn.js?" +
        COLUMN_CREATE_REQUEST_QUERY +
        "=req-a"
    };
    vm.runInContext(source, context);
    document.currentScript = {
      src:
        "https://extension.test/createColumn.js?" +
        COLUMN_CREATE_REQUEST_QUERY +
        "=req-b"
    };
    vm.runInContext(source, context);

    await waitForPosts(posts, 2);
    const byRequest = Object.fromEntries(posts.map((post) => [post.requestId, post]));
    assert.equal(byRequest["req-a"].ok, true);
    assert.equal(byRequest["req-b"].ok, true);
    assert.equal(byRequest["req-a"].internalName, "Department");
    assert.equal(byRequest["req-b"].internalName, "Owner");
    assert.equal(createBodies.includes(schemaA), true);
    assert.equal(createBodies.includes(schemaB), true);
    assert.equal(createBodies.includes(staleSchema), false);
  });
});

describe("content.js create-column correlation wiring", () => {
  const contentSource = readFileSync(join(root, "content.js"), "utf8");
  const createSource = readFileSync(join(root, "createColumn.js"), "utf8");

  it("content.js scopes create-column params and ignores mismatched responses", () => {
    assert.match(contentSource, /function getColumnCreateParamsScriptId\s*\(/);
    assert.match(contentSource, /function createColumnRequestId\s*\(/);
    assert.match(contentSource, /function attachSpColumnCreateParamsScript\s*\(/);
    assert.match(contentSource, /attachSpColumnCreateParamsScript\(message,\s*requestId\)/);
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
      /createColumn[\s\S]{0,800}el\.id\s*=\s*SP_COLUMN_CREATE_PARAMS_SCRIPT_ID;/
    );
  });

  it("createColumn.js echoes requestId and reads the request-scoped params node", () => {
    assert.match(createSource, /spcsvRequestId/);
    assert.match(createSource, /function getParamsScriptId\s*\(/);
    assert.match(createSource, /PARAMS_SCRIPT_ID\s*\+\s*"-"\s*\+\s*requestId/);
    assert.match(createSource, /msg\.requestId\s*=\s*activeRequestId/);
    assert.match(createSource, /getElementById\(getParamsScriptId\(activeRequestId\)\)/);
  });
});
