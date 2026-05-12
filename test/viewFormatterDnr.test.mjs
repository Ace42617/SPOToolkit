import { describe, it } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");

function loadBackground() {
  const listeners = {};
  const calls = {
    create: [],
    update: [],
    remove: [],
    dnr: [],
  };
  const chrome = {
    runtime: {
      lastError: null,
      getURL: (p) => "chrome-extension://extension-id/" + p,
      onMessage: {
        addListener(fn) {
          listeners.message = fn;
        },
      },
    },
    commands: {
      onCommand: {
        addListener(fn) {
          listeners.command = fn;
        },
      },
    },
    storage: {
      local: {
        get() {},
        set() {},
      },
    },
    action: {
      openPopup() {
        return Promise.resolve();
      },
    },
    tabs: {
      create(options, cb) {
        calls.create.push(options);
        cb({ id: 321 });
      },
      update(tabId, options, cb) {
        calls.update.push({ tabId, options });
        if (cb) cb({ id: tabId });
      },
      remove(tabId, cb) {
        calls.remove.push(tabId);
        if (cb) cb();
      },
      get() {},
      onUpdated: {
        addListener() {},
        removeListener() {},
      },
      onRemoved: {
        addListener(fn) {
          listeners.removed = fn;
        },
      },
    },
    webNavigation: {
      onCommitted: {
        addListener() {},
        removeListener() {},
      },
    },
    scripting: {
      executeScript() {},
    },
    declarativeNetRequest: {
      updateSessionRules(options, cb) {
        calls.dnr.push(options);
        if (cb) cb();
      },
    },
  };

  const source = fs.readFileSync(path.join(root, "background.js"), "utf8");
  vm.runInNewContext(source, { chrome, URL, setTimeout, clearTimeout }, { filename: "background.js" });
  return { calls, listeners };
}

describe("View formatter DNR rules", () => {
  it("does not ship static SharePoint header-stripping rules", () => {
    const manifest = JSON.parse(fs.readFileSync(path.join(root, "manifest.json"), "utf8"));
    assert.equal(manifest.declarative_net_request, undefined);
  });

  it("installs header-stripping only for the formatter tab and preview host", () => {
    const { calls, listeners } = loadBackground();
    let response;
    const previewUrl = "https://contoso.sharepoint.com/sites/demo/Lists/AllItems.aspx";

    const keepAlive = listeners.message(
      { type: "SPOToolkitOpenViewFormatter", previewUrl },
      { tab: { id: 11, url: previewUrl } },
      (r) => {
        response = r;
      }
    );

    assert.equal(keepAlive, true);
    assert.equal(response.ok, true);
    assert.equal(calls.create.length, 1);
    assert.equal(calls.create[0].url, "about:blank");

    assert.equal(calls.dnr.length, 1);
    const rule = calls.dnr[0].addRules[0];
    assert.deepEqual(Array.from(calls.dnr[0].removeRuleIds), [100321]);
    assert.equal(rule.id, 100321);
    assert.equal(rule.condition.urlFilter, "||contoso.sharepoint.com^");
    assert.deepEqual(Array.from(rule.condition.resourceTypes), ["sub_frame"]);
    assert.deepEqual(Array.from(rule.condition.tabIds), [321]);
    assert.deepEqual(
      Array.from(rule.action.responseHeaders, (h) => h.header),
      ["x-frame-options", "content-security-policy", "content-security-policy-report-only"]
    );

    assert.equal(calls.update.length, 1);
    assert.equal(calls.update[0].tabId, 321);
    const formatterUrl = new URL(calls.update[0].options.url);
    assert.equal(formatterUrl.protocol, "chrome-extension:");
    assert.equal(formatterUrl.pathname, "/view-formatter.html");
    assert.equal(formatterUrl.searchParams.get("src"), previewUrl);
    assert.equal(formatterUrl.searchParams.get("tabId"), "11");

    listeners.removed(321);
    assert.deepEqual(Array.from(calls.dnr[1].removeRuleIds), [100321]);
  });

  it("rejects lookalike SharePoint hosts before creating a preview tab", () => {
    const { calls, listeners } = loadBackground();
    let response;

    const keepAlive = listeners.message(
      {
        type: "SPOToolkitOpenViewFormatter",
        previewUrl: "https://contoso.sharepoint.com.evil.example/sites/demo",
      },
      { tab: { id: 11, url: "https://contoso.sharepoint.com.evil.example/sites/demo" } },
      (r) => {
        response = r;
      }
    );

    assert.equal(keepAlive, true);
    assert.equal(response.ok, false);
    assert.equal(response.error, "Need a SharePoint URL");
    assert.equal(calls.create.length, 0);
    assert.equal(calls.dnr.length, 0);
    assert.equal(calls.update.length, 0);
  });
});
