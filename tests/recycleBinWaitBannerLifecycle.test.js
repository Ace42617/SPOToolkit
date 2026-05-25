import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { describe, it } from "node:test";
import vm from "node:vm";

const KEY = "__SPO_TOOLKIT_RB_WAIT__";
const TS_KEY = "__SPO_TOOLKIT_RB_WAIT_TS__";
const MAX_WAIT_MS = 60000;

async function loadBanner(relativePath) {
  return readFile(new URL(relativePath, import.meta.url), "utf8");
}

function runBanner(source, now, initialStorage) {
  const store = new Map(Object.entries(initialStorage || {}));
  const elements = new Map();
  const timers = [];

  function createElement(tagName) {
    return {
      tagName: String(tagName).toUpperCase(),
      id: "",
      attributes: {},
      children: [],
      style: {},
      textContent: "",
      setAttribute(name, value) {
        this.attributes[name] = String(value);
      },
      appendChild(child) {
        this.children.push(child);
        if (child.id) elements.set(child.id, child);
        return child;
      },
      remove() {
        if (this.id) elements.delete(this.id);
      },
    };
  }

  const document = {
    body: createElement("body"),
    documentElement: createElement("html"),
    createElement,
    getElementById(id) {
      return elements.get(id) || null;
    },
  };

  const context = {
    Date: { now: () => now },
    Number,
    Math,
    Object,
    document,
    sessionStorage: {
      getItem(key) {
        return store.has(key) ? store.get(key) : null;
      },
      setItem(key, value) {
        store.set(key, String(value));
      },
      removeItem(key) {
        store.delete(key);
      },
    },
    setTimeout(fn, delay) {
      timers.push({ fn, delay });
      return timers.length;
    },
  };

  vm.runInNewContext(source, context);
  return { document, store, timers };
}

describe("recycle bin wait banner lifecycle", () => {
  for (const [label, relativePath] of [
    ["full", "../recycleBinWaitBanner.js"],
    ["lite", "../SP-Developer-Toolkit-Lite/recycleBinWaitBanner.js"],
  ]) {
    it(`${label}: clears stale wait state instead of mounting`, async () => {
      const source = await loadBanner(relativePath);
      const result = runBanner(source, 100000, {
        [KEY]: "1",
        [TS_KEY]: String(100000 - MAX_WAIT_MS - 1),
      });

      assert.equal(result.document.getElementById("sp-toolkit-rb-wait-root"), null);
      assert.equal(result.store.has(KEY), false);
      assert.equal(result.store.has(TS_KEY), false);
    });

    it(`${label}: self-clears active wait state`, async () => {
      const source = await loadBanner(relativePath);
      const result = runBanner(source, 100000, {
        [KEY]: "1",
        [TS_KEY]: String(100000),
      });

      assert.ok(result.document.getElementById("sp-toolkit-rb-wait-root"));
      assert.equal(result.timers.length, 1);
      assert.equal(result.timers[0].delay, MAX_WAIT_MS);

      result.timers[0].fn();

      assert.equal(result.document.getElementById("sp-toolkit-rb-wait-root"), null);
      assert.equal(result.store.has(KEY), false);
      assert.equal(result.store.has(TS_KEY), false);
    });
  }
});
