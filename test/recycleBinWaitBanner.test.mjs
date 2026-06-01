import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const KEY = "__SPO_TOOLKIT_RB_WAIT__";
const TS_KEY = "__SPO_TOOLKIT_RB_WAIT_TS__";
const ROOT_ID = "sp-toolkit-rb-wait-root";

function makeElement(tagName, documentRef) {
  return {
    tagName,
    id: "",
    style: {},
    children: [],
    textContent: "",
    innerHTML: "",
    setAttribute(name, value) {
      this[name] = value;
    },
    appendChild(child) {
      this.children.push(child);
      if (child.id) documentRef.nodes.set(child.id, child);
      return child;
    },
    remove() {
      if (this.id) documentRef.nodes.delete(this.id);
      this.removed = true;
    },
  };
}

function runBanner(path, storageEntries, now = 100000) {
  const store = new Map(Object.entries(storageEntries));
  const timers = [];
  const documentRef = {
    nodes: new Map(),
    createElement(tagName) {
      return makeElement(tagName, documentRef);
    },
    getElementById(id) {
      return documentRef.nodes.get(id) || null;
    },
  };
  documentRef.body = makeElement("body", documentRef);
  documentRef.documentElement = makeElement("html", documentRef);

  const context = {
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
    document: documentRef,
    window: {
      setTimeout(fn, delay) {
        timers.push({ fn, delay });
        return timers.length;
      },
    },
    Date: {
      now() {
        return now;
      },
    },
    Number,
    Object,
    String,
    Math,
    isFinite,
  };
  vm.createContext(context);
  vm.runInContext(readFileSync(path, "utf8"), context, { filename: path });
  return { document: documentRef, store, timers };
}

describe("recycleBinWaitBanner", () => {
  for (const path of ["recycleBinWaitBanner.js", "SP-Developer-Toolkit-Lite/recycleBinWaitBanner.js"]) {
    it(`${path} mounts only for a fresh primed session and expires it`, () => {
      const result = runBanner(path, { [KEY]: "1", [TS_KEY]: "90000" });

      assert.ok(result.document.getElementById(ROOT_ID), "fresh marker should mount overlay");
      assert.equal(result.timers.length, 1, "fresh marker should schedule teardown");
      assert.ok(result.timers[0].delay > 0, "teardown should wait until marker expiry");

      result.timers[0].fn();
      assert.equal(result.document.getElementById(ROOT_ID), null);
      assert.equal(result.store.has(KEY), false);
      assert.equal(result.store.has(TS_KEY), false);
    });

    it(`${path} clears stale or legacy markers without mounting`, () => {
      const stale = runBanner(path, { [KEY]: "1", [TS_KEY]: "1" });
      assert.equal(stale.document.getElementById(ROOT_ID), null);
      assert.equal(stale.store.has(KEY), false);
      assert.equal(stale.store.has(TS_KEY), false);

      const legacy = runBanner(path, { [KEY]: "1" });
      assert.equal(legacy.document.getElementById(ROOT_ID), null);
      assert.equal(legacy.store.has(KEY), false);
    });
  }
});
