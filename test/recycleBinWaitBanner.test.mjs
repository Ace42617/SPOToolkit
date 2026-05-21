import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

const WAIT_KEY = "__SPO_TOOLKIT_RB_WAIT__";
const WAIT_TS_KEY = "__SPO_TOOLKIT_RB_WAIT_TS__";
const ROOT_ID = "sp-toolkit-rb-wait-root";

class Element {
  constructor(tagName, document) {
    this.tagName = tagName;
    this.document = document;
    this.children = [];
    this.style = {};
    this.attributes = {};
    this.id = "";
    this.textContent = "";
  }

  setAttribute(name, value) {
    this.attributes[name] = value;
  }

  appendChild(child) {
    child.parentNode = this;
    this.children.push(child);
    if (child.id) this.document.nodes.set(child.id, child);
    return child;
  }

  remove() {
    if (this.parentNode) {
      this.parentNode.children = this.parentNode.children.filter((child) => child !== this);
    }
    if (this.id) this.document.nodes.delete(this.id);
  }
}

function createDocument() {
  const document = {
    nodes: new Map(),
    createElement(tagName) {
      return new Element(tagName, document);
    },
    getElementById(id) {
      return document.nodes.get(id) || null;
    },
  };
  document.body = new Element("body", document);
  document.documentElement = new Element("html", document);
  return document;
}

function createStorage(entries) {
  const store = new Map(Object.entries(entries));
  return {
    getItem(key) {
      return store.has(key) ? store.get(key) : null;
    },
    setItem(key, value) {
      store.set(key, String(value));
    },
    removeItem(key) {
      store.delete(key);
    },
  };
}

function runBanner(relativePath, { now, entries }) {
  const source = readFileSync(resolve(relativePath), "utf8");
  const document = createDocument();
  const sessionStorage = createStorage(entries);
  const timers = [];
  const window = {
    setTimeout(fn, ms) {
      timers.push({ fn, ms });
      return timers.length;
    },
  };
  const DateMock = { now: () => now };

  Function("sessionStorage", "document", "window", "Date", source)(
    sessionStorage,
    document,
    window,
    DateMock
  );

  return { document, sessionStorage, timers };
}

for (const bannerPath of [
  "recycleBinWaitBanner.js",
  "SP-Developer-Toolkit-Lite/recycleBinWaitBanner.js",
]) {
  describe(`${bannerPath} wait marker expiry`, () => {
    it("mounts for a fresh marker and clears it when the failsafe fires", () => {
      const result = runBanner(bannerPath, {
        now: 1000,
        entries: { [WAIT_KEY]: "1", [WAIT_TS_KEY]: "1000" },
      });

      assert.ok(result.document.getElementById(ROOT_ID));
      assert.equal(result.sessionStorage.getItem(WAIT_KEY), "1");
      assert.equal(result.sessionStorage.getItem(WAIT_TS_KEY), "1000");
      assert.equal(result.timers.length, 1);
      assert.equal(result.timers[0].ms, 30000);

      result.timers[0].fn();

      assert.equal(result.document.getElementById(ROOT_ID), null);
      assert.equal(result.sessionStorage.getItem(WAIT_KEY), null);
      assert.equal(result.sessionStorage.getItem(WAIT_TS_KEY), null);
    });

    it("clears expired or malformed markers without mounting", () => {
      for (const entries of [
        { [WAIT_KEY]: "1", [WAIT_TS_KEY]: "1000" },
        { [WAIT_KEY]: "1" },
      ]) {
        const result = runBanner(bannerPath, { now: 31001, entries });

        assert.equal(result.document.getElementById(ROOT_ID), null);
        assert.equal(result.sessionStorage.getItem(WAIT_KEY), null);
        assert.equal(result.sessionStorage.getItem(WAIT_TS_KEY), null);
        assert.equal(result.timers.length, 0);
      }
    });
  });
}
