import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const source = await readFile(new URL("../recycleBinWaitBanner.js", import.meta.url), "utf8");
const WAIT_KEY = "__SPO_TOOLKIT_RB_WAIT__";
const WAIT_TS_KEY = "__SPO_TOOLKIT_RB_WAIT_TS__";
const ROOT_ID = "sp-toolkit-rb-wait-root";

class FakeElement {
  constructor(tagName, registry) {
    this.tagName = tagName;
    this.registry = registry;
    this.children = [];
    this.parentNode = null;
    this.style = {};
    this.attributes = {};
    this.textContent = "";
    this._id = "";
  }

  get id() {
    return this._id;
  }

  set id(value) {
    if (this._id) this.registry.delete(this._id);
    this._id = String(value);
    if (this._id) this.registry.set(this._id, this);
  }

  setAttribute(name, value) {
    this.attributes[name] = String(value);
    if (name === "id") this.id = value;
  }

  appendChild(child) {
    child.parentNode = this;
    this.children.push(child);
    return child;
  }

  remove() {
    if (this.parentNode) {
      const i = this.parentNode.children.indexOf(this);
      if (i >= 0) this.parentNode.children.splice(i, 1);
    }
    if (this.id) this.registry.delete(this.id);
    this.parentNode = null;
  }
}

function createHarness(initialStorage, now) {
  const registry = new Map();
  const store = new Map(Object.entries(initialStorage || {}));
  const timers = [];
  const document = {
    body: new FakeElement("body", registry),
    documentElement: new FakeElement("html", registry),
    createElement(tagName) {
      return new FakeElement(tagName, registry);
    },
    getElementById(id) {
      return registry.get(id) || null;
    },
  };
  const sessionStorage = {
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
  const window = {
    setTimeout(fn, delay) {
      timers.push({ fn, delay });
      return timers.length;
    },
  };

  return { document, sessionStorage, timers, window, now };
}

function runBanner(harness) {
  const originalNow = Date.now;
  Date.now = () => harness.now;
  try {
    Function("sessionStorage", "document", "window", source)(
      harness.sessionStorage,
      harness.document,
      harness.window
    );
  } finally {
    Date.now = originalNow;
  }
}

describe("recycleBinWaitBanner", () => {
  it("mounts for a fresh wait marker and clears itself at expiry", () => {
    const h = createHarness({ [WAIT_KEY]: "1", [WAIT_TS_KEY]: "1000" }, 1250);

    runBanner(h);

    assert.ok(h.document.getElementById(ROOT_ID));
    assert.equal(h.timers.length, 1);
    assert.equal(h.timers[0].delay, 59750);

    h.timers[0].fn();

    assert.equal(h.sessionStorage.getItem(WAIT_KEY), null);
    assert.equal(h.sessionStorage.getItem(WAIT_TS_KEY), null);
    assert.equal(h.document.getElementById(ROOT_ID), null);
  });

  it("clears stale wait markers instead of mounting a blocking overlay", () => {
    const h = createHarness({ [WAIT_KEY]: "1", [WAIT_TS_KEY]: "1000" }, 62000);

    runBanner(h);

    assert.equal(h.document.getElementById(ROOT_ID), null);
    assert.equal(h.sessionStorage.getItem(WAIT_KEY), null);
    assert.equal(h.sessionStorage.getItem(WAIT_TS_KEY), null);
    assert.equal(h.timers.length, 0);
  });

  it("treats legacy wait markers without timestamps as stale", () => {
    const h = createHarness({ [WAIT_KEY]: "1" }, 1250);

    runBanner(h);

    assert.equal(h.document.getElementById(ROOT_ID), null);
    assert.equal(h.sessionStorage.getItem(WAIT_KEY), null);
    assert.equal(h.sessionStorage.getItem(WAIT_TS_KEY), null);
  });
});
