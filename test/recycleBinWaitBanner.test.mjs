import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const KEY = "__SPO_TOOLKIT_RB_WAIT__";
const TS_KEY = "__SPO_TOOLKIT_RB_WAIT_TS__";
const ROOT_ID = "sp-toolkit-rb-wait-root";

const bannerFiles = [
  ["full", new URL("../recycleBinWaitBanner.js", import.meta.url)],
  ["lite", new URL("../SP-Developer-Toolkit-Lite/recycleBinWaitBanner.js", import.meta.url)],
];

function createElement(tagName, nodes) {
  return {
    tagName,
    id: "",
    style: {},
    children: [],
    attributes: {},
    textContent: "",
    parentNode: null,
    setAttribute(name, value) {
      this.attributes[name] = String(value);
    },
    appendChild(child) {
      child.parentNode = this;
      this.children.push(child);
      if (child.id) nodes.set(child.id, child);
      return child;
    },
    remove() {
      if (this.id) nodes.delete(this.id);
      if (this.parentNode) {
        this.parentNode.children = this.parentNode.children.filter((child) => child !== this);
      }
    },
  };
}

async function runBanner(fileUrl, initialStorage, now) {
  const source = await readFile(fileUrl, "utf8");
  const storage = new Map(Object.entries(initialStorage));
  const nodes = new Map();
  const timers = [];
  const document = {
    body: createElement("body", nodes),
    documentElement: createElement("html", nodes),
    createElement(tagName) {
      return createElement(tagName, nodes);
    },
    getElementById(id) {
      return nodes.get(id) || null;
    },
  };
  const context = {
    Date: { now: () => now.value },
    Object,
    Number,
    String,
    Infinity,
    isFinite,
    sessionStorage: {
      getItem(key) {
        return storage.has(key) ? storage.get(key) : null;
      },
      setItem(key, value) {
        storage.set(key, String(value));
      },
      removeItem(key) {
        storage.delete(key);
      },
    },
    document,
    window: {
      setTimeout(callback, delay) {
        timers.push({ callback, delay });
        return timers.length;
      },
    },
  };
  vm.runInNewContext(source, context, { filename: fileUrl.pathname });
  return { document, storage, timers };
}

for (const [name, fileUrl] of bannerFiles) {
  test(`${name} recycle-bin wait banner clears legacy markers without a timestamp`, async () => {
    const now = { value: 1_000_000 };
    const { document, storage } = await runBanner(fileUrl, { [KEY]: "1" }, now);

    assert.equal(document.getElementById(ROOT_ID), null);
    assert.equal(storage.has(KEY), false);
    assert.equal(storage.has(TS_KEY), false);
  });

  test(`${name} recycle-bin wait banner removes itself after marker expiry`, async () => {
    const now = { value: 1_000_000 };
    const { document, storage, timers } = await runBanner(
      fileUrl,
      { [KEY]: "1", [TS_KEY]: String(now.value) },
      now
    );

    assert.ok(document.getElementById(ROOT_ID));
    assert.equal(timers.length, 1);

    now.value += 45_001;
    timers[0].callback();

    assert.equal(document.getElementById(ROOT_ID), null);
    assert.equal(storage.has(KEY), false);
    assert.equal(storage.has(TS_KEY), false);
  });
}
