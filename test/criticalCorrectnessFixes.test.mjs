import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import { constants } from "node:fs";
import { test } from "node:test";

const root = new URL("../", import.meta.url);

async function readText(path) {
  return readFile(new URL(path, root), "utf8");
}

test("view formatter framing rules are not installed as global static rules", async () => {
  const manifest = JSON.parse(await readText("manifest.json"));
  assert.equal(manifest.declarative_net_request, undefined);
  await assert.rejects(access(new URL("rules/view-formatter-iframe.json", root), constants.F_OK));

  const background = await readText("background.js");
  assert.match(background, /updateSessionRules/);
  assert.match(background, /tabIds:\s*\[tabId\]/);
});

test("view formatter save fails closed when the current view id is unknown", async () => {
  const source = await readText("view-formatter.js");
  assert.match(source, /Could not determine the current view id/);
  assert.doesNotMatch(source, /DefaultView\?\$select=Id/);
});

function makeStorage(entries) {
  const map = new Map(Object.entries(entries || {}));
  return {
    getItem(key) {
      return map.has(key) ? map.get(key) : null;
    },
    setItem(key, value) {
      map.set(key, String(value));
    },
    removeItem(key) {
      map.delete(key);
    },
    has(key) {
      return map.has(key);
    },
  };
}

function makeDocument() {
  const nodes = new Map();
  function store(node) {
    if (node.id) nodes.set(node.id, node);
  }
  function makeHost() {
    return {
      appendChild(node) {
        store(node);
      },
    };
  }
  return {
    body: makeHost(),
    documentElement: makeHost(),
    getElementById(id) {
      return nodes.get(id) || null;
    },
    createElement(tag) {
      const node = {
        tagName: tag,
        id: "",
        style: {},
        children: [],
        textContent: "",
        setAttribute() {},
        appendChild(child) {
          this.children.push(child);
          store(child);
        },
        remove() {
          if (this.id) nodes.delete(this.id);
        },
      };
      return node;
    },
  };
}

async function runRecycleBanner(storageEntries, now) {
  const source = await readText("recycleBinWaitBanner.js");
  const storage = makeStorage(storageEntries);
  const document = makeDocument();
  const timeouts = [];
  const window = {
    setTimeout(fn, ms) {
      timeouts.push({ fn, ms });
      return timeouts.length;
    },
  };
  const DateShim = { now: () => now };
  Function("sessionStorage", "document", "window", "Date", source)(storage, document, window, DateShim);
  return { storage, document, timeouts };
}

test("recycle-bin wait overlay clears stale session markers instead of remounting forever", async () => {
  const { storage, document, timeouts } = await runRecycleBanner(
    {
      __SPO_TOOLKIT_RB_WAIT__: "1",
      __SPO_TOOLKIT_RB_WAIT_TS__: "1000",
    },
    47001
  );

  assert.equal(document.getElementById("sp-toolkit-rb-wait-root"), null);
  assert.equal(storage.has("__SPO_TOOLKIT_RB_WAIT__"), false);
  assert.equal(storage.has("__SPO_TOOLKIT_RB_WAIT_TS__"), false);
  assert.equal(timeouts.length, 0);
});

test("recycle-bin wait overlay mounts fresh markers with an expiry timer", async () => {
  const { storage, document, timeouts } = await runRecycleBanner(
    {
      __SPO_TOOLKIT_RB_WAIT__: "1",
      __SPO_TOOLKIT_RB_WAIT_TS__: "1000",
    },
    2000
  );

  assert.ok(document.getElementById("sp-toolkit-rb-wait-root"));
  assert.equal(storage.has("__SPO_TOOLKIT_RB_WAIT__"), true);
  assert.equal(storage.has("__SPO_TOOLKIT_RB_WAIT_TS__"), true);
  assert.equal(timeouts.length, 1);
  assert.equal(timeouts[0].ms, 44000);
});
