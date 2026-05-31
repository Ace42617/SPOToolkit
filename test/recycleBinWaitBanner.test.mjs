import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";
import vm from "node:vm";

const KEY = "__SPO_TOOLKIT_RB_WAIT__";
const TS_KEY = "__SPO_TOOLKIT_RB_WAIT_TS__";
const ROOT_ID = "sp-toolkit-rb-wait-root";

function makeElement(document, tag) {
  return {
    tag,
    id: "",
    style: {},
    children: [],
    textContent: "",
    setAttribute(name, value) {
      this[name] = value;
    },
    appendChild(child) {
      this.children.push(child);
      if (child.id) document.elements.set(child.id, child);
      return child;
    },
    remove() {
      if (this.id) document.elements.delete(this.id);
    },
  };
}

function runBanner(path, sessionEntries, now = 1_000_000) {
  const source = fs.readFileSync(path, "utf8");
  const session = new Map(Object.entries(sessionEntries));
  const timers = [];
  const document = {
    elements: new Map(),
    createElement(tag) {
      return makeElement(document, tag);
    },
    getElementById(id) {
      return this.elements.get(id) || null;
    },
  };
  document.body = makeElement(document, "body");
  document.documentElement = makeElement(document, "html");

  const context = {
    Date: class extends Date {
      static now() {
        return now;
      }
    },
    document,
    sessionStorage: {
      getItem(key) {
        return session.has(key) ? session.get(key) : null;
      },
      setItem(key, value) {
        session.set(key, String(value));
      },
      removeItem(key) {
        session.delete(key);
      },
    },
    window: {
      setTimeout(fn, ms) {
        timers.push({ fn, ms });
        return timers.length;
      },
    },
  };
  vm.runInNewContext(source, context, { filename: path });
  return { document, session, timers };
}

for (const path of ["recycleBinWaitBanner.js", "SP-Developer-Toolkit-Lite/recycleBinWaitBanner.js"]) {
  test(`${path} ignores and clears wait markers without timestamps`, () => {
    const { document, session } = runBanner(path, { [KEY]: "1" });

    assert.equal(document.getElementById(ROOT_ID), null);
    assert.equal(session.has(KEY), false);
    assert.equal(session.has(TS_KEY), false);
  });

  test(`${path} mounts only for fresh wait markers and self-clears`, () => {
    const { document, session, timers } = runBanner(path, {
      [KEY]: "1",
      [TS_KEY]: "1000000",
    });

    assert.ok(document.getElementById(ROOT_ID));
    assert.equal(timers.length, 1);
    assert.ok(timers[0].ms > 0);

    timers[0].fn();
    assert.equal(document.getElementById(ROOT_ID), null);
    assert.equal(session.has(KEY), false);
    assert.equal(session.has(TS_KEY), false);
  });

  test(`${path} clears expired wait markers`, () => {
    const { document, session, timers } = runBanner(path, {
      [KEY]: "1",
      [TS_KEY]: "900000",
    });

    assert.equal(document.getElementById(ROOT_ID), null);
    assert.equal(timers.length, 0);
    assert.equal(session.has(KEY), false);
    assert.equal(session.has(TS_KEY), false);
  });
}
