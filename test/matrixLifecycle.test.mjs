import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";
import vm from "node:vm";

const root = new URL("../", import.meta.url);
const backgroundSource = readFileSync(new URL("background.js", root), "utf8");

function loadBackground() {
  const runtimeMessages = [];
  const tabMessages = [];
  const removedTabs = [];
  const timers = [];
  const session = {};
  let runtimeListener = null;
  let nextTabId = 100;
  let nextRunId = 1;

  function event(capture) {
    return {
      addListener(fn) {
        if (capture) capture(fn);
      },
      removeListener() {},
    };
  }

  const chrome = {
    action: { openPopup: async () => {} },
    commands: { onCommand: event() },
    runtime: {
      lastError: null,
      getURL: (path) => `chrome-extension://test/${path}`,
      onConnect: event(),
      onMessage: event((fn) => {
        runtimeListener = fn;
      }),
    },
    scripting: { executeScript(_opts, callback) { callback?.(); } },
    storage: {
      local: {
        get(_key, callback) { callback({}); },
        set(_value, callback) { callback?.(); },
      },
      session: {
        get(key, callback) { callback({ [key]: session[key] }); },
        set(value, callback) {
          Object.assign(session, value);
          callback?.();
        },
        remove(key, callback) {
          delete session[key];
          callback?.();
        },
      },
    },
    tabs: {
      create(_options, callback) { callback?.({ id: nextTabId++ }); },
      get(id, callback) {
        callback({ id, url: "https://contoso.sharepoint.com/sites/test/Lists/tasks/AllItems.aspx" });
      },
      onUpdated: event(),
      remove(id, callback) {
        removedTabs.push(id);
        callback?.();
      },
      sendMessage(id, message, callback) {
        tabMessages.push({ id, message });
        callback?.({ ok: true, running: true });
      },
      update(_id, _options, callback) { callback?.(); },
    },
    webNavigation: { onCommitted: event() },
  };

  const context = {
    URL,
    chrome,
    console,
    crypto: { randomUUID: () => `run-${nextRunId++}` },
    Math,
    setInterval() { return 1; },
    clearInterval() {},
    setTimeout(fn, ms) {
      const timer = { fn, ms, cancelled: false };
      timers.push(timer);
      return timer;
    },
    clearTimeout(timer) {
      if (timer) timer.cancelled = true;
    },
  };
  vm.runInNewContext(backgroundSource, context, { filename: "background.js" });
  assert.equal(typeof runtimeListener, "function");

  function send(message, tabId = 1) {
    let response;
    runtimeMessages.push({ message, tabId });
    runtimeListener(message, { tab: { id: tabId } }, (value) => {
      response = value;
    });
    return response;
  }

  function flushTimers(ms) {
    for (const timer of timers) {
      if (!timer.cancelled && timer.ms === ms) {
        timer.cancelled = true;
        timer.fn();
      }
    }
  }

  function progress() {
    flushTimers(120);
    let value;
    runtimeListener({ type: "SPCSVExportProgressGet" }, {}, (result) => {
      value = result;
    });
    return value;
  }

  return { progress, removedTabs, send, tabMessages, flushTimers };
}

test("late matrix lifecycle messages cannot terminate a newer run", () => {
  const bg = loadBackground();

  bg.send({
    type: "SPCSVStartMatrixExport",
    exportMessage: {
      report: "permissionsMatrix",
      siteUrl: "https://contoso.sharepoint.com/sites/test",
    },
  });
  const firstStart = bg.tabMessages.find((entry) => entry.message.action === "matrixWorkerLockAndRun");
  assert.equal(firstStart.message.exportMessage.matrixRunId, "run-1");

  bg.send({
    type: "SPCSVExportDone",
    report: "permissionsMatrix",
    matrixRunId: "run-1",
    success: true,
  });

  bg.send({
    type: "SPCSVStartMatrixExport",
    exportMessage: {
      report: "permissionsMatrix",
      siteUrl: "https://contoso.sharepoint.com/sites/test",
    },
  });
  const starts = bg.tabMessages.filter((entry) => entry.message.action === "matrixWorkerLockAndRun");
  assert.equal(starts[1].message.exportMessage.matrixRunId, "run-2");

  // The delayed close from run 1 must not remove a tab now running run 2.
  bg.flushTimers(5000);
  assert.deepEqual(bg.removedTabs, []);

  const staleResponse = bg.send({
    type: "SPCSVExportDone",
    report: "permissionsMatrix",
    matrixRunId: "run-1",
    success: true,
  });
  assert.equal(staleResponse.ignored, true);
  assert.equal(bg.progress().active, true);
  assert.equal(bg.progress().matrixRunId, "run-2");

  const wrongTabResponse = bg.send({
    type: "SPCSVExportProgress",
    report: "permissionsMatrix",
    matrixRunId: "run-2",
    percent: 90,
  }, 2);
  assert.equal(wrongTabResponse.ignored, true);

  const csvDoneResponse = bg.send({
    type: "SPCSVExportDone",
    report: "exportCSV",
    success: true,
  }, 3);
  assert.equal(csvDoneResponse.ignored, true);
  assert.equal(bg.progress().active, true);

  bg.send({
    type: "SPCSVExportDone",
    report: "permissionsMatrix",
    matrixRunId: "run-2",
    success: true,
  });
  assert.equal(bg.progress().active, false);
  const latePulseResponse = bg.send({
    type: "SPCSVExportProgress",
    report: "permissionsMatrix",
    matrixRunId: "run-2",
    percent: 99,
  });
  assert.equal(latePulseResponse.ignored, true);
  assert.equal(bg.progress().active, false);
  assert.equal(
    bg.tabMessages.filter((entry) => entry.message.action === "matrixWorkerFinish").length,
    2
  );
});

test("matrix run id is preserved through page and content lifecycle messages", () => {
  const content = readFileSync(new URL("content.js", root), "utf8");
  const exporter = readFileSync(new URL("permissionsMatrixExport.js", root), "utf8");

  assert.match(content, /params\.matrixRunId = activeMatrixRunId/);
  assert.match(content, /matrixRunId: d\.matrixRunId \|\| undefined/);
  assert.match(content, /d\.matrixRunId !== activeMatrixRunId/);
  assert.match(exporter, /matrixRunId: MATRIX_RUN_ID/);
  assert.match(exporter, /report: "permissionsMatrix"/);
});
