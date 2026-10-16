import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

const source = readFileSync(new URL("../exportCSV.js", import.meta.url), "utf8");

function sliceFunction(name) {
  const start = source.indexOf("async function " + name + "(");
  assert.notEqual(start, -1, name + " should be async");
  const next = source.indexOf("\n  async function ", start + 1);
  const nextSync = source.indexOf("\n  function ", start + 1);
  const end = [next, nextSync].filter((n) => n > start).sort((a, b) => a - b)[0];
  return source.slice(start, end);
}

test("list export delivers the file before reporting success", () => {
  const downloadExport = sliceFunction("downloadExport");
  assert.match(downloadExport, /if \(exportCaptureMode\)/);
  const captureReturn = downloadExport.indexOf("return;");
  const deliver = downloadExport.indexOf("await downloadBlobAs(blob, fn);");
  assert.ok(captureReturn > 0 && captureReturn < deliver, "capture mode must not download");
  assert.equal(downloadExport.includes("link.click"), false);
  assert.equal(downloadExport.includes("revokeObjectURL"), false);
  assert.equal(downloadExport.includes("reportDone("), false);

  const withoutDef = source.replace("async function downloadExport(", "");
  const calls = withoutDef.match(/downloadExport\(/g) || [];
  const awaited = withoutDef.match(/await downloadExport\(/g) || [];
  assert.equal(calls.length, awaited.length);
  assert.ok(awaited.length >= 15);

  const finish = sliceFunction("finishOwssvrExport");
  const finishDeliver = finish.indexOf("await downloadExport(");
  const finishDone = finish.indexOf("reportDone(true");
  assert.ok(finishDeliver > 0 && finishDeliver < finishDone);

  const rest = sliceFunction("exportViaRest");
  let from = 0;
  let seen = 0;
  while (from < rest.length) {
    const doneAt = rest.indexOf("reportDone(true", from);
    if (doneAt < 0) break;
    const prior = rest.lastIndexOf("await downloadExport(", doneAt);
    assert.ok(prior > from - 1 && prior < doneAt, "REST success must follow an awaited download");
    seen++;
    from = doneAt + 1;
  }
  assert.ok(seen >= 4);
});
