import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { describe, it } from "node:test";
import { fileURLToPath } from "node:url";
import {
  classifySpecialView,
  specialViewSaveBlockMessage,
} from "../lib/viewManagerSpecialViews.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const viewsJs = readFileSync(join(root, "views.js"), "utf8");

const CALENDAR_VIEW_QUERY =
  '<Where><DateRangesOverlap><FieldRef Name="EventDate"/><FieldRef Name="EndDate"/>' +
  '<FieldRef Name="RecurrenceID"/><Value Type="DateTime"><Month/></Value></DateRangesOverlap></Where>';

describe("classifySpecialView", () => {
  it("treats classic calendar ViewQuery as calendar", () => {
    assert.equal(classifySpecialView({ viewQuery: CALENDAR_VIEW_QUERY }), "calendar");
  });

  it("treats CalendarSettings as calendar even without DateRangesOverlap", () => {
    assert.equal(classifySpecialView({ calendarSettings: "<CalendarSettings/>" }), "calendar");
  });

  it("treats modern ViewType2 values as special", () => {
    assert.equal(classifySpecialView({ viewType2: "MODERNCALENDAR" }), "calendar");
    assert.equal(classifySpecialView({ viewType2: "KANBAN" }), "board");
    assert.equal(classifySpecialView({ viewType2: "TILES" }), "gallery");
    assert.equal(classifySpecialView({ viewType2: "GALLERY" }), "gallery");
    assert.equal(classifySpecialView({ viewType2: "GANTT" }), "gantt");
  });

  it("treats Gantt ViewData as gantt", () => {
    assert.equal(
      classifySpecialView({ viewData: '<FieldRef Name="Title" Type="GanttTitle"/>' }),
      "gantt"
    );
  });

  it("allows standard list views", () => {
    assert.equal(classifySpecialView({
      viewType2: "",
      viewQuery: '<Where><Eq><FieldRef Name="FSObjType"/><Value Type="Integer">0</Value></Eq></Where><OrderBy><FieldRef Name="FileLeafRef" Ascending="True"/></OrderBy>',
      calendarSettings: "",
      viewData: "",
    }), "");
    assert.equal(classifySpecialView({ viewType2: "LIST" }), "");
    assert.equal(classifySpecialView({ viewType2: "COMPACTLIST" }), "");
    assert.equal(classifySpecialView({}), "");
  });
});

describe("specialViewSaveBlockMessage", () => {
  it("names the view kind and refuses save", () => {
    const msg = specialViewSaveBlockMessage("calendar");
    assert.match(msg, /cannot save calendar views/i);
    assert.match(msg, /breaks the calendar\/board layout/i);
  });

  it("is empty when the view is safe to save", () => {
    assert.equal(specialViewSaveBlockMessage(""), "");
  });
});

describe("views.js source sync", () => {
  it("classifies DateRangesOverlap and ViewType2 before save", () => {
    assert.match(viewsJs, /DateRangesOverlap/i);
    assert.match(viewsJs, /ViewType2/);
    assert.match(viewsJs, /CalendarSettings/);
    assert.match(viewsJs, /specialViewKind/);
    assert.match(viewsJs, /View Manager cannot save /);
  });

  it("refuses save before PATCH / RemoveAllViewFields when the view is special", () => {
    const saveIdx = viewsJs.indexOf("async function saveView()");
    const patchIdx = viewsJs.indexOf("ViewQuery: viewQuery", saveIdx);
    const specialGuard = viewsJs.indexOf("selectedViewId && specialViewKind", saveIdx);
    assert.ok(saveIdx >= 0 && patchIdx > saveIdx, "saveView PATCHes ViewQuery");
    assert.ok(specialGuard > saveIdx && specialGuard < patchIdx, "special-view guard runs before ViewQuery PATCH");
  });
});
