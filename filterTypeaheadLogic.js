/**
 * SharePoint list view filter value tokens — scoped by column TypeAsString (REST).
 * Loaded before views.js; also runnable under Node for tests (module.exports).
 */
(function (root) {
  "use strict";

  /** @type {{ kind: string, token: string, hint: string, kw?: string, ord: number }[]} */
  var LIBRARY = [
    { kind: "date", token: "[Today]", hint: "Current date (re-evaluated when the view runs).", kw: "date due now", ord: 1 },
    { kind: "date", token: "[Today]+7", hint: "+7 days (like CAML Today OffsetDays=\"7\").", kw: "week future deadline", ord: 2 },
    { kind: "date", token: "[Today]-7", hint: "−7 days (OffsetDays=\"-7\").", kw: "week ago past", ord: 3 },
    { kind: "date", token: "[Today]+1", hint: "+1 day (often same as “tomorrow” UI).", kw: "tomorrow", ord: 4 },
    { kind: "date", token: "[Today]-1", hint: "−1 day (often same as “yesterday” UI).", kw: "yesterday", ord: 5 },
    { kind: "date", token: "[Today]+14", hint: "+14 days.", kw: "fortnight sprint", ord: 6 },
    { kind: "date", token: "[Today]-14", hint: "−14 days.", kw: "fortnight", ord: 7 },
    { kind: "date", token: "[Today]+30", hint: "+30 days (~1 month).", kw: "month", ord: 8 },
    { kind: "date", token: "[Today]-30", hint: "−30 days.", kw: "month overdue", ord: 9 },
    { kind: "date", token: "[Today]+90", hint: "+90 days (~quarter).", kw: "quarter fiscal", ord: 10 },
    { kind: "date", token: "[Today]-90", hint: "−90 days.", kw: "quarter", ord: 11 },
    { kind: "date", token: "[Today]+180", hint: "+180 days (~6 months).", kw: "half year", ord: 12 },
    { kind: "date", token: "[Today]-180", hint: "−180 days.", kw: "retention archive", ord: 13 },
    { kind: "date", token: "[Today]+365", hint: "+365 days (~1 year).", kw: "year anniversary", ord: 14 },
    { kind: "date", token: "[Today]-365", hint: "−365 days.", kw: "year audit", ord: 15 },
    { kind: "date", token: "[Today]+2", hint: "+2 days.", kw: "", ord: 20 },
    { kind: "date", token: "[Today]-2", hint: "−2 days.", kw: "", ord: 21 },
    { kind: "date", token: "[Today]+3", hint: "+3 days.", kw: "", ord: 22 },
    { kind: "date", token: "[Today]-3", hint: "−3 days.", kw: "", ord: 23 },
    { kind: "date", token: "[Tomorrow]", hint: "Classic UI alias (often equals [Today]+1).", kw: "tomorrow", ord: 30 },
    { kind: "date", token: "[Yesterday]", hint: "Classic UI alias (often equals [Today]-1).", kw: "yesterday", ord: 31 },
    { kind: "date", token: "[Now]", hint: "Date/time token — behavior varies; verify on your DateTime column.", kw: "time datetime", ord: 32 },
    { kind: "date", token: "[Today]+10", hint: "+10 days (e.g. net-10 style windows).", kw: "invoice net10", ord: 40 },
    { kind: "date", token: "[Today]+15", hint: "+15 days.", kw: "net15", ord: 41 },
    { kind: "date", token: "[Today]+45", hint: "+45 days.", kw: "invoice", ord: 42 },
    { kind: "date", token: "[Today]-60", hint: "−60 days.", kw: "rolling", ord: 43 },
    { kind: "date", token: "[Today]-3650", hint: "−3650 days (~10 years).", kw: "archive", ord: 50 },
    { kind: "person", token: "[Me]", hint: "Current user — Person / People columns only.", kw: "me user assignee author owner current", ord: 1 }
  ];

  var MAX_ITEMS = 16;

  /**
   * Which suggestion kinds apply to this SharePoint field type (REST TypeAsString).
   * @returns {string[]} subset of "date" | "person"
   */
  function kindsFromTypeAsString(typeAsString) {
    var raw = String(typeAsString == null ? "" : typeAsString).trim();
    var t = raw.toLowerCase();
    var out = [];
    if (t.indexOf("datetime") >= 0 || t === "date") {
      out.push("date");
    }
    if (t === "user" || t === "usermulti") {
      out.push("person");
    }
    return out;
  }

  /**
   * @param {string[]} kinds from kindsFromTypeAsString
   * @param {string} queryRaw token / word being typed (may include "[")
   */
  function suggest(kinds, queryRaw) {
    var kindSet = {};
    for (var i = 0; i < kinds.length; i++) kindSet[kinds[i]] = true;
    if (!kinds.length) return [];

    var q = String(queryRaw || "").trim().toLowerCase();
    var list = [];
    for (var j = 0; j < LIBRARY.length; j++) {
      var s = LIBRARY[j];
      if (!kindSet[s.kind]) continue;
      list.push(s);
    }

    if (q) {
      list = list.filter(function (s) {
        if (s.token.toLowerCase().indexOf(q) >= 0) return true;
        if ((s.hint || "").toLowerCase().indexOf(q) >= 0) return true;
        if ((s.kw || "").toLowerCase().indexOf(q) >= 0) return true;
        return false;
      });
      list.sort(function (a, b) {
        var ta = a.token.toLowerCase().indexOf(q);
        var tb = b.token.toLowerCase().indexOf(q);
        var ia = ta < 0 ? 9999 : ta;
        var ib = tb < 0 ? 9999 : tb;
        if (ia !== ib) return ia - ib;
        return (a.ord || 0) - (b.ord || 0);
      });
    } else {
      list.sort(function (a, b) { return (a.ord || 0) - (b.ord || 0); });
    }

    return list.slice(0, MAX_ITEMS);
  }

  var api = { LIBRARY: LIBRARY, kindsFromTypeAsString: kindsFromTypeAsString, suggest: suggest, MAX_ITEMS: MAX_ITEMS };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  if (typeof window !== "undefined") {
    window.SPFilterTypeahead = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : this);
