/**
 * Tests for View Manager REST helpers (lib/viewsDataCore.mjs).
 */
import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  normalizeWebUrl,
  parseInjectParams,
  resolveSiteUrl,
  normalizeGuidParam,
  isBoardScaffoldField,
  mergeStubFieldsForView,
  mapODataViewRow,
  mapFieldRowToModel,
  rawFieldsToModels,
  fetchODataAllPages,
  fetchJson,
  resolveListContext,
  viewFieldsResponseToNames,
  parseViewQueryParts,
  runGetViewsDataPipeline,
  ACCEPT_NOMETADATA,
} from "../lib/viewsDataCore.mjs";

describe("normalizeWebUrl", () => {
  it("strips trailing slash", () => {
    assert.equal(normalizeWebUrl("https://x/sites/a/"), "https://x/sites/a");
  });
  it("handles empty", () => {
    assert.equal(normalizeWebUrl(""), "");
  });
});

describe("parseInjectParams", () => {
  it("parses valid JSON", () => {
    assert.deepEqual(parseInjectParams('{"listId":"abc"}'), { listId: "abc" });
  });
  it("returns {} on bad JSON", () => {
    assert.deepEqual(parseInjectParams("{"), {});
  });
});

describe("resolveSiteUrl", () => {
  it("prefers param web over page context", () => {
    assert.equal(
      resolveSiteUrl("https://tenant/sites/hr", { webAbsoluteUrl: "https://tenant/sites/root" }),
      "https://tenant/sites/hr"
    );
  });
  it("falls back to page context", () => {
    assert.equal(resolveSiteUrl("", { webAbsoluteUrl: "https://t/w" }), "https://t/w");
  });
});

describe("normalizeGuidParam", () => {
  it("removes braces", () => {
    assert.equal(normalizeGuidParam("{abc-def}"), "abc-def");
  });
});

describe("isBoardScaffoldField", () => {
  it("detects board view choice title", () => {
    assert.equal(isBoardScaffoldField({ Title: "Board View Choice 1", InternalName: "x", TypeAsString: "Text" }), true);
  });
  it("keeps normal field", () => {
    assert.equal(isBoardScaffoldField({ Title: "Job Title", InternalName: "JobTitle", TypeAsString: "Text" }), false);
  });
});

describe("mapFieldRowToModel", () => {
  it("skips internal prefixes", () => {
    assert.equal(mapFieldRowToModel({ InternalName: "_Hidden", Title: "H" }), null);
  });
  it("maps visible field", () => {
    assert.deepEqual(mapFieldRowToModel({ InternalName: "Title", Title: "Title", TypeAsString: "Text" }), {
      internalName: "Title",
      title: "Title",
      typeAsString: "Text",
    });
  });
});

describe("fetchODataAllPages", () => {
  it("follows nextLink", async () => {
    const urls = [];
    const fetchImpl = async (url) => {
      urls.push(url);
      if (url.includes("page=1")) {
        return {
          json: async () => ({
            value: [{ Id: "a", Title: "A" }],
            "@odata.nextLink": "https://x/next",
          }),
        };
      }
      return {
        json: async () => ({
          value: [{ Id: "b", Title: "B" }],
        }),
      };
    };
    const rows = await fetchODataAllPages(fetchImpl, "https://x/page=1", ACCEPT_NOMETADATA);
    assert.equal(rows.length, 2);
    assert.equal(rows[0].Title, "A");
    assert.equal(rows[1].Title, "B");
    assert.equal(urls.length, 2);
  });

  it("throws instead of truncating when a later page fails", async () => {
    const fetchImpl = async (url) => {
      if (url.includes("page=1")) {
        return {
          ok: true,
          json: async () => ({
            value: [{ Id: "a", Title: "A" }],
            "@odata.nextLink": "https://x/page=2",
          }),
        };
      }
      return {
        ok: false,
        status: 429,
        json: async () => ({ error: { message: "throttled" } }),
      };
    };
    await assert.rejects(
      () => fetchODataAllPages(fetchImpl, "https://x/page=1", ACCEPT_NOMETADATA),
      /HTTP 429: throttled/
    );
  });
});

describe("fetchJson", () => {
  it("throws on OData error bodies even when HTTP succeeds", async () => {
    const fetchImpl = async () => ({
      ok: true,
      json: async () => ({ error: { message: { value: "Access denied" } } }),
    });
    await assert.rejects(() => fetchJson(fetchImpl, "https://x", ACCEPT_NOMETADATA), /Access denied/);
  });
});

describe("resolveListContext", () => {
  it("uses forced list id path", async () => {
    const fetchImpl = async (url) => {
      if (url.includes("/lists(guid'lid1')?$select=Title")) {
        return { json: async () => ({ Title: "Docs" }) };
      }
      throw new Error("unexpected " + url);
    };
    const ctx = await resolveListContext({
      siteUrl: "https://t/s",
      listUrl: null,
      forcedListId: "lid1",
      fetchImpl,
    });
    assert.equal(ctx.listId, "lid1");
    assert.ok(ctx.listTitle.includes("Docs"));
  });
});

describe("viewFieldsResponseToNames", () => {
  it("reads Items array", () => {
    assert.deepEqual(viewFieldsResponseToNames({ Items: ["LinkTitle", "Modified"] }), ["LinkTitle", "Modified"]);
  });
  it("reads value array", () => {
    assert.deepEqual(viewFieldsResponseToNames({ value: ["a"] }), ["a"]);
  });
});

describe("parseViewQueryParts", () => {
  it("parses OrderBy and Where", () => {
    const q =
      '<Where><Eq><FieldRef Name="Foo"/><Value Type="Text">bar</Value></Eq></Where><OrderBy><FieldRef Name="Modified" Ascending="False"/></OrderBy>';
    const p = parseViewQueryParts(q);
    assert.equal(p.orderBy.field, "Modified");
    assert.equal(p.orderBy.ascending, false);
    assert.equal(p.filters.length, 1);
    assert.equal(p.filters[0].field, "Foo");
    assert.equal(p.filters[0].value, "bar");
  });
});

describe("runGetViewsDataPipeline", () => {
  it("sends error when no site/list context", async () => {
    const sent = [];
    await runGetViewsDataPipeline({
      send: (x) => sent.push(x),
      params: {},
      pageContext: {},
      fetchImpl: async () => ({ json: async () => ({}) }),
    });
    assert.equal(sent.length, 1);
    assert.match(sent[0].error, /Open a list/);
  });

  it("sends views and fields without viewId", async () => {
    const sent = [];
    const fetchImpl = async (url) => {
      if (url.includes("GetList(@u)")) return { json: async () => ({ value: "{abc}" }) };
      if (url.includes("/lists(guid'abc')?$select=Title")) return { json: async () => ({ Title: "L1" }) };
      if (url.includes("/views?$select")) return { json: async () => ({ value: [{ Id: "{v1}", Title: "All", DefaultView: true }] }) };
      if (url.includes("/fields?$select")) return { json: async () => ({ value: [{ InternalName: "Title", Title: "T", TypeAsString: "Text" }] }) };
      return { json: async () => ({ value: [] }) };
    };
    await runGetViewsDataPipeline({
      send: (x) => sent.push(x),
      params: {},
      pageContext: { webAbsoluteUrl: "https://t/sites/x", listUrl: "/sites/x/Lists/MyList" },
      fetchImpl,
    });
    assert.equal(sent.length, 1);
    assert.equal(sent[0].error, undefined);
    assert.equal(sent[0].listId, "abc");
    assert.equal(sent[0].views.length, 1);
    assert.equal(sent[0].views[0].id, "v1");
    assert.equal(sent[0].fields.length, 1);
  });

  it("uses webAbsoluteUrl from params for forced list", async () => {
    const sent = [];
    const fetchImpl = async (url) => {
      assert.ok(url.startsWith("https://correct-site"), url);
      if (url.includes("/lists(guid'zzz')?$select=Title")) return { json: async () => ({ Title: "Lib" }) };
      if (url.includes("/views?$select")) return { json: async () => ({ value: [] }) };
      if (url.includes("/fields?$select")) return { json: async () => ({ value: [] }) };
      return { json: async () => ({}) };
    };
    await runGetViewsDataPipeline({
      send: (x) => sent.push(x),
      params: { listId: "zzz", webAbsoluteUrl: "https://correct-site/sites/hr" },
      pageContext: { webAbsoluteUrl: "https://wrong/sites/root" },
      fetchImpl,
    });
    assert.equal(sent[0].error, undefined);
    assert.equal(sent[0].listId, "zzz");
  });
});

describe("mergeStubFieldsForView", () => {
  it("adds stubs for view-only fields", () => {
    const fields = [{ internalName: "Title", title: "T", typeAsString: "Text" }];
    mergeStubFieldsForView(fields, ["MissingCol"], [], null, null);
    assert.equal(fields.length, 2);
    assert.equal(fields[1].internalName, "MissingCol");
  });
});
