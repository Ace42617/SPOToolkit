/**
 * View Manager data loading (SharePoint REST).
 * Tested here; injected script is getViewsData.js (same behavior, classic IIFE for the page).
 */

export const ACCEPT_NOMETADATA = "application/json;odata=nometadata";
export const RESULT_TYPE = "SPCSVViewsDataResult";

/** Strip trailing slash for stable URL joins. */
export function normalizeWebUrl(url) {
  return String(url || "").replace(/\/$/, "");
}

/** Safe JSON parse for #sp-views-params script text. */
export function parseInjectParams(text) {
  try {
    if (text == null || String(text).trim() === "") return {};
    return JSON.parse(String(text));
  } catch {
    return {};
  }
}

/**
 * Prefer explicit web URL from extension (off-list View Manager); else page context.
 */
export function resolveSiteUrl(paramWeb, pageContext) {
  const pc = pageContext || {};
  const p =
    paramWeb != null && String(paramWeb).trim() !== ""
      ? normalizeWebUrl(paramWeb)
      : "";
  return normalizeWebUrl(p || pc.webAbsoluteUrl || pc.siteAbsoluteUrl || "");
}

/** GUID without braces for REST path segments. */
export function normalizeGuidParam(raw) {
  if (raw == null || raw === "") return "";
  return String(raw).replace(/[{}]/g, "").trim();
}

export function sqlEscapeListTitle(title) {
  return String(title || "").replace(/'/g, "''");
}

export function unescapeListTitleForDisplay(sqlTitle) {
  return String(sqlTitle || "").replace(/''/g, "'");
}

/** Board / lane scaffold columns — omit from editor list. */
export function isBoardScaffoldField(f) {
  const title = (f.Title || "").trim();
  const iname = (f.InternalName || "").trim();
  const norm = iname.replace(/_x0020_/gi, " ").replace(/_x005f_/gi, "_");
  const lowT = title.toLowerCase();
  if (lowT.includes("board view choice") || norm.toLowerCase().includes("board view choice")) return true;
  const ts = f.TypeAsString || "";
  if (/\schoice\d+$/i.test(title) && /^Choice$/i.test(ts) && (f.ReadOnlyField === true || f.CanBeDeleted === false))
    return true;
  return false;
}

export function mergeStubFieldsForView(fieldsArr, viewFields, filters, orderBy, groupBy, orderByLevels, groupByLevels) {
  const known = Object.create(null);
  for (let i = 0; i < fieldsArr.length; i++) known[fieldsArr[i].internalName] = true;
  function addStub(name) {
    if (!name || known[name]) return;
    known[name] = true;
    fieldsArr.push({ internalName: name, title: name, typeAsString: "Text" });
  }
  let j;
  if (viewFields) for (j = 0; j < viewFields.length; j++) addStub(viewFields[j]);
  if (filters) for (j = 0; j < filters.length; j++) if (filters[j].field) addStub(filters[j].field);
  if (orderByLevels && orderByLevels.length) {
    for (j = 0; j < orderByLevels.length; j++) if (orderByLevels[j] && orderByLevels[j].field) addStub(orderByLevels[j].field);
  } else if (orderBy && orderBy.field) addStub(orderBy.field);
  if (groupByLevels && groupByLevels.length) {
    for (j = 0; j < groupByLevels.length; j++) {
      const g = groupByLevels[j];
      addStub(g && g.field ? g.field : g);
    }
  } else if (groupBy) addStub(groupBy);
}

export function mapODataViewRow(v) {
  return {
    id: String(v.Id || "").replace(/[{}]/g, ""),
    title: v.Title || "",
    defaultView: !!v.DefaultView,
  };
}

const SKIP_INTERNAL = /^_|^vti_|^ows_|^tp_/;

export function mapFieldRowToModel(f) {
  const iname = f.InternalName || f.Title || "";
  if (!iname || SKIP_INTERNAL.test(iname)) return null;
  if (isBoardScaffoldField(f)) return null;
  return { internalName: iname, title: f.Title || iname, typeAsString: f.TypeAsString || "Text" };
}

export function rawFieldsToModels(rows) {
  const fields = [];
  for (let i = 0; i < rows.length; i++) {
    const m = mapFieldRowToModel(rows[i]);
    if (m) fields.push(m);
  }
  return fields;
}

/** Follow @odata.nextLink until exhausted. */
export async function fetchODataAllPages(fetchImpl, startUrl, accept = ACCEPT_NOMETADATA) {
  const accum = [];
  let url = startUrl;
  while (url) {
    const r = await fetchImpl(url, { credentials: "include", headers: { Accept: accept } });
    const j = await r.json();
    const batch = j.value || (j.d && j.d.results) || [];
    for (let i = 0; i < batch.length; i++) accum.push(batch[i]);
    url = j["@odata.nextLink"] || j["odata.nextLink"] || null;
  }
  return accum;
}

export async function fetchJson(fetchImpl, url, accept = ACCEPT_NOMETADATA) {
  const r = await fetchImpl(url, { credentials: "include", headers: { Accept: accept } });
  return r.json();
}

/**
 * Resolve { listId, listTitle } (listTitle SQL-escaped for future URL use in callers).
 */
export async function resolveListContext({ siteUrl, listUrl, forcedListId, fetchImpl, accept = ACCEPT_NOMETADATA }) {
  if (forcedListId && siteUrl) {
    const listData = await fetchJson(
      fetchImpl,
      `${siteUrl}/_api/web/lists(guid'${forcedListId}')?$select=Title`,
      accept
    );
    const listTitle = sqlEscapeListTitle(listData.Title || listData.title || "");
    return { listId: forcedListId, listTitle };
  }
  if (!siteUrl || !listUrl) {
    throw new Error("MISSING_LIST_CONTEXT");
  }
  const enc = encodeURIComponent("'" + String(listUrl).replace(/'/g, "''") + "'");
  const j = await fetchJson(fetchImpl, `${siteUrl}/_api/web/GetList(@u)/Id?@u=${enc}`, accept);
  const listId = String(j.value || "").replace(/[{}]/g, "");
  if (!listId) throw new Error("Could not get list ID");
  const listData = await fetchJson(
    fetchImpl,
    `${siteUrl}/_api/web/lists(guid'${listId}')?$select=Title`,
    accept
  );
  const listTitle = sqlEscapeListTitle(listData.Title || listData.title || "");
  return { listId, listTitle };
}

const FIELDS_SELECT =
  "InternalName,Title,Hidden,TypeAsString,ReadOnlyField,CanBeDeleted&$orderby=Title&$filter=Hidden eq false";

export async function loadViewsAndFieldModels({ siteUrl, ctx, fetchImpl, accept = ACCEPT_NOMETADATA }) {
  const lb = `${siteUrl}/_api/web/lists(guid'${ctx.listId}')`;
  const [viewRows, rawFields] = await Promise.all([
    fetchODataAllPages(fetchImpl, `${lb}/views?$select=Id,Title,DefaultView`, accept),
    fetchODataAllPages(fetchImpl, `${lb}/fields?$select=${FIELDS_SELECT}`, accept),
  ]);
  return {
    ctx,
    views: viewRows.map(mapODataViewRow),
    fields: rawFieldsToModels(rawFields),
  };
}

export function viewFieldsResponseToNames(vfJson) {
  const d = vfJson && vfJson.d;
  const vfItems = (vfJson && (vfJson.Items || vfJson.value || (d && d.results))) || [];
  return []
    .concat(vfItems)
    .filter(Boolean)
    .map((x) => String(typeof x === "object" ? x.Name || x : x));
}

/** Parse CAML FieldRef nodes; Ascending omitted defaults to true (SharePoint). */
export function parseCamlFieldRefs(xml) {
  const refs = [];
  const re = /<FieldRef\b([^>]*?)\/?\s*>/gi;
  let m;
  while ((m = re.exec(String(xml || ""))) !== null) {
    const attrs = m[1] || "";
    const nameM = /\bName\s*=\s*["']([^"']+)["']/i.exec(attrs);
    if (!nameM) continue;
    const ascM = /\bAscending\s*=\s*["'](True|False)["']/i.exec(attrs);
    refs.push({
      field: nameM[1],
      ascending: ascM ? ascM[1].toLowerCase() === "true" : true,
    });
  }
  return refs;
}

function camlAttrTrueFalse(attrs, name) {
  const re = new RegExp("\\b" + name + "\\s*=\\s*[\"'](True|False)[\"']", "i");
  const m = re.exec(attrs || "");
  if (!m) return null;
  return m[1].toLowerCase() === "true";
}

export function parseViewQueryParts(viewQuery) {
  const q = String(viewQuery || "").replace(/\s+/g, " ");
  const filters = [];
  let orderByLevels = [];
  const orderBlock = /<OrderBy\b[^>]*>([\s\S]*?)<\/OrderBy>/i.exec(q);
  if (orderBlock) orderByLevels = parseCamlFieldRefs(orderBlock[1]);
  const orderBy = orderByLevels.length ? orderByLevels[0] : null;

  let groupBy = null;
  let groupByLevels = [];
  let groupExpand = true;
  let groupLimit = null;
  const groupBlock = /<GroupBy\b([^>]*)>([\s\S]*?)<\/GroupBy>/i.exec(q);
  if (groupBlock) {
    const gAttrs = groupBlock[1] || "";
    const collapse = camlAttrTrueFalse(gAttrs, "Collapse");
    // SharePoint default Collapse=TRUE (groups collapsed) when the attribute is omitted.
    groupExpand = collapse === false;
    const lim = /\bGroupLimit\s*=\s*["'](\d+)["']/i.exec(gAttrs);
    if (lim) groupLimit = parseInt(lim[1], 10);
    groupByLevels = parseCamlFieldRefs(groupBlock[2]);
    if (groupByLevels.length) groupBy = groupByLevels[0].field;
  }

  const condRegex =
    /<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)>\s*<FieldRef\s+Name="([^"]+)"\s*\/>\s*<Value\s+Type="([^"]*)">([^<]*)<\/Value>/gi;
  let m;
  while ((m = condRegex.exec(q)) !== null) {
    filters.push({
      op: m[1],
      field: m[2],
      valueType: m[3] || "Text",
      value: (m[4] || "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">"),
    });
  }
  return { viewQuery: q, orderBy, orderByLevels, filters, groupBy, groupByLevels, groupExpand, groupLimit };
}

/**
 * Map loaded viewDetails into View Manager editor sort/group state.
 * Used by views.js (inlined) so a no-op save does not drop secondary sorts or grouping.
 */
export function editorQueryStateFromViewDetails(viewDetails) {
  const d = viewDetails || {};
  let sortLevels;
  if (d.orderByLevels && d.orderByLevels.length) {
    sortLevels = d.orderByLevels.map(function (s) {
      return { field: s.field, ascending: s.ascending !== false };
    });
  } else {
    sortLevels = d.orderBy ? [{ field: d.orderBy.field, ascending: d.orderBy.ascending !== false }] : [];
  }
  const groupByLevels = d.groupByLevels && d.groupByLevels.length ? d.groupByLevels : [];
  const groupByColumn = d.groupBy || (groupByLevels[0] && groupByLevels[0].field) || "";
  return {
    sortLevels: sortLevels,
    groupByColumn: groupByColumn,
    groupExpand: groupByColumn ? d.groupExpand !== false : true,
    groupByAscending: groupByLevels[0] ? groupByLevels[0].ascending !== false : true,
    groupByExtraLevels: groupByLevels.slice(1).map(function (s) {
      return { field: s.field, ascending: s.ascending !== false };
    }),
    groupLimit: d.groupLimit != null && d.groupLimit !== "" ? d.groupLimit : null,
  };
}

export async function loadSingleViewDetails({
  siteUrl,
  ctx,
  viewId,
  fields,
  fetchImpl,
  accept = ACCEPT_NOMETADATA,
}) {
  const lb = `${siteUrl}/_api/web/lists(guid'${ctx.listId}')`;
  const vBase = `${lb}/views(guid'${viewId}')`;
  const [vfJson, metaJson] = await Promise.all([
    fetchJson(fetchImpl, `${vBase}/ViewFields`, accept),
    fetchJson(fetchImpl, `${vBase}?$select=ViewQuery,Title`, accept),
  ]);
  const viewFields = viewFieldsResponseToNames(vfJson);
  const viewQuery = (metaJson.ViewQuery || "").replace(/\s+/g, " ");
  const viewTitle = metaJson.Title || "";
  const parsed = parseViewQueryParts(viewQuery);
  const { orderBy, orderByLevels, filters, groupBy, groupByLevels, groupExpand, groupLimit } = parsed;
  const fieldsOut = fields.slice();
  mergeStubFieldsForView(fieldsOut, viewFields, filters, orderBy, groupBy, orderByLevels, groupByLevels);
  return {
    viewFields,
    viewQuery,
    viewTitle,
    orderBy,
    orderByLevels,
    filters,
    groupBy,
    groupByLevels,
    groupExpand,
    groupLimit,
    fieldsOut,
  };
}

/**
 * Full pipeline: resolve list → views + fields → optional view details.
 * @param {object} deps
 * @param {(data: object) => void} deps.send — postMessage payload body (without __spcsv wrapper)
 * @param {object} deps.params — { viewId?, listId?, webAbsoluteUrl? }
 * @param {object} deps.pageContext — _spPageContextInfo
 * @param {typeof fetch} deps.fetchImpl
 */
export async function runGetViewsDataPipeline({ send, params, pageContext, fetchImpl }) {
  const accept = ACCEPT_NOMETADATA;
  const siteUrl = resolveSiteUrl(params.webAbsoluteUrl, pageContext);
  const listUrl = pageContext.listUrl || pageContext.listServerRelativeUrl;
  const forcedListId = normalizeGuidParam(params.listId);

  let ctx;
  try {
    ctx = await resolveListContext({
      siteUrl,
      listUrl,
      forcedListId: forcedListId || null,
      fetchImpl,
      accept,
    });
  } catch (e) {
    if (e && e.message === "MISSING_LIST_CONTEXT") {
      send({ error: "Open a list or library page.", views: [], fields: [] });
      return;
    }
    send({ error: (e && e.message) || String(e), views: [], fields: [] });
    return;
  }

  let data;
  try {
    data = await loadViewsAndFieldModels({ siteUrl, ctx, fetchImpl, accept });
  } catch (e) {
    send({ error: (e && e.message) || String(e), views: [], fields: [] });
    return;
  }

  const viewId = params.viewId ? normalizeGuidParam(params.viewId) : "";
  if (!viewId) {
    send({
      views: data.views,
      fields: data.fields,
      listId: data.ctx.listId,
      listTitle: unescapeListTitleForDisplay(data.ctx.listTitle),
    });
    return;
  }

  try {
    const details = await loadSingleViewDetails({
      siteUrl,
      ctx: data.ctx,
      viewId,
      fields: data.fields,
      fetchImpl,
      accept,
    });
    send({
      views: data.views,
      fields: details.fieldsOut,
      listId: data.ctx.listId,
      listTitle: unescapeListTitleForDisplay(data.ctx.listTitle),
      viewDetails: {
        viewFields: details.viewFields,
        viewQuery: details.viewQuery,
        viewTitle: details.viewTitle,
        orderBy: details.orderBy,
        orderByLevels: details.orderByLevels,
        filters: details.filters,
        groupBy: details.groupBy,
        groupByLevels: details.groupByLevels,
        groupExpand: details.groupExpand,
        groupLimit: details.groupLimit,
      },
    });
  } catch (err) {
    send({
      views: data.views,
      fields: data.fields,
      listId: data.ctx.listId,
      listTitle: unescapeListTitleForDisplay(data.ctx.listTitle),
      error: (err && err.message) || "Failed to load view details",
    });
  }
}
