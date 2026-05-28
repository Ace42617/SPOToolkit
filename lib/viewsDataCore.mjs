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

export function mergeStubFieldsForView(fieldsArr, viewFields, filters, orderBy, groupBy) {
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
  if (orderBy && orderBy.field) addStub(orderBy.field);
  if (groupBy) addStub(groupBy);
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

export function formatFetchError(response, json) {
  const status = response && response.status ? `HTTP ${response.status}` : "Request failed";
  const err = json && json.error;
  let msg = "";
  if (typeof err === "string") msg = err;
  else if (err && typeof err.message === "string") msg = err.message;
  else if (err && err.message && typeof err.message.value === "string") msg = err.message.value;
  return msg ? `${status}: ${msg}` : status;
}

export async function parseCheckedJson(response) {
  const json = await response.json();
  if (response.ok === false || (json && json.error)) throw new Error(formatFetchError(response, json));
  return json;
}

/** Follow @odata.nextLink until exhausted. */
export async function fetchODataAllPages(fetchImpl, startUrl, accept = ACCEPT_NOMETADATA) {
  const accum = [];
  let url = startUrl;
  while (url) {
    const r = await fetchImpl(url, { credentials: "include", headers: { Accept: accept } });
    const j = await parseCheckedJson(r);
    const batch = j.value || (j.d && j.d.results) || [];
    for (let i = 0; i < batch.length; i++) accum.push(batch[i]);
    url = j["@odata.nextLink"] || j["odata.nextLink"] || null;
  }
  return accum;
}

export async function fetchJson(fetchImpl, url, accept = ACCEPT_NOMETADATA) {
  const r = await fetchImpl(url, { credentials: "include", headers: { Accept: accept } });
  return parseCheckedJson(r);
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

export function parseViewQueryParts(viewQuery) {
  const q = String(viewQuery || "").replace(/\s+/g, " ");
  let orderBy = null;
  let groupBy = null;
  const filters = [];
  const orderMatch = /<OrderBy>\s*<FieldRef\s+Name="([^"]+)"\s+Ascending="(True|False)"/i.exec(q);
  if (orderMatch) orderBy = { field: orderMatch[1], ascending: orderMatch[2].toLowerCase() === "true" };
  const groupMatch = /<GroupBy>\s*<FieldRef\s+Name="([^"]+)"/i.exec(q);
  if (groupMatch) groupBy = groupMatch[1];
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
  return { viewQuery: q, orderBy, filters, groupBy };
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
  const { orderBy, filters, groupBy } = parseViewQueryParts(viewQuery);
  const fieldsOut = fields.slice();
  mergeStubFieldsForView(fieldsOut, viewFields, filters, orderBy, groupBy);
  return {
    viewFields,
    viewQuery,
    viewTitle,
    orderBy,
    filters,
    groupBy,
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
        filters: details.filters,
        groupBy: details.groupBy,
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
