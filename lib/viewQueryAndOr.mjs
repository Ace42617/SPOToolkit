/**
 * View Manager flattens nested CAML And/Or into a linear row list, then rebuilds
 * left-associatively. Mixed grouping such as A OR (B AND C) round-trips as
 * (A OR B) AND C, so Save would change which items the view shows.
 */

export function extractViewQueryWhereXml(viewQuery) {
  const q = String(viewQuery || "");
  const m = q.match(/<Where\b[^>]*>[\s\S]*?<\/Where>/i);
  return m ? m[0] : "";
}

function matchBalancedOuter(s, tag) {
  const openRe = new RegExp("^<" + tag + "\\s*>", "i");
  const om = s.match(openRe);
  if (!om) return null;
  let i = om[0].length;
  let depth = 1;
  while (i < s.length && depth > 0) {
    const rest = s.slice(i);
    const no = rest.match(new RegExp("^<" + tag + "\\s*>", "i"));
    const nc = rest.match(new RegExp("^</" + tag + "\\s*>", "i"));
    if (nc && (!no || nc.index < no.index)) {
      depth--;
      i += nc[0].length;
      if (depth === 0) return s.slice(0, i);
      continue;
    }
    if (no) {
      depth++;
      i += no[0].length;
      continue;
    }
    i++;
  }
  return null;
}

function takeFirstFilterSegment(s) {
  s = s.trim();
  const leafTags = ["Eq", "Neq", "Gt", "Geq", "Lt", "Leq", "Contains", "BeginsWith"];
  for (let t = 0; t < leafTags.length; t++) {
    const tag = leafTags[t];
    const re = new RegExp("^<" + tag + "\\s*>[\\s\\S]*?</" + tag + "\\s*>", "i");
    const m = s.match(re);
    if (m) return [m[0], s.slice(m[0].length).trim()];
  }
  const tries = ["And", "Or"];
  for (let t = 0; t < tries.length; t++) {
    const block = matchBalancedOuter(s, tries[t]);
    if (block) return [block, s.slice(block.length).trim()];
  }
  return null;
}

function takeAllFilterSegments(body) {
  const segs = [];
  let rest = body.trim();
  while (rest.length) {
    const seg = takeFirstFilterSegment(rest);
    if (!seg) return null;
    segs.push(seg[0].trim());
    rest = seg[1].trim();
  }
  return segs.length ? segs : null;
}

function parseLeafCondString(s) {
  const norm = s.replace(/\s+/g, " ").trim();
  const re = /^<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)\s*>\s*<FieldRef\s+Name="([^"]+)"\s*\/>\s*<Value\s+Type="([^"]*)">([^<]*)<\/Value>\s*<\/\1\s*>$/i;
  const m = norm.match(re);
  if (!m) return null;
  const v = (m[4] || "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">");
  return { type: "leaf", field: m[2], op: m[1], value: v };
}

export function parseWhereExpr(inner) {
  inner = String(inner || "").trim();
  const leaf = parseLeafCondString(inner);
  if (leaf) return leaf;
  const tries = ["And", "Or"];
  for (let ti = 0; ti < tries.length; ti++) {
    const tag = tries[ti];
    const full = matchBalancedOuter(inner, tag);
    if (!full || full.length !== inner.length) continue;
    const openM = inner.match(new RegExp("^<" + tag + "\\s*>", "i"));
    const body = inner.slice(openM[0].length, inner.length - ("</" + tag + ">").length).trim();
    const segs = takeAllFilterSegments(body);
    if (!segs || segs.length === 0) return null;
    const nodes = segs.map(function (seg) { return parseWhereExpr(seg); });
    if (nodes.some(function (n) { return !n; })) return null;
    let acc = nodes[0];
    for (let k = 1; k < nodes.length; k++) {
      acc = { type: "bin", op: tag, left: acc, right: nodes[k] };
    }
    return acc;
  }
  return null;
}

export function binTreeToFilterRows(node) {
  if (!node) return [];
  if (node.type === "leaf") {
    return [{ field: node.field, op: node.op, value: node.value }];
  }
  const L = binTreeToFilterRows(node.left);
  const R = binTreeToFilterRows(node.right);
  if (R.length === 0) return L;
  R[0].join = node.op;
  return L.concat(R);
}

export function filterRowsToLeftAssocTree(rows) {
  if (!rows || !rows.length) return null;
  let acc = { type: "leaf", field: rows[0].field, op: rows[0].op, value: rows[0].value };
  for (let k = 1; k < rows.length; k++) {
    const r = rows[k];
    acc = {
      type: "bin",
      op: r.join || "And",
      left: acc,
      right: { type: "leaf", field: r.field, op: r.op, value: r.value }
    };
  }
  return acc;
}

export function whereTreesEqual(a, b) {
  if (!a || !b) return a === b;
  if (a.type !== b.type) return false;
  if (a.type === "leaf") {
    return a.field === b.field
      && String(a.op).toLowerCase() === String(b.op).toLowerCase()
      && a.value === b.value;
  }
  return String(a.op).toLowerCase() === String(b.op).toLowerCase()
    && whereTreesEqual(a.left, b.left)
    && whereTreesEqual(a.right, b.right);
}

export function groupedWhereIsUniformAssoc(node) {
  if (!node || node.type === "leaf") return true;
  const op = String(node.op).toLowerCase();
  function walk(n) {
    if (!n || n.type === "leaf") return true;
    if (String(n.op).toLowerCase() !== op) return false;
    return walk(n.left) && walk(n.right);
  }
  return walk(node);
}

function extractWhereInner(viewQuery) {
  const xml = extractViewQueryWhereXml(viewQuery);
  if (!xml) return "";
  const m = xml.match(/^<Where\b[^>]*>([\s\S]*)<\/Where>$/i);
  return m ? m[1].trim() : "";
}

/**
 * Left-associative rebuild matching views.js buildFilterWhereXml wrapping.
 * Leaf XML is a simplified Type=Text form used only to compare grouping.
 */
export function rebuildLeftAssocWhereXml(rows) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return "";
  function leafXml(r) {
    const field = String(r.field || "").replace(/"/g, "");
    const op = r.op || "Eq";
    const inner = String(r.value == null ? "" : r.value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    return "<" + op + "><FieldRef Name=\"" + field + "\"/><Value Type=\"Text\">" + inner + "</Value></" + op + ">";
  }
  let acc = leafXml(list[0]);
  for (let k = 1; k < list.length; k++) {
    const op = list[k].join || "And";
    acc = "<" + op + ">" + acc + leafXml(list[k]) + "</" + op + ">";
  }
  return "<Where>" + acc + "</Where>";
}

export function flattenAndRebuildWhereXml(viewQuery) {
  const inner = extractWhereInner(viewQuery);
  if (!inner) return "";
  const tree = parseWhereExpr(inner);
  if (!tree) return "";
  return rebuildLeftAssocWhereXml(binTreeToFilterRows(tree));
}

export function shouldPreserveGroupedWhere(viewQuery) {
  const inner = extractWhereInner(viewQuery);
  if (!inner) return false;
  const tree = parseWhereExpr(inner);
  if (!tree) return false;
  if (groupedWhereIsUniformAssoc(tree)) return false;
  const rebuilt = filterRowsToLeftAssocTree(binTreeToFilterRows(tree));
  return !whereTreesEqual(tree, rebuilt);
}

export function groupedWhereLoadMessage() {
  return (
    "This view uses grouped And/Or filters that cannot be edited here without changing which items appear. " +
    "Save will keep the existing filter as-is."
  );
}

export function groupedWhereSaveBlockMessage() {
  return (
    "Cannot save: this view's filter uses grouped And/Or that View Manager cannot edit without changing which items appear. " +
    "Remove the new filter rows to keep the original filter, or change the filter in SharePoint."
  );
}

/**
 * @param {{ editorFilters: Array, preservedWhereXml: string, builtWhereXml: string }} input
 * @returns {{ ok: true, whereXml: string } | { ok: false, error: string }}
 */
export function resolveGroupedWhereSave(input) {
  const src = input || {};
  const preserved = String(src.preservedWhereXml || "");
  const editorFilters = Array.isArray(src.editorFilters) ? src.editorFilters : [];
  if (preserved) {
    if (editorFilters.length > 0) {
      return { ok: false, error: groupedWhereSaveBlockMessage() };
    }
    return { ok: true, whereXml: preserved };
  }
  return { ok: true, whereXml: String(src.builtWhereXml || "") };
}
