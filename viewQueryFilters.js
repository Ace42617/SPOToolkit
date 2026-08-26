/**
 * View Manager CAML Where parse/build — loaded before views.js; also runnable under Node (module.exports).
 *
 * SharePoint views commonly use IsNull / IsNotNull (is empty / is not empty) and nested
 * <Today/> / <UserID/> tokens. The editor previously only recognized Eq/Neq/… with a
 * text Value, so those filters loaded as empty and Save wrote ViewQuery without <Where>.
 */
(function (root) {
  "use strict";

  var VALUE_OPS = ["Eq", "Neq", "Gt", "Geq", "Lt", "Leq", "Contains", "BeginsWith"];
  var UNARY_OPS = ["IsNull", "IsNotNull"];
  var LEAF_OPS = VALUE_OPS.concat(UNARY_OPS);

  var FILTER_OPS = [
    { value: "Eq", label: "equals" },
    { value: "Neq", label: "not equals" },
    { value: "Gt", label: "greater than" },
    { value: "Geq", label: "greater or equal" },
    { value: "Lt", label: "less than" },
    { value: "Leq", label: "less or equal" },
    { value: "Contains", label: "contains" },
    { value: "BeginsWith", label: "begins with" },
    { value: "IsNull", label: "is empty" },
    { value: "IsNotNull", label: "is not empty" }
  ];

  var OP_CANON = {};
  LEAF_OPS.forEach(function (op) {
    OP_CANON[op.toLowerCase()] = op;
  });

  function canonicalOp(op) {
    if (!op) return "";
    return OP_CANON[String(op).toLowerCase()] || "";
  }

  function isUnaryFilterOp(op) {
    var c = canonicalOp(op);
    return c === "IsNull" || c === "IsNotNull";
  }

  function filterRowParticipates(f) {
    if (!f || !f.field) return false;
    if (isUnaryFilterOp(f.op)) return true;
    return String(f.value == null ? "" : f.value).trim() !== "";
  }

  function escapeXmlAttr(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function escapeXmlText(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function unescapeXmlText(s) {
    return String(s == null ? "" : s)
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">");
  }

  function fieldRefNameFromAttrs(attrBlob) {
    var m = /\bName\s*=\s*"([^"]+)"/i.exec(attrBlob || "");
    return m ? m[1] : "";
  }

  /**
   * Map CAML Value inner XML to the editor token/text.
   * <Today OffsetDays="-7"/> → [Today]-7; <UserID/> → [Me].
   */
  function camlValueInnerToEditor(inner) {
    var s = String(inner == null ? "" : inner).trim();
    var today = /^<Today\b([^>]*)\/\s*>$/i.exec(s);
    if (!today) today = /^<Today\b([^>]*)>\s*<\/Today\s*>$/i.exec(s);
    if (today) {
      var offsetM = /\bOffsetDays\s*=\s*"(-?\d+)"/i.exec(today[1] || "");
      var n = offsetM ? parseInt(offsetM[1], 10) : 0;
      if (!n) return "[Today]";
      return n > 0 ? "[Today]+" + n : "[Today]" + n;
    }
    if (/^<Now\b[^>]*\/\s*>$/i.test(s) || /^<Now\b[^>]*>\s*<\/Now\s*>$/i.test(s)) return "[Now]";
    if (/^<UserID\b[^>]*\/\s*>$/i.test(s) || /^<UserID\b[^>]*>\s*<\/UserID\s*>$/i.test(s)) return "[Me]";
    return unescapeXmlText(s);
  }

  /**
   * Map editor tokens to CAML Value inner XML. Returns null when the value is plain text.
   */
  function editorValueToCamlToken(value) {
    var t = String(value == null ? "" : value).trim();
    var today = /^\[Today\]\s*([+-]\d+)?$/i.exec(t);
    if (today) {
      var n = today[1] ? parseInt(today[1], 10) : 0;
      if (!n) return { kind: "today", xml: "<Today />" };
      return { kind: "today", xml: '<Today OffsetDays="' + n + '" />' };
    }
    if (/^\[Now\]$/i.test(t)) return { kind: "today", xml: "<Now />" };
    if (/^\[Tomorrow\]$/i.test(t)) return { kind: "today", xml: '<Today OffsetDays="1" />' };
    if (/^\[Yesterday\]$/i.test(t)) return { kind: "today", xml: '<Today OffsetDays="-1" />' };
    if (/^\[Me\]$/i.test(t)) return { kind: "me", xml: "<UserID />" };
    return null;
  }

  function matchBalancedOuter(s, tag) {
    var openRe = new RegExp("^<" + tag + "\\s*>", "i");
    var om = s.match(openRe);
    if (!om) return null;
    var i = om[0].length;
    var depth = 1;
    while (i < s.length && depth > 0) {
      var rest = s.slice(i);
      var no = rest.match(new RegExp("^<" + tag + "\\s*>", "i"));
      var nc = rest.match(new RegExp("^</" + tag + "\\s*>", "i"));
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
    var t;
    for (t = 0; t < LEAF_OPS.length; t++) {
      var tag = LEAF_OPS[t];
      var re = new RegExp("^<" + tag + "\\s*>[\\s\\S]*?</" + tag + "\\s*>", "i");
      var m = s.match(re);
      if (m) return [m[0], s.slice(m[0].length).trim()];
    }
    var tries = ["And", "Or"];
    for (t = 0; t < tries.length; t++) {
      var block = matchBalancedOuter(s, tries[t]);
      if (block) return [block, s.slice(block.length).trim()];
    }
    return null;
  }

  function takeAllFilterSegments(body) {
    var segs = [];
    var rest = body.trim();
    while (rest.length) {
      var seg = takeFirstFilterSegment(rest);
      if (!seg) return null;
      segs.push(seg[0].trim());
      rest = seg[1].trim();
    }
    return segs.length ? segs : null;
  }

  function parseLeafCondString(s) {
    var norm = s.replace(/\s+/g, " ").trim();
    var unary = /^<(IsNull|IsNotNull)\s*>\s*<FieldRef\b([^>]*?)(?:\s*\/>|><\/FieldRef\s*>)\s*<\/\1\s*>$/i.exec(norm);
    if (unary) {
      var uField = fieldRefNameFromAttrs(unary[2]);
      if (!uField) return null;
      return { type: "leaf", field: uField, op: canonicalOp(unary[1]), value: "" };
    }
    var valueOp =
      /^<(Eq|Neq|Gt|Geq|Lt|Leq|Contains|BeginsWith)\s*>\s*<FieldRef\b([^>]*?)(?:\s*\/>|><\/FieldRef\s*>)\s*<Value\b([^>]*)>([\s\S]*?)<\/Value>\s*<\/\1\s*>$/i.exec(
        norm
      );
    if (!valueOp) return null;
    var vField = fieldRefNameFromAttrs(valueOp[2]);
    if (!vField) return null;
    return {
      type: "leaf",
      field: vField,
      op: canonicalOp(valueOp[1]),
      value: camlValueInnerToEditor(valueOp[4])
    };
  }

  function parseWhereExpr(inner) {
    inner = inner.trim();
    var leaf = parseLeafCondString(inner);
    if (leaf) return leaf;
    var tries = ["And", "Or"];
    for (var ti = 0; ti < tries.length; ti++) {
      var tag = tries[ti];
      var full = matchBalancedOuter(inner, tag);
      if (!full || full.length !== inner.length) continue;
      var openM = inner.match(new RegExp("^<" + tag + "\\s*>", "i"));
      var body = inner.slice(openM[0].length, inner.length - ("</" + tag + ">").length).trim();
      var segs = takeAllFilterSegments(body);
      if (!segs || segs.length === 0) return null;
      var nodes = segs.map(function (seg) {
        return parseWhereExpr(seg);
      });
      if (nodes.some(function (n) {
        return !n;
      })) return null;
      var acc = nodes[0];
      for (var k = 1; k < nodes.length; k++) {
        acc = { type: "bin", op: tag, left: acc, right: nodes[k] };
      }
      return acc;
    }
    return null;
  }

  function binTreeToFilterRows(node) {
    if (node.type === "leaf") {
      return [{ field: node.field, op: node.op, value: node.value }];
    }
    var L = binTreeToFilterRows(node.left);
    var R = binTreeToFilterRows(node.right);
    if (R.length === 0) return L;
    R[0].join = node.op;
    return L.concat(R);
  }

  function parseViewQueryToFilterRows(viewQuery) {
    if (!viewQuery || typeof viewQuery !== "string") return null;
    var wm = viewQuery.replace(/\s+/g, " ").match(/<Where>\s*([\s\S]*?)\s*<\/Where>/i);
    if (!wm) return null;
    var tree = parseWhereExpr(wm[1].trim());
    if (!tree) return null;
    return binTreeToFilterRows(tree);
  }

  function buildOneFilterCondXml(f, options) {
    options = options || {};
    var op = canonicalOp(f && f.op);
    if (!op || !f || !f.field) return null;
    var fieldXml = escapeXmlAttr(f.field);
    if (isUnaryFilterOp(op)) {
      return "<" + op + "><FieldRef Name=\"" + fieldXml + "\"/></" + op + ">";
    }
    var token = editorValueToCamlToken(f.value);
    if (token && token.kind === "today") {
      return (
        "<" + op + "><FieldRef Name=\"" + fieldXml + "\"/><Value Type=\"DateTime\">" + token.xml + "</Value></" + op + ">"
      );
    }
    if (token && token.kind === "me") {
      return (
        "<" +
        op +
        "><FieldRef Name=\"" +
        fieldXml +
        "\" LookupId=\"TRUE\"/><Value Type=\"Integer\">" +
        token.xml +
        "</Value></" +
        op +
        ">"
      );
    }
    var type;
    var inner;
    if (options.isBooleanField && options.isBooleanField(f.field)) {
      var camlInt = options.booleanDisplayToCamlInteger
        ? options.booleanDisplayToCamlInteger(f.value)
        : null;
      if (camlInt == null) return null;
      inner = camlInt;
      type = "Integer";
    } else {
      type = String(options.getValueType ? options.getValueType(f.field) : "Text").replace(/"/g, "") || "Text";
      inner = escapeXmlText(String(f.value));
    }
    return (
      "<" + op + "><FieldRef Name=\"" + fieldXml + "\"/><Value Type=\"" + type + "\">" + inner + "</Value></" + op + ">"
    );
  }

  function buildFilterWhereXml(filterRows, options) {
    var chunks = [];
    var rows = filterRows || [];
    for (var i = 0; i < rows.length; i++) {
      var f = rows[i];
      if (!filterRowParticipates(f)) continue;
      var xml = buildOneFilterCondXml(f, options);
      if (!xml) continue;
      chunks.push({ xml: xml, join: chunks.length === 0 ? null : f.join || "And" });
    }
    if (chunks.length === 0) return "";
    var acc = chunks[0].xml;
    for (var k = 1; k < chunks.length; k++) {
      var joinOp = chunks[k].join || "And";
      acc = "<" + joinOp + ">" + acc + chunks[k].xml + "</" + joinOp + ">";
    }
    return "<Where>" + acc + "</Where>";
  }

  var api = {
    VALUE_OPS: VALUE_OPS,
    UNARY_OPS: UNARY_OPS,
    FILTER_OPS: FILTER_OPS,
    canonicalOp: canonicalOp,
    isUnaryFilterOp: isUnaryFilterOp,
    filterRowParticipates: filterRowParticipates,
    camlValueInnerToEditor: camlValueInnerToEditor,
    editorValueToCamlToken: editorValueToCamlToken,
    parseLeafCondString: parseLeafCondString,
    parseViewQueryToFilterRows: parseViewQueryToFilterRows,
    buildOneFilterCondXml: buildOneFilterCondXml,
    buildFilterWhereXml: buildFilterWhereXml
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  if (typeof window !== "undefined") {
    window.SPViewQueryFilters = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : this);
