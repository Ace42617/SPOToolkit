/**
 * SharePoint column types and SchemaXml builders for REST field creation.
 */

/** @typedef {{ id: string, label: string, type: string, fieldTypeKind: number, supportsChoices?: boolean, supportsNumberFormat?: boolean, supportsDateFormat?: boolean, supportsUserMode?: boolean }} ColumnTypeDef */

/** @type {ColumnTypeDef[]} */
export const COLUMN_TYPES = [
  { id: "Text", label: "Single line of text", type: "Text", fieldTypeKind: 2 },
  { id: "Note", label: "Multiple lines of text", type: "Note", fieldTypeKind: 3 },
  { id: "Number", label: "Number", type: "Number", fieldTypeKind: 9, supportsNumberFormat: true },
  { id: "Currency", label: "Currency", type: "Currency", fieldTypeKind: 10, supportsNumberFormat: true },
  { id: "DateTime", label: "Date and Time", type: "DateTime", fieldTypeKind: 4, supportsDateFormat: true },
  { id: "Choice", label: "Choice", type: "Choice", fieldTypeKind: 6, supportsChoices: true },
  { id: "MultiChoice", label: "Multiple choice", type: "MultiChoice", fieldTypeKind: 15, supportsChoices: true },
  { id: "Boolean", label: "Yes/No", type: "Boolean", fieldTypeKind: 8 },
  { id: "URL", label: "Hyperlink or Picture", type: "URL", fieldTypeKind: 11 },
  { id: "User", label: "Person or Group", type: "User", fieldTypeKind: 20, supportsUserMode: true },
];

export const DEFAULT_COLUMN_GROUP = "Custom Columns";

export function getColumnTypeDef(typeId) {
  return COLUMN_TYPES.find((t) => t.id === typeId) || COLUMN_TYPES[0];
}

function escapeXmlAttr(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeXmlText(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function choiceXml(choices) {
  const items = (choices || [])
    .map((c) => String(c || "").trim())
    .filter(Boolean);
  if (!items.length) return "<CHOICES><CHOICE>Option 1</CHOICE></CHOICES>";
  return "<CHOICES>" + items.map((c) => "<CHOICE>" + escapeXmlText(c) + "</CHOICE>").join("") + "</CHOICES>";
}

/**
 * @param {object} input
 * @param {string} input.internalName
 * @param {string} input.title
 * @param {string} input.typeId
 * @param {boolean} [input.required]
 * @param {string} [input.group]
 * @param {string} [input.description]
 * @param {string[]} [input.choices]
 * @param {number} [input.decimals]
 * @param {'DateOnly'|'DateTime'} [input.dateFormat]
 * @param {'PeopleOnly'|'PeopleAndGroups'} [input.userSelectionMode]
 * @param {number} [input.numLines]
 * @param {string|number|boolean|null} [input.defaultValue]
 */
export function buildFieldSchemaXml(input) {
  const def = getColumnTypeDef(input?.typeId);
  const name = String(input?.internalName || "").trim();
  const title = String(input?.title || "").trim() || name;
  const group = String(input?.group || DEFAULT_COLUMN_GROUP).trim() || DEFAULT_COLUMN_GROUP;
  const required = !!input?.required;
  const description = String(input?.description || "").trim();

  const attrs = [
    'Type="' + escapeXmlAttr(def.type) + '"',
    'Name="' + escapeXmlAttr(name) + '"',
    'DisplayName="' + escapeXmlAttr(title) + '"',
    'Group="' + escapeXmlAttr(group) + '"',
    'Required="' + (required ? "TRUE" : "FALSE") + '"',
  ];

  if (description) attrs.push('Description="' + escapeXmlAttr(description) + '"');

  if (def.type === "Note") {
    const lines = parseInt(input?.numLines, 10);
    attrs.push('NumLines="' + (lines > 0 ? lines : 6) + '"');
  }
  if (def.supportsNumberFormat) {
    const decimals = parseInt(input?.decimals, 10);
    if (decimals >= 0 && decimals <= 5) attrs.push('Decimals="' + decimals + '"');
  }
  if (def.supportsDateFormat) {
    const fmt = input?.dateFormat === "DateOnly" ? "DateOnly" : "DateTime";
    attrs.push('Format="' + fmt + '"');
  }
  if (def.supportsUserMode) {
    const mode = input?.userSelectionMode === "PeopleOnly" ? "PeopleOnly" : "PeopleAndGroups";
    attrs.push('UserSelectionMode="' + mode + '"');
  }

  let inner = "";
  if (def.supportsChoices) inner = choiceXml(input?.choices);
  if (input?.defaultValue != null && input.defaultValue !== "") {
  }

  return "<Field " + attrs.join(" ") + ">" + inner + "</Field>";
}

/**
 * @param {object} input
 * @param {'site'|'list'|'listFromSite'} input.target
 * @param {string} input.siteUrl
 * @param {string} [input.listId]
 * @param {string} input.schemaXml
 */
export function buildCreateFieldRequest(input) {
  const siteUrl = String(input?.siteUrl || "").replace(/\/$/, "");
  const listId = String(input?.listId || "").replace(/[{}]/g, "");
  const target = input?.target === "site" || input?.target === "listFromSite" ? input.target : "list";
  const schemaXml = String(input?.schemaXml || "");

  if (!siteUrl) {
    return { ok: false, error: "Missing site URL." };
  }

  if (target === "listFromSite") {
    const siteFieldInternal = String(input?.siteFieldInternal || "").trim();
    if (!siteFieldInternal) {
      return { ok: false, error: "Missing site column internal name." };
    }
  } else if (!schemaXml) {
    return { ok: false, error: "Missing field schema." };
  }

  if (target === "list" || target === "listFromSite") {
    if (!listId) return { ok: false, error: "Open a list or library to create a list column." };
    return {
      ok: true,
      endpoint: siteUrl + "/_api/web/lists(guid'" + listId + "')/fields/createfieldasxml",
      method: "POST",
      body: { parameters: { SchemaXml: schemaXml } },
    };
  }

  return {
    ok: true,
    endpoint: siteUrl + "/_api/web/fields/createfieldasxml",
    method: "POST",
    body: { parameters: { SchemaXml: schemaXml } },
  };
}
