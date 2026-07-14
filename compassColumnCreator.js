/* global window */
(function () {
  const MAX_INTERNAL_NAME_LENGTH = 32;

  function isXmlNameStartChar(code) {
    return (
      (code >= 0x41 && code <= 0x5a) ||
      (code >= 0x61 && code <= 0x7a) ||
      code === 0x5f
    );
  }

  function isXmlNameChar(code) {
    return (
      isXmlNameStartChar(code) ||
      (code >= 0x30 && code <= 0x39) ||
      code === 0x2d ||
      code === 0x2e ||
      code === 0x3a
    );
  }

  function encodeCodePoint(code) {
    return "_x" + code.toString(16).padStart(4, "0") + "_";
  }

  function encodeInternalNameFromTitle(title) {
    const raw = String(title || "").trim();
    if (!raw) return "";
    let out = "";
    for (let i = 0; i < raw.length; i++) {
      const ch = raw[i];
      const code = raw.charCodeAt(i);
      const atStart = out.length === 0;
      const allowed = atStart ? isXmlNameStartChar(code) : isXmlNameChar(code);
      out += allowed ? ch : encodeCodePoint(code);
    }
    if (!out) out = encodeCodePoint(raw.charCodeAt(0));
    if (out.length > MAX_INTERNAL_NAME_LENGTH) out = out.slice(0, MAX_INTERNAL_NAME_LENGTH);
    return out;
  }

  function normalizeKey(name) {
    return String(name || "").trim().toLowerCase();
  }

  function fieldIndex(fields) {
    const byInternal = new Map();
    const byTitle = new Map();
    for (const field of fields || []) {
      const internal = String(field.internalName || field.InternalName || "").trim();
      const title = String(field.title || field.Title || "").trim();
      if (internal) byInternal.set(normalizeKey(internal), field);
      if (title) byTitle.set(normalizeKey(title), field);
    }
    return { byInternal: byInternal, byTitle: byTitle };
  }

  function existingInternalNames(fields) {
    const names = new Set();
    for (const field of fields || []) {
      const internal = String(field.internalName || field.InternalName || "").trim();
      if (internal) names.add(internal);
    }
    return names;
  }

  function fitWithSuffix(base, suffix) {
    const maxBase = Math.max(1, MAX_INTERNAL_NAME_LENGTH - suffix.length);
    return base.slice(0, maxBase) + suffix;
  }

  function resolveUniqueInternalName(base, taken) {
    const encoded = encodeInternalNameFromTitle(base);
    if (!encoded) return "";
    const used = new Set();
    for (const name of taken || []) {
      const key = String(name || "").trim();
      if (key) used.add(normalizeKey(key));
    }
    if (!used.has(normalizeKey(encoded))) return encoded;
    for (let i = 0; i < 1000; i++) {
      const suffix = String(i);
      const candidate = fitWithSuffix(encoded, suffix);
      if (!used.has(normalizeKey(candidate))) return candidate;
    }
    return fitWithSuffix(encoded, "0");
  }

  function crawledForInternal(internalName) {
    const internal = String(internalName || "");
    if (!internal) return "";
    let crawled = "ows_" + internal;
    if (/_x003a__x0020_/i.test(internal)) {
      crawled = "ows_" + internal.replace(/_x003a__x0020_/gi, ":_x0020_");
    }
    return crawled;
  }

  function predictColumnInternalName(input) {
    const title = String((input && input.title) || "").trim();
    const placement = input && (input.placement === "site" || input.placement === "listFromSite") ? input.placement : "listNew";
    const siteFields = Array.isArray(input && input.siteFields) ? input.siteFields : [];
    const listFields = Array.isArray(input && input.listFields) ? input.listFields : [];
    const siteIndex = fieldIndex(siteFields);
    const listIndex = fieldIndex(listFields);
    const listNames = existingInternalNames(listFields);
    const siteNames = existingInternalNames(siteFields);
    const notes = [];

    if (!title) {
      return {
        internalName: "",
        placement: placement,
        reusedSiteColumn: false,
        alreadyOnList: false,
        conflictResolved: false,
        notes: ["Enter a display name to preview the internal name."],
        crawledProperty: "",
      };
    }

    const siteMatch = siteIndex.byTitle.get(normalizeKey(title));
    const listMatch = listIndex.byTitle.get(normalizeKey(title));

    if (placement === "listFromSite") {
      if (!siteMatch) {
        const encoded = resolveUniqueInternalName(title, listNames);
        notes.push("No site column with this display name yet; a new site column would be created first.");
        return {
          internalName: encoded,
          placement: placement,
          reusedSiteColumn: false,
          alreadyOnList: listIndex.byInternal.has(normalizeKey(encoded)),
          conflictResolved: encoded !== encodeInternalNameFromTitle(title),
          notes: notes,
          crawledProperty: crawledForInternal(encoded),
        };
      }
      const internal = String(siteMatch.internalName || siteMatch.InternalName || "").trim();
      const onList = listIndex.byInternal.has(normalizeKey(internal));
      if (onList) notes.push("This site column is already on the current list.");
      else notes.push("Reuses the existing site column static name when added to the list.");
      return {
        internalName: internal,
        placement: placement,
        reusedSiteColumn: true,
        alreadyOnList: onList,
        conflictResolved: false,
        notes: notes,
        crawledProperty: crawledForInternal(internal),
      };
    }

    if (placement === "site") {
      if (siteMatch) {
        const internal = String(siteMatch.internalName || siteMatch.InternalName || "").trim();
        notes.push("A site column with this display name already exists.");
        return {
          internalName: internal,
          placement: placement,
          reusedSiteColumn: true,
          alreadyOnList: listIndex.byInternal.has(normalizeKey(internal)),
          conflictResolved: false,
          notes: notes,
          crawledProperty: crawledForInternal(internal),
        };
      }
      const encoded = resolveUniqueInternalName(title, siteNames);
      if (encoded !== encodeInternalNameFromTitle(title)) notes.push("Adjusted to avoid a duplicate static name on the site.");
      return {
        internalName: encoded,
        placement: placement,
        reusedSiteColumn: false,
        alreadyOnList: listIndex.byInternal.has(normalizeKey(encoded)),
        conflictResolved: encoded !== encodeInternalNameFromTitle(title),
        notes: notes,
        crawledProperty: crawledForInternal(encoded),
      };
    }

    if (siteMatch) {
      const siteInternal = String(siteMatch.internalName || siteMatch.InternalName || "").trim();
      if (!listIndex.byInternal.has(normalizeKey(siteInternal))) {
        notes.push("A site column with this display name exists; add it from the site instead to reuse " + siteInternal + ".");
      }
    }
    if (listMatch) {
      const internal = String(listMatch.internalName || listMatch.InternalName || "").trim();
      notes.push("A list column with this display name already exists.");
      return {
        internalName: internal,
        placement: placement,
        reusedSiteColumn: false,
        alreadyOnList: true,
        conflictResolved: false,
        notes: notes,
        crawledProperty: crawledForInternal(internal),
      };
    }

    const encoded = resolveUniqueInternalName(title, listNames);
    if (encoded !== encodeInternalNameFromTitle(title)) notes.push("Adjusted to avoid a duplicate static name on this list.");
    return {
      internalName: encoded,
      placement: placement,
      reusedSiteColumn: false,
      alreadyOnList: false,
      conflictResolved: encoded !== encodeInternalNameFromTitle(title),
      notes: notes,
      crawledProperty: crawledForInternal(encoded),
    };
  }

  const COLUMN_TYPES = [
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
  const DEFAULT_COLUMN_GROUP = "Custom Columns";

  function getColumnTypeDef(typeId) {
    for (let i = 0; i < COLUMN_TYPES.length; i++) {
      if (COLUMN_TYPES[i].id === typeId) return COLUMN_TYPES[i];
    }
    return COLUMN_TYPES[0];
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
      .map(function (c) { return String(c || "").trim(); })
      .filter(Boolean);
    if (!items.length) return "<CHOICES><CHOICE>Option 1</CHOICE></CHOICES>";
    return "<CHOICES>" + items.map(function (c) { return "<CHOICE>" + escapeXmlText(c) + "</CHOICE>"; }).join("") + "</CHOICES>";
  }

  function buildFieldSchemaXml(input) {
    const def = getColumnTypeDef(input && input.typeId);
    const name = String((input && input.internalName) || "").trim();
    const title = String((input && input.title) || "").trim() || name;
    const group = String((input && input.group) || DEFAULT_COLUMN_GROUP).trim() || DEFAULT_COLUMN_GROUP;
    const required = !!(input && input.required);
    const description = String((input && input.description) || "").trim();
    const attrs = [
      'Type="' + escapeXmlAttr(def.type) + '"',
      'Name="' + escapeXmlAttr(name) + '"',
      'DisplayName="' + escapeXmlAttr(title) + '"',
      'Group="' + escapeXmlAttr(group) + '"',
      'Required="' + (required ? "TRUE" : "FALSE") + '"',
    ];
    if (description) attrs.push('Description="' + escapeXmlAttr(description) + '"');
    if (def.type === "Note") {
      const lines = parseInt(input && input.numLines, 10);
      attrs.push('NumLines="' + (lines > 0 ? lines : 6) + '"');
    }
    if (def.supportsNumberFormat) {
      const decimals = parseInt(input && input.decimals, 10);
      if (decimals >= 0 && decimals <= 5) attrs.push('Decimals="' + decimals + '"');
    }
    if (def.supportsDateFormat) {
      const fmt = input && input.dateFormat === "DateOnly" ? "DateOnly" : "DateTime";
      attrs.push('Format="' + fmt + '"');
    }
    if (def.supportsUserMode) {
      const mode = input && input.userSelectionMode === "PeopleOnly" ? "PeopleOnly" : "PeopleAndGroups";
      attrs.push('UserSelectionMode="' + mode + '"');
    }
    let inner = "";
    if (def.supportsChoices) inner = choiceXml(input && input.choices);
    return "<Field " + attrs.join(" ") + ">" + inner + "</Field>";
  }

  function escapeHtml(s) {
    const d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function parseChoices(raw) {
    return String(raw || "")
      .split(/\r?\n/)
      .map(function (line) { return line.trim(); })
      .filter(Boolean);
  }

  function mountColumnCreator(host, options) {
    if (!host || host.dataset.columnCreatorMounted === "1") return;
    host.dataset.columnCreatorMounted = "1";
    const invoke = options.invoke;
    const onCreated = options.onCreated || function () {};
    const hasList = !!options.hasList;

    const toggleBtn = options.toggleBtn || null;

    host.innerHTML =
      '<div class="column-creator" hidden>' +
      '<label class="column-creator-label" for="columnCreatorPlacement">Where to create</label>' +
      '<select id="columnCreatorPlacement" class="column-creator-placement">' +
      (hasList
        ? '<option value="listNew">This list / library only</option><option value="listFromSite">Add existing site column</option>'
        : "") +
      '<option value="site">' +
      (hasList
        ? "Site column on this list (available for other lists)"
        : "Site column catalog (add to lists when needed)") +
      "</option>" +
      "</select>" +
      '<div class="column-creator-create-fields">' +
      '<label class="column-creator-label" for="columnCreatorTitle">Display name</label>' +
      '<input id="columnCreatorTitle" class="column-creator-title-input" type="text" autocomplete="off" placeholder="e.g. Project phase" />' +
      '<div class="column-creator-preview">' +
      '<span class="column-creator-preview-label">Predicted internal name</span>' +
      '<code id="columnCreatorInternalPreview" class="column-creator-internal"></code>' +
      '<div id="columnCreatorCrawledPreview" class="column-creator-crawled"></div>' +
      "</div>" +
      '<div id="columnCreatorNotes" class="column-creator-notes" aria-live="polite"></div>' +
      '<label class="column-creator-label" for="columnCreatorType">Type</label>' +
      '<select id="columnCreatorType" class="column-creator-type"></select>' +
      '<div id="columnCreatorTypeOptions" class="column-creator-type-options"></div>' +
      '<label class="column-creator-label" for="columnCreatorGroup">Group</label>' +
      '<input id="columnCreatorGroup" class="column-creator-group" type="text" value="' +
      escapeHtml(DEFAULT_COLUMN_GROUP) +
      '" />' +
      '<label class="column-creator-check"><input type="checkbox" id="columnCreatorRequired" /><span>Required</span></label>' +
      '<label class="column-creator-label" for="columnCreatorDescription">Description</label>' +
      '<textarea id="columnCreatorDescription" class="column-creator-description" rows="2" placeholder="Optional"></textarea>' +
      '<div class="column-creator-actions">' +
      '<button type="button" id="columnCreatorSubmit" class="column-creator-submit">Create column</button>' +
      '<span id="columnCreatorStatus" class="column-creator-status" aria-live="polite"></span>' +
      "</div></div>" +
      '<div class="column-creator-site-pick" hidden>' +
      '<input id="columnCreatorSiteFilter" class="column-creator-site-filter" type="text" autocomplete="off" placeholder="Filter site columns…" />' +
      '<div id="columnCreatorSiteList" class="column-creator-site-list" role="list"></div>' +
      '<span id="columnCreatorSiteStatus" class="column-creator-status" aria-live="polite"></span>' +
      "</div></div>";

    const panel = host.querySelector(".column-creator");
    const createFields = host.querySelector(".column-creator-create-fields");
    const sitePick = host.querySelector(".column-creator-site-pick");
    const placementEl = host.querySelector("#columnCreatorPlacement");
    const titleEl = host.querySelector("#columnCreatorTitle");
    const internalEl = host.querySelector("#columnCreatorInternalPreview");
    const crawledEl = host.querySelector("#columnCreatorCrawledPreview");
    const notesEl = host.querySelector("#columnCreatorNotes");
    const typeEl = host.querySelector("#columnCreatorType");
    const typeOptionsEl = host.querySelector("#columnCreatorTypeOptions");
    const groupEl = host.querySelector("#columnCreatorGroup");
    const requiredEl = host.querySelector("#columnCreatorRequired");
    const descriptionEl = host.querySelector("#columnCreatorDescription");
    const submitBtn = host.querySelector("#columnCreatorSubmit");
    const statusEl = host.querySelector("#columnCreatorStatus");
    const siteFilterEl = host.querySelector("#columnCreatorSiteFilter");
    const siteListEl = host.querySelector("#columnCreatorSiteList");
    const siteStatusEl = host.querySelector("#columnCreatorSiteStatus");

    COLUMN_TYPES.forEach(function (t) {
      const opt = document.createElement("option");
      opt.value = t.id;
      opt.textContent = t.label;
      typeEl.appendChild(opt);
    });

    let context = { siteFields: [], listFields: [], siteUrl: "", listId: "" };
    let prediction = predictColumnInternalName({
      title: "",
      placement: hasList ? "listNew" : "site",
      siteFields: [],
      listFields: [],
    });

    let contextLoaded = false;

    function placementValue() {
      const v = placementEl ? placementEl.value : "site";
      return v === "site" || v === "listFromSite" ? v : "listNew";
    }

    function listInternalKeys() {
      const keys = new Set();
      (context.listFields || []).forEach(function (field) {
        const internal = String(field.internalName || field.InternalName || "").trim();
        if (internal) keys.add(normalizeKey(internal));
      });
      return keys;
    }

    function addableSiteFields() {
      const onList = listInternalKeys();
      return (context.siteFields || []).filter(function (field) {
        const internal = String(field.internalName || field.InternalName || "").trim();
        return internal && !onList.has(normalizeKey(internal));
      });
    }

    function renderTypeOptions() {
      const def = getColumnTypeDef(typeEl.value);
      typeOptionsEl.innerHTML = "";
      if (def.supportsChoices) {
        typeOptionsEl.innerHTML =
          '<label class="column-creator-label" for="columnCreatorChoices">Choices (one per line)</label>' +
          '<textarea id="columnCreatorChoices" class="column-creator-choices" rows="4">Option 1\nOption 2</textarea>';
      } else if (def.supportsDateFormat) {
        typeOptionsEl.innerHTML =
          '<label class="column-creator-label" for="columnCreatorDateFormat">Date format</label>' +
          '<select id="columnCreatorDateFormat"><option value="DateOnly">Date only</option><option value="DateTime">Date &amp; time</option></select>';
      } else if (def.supportsNumberFormat) {
        typeOptionsEl.innerHTML =
          '<label class="column-creator-label" for="columnCreatorDecimals">Decimal places</label>' +
          '<input id="columnCreatorDecimals" type="number" min="0" max="5" value="0" />';
      } else if (def.supportsUserMode) {
        typeOptionsEl.innerHTML =
          '<label class="column-creator-label" for="columnCreatorUserMode">Allow selection of</label>' +
          '<select id="columnCreatorUserMode"><option value="PeopleAndGroups">People and Groups</option><option value="PeopleOnly">People only</option></select>';
      } else if (def.type === "Note") {
        typeOptionsEl.innerHTML =
          '<label class="column-creator-label" for="columnCreatorNumLines">Number of lines</label>' +
          '<input id="columnCreatorNumLines" type="number" min="3" max="20" value="6" />';
      }
    }

    function renderPrediction() {
      prediction = predictColumnInternalName({
        title: titleEl.value,
        placement: placementValue(),
        siteFields: context.siteFields,
        listFields: context.listFields,
      });
      if (internalEl) internalEl.textContent = prediction.internalName || "—";
      if (crawledEl) {
        crawledEl.textContent = prediction.crawledProperty
          ? "Crawled property: " + prediction.crawledProperty
          : "";
      }
      if (notesEl) {
        notesEl.innerHTML = (prediction.notes || [])
          .map(function (n) { return '<p class="column-creator-note">' + escapeHtml(n) + "</p>"; })
          .join("");
      }
      if (submitBtn) submitBtn.disabled = !String(titleEl.value || "").trim();
    }

    function setCreateStatus(text, isError) {
      if (!statusEl) return;
      statusEl.textContent = text || "";
      statusEl.className = "column-creator-status" + (isError ? " err" : text ? " ok" : "");
    }

    function setSiteStatus(text, isError) {
      if (!siteStatusEl) return;
      siteStatusEl.textContent = text || "";
      siteStatusEl.className = "column-creator-status" + (isError ? " err" : text ? " ok" : "");
    }

    function renderSiteColumnList() {
      if (!siteListEl) return;
      const q = siteFilterEl ? String(siteFilterEl.value || "").toLowerCase().trim() : "";
      const fields = addableSiteFields().filter(function (field) {
        if (!q) return true;
        const title = String(field.title || field.Title || "").toLowerCase();
        const internal = String(field.internalName || field.InternalName || "").toLowerCase();
        const type = String(field.type || field.TypeAsString || "").toLowerCase();
        return title.indexOf(q) >= 0 || internal.indexOf(q) >= 0 || type.indexOf(q) >= 0;
      });
      siteListEl.innerHTML = "";
      if (!fields.length) {
        siteListEl.innerHTML = '<div class="column-creator-site-empty">No site columns are available to add.</div>';
        return;
      }
      fields.forEach(function (field) {
        const internal = String(field.internalName || field.InternalName || "").trim();
        const title = String(field.title || field.Title || internal).trim();
        const type = String(field.type || field.TypeAsString || "").trim();
        const row = document.createElement("div");
        row.className = "column-creator-site-row";
        row.setAttribute("role", "listitem");
        row.innerHTML =
          '<span class="column-creator-site-title">' + escapeHtml(title) + "</span>" +
          '<span class="column-creator-site-internal">' + escapeHtml(internal) + "</span>" +
          '<span class="column-creator-site-type">' + escapeHtml(type) + "</span>" +
          '<button type="button" class="column-creator-site-add">Add</button>';
        const addBtn = row.querySelector(".column-creator-site-add");
        addBtn.addEventListener("click", function () {
          addSiteColumn(internal, title, addBtn);
        });
        siteListEl.appendChild(row);
      });
    }

    function addSiteColumn(internalName, title, btn) {
      if (!internalName) return;
      setSiteStatus("Checking current list…", false);
      if (btn) btn.disabled = true;
      loadContext().then(function (loaded) {
        if (!loaded) {
          if (btn) btn.disabled = false;
          return;
        }
        const freshSiteField = (context.siteFields || []).find(function (field) {
          return normalizeKey(field.internalName || field.InternalName) === normalizeKey(internalName);
        });
        if (!context.listId || !freshSiteField) {
          if (btn) btn.disabled = false;
          setSiteStatus(
            !context.listId
              ? "Open a list or library view before adding a site column."
              : "That site column is not available on the current site.",
            true
          );
          return;
        }
        setSiteStatus("Adding…", false);
        return invoke({
          action: "createColumn",
          siteUrl: context.siteUrl,
          listId: context.listId,
          target: "listFromSite",
          schemaXml: "",
          siteFieldInternal: internalName,
        }).then(function (res) {
          if (btn) btn.disabled = false;
          if (!res || !res.ok) {
            setSiteStatus((res && res.error) || "Add failed.", true);
            return;
          }
          setSiteStatus("Added " + (res.title || title || internalName) + ".", false);
          loadContext().then(function () {
            renderSiteColumnList();
            onCreated();
          });
        });
      }).catch(function (err) {
        if (btn) btn.disabled = false;
        setSiteStatus((err && err.message) || "Add failed.", true);
      });
    }

    function updatePlacementUI() {
      const placement = placementValue();
      if (placement === "listFromSite") {
        if (createFields) createFields.hidden = true;
        if (sitePick) sitePick.hidden = false;
        renderSiteColumnList();
        return;
      }
      if (createFields) createFields.hidden = false;
      if (sitePick) sitePick.hidden = true;
      renderPrediction();
    }

    function loadContext() {
      const resolveContext = window.SPOT_resolveCompassActionContext;
      if (typeof resolveContext !== "function") {
        contextLoaded = false;
        setCreateStatus("Could not verify the current SharePoint page.", true);
        setSiteStatus("Could not verify the current SharePoint page.", true);
        return Promise.resolve(false);
      }
      return resolveContext(invoke, "getColumnCreatorContext", false).then(function (res) {
        if (!res) {
          context = { siteFields: [], listFields: [], siteUrl: "", listId: "" };
          contextLoaded = false;
          setCreateStatus("Could not load column context.", true);
          setSiteStatus("Could not load column context.", true);
          renderPrediction();
          renderSiteColumnList();
          return false;
        }
        context = {
          siteFields: res.siteFields || [],
          listFields: res.listFields || [],
          siteUrl: res.siteUrl || "",
          listId: res.listId || "",
        };
        contextLoaded = true;
        setCreateStatus("", false);
        setSiteStatus("", false);
        updatePlacementUI();
        return true;
      }).catch(function (err) {
        context = { siteFields: [], listFields: [], siteUrl: "", listId: "" };
        contextLoaded = false;
        const message = (err && err.message) || "Could not load column context.";
        setCreateStatus(message, true);
        setSiteStatus(message, true);
        renderPrediction();
        renderSiteColumnList();
        return false;
      });
    }

    function submit() {
      const title = String(titleEl.value || "").trim();
      if (!title) return;
      setCreateStatus("Checking current page…", false);
      if (submitBtn) submitBtn.disabled = true;
      loadContext().then(function (loaded) {
        if (!loaded) {
          if (submitBtn) submitBtn.disabled = false;
          return;
        }
        renderPrediction();
        if (!prediction.internalName) {
          if (submitBtn) submitBtn.disabled = false;
          return;
        }

        const choicesEl = host.querySelector("#columnCreatorChoices");
        const dateEl = host.querySelector("#columnCreatorDateFormat");
        const decimalsEl = host.querySelector("#columnCreatorDecimals");
        const userModeEl = host.querySelector("#columnCreatorUserMode");
        const numLinesEl = host.querySelector("#columnCreatorNumLines");
        const schemaXml = buildFieldSchemaXml({
          internalName: prediction.internalName,
          title: title,
          typeId: typeEl.value,
          required: !!(requiredEl && requiredEl.checked),
          group: groupEl ? groupEl.value : DEFAULT_COLUMN_GROUP,
          description: descriptionEl ? descriptionEl.value : "",
          choices: choicesEl ? parseChoices(choicesEl.value) : undefined,
          dateFormat: dateEl ? dateEl.value : undefined,
          decimals: decimalsEl ? parseInt(decimalsEl.value, 10) : undefined,
          userSelectionMode: userModeEl ? userModeEl.value : undefined,
          numLines: numLinesEl ? parseInt(numLinesEl.value, 10) : undefined,
        });
        const placement = placementValue();
        const target = placement === "site" ? "site" : "list";
        if (target === "list" && !context.listId) {
          if (submitBtn) submitBtn.disabled = false;
          setCreateStatus("Open a list or library view before creating a list column.", true);
          return;
        }
        setCreateStatus("Creating…", false);
        return invoke({
          action: "createColumn",
          siteUrl: context.siteUrl,
          listId: context.listId,
          target: target,
          schemaXml: schemaXml,
          siteFieldInternal: prediction.internalName,
        }).then(function (res) {
          if (submitBtn) submitBtn.disabled = false;
          if (!res || !res.ok) {
            setCreateStatus((res && res.error) || "Create failed.", true);
            renderPrediction();
            return;
          }
          setCreateStatus("Created " + (res.title || res.internalName || title) + ".", false);
          titleEl.value = "";
          loadContext().then(function () {
            onCreated();
            renderPrediction();
          });
        });
      }).catch(function (err) {
        if (submitBtn) submitBtn.disabled = false;
        setCreateStatus((err && err.message) || "Create failed.", true);
      });
    }

    function setPanelOpen(open) {
      if (!panel) return;
      if (open) panel.removeAttribute("hidden");
      else panel.setAttribute("hidden", "");
      if (toggleBtn) {
        toggleBtn.setAttribute("aria-expanded", open ? "true" : "false");
        toggleBtn.classList.toggle("column-creator-open", open);
      }
      if (open && !contextLoaded) loadContext();
      else if (open) updatePlacementUI();
    }

    if (toggleBtn) {
      toggleBtn.addEventListener("click", function () {
        const open = panel && panel.hasAttribute("hidden");
        setPanelOpen(!!open);
      });
    }
    if (placementEl) placementEl.addEventListener("change", updatePlacementUI);
    if (titleEl) titleEl.addEventListener("input", renderPrediction);
    if (typeEl) {
      typeEl.addEventListener("change", function () {
        renderTypeOptions();
        renderPrediction();
      });
    }
    if (submitBtn) submitBtn.addEventListener("click", submit);
    if (siteFilterEl) siteFilterEl.addEventListener("input", renderSiteColumnList);

    renderTypeOptions();
    renderPrediction();
  }

  window.SPOT_mountColumnCreator = mountColumnCreator;
})();
