import { predictColumnInternalName } from "./columnInternalName.mjs";
import {
  COLUMN_TYPES,
  DEFAULT_COLUMN_GROUP,
  buildFieldSchemaXml,
  getColumnTypeDef,
} from "./columnFieldSchema.mjs";

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s == null ? "" : String(s);
  return d.innerHTML;
}

function parseChoices(raw) {
  return String(raw || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function normalizeKey(name) {
  return String(name || "").trim().toLowerCase();
}

/**
 * @param {HTMLElement} host
 * @param {object} options
 * @param {(msg: object) => Promise<object>} options.invoke
 * @param {() => void} [options.onCreated]
 * @param {boolean} [options.hasList]
 * @param {HTMLElement} [options.toggleBtn]
 */
export function mountColumnCreator(host, options) {
  if (!host || host.dataset.columnCreatorMounted === "1") return;
  host.dataset.columnCreatorMounted = "1";
  const invoke = options.invoke;
  const onCreated = options.onCreated || (() => {});
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
    "</div>" +
    '<div class="column-creator-site-pick" hidden>' +
    '<input id="columnCreatorSiteFilter" class="column-creator-site-filter" type="text" autocomplete="off" placeholder="Filter site columns…" />' +
    '<div id="columnCreatorSiteList" class="column-creator-site-list" role="list"></div>' +
    '<span id="columnCreatorSiteStatus" class="column-creator-status" aria-live="polite"></span>' +
    "</div>" +
    '<div class="column-creator-actions">' +
    '<button type="button" id="columnCreatorSubmit" class="column-creator-submit">Create column</button>' +
    '<button type="button" id="columnCreatorCancel" class="column-creator-cancel">Cancel</button>' +
    '<span id="columnCreatorStatus" class="column-creator-status" aria-live="polite"></span>' +
    "</div></div>";

  const panel = host.querySelector(".column-creator");
  const placementEl = host.querySelector("#columnCreatorPlacement");
  const createFields = host.querySelector(".column-creator-create-fields");
  const sitePick = host.querySelector(".column-creator-site-pick");
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
  const cancelBtn = host.querySelector("#columnCreatorCancel");
  const statusEl = host.querySelector("#columnCreatorStatus");
  const siteFilterEl = host.querySelector("#columnCreatorSiteFilter");
  const siteListEl = host.querySelector("#columnCreatorSiteList");
  const siteStatusEl = host.querySelector("#columnCreatorSiteStatus");

  COLUMN_TYPES.forEach((t) => {
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
    (context.listFields || []).forEach((field) => {
      const internal = String(field.internalName || field.InternalName || "").trim();
      if (internal) keys.add(normalizeKey(internal));
    });
    return keys;
  }

  function addableSiteFields() {
    const onList = listInternalKeys();
    return (context.siteFields || []).filter((field) => {
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
        .map((n) => '<p class="column-creator-note">' + escapeHtml(n) + "</p>")
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
    const fields = addableSiteFields().filter((field) => {
      if (!q) return true;
      const title = String(field.title || field.Title || "").toLowerCase();
      const internal = String(field.internalName || field.InternalName || "").toLowerCase();
      const type = String(field.type || field.TypeAsString || "").toLowerCase();
      return title.includes(q) || internal.includes(q) || type.includes(q);
    });
    siteListEl.innerHTML = "";
    if (!fields.length) {
      siteListEl.innerHTML = '<div class="column-creator-site-empty">No site columns are available to add.</div>';
      return;
    }
    fields.forEach((field) => {
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
      addBtn.addEventListener("click", () => {
        addSiteColumn(internal, title, addBtn);
      });
      siteListEl.appendChild(row);
    });
  }

  async function addSiteColumn(internalName, title, btn) {
    if (!internalName) return;
    setSiteStatus("Adding…", false);
    if (btn) btn.disabled = true;
    const res = await invoke({
      action: "createColumn",
      siteUrl: context.siteUrl,
      listId: context.listId,
      target: "listFromSite",
      schemaXml: "",
      siteFieldInternal: internalName,
    }).catch((err) => ({ ok: false, error: err?.message || "Add failed." }));
    if (btn) btn.disabled = false;
    if (!res?.ok) {
      setSiteStatus(res?.error || "Add failed.", true);
      return;
    }
    setSiteStatus("Added " + (res.title || title || internalName) + ".", false);
    await loadContext();
    renderSiteColumnList();
    onCreated();
  }

  function updatePlacementUI() {
    const placement = placementValue();
    if (placement === "listFromSite") {
      if (createFields) createFields.hidden = true;
      if (sitePick) sitePick.hidden = false;
      if (submitBtn) submitBtn.hidden = true;
      renderSiteColumnList();
      return;
    }
    if (createFields) createFields.hidden = false;
    if (sitePick) sitePick.hidden = true;
    if (submitBtn) submitBtn.hidden = false;
    renderPrediction();
  }

  async function loadContext() {
    const res = await invoke({ action: "getColumnCreatorContext" });
    if (!res?.ok) {
      context = { siteFields: [], listFields: [], siteUrl: "", listId: "" };
      setCreateStatus(res?.error || "Could not load column context.", true);
      setSiteStatus(res?.error || "Could not load column context.", true);
      renderPrediction();
      renderSiteColumnList();
      return;
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
  }

  async function submit() {
    const title = String(titleEl.value || "").trim();
    if (!title) return;
    renderPrediction();
    if (!prediction.internalName) return;

    const choicesEl = host.querySelector("#columnCreatorChoices");
    const dateEl = host.querySelector("#columnCreatorDateFormat");
    const decimalsEl = host.querySelector("#columnCreatorDecimals");
    const userModeEl = host.querySelector("#columnCreatorUserMode");
    const numLinesEl = host.querySelector("#columnCreatorNumLines");
    const schemaXml = buildFieldSchemaXml({
      internalName: prediction.internalName,
      title,
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
    setCreateStatus("Creating…", false);
    if (submitBtn) submitBtn.disabled = true;
    const res = await invoke({
      action: "createColumn",
      siteUrl: context.siteUrl,
      listId: context.listId,
      target,
      schemaXml,
      siteFieldInternal: prediction.internalName,
    });
    if (submitBtn) submitBtn.disabled = false;
    if (!res?.ok) {
      setCreateStatus(res?.error || "Create failed.", true);
      renderPrediction();
      return;
    }
    setCreateStatus("Created " + (res.title || res.internalName || title) + ".", false);
    titleEl.value = "";
    await loadContext();
    onCreated();
    renderPrediction();
  }

  function setPanelOpen(open) {
    if (!panel) return;
    if (open) panel.removeAttribute("hidden");
    else panel.setAttribute("hidden", "");
    if (toggleBtn) {
      toggleBtn.setAttribute("aria-expanded", open ? "true" : "false");
      toggleBtn.classList.toggle("column-creator-open", open);
    }
    if (open && !contextLoaded) void loadContext();
    else if (open) updatePlacementUI();
  }

  if (toggleBtn) {
    toggleBtn.addEventListener("click", () => {
      const open = panel && panel.hasAttribute("hidden");
      setPanelOpen(!!open);
    });
  }

  placementEl?.addEventListener("change", updatePlacementUI);
  titleEl?.addEventListener("input", renderPrediction);
  typeEl?.addEventListener("change", () => {
    renderTypeOptions();
    renderPrediction();
  });
  submitBtn?.addEventListener("click", submit);
  cancelBtn?.addEventListener("click", () => setPanelOpen(false));
  siteFilterEl?.addEventListener("input", renderSiteColumnList);

  renderTypeOptions();
  renderPrediction();
}
