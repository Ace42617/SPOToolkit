// Injected: creates a SharePoint field from SchemaXml (site or list).
(function () {
  var PARAMS_SCRIPT_ID = "sp-column-create-params";
  var activeRequestId = getCurrentRequestId();

  function getCurrentRequestId() {
    try {
      var script = document.currentScript;
      var src = script && script.src ? String(script.src) : "";
      if (!src) return "";
      return new URL(src, window.location.href).searchParams.get("spcsvRequestId") || "";
    } catch (e) {
      return "";
    }
  }

  function getParamsScriptId(requestId) {
    return requestId ? PARAMS_SCRIPT_ID + "-" + requestId : PARAMS_SCRIPT_ID;
  }

  function readParams() {
    try {
      var el = document.getElementById(getParamsScriptId(activeRequestId));
      if (!el && activeRequestId) el = document.getElementById(PARAMS_SCRIPT_ID);
      if (el && el.textContent) {
        var parsed = JSON.parse(el.textContent);
        if (!activeRequestId && parsed && parsed.requestId) {
          activeRequestId = String(parsed.requestId);
        }
        return parsed;
      }
    } catch (_) {}
    return {};
  }

  var params = readParams();

  function post(payload) {
    var msg = Object.assign({ __spcsv: true, type: "SPCSVCreateColumnResult" }, payload);
    if (activeRequestId) msg.requestId = activeRequestId;
    window.postMessage(msg, "*");
  }

  function withTimeout(promise, ms, label) {
    return new Promise(function (resolve, reject) {
      var timer = setTimeout(function () {
        reject(new Error(label + " timed out."));
      }, ms);
      promise.then(
        function (value) {
          clearTimeout(timer);
          resolve(value);
        },
        function (err) {
          clearTimeout(timer);
          reject(err);
        }
      );
    });
  }

  function readDigest(j) {
    if (!j) return "";
    if (j.FormDigestValue) return j.FormDigestValue;
    if (j.d && j.d.GetContextWebInformation && j.d.GetContextWebInformation.FormDigestValue) {
      return j.d.GetContextWebInformation.FormDigestValue;
    }
    return "";
  }

  function readErrorMessage(j, text, statusText) {
    if (j && j.error) {
      if (j.error.message && j.error.message.value) return j.error.message.value;
      if (j.error.message) return j.error.message;
    }
    if (j && j["odata.error"] && j["odata.error"].message && j["odata.error"].message.value) {
      return j["odata.error"].message.value;
    }
    if (text) return text;
    return statusText || "Request failed.";
  }

  var siteUrl = String(params.siteUrl || "").replace(/\/$/, "");
  var listId = String(params.listId || "").replace(/[{}]/g, "");
  var target = params.target === "site" || params.target === "listFromSite" ? params.target : "list";
  var schemaXml = String(params.schemaXml || "");
  var siteFieldInternal = String(params.siteFieldInternal || "").trim();

  if (!siteUrl) {
    post({ ok: false, error: "Missing site URL." });
    return;
  }
  if ((target === "list" || target === "listFromSite") && !listId) {
    post({ ok: false, error: "Open a list or library to create a list column." });
    return;
  }
  if (target === "listFromSite") {
    if (!siteFieldInternal) {
      post({ ok: false, error: "Missing site column internal name." });
      return;
    }
  } else if (!schemaXml) {
    post({ ok: false, error: "Missing field schema." });
    return;
  }

  function requestDigest() {
    return fetch(siteUrl + "/_api/contextinfo", {
      method: "POST",
      credentials: "include",
      headers: { Accept: "application/json;odata=nometadata" },
    })
      .then(function (r) {
        if (!r.ok) throw new Error(r.status + " " + r.statusText);
        return r.json();
      })
      .then(function (j) {
        var digest = readDigest(j);
        if (!digest) throw new Error("Could not get request digest");
        return digest;
      });
  }

  function createFieldAsXml(endpoint, digest, xml) {
    return fetch(endpoint, {
      method: "POST",
      credentials: "include",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
      },
      body: JSON.stringify({
        parameters: {
          SchemaXml: xml,
        },
      }),
    }).then(function (r) {
      return r.text().then(function (text) {
        var j = {};
        try {
          j = JSON.parse(text);
        } catch (_) {}
        if (!r.ok) throw new Error(readErrorMessage(j, text, r.status + " " + r.statusText));
        return j;
      });
    });
  }

  function loadSiteFieldSchema(internalName) {
    var enc = encodeURIComponent("'" + internalName.replace(/'/g, "''") + "'");
    return fetch(siteUrl + "/_api/web/fields/getbyinternalnameortitle(@v)?@v=" + enc + "&$select=SchemaXml,InternalName,Title", {
      credentials: "include",
      headers: { Accept: "application/json;odata=nometadata" },
    })
      .then(function (r) {
        if (!r.ok) throw new Error(r.status + " " + r.statusText);
        return r.json();
      })
      .then(function (j) {
        var xml = j.SchemaXml || "";
        if (!xml) throw new Error("Site column schema was not found.");
        return xml;
      });
  }

  withTimeout(
    requestDigest().then(function (digest) {
      if (target === "site") {
        return createFieldAsXml(siteUrl + "/_api/web/fields/createfieldasxml", digest, schemaXml);
      }
      if (target === "listFromSite") {
        var internal = siteFieldInternal || (schemaXml.match(/Name=\"([^\"]+)\"/i) || [])[1] || "";
        if (!internal) throw new Error("Missing site column internal name.");
        return loadSiteFieldSchema(internal).then(function (siteXml) {
          return createFieldAsXml(
            siteUrl + "/_api/web/lists(guid'" + listId + "')/fields/createfieldasxml",
            digest,
            siteXml
          );
        });
      }
      return createFieldAsXml(
        siteUrl + "/_api/web/lists(guid'" + listId + "')/fields/createfieldasxml",
        digest,
        schemaXml
      );
    }),
    55000,
    "Creating the column"
  )
    .then(function (result) {
      var field = result.d || result;
      post({
        ok: true,
        internalName: field.InternalName || field.Title || "",
        title: field.Title || "",
        type: field.TypeAsString || "",
      });
    })
    .catch(function (err) {
      post({ ok: false, error: (err && err.message) || String(err) });
    });
})();
