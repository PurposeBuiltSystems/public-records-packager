/*
 * Public Records Packager — Microsoft Graph data layer.
 *
 * AUTH: Nested App Authentication (NAA) via MSAL — no backend; identical
 * pattern to the other PurposeBuilt add-ins. Delegated scopes:
 *   Mail.Read        — search the signed-in user's own mail and read MIME
 *   Files.ReadWrite  — write the records bundle to the user's own OneDrive
 * Records never leave the user's Microsoft 365 tenant.
 *
 * Exposes a global `GraphData` object.
 */
/* global msal */
(function (root) {
  "use strict";

  var CLIENT_ID = "3285ecbd-b4b2-4342-8bf6-75b83ac26af7"; // "Public Records Packager" Entra app (purposebuilt.systems tenant)
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var SCOPES = ["Mail.Read", "Files.ReadWrite"];

  var pcaPromise = null;

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: "https://login.microsoftonline.com/common",
        },
      });
    }
    return pcaPromise;
  }

  async function getToken() {
    var pca = await getPca();
    try {
      var silent = await pca.acquireTokenSilent({ scopes: SCOPES });
      return silent.accessToken;
    } catch (e) {
      var interactive = await pca.acquireTokenPopup({ scopes: SCOPES });
      return interactive.accessToken;
    }
  }

  async function graphJson(token, method, path, body) {
    var res = await fetch(GRAPH + path, {
      method: method,
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (!res.ok) { throw new Error("Graph " + method + " " + path + " -> " + res.status + " " + (await res.text())); }
    return res.status === 204 ? null : res.json();
  }

  // ---------- mail ----------

  /**
   * Full-text mailbox search. $search can't combine with $filter, so date and
   * participant narrowing happen client-side (Packager.applyFilters).
   * Capped at ~1000 results; the UI surfaces the cap.
   */
  async function searchMessages(token, query, onProgress) {
    var items = [];
    var url = GRAPH + "/me/messages?$search=" + encodeURIComponent('"' + query.replace(/"/g, "") + '"') +
      "&$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments&$top=25";
    var guard = 0;
    while (url && guard++ < 40) {
      var res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
      if (!res.ok) { throw new Error("Graph search -> " + res.status + " " + (await res.text())); }
      var page = await res.json();
      items = items.concat(page.value || []);
      if (onProgress) { onProgress(items.length); }
      url = page["@odata.nextLink"] || null;
    }
    return { messages: items, capped: guard >= 40 };
  }

  /** Exact MIME (.eml) content of a message. */
  async function getMime(token, messageId) {
    var res = await fetch(GRAPH + "/me/messages/" + messageId + "/$value", {
      headers: { Authorization: "Bearer " + token },
    });
    if (!res.ok) { throw new Error("Graph MIME -> " + res.status); }
    return res.text();
  }

  /** Message body HTML (for the printable copy). */
  async function getBody(token, messageId) {
    var m = await graphJson(token, "GET", "/me/messages/" + messageId + "?$select=body");
    return (m && m.body && m.body.content) || "";
  }

  /** Real file attachments with bytes (fetch individually when omitted). */
  async function getAttachments(token, messageId) {
    var list = await graphJson(token, "GET", "/me/messages/" + messageId + "/attachments");
    var out = [];
    for (var i = 0; i < (list.value || []).length; i++) {
      var a = list.value[i];
      if (a["@odata.type"] !== "#microsoft.graph.fileAttachment") { continue; }
      if (a.isInline) { continue; }
      var bytes = a.contentBytes;
      if (!bytes) {
        var full = await graphJson(token, "GET", "/me/messages/" + messageId + "/attachments/" + a.id);
        bytes = full && full.contentBytes;
      }
      if (bytes) { out.push({ name: a.name, contentBytes: bytes }); }
    }
    return out;
  }

  // ---------- OneDrive ----------

  function b64ToBytes(b64) {
    var bin = atob(b64);
    var arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) { arr[i] = bin.charCodeAt(i); }
    return arr;
  }

  var SMALL_LIMIT = 3800000; // simple PUT is fine below ~4MB

  /**
   * Fail fast with a human message when the user's OneDrive has never been
   * provisioned (licensed but never opened once — /me/drive returns 404).
   */
  async function ensureDrive(token) {
    var res = await fetch(GRAPH + "/me/drive?$select=id", {
      headers: { Authorization: "Bearer " + token },
    });
    if (res.status === 404) {
      throw new Error("Your OneDrive isn't set up yet. Open OneDrive once at https://www.office.com (choose the OneDrive app and let it load) to activate it, then build again. If OneDrive won't open, your Microsoft 365 plan may not include it.");
    }
    if (!res.ok) { throw new Error("OneDrive check -> " + res.status + " " + (await res.text())); }
  }

  var ensuredFolders = {};

  /** Create every folder in `path` that doesn't exist yet (root-relative). */
  async function ensureFolder(token, path) {
    var segs = path.split("/").filter(Boolean);
    var soFar = "";
    for (var i = 0; i < segs.length; i++) {
      var parent = soFar;
      soFar = soFar ? soFar + "/" + segs[i] : segs[i];
      if (ensuredFolders[soFar]) { continue; }
      var url = parent
        ? "/me/drive/root:/" + parent.split("/").map(encodeURIComponent).join("/") + ":/children"
        : "/me/drive/root/children";
      var res = await fetch(GRAPH + url, {
        method: "POST",
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
        body: JSON.stringify({ name: segs[i], folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
      });
      // 409 = already exists — exactly what we want.
      if (!res.ok && res.status !== 409) {
        throw new Error("OneDrive folder \"" + soFar + "\" -> " + res.status + " " + (await res.text()));
      }
      ensuredFolders[soFar] = true;
    }
  }

  /**
   * Upload content to /me/drive at the given path (parent folders must exist —
   * call ensureFolder first). `content` is a string (text) or Uint8Array.
   */
  async function uploadFile(token, path, content) {
    var bytes = typeof content === "string" ? new TextEncoder().encode(content) : content;
    var encodedPath = path.split("/").map(encodeURIComponent).join("/");
    if (bytes.length <= SMALL_LIMIT) {
      var res = await fetch(GRAPH + "/me/drive/root:/" + encodedPath + ":/content", {
        method: "PUT",
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/octet-stream" },
        body: bytes,
      });
      if (!res.ok) { throw new Error("OneDrive PUT " + path + " -> " + res.status + " " + (await res.text())); }
      return res.json();
    }
    // Large file: upload session in 5MB chunks.
    var session = await graphJson(token, "POST", "/me/drive/root:/" + encodedPath + ":/createUploadSession", {
      item: { "@microsoft.graph.conflictBehavior": "replace" },
    });
    var CHUNK = 5 * 1024 * 1024;
    for (var start = 0; start < bytes.length; start += CHUNK) {
      var end = Math.min(start + CHUNK, bytes.length);
      var res2 = await fetch(session.uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(end - start),
          "Content-Range": "bytes " + start + "-" + (end - 1) + "/" + bytes.length,
        },
        body: bytes.slice(start, end),
      });
      if (!res2.ok) { throw new Error("OneDrive chunk " + path + " -> " + res2.status); }
    }
    return {};
  }

  /** Web URL of a OneDrive folder (to show the user where the bundle is). */
  async function folderLink(token, path) {
    try {
      var encodedPath = path.split("/").map(encodeURIComponent).join("/");
      var item = await graphJson(token, "GET", "/me/drive/root:/" + encodedPath + "?$select=webUrl");
      return (item && item.webUrl) || "";
    } catch (e) { return ""; }
  }

  root.GraphData = {
    getToken: getToken,
    ensureDrive: ensureDrive,
    ensureFolder: ensureFolder,
    searchMessages: searchMessages,
    getMime: getMime,
    getBody: getBody,
    getAttachments: getAttachments,
    uploadFile: uploadFile,
    b64ToBytes: b64ToBytes,
    folderLink: folderLink,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
