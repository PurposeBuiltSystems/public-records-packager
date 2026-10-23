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
  var authEpoch = 0;
  var authenticatedAs = null;

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

  // --- add-in sign-out state -------------------------------------------
  //
  // Certification rejected the naive version on a sibling add-in: "after
  // clicking sign-out there is no response or not signed out." The reason is
  // structural. Under nested app authentication Outlook owns the session and
  // getAllAccounts() reports the HOST's account, not a cache this add-in
  // controls - so clearing MSAL's cache changes nothing visible, the next
  // silent acquisition succeeds anyway, and the pane redraws as signed in.
  //
  // A sign-out this add-in cannot deliver should not be offered. What it CAN
  // deliver is refusing to act until the user authenticates again: while
  // signed out it reports itself signed out and will not use a silent token,
  // so the next action raises a real prompt. Outlook's own session is
  // untouched, and the pane says so.
  var SIGNED_OUT_KEY = "addinSignedOut";
  var signedOut = false;
  try { signedOut = Office.context.roamingSettings.get(SIGNED_OUT_KEY) === true; } catch (e) { signedOut = false; }

  function setSignedOut(v) {
    signedOut = !!v;
    try {
      Office.context.roamingSettings.set(SIGNED_OUT_KEY, signedOut);
      Office.context.roamingSettings.saveAsync(function () {});
    } catch (e) { /* in-memory is still correct for this session */ }
  }

  /** Signed-in account, or null. Reports null while signed out, by design. */

  /**
   * Microsoft Graph throttles, and this add-in makes bursts of calls - a
   * records bundle or a bulk post is dozens to hundreds. An unretried 429
   * aborts the whole run part-way, which is the worst possible failure for
   * work that is half-written. One respectful retry honouring Retry-After
   * absorbs the overwhelming majority of throttling without hammering the
   * service; anything past that is a real outage and should surface.
   */
  async function fetchRetry(url, opts) {
    var res = await fetch(url, opts);
    if (res.status === 429 || res.status === 503) {
      var retryAfter = res.headers.get("Retry-After");
      var wait = retryAfter && /^\d+(\.\d+)?$/.test(retryAfter.trim())
        ? Number(retryAfter) * 1000 : Date.parse(retryAfter) - Date.now();
      if (!isFinite(wait)) { wait = 3000; }
      wait = Math.max(0, wait);
      // Surface long throttles rather than retrying before the server allows it.
      if (wait > 15000) { return res; }
      await new Promise(function (r) { setTimeout(r, wait); });
      res = await fetch(url, opts);
    }
    return res;
  }

  async function currentAccount() {
    if (signedOut) { return null; }
    if (authenticatedAs) { return authenticatedAs; }
    try {
      var pca = await getPca();
      var accts = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      return accts.length ? (accts[0].username || accts[0].name || "signed in") : null;
    } catch (e) { return null; }
  }

  /**
   * Sign out of the add-in. The state flips SYNCHRONOUSLY before any awaiting
   * so the pane can respond instantly - awaiting a broker handshake first is
   * the "no response" half of the finding. Cache clearing is best-effort on
   * top; the enforced state is what makes this real.
   */
  function signOut() {
    ensuredFolders = {};
    authEpoch++;
    authenticatedAs = null;
    setSignedOut(true);
    var pending = pcaPromise;      // only clear what exists; never start a handshake here
    pcaPromise = null;
    if (!pending) { return Promise.resolve(true); }
    return Promise.resolve(pending).then(function (pca) {
      var accts = (pca && pca.getAllAccounts && pca.getAllAccounts()) || [];
      var chain = Promise.resolve();
      accts.forEach(function (a) {
        chain = chain.then(function () {
          if (pca.clearCache) { return pca.clearCache({ account: a }); }
          if (pca.logoutPopup) { return pca.logoutPopup({ account: a }); }
        }).catch(function () { /* best effort; the enforced state stands */ });
      });
      return chain.then(function () { return true; });
    }).catch(function () { return true; });
  }

  /** True while the user has signed the add-in out. */
  function isSignedOut() { return signedOut; }




  /**
   * Sign-in must never hang the pane. An un-timed await on the popup flow
   * leaves a button disabled with nothing visible happening — which reads
   * to the user as "the button does nothing" and gives them nothing to act
   * on. Fail loudly instead, naming the two things that actually fix it.
   */
  function withTimeout(promise, ms, message) {
    var timer;
    return Promise.race([
      promise.then(function (v) { clearTimeout(timer); return v; },
                   function (e) { clearTimeout(timer); throw e; }),
      new Promise(function (_, reject) {
        timer = setTimeout(function () { reject(new Error(message)); }, ms);
      }),
    ]);
  }

  function accountLabel(result) {
    if (result.account) { return result.account.username || result.account.name || "signed in"; }
    try { return Office.context.mailbox.userProfile.emailAddress || "signed in"; }
    catch (e) { return "signed in"; }
  }

  async function getToken() {
    var epoch = authEpoch;
    function checkSession() {
      if (epoch !== authEpoch) { throw new Error("Sign-in cancelled because the add-in session changed."); }
    }
    var pca = await withTimeout(getPca(), 20000,
      "Sign-in didn't start. Fully quit Outlook (Cmd+Q) and reopen, then try again.");
    checkSession();
    try {
      // Signed out means signed out: skip silent so the user must re-authenticate.
      if (signedOut) { throw new Error("signed out of the add-in"); }
      var silent = await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout");
      checkSession();
      authenticatedAs = accountLabel(silent);
      return silent.accessToken;
    } catch (e) {
      checkSession();
      var interactive = await withTimeout(
        pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
        "Sign-in didn't finish. A Microsoft sign-in window may have opened behind Outlook — " +
        "check for it (or Mission Control), finish signing in, and click again. If no window " +
        "appeared at all, fully quit Outlook (Cmd+Q), reopen, and retry.");
      checkSession();
      authenticatedAs = accountLabel(interactive);
      setSignedOut(false);   // a real interactive sign-in ends the signed-out state
      return interactive.accessToken;
    }
  }

  async function graphJson(token, method, path, body) {
    var res = await fetchRetry(GRAPH + path, {
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
      if (!/^https:\/\/graph\.microsoft\.com\/v1\.0\//i.test(url) || /[\u0000-\u001f]/.test(url)) {
        throw new Error("Graph returned an unsupported paging URL.");
      }
      var res = await fetchRetry(url, { headers: { Authorization: "Bearer " + token } });
      if (!res.ok) { throw new Error("Graph search -> " + res.status + " " + (await res.text())); }
      var page = await res.json();
      items = items.concat(page.value || []);
      if (onProgress) { onProgress(items.length); }
      url = page["@odata.nextLink"] || null;
    }
    return { messages: items, capped: !!url };
  }

  /** Exact MIME (.eml) content of a message. */
  async function getMime(token, messageId) {
    var res = await fetchRetry(GRAPH + "/me/messages/" + messageId + "/$value", {
      headers: { Authorization: "Bearer " + token },
    });
    if (!res.ok) { throw new Error("Graph MIME -> " + res.status); }
    return res.text();
  }

  /** Message body HTML (for the printable copy). */
  async function getBody(token, messageId) {
    var m = await graphJson(token, "GET", "/me/messages/" + messageId + "?$select=body");
    var body = (m && m.body) || {};
    var content = body.content || "";
    if (String(body.contentType).toLowerCase() === "text") {
      return "<pre>" + content.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;") + "</pre>";
    }
    return content;
  }

  /** File bytes plus an explicit record of attachment types not exported. */
  async function getAttachmentExport(token, messageId) {
    var values = [];
    var next = "/me/messages/" + messageId + "/attachments";
    var pages = 0;
    while (next && pages++ < 100) {
      var list = await graphJson(token, "GET", next);
      values = values.concat(list.value || []);
      var link = list["@odata.nextLink"];
      if (link && link.indexOf(GRAPH + "/") !== 0) { throw new Error("Unexpected Graph paging URL."); }
      next = link ? link.slice(GRAPH.length) : null;
    }
    if (next) { throw new Error("Attachment results exceeded the scan limit."); }
    var out = [];
    var notExported = [];
    for (var i = 0; i < values.length; i++) {
      var a = values[i];
      if (a.isInline) { continue; }
      if (a["@odata.type"] !== "#microsoft.graph.fileAttachment") {
        notExported.push({ name: a.name || "(unnamed)", type: a["@odata.type"] || "unknown" });
        continue;
      }
      var bytes = a.contentBytes;
      if (typeof bytes !== "string") {
        var full = await graphJson(token, "GET", "/me/messages/" + messageId + "/attachments/" + a.id);
        bytes = full && full.contentBytes;
      }
      if (typeof bytes !== "string") { throw new Error("Attachment content unavailable: " + a.name); }
      out.push({ name: a.name, contentBytes: bytes });
    }
    return { files: out, notExported: notExported };
  }

  // Keep cached older task panes compatible with the original array API.
  async function getAttachments(token, messageId) {
    return (await getAttachmentExport(token, messageId)).files;
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
    var res = await fetchRetry(GRAPH + "/me/drive?$select=id", {
      headers: { Authorization: "Bearer " + token },
    });
    if (res.status === 404) {
      throw new Error("Your OneDrive isn't set up yet. Open OneDrive once at https://www.office.com (choose the OneDrive app and let it load) to activate it, then build again. If OneDrive won't open, your Microsoft 365 plan may not include it.");
    }
    if (!res.ok) { throw new Error("OneDrive check -> " + res.status + " " + (await res.text())); }
  }

  var ensuredFolders = {};

  /** Reserve a new output folder. Never merge a new export with an old bundle. */
  async function createBundleFolder(token, name) {
    await ensureFolder(token, "Records Requests");
    var res = await fetchRetry(GRAPH + "/me/drive/root:/Records%20Requests:/children", {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ name: name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
    });
    if (res.status === 409) {
      throw new Error("That bundle folder already exists. Choose a new folder name so earlier records and review notes are preserved.");
    }
    if (!res.ok) { throw new Error("Could not create the bundle folder (" + res.status + "). Check OneDrive before retrying."); }
    return "Records Requests/" + name;
  }

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
      var res = await fetchRetry(GRAPH + url, {
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
      var res = await fetchRetry(GRAPH + "/me/drive/root:/" + encodedPath + ":/content", {
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
      var res2 = await fetchRetry(session.uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(end - start),
          "Content-Range": "bytes " + start + "-" + (end - 1) + "/" + bytes.length,
        },
        body: bytes.slice(start, end),
      });
      if (!res2.ok) { throw new Error("OneDrive chunk " + path + " -> " + res2.status); }
      if (end === bytes.length && res2.status !== 200 && res2.status !== 201) {
        throw new Error("OneDrive did not confirm completion of " + path + ". The bundle is incomplete.");
      }
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
    signOut: signOut,
    currentAccount: currentAccount,
    isSignedOut: isSignedOut,
    getToken: getToken,
    ensureDrive: ensureDrive,
    ensureFolder: ensureFolder,
    createBundleFolder: createBundleFolder,
    searchMessages: searchMessages,
    getMime: getMime,
    getBody: getBody,
    getAttachments: getAttachments,
    getAttachmentExport: getAttachmentExport,
    uploadFile: uploadFile,
    b64ToBytes: b64ToBytes,
    folderLink: folderLink,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
