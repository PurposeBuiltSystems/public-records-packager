/*
 * Public Records Packager — task pane UI wiring.
 *
 * Search → review checklist → build: for each selected message, upload the
 * exact .eml, a printable .html, and every attachment to a numbered,
 * chronological OneDrive folder with a manifest.csv for exemption review.
 */
/* global Office, GraphData, Packager, document */
(function () {
  "use strict";

  var found = [];   // search results
  var checked = {}; // message id -> bool

  /**
   * Account row. Certification policy 1100.5.7.1 requires a visible way out
   * wherever an add-in signs a user in. Every element access is guarded:
   * Outlook desktop caches the pane HTML while ?v= fetches fresh JS, so this
   * code can run against a page that predates these controls, and an
   * unguarded dereference here would throw inside Office.onReady and take
   * the whole pane down as "Add-in Error".
   */
  function authSet(id, k, v) { var e = document.getElementById(id); if (e) { e[k] = v; } }

  async function renderAuthState() {
    var who = null;
    try { who = await GraphData.currentAccount(); } catch (e) { who = null; }
    authSet("authWho", "textContent", who ? ("Signed in as " + who) : "Not signed in");
    authSet("signOut", "hidden", !who);
    authSet("signIn", "hidden", !!who);
  }

  async function doSignIn() {
    authSet("signIn", "disabled", true);
    try { await GraphData.getToken(); }
    catch (e) { authSet("authWho", "textContent", "Sign-in failed: " + ((e && e.message) || e)); }
    finally { authSet("signIn", "disabled", false); renderAuthState(); }
  }

  async function doSignOut() {
    authSet("signOut", "disabled", true);
    try {
      await GraphData.signOut();
      authSet("authWho", "textContent", "Signed out \u2014 this add-in's saved tokens are cleared. " +
        "Your Outlook session is separate and is not affected; no add-in can end it.");
    } catch (e) {
      authSet("authWho", "textContent", "Sign-out failed: " + ((e && e.message) || e));
    } finally {
      authSet("signOut", "disabled", false);
      setTimeout(renderAuthState, 2500);
    }
  }


  Office.onReady(function () {
    // Certification 1100.5.7.1 - sign-out must be reachable.
    var _si = document.getElementById("signIn");
    if (_si) { _si.addEventListener("click", doSignIn); }
    var _so = document.getElementById("signOut");
    if (_so) { _so.addEventListener("click", doSignOut); }
    renderAuthState();
    on("search", "click", search);
    on("selectAll", "click", selectAll);
    on("build", "click", build);
    setVal("bundleName", "Records Request " + new Date().toISOString().slice(0, 10));
  });

  function byId(id) { return document.getElementById(id); }

  /**
   * Guarded element access. Outlook desktop caches the pane HTML far harder
   * than the web client while ?v= still fetches today's JavaScript, so startup
   * routinely runs new code against an old page. One unguarded
   * `byId(x).value` there throws inside Office.onReady, and Outlook reports
   * that as "Add-in Error" - the whole pane, not one field. This is the exact
   * cause of certification finding 1120.3.7.8 on a sibling add-in.
   */
  function setVal(id, v) { var el = byId(id); if (el) { el.value = v; } }
  function setProp(id, k, v) { var el = byId(id); if (el) { el[k] = v; } }
  function setAttrIf(id, n, v) { var el = byId(id); if (el) { el.setAttribute(n, v); } }
  function rmAttrIf(id, n) { var el = byId(id); if (el) { el.removeAttribute(n); } }
  function isChecked(id) { var el = byId(id); return !!(el && el.checked); }

  /**
   * Outlook caches the pane HTML but the ?v= query string makes it fetch
   * JavaScript fresh, so a returning user can run today's JS against
   * yesterday's page. Binding through this helper means a missing element
   * costs one feature instead of throwing and leaving every later button
   * unbound — a whole dead pane.
   */
  function on(id, ev, fn) {
    var el = byId(id);
    if (el) { el.addEventListener(ev, fn); }
    return el;
  }

  function setStatus(kind, text) {
    var el = byId("status");
    if (!text) { el.hidden = true; return; }
    el.hidden = false;
    el.className = "status " + kind;
    el.textContent = text;
  }

  async function search() {
    var query = byId("query").value.trim();
    if (!query) { setStatus("error", "Enter search terms first."); return; }
    byId("search").disabled = true;
    byId("results").hidden = true;
    found = []; checked = {};
    try {
      setStatus("work", "Searching your mailbox…");
      var token = await GraphData.getToken();
      var res = await GraphData.searchMessages(token, query, function (n) {
        setStatus("work", "Searching your mailbox… " + n + " matches so far");
      });
      var filtered = Packager.applyFilters(res.messages, {
        fromDate: byId("fromDate").value || null,
        toDate: byId("toDate").value || null,
        participant: byId("participant").value || null,
      });
      // chronological for review, same order the bundle will use
      filtered.sort(function (a, b) {
        return (Date.parse(a.receivedDateTime || 0) || 0) - (Date.parse(b.receivedDateTime || 0) || 0);
      });
      found = filtered;
      found.forEach(function (m) { checked[m.id] = true; });
      if (!found.length) {
        setStatus("info", "No messages matched. Broaden the terms or the date range.");
        return;
      }
      render();
      byId("results").hidden = false;
      setStatus("info", found.length + " responsive message(s)" +
        (res.capped ? " (search capped at ~1000 — narrow the terms if this request is larger)" : "") +
        ". Uncheck anything out of scope, then build the bundle.");
      byId("options").removeAttribute("open");
    } catch (e) {
      var msg = (e && e.message) || String(e);
      if (/REPLACE_WITH_ENTRA_CLIENT_ID/.test(GraphData._config.clientId)) {
        msg = "Set your Entra client ID in src/graph.js before running. (" + msg + ")";
      }
      setStatus("error", "Search failed: " + msg);
    } finally {
      byId("search").disabled = false;
    }
  }

  function fmtDate(iso) {
    var d = new Date(iso);
    return isNaN(d) ? "" : d.toLocaleDateString();
  }

  function render() {
    var host = byId("messages");
    host.innerHTML = "";
    found.forEach(function (m) {
      var div = document.createElement("div");
      div.className = "person";
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = !!checked[m.id];
      cb.addEventListener("change", function () { checked[m.id] = cb.checked; });
      var who = document.createElement("div");
      who.className = "who";
      var name = document.createElement("div");
      name.className = "name";
      name.textContent = m.subject || "(no subject)";
      var meta = document.createElement("div");
      meta.className = "email";
      var from = ((m.from || {}).emailAddress || {});
      meta.textContent = fmtDate(m.receivedDateTime) + " — " + (from.name || from.address || "?") +
        (m.hasAttachments ? " — 📎" : "");
      who.appendChild(name);
      who.appendChild(meta);
      div.appendChild(cb);
      div.appendChild(who);
      host.appendChild(div);
    });
  }

  function selectAll() {
    var anyOff = found.some(function (m) { return !checked[m.id]; });
    found.forEach(function (m) { checked[m.id] = anyOff; });
    render();
  }

  async function build() {
    var selected = found.filter(function (m) { return checked[m.id]; });
    if (!selected.length) { setStatus("error", "Nothing selected."); return; }
    var bundleName = Packager.sanitize(byId("bundleName").value, 80) || "Records Request";
    var root = "Records Requests/" + bundleName;
    byId("build").disabled = true;
    try {
      var token = await GraphData.getToken();
      setStatus("work", "Checking OneDrive…");
      await GraphData.ensureDrive(token);
      await GraphData.ensureFolder(token, root);
      var plan = Packager.plan(selected);

      for (var i = 0; i < plan.length; i++) {
        var p = plan[i];
        setStatus("work", "Packaging " + (i + 1) + " of " + plan.length + ": " + p.base);

        // exact archival copy
        var mime = await GraphData.getMime(token, p.message.id);
        await GraphData.uploadFile(token, root + "/" + p.emlName, mime);

        // printable copy
        var body = await GraphData.getBody(token, p.message.id);
        await GraphData.uploadFile(token, root + "/" + p.htmlName, Packager.messageHtml(p.message, body));

        // attachments in original format
        if (p.message.hasAttachments) {
          var atts = await GraphData.getAttachments(token, p.message.id);
          var names = Packager.attachmentNames(atts.map(function (a) { return a.name; }));
          p.attachmentNames = names;
          if (atts.length) { await GraphData.ensureFolder(token, root + "/" + p.attDir); }
          for (var j = 0; j < atts.length; j++) {
            await GraphData.uploadFile(token, root + "/" + p.attDir + "/" + names[j], GraphData.b64ToBytes(atts[j].contentBytes));
          }
        }
      }

      setStatus("work", "Writing manifest…");
      await GraphData.uploadFile(token, root + "/manifest.csv", Packager.manifestCsv(plan));
      await GraphData.uploadFile(token, root + "/README.txt", Packager.readmeText(bundleName, plan.length));

      var link = await GraphData.folderLink(token, root);
      setStatus("info", "Bundle complete: " + plan.length + " messages in OneDrive → " + root +
        (link ? "  (" + link + ")" : ""));
      if (link) {
        try {
          if (Office.context.ui && Office.context.ui.openBrowserWindow) {
            Office.context.ui.openBrowserWindow(link);
          }
        } catch (e) { /* link is in the status line */ }
      }
    } catch (e) {
      setStatus("error", "Bundle failed: " + ((e && e.message) || e) + " — already-uploaded files remain in OneDrive; fix and rebuild into a new folder name.");
    } finally {
      byId("build").disabled = false;
    }
  }
})();
